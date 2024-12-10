#include "OutlookAPI.h"
#include <QInputDialog>
#include <QMessageBox>
#include <QDebug>
#include <QWidget>

#include <QStringView>
#include <QMetaProperty>
#include <QSettings>

#include <QVariant>
#include <iostream>
#include <oaidl.h>
#include "MSOUTL.h"
#include <QDebug>
std::shared_ptr< COutlookAPI > COutlookAPI::sInstance;
Q_DECLARE_METATYPE( std::shared_ptr< Outlook::Rule > );

COutlookAPI::COutlookAPI( QWidget *parent )
{
    getApplication();
    fParentWidget = parent;

    QSettings settings;
    setOnlyProcessUnread( settings.value( "OnlyProcessUnread", true ).toBool(), false );
    setProcessAllEmailWhenLessThan200Emails( settings.value( "ProcessAllEmailWhenLessThan200Emails", true ).toBool(), false );
    setRootFolder( settings.value( "RootFolder", R"(\Inbox)" ).toString(), false );

    qRegisterMetaType< std::shared_ptr< Outlook::Rule > >();
    qRegisterMetaType< std::shared_ptr< Outlook::Rule > >( "std::shared_ptr<Outlook::Rule>const&" );
}

std::shared_ptr< COutlookAPI > COutlookAPI::getInstance( QWidget *parent )
{
    if ( !sInstance )
    {
        Q_ASSERT( parent );
        sInstance = std::make_shared< COutlookAPI >( parent );
    }
    else
    {
        Q_ASSERT( !parent );
    }
    return sInstance;
}

COutlookAPI::~COutlookAPI()
{
    logout( false );
}

void COutlookAPI::logout( bool andNotify )
{
    fAccount.reset();
    fInbox.reset();
    fRootFolder.reset();
    fJunkFolder.reset();
    fContacts.reset();
    fRules.reset();

    if ( fLoggedIn && fOutlookApp && !fOutlookApp->isNull() && fOutlookApp->Session() )
    {
        Outlook::NameSpace( fOutlookApp->Session() ).Logoff();
        fLoggedIn = false;
        if ( andNotify )
            emit sigAccountChanged();
    }
}

QString COutlookAPI::accountName() const
{
    if ( !accountSelected() )
        return {};
    return fAccount->DisplayName();
}

bool COutlookAPI::accountSelected() const
{
    return fAccount.operator bool();
}

std::shared_ptr< Outlook::Account > COutlookAPI::selectAccount( bool notifyOnChange )
{
    logout( notifyOnChange );
    if ( fOutlookApp->isNull() )
        return {};

    auto profileName = fOutlookApp->DefaultProfileName();
    Outlook::NameSpace session( fOutlookApp->Session() );
    session.Logon( profileName );
    fLoggedIn = true;

    std::vector< std::shared_ptr< Outlook::Account > > allAccounts;

    auto accounts = session.Accounts();
    if ( !accounts )
    {
        logout( notifyOnChange );
        return {};
    }

    auto numAccounts = accounts->Count();
    allAccounts.reserve( numAccounts );

    QSettings settings;
    auto lastAccount = settings.value( "Account", QString() ).toString();
    int accountPos = 0;
    QStringList accountNames;
    std::map< QString, std::shared_ptr< Outlook::Account > > accountMap;
    for ( auto ii = 1; ii <= numAccounts; ++ii )
    {
        auto item = accounts->Item( ii );
        if ( !item )
            continue;

        auto account = getAccount( item );

        if ( account->AccountType() != Outlook::OlAccountType::olExchange )
            continue;

        switch ( account->ExchangeConnectionMode() )
        {
            case Outlook::OlExchangeConnectionMode::olNoExchange:
            case Outlook::OlExchangeConnectionMode::olOffline:
            case Outlook::OlExchangeConnectionMode::olCachedOffline:
            case Outlook::OlExchangeConnectionMode::olDisconnected:
            case Outlook::OlExchangeConnectionMode::olCachedDisconnected:
                continue;
            default:
                break;
        }

        allAccounts.push_back( account );
        auto accountName = account->DisplayName();
        accountNames << accountName;
        accountMap[ accountName ] = account;
        if ( accountName == lastAccount )
        {
            accountPos = ii - 1;
        }
    }

    if ( allAccounts.size() == 0 )
    {
        logout( notifyOnChange );
        return {};
    }

    if ( allAccounts.size() == 1 )
    {
        fAccount = allAccounts.front();
        settings.setValue( "Account", allAccounts.front()->DisplayName() );
        return fAccount;
    }

    bool aOK{ false };
    auto account = QInputDialog::getItem( fParentWidget, QString( "Select Account:" ), "Account:", accountNames, accountPos, false, &aOK );
    if ( !aOK || account.isEmpty() )
    {
        logout( notifyOnChange );
        return {};
    }
    auto pos = accountMap.find( account );
    if ( pos == accountMap.end() )
    {
        logout( notifyOnChange );
        return {};
    }
    settings.setValue( "Account", account );
    fAccount = ( *pos ).second;

    if ( notifyOnChange )
        emit sigAccountChanged();
    return fAccount;
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getContacts()
{
    if ( !fContacts )
        fContacts = selectContacts( false ).first;
    return fContacts;
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookAPI::selectContacts( bool singleOnly )
{
    if ( !fAccount )
    {
        if ( !selectAccount( true ) )
            return {};
    }

    return selectFolder(
        "Contacts",
        []( const std::shared_ptr< Outlook::Folder > &folder )
        {
            if ( !folder )
                return false;

            if ( folder->DefaultItemType() != Outlook::OlItemType::olContactItem )
                return false;

            auto items = folder->Items();
            if ( items->Count() == 0 )
                return false;

            if ( folder->Name().contains( "meta", Qt::CaseSensitivity::CaseInsensitive ) )
                return false;

            return true;
        },
        {}, singleOnly );
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getInbox()
{
    if ( !fInbox )
        fInbox = selectInbox( false ).first;
    return fInbox;
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookAPI::selectInbox( bool singleOnly )
{
    if ( !fAccount )
    {
        if ( !selectAccount( true ) )
            return {};
    }

    return getMailFolder( "Inbox", R"(Inbox)", singleOnly );
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getJunkFolder()
{
    if ( !fJunkFolder )
    {
        fJunkFolder = getMailFolder( "Junk Email", R"(Junk Email)", false ).first;
    }
    return fJunkFolder;
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookAPI::getMailFolder( const QString &folderLabel, const QString &path, bool singleOnly )
{
    if ( !accountSelected() )
        return {};

    auto retVal = selectFolder(
        folderLabel,
        [ this, path ]( const std::shared_ptr< Outlook::Folder > &folder )
        {
            if ( !folder )
                return false;
            if ( folder->DefaultItemType() != Outlook::OlItemType::olMailItem )
                return false;

            return isFolder( folder, path );
        },
        {}, singleOnly );
    return retVal;
}

bool COutlookAPI::isFolder( const std::shared_ptr< Outlook::Folder > &folder, const QString &path ) const
{
    return ( getFolderPath( folder, true ) == path ) || ( getFolderPath( folder, false ) == path );
}

std::shared_ptr< Outlook::Rules > COutlookAPI::getRules()
{
    if ( !fRules )
        fRules = selectRules();
    return fRules;
}

std::shared_ptr< Outlook::Folder > COutlookAPI::rootProcessFolder()
{
    if ( !accountSelected() )
        return fInbox;

    if ( fRootFolder )
        return fRootFolder;

    return getInbox();
}

QString COutlookAPI::rootProcessFolderName()
{
    return getFolderPath( rootProcessFolder() );
}

QString COutlookAPI::getFolderPath( const std::shared_ptr< Outlook::Folder > &folder, bool removeLeadingSlashes ) const
{
    if ( !folder )
        return {};

    auto retVal = folder->FullFolderPath();

    auto accountName = this->accountName();
    auto pos = retVal.indexOf( accountName + R"(\)" );
    if ( pos != -1 )
        retVal = retVal.remove( pos, accountName.length() + 1 );

    while ( removeLeadingSlashes && retVal.startsWith( R"(\)" ) )
        retVal = retVal.mid( 1 );
    return retVal;
}

void COutlookAPI::slotHandleException( int code, const QString &source, const QString &desc, const QString &help )
{
    auto msg = QString( "%1 - %2: %3" ).arg( source ).arg( code );
    auto txt = "<br>" + desc + "</br>";
    if ( !help.isEmpty() )
        txt += "<br>" + help + "</br>";
    msg = msg.arg( txt );

    QMessageBox::critical( nullptr, "Exception Thrown", msg );
}

std::shared_ptr< Outlook::Rules > COutlookAPI::selectRules()
{
    if ( !fAccount )
    {
        if ( !selectAccount( true ) )
            return {};
    }

    if ( !fAccount || fAccount->isNull() )
        return {};

    auto store = connectToException( fAccount->DeliveryStore() );
    if ( !store )
        return {};

    auto rules = store->GetRules();
    return getRules( rules );
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookAPI::selectFolder( const QString &folderName, const TFolderFunc &acceptFolder, const TFolderFunc &checkChildFolders, bool singleOnly )
{
    auto &&folders = getFolders( false, acceptFolder, checkChildFolders );
    return selectFolder( folderName, folders, singleOnly );
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookAPI::selectFolder( const QString &folderName, const std::list< std::shared_ptr< Outlook::Folder > > &folders, bool singleOnly )
{
    if ( folders.empty() )
    {
        QMessageBox::critical( fParentWidget, QString( "Could not find %1" ).arg( folderName.toLower() ), folderName + " not found" );
        return { {}, false };
    }
    if ( folders.size() == 1 )
        return { folders.front(), false };
    if ( singleOnly )
        return { {}, false };

    QStringList folderNames;
    std::map< QString, std::shared_ptr< Outlook::Folder > > folderMap;

    for ( auto &&ii : folders )
    {
        auto path = ii->FullFolderPath();
        folderNames << path;
        folderMap[ path ] = ii;
    }
    bool aOK{ false };
    auto item = QInputDialog::getItem( fParentWidget, QString( "Select %1 Folder" ).arg( folderName ), folderName + " Folder:", folderNames, 0, false, &aOK );
    if ( !aOK )
        return { {}, false };
    auto pos = folderMap.find( item );
    if ( pos == folderMap.end() )
        return { {}, false };
    return { ( *pos ).second, true };
}

std::list< std::shared_ptr< Outlook::Folder > > COutlookAPI::getFolders( bool recursive, const TFolderFunc &acceptFolder, const TFolderFunc &checkChildFolders )
{
    if ( !fAccount )
    {
        if ( !selectAccount( true ) )
            return {};
    }

    if ( !fAccount || fAccount->isNull() )
        return {};

    auto store = fAccount->DeliveryStore();
    if ( !store )
        return {};

    auto root = getMailFolder( store->GetRootFolder() );
    auto retVal = getFolders( root, recursive, acceptFolder, checkChildFolders );

    return retVal;
}

std::list< std::shared_ptr< Outlook::Folder > > COutlookAPI::getFolders( const std::shared_ptr< Outlook::Folder > &parent, bool recursive, const TFolderFunc &acceptFolder, const TFolderFunc &checkChildFolders )
{
    if ( !parent )
        return {};

    std::list< std::shared_ptr< Outlook::Folder > > retVal;

    auto folders = parent->Folders();
    auto folderCount = folders->Count();
    for ( auto jj = 1; jj <= folderCount; ++jj )
    {
        auto folder = getMailFolder( folders->Item( jj ) );

        bool isMatch = !acceptFolder || ( acceptFolder && acceptFolder( folder ) );
        bool cont = recursive && ( !checkChildFolders || ( checkChildFolders && checkChildFolders( folder ) ) );

        if ( isMatch )
            retVal.push_back( folder );
        if ( cont )
        {
            auto &&subFolders = getFolders( folder, true, acceptFolder );
            retVal.insert( retVal.end(), subFolders.begin(), subFolders.end() );
        }
    }

    retVal.sort(
        []( const std::shared_ptr< Outlook::Folder > &lhs, const std::shared_ptr< Outlook::Folder > &rhs )
        {
            if ( !lhs )
                return false;
            if ( !rhs )
                return true;
            return lhs->FullFolderPath() < rhs->FullFolderPath();
        } );
    return retVal;
}

int COutlookAPI::recursiveSubFolderCount( const Outlook::Folder *parent )
{
    if ( !parent )
        return 0;

    emit sigInitStatus( "Counting Folders:", 0 );
    auto retVal = subFolderCount( parent, true );
    emit sigStatusFinished( "Counting Folders:" );
    return retVal;
}

int COutlookAPI::subFolderCount( const Outlook::Folder *parent, bool recursive )
{
    if ( !parent )
        return 0;

    if ( !recursive )
        emit sigInitStatus( "Counting Folders:", 0 );

    auto folders = parent->Folders();
    auto folderCount = folders->Count();

    int retVal = folderCount;
    for ( auto jj = 1; recursive && ( jj <= folderCount ); ++jj )
    {
        auto folder = reinterpret_cast< Outlook::Folder * >( folders->Item( jj ) );
        if ( !folder )
            continue;
        retVal += subFolderCount( folder, recursive );
    }

    if ( !recursive )
        emit sigStatusFinished( "Counting Folders:" );

    return retVal;
}

bool validEmail( const QString &address )
{
    auto split = address.splitRef( "@", QString::SkipEmptyParts );
    return split.size() == 2;
}

bool validEmails( const QStringList &addresses )
{
    for ( auto &&ii : addresses )
    {
        if ( !validEmail( ii ) )
            return false;
    }
    return true;
}

QString COutlookAPI::ruleNameForFolder( const std::shared_ptr< Outlook::Folder > &folder )
{
    return ruleNameForFolder( folder.get() );
}

QString COutlookAPI::ruleNameForFolder( Outlook::Folder *folder )
{
    if ( !folder )
        return {};
    auto path = folder->FullFolderPath();

    auto pos = path.indexOf( "Inbox" );
    QString ruleName;
    if ( pos != -1 )
    {
        ruleName = path.mid( pos + 6 ).replace( R"(\)", "-" );
        if ( ruleName.isEmpty() )
            ruleName = "Inbox";
    }
    else
    {
        pos = path.lastIndexOf( R"(\)" );

        if ( pos != -1 )
            ruleName = path.mid( pos + 1 );
        else
            ruleName = path;
    }

    ruleName = ruleName.replace( "%2F", "/" );
    return ruleName;
}

QString COutlookAPI::folderName( const std::shared_ptr< Outlook::Folder > &folder )
{
    return folderName( folder.get() );
}

QString COutlookAPI::folderName( const Outlook::Folder *folder )
{
    if ( !folder )
        return {};
    auto retVal = folder->Name();
    retVal = retVal.replace( "%2F", "/" );
    return retVal;
}

bool COutlookAPI::addRule( const std::shared_ptr< Outlook::Folder > &folder, const QStringList &rules, QStringList &msgs )
{
    if ( !folder )
        return false;

    auto ruleName = ruleNameForFolder( folder );

    auto rule = std::shared_ptr< Outlook::Rule >( fRules->Create( ruleName, Outlook::OlRuleType::olRuleReceive ) );
    if ( !rule )
    {
        msgs.push_back( QString( "Could not create rule '%1'" ).arg( ruleName ) );
        return false;
    }

    auto moveAction = rule->Actions()->MoveToFolder();
    if ( !moveAction )
    {
        msgs.push_back( QString( "Internal error" ) );
        return false;
    }
    moveAction->SetEnabled( true );
    moveAction->SetFolder( reinterpret_cast< Outlook::MAPIFolder * >( folder.get() ) );

    rule->Actions()->Stop()->SetEnabled( true );

    if ( !addRecipientsToRule( rule.get(), rules, msgs ) )
        return false;

    auto name = ruleNameForRule( rule );
    if ( rule->Name() != name )
        rule->SetName( name );

    saveRules();

    bool retVal = runRule( rule );
    emit sigRuleAdded( rule );
    return retVal;
}

void COutlookAPI::saveRules()
{
    emit sigStatusMessage( QString( "Saving Rules" ) );
    fRules->Save( true );
}

bool COutlookAPI::addToRule( std::shared_ptr< Outlook::Rule > rule, const QStringList &rules, QStringList &msgs )
{
    if ( !rule || rules.isEmpty() || !fRules )
    {
        msgs.push_back( "Parameters not set" );
        return false;
    }

    if ( !addRecipientsToRule( rule.get(), rules, msgs ) )
        return false;

    saveRules();

    bool retVal = runRule( rule );
    emit sigRuleChanged( rule );
    return retVal;
}

bool COutlookAPI::deleteRule( std::shared_ptr< Outlook::Rule > rule )
{
    if ( !rule || !fRules )
        return false;
    auto name = rule->Name();
    auto idx = rule->ExecutionOrder();
    auto ruleName = QString( "%1 (%2)" ).arg( name ).arg( idx );

    emit sigStatusMessage( QString( "Deleting Rule: %1" ).arg( ruleName ) );
    fRules->Remove( idx );

    saveRules();
    QMessageBox::information( fParentWidget, "Deleted Rule", QString( "Deleted Rule: %1" ).arg( ruleName ) );

    emit sigRuleDeleted( rule );
    return true;
}

bool COutlookAPI::runAllRulesOnAllFolders()
{
    auto allRules = getAllRules();
    auto inbox = getInbox();
    auto junk = getJunkFolder();

    bool retVal = true;

    int numFolders = recursiveSubFolderCount( inbox.get() );

    auto msg = QString( "Running All Rules on All Folders:" );
    auto totalFolders = numFolders + ( junk ? 1 : 0 );
    emit sigInitStatus( msg, totalFolders );

    if ( inbox )
        retVal = runRules( allRules, inbox, true, msg ) && retVal;

    if ( junk )
        retVal = runRules( allRules, junk, false, msg ) && retVal;
    return retVal;
}

bool COutlookAPI::runAllRules( const std::shared_ptr< Outlook::Folder > &folder )
{
    return runRules( {}, folder );
}

std::vector< std::shared_ptr< Outlook::Rule > > COutlookAPI::getAllRules()
{
    if ( !fRules )
        return {};

    std::vector< std::shared_ptr< Outlook::Rule > > rules;
    rules.reserve( fRules->Count() );
    auto numRules = fRules->Count();
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        auto rule = getRule( fRules->Item( ii ) );
        rules.push_back( rule );
    }
    return rules;
}

bool COutlookAPI::runRule( std::shared_ptr< Outlook::Rule > rule, const std::shared_ptr< Outlook::Folder > &folder )
{
    return runRules( std::vector< std::shared_ptr< Outlook::Rule > >( { rule } ), folder );
}

bool COutlookAPI::runRules( std::vector< std::shared_ptr< Outlook::Rule > > rules, std::shared_ptr< Outlook::Folder > folder, bool recursive, const std::optional< QString > & perFolderMsg /*={}*/ )
{
    if ( !folder )
        folder = rootProcessFolder();

    if ( !folder )
        return false;

    auto folderPtr = reinterpret_cast< Outlook::MAPIFolder * >( folder.get() );
    auto folderTypeID = qRegisterMetaType< Outlook::MAPIFolder * >( "MAPIFolder*", &folderPtr );

    auto msg = QString( "Running Rules on '%1':" ).arg( getFolderPath( folder ) );
    emit sigInitStatus( msg, static_cast< int >( rules.size() ) );


    if ( perFolderMsg.has_value() )
    {
        emit sigIncStatusValue( perFolderMsg.value() );
    }

    if ( rules.empty() )
        rules = getAllRules();

    for ( auto &&rule : rules )
    {
        if ( canceled() )
            return false;

        if ( !rule || !rule->Enabled() )
            continue;

        auto inboxPtr = fInbox.get();
        emit sigStatusMessage( QString( "Running Rule: %1 on Folder: %2" ).arg( rule->Name() ).arg( getFolderPath( folder ) ) );
        rule->Execute( false, QVariant( folderTypeID, &folderPtr ) );
        emit sigIncStatusValue( msg );
    }

    bool retVal = true;
    if ( recursive )
    {
        auto childFolders = getFolders( folder, false );

        for ( auto &&ii : childFolders )
        {
            retVal = runRules( rules, ii, recursive, perFolderMsg ) && retVal;
        }
    }

    return retVal;
}

bool COutlookAPI::addRecipientsToRule( Outlook::Rule *rule, const QStringList &recipients, QStringList &msgs )
{
    if ( recipients.isEmpty() )
        return true;

    if ( !rule || !rule->Conditions() )
        return false;

    auto cond = rule->Conditions()->SenderAddress();
    if ( !cond )
    {
        msgs.push_back( QString( "Internal error" ) );
        return false;
    }

    auto addresses = mergeRecipients( rule, recipients, &msgs );
    if ( !addresses.has_value() )
        return false;

    cond->SetAddress( addresses.value() );
    cond->SetEnabled( true );

    return true;
}

std::optional< QStringList > COutlookAPI::getRecipients( Outlook::Rule *rule, QStringList *msgs )
{
    if ( !rule || !rule->Conditions() )
        return {};

    auto cond = rule->Conditions()->SenderAddress();
    if ( !cond )
    {
        if ( msgs )
            msgs->push_back( QString( "Internal error" ) );
        return {};
    }

    QStringList addresses;
    if ( cond->Enabled() )
    {
        auto variant = cond->Address();
        if ( variant.type() == QVariant::Type::String )
            addresses << variant.toString();
        else if ( variant.type() == QVariant::Type::StringList )
            addresses << variant.toStringList();
    }
    return addresses;
}

std::optional< QStringList > COutlookAPI::mergeRecipients( Outlook::Rule *lhs, const QStringList &rhs, QStringList *msgs )
{
    auto lhsRecipients = getRecipients( lhs, msgs );
    if ( !lhsRecipients.has_value() )
        return {};
    if ( !lhsRecipients )
        return rhs;
    lhsRecipients.value() << rhs;
    lhsRecipients.value().removeDuplicates();
    return lhsRecipients;
}

std::optional< QStringList > COutlookAPI::mergeRecipients( Outlook::Rule *lhs, Outlook::Rule *rhs, QStringList *msgs )
{
    auto lhsRecipients = getRecipients( lhs, msgs );
    auto rhsRecipients = getRecipients( rhs, msgs );
    if ( !lhsRecipients.has_value() && !rhsRecipients.has_value() )
        return {};
    if ( lhsRecipients && !rhsRecipients )
        return lhsRecipients;
    if ( !lhsRecipients && rhsRecipients )
        return rhsRecipients;
    lhsRecipients.value() << rhsRecipients.value();
    lhsRecipients.value().removeDuplicates();
    return lhsRecipients;
}

bool hasProperty( QObject *item, const char *propName )
{
    auto mo = item->metaObject();
    //auto aoxMO = dynamic_cast< c * >( mo );
    auto idx = mo->indexOfProperty( propName );
    return idx != -1;
}

Outlook::OlObjectClass COutlookAPI::getObjectClass( IDispatch *item )
{
    if ( !item )
        return {};

    IDispatch *pdisp = (IDispatch *)NULL;
    DISPID dispid;
    OLECHAR *szMember = L"Class";
    auto result = item->GetIDsOfNames( IID_NULL, &szMember, 1, LOCALE_SYSTEM_DEFAULT, &dispid );

    if ( result == S_OK )
    {
        VARIANT resultant{};
        DISPPARAMS params{ 0 };
        EXCEPINFO excepInfo{};
        UINT argErr{ 0 };
        result = item->Invoke( dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD | DISPATCH_PROPERTYGET, &params, &resultant, &excepInfo, &argErr );
        if ( result == S_OK )
        {
            return static_cast< Outlook::OlObjectClass >( resultant.lVal );
        }
    }

    auto retVal = QAxObject( item ).property( "Class" );

    return static_cast< Outlook::OlObjectClass >( retVal.toInt() );
}

void COutlookAPI::dumpSession( Outlook::NameSpace &session )
{
    auto stores = session.Stores();
    auto numStores = stores->Count();
    for ( auto ii = 1; ii <= numStores; ++ii )
    {
        auto store = stores->Item( ii );
        if ( !store )
            continue;
        auto root = store->GetRootFolder();
        //qDebug() << root->FullFolderPath();
        dumpFolder( reinterpret_cast< Outlook::Folder * >( root ) );
    }
}

void COutlookAPI::dumpFolder( Outlook::Folder *parent )
{
    if ( !parent )
        return;

    auto folders = parent->Folders();
    auto folderCount = folders->Count();
    for ( auto jj = 1; jj <= folderCount; ++jj )
    {
        auto folder = reinterpret_cast< Outlook::Folder * >( folders->Item( jj ) );
        qDebug() << folder->FullFolderPath() << toString( folder->DefaultItemType() );
        dumpFolder( folder );
    }
}

QString toString( Outlook::OlItemType olItemType )
{
    switch ( olItemType )
    {
        case Outlook::OlItemType::olMailItem:
            return "Mail";
        case Outlook::OlItemType::olAppointmentItem:
            return "Appointment";
        case Outlook::OlItemType::olContactItem:
            return "Contact";
        case Outlook::OlItemType::olTaskItem:
            return "Task";
        case Outlook::OlItemType::olJournalItem:
            return "Journal";
        case Outlook::OlItemType::olNoteItem:
            return "Note";
        case Outlook::OlItemType::olPostItem:
            return "Post";
        case Outlook::OlItemType::olDistributionListItem:
            return "Distribution List";
        case Outlook::OlItemType::olMobileItemSMS:
            return "Mobile Item SMS";
        case Outlook::OlItemType::olMobileItemMMS:
            return "Mobile Item MMS";
    }
    return "<UNKNOWN>";
}

QString toString( Outlook::OlRuleConditionType olRuleConditionType )
{
    switch ( olRuleConditionType )
    {
        case Outlook::OlRuleConditionType::olConditionUnknown:
            return "ConditionUnknown";
        case Outlook::OlRuleConditionType::olConditionFrom:
            return "ConditionFrom";
        case Outlook::OlRuleConditionType::olConditionSubject:
            return "ConditionSubject";
        case Outlook::OlRuleConditionType::olConditionAccount:
            return "ConditionAccount";
        case Outlook::OlRuleConditionType::olConditionOnlyToMe:
            return "ConditionOnlyToMe";
        case Outlook::OlRuleConditionType::olConditionTo:
            return "ConditionTo";
        case Outlook::OlRuleConditionType::olConditionImportance:
            return "ConditionImportance";
        case Outlook::OlRuleConditionType::olConditionSensitivity:
            return "ConditionSensitivity";
        case Outlook::OlRuleConditionType::olConditionFlaggedForAction:
            return "ConditionFlaggedForAction";
        case Outlook::OlRuleConditionType::olConditionCc:
            return "ConditionCc";
        case Outlook::OlRuleConditionType::olConditionToOrCc:
            return "ConditionToOrCc";
        case Outlook::OlRuleConditionType::olConditionNotTo:
            return "ConditionNotTo";
        case Outlook::OlRuleConditionType::olConditionSentTo:
            return "ConditionSentTo";
        case Outlook::OlRuleConditionType::olConditionBody:
            return "ConditionBody";
        case Outlook::OlRuleConditionType::olConditionBodyOrSubject:
            return "ConditionBodyOrSubject";
        case Outlook::OlRuleConditionType::olConditionMessageHeader:
            return "ConditionMessageHeader";
        case Outlook::OlRuleConditionType::olConditionRecipientAddress:
            return "ConditionRecipientAddress";
        case Outlook::OlRuleConditionType::olConditionSenderAddress:
            return "ConditionSenderAddress";
        case Outlook::OlRuleConditionType::olConditionCategory:
            return "ConditionCategory";
        case Outlook::OlRuleConditionType::olConditionOOF:
            return "ConditionOOF";
        case Outlook::OlRuleConditionType::olConditionHasAttachment:
            return "ConditionHasAttachment";
        case Outlook::OlRuleConditionType::olConditionSizeRange:
            return "ConditionSizeRange";
        case Outlook::OlRuleConditionType::olConditionDateRange:
            return "ConditionDateRange";
        case Outlook::OlRuleConditionType::olConditionFormName:
            return "ConditionFormName";
        case Outlook::OlRuleConditionType::olConditionProperty:
            return "ConditionProperty";
        case Outlook::OlRuleConditionType::olConditionSenderInAddressBook:
            return "ConditionSenderInAddressBook";
        case Outlook::OlRuleConditionType::olConditionMeetingInviteOrUpdate:
            return "ConditionMeetingInviteOrUpdate";
        case Outlook::OlRuleConditionType::olConditionLocalMachineOnly:
            return "ConditionLocalMachineOnly";
        case Outlook::OlRuleConditionType::olConditionOtherMachine:
            return "ConditionOtherMachine";
        case Outlook::OlRuleConditionType::olConditionAnyCategory:
            return "ConditionAnyCategory";
        case Outlook::OlRuleConditionType::olConditionFromRssFeed:
            return "ConditionFromRssFeed";
        case Outlook::OlRuleConditionType::olConditionFromAnyRssFeed:
            return "ConditionFromAnyRssFeed";
    }
    return "<UNKNOWN>";
}

QString toString( Outlook::OlImportance importance )
{
    switch ( importance )
    {
        case Outlook::OlImportance::olImportanceLow:
            return "Low";
        case Outlook::OlImportance::olImportanceNormal:
            return "Normal";
        case Outlook::OlImportance::olImportanceHigh:
            return "High";
    }

    return "<UNKNOWN>";
}

QString toString( Outlook::OlSensitivity sensitivity )
{
    switch ( sensitivity )
    {
        case Outlook::OlSensitivity::olPersonal:
            return "Personal";
        case Outlook::OlSensitivity::olNormal:
            return "Normal";
        case Outlook::OlSensitivity::olPrivate:
            return "Private";
        case Outlook::OlSensitivity::olConfidential:
            return "Confidential";
    }

    return "<UNKNOWN>";
}

QString toString( Outlook::OlMarkInterval markInterval )
{
    switch ( markInterval )
    {
        case Outlook::OlMarkInterval::olMarkToday:
            return "Mark Today";
        case Outlook::OlMarkInterval::olMarkTomorrow:
            return "Mark Tomorrow";
        case Outlook::OlMarkInterval::olMarkThisWeek:
            return "Mark This Week";
        case Outlook::OlMarkInterval::olMarkNextWeek:
            return "Mark Next Week";
        case Outlook::OlMarkInterval::olMarkNoDate:
            return "Mark No Date";
        case Outlook::OlMarkInterval::olMarkComplete:
            return "Mark Complete";
    }

    return "<UNKNOWN>";
}

QString getValue( const QVariant &variant, const QString &joinSeparator )
{
    QString retVal;
    if ( variant.type() == QVariant::Type::StringList )
        retVal = variant.toStringList().join( joinSeparator );
    else
    {
        Q_ASSERT( variant.canConvert( QVariant::Type::String ) );
        retVal = variant.toString();
    }
    return retVal;
}

void dumpMetaMethods( QObject *object )
{
    if ( !object )
        return;
    auto metaObject = object->metaObject();

    QStringList sigs;
    QStringList slotList;
    QStringList constructors;
    QStringList methods;

    for ( int methodIdx = metaObject->methodOffset(); methodIdx < metaObject->methodCount(); ++methodIdx )
    {
        auto mmTest = metaObject->method( methodIdx );
        auto signature = QString( mmTest.methodSignature() );
        switch ( mmTest.methodType() )
        {
            case QMetaMethod::Signal:
                sigs << signature;
                break;
            case QMetaMethod::Slot:
                slotList << signature;
                break;
            case QMetaMethod::Constructor:
                constructors << signature;
                break;
            case QMetaMethod::Method:
                methods << signature;
                break;
        }
    }
    qDebug() << object;
    qDebug() << "Signals:";
    for ( auto &&ii : sigs )
        qDebug() << ii;

    qDebug() << "Slots:";
    for ( auto &&ii : slotList )
        qDebug() << ii;

    qDebug() << "Constructors:";
    for ( auto &&ii : constructors )
        qDebug() << ii;

    qDebug() << "Methods:";
    for ( auto &&ii : methods )
        qDebug() << ii;
}

COutlookAPI::EAddressTypes operator|( const COutlookAPI::EAddressTypes &lhs, const COutlookAPI::EAddressTypes &rhs )
{
    auto lhsA = static_cast< int >( lhs );
    auto rhsA = static_cast< int >( rhs );
    return static_cast< COutlookAPI::EAddressTypes >( lhsA | rhsA );
}

bool COutlookAPI::renameRules()
{
    if ( !fRules )
        return false;

    auto numRules = fRules->Count();

    emit sigInitStatus( "Analyzing Rule Names:", numRules );

    std::list< std::pair< std::shared_ptr< Outlook::Rule >, QString > > changes;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;

        emit sigIncStatusValue( "Analyzing Rule Names:" );
        auto rule = getRule( fRules->Item( ii ) );
        if ( !rule )
            continue;

        auto ruleName = ruleNameForRule( rule );
        auto currName = rule->Name();
        if ( ruleName != currName )
        {
            changes.emplace_back( rule, ruleName );
        }
    }
    if ( canceled() )
        return false;

    if ( changes.empty() )
    {
        QMessageBox::information( fParentWidget, "Renamed Rules", QString( "No rules needed renaming" ) );
        return 0;
    }
    QStringList tmp;
    for ( auto &&ii : changes )
    {
        tmp << "<li>" + ii.first->Name() + " => " + ii.second + "</li>";
    }
    auto msg = QString( "Rules to be changed:<ul>%1</ul>Continue?" ).arg( tmp.join( "\n" ) );
    auto process = QMessageBox::information( fParentWidget, "Renamed Rules", msg, QMessageBox::Yes | QMessageBox::No );
    if ( process == QMessageBox::No )
        return 0;

    emit sigInitStatus( "Renaming Rules:", static_cast< int >( changes.size() ) );
    for ( auto &&ii : changes )
    {
        ii.first->SetName( ii.second );
        emit sigIncStatusValue( "Renaming Rules:" );
    }
    saveRules();

    return changes.size();
}

template< typename T >
QString addConditionBase( T *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( condition && condition->Enabled() )
    {
        if ( forDisplayOnly )
            return "<" + conditionStr + ">";
        else
            return "(" + conditionStr + ")";
    }
    return {};
}

QString addCondition( Outlook::AccountRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + toString( condition->ConditionType() );
    return addConditionBase( condition, retVal, forDisplayOnly );
}

QString addCondition( Outlook::RuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=Yes";
    return addConditionBase( condition, retVal, forDisplayOnly );
}

QString addCondition( Outlook::TextRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + getValue( condition->Text(), " or " );
    return addConditionBase( condition, retVal, forDisplayOnly );
}

QString addCondition( Outlook::CategoryRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + getValue( condition->Categories(), " or " );
    return addConditionBase( condition, retVal, forDisplayOnly );
}

QString addCondition( Outlook::ToOrFromRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=";

    auto recipients = COutlookAPI::getRecipientEmails( condition->Recipients(), {}, false );
    retVal += recipients.join( " or " );

    return addConditionBase( condition, retVal, forDisplayOnly );
}

QString addCondition( Outlook::FormNameRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + getValue( condition->FormName(), " or " );
    return addConditionBase( condition, retVal, forDisplayOnly );
}

QString addCondition( Outlook::FromRssFeedRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + getValue( condition->FromRssFeed(), " or " );
    return addConditionBase( condition, retVal, forDisplayOnly );
}

QString addCondition( Outlook::ImportanceRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + toString( condition->Importance() );
    return addConditionBase( condition, retVal, forDisplayOnly );
}

QString addCondition( Outlook::AddressRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + getValue( condition->Address(), " or " );
    return addConditionBase( condition, retVal, forDisplayOnly );
}

QString addCondition( Outlook::SenderInAddressListRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto addresses = COutlookAPI::getInstance()->getEmailAddresses( condition->AddressList(), false );
    auto retVal = conditionStr + "=";
    retVal += addresses.join( " or " );

    return addConditionBase( condition, retVal, forDisplayOnly );
}

QString addCondition( Outlook::SensitivityRuleCondition *condition, const QString &conditionStr, bool forDisplayOnly )
{
    if ( !condition || !condition->Enabled() )
        return {};

    auto retVal = conditionStr + "=" + toString( condition->Sensitivity() );
    return addConditionBase( condition, retVal, forDisplayOnly );
}

QString COutlookAPI::ruleNameForRule( std::shared_ptr< Outlook::Rule > rule, bool forDisplay )
{
    QStringList addOns;
    if ( !rule )
        addOns << "INV-NULLPTR";

    bool isEnabled = rule ? rule->Enabled() : false;
    auto actions = rule ? rule->Actions() : nullptr;
    if ( !actions )
    {
        addOns << "INV-NOACTIONS";
    }

    Outlook::MAPIFolder *destFolder = nullptr;
    auto mvToFolderAction = actions ? actions->MoveToFolder() : nullptr;
    if ( mvToFolderAction )
    {
        isEnabled = isEnabled && mvToFolderAction->Enabled();
        destFolder = mvToFolderAction->Folder();
        if ( !destFolder )
            addOns << "NOFOLDER";
    }
    else
    {
        addOns << "INV-NOMOVEACTION";
    }

    QStringList conditions;
    if ( !forDisplay && rule && rule->Conditions() )
    {
        conditions << addCondition( rule->Conditions()->Account(), "Account", forDisplay );
        conditions << addCondition( rule->Conditions()->AnyCategory(), "AnyCategory", forDisplay );
        conditions << addCondition( rule->Conditions()->Body(), "Body", forDisplay );
        conditions << addCondition( rule->Conditions()->BodyOrSubject(), "BodyOrSubject", forDisplay );
        conditions << addCondition( rule->Conditions()->CC(), "CC", forDisplay );
        conditions << addCondition( rule->Conditions()->Category(), "Category", forDisplay );
        conditions << addCondition( rule->Conditions()->FormName(), "FormName", forDisplay );
        conditions << addCondition( rule->Conditions()->From(), "From", forDisplay );
        conditions << addCondition( rule->Conditions()->FromAnyRSSFeed(), "FromAnyRSSFeed", forDisplay );
        conditions << addCondition( rule->Conditions()->FromRssFeed(), "FromRssFeed", forDisplay );
        conditions << addCondition( rule->Conditions()->HasAttachment(), "HasAttachment", forDisplay );
        conditions << addCondition( rule->Conditions()->Importance(), "Importance", forDisplay );
        conditions << addCondition( rule->Conditions()->MeetingInviteOrUpdate(), "MeetingInviteOrUpdate", forDisplay );
        conditions << addCondition( rule->Conditions()->MessageHeader(), "MessageHeader", forDisplay );
        conditions << addCondition( rule->Conditions()->NotTo(), "NotTo", forDisplay );
        conditions << addCondition( rule->Conditions()->OnLocalMachine(), "OnLocalMachine", forDisplay );
        conditions << addCondition( rule->Conditions()->OnOtherMachine(), "OnOtherMachine", forDisplay );
        conditions << addCondition( rule->Conditions()->OnlyToMe(), "OnlyToMe", forDisplay );
        conditions << addCondition( rule->Conditions()->RecipientAddress(), "RecipientAddress", forDisplay );
        //conditions << addCondition( rule->Conditions()->SenderAddress(), "SenderAddress", forDisplay );
        conditions << addCondition( rule->Conditions()->SenderInAddressList(), "SenderInAddressList", forDisplay );
        conditions << addCondition( rule->Conditions()->Sensitivity(), "Sensitivity", forDisplay );
        conditions << addCondition( rule->Conditions()->SentTo(), "SentTo", forDisplay );
        conditions << addCondition( rule->Conditions()->Subject(), "Subject", forDisplay );
        conditions << addCondition( rule->Conditions()->ToMe(), "ToMe", forDisplay );
        conditions << addCondition( rule->Conditions()->ToOrCc(), "ToOrCc", forDisplay );
    }

    if ( !isEnabled )
        conditions << ( forDisplay ? "(Disabled)" : "<Disabled>" );

    QString ruleName;
    if ( forDisplay && rule )
        ruleName = rule->Name();
    else
        ruleName = ruleNameForFolder( reinterpret_cast< Outlook::Folder * >( destFolder ) );

    if ( ruleName.isEmpty() )
        ruleName = "<UNNAMED RULE>";

    conditions.removeAll( QString() );
    conditions.sort();

    addOns.removeAll( QString() );
    addOns.sort();

    auto suffixes = QStringList() << ruleName << addOns.join( " " ) << conditions.join( " " ) << ( ( forDisplay ) ? ( rule ? QString( "(%1)" ).arg( rule->ExecutionOrder() ) : QString( "(INV_EXECUTION_ORDER)" ) ) : QString() );
    suffixes.removeAll( QString() );

    return suffixes.join( " " ).trimmed();
}

bool COutlookAPI::sortRules()
{
    if ( !fRules )
        return false;

    auto numRules = fRules->Count();
    emit sigInitStatus( "Sorting Rules:", numRules );

    std::list< Outlook::_Rule * > rules;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;
        auto rule = fRules->Item( ii );
        emit sigIncStatusValue( "Sorting Rules:" );
        if ( !rule )
            continue;
        rules.push_back( rule );
    }
    if ( canceled() )
        return false;

    rules.sort(
        []( Outlook::_Rule *lhs, Outlook::_Rule *rhs )
        {
            if ( !lhs )
                return false;
            if ( !rhs )
                return true;
            auto lhsName = lhs->Name();
            auto rhsName = rhs->Name();
            if ( lhsName.startsWith( rhsName ) && ( lhsName != rhsName ) )
                return true;
            else if ( rhsName.startsWith( lhsName ) && ( lhsName != rhsName ) )
                return false;
            else
                return lhsName < rhsName;
        } );

    if ( canceled() )
        return false;
    bool changed = false;
    auto pos = 1;
    emit sigInitStatus( "Recomputing Execution Order:", numRules );

    for ( auto &&ii : rules )
    {
        if ( canceled() )
            return false;

        changed = changed || ( ii->ExecutionOrder() != pos );
        ii->SetExecutionOrder( pos++ );
        emit sigIncStatusValue( "Recomputing Execution Order:" );
    }
    if ( changed )
        saveRules();
    return changed;
}

bool COutlookAPI::enableAllRules()
{
    if ( !fRules )
        return false;

    auto numRules = fRules->Count();
    emit sigInitStatus( "Enabling Rules:", numRules );

    std::list< Outlook::_Rule * > rules;
    int numChanged = 0;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;
        auto rule = fRules->Item( ii );
        emit sigIncStatusValue( "Enabling Rules:" );
        if ( !rule )
            continue;
        if ( rule->Enabled() )
            continue;
        rule->SetEnabled( true );
        numChanged++;
    }
    if ( canceled() )
        return false;

    if ( numChanged != 0 )
        saveRules();

    QMessageBox::information( fParentWidget, R"(Enable All Rules)", QString( "%1 rules enabled" ).arg( numChanged ) );

    return numChanged != 0;
}

bool COutlookAPI::moveFromToAddress()
{
    if ( !fRules )
        return false;

    auto numRules = fRules->Count();
    emit sigInitStatus( "Fixing Rules:", numRules );
    int numChanged = 0;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;

        auto rule = getRule( fRules->Item( ii ) );
        if ( !rule )
            continue;

        emit sigIncStatusValue( "Fixing Rules:" );

        auto conditions = rule->Conditions();
        if ( !conditions )
            continue;

        auto from = conditions->From();
        if ( !from->Enabled() )
            continue;

        auto fromEmails = getRecipientEmails( from->Recipients(), {}, true );
        if ( fromEmails.isEmpty() )
            continue;

        QStringList msgs;
        if ( !addRecipientsToRule( rule.get(), fromEmails, msgs ) )
            return false;

        from->SetEnabled( false );
        numChanged++;
    }
    if ( numChanged )
        saveRules();
    QMessageBox::information( fParentWidget, R"(Move "From" to "Address")", QString( "%1 rules modified" ).arg( numChanged ) );
    return numChanged;
}

bool COutlookAPI::mergeRules()
{
    if ( !fRules )
        return false;

    auto numRules = fRules->Count();
    emit sigInitStatus( "Merging Rules:", numRules );
    std::map< QString, std::shared_ptr< Outlook::Rule > > rules;
    std::list< int > toRemove;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        emit sigIncStatusValue( "Merging Rules:" );
        if ( canceled() )
            return false;

        auto rule = getRule( fRules->Item( ii ) );
        if ( !rule || !rule->Enabled() )
            continue;

        auto from = rule->Conditions()->SenderAddress();
        if ( !from || !from->Enabled() )
            continue;

        auto moveAction = rule->Actions()->MoveToFolder();
        if ( !moveAction || !moveAction->Enabled() )
            continue;

        auto key = moveAction->Folder()->FullFolderPath();
        auto pos = rules.find( key );
        if ( pos == rules.end() )
        {
            rules[ key ] = rule;
        }
        else
        {
            auto mergedRecipients = mergeRecipients( ( *pos ).second.get(), rule.get(), nullptr );
            if ( !mergedRecipients.has_value() )
                continue;

            rule->SetEnabled( false );
            ( *pos ).second->Conditions()->SenderAddress()->SetAddress( mergedRecipients.value() );
            toRemove.push_front( ii );
        }
    }
    if ( canceled() )
        return false;
    auto numChanged = toRemove.size();
    for ( auto &&ii : toRemove )
    {
        if ( canceled() )
            return false;
        fRules->Remove( ii );
    }

    if ( !toRemove.empty() )
        saveRules();

    QMessageBox::information( fParentWidget, R"(Merge Rules by Target Folder)", QString( "%1 rules deleted" ).arg( numChanged ) );

    return !toRemove.empty();
}

std::shared_ptr< Outlook::Application > COutlookAPI::getApplication()
{
    if ( !fOutlookApp )
        fOutlookApp = connectToException( std::make_shared< Outlook::Application >() );
    return fOutlookApp;
}

std::shared_ptr< Outlook::Account > COutlookAPI::getAccount( Outlook::_Account *item )
{
    if ( !item )
        return {};

    return connectToException( std::make_shared< Outlook::Account >( item ) );
}

std::shared_ptr< Outlook::MailItem > COutlookAPI::getMailItem( IDispatch *item )
{
    if ( !item )
        return {};
    return connectToException( std::make_shared< Outlook::MailItem >( item ) );
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getMailFolder( const Outlook::Folder *item )
{
    if ( !item )
        return {};
    return connectToException( std::shared_ptr< Outlook::Folder >( const_cast< Outlook::Folder * >( item ) ) );
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getMailFolder( const Outlook::MAPIFolder *item )
{
    if ( !item )
        return {};
    return getMailFolder( reinterpret_cast< const Outlook::Folder * >( item ) );
}

std::shared_ptr< Outlook::Items > COutlookAPI::getItems( Outlook::_Items *item )
{
    if ( !item )
        return {};
    return connectToException( std::make_shared< Outlook::Items >( item ) );
}

std::shared_ptr< Outlook::Rules > COutlookAPI::getRules( Outlook::Rules *item )
{
    if ( !item )
        return {};
    return connectToException( std::shared_ptr< Outlook::Rules >( item ) );
}

std::shared_ptr< Outlook::Rule > COutlookAPI::getRule( Outlook::_Rule *item )
{
    if ( !item )
        return {};
    return connectToException( std::make_shared< Outlook::Rule >( item ) );
}

void COutlookAPI::setOnlyProcessUnread( bool value, bool update )
{
    fOnlyProcessUnread = value;
    QSettings settings;
    settings.setValue( "OnlyProcessUnread", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setProcessAllEmailWhenLessThan200Emails( bool value, bool update )
{
    fProcessAllEmailWhenLessThan200Emails = value;
    QSettings settings;
    settings.setValue( "ProcessAllEmailWhenLessThan200Emails", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setRootFolder( const std::shared_ptr< Outlook::Folder > &folder, bool update )
{
    fRootFolder = folder;

    QSettings settings;
    if ( folder )
        settings.setValue( "RootFolder", getFolderPath( folder ) );
    else
        settings.remove( "RootFolder" );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setRootFolder( const QString &folderPath, bool update )
{
    if ( !accountSelected() )
        return;

    auto folder = getMailFolder( "Folder", folderPath, false ).first;
    setRootFolder( folder, update );
}
