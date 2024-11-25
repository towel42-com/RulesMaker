#include "OutlookAPI.h"
#include <QInputDialog>
#include <QMessageBox>
#include <QDebug>

#include <QMetaProperty>
#include <QSettings>

#include <QVariant>
#include <iostream>
#include <oaidl.h>
#include "MSOUTL.h"
#include <QDebug>
std::shared_ptr< COutlookAPI > COutlookAPI::sInstance;

COutlookAPI::COutlookAPI() :
    fOutlookApp( NWrappers::getApplication() )
{
}

std::shared_ptr< COutlookAPI > COutlookAPI::getInstance()
{
    if ( !sInstance )
        sInstance = std::make_shared< COutlookAPI >();
    return sInstance;
}

COutlookAPI::~COutlookAPI()
{
    logout( false );
}

void COutlookAPI::logout( bool andNotify )
{
    NWrappers::clearApplication();

    fAccount.reset();
    fInbox.reset();
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

std::shared_ptr< Outlook::Account > COutlookAPI::selectAccount( bool notifyOnChange, QWidget *parent )
{
    logout( notifyOnChange );
    if ( fOutlookApp->isNull() )
        return {};

    Outlook::NameSpace session( fOutlookApp->Session() );
    session.Logon();
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

        auto account = NWrappers::getAccount( item );

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
    auto account = QInputDialog::getItem( parent, QString( "Select Account:" ), "Account:", accountNames, accountPos, false, &aOK );
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

std::shared_ptr< Outlook::Folder > COutlookAPI::getContacts( QWidget *parent )
{
    if ( !fContacts )
        fContacts = selectContacts( parent, false ).first;
    return fContacts;
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookAPI::selectContacts( QWidget *parent, bool singleOnly )
{
    if ( !fAccount )
    {
        if ( !selectAccount( true, parent ) )
            return {};
    }

    return selectFolder(
        parent, "Contact",
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

std::shared_ptr< Outlook::Folder > COutlookAPI::getInbox( QWidget *parent )
{
    if ( !fInbox )
        fInbox = selectInbox( parent, false ).first;
    return fInbox;
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookAPI::selectInbox( QWidget *parent, bool singleOnly )
{
    if ( !fAccount )
    {
        if ( !selectAccount( true, parent ) )
            return {};
    }

    return selectFolder(
        parent, "Inbox",
        []( const std::shared_ptr< Outlook::Folder > &folder )
        {
            if ( !folder )
                return false;
            if ( folder->DefaultItemType() != Outlook::OlItemType::olMailItem )
                return false;
            return ( folder->Name() == "Inbox" );
        },
        {}, singleOnly );
}

std::shared_ptr< Outlook::Rules > COutlookAPI::getRules( QWidget *parent )
{
    if ( !fRules )
        fRules = selectRules( parent );
    return fRules;
}

std::shared_ptr< Outlook::Folder > COutlookAPI::rootFolder()
{
    if ( fRootFolder )
        return fRootFolder;
    return fInbox;
}

std::shared_ptr< Outlook::Rules > COutlookAPI::selectRules( QWidget *parent )
{
    if ( !fAccount )
    {
        if ( !selectAccount( true, parent ) )
            return {};
    }

    if ( fOutlookApp->isNull() )
        return {};
    if ( fAccount->isNull() )
        return {};

    auto store = fAccount->DeliveryStore();
    if ( !store )
        return {};

    auto rules = store->GetRules();
    return NWrappers::getRules( rules );
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookAPI::selectFolder( QWidget *parent, const QString &folderName, std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) > acceptFolder, std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) > checkChildFolders, bool singleOnly )
{
    auto && folders = getFolders( false, acceptFolder, checkChildFolders );
    return selectFolder( parent, folderName, folders, singleOnly );
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookAPI::selectFolder( QWidget *parent, const QString &folderName, const std::list< std::shared_ptr< Outlook::Folder > > &folders, bool singleOnly )
{
    if ( folders.empty() )
    {
        QMessageBox::critical( parent, QString( "Could not find %1" ).arg( folderName.toLower() ), folderName + " not found" );
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
    auto item = QInputDialog::getItem( parent, QString( "Select %1 Folder" ).arg( folderName ), folderName + " Folder:", folderNames, 0, false, &aOK );
    if ( !aOK )
        return { {}, false };
    auto pos = folderMap.find( item );
    if ( pos == folderMap.end() )
        return { {}, false };
    return { ( *pos ).second, true };
}

std::list< std::shared_ptr< Outlook::Folder > > COutlookAPI::getFolders( bool recursive, std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) > acceptFolder, std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) > checkChildFolders )
{
    if ( fOutlookApp->isNull() )
        return {};
    if ( fAccount->isNull() )
        return {};

    auto store = fAccount->DeliveryStore();
    if ( !store )
        return {};

    auto root = NWrappers::getFolder( store->GetRootFolder() );
    auto retVal = getFolders( root, recursive, acceptFolder, checkChildFolders );

    return retVal;
}

std::list< std::shared_ptr< Outlook::Folder > > COutlookAPI::getFolders( const std::shared_ptr< Outlook::Folder > &parent, bool recursive, std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) > acceptFolder, std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) > checkChildFolders )
{
    if ( !parent )
        return {};

    std::list< std::shared_ptr< Outlook::Folder > > retVal;

    auto folders = parent->Folders();
    auto folderCount = folders->Count();
    for ( auto jj = 1; jj <= folderCount; ++jj )
    {
        auto folder = NWrappers::getFolder( folders->Item( jj ) );

        bool isMatch = !acceptFolder || ( acceptFolder && acceptFolder( folder ) );
        bool cont = recursive && ( !checkChildFolders || ( checkChildFolders && checkChildFolders( folder ) ) );

        if ( isMatch )
            retVal.push_back( folder );
        if ( cont )
        {
            auto && subFolders = getFolders( folder, true, acceptFolder );
            retVal.insert( retVal.end(), subFolders.begin(), subFolders.end() );
        }
    }
    for ( auto &&ii : retVal )
    {
        if ( !ii )
            int xyz = 0;
    }

    return retVal;
}

int COutlookAPI::subFolderCount( const std::shared_ptr< Outlook::Folder > &parent, std::function< bool( Outlook::Folder *folder ) > acceptFolder /*= {}*/ )
{
    if ( !parent )
        return 0;

    auto folders = parent->Folders();
    auto folderCount = folders->Count();
    if ( !acceptFolder )
        return folderCount;

    int retVal = 0;
    for ( auto jj = 1; jj <= folderCount; ++jj )
    {
        bool isMatch = !acceptFolder || ( acceptFolder && acceptFolder( reinterpret_cast< Outlook::Folder * >( folders->Item( jj ) ) ) );
        if ( isMatch )
            retVal++;
    }
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
    }
    else
    {
        pos = path.lastIndexOf( R"(\)" );

        if ( pos != -1 )
            ruleName = path.mid( pos + 1 );
        else
            ruleName = path;
    }
    return ruleName;
}

std::pair< std::shared_ptr< Outlook::Rule >, bool > COutlookAPI::addRule( const QString &destFolder, const QStringList &rules, QStringList &msgs )
{
    auto retVal = std::make_pair( std::shared_ptr< Outlook::Rule >(), false );
    if ( destFolder.isEmpty() || rules.isEmpty() || !fRules )
    {
        msgs.push_back( "Parameters not set" );
        return retVal;
    }

    auto && folders = getFolders(
        true,
        [ = ]( std::shared_ptr< Outlook::Folder > folder )
        {
            if ( !folder )
                return false;
            auto curr = folder->FullFolderPath();
            return ( curr == destFolder );
        },
        [ = ]( std::shared_ptr< Outlook::Folder > folder )
        {
            if ( !folder )
                return false;
            auto curr = folder->FullFolderPath();
            return destFolder.startsWith( curr );
        } );
    if ( folders.empty() )
    {
        msgs.push_back( QString( "Could not find folder '%1'" ).arg( destFolder ) );
        return retVal;
    }
    auto folder = folders.front();
    auto ruleName = ruleNameForFolder( folder );

    auto rulePtr = fRules->Create( ruleName, Outlook::OlRuleType::olRuleReceive );
    if ( !rulePtr )
    {
        msgs.push_back( QString( "Could not create rule '%1'" ).arg( ruleName ) );
        return retVal;
    }

    auto rule = NWrappers::getRule( rulePtr );
    auto moveAction = rule->Actions()->MoveToFolder();
    if ( !moveAction )
    {
        msgs.push_back( QString( "Internal error" ) );
        return retVal;
    }
    retVal.first = rule;
    moveAction->SetEnabled( true );
    moveAction->SetFolder( reinterpret_cast< Outlook::MAPIFolder * >( folder.get() ) );

    rule->Actions()->Stop()->SetEnabled( true );

    if ( !addRecipientsToRule( rule.get(), rules, msgs ) )
        return retVal;

    auto name = ruleNameForRule( rule );
    if ( name.has_value() && ( rule->Name() != name ) )
        rule->SetName( name.value() );

    fRules->Save( true );
    retVal.second = execute( rule );
    return retVal;
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

    fRules->Save( true );

    return execute( rule );
}

bool COutlookAPI::execute( std::shared_ptr< Outlook::Rule > rule )
{
    return execute( rule.get() );
}

bool COutlookAPI::execute( Outlook::Rule *rule )
{
    if ( !fInbox )
        return false;

    auto folder = reinterpret_cast< Outlook::MAPIFolder * >( fInbox.get() );
    int typeId = qRegisterMetaType< Outlook::MAPIFolder * >( "MAPIFolder*", &folder );

    auto inboxPtr = fInbox.get();
    rule->Execute( true, QVariant( typeId, &inboxPtr ) );
    return true;
}

bool COutlookAPI::addRecipientsToRule( Outlook::Rule *rule, const QStringList &recipients, QStringList &msgs )
{
    if ( !rule || !rule->Conditions() )
        return {};

    auto cond = rule->Conditions()->SenderAddress();
    if ( !cond )
    {
        msgs.push_back( QString( "Internal error" ) );
        return {};
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

    auto retVal = QAxObject( item ).property( "Class" );

    return static_cast< Outlook::OlObjectClass >( retVal.toInt() );
}

QStringList COutlookAPI::getSenderEmailAddresses( Outlook::MailItem *mailItem )
{
    if ( !mailItem )
        return {};

    auto email = getEmailAddresses( mailItem->Sender() );
    email << mailItem->SenderEmailAddress();
    email.removeDuplicates();
    return email;
}

QStringList COutlookAPI::getEmailAddresses( Outlook::AddressEntry *address )
{
    if ( !address )
        return {};
    auto type = address->AddressEntryUserType();
    QStringList retVal;
    switch ( type )
    {
        case Outlook::OlAddressEntryUserType::olExchangeAgentAddressEntry:
        case Outlook::OlAddressEntryUserType::olExchangeRemoteUserAddressEntry:
        case Outlook::OlAddressEntryUserType::olExchangeUserAddressEntry:
        case Outlook::OlAddressEntryUserType::olOutlookContactAddressEntry:
            {
                if ( address->GetExchangeUser() )
                {
                    retVal.push_back( address->GetExchangeUser()->PrimarySmtpAddress() );
                }
            }
            break;
        case Outlook::OlAddressEntryUserType::olOutlookDistributionListAddressEntry:
        case Outlook::OlAddressEntryUserType::olExchangeDistributionListAddressEntry:
            {
                auto list = address->GetExchangeDistributionList();
                if ( list )
                {
                    retVal << list->PrimarySmtpAddress();
                }
            }
            break;
        case Outlook::OlAddressEntryUserType::olSmtpAddressEntry:
            {
                retVal.push_back( address->Address() );
            }
        case Outlook::OlAddressEntryUserType::olExchangeOrganizationAddressEntry:
        case Outlook::OlAddressEntryUserType::olExchangePublicFolderAddressEntry:
        case Outlook::OlAddressEntryUserType::olLdapAddressEntry:
        case Outlook::OlAddressEntryUserType::olOtherAddressEntry:
            break;
    }

    retVal.removeAll( QString() );
    return retVal;
}

QStringList COutlookAPI::getEmailAddresses( Outlook::AddressEntries *entries )
{
    if ( !entries )
        return {};
    QStringList retVal;
    auto num = entries->Count();
    for ( int ii = 1; ii <= num; ++ii )
    {
        auto currItem = entries->Item( ii );
        if ( !currItem )
            continue;
        auto currEmails = getEmailAddresses( currItem );
        retVal << currEmails;
    }
    retVal.removeAll( QString() );
    return retVal;
}

QString COutlookAPI::getEmailAddress( Outlook::Recipient *recipient )
{
    if ( !recipient )
        return {};

    QString retVal;
    auto entryEmail = getEmailAddresses( recipient->AddressEntry() );
    if ( !entryEmail.isEmpty() )
        return entryEmail.front();
    else
        return recipient->Address();

    return retVal;
}

QStringList COutlookAPI::getEmailAddresses( Outlook::AddressList *addresses )
{
    if ( !addresses )
        return {};

    auto entries = addresses->AddressEntries();
    if ( !entries )
        return {};
    auto count = entries->Count();

    QStringList retVal;
    for ( int ii = 1; ii <= count; ++ii )
    {
        auto entry = entries->Item( ii );
        if ( !entry )
            continue;
        auto currEmails = getEmailAddresses( entry );
        retVal << currEmails;
    }
    return retVal;
}

QStringList COutlookAPI::getRecipientEmails( Outlook::MailItem *mailItem, Outlook::OlMailRecipientType recipientType )
{
    if ( !mailItem )
        return {};
    auto recipients = mailItem->Recipients();
    if ( !recipients )
        return {};

    return getRecipientEmails( recipients, recipientType );
}

QStringList COutlookAPI::getRecipientEmails( Outlook::Recipients *recipients, std::optional< Outlook::OlMailRecipientType > recipientType )
{
    if ( !recipients )
        return {};
    auto numRecipients = recipients->Count();
    QStringList retVal;
    for ( int ii = 1; ii <= numRecipients; ++ii )
    {
        auto recipient = recipients->Item( ii );
        if ( !recipient )
            continue;
        if ( recipientType.has_value() && ( recipientType.value() != static_cast< Outlook::OlMailRecipientType >( recipient->Type() ) ) )
            continue;

        auto addresses = getEmailAddresses( recipient->AddressEntry() );
        retVal << addresses;
    }
    return retVal;
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
    if ( variant.type() == QVariant::Type::String )
        retVal = variant.toString();
    else if ( variant.type() == QVariant::Type::StringList )
        retVal = variant.toStringList().join( joinSeparator );
    else
        int xyz = 0;
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

void COutlookAPI::renameRules()
{
    if ( !fRules )
        return;

    auto numRules = fRules->Count();
    bool changed = false;

    for ( int ii = 1; ii <= numRules; ++ii )
    {
        auto rule = fRules->Item( ii );
        if ( !rule )
            continue;

        auto ruleName = ruleNameForRule( NWrappers::getRule( rule ) );
        if ( !ruleName.has_value() )
            continue;

        auto currName = rule->Name();
        if ( ruleName.value() != currName )
        {
            changed = true;
            rule->SetName( ruleName.value() );
        }
    }
    if ( changed )
        fRules->Save();
}

std::optional< QString > COutlookAPI::ruleNameForRule( std::shared_ptr< Outlook::Rule > rule )
{
    if ( !rule->Enabled() )
        return {};

    auto actions = rule->Actions();
    if ( !actions )
        return {};

    auto action = actions->MoveToFolder();
    if ( !action || !action->Enabled() )
        return {};

    auto folder = action->Folder();
    if ( !folder )
        return {};

    auto ruleName = ruleNameForFolder( reinterpret_cast< Outlook::Folder * >( folder ) );
    if ( rule->Conditions() )
    {
        if ( rule->Conditions()->From() && rule->Conditions()->From()->Enabled() )
            ruleName += " (From)";

        if ( rule->Conditions()->SentTo() && rule->Conditions()->SentTo()->Enabled() )
            ruleName += " (SentTo)";
    }

    ruleName = ruleName.replace( "%2F", "/" );
    return ruleName;
}

void COutlookAPI::sortRules()
{
    if ( !fRules )
        return;

    std::list< Outlook::_Rule * > rules;
    auto numRules = fRules->Count();
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        auto rule = fRules->Item( ii );
        if ( !rule )
            continue;
        rules.push_back( rule );
    }
    rules.sort(
        []( Outlook::_Rule *lhs, Outlook::_Rule *rhs )
        {
            if ( !lhs || !rhs )
                return false;
            auto lhsName = lhs->Name();
            auto rhsName = rhs->Name();
            if ( lhsName.startsWith( rhsName ) && ( lhsName != rhsName ) )
                return true;
            else if ( rhsName.startsWith( lhsName ) && ( lhsName != rhsName ) )
                return false;
            else
                return lhsName < rhsName;
        } );
    bool changed = false;
    auto pos = 1;
    for ( auto &&ii : rules )
    {
        changed = changed || ( ii->ExecutionOrder() != pos );
        ii->SetExecutionOrder( pos++ );
    }
    if ( changed )
        fRules->Save();
}

void COutlookAPI::moveFromToAddress()
{
    if ( !fRules )
        return;

    auto numRules = fRules->Count();
    bool changed = false;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        auto rule = NWrappers::getRule( fRules->Item( ii ) );
        if ( !rule )
            continue;

        auto conditions = rule->Conditions();
        if ( !conditions )
            continue;

        auto from = conditions->From();
        if ( !from->Enabled() )
            continue;

        from->SetEnabled( false );
        changed = true;
        auto fromEmails = getRecipientEmails( from->Recipients(), {} );
        QStringList msgs;
        if ( !addRecipientsToRule( rule.get(), fromEmails, msgs ) )
            return;
    }
    if ( changed )
        fRules->Save();
}

void COutlookAPI::mergeRules()
{
    if ( !fRules )
        return;

    auto numRules = fRules->Count();
    bool changed = false;
    std::map< QString, std::shared_ptr< Outlook::Rule > > rules;
    std::list< int > toRemove;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        auto rule = NWrappers::getRule( fRules->Item( ii ) );
        if ( !rule || !rule->Enabled() )
            continue;

        auto from = rule->Conditions()->SenderAddress();
        if ( !from || !from->Enabled() )
            continue;

        auto moveAction = rule->Actions()->MoveToFolder();
        if ( !moveAction || !moveAction->Enabled() )
            continue;

        auto pos = rules.find( rule->Name() );
        if ( pos == rules.end() )
        {
            rules[ rule->Name() ] = rule;
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
    for ( auto &&ii : toRemove )
        fRules->Remove( ii );
}
