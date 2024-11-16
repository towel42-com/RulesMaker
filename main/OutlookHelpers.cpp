#include "OutlookHelpers.h"
#include <QInputDialog>
#include <QMessageBox>
#include <QDebug>

#include <QMetaProperty>
#include <QAxObject>

#include <QVariant>
#include <iostream>
#include <oaidl.h>
#include "MSOUTL.h"
#include <QDebug>
std::shared_ptr< COutlookHelpers > COutlookHelpers::sInstance;

COutlookHelpers::COutlookHelpers() :
    fOutlookApp( std::make_shared< Outlook::Application >() )
{
}

std::shared_ptr< COutlookHelpers > COutlookHelpers::getInstance()
{
    if ( !sInstance )
        sInstance = std::make_shared< COutlookHelpers >();
    return sInstance;
}

COutlookHelpers::~COutlookHelpers()
{
    logout( false );
}

void COutlookHelpers::logout( bool andNotify )
{
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

bool COutlookHelpers::accountSelected() const
{
    return fAccount.operator bool();
}

std::shared_ptr< Outlook::Account > COutlookHelpers::selectAccount( bool notifyOnChange, QWidget *parent )
{
    logout( notifyOnChange );
    if ( fOutlookApp->isNull() )
        return {};

    Outlook::NameSpace session( fOutlookApp->Session() );
    session.Logon();
    fLoggedIn = true;

    std::list< std::shared_ptr< Outlook::Account > > allAccounts;

    auto accounts = session.Accounts();
    if ( !accounts )
    {
        logout( notifyOnChange );
        return {};
    }

    auto numAccounts = accounts->Count();
    for ( auto ii = 1; ii <= numAccounts; ++ii )
    {
        auto account = accounts->Item( ii );
        if ( !account )
            continue;

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

        allAccounts.push_back( std::shared_ptr< Outlook::Account >( new Outlook::Account( account ) ) );
    }

    if ( allAccounts.size() == 0 )
    {
        logout( notifyOnChange );
        return {};
    }
    if ( allAccounts.size() == 1 )
        fAccount = allAccounts.front();

    QStringList accountNames;
    std::map< QString, std::shared_ptr< Outlook::Account > > accountMap;

    for ( auto &&ii : allAccounts )
    {
        auto path = ii->DisplayName();
        accountNames << path;
        accountMap[ path ] = ii;
    }
    bool aOK{ false };
    auto item = QInputDialog::getItem( parent, QString( "Select Account:" ), "Account:", accountNames, 0, false, &aOK );
    if ( !aOK )
    {
        logout( notifyOnChange );
        return {};
    }
    auto pos = accountMap.find( item );
    if ( pos == accountMap.end() )
    {
        logout( notifyOnChange );
        return {};
    }
    fAccount = ( *pos ).second;
    if ( notifyOnChange )
        emit sigAccountChanged();
    return fAccount;
}

std::shared_ptr< Outlook::Folder > COutlookHelpers::getContacts( QWidget *parent )
{
    if ( !fContacts )
        fContacts = selectContacts( parent, false ).first;
    return fContacts;
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookHelpers::selectContacts( QWidget *parent, bool singleOnly )
{
    if ( !fAccount )
    {
        if ( !selectAccount( true, parent ) )
            return {};
    }

    return selectFolder(
        parent, "Contact",
        []( std::shared_ptr< Outlook::Folder > folder )
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

std::shared_ptr< Outlook::Folder > COutlookHelpers::getInbox( QWidget *parent )
{
    if ( !fInbox )
        fInbox = selectInbox( parent, false ).first;
    return fInbox;
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookHelpers::selectInbox( QWidget *parent, bool singleOnly )
{
    if ( !fAccount )
    {
        if ( !selectAccount( true, parent ) )
            return {};
    }

    return selectFolder(
        parent, "Inbox",
        []( std::shared_ptr< Outlook::Folder > folder )
        {
            if ( !folder )
                return false;
            if ( folder->DefaultItemType() != Outlook::OlItemType::olMailItem )
                return false;
            return ( folder->Name() == "Inbox" );
        },
        {}, singleOnly );
}

std::shared_ptr< Outlook::Rules > COutlookHelpers::getRules( QWidget *parent )
{
    if ( !fRules )
        fRules = selectRules( parent );
    return fRules;
}

std::shared_ptr< Outlook::Folder > COutlookHelpers::rootFolder()
{
    if ( fRootFolder )
        return fRootFolder;
    return fInbox;
}

std::shared_ptr< Outlook::Rules > COutlookHelpers::selectRules( QWidget *parent )
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
    return std::shared_ptr< Outlook::Rules >( rules );
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookHelpers::selectFolder( QWidget *parent, const QString &folderName, std::function< bool( std::shared_ptr< Outlook::Folder > folder ) > acceptFolder, std::function< bool( std::shared_ptr< Outlook::Folder > folder ) > checkChildFolders, bool singleOnly )
{
    auto folders = getFolders( false, acceptFolder, checkChildFolders );
    return selectFolder( parent, folderName, folders, singleOnly );
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookHelpers::selectFolder( QWidget *parent, const QString &folderName, const std::list< std::shared_ptr< Outlook::Folder > > &folders, bool singleOnly )
{
    if ( folders.empty() )
    {
        QMessageBox::critical( parent, QString( "Could not find %1" ).arg( folderName.toLower() ), folderName + " not found" );
        return { nullptr, false };
    }
    if ( folders.size() == 1 )
        return { std::shared_ptr< Outlook::Folder >( folders.front() ), false };
    if ( singleOnly )
        return { nullptr, false };

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
        return { nullptr, false };
    auto pos = folderMap.find( item );
    if ( pos == folderMap.end() )
        return { nullptr, false };
    return { ( *pos ).second, true };
}

std::list< std::shared_ptr< Outlook::Folder > > COutlookHelpers::getFolders( bool recursive, std::function< bool( std::shared_ptr< Outlook::Folder > folder ) > acceptFolder, std::function< bool( std::shared_ptr< Outlook::Folder > folder ) > checkChildFolders )
{
    if ( fOutlookApp->isNull() )
        return {};
    if ( fAccount->isNull() )
        return {};

    auto store = fAccount->DeliveryStore();
    if ( !store )
        return {};

    auto root = std::shared_ptr< Outlook::Folder >( reinterpret_cast< Outlook::Folder * >( store->GetRootFolder() ) );
    auto retVal = getFolders( root, recursive, acceptFolder, checkChildFolders );

    return retVal;
}

std::list< std::shared_ptr< Outlook::Folder > > COutlookHelpers::getFolders( std::shared_ptr< Outlook::Folder > parent, bool recursive, std::function< bool( std::shared_ptr< Outlook::Folder > folder ) > acceptFolder, std::function< bool( std::shared_ptr< Outlook::Folder > folder ) > checkChildFolders )
{
    if ( !parent )
        return {};

    std::list< std::shared_ptr< Outlook::Folder > > retVal;

    auto folders = parent->Folders();
    auto folderCount = folders->Count();
    for ( auto jj = 1; jj <= folderCount; ++jj )
    {
        auto folder = std::shared_ptr< Outlook::Folder >( reinterpret_cast< Outlook::Folder * >( folders->Item( jj ) ) );

        bool isMatch = !acceptFolder || ( acceptFolder && acceptFolder( folder ) );
        bool cont = recursive && ( !checkChildFolders || ( checkChildFolders && checkChildFolders( folder ) ) );

        if ( isMatch )
            retVal.push_back( folder );
        if ( cont )
        {
            auto subFolders = getFolders( folder, true, acceptFolder );
            retVal.insert( retVal.end(), subFolders.begin(), subFolders.end() );
        }
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

QString COutlookHelpers::ruleNameForFolder( std::shared_ptr< Outlook::Folder > folder )
{
    return ruleNameForFolder( folder.get() );
}

QString COutlookHelpers::ruleNameForFolder( Outlook::Folder *folder )
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

bool COutlookHelpers::addRule( const QString &destFolder, const QStringList &rules, QStringList &msgs )
{
    if ( destFolder.isEmpty() || rules.isEmpty() || !fRules )
    {
        msgs.push_back( "Parameters not set" );
        return false;
    }

    auto folders = getFolders(
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
        return false;
    }
    auto folder = folders.front();
    auto ruleName = ruleNameForFolder( folder );

    auto rule = fRules->Create( ruleName, Outlook::OlRuleType::olRuleReceive );
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

    if ( !addRecipientsToRule( rule, rules, msgs ) )
        return false;

    fRules->Save( true );
    return execute( rule );
}

bool COutlookHelpers::addToRule( std::shared_ptr< Outlook::Rule > rule, const QStringList &rules, QStringList &msgs )
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

bool COutlookHelpers::execute( std::shared_ptr< Outlook::Rule > rule )
{
    return execute( rule.get() );
}

bool COutlookHelpers::execute( Outlook::Rule *rule )
{
    if ( !fInbox )
        return false;

    auto folder = reinterpret_cast< Outlook::MAPIFolder * >( fInbox.get() );
    int typeId = qRegisterMetaType< Outlook::MAPIFolder * >( "MAPIFolder*", &folder );

    auto inboxPtr = fInbox.get();
    rule->Execute( true, QVariant( typeId, &inboxPtr ) );
    return true;
}

bool COutlookHelpers::addRecipientsToRule( Outlook::Rule *rule, const QStringList &recipients, QStringList &msgs )
{
    auto cond = rule->Conditions()->SenderAddress();
    if ( !cond )
    {
        msgs.push_back( QString( "Internal error" ) );
        return false;
    }

    QStringList addresses;
    if ( cond->Enabled() )
    {
        auto variant = cond->Address();
        if ( variant.type() == QVariant::Type::String )
            addresses << variant.toString();
        else if ( variant.type() == QVariant::Type::StringList )
            addresses << variant.toStringList();
        addresses << recipients;
    }
    else
        addresses = recipients;
    addresses.removeDuplicates();
    cond->SetAddress( addresses );
    cond->SetEnabled( true );

    return true;
}

bool hasProperty( QObject *item, const char *propName )
{
    auto mo = item->metaObject();
    //auto aoxMO = dynamic_cast< c * >( mo );
    auto idx = mo->indexOfProperty( propName );
    return idx != -1;
}

Outlook::OlObjectClass COutlookHelpers::getObjectClass( IDispatch *item )
{
    if ( !item )
        return {};

    auto retVal = QAxObject( item ).property( "Class" );

    return static_cast< Outlook::OlObjectClass >( retVal.toInt() );
}

QStringList COutlookHelpers::getSenderEmailAddresses( Outlook::MailItem *mailItem )
{
    if ( !mailItem )
        return {};

    auto email = getEmailAddresses( mailItem->Sender() );
    email << mailItem->SenderEmailAddress();
    email.removeDuplicates();
    return email;
}

QStringList COutlookHelpers::getEmailAddresses( Outlook::AddressEntry *address )
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

QStringList COutlookHelpers::getEmailAddresses( Outlook::AddressEntries *entries )
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

QString COutlookHelpers::getEmailAddress( Outlook::Recipient *recipient )
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

QStringList COutlookHelpers::getEmailAddresses( Outlook::AddressList *addresses )
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

QStringList COutlookHelpers::getRecipientEmails( Outlook::MailItem *mailItem, Outlook::OlMailRecipientType recipientType )
{
    if ( !mailItem )
        return {};
    auto recipients = mailItem->Recipients();
    if ( !recipients )
        return {};

    return getRecipientEmails( recipients, recipientType );
}

QStringList COutlookHelpers::getRecipientEmails( Outlook::Recipients *recipients, std::optional< Outlook::OlMailRecipientType > recipientType )
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

void COutlookHelpers::dumpSession( Outlook::NameSpace &session )
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

void COutlookHelpers::dumpFolder( Outlook::Folder *parent )
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

void COutlookHelpers::renameRules()
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

        auto actions = rule->Actions();
        if ( !actions )
            continue;

        auto action = actions->MoveToFolder();
        if ( !action )
            continue;

        auto folder = action->Folder();
        if ( !folder )
            continue;

        auto ruleName = ruleNameForFolder( reinterpret_cast< Outlook::Folder * >( folder ) );
        auto currName = rule->Name();
        if ( ruleName != currName )
        {
            changed = true;
            rule->SetName( ruleName );
        }
    }
    if ( changed )
        fRules->Save();
}

void COutlookHelpers::sortRules()
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

void COutlookHelpers::moveFromToAddress()
{
    if ( !fRules )
        return;

    auto numRules = fRules->Count();
    bool changed = false;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        auto rule = std::make_shared< Outlook::Rule >( fRules->Item( ii ) );
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
