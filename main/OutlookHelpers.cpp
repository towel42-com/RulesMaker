#include "OutlookHelpers.h"
#include <QInputDialog>
#include <QMessageBox>
#include <QDebug>

#include <QMetaProperty>
#include <QAxObject>

#include <iostream>
#include <oaidl.h>
#include "MSOUTL.h"

std::shared_ptr< COutlookHelpers > COutlookHelpers::sInstance;

COutlookHelpers::COutlookHelpers() :
    fOutlook( std::make_shared< Outlook::Application >() )
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
    //if ( fOutlook && !fOutlook->isNull() )
    //    Outlook::NameSpace( fOutlook->Session() ).Logoff();
}

std::shared_ptr< Outlook::Account > COutlookHelpers::selectAccount( QWidget *parent )
{
    fAccount.reset();
    if ( fOutlook->isNull() )
        return {};

    Outlook::NameSpace session( fOutlook->Session() );
    session.Logon();

    std::list< std::shared_ptr< Outlook::Account > > allAccounts;

    auto accounts = session.Accounts();
    if ( !accounts )
        return {};

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
        return {};
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
        return {};
    auto pos = accountMap.find( item );
    if ( pos == accountMap.end() )
        return {};
    fAccount = ( *pos ).second;
    fInbox.reset();
    fContacts.reset();
    return fAccount;
}

std::shared_ptr< Outlook::MAPIFolder > COutlookHelpers::getContacts( QWidget *parent )
{
    if ( !fContacts )
        fContacts = selectContacts( parent, false ).first;
    return fContacts;
}

std::pair< std::shared_ptr< Outlook::MAPIFolder >, bool > COutlookHelpers::selectContacts( QWidget *parent, bool singleOnly )
{
    if ( !fAccount )
    {
        if ( !selectAccount( parent ) )
            return {};
    }

    return selectFolder(
        parent, "Contact",
        []( std::shared_ptr< Outlook::MAPIFolder > folder )
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
        singleOnly );
}

std::shared_ptr< Outlook::MAPIFolder > COutlookHelpers::getInbox( QWidget *parent )
{
    if ( !fInbox )
        fInbox = selectInbox( parent, false ).first;
    return fInbox;
}

std::pair< std::shared_ptr< Outlook::MAPIFolder >, bool > COutlookHelpers::selectInbox( QWidget *parent, bool singleOnly )
{
    if ( !fAccount )
    {
        if ( !selectAccount( parent ) )
            return {};
    }

    return selectFolder(
        parent, "Inbox",
        []( std::shared_ptr< Outlook::MAPIFolder > folder )
        {
            if ( !folder )
                return false;
            if ( folder->DefaultItemType() != Outlook::OlItemType::olMailItem )
                return false;
            return ( folder->Name() == "Inbox" );
        },
        singleOnly );
}

std::pair< std::shared_ptr< Outlook::MAPIFolder >, bool > COutlookHelpers::selectFolder( QWidget *parent, const QString &folderName, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder, bool singleOnly )
{
    auto folders = getFolders( acceptFolder );
    return selectFolder( parent, folderName, folders, singleOnly );
}

std::pair< std::shared_ptr< Outlook::MAPIFolder >, bool > COutlookHelpers::selectFolder( QWidget *parent, const QString &folderName, const std::list< std::shared_ptr< Outlook::MAPIFolder > > &folders, bool singleOnly )
{
    if ( folders.empty() )
    {
        QMessageBox::critical( parent, QString( "Could not find %1" ).arg( folderName.toLower() ), folderName + " not found" );
        return { nullptr, false };
    }
    if ( folders.size() == 1 )
        return { std::shared_ptr< Outlook::MAPIFolder >( folders.front() ), false };
    if ( singleOnly )
        return { nullptr, false };

    QStringList folderNames;
    std::map< QString, std::shared_ptr< Outlook::MAPIFolder > > folderMap;

    for ( auto &&ii : folders )
    {
        auto path = ii->FolderPath();
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

std::list< std::shared_ptr< Outlook::MAPIFolder > > COutlookHelpers::getFolders( std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder )
{
    if ( fOutlook->isNull() )
        return {};
    if ( fAccount->isNull() )
        return {};

    //Outlook::NameSpace session( fOutlook->Session() );
    //session.Logon();

    //std::list< std::shared_ptr< Outlook::MAPIFolder > > retVal;
    //auto stores = session.Stores();
    //auto numStores = stores->Count();
    //for ( auto ii = 1; ii <= numStores; ++ii )
    //{
    auto store = fAccount->DeliveryStore();
    if ( !store )
        return {};

    auto root = std::shared_ptr< Outlook::MAPIFolder >( store->GetRootFolder() );
    auto retVal = getFolders( root, false, acceptFolder );
    //retVal.insert( retVal.end(), currFolders.begin(), currFolders.end() );

    return retVal;
}

std::list< std::shared_ptr< Outlook::MAPIFolder > > COutlookHelpers::getFolders( std::shared_ptr< Outlook::MAPIFolder > parent, bool recursive, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder )
{
    if ( !parent )
        return {};

    std::list< std::shared_ptr< Outlook::MAPIFolder > > retVal;

    auto folders = parent->Folders();
    auto folderCount = folders->Count();
    for ( auto jj = 1; jj < folderCount; ++jj )
    {
        auto folder = std::shared_ptr< Outlook::MAPIFolder >( folders->Item( jj ) );

        if ( acceptFolder && !acceptFolder( folder ) )
            continue;

        retVal.push_back( folder );
        if ( recursive )
        {
            auto subFolders = getFolders( folder, true, acceptFolder );
            retVal.insert( retVal.end(), subFolders.begin(), subFolders.end() );
        }
    }
    return retVal;
}

bool hasProperty( QObject *item, const char *propName )
{
    auto mo = item->metaObject();
    //auto aoxMO = dynamic_cast< c * >( mo );
    auto idx = mo->indexOfProperty( propName );
    return idx != -1;
    do
    {
        for ( int ii = mo->propertyOffset(); ii < mo->propertyCount(); ++ii )
        {
            if ( propName == mo->property( ii ).name() )
                return true;
        }
    }
    while ( ( mo = mo->superClass() ) );
    return false;
}

Outlook::OlObjectClass COutlookHelpers::getObjectClass( IDispatch *item )
{
    if ( !item )
        return {};

    auto tmp = new QAxObject( item );
    bool t2 = hasProperty( tmp, "Class" );
    auto retVal = tmp->property( "Class" );
    delete tmp;

    return static_cast< Outlook::OlObjectClass >( retVal.toInt() );
}

QString COutlookHelpers::getSenderEmailAddress( Outlook::MailItem *mailItem )
{
    if ( !mailItem )
        return {};

    auto email = getEmailAddress( mailItem->Sender() );
    if ( email.has_value() )
        return email.value();

    bool hasProperty = ::hasProperty( mailItem, "SenderEmailAddress" );
    if ( !hasProperty )
        return {};

    auto retVal = mailItem->SenderEmailAddress();
    return retVal;
}

bool COutlookHelpers::isExchangeUser( Outlook::AddressEntry *address )
{
    if ( !address )
        return false;

    auto user = address->GetExchangeUser();
    return user != nullptr;
}

std::optional< QString > COutlookHelpers::getEmailAddress( Outlook::AddressEntry *address )
{
    if ( !isExchangeUser( address ) )
        return {};

    auto user = address->GetExchangeUser();
    if ( user )
    {
        return user->PrimarySmtpAddress();
    }
    return {};
}

QString COutlookHelpers::getEmailAddress( Outlook::Recipient *recipient )
{
    if ( !recipient )
        return {};

    auto exchUserEmail = getEmailAddress( recipient->AddressEntry() );
    if ( exchUserEmail.has_value() )
        return exchUserEmail.value();
    return recipient->Address();
}

QStringList COutlookHelpers::getRecipients( Outlook::MailItem *mailItem, Outlook::OlMailRecipientType recipientType )
{
    if ( !mailItem )
        return {};
    auto recipients = mailItem->Recipients();
    if ( !recipients )
        return {};

    auto numRecipients = recipients->Count();
    QStringList retVal;
    for ( int ii = 1; ii <= numRecipients; ++ii )
    {
        auto recipient = recipients->Item( ii );
        if ( !recipient )
            continue;
        if ( recipientType != static_cast< Outlook::OlMailRecipientType >( recipient->Type() ) )
            continue;

        retVal << getEmailAddress( recipient );
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
        qDebug() << root->FolderPath();
        dumpFolder( root );
    }
}

QString COutlookHelpers::toString( Outlook::OlItemType olItemType )
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

void COutlookHelpers::dumpFolder( Outlook::MAPIFolder *parent )
{
    if ( !parent )
        return;

    auto folders = parent->Folders();
    auto folderCount = folders->Count();
    for ( auto jj = 1; jj < folderCount; ++jj )
    {
        auto folder = folders->Item( jj );
        qDebug() << folder->FolderPath() << toString( folder->DefaultItemType() );
        dumpFolder( folder );
    }
}
