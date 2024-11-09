#include "OutlookHelpers.h"
#include <QInputDialog>
#include <QMessageBox>
#include <QDebug>

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

std::shared_ptr< Outlook::MAPIFolder > COutlookHelpers::selectContactFolder( QWidget *parent )
{
    return selectFolder(
        parent, Outlook::OlItemType::olContactItem, "Contact",
        []( std::shared_ptr< Outlook::MAPIFolder > folder )
        {
            if ( !folder )
                return false;
            auto items = folder->Items();
            if ( items->Count() == 0 )
                return false;
            if ( folder->Name().contains( "meta", Qt::CaseSensitivity::CaseInsensitive ) )
                return false;
            return true;
        } );
}

std::shared_ptr< Outlook::MAPIFolder > COutlookHelpers::selectInboxFolder( QWidget *parent )
{
    return selectFolder(
        parent, Outlook::OlItemType::olMailItem, "Inbox",
        []( std::shared_ptr< Outlook::MAPIFolder > folder )
        {
            if ( !folder )
                return false;
            return ( folder->Name() == "Inbox" );
        } );
}

std::shared_ptr< Outlook::MAPIFolder > COutlookHelpers::selectFolder( QWidget *parent, const Outlook::OlItemType &itemType, const QString &folderName, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder )
{
    auto folders = getFolders( itemType, acceptFolder );
    return selectFolder( parent, folderName, folders );
}

std::shared_ptr< Outlook::MAPIFolder > COutlookHelpers::selectFolder( QWidget *parent, const QString &folderName, const std::list< std::shared_ptr< Outlook::MAPIFolder > > &folders )
{
    if ( folders.empty() )
    {
        QMessageBox::critical( parent, QString( "Could not find %1" ).arg( folderName.toLower() ), folderName + " not found" );
        return nullptr;
    }
    if ( folders.size() == 1 )
        return std::shared_ptr< Outlook::MAPIFolder >( folders.front() );
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
        return nullptr;
    auto pos = folderMap.find( item );
    if ( pos == folderMap.end() )
        return nullptr;
    return ( *pos ).second;
}

std::list< std::shared_ptr< Outlook::MAPIFolder > > COutlookHelpers::getFolders( Outlook::OlItemType itemType, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder )
{
    if ( fOutlook->isNull() )
        return {};
    Outlook::NameSpace session( fOutlook->Session() );
    session.Logon();

    std::list< std::shared_ptr< Outlook::MAPIFolder > > retVal;
    auto stores = session.Stores();
    auto numStores = stores->Count();
    for ( auto ii = 1; ii <= numStores; ++ii )
    {
        auto store = stores->Item( ii );
        if ( !store )
            continue;
        auto root = std::shared_ptr< Outlook::MAPIFolder >( store->GetRootFolder() );
        auto currFolders = getFolders( itemType, root, false, acceptFolder );
        retVal.insert( retVal.end(), currFolders.begin(), currFolders.end() );
    }
    return retVal;
}

std::list< std::shared_ptr< Outlook::MAPIFolder > > COutlookHelpers::getFolders( Outlook::OlItemType itemType, std::shared_ptr< Outlook::MAPIFolder > parent, bool recursive, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder )
{
    if ( !parent )
        return {};

    std::list< std::shared_ptr< Outlook::MAPIFolder > > retVal;

    auto folders = parent->Folders();
    auto folderCount = folders->Count();
    for ( auto jj = 1; jj < folderCount; ++jj )
    {
        auto folder = std::shared_ptr< Outlook::MAPIFolder >( folders->Item( jj ) );
        if ( folder->DefaultItemType() != itemType )
            continue;

        if ( acceptFolder && !acceptFolder( folder ) )
            continue;

        retVal.push_back( folder );
        if ( recursive )
        {
            auto subFolders = getFolders( itemType, folder, true, acceptFolder );
            retVal.insert( retVal.end(), subFolders.begin(), subFolders.end() );
        }
    }
    return retVal;
}

std::list< std::shared_ptr< Outlook::MAPIFolder > > COutlookHelpers::getFolders( std::shared_ptr< Outlook::MAPIFolder > parent, bool recursive, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder /*= {} */ )
{
    if ( !parent )
        return {};
    return getFolders( parent->DefaultItemType(), parent, recursive, acceptFolder );
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
