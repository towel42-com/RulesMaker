#include "OutlookAPI.h"

#include <QInputDialog>
#include <QMessageBox>
#include <QDebug>
#include <QSettings>

#include "MSOUTL.h"

std::shared_ptr< Outlook::Folder > COutlookAPI::rootFolder()
{
    if ( !accountSelected() )
        return fInbox;

    if ( fRootFolder )
        return fRootFolder;

    return getInbox();
}

QString COutlookAPI::rootFolderName()
{
    return folderDisplayPath( rootFolder() );
}

void COutlookAPI::setRootFolder( const std::shared_ptr< Outlook::Folder > &folder, bool update )
{
    fRootFolder = folder;

    QSettings settings;
    if ( folder )
        settings.setValue( "RootFolder", folderDisplayPath( folder ) );
    else
        settings.remove( "RootFolder" );
    if ( update )
        emit sigOptionChanged();
}

std::shared_ptr< Outlook::Folder > COutlookAPI::findFolder( const QString &folderName, std::shared_ptr< Outlook::Folder > parentFolder )
{
    if ( !accountSelected() )
        return {};

    if ( !parentFolder )
        parentFolder = rootFolder();

    if ( !parentFolder )
        return {};

    auto folders = parentFolder->Folders();
    auto folderCount = folders->Count();

    for ( int ii = 1; ii <= folderCount; ++ii )
    {
        if ( canceled() )
            break;
        auto subFolder = folders->Item( ii );
        if ( !subFolder )
            continue;
        auto name = subFolder->FullFolderPath();
        if ( folderName == name )
            return getFolder( subFolder );
    }
    return {};
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getFolder( const Outlook::Folder *item )
{
    if ( !item )
        return {};
    return connectToException( std::shared_ptr< Outlook::Folder >( const_cast< Outlook::Folder * >( item ) ) );
}

std::list< std::shared_ptr< Outlook::Folder > > COutlookAPI::getFolders( const std::shared_ptr< Outlook::Folder > &parent, bool updateStatus, bool recursive, const TFolderFunc &acceptFolder, const TFolderFunc &checkChildFolders )
{
    if ( !parent )
        return {};

    std::list< std::shared_ptr< Outlook::Folder > > retVal;

    auto folders = parent->Folders();
    auto folderCount = folders->Count();
    QString msg;
    if ( parent )
        msg = QString( "Finding Sub-Folders of %1:" ).arg( parent->FolderPath() );
    else
        msg = QString( "Finding Folders:" );
    if ( updateStatus )
        emit sigInitStatus( msg, folderCount );
    for ( auto jj = 1; jj <= folderCount; ++jj )
    {
        if ( canceled() )
            break;

        if ( updateStatus )
        {
            emit sigIncStatusValue( msg );
        }
        
        auto folder = getFolder( folders->Item( jj ) );

        bool isMatch = !acceptFolder || ( acceptFolder && acceptFolder( folder ) );
        bool cont = recursive && ( !checkChildFolders || ( checkChildFolders && checkChildFolders( folder ) ) );

        if ( isMatch )
            retVal.push_back( folder );
        if ( cont )
        {
            auto &&subFolders = getFolders( folder, false, true, acceptFolder );
            retVal.insert( retVal.end(), subFolders.begin(), subFolders.end() );
        }
    }
    if ( updateStatus )
        emit sigStatusFinished( msg );

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

std::shared_ptr< Outlook::Folder > COutlookAPI::addFolder( const std::shared_ptr< Outlook::Folder > &parent, const QString &folderName )
{
    if ( !parent )
        return {};
    auto folder = parent->Folders()->Add( folderName );
    if ( !folder )
        return {};
    return getFolder( folder );
}

std::shared_ptr< Outlook::Folder > COutlookAPI::parentFolder( const std::shared_ptr< Outlook::Folder > &folder )
{
    if ( !folder )
        return {};

    auto parentObj = folder->Parent();
    if ( !parentObj )
        return {};

    auto parentFolder = new Outlook::Folder( parentObj );
    if ( parentFolder->Class() != Outlook::OlObjectClass::olFolder )
    {
        delete parentFolder;
        return {};
    }

    return getFolder( parentFolder );
}

QString COutlookAPI::nameForFolder( const std::shared_ptr< Outlook::Folder > &folder ) const
{
    if ( !folder )
        return {};
    return folder->Name();
}

QString COutlookAPI::rawPathForFolder( const std::shared_ptr< Outlook::Folder > &folder ) const
{
    if ( !folder )
        return {};
    return folder->FullFolderPath();
}

QString COutlookAPI::folderDisplayName( const std::shared_ptr< Outlook::Folder > &folder )
{
    return folderDisplayName( folder.get() );
}

bool COutlookAPI::isFolder( const std::shared_ptr< Outlook::Folder > &folder, const QString &path ) const
{
    return ( folderDisplayPath( folder, true ) == path ) || ( folderDisplayPath( folder, false ) == path );
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getFolder( const Outlook::MAPIFolder *item )
{
    if ( !item )
        return {};
    return getFolder( reinterpret_cast< const Outlook::Folder * >( item ) );
}

void COutlookAPI::setRootFolder( const QString &folderPath, bool update )
{
    if ( !accountSelected() )
        return;

    auto folder = getMailFolder( "Folder", folderPath, false ).first;
    setRootFolder( folder, update );
}

QString COutlookAPI::folderDisplayPath( const std::shared_ptr< Outlook::Folder > &folder, bool removeLeadingSlashes ) const
{
    if ( !folder )
        return {};

    auto retVal = folder->FullFolderPath();
    retVal = retVal.replace( "%2F", "/" );

    auto accountName = this->accountName();
    auto pos = retVal.indexOf( accountName + R"(\)" );
    if ( pos != -1 )
        retVal = retVal.remove( pos, accountName.length() + 1 );

    while ( removeLeadingSlashes && retVal.startsWith( R"(\)" ) )
        retVal = retVal.mid( 1 );
    return retVal;
}

QString COutlookAPI::folderDisplayName( const Outlook::Folder *folder )
{
    if ( !folder )
        return {};
    auto retVal = folder->Name();
    retVal = retVal.replace( "%2F", "/" );
    return retVal;
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getDefaultFolder( Outlook::OlDefaultFolders folderType )
{
    if ( !accountSelected() )
        return {};

    auto store = fAccount->DeliveryStore();
    if ( !store )
        return {};

    return getFolder( store->GetDefaultFolder( folderType ) );
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

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookAPI::selectFolder( const QString &folderName, const TFolderFunc &acceptFolder, const TFolderFunc &checkChildFolders, bool singleOnly )
{
    auto &&folders = getFolders( false, false, acceptFolder, checkChildFolders );
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
    auto item = QInputDialog::getItem( fParentWidget, QString( "Select %1 Folder" ).arg( folderName ), folderName + " Folder", folderNames, 0, false, &aOK );
    if ( !aOK )
        return { {}, false };
    auto pos = folderMap.find( item );
    if ( pos == folderMap.end() )
        return { {}, false };
    return { ( *pos ).second, true };
}

std::list< std::shared_ptr< Outlook::Folder > > COutlookAPI::getFolders( bool updateStatus, bool recursive, const TFolderFunc &acceptFolder, const TFolderFunc &checkChildFolders )
{
    if ( !selectAccount( true ) )
        return {};

    auto store = fAccount->DeliveryStore();
    if ( !store )
        return {};

    auto root = getFolder( store->GetRootFolder() );
    auto retVal = getFolders( root, updateStatus, recursive, acceptFolder, checkChildFolders );

    return retVal;
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

int COutlookAPI::recursiveSubFolderCount( const Outlook::Folder *parent )
{
    if ( !parent )
        return 0;

    emit sigInitStatus( "Counting Folders", 0 );
    auto retVal = subFolderCount( parent, true );
    emit sigStatusFinished( "Counting Folders" );
    return retVal;
}

int COutlookAPI::subFolderCount( const Outlook::Folder *parent, bool recursive )
{
    if ( !parent )
        return 0;

    emit sigInitStatus( "Counting Folders", 0 );

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
        emit sigStatusFinished( "Counting Folders" );

    return retVal;
}

bool COutlookAPI::emptyTrash()
{
    slotClearCanceled();
    fIgnoreExceptions = true;
    auto trash = getTrashFolder();
    auto retVal = emptyFolder( trash );
    fIgnoreExceptions = false;
    return retVal;
}

bool COutlookAPI::emptyJunk()
{
    slotClearCanceled();
    //fIgnoreExceptions = true;
    auto junk = getJunkFolder();
    auto retVal = emptyFolder( junk );
    fIgnoreExceptions = false;
    return retVal;
}

bool COutlookAPI::emptyFolder( std::shared_ptr< Outlook::Folder > &folder )
{
    if ( !folder )
        return false;

    auto subFolders = folder->Folders();
    auto msg = tr( "Emptying Folder - %1" ).arg( folder->Name() );
    emit sigStatusMessage( msg );
    int numFoldersDeleted = 0;
    if ( subFolders && subFolders->Count() )
    {
        auto msg = tr( "Emptying Folder - %1 - Deleting Sub-Folders" ).arg( folder->Name() );
        auto count = subFolders->Count();
        emit sigInitStatus( msg, count );
        int itemNum = 1;
        while ( subFolders->Count() && ( itemNum <= subFolders->Count() ) )
        {
            if ( canceled() )
                break;
            auto subFolder = subFolders->Item( itemNum );
            if ( !subFolder )
                continue;
            auto subFolderName = subFolder->Name();
            emit sigStatusMessage( tr( "Deleting Folder - %1" ).arg( subFolderName ) );
            emit sigIncStatusValue( msg );
            subFolder->Delete();
            numFoldersDeleted++;
        }
    }
    emit sigStatusMessage( QString( "%1 folders deleted" ).arg( numFoldersDeleted ) );
    auto items = getItems( folder->Items() );
    int numItemsDeleted = 0;
    int numItemsSkipped = 0;
    if ( items && items->Count() )
    {
        auto count = items->Count();
        auto msg = tr( "Emptying Folder - %1 - Deleting items" ).arg( folder->Name() );
        emit sigInitStatus( msg, count );
        int itemNum = 1;
        while ( items->Count() && ( itemNum <= items->Count() ) ) 
        {
            if ( canceled() )
                break;
            auto item = getItem( items, itemNum );
            auto && [ aOK, desc ] = canDeleteItem( item );
            if ( !aOK )
            {
                itemNum++;
                numItemsSkipped++;
                continue;
            }
            emit sigStatusMessage( tr( "Deleting item - %1" ).arg( desc ) );
            emit sigIncStatusValue( msg );
            deleteItem( item );
            numItemsDeleted++;
        }
    }
    emit sigStatusMessage( QString( "%1 items deleted, %2 items skipped" ).arg( numItemsDeleted ).arg( numItemsSkipped ) );
    emit sigStatusFinished( msg );
    return !canceled();
}

void COutlookAPI::displayFolder( const std::shared_ptr< Outlook::Folder > &folder ) const
{
    if ( !folder )
        return;
    folder->Display();
}
