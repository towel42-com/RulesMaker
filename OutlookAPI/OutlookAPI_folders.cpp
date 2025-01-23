#include "OutlookAPI.h"
#include "ExceptionHandler.h"

#include <QInputDialog>
#include <QMessageBox>
#include <QDebug>
#include <QSettings>

#include "MSOUTL.h"

COutlookObj< Outlook::MAPIFolder > COutlookAPI::rootFolder()
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

void COutlookAPI::setRootFolder( const COutlookObj< Outlook::MAPIFolder > &folder, bool update )
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

COutlookObj< Outlook::MAPIFolder > COutlookAPI::findFolder( const QString &folderName, COutlookObj< Outlook::MAPIFolder > parentFolder )
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

COutlookObj< Outlook::MAPIFolder > COutlookAPI::getFolder( const Outlook::MAPIFolder *item )
{
    if ( !item )
        return {};
    return COutlookObj< Outlook::MAPIFolder >( const_cast< Outlook::MAPIFolder * >( item ) );
}

std::list< COutlookObj< Outlook::MAPIFolder > > COutlookAPI::getFolders( const COutlookObj< Outlook::MAPIFolder > &parent, bool recursive, const TFolderFunc &acceptFolder, const TFolderFunc &checkChildFolders )
{
    if ( !parent )
        return {};

    std::list< COutlookObj< Outlook::MAPIFolder > > retVal;

    auto folders = parent->Folders();
    auto folderCount = folders->Count();
    for ( auto jj = 1; jj <= folderCount; ++jj )
    {
        auto folder = getFolder( folders->Item( jj ) );

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
        []( const COutlookObj< Outlook::MAPIFolder > &lhs, const COutlookObj< Outlook::MAPIFolder > &rhs )
        {
            if ( !lhs )
                return false;
            if ( !rhs )
                return true;
            return lhs->FullFolderPath() < rhs->FullFolderPath();
        } );
    return retVal;
}

COutlookObj< Outlook::MAPIFolder > COutlookAPI::addFolder( const COutlookObj< Outlook::MAPIFolder > &parent, const QString &folderName )
{
    if ( !parent )
        return {};
    auto folder = parent->Folders()->Add( folderName );
    if ( !folder )
        return {};
    return getFolder( folder );
}

COutlookObj< Outlook::MAPIFolder > COutlookAPI::parentFolder( const COutlookObj< Outlook::MAPIFolder > &folder )
{
    if ( !folder )
        return {};

    auto parentObj = folder->Parent();
    if ( !parentObj )
        return {};

    auto parentFolder = new Outlook::MAPIFolder( parentObj );
    if ( parentFolder->Class() != Outlook::OlObjectClass::olFolder )
    {
        delete parentFolder;
        return {};
    }

    return getFolder( parentFolder );
}

QString COutlookAPI::nameForFolder( const COutlookObj< Outlook::MAPIFolder > &folder ) const
{
    if ( !folder )
        return {};
    return folder->Name();
}

QString COutlookAPI::rawPathForFolder( const COutlookObj< Outlook::MAPIFolder > &folder ) const
{
    if ( !folder )
        return {};
    return folder->FullFolderPath();
}

QString COutlookAPI::folderDisplayName( const COutlookObj< Outlook::MAPIFolder > &folder )
{
    return folderDisplayName( folder.get() );
}

bool COutlookAPI::isFolder( const COutlookObj< Outlook::MAPIFolder > &folder, const QString &path ) const
{
    return ( folderDisplayPath( folder, true ) == path ) || ( folderDisplayPath( folder, false ) == path );
}

void COutlookAPI::setRootFolder( const QString &folderPath, bool update )
{
    if ( !accountSelected() )
        return;

    auto folder = getMailFolder( "Folder", folderPath, false ).first;
    setRootFolder( folder, update );
}

QString COutlookAPI::folderDisplayPath( const COutlookObj< Outlook::MAPIFolder > &folder, bool removeLeadingSlashes ) const
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

QString COutlookAPI::folderDisplayName( const Outlook::MAPIFolder *folder )
{
    if ( !folder )
        return {};
    auto retVal = folder->Name();
    retVal = retVal.replace( "%2F", "/" );
    return retVal;
}

COutlookObj< Outlook::MAPIFolder > COutlookAPI::getDefaultFolder( Outlook::OlDefaultFolders folderType )
{
    if ( !accountSelected() )
        return {};

    auto store = fAccount->DeliveryStore();
    if ( !store )
        return {};

    return getFolder( store->GetDefaultFolder( folderType ) );
}

std::pair< COutlookObj< Outlook::MAPIFolder >, bool > COutlookAPI::getMailFolder( const QString &folderLabel, const QString &path, bool singleOnly )
{
    if ( !accountSelected() )
        return {};

    auto retVal = selectFolder(
        folderLabel,
        [ this, path ]( const COutlookObj< Outlook::MAPIFolder > &folder )
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

std::pair< COutlookObj< Outlook::MAPIFolder >, bool > COutlookAPI::selectFolder( const QString &folderName, const TFolderFunc &acceptFolder, const TFolderFunc &checkChildFolders, bool singleOnly )
{
    auto &&folders = getFolders( false, acceptFolder, checkChildFolders );
    return selectFolder( folderName, folders, singleOnly );
}

std::pair< COutlookObj< Outlook::MAPIFolder >, bool > COutlookAPI::selectFolder( const QString &folderName, const std::list< COutlookObj< Outlook::MAPIFolder > > &folders, bool singleOnly )
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
    std::map< QString, COutlookObj< Outlook::MAPIFolder > > folderMap;

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

std::list< COutlookObj< Outlook::MAPIFolder > > COutlookAPI::getFolders( bool recursive, const TFolderFunc &acceptFolder, const TFolderFunc &checkChildFolders )
{
    if ( !selectAccount( true ) )
        return {};

    auto store = fAccount->DeliveryStore();
    if ( !store )
        return {};

    auto root = getFolder( store->GetRootFolder() );
    auto retVal = getFolders( root, recursive, acceptFolder, checkChildFolders );

    return retVal;
}

QString COutlookAPI::ruleNameForFolder( const COutlookObj< Outlook::MAPIFolder > &folder )
{
    return ruleNameForFolder( folder.get() );
}

QString COutlookAPI::ruleNameForFolder( Outlook::MAPIFolder *folder )
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

int COutlookAPI::recursiveSubFolderCount( const Outlook::MAPIFolder *parent )
{
    if ( !parent )
        return 0;

    emit sigInitStatus( "Counting Folders:", 0 );
    auto retVal = subFolderCount( parent, true );
    emit sigStatusFinished( "Counting Folders:" );
    return retVal;
}

int COutlookAPI::subFolderCount( const Outlook::MAPIFolder *parent, bool recursive )
{
    if ( !parent )
        return 0;

    emit sigInitStatus( "Counting Folders:", 0 );

    auto folders = parent->Folders();
    if ( !folders )
        return 0;

    auto folderCount = folders->Count();

    int retVal = folderCount;
    for ( auto jj = 1; recursive && ( jj <= folderCount ); ++jj )
    {
        auto folder = reinterpret_cast< Outlook::MAPIFolder * >( folders->Item( jj ) );
        if ( !folder )
            continue;
        retVal += subFolderCount( folder, recursive );
    }

    if ( !recursive )
        emit sigStatusFinished( "Counting Folders:" );

    return retVal;
}

bool COutlookAPI::emptyTrash()
{
    slotClearCanceled();
    CExceptionHandler::instance()->setIgnoreExceptions( true );
    auto trash = getTrashFolder();
    auto retVal = emptyFolder( trash );
    CExceptionHandler::instance()->setIgnoreExceptions( false );
    return retVal;
}

bool COutlookAPI::emptyJunk()
{
    slotClearCanceled();
    CExceptionHandler::instance()->setIgnoreExceptions( true );
    auto junk = getJunkFolder();
    auto retVal = emptyFolder( junk );
    CExceptionHandler::instance()->setIgnoreExceptions( false );
    return retVal;
}

bool COutlookAPI::emptyFolder( const COutlookObj< Outlook::MAPIFolder > &folder )
{
    if ( !folder )
        return false;

    auto subFolders = folder->Folders();
    auto msg = tr( "Emptying Folder - %1:" ).arg( folder->Name() );
    emit sigStatusMessage( msg );
    int numFoldersDeleted = 0;
    if ( subFolders && subFolders->Count() )
    {
        auto msg = tr( "Emptying Folder - %1 - Deleting Sub-Folders:" ).arg( folder->Name() );
        emit sigInitStatus( msg, subFolders->Count() );
        int itemNum = 1;
        while ( !canceled() && ( itemNum <= subFolders->Count() ) )
        {
            auto subFolder = subFolders->Item( itemNum );
            if ( !subFolder )
            {
                emit sigIncStatusValue( msg );
                itemNum++;
                continue;
            }
            auto subFolderName = subFolder->Name();
            emit sigStatusMessage( tr( "Deleting Folder - %1" ).arg( subFolderName ) );
            emit sigIncStatusValue( msg );
            subFolder->Delete();
            numFoldersDeleted++;
        }
        emit sigStatusFinished( msg );
    }

    emit sigStatusMessage( QString( "%1 folders deleted" ).arg( numFoldersDeleted ) );
    auto items = getItems( folder->Items() );
    int numEmailsDeleted = 0;
    if ( items && items->Count() )
    {
        auto msg = tr( "Emptying Folder - %1 - Deleting Items:" ).arg( folder->Name() );
        emit sigInitStatus( msg, items->Count() );
        int itemNum = 1;
        while ( !canceled() && ( itemNum <= items->Count() ) )
        {
            auto item = items->Item( itemNum );
            if ( !item )
            {
                emit sigIncStatusValue( msg );
                itemNum++;
                continue;
            }
            COutlookObj< Outlook::MailItem > mailItem( item );
            if ( mailItem.deleteItem( [ msg, this ]() { emit sigIncStatusValue( msg ); } ) )
                numEmailsDeleted++;
            else
                itemNum++;
        }
        emit sigStatusFinished( msg );
    }
    emit sigStatusMessage( QString( "%1 emails deleted" ).arg( numEmailsDeleted ) );
    emit sigStatusFinished( msg );
    return !canceled();
}
