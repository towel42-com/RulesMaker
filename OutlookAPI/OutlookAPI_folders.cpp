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

std::shared_ptr< Outlook::Folder > COutlookAPI::getFolder( const Outlook::Folder *item )
{
    if ( !item )
        return {};
    return connectToException( std::shared_ptr< Outlook::Folder >( const_cast< Outlook::Folder * >( item ) ) );
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

    auto root = getFolder( store->GetRootFolder() );
    auto retVal = getFolders( root, recursive, acceptFolder, checkChildFolders );

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

    emit sigInitStatus( "Counting Folders:", 0 );
    auto retVal = subFolderCount( parent, true );
    emit sigStatusFinished( "Counting Folders:" );
    return retVal;
}

int COutlookAPI::subFolderCount( const Outlook::Folder *parent, bool recursive )
{
    if ( !parent )
        return 0;

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