#include "Wrappers.h"
#include "MSOUTL.h"

#include <QDebug>

namespace NWrappers
{
    std::shared_ptr< Outlook::Application > fApplication;
    std::shared_ptr< Outlook::Application > getApplication()
    {
        if ( !fApplication )
            fApplication = std::make_shared< Outlook::Application >();
        return fApplication;
    }
    void clearApplication()
    {
        fApplication.reset();
    }

    std::shared_ptr< Outlook::MailItem > getMailItem( IDispatch *item )
    {
        return std::make_shared< Outlook::MailItem >( item );
    }

    std::unordered_map< Outlook::Folder *, std::weak_ptr< Outlook::Folder > > sFolderMap;
    std::shared_ptr< Outlook::Folder > getFolder( Outlook::Folder *item )
    {
        auto pos = sFolderMap.find( item );
        if ( ( pos == sFolderMap.end() ) || ( *pos ).second.expired() )
        {
            auto retVal = std::shared_ptr< Outlook::Folder >( item );
            sFolderMap[ item ].operator=( retVal );
            return retVal;
        }
        else
            return ( *pos ).second.lock();
    }

    std::shared_ptr< Outlook::Folder > getFolder( Outlook::MAPIFolder *item )
    {
        return getFolder( reinterpret_cast< Outlook::Folder * >( item ) );
    }

    void clearFolderCache()
    {
        sFolderMap.clear();
    }

    std::unordered_map< Outlook::_Items *, Outlook::Items * > sItemsMap1;
    std::unordered_map< Outlook::Items *, std::weak_ptr< Outlook::Items > > sItemsMap2;

    std::shared_ptr< Outlook::Items > getItems( Outlook::Items *items )
    {
        auto pos = sItemsMap2.find( items );
        if ( ( pos == sItemsMap2.end() ) || ( ( *pos ).second.expired() ) )
        {
            auto retVal = std::shared_ptr< Outlook::Items >( items );
            sItemsMap2[ items ].operator=( retVal );
            return retVal;
        }
        else
            return ( *pos ).second.lock();
    }

    std::shared_ptr< Outlook::Items > getItems( Outlook::_Items *items )
    {
        auto pos = sItemsMap1.find( items );
        if ( pos == sItemsMap1.end() )
            return getItems( new Outlook::Items( items ) );
        else
            return getItems( ( *pos ).second );
    }

    void clearItemsCache()
    {
        sItemsMap1.clear();
        sItemsMap2.clear();
    }
}