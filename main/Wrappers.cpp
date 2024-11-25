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

    std::shared_ptr< Outlook::Account > getAccount( Outlook::_Account *accountItem )
    {
        if ( !accountItem )
            return {};

        return std::make_shared< Outlook::Account >( accountItem );
    }

    void clearAccountCache()
    {
    }

    std::shared_ptr< Outlook::MailItem > getMailItem( IDispatch *item )
    {
        return std::make_shared< Outlook::MailItem >( item );
    }

    std::shared_ptr< Outlook::Folder > getFolder( Outlook::Folder *item )
    {
        auto retVal = std::shared_ptr< Outlook::Folder >( item );
        return retVal;
    }

    std::shared_ptr< Outlook::Folder > getFolder( Outlook::MAPIFolder *item )
    {
        return getFolder( reinterpret_cast< Outlook::Folder * >( item ) );
    }

    std::shared_ptr< Outlook::Items > getItems( Outlook::Items *items )
    {
        auto retVal = std::shared_ptr< Outlook::Items >( items );
        return retVal;
    }

    std::shared_ptr< Outlook::Items > getItems( Outlook::_Items *items )
    {
        return std::make_shared< Outlook::Items >( items );
    }

    std::shared_ptr< Outlook::Rules > getRules( Outlook::Rules *item )
    {
        auto retVal = std::shared_ptr< Outlook::Rules >( item );
        return retVal;
    }

    std::shared_ptr< Outlook::Rule > getRule( Outlook::_Rule *item )
    {
        auto retVal = std::make_shared< Outlook::Rule >( item );
        return retVal;
    }
}