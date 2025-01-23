#ifndef OUTLOOKOBJ_H
#define OUTLOOKOBJ_H

#include "ExceptionHandler.h"

#include <QAxObject>
#include <cassert>
#include <utility>
#include <optional>
#include <type_traits>
#include <QString>
#include <memory>
#include <functional>

#include "OutlookLib/MSOUTL.h"

class QAxObject;
class COutlookAPI;
struct IDispatch;

template< typename T >
class has_delete
{
private:
    template< typename U >
    static auto check( U *u ) -> decltype( u->Delete(), std::true_type{} );
    template< typename >
    static std::false_type check( ... );

public:
    static constexpr bool value = decltype( check< T >( nullptr ) )::value;
};

template< typename T >
class use_name_desc
{
private:
    template< typename U >
    static auto check( U *u ) -> decltype( u->Name(), std::true_type{} );
    template< typename >
    static std::false_type check( ... );

public:
    static constexpr bool value = decltype( check< T >( nullptr ) )::value;
};

template< typename T >
class use_displayname_desc
{
private:
    template< typename U >
    static auto check( U *u ) -> decltype( u->DisplayName(), std::true_type{} );
    template< typename >
    static std::false_type check( ... );

public:
    static constexpr bool value = decltype( check< T >( nullptr ) )::value;
};

template< typename T >
class use_profile_name_desc
{
private:
    template< typename U >
    static auto check( U *u ) -> decltype( u->CurrentProfileName(), std::true_type{} );
    template< typename >
    static std::false_type check( ... );

public:
    static constexpr bool value = decltype( check< T >( nullptr ) )::value;
};

template< typename T >
class use_subject_desc
{
private:
    template< typename U >
    static auto check( U *u ) -> decltype( u->Subject(), std::true_type{} );
    template< typename >
    static std::false_type check( ... );

public:
    static constexpr bool value = decltype( check< T >( nullptr ) )::value;
};

template< typename T >
class use_caption_desc
{
private:
    template< typename U >
    static auto check( U *u ) -> decltype( u->Caption(), std::true_type{} );
    template< typename >
    static std::false_type check( ... );

public:
    static constexpr bool value = decltype( check< T >( nullptr ) )::value;
};

template< typename T >
class use_formula_desc
{
private:
    template< typename U >
    static auto check( U *u ) -> decltype( u->Formula(), std::true_type{} );
    template< typename >
    static std::false_type check( ... );

public:
    static constexpr bool value = decltype( check< T >( nullptr ) )::value;
};

template< typename T >
class use_filter_desc
{
private:
    template< typename U >
    static auto check( U *u ) -> decltype( u->Filter(), std::true_type{} );
    template< typename >
    static std::false_type check( ... );

public:
    static constexpr bool value = decltype( check< T >( nullptr ) )::value;
};

template< typename T >
class use_conversationid_desc
{
private:
    template< typename U >
    static auto check( U *u ) -> decltype( u->ConversationID(), std::true_type{} );
    template< typename >
    static std::false_type check( ... );

public:
    static constexpr bool value = decltype( check< T >( nullptr ) )::value;
};

template< typename T >
class use_conversationtopic_desc
{
private:
    template< typename U >
    static auto check( U *u ) -> decltype( u->ConversationTopic(), std::true_type{} );
    template< typename >
    static std::false_type check( ... );

public:
    static constexpr bool value = decltype( check< T >( nullptr ) )::value;
};

template< typename T, typename = typename std::enable_if< std::is_base_of< QAxObject, T >::value >::type >
using has_description = std::disjunction< use_name_desc< T >, use_displayname_desc< T >, use_profile_name_desc< T >, use_subject_desc< T >, use_caption_desc< T >, use_formula_desc< T >, use_filter_desc< T >, use_conversationtopic_desc< T >, use_conversationid_desc< T > >;

std::optional< Outlook::OlObjectClass > getObjectClass( IDispatch *item );
std::optional< Outlook::OlObjectClass > getObjectClass( QAxObject *item );
bool hasDelete( IDispatch *item );
bool deleteItem( IDispatch *item );
QString getDescription( IDispatch *item );

template< typename T, typename = typename std::enable_if< std::is_base_of< QAxObject, T >::value >::type >
class COutlookObj
{
public:
    COutlookObj() { initApplication(); }
    COutlookObj( const std::initializer_list< COutlookObj< T > > & ) { initApplication(); }

    COutlookObj( QAxObject *object ) { setObject( object ); }

    COutlookObj( IDispatch *baseItem ) { setObject( constructItem( baseItem ) ); }
    ~COutlookObj() { reset(); }

    operator bool() const { return isValid(); }

    bool isValid() const { return fClassType.has_value() && isValid( fObject.get() ); }

    T *get() const
    {
        if ( !fObject )
            return nullptr;
        auto retVal = dynamic_cast< T * >( fObject.get() );
        return retVal;
    }
    T *operator->() const { return get(); }

    void reset() { setObject( nullptr ); }

    template< typename U = T >
    std::enable_if_t< use_displayname_desc< U >::value, QString > getDescription() const
    {
        return get() ? get()->DisplayName() : QString();
    }
    template< typename U = T >
    std::enable_if_t< use_name_desc< U >::value && !use_displayname_desc< U >::value, QString > getDescription() const
    {
        return get() ? get()->Name() : QString();
    }
    template< typename U = T >
    std::enable_if_t< use_profile_name_desc< U >::value, QString > getDescription() const
    {
        return get() ? get()->CurrentProfileName() : QString();
    }
    template< typename U = T >
    std::enable_if_t< use_subject_desc< U >::value, QString > getDescription() const
    {
        return get() ? get()->Subject() : QString();
    }
    template< typename U = T >
    std::enable_if_t< use_caption_desc< U >::value, QString > getDescription() const
    {
        return get() ? get()->Caption() : QString();
    }
    template< typename U = T >
    std::enable_if_t< use_formula_desc< U >::value && !use_name_desc< U >::value, QString > getDescription() const
    {
        return get() ? get()->Formula() : QString();
    }
    template< typename U = T >
    std::enable_if_t< use_filter_desc< U >::value && !use_name_desc< U >::value, QString > getDescription() const
    {
        return get() ? get()->Filter() : QString();
    }
    template< typename U = T >
    std::enable_if_t< use_conversationtopic_desc< U >::value && !use_subject_desc< U >::value, QString > getDescription() const
    {
        return get() ? get()->ConversationTopic() : QString();
    }
    template< typename U = T >
    std::enable_if_t< use_conversationid_desc< U >::value && !use_conversationtopic_desc< U >::value && !use_subject_desc< U >::value, QString > getDescription() const
    {
        return get() ? get()->ConversationID() : QString();
    }

    template< typename U = T >
    std::enable_if_t< !has_description< U >::value, QString > getDescription() const
    {
        return {};
    }

    template< typename U = T >
    std::enable_if_t< has_delete< U >::value, bool > deleteItem( const std::function< void() > &onDelete = {} )
    {
        if ( !get() )
            return false;
        QString typeString;
        if ( fClassType.has_value() )
            typeString = toString( fClassType.value() );
        auto desc = getDescription();
        auto descText = ( QStringList() << typeString << desc ).join( " - " );

        COutlookAPI::instance()->sendStatusMesage( QObject::tr( "Deleting Item - %1" ).arg( descText ) );
        get()->Delete();
        if ( onDelete )
            onDelete();
        return true;
    }

    template< typename U = T >
    std::enable_if_t< !has_delete< U >::value, bool > deleteItem( const std::function< void() > & = {} )
    {
        return false;
    }

private:
    bool isValid( QAxObject *object ) const { return dynamic_cast< T * >( object ) != nullptr; }

    void connectToExceptionHandler()
    {
        if ( !fObject )
            return;
        CExceptionHandler::instance()->connectToException( get() );
    }

    void initApplication()
    {
        if constexpr ( std::is_same_v< T, Outlook::Application > )
        {
            if ( !isValid() )
            {
                setObject( new Outlook::Application() );
                Q_ASSERT( fClassType == Outlook::OlObjectClass::olApplication );
            }
        }
    }

    QAxObject *updateObject( QAxObject *origObject )
    {
        if ( !origObject )
            return nullptr;

        QAxObject *retVal = nullptr;
        if ( constexpr( std::is_same_v< T, Outlook::TimeZone > ) )
        {
            auto obj = dynamic_cast< Outlook::_TimeZone * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::TimeZone( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Application > ) )
        {
            auto obj = dynamic_cast< Outlook::_Application * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Application( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::NameSpace > ) )
        {
            auto obj = dynamic_cast< Outlook::_NameSpace * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::NameSpace( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::ContactItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_ContactItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::ContactItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::PropertyAccessor > ) )
        {
            auto obj = dynamic_cast< Outlook::_PropertyAccessor * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::PropertyAccessor( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Inspector > ) )
        {
            auto obj = dynamic_cast< Outlook::_Inspector * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Inspector( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::AttachmentSelection > ) )
        {
            auto obj = dynamic_cast< Outlook::_AttachmentSelection * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::AttachmentSelection( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Folders > ) )
        {
            auto obj = dynamic_cast< Outlook::_Folders * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Folders( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Items > ) )
        {
            auto obj = dynamic_cast< Outlook::_Items * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Items( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Explorer > ) )
        {
            auto obj = dynamic_cast< Outlook::_Explorer * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Explorer( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::NavigationPane > ) )
        {
            auto obj = dynamic_cast< Outlook::_NavigationPane * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::NavigationPane( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::NavigationModule > ) )
        {
            auto obj = dynamic_cast< Outlook::_NavigationModule * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::NavigationModule( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::NavigationModules > ) )
        {
            auto obj = dynamic_cast< Outlook::_NavigationModules * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::NavigationModules( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::AccountSelector > ) )
        {
            auto obj = dynamic_cast< Outlook::_AccountSelector * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::AccountSelector( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Account > ) )
        {
            auto obj = dynamic_cast< Outlook::_Account * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Account( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Store > ) )
        {
            auto obj = dynamic_cast< Outlook::_Store * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Store( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Rules > ) )
        {
            auto obj = dynamic_cast< Outlook::_Rules * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Rules( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Rule > ) )
        {
            auto obj = dynamic_cast< Outlook::_Rule * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Rule( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::RuleActions > ) )
        {
            auto obj = dynamic_cast< Outlook::_RuleActions * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::RuleActions( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::RuleAction > ) )
        {
            auto obj = dynamic_cast< Outlook::_RuleAction * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::RuleAction( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::MoveOrCopyRuleAction > ) )
        {
            auto obj = dynamic_cast< Outlook::_MoveOrCopyRuleAction * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::MoveOrCopyRuleAction( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::SendRuleAction > ) )
        {
            auto obj = dynamic_cast< Outlook::_SendRuleAction * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::SendRuleAction( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::AssignToCategoryRuleAction > ) )
        {
            auto obj = dynamic_cast< Outlook::_AssignToCategoryRuleAction * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::AssignToCategoryRuleAction( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::PlaySoundRuleAction > ) )
        {
            auto obj = dynamic_cast< Outlook::_PlaySoundRuleAction * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::PlaySoundRuleAction( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::MarkAsTaskRuleAction > ) )
        {
            auto obj = dynamic_cast< Outlook::_MarkAsTaskRuleAction * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::MarkAsTaskRuleAction( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::NewItemAlertRuleAction > ) )
        {
            auto obj = dynamic_cast< Outlook::_NewItemAlertRuleAction * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::NewItemAlertRuleAction( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::RuleConditions > ) )
        {
            auto obj = dynamic_cast< Outlook::_RuleConditions * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::RuleConditions( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::RuleCondition > ) )
        {
            auto obj = dynamic_cast< Outlook::_RuleCondition * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::RuleCondition( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::ImportanceRuleCondition > ) )
        {
            auto obj = dynamic_cast< Outlook::_ImportanceRuleCondition * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::ImportanceRuleCondition( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::AccountRuleCondition > ) )
        {
            auto obj = dynamic_cast< Outlook::_AccountRuleCondition * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::AccountRuleCondition( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::TextRuleCondition > ) )
        {
            auto obj = dynamic_cast< Outlook::_TextRuleCondition * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::TextRuleCondition( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::CategoryRuleCondition > ) )
        {
            auto obj = dynamic_cast< Outlook::_CategoryRuleCondition * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::CategoryRuleCondition( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::FormNameRuleCondition > ) )
        {
            auto obj = dynamic_cast< Outlook::_FormNameRuleCondition * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::FormNameRuleCondition( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::ToOrFromRuleCondition > ) )
        {
            auto obj = dynamic_cast< Outlook::_ToOrFromRuleCondition * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::ToOrFromRuleCondition( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::AddressRuleCondition > ) )
        {
            auto obj = dynamic_cast< Outlook::_AddressRuleCondition * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::AddressRuleCondition( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::SenderInAddressListRuleCondition > ) )
        {
            auto obj = dynamic_cast< Outlook::_SenderInAddressListRuleCondition * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::SenderInAddressListRuleCondition( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::FromRssFeedRuleCondition > ) )
        {
            auto obj = dynamic_cast< Outlook::_FromRssFeedRuleCondition * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::FromRssFeedRuleCondition( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::SensitivityRuleCondition > ) )
        {
            auto obj = dynamic_cast< Outlook::_SensitivityRuleCondition * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::SensitivityRuleCondition( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Categories > ) )
        {
            auto obj = dynamic_cast< Outlook::_Categories * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Categories( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Category > ) )
        {
            auto obj = dynamic_cast< Outlook::_Category * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Category( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::PreviewPane > ) )
        {
            auto obj = dynamic_cast< Outlook::_PreviewPane * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::PreviewPane( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Views > ) )
        {
            auto obj = dynamic_cast< Outlook::_Views * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Views( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::StorageItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_StorageItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::StorageItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Table > ) )
        {
            auto obj = dynamic_cast< Outlook::_Table * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Table( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Row > ) )
        {
            auto obj = dynamic_cast< Outlook::_Row * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Row( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Columns > ) )
        {
            auto obj = dynamic_cast< Outlook::_Columns * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Columns( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Column > ) )
        {
            auto obj = dynamic_cast< Outlook::_Column * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Column( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::CalendarSharing > ) )
        {
            auto obj = dynamic_cast< Outlook::_CalendarSharing * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::CalendarSharing( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::MailItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_MailItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::MailItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Conversation > ) )
        {
            auto obj = dynamic_cast< Outlook::_Conversation * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Conversation( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::SimpleItems > ) )
        {
            auto obj = dynamic_cast< Outlook::_SimpleItems * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::SimpleItems( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::UserDefinedProperties > ) )
        {
            auto obj = dynamic_cast< Outlook::_UserDefinedProperties * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::UserDefinedProperties( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::UserDefinedProperty > ) )
        {
            auto obj = dynamic_cast< Outlook::_UserDefinedProperty * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::UserDefinedProperty( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::ExchangeUser > ) )
        {
            auto obj = dynamic_cast< Outlook::_ExchangeUser * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::ExchangeUser( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::ExchangeDistributionList > ) )
        {
            auto obj = dynamic_cast< Outlook::_ExchangeDistributionList * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::ExchangeDistributionList( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::SyncObject > ) )
        {
            auto obj = dynamic_cast< Outlook::_SyncObject * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::SyncObject( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Accounts > ) )
        {
            auto obj = dynamic_cast< Outlook::_Accounts * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Accounts( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Stores > ) )
        {
            auto obj = dynamic_cast< Outlook::_Stores * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Stores( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::SelectNamesDialog > ) )
        {
            auto obj = dynamic_cast< Outlook::_SelectNamesDialog * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::SelectNamesDialog( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::SharingItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_SharingItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::SharingItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Explorers > ) )
        {
            auto obj = dynamic_cast< Outlook::_Explorers * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Explorers( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Inspectors > ) )
        {
            auto obj = dynamic_cast< Outlook::_Inspectors * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Inspectors( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Results > ) )
        {
            auto obj = dynamic_cast< Outlook::_Results * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Results( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Reminders > ) )
        {
            auto obj = dynamic_cast< Outlook::_Reminders * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Reminders( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::Reminder > ) )
        {
            auto obj = dynamic_cast< Outlook::_Reminder * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::Reminder( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::TimeZones > ) )
        {
            auto obj = dynamic_cast< Outlook::_TimeZones * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::TimeZones( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::AppointmentItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_AppointmentItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::AppointmentItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::MeetingItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_MeetingItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::MeetingItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::OutlookBarShortcuts > ) )
        {
            auto obj = dynamic_cast< Outlook::_OutlookBarShortcuts * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::OutlookBarShortcuts( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::OutlookBarGroups > ) )
        {
            auto obj = dynamic_cast< Outlook::_OutlookBarGroups * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::OutlookBarGroups( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::OutlookBarPane > ) )
        {
            auto obj = dynamic_cast< Outlook::_OutlookBarPane * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::OutlookBarPane( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::DocumentItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_DocumentItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::DocumentItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::NoteItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_NoteItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::NoteItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::ViewField > ) )
        {
            auto obj = dynamic_cast< Outlook::_ViewField * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::ViewField( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::ColumnFormat > ) )
        {
            auto obj = dynamic_cast< Outlook::_ColumnFormat * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::ColumnFormat( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::ViewFields > ) )
        {
            auto obj = dynamic_cast< Outlook::_ViewFields * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::ViewFields( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::IconView > ) )
        {
            auto obj = dynamic_cast< Outlook::_IconView * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::IconView( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::OrderFields > ) )
        {
            auto obj = dynamic_cast< Outlook::_OrderFields * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::OrderFields( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::OrderField > ) )
        {
            auto obj = dynamic_cast< Outlook::_OrderField * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::OrderField( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::CardView > ) )
        {
            auto obj = dynamic_cast< Outlook::_CardView * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::CardView( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::ViewFont > ) )
        {
            auto obj = dynamic_cast< Outlook::_ViewFont * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::ViewFont( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::AutoFormatRules > ) )
        {
            auto obj = dynamic_cast< Outlook::_AutoFormatRules * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::AutoFormatRules( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::AutoFormatRule > ) )
        {
            auto obj = dynamic_cast< Outlook::_AutoFormatRule * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::AutoFormatRule( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::TimelineView > ) )
        {
            auto obj = dynamic_cast< Outlook::_TimelineView * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::TimelineView( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::MailModule > ) )
        {
            auto obj = dynamic_cast< Outlook::_MailModule * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::MailModule( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::NavigationGroups > ) )
        {
            auto obj = dynamic_cast< Outlook::_NavigationGroups * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::NavigationGroups( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::NavigationGroup > ) )
        {
            auto obj = dynamic_cast< Outlook::_NavigationGroup * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::NavigationGroup( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::NavigationFolders > ) )
        {
            auto obj = dynamic_cast< Outlook::_NavigationFolders * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::NavigationFolders( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::NavigationFolder > ) )
        {
            auto obj = dynamic_cast< Outlook::_NavigationFolder * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::NavigationFolder( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::CalendarModule > ) )
        {
            auto obj = dynamic_cast< Outlook::_CalendarModule * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::CalendarModule( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::ContactsModule > ) )
        {
            auto obj = dynamic_cast< Outlook::_ContactsModule * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::ContactsModule( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::TasksModule > ) )
        {
            auto obj = dynamic_cast< Outlook::_TasksModule * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::TasksModule( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::JournalModule > ) )
        {
            auto obj = dynamic_cast< Outlook::_JournalModule * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::JournalModule( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::NotesModule > ) )
        {
            auto obj = dynamic_cast< Outlook::_NotesModule * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::NotesModule( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::BusinessCardView > ) )
        {
            auto obj = dynamic_cast< Outlook::_BusinessCardView * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::BusinessCardView( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::FormRegionStartup > ) )
        {
            auto obj = dynamic_cast< Outlook::_FormRegionStartup * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::FormRegionStartup( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::FormRegion > ) )
        {
            auto obj = dynamic_cast< Outlook::_FormRegion * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::FormRegion( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::SolutionsModule > ) )
        {
            auto obj = dynamic_cast< Outlook::_SolutionsModule * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::SolutionsModule( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::CalendarView > ) )
        {
            auto obj = dynamic_cast< Outlook::_CalendarView * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::CalendarView( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::TableView > ) )
        {
            auto obj = dynamic_cast< Outlook::_TableView * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::TableView( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::MobileItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_MobileItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::MobileItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::JournalItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_JournalItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::JournalItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::PostItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_PostItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::PostItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::TaskItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_TaskItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::TaskItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::DistListItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_DistListItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::DistListItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::ReportItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_ReportItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::ReportItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::RemoteItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_RemoteItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::RemoteItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::TaskRequestItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_TaskRequestItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::TaskRequestItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::TaskRequestAcceptItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_TaskRequestAcceptItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::TaskRequestAcceptItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::TaskRequestDeclineItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_TaskRequestDeclineItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::TaskRequestDeclineItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::TaskRequestUpdateItem > ) )
        {
            auto obj = dynamic_cast< Outlook::_TaskRequestUpdateItem * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::TaskRequestUpdateItem( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::ConversationHeader > ) )
        {
            auto obj = dynamic_cast< Outlook::_ConversationHeader * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::ConversationHeader( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::PeopleView > ) )
        {
            auto obj = dynamic_cast< Outlook::_PeopleView * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::PeopleView( obj );
            }
        }
        if ( constexpr( std::is_same_v< T, Outlook::SearchView > ) )
        {
            auto obj = dynamic_cast< Outlook::_SearchView * >( origObject );

            if ( obj )
            {
                retVal = new Outlook::SearchView( obj );
            }
        }
        return retVal;
    }

    QAxObject *constructItem( IDispatch *baseItem )
    {
        auto classType = getObjectClass( baseItem );
        if ( !classType.has_value() )
            return nullptr;

        QAxObject *retVal = nullptr;
        switch ( classType.value() )
        {
            case Outlook::OlObjectClass::olAccount:
                {
                    retVal = new Outlook::Account( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAccountRuleCondition:
                {
                    retVal = new Outlook::AccountRuleCondition( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAccounts:
                {
                    retVal = new Outlook::Accounts( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAction:
                {
                    retVal = new Outlook::Action( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olActions:
                {
                    retVal = new Outlook::Actions( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAddressEntries:
                {
                    retVal = new Outlook::AddressEntries( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAddressEntry:
                {
                    retVal = new Outlook::AddressEntry( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAddressList:
                {
                    retVal = new Outlook::AddressList( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAddressLists:
                {
                    retVal = new Outlook::AddressLists( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAddressRuleCondition:
                {
                    retVal = new Outlook::AddressRuleCondition( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olApplication:
                {
                    retVal = new Outlook::_Application( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAppointment:
                {
                    retVal = new Outlook::AppointmentItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olMeetingCancellation:
                {
                    retVal = new Outlook::MeetingItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olMeetingForwardNotification:
                {
                    retVal = new Outlook::MeetingItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olMeetingRequest:
                {
                    retVal = new Outlook::MeetingItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olMeetingResponseNegative:
                {
                    retVal = new Outlook::MeetingItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olMeetingResponsePositive:
                {
                    retVal = new Outlook::MeetingItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olMeetingResponseTentative:
                {
                    retVal = new Outlook::MeetingItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAssignToCategoryRuleAction:
                {
                    retVal = new Outlook::AssignToCategoryRuleAction( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAttachment:
                {
                    retVal = new Outlook::Attachment( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAttachments:
                {
                    retVal = new Outlook::Attachments( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAttachmentSelection:
                {
                    retVal = new Outlook::AttachmentSelection( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAutoFormatRule:
                {
                    retVal = new Outlook::AutoFormatRule( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olAutoFormatRules:
                {
                    retVal = new Outlook::AutoFormatRules( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olCalendarModule:
                {
                    retVal = new Outlook::CalendarModule( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olCalendarSharing:
                {
                    retVal = new Outlook::CalendarSharing( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olCategories:
                {
                    retVal = new Outlook::Categories( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olCategory:
                {
                    retVal = new Outlook::Category( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olCategoryRuleCondition:
                {
                    retVal = new Outlook::CategoryRuleCondition( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olClassBusinessCardView:
                {
                    retVal = new Outlook::BusinessCardView( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olClassCalendarView:
                {
                    retVal = new Outlook::CalendarView( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olClassCardView:
                {
                    retVal = new Outlook::CardView( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olClassIconView:
                {
                    retVal = new Outlook::IconView( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olClassNavigationPane:
                {
                    retVal = new Outlook::NavigationPane( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olClassPeopleView:
                {
                    retVal = new Outlook::PeopleView( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olClassTableView:
                {
                    retVal = new Outlook::TableView( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olClassTimeLineView:
                {
                    retVal = new Outlook::TimelineView( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olClassTimeZone:
                {
                    retVal = new Outlook::TimeZone( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olClassTimeZones:
                {
                    retVal = new Outlook::TimeZones( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olColumn:
                {
                    retVal = new Outlook::Column( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olColumnFormat:
                {
                    retVal = new Outlook::ColumnFormat( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olColumns:
                {
                    retVal = new Outlook::Columns( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olConflict:
                {
                    retVal = new Outlook::Conflict( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olConflicts:
                {
                    retVal = new Outlook::Conflicts( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olContact:
                {
                    retVal = new Outlook::ContactItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olContactsModule:
                {
                    retVal = new Outlook::ContactsModule( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olConversation:
                {
                    retVal = new Outlook::Conversation( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olConversationHeader:
                {
                    retVal = new Outlook::ConversationHeader( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olDistributionList:
                {
                    retVal = new Outlook::ExchangeDistributionList( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olDocument:
                {
                    retVal = new Outlook::DocumentItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olException:
                {
                    retVal = new Outlook::Exception( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olExceptions:
                {
                    retVal = new Outlook::Exceptions( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olExchangeDistributionList:
                {
                    retVal = new Outlook::ExchangeDistributionList( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olExchangeUser:
                {
                    retVal = new Outlook::ExchangeUser( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olExplorer:
                {
                    retVal = new Outlook::Explorer( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olExplorers:
                {
                    retVal = new Outlook::Explorers( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olFolder:
                {
                    retVal = new Outlook::MAPIFolder( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olFolders:
                {
                    retVal = new Outlook::Folders( baseItem );
                    break;
                }
            //case Outlook::OlObjectClass::olFolderUserProperties:
            //    {
            //          retVal = new Outlook::UserDefinedProperties( baseItem );
            //        break;
            //    }
            //case Outlook::OlObjectClass::olFolderUserProperty:
            //    {
            //          retVal = new Outlook::UserDefinedProperty( baseItem );
            //        break;
            //    }
            case Outlook::OlObjectClass::olFormDescription:
                {
                    retVal = new Outlook::FormDescription( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olFormNameRuleCondition:
                {
                    retVal = new Outlook::FormNameRuleCondition( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olFormRegion:
                {
                    retVal = new Outlook::FormRegion( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olFromRssFeedRuleCondition:
                {
                    retVal = new Outlook::FromRssFeedRuleCondition( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olFromRuleCondition:
                {
                    retVal = new Outlook::ToOrFromRuleCondition( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olImportanceRuleCondition:
                {
                    retVal = new Outlook::ImportanceRuleCondition( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olInspector:
                {
                    retVal = new Outlook::Inspector( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olInspectors:
                {
                    retVal = new Outlook::Inspectors( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olItemProperties:
                {
                    retVal = new Outlook::ItemProperties( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olItemProperty:
                {
                    retVal = new Outlook::ItemProperty( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olItems:
                {
                    retVal = new Outlook::_Items( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olJournal:
                {
                    retVal = new Outlook::JournalItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olJournalModule:
                {
                    retVal = new Outlook::JournalModule( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olMail:
                {
                    retVal = new Outlook::MailItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olMailModule:
                {
                    retVal = new Outlook::MailModule( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olMarkAsTaskRuleAction:
                {
                    retVal = new Outlook::MarkAsTaskRuleAction( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olMoveOrCopyRuleAction:
                {
                    retVal = new Outlook::MoveOrCopyRuleAction( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olNamespace:
                {
                    retVal = new Outlook::NameSpace( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olNavigationFolder:
                {
                    retVal = new Outlook::NavigationFolder( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olNavigationFolders:
                {
                    retVal = new Outlook::NavigationFolders( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olNavigationGroup:
                {
                    retVal = new Outlook::NavigationGroup( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olNavigationGroups:
                {
                    retVal = new Outlook::NavigationGroups( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olNavigationModule:
                {
                    retVal = new Outlook::NavigationModule( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olNavigationModules:
                {
                    retVal = new Outlook::NavigationModules( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olNewItemAlertRuleAction:
                {
                    retVal = new Outlook::NewItemAlertRuleAction( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olNote:
                {
                    retVal = new Outlook::NoteItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olNotesModule:
                {
                    retVal = new Outlook::NotesModule( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olOrderField:
                {
                    retVal = new Outlook::OrderField( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olOrderFields:
                {
                    retVal = new Outlook::OrderFields( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olOutlookBarGroup:
                {
                    retVal = new Outlook::OutlookBarGroup( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olOutlookBarGroups:
                {
                    retVal = new Outlook::OutlookBarGroups( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olOutlookBarPane:
                {
                    retVal = new Outlook::OutlookBarPane( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olOutlookBarShortcut:
                {
                    retVal = new Outlook::OutlookBarShortcut( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olOutlookBarShortcuts:
                {
                    retVal = new Outlook::OutlookBarShortcuts( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olOutlookBarStorage:
                {
                    retVal = new Outlook::OutlookBarStorage( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olOutspace:
                {
                    retVal = new Outlook::AccountSelector( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olPages:
                {
                    retVal = new Outlook::Pages( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olPanes:
                {
                    retVal = new Outlook::Panes( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olPlaySoundRuleAction:
                {
                    retVal = new Outlook::PlaySoundRuleAction( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olPost:
                {
                    retVal = new Outlook::PostItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olPropertyAccessor:
                {
                    retVal = new Outlook::PropertyAccessor( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olPropertyPages:
                {
                    retVal = new Outlook::PropertyPages( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olPropertyPageSite:
                {
                    retVal = new Outlook::PropertyPageSite( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olRecipient:
                {
                    retVal = new Outlook::Recipient( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olRecipients:
                {
                    retVal = new Outlook::Recipients( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olRecurrencePattern:
                {
                    retVal = new Outlook::RecurrencePattern( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olReminder:
                {
                    retVal = new Outlook::Reminder( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olReminders:
                {
                    retVal = new Outlook::Reminders( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olRemote:
                {
                    retVal = new Outlook::RemoteItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olReport:
                {
                    retVal = new Outlook::ReportItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olResults:
                {
                    retVal = new Outlook::Results( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olRow:
                {
                    retVal = new Outlook::Row( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olRule:
                {
                    retVal = new Outlook::Rule( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olRuleAction:
                {
                    retVal = new Outlook::RuleAction( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olRuleActions:
                {
                    retVal = new Outlook::RuleActions( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olRuleCondition:
                {
                    retVal = new Outlook::RuleCondition( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olRuleConditions:
                {
                    retVal = new Outlook::RuleConditions( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olRules:
                {
                    retVal = new Outlook::Rules( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olSearch:
                {
                    retVal = new Outlook::Search( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olSelection:
                {
                    retVal = new Outlook::Selection( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olSelectNamesDialog:
                {
                    retVal = new Outlook::SelectNamesDialog( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olSenderInAddressListRuleCondition:
                {
                    retVal = new Outlook::SenderInAddressListRuleCondition( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olSendRuleAction:
                {
                    retVal = new Outlook::SendRuleAction( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olSharing:
                {
                    retVal = new Outlook::SharingItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olSimpleItems:
                {
                    retVal = new Outlook::SimpleItems( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olSolutionsModule:
                {
                    retVal = new Outlook::SolutionsModule( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olStorageItem:
                {
                    retVal = new Outlook::StorageItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olStore:
                {
                    retVal = new Outlook::Store( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olStores:
                {
                    retVal = new Outlook::Stores( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olSyncObject:
                {
                    retVal = new Outlook::SyncObject( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olSyncObjects:
                {
                    retVal = new Outlook::SyncObjects( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olTable:
                {
                    retVal = new Outlook::Table( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olTask:
                {
                    retVal = new Outlook::TaskItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olTaskRequest:
                {
                    retVal = new Outlook::TaskRequestItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olTaskRequestAccept:
                {
                    retVal = new Outlook::TaskRequestAcceptItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olTaskRequestDecline:
                {
                    retVal = new Outlook::TaskRequestDeclineItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olTaskRequestUpdate:
                {
                    retVal = new Outlook::TaskRequestUpdateItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olTasksModule:
                {
                    retVal = new Outlook::TasksModule( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olTextRuleCondition:
                {
                    retVal = new Outlook::TextRuleCondition( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olUserDefinedProperties:
                {
                    retVal = new Outlook::UserDefinedProperties( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olUserDefinedProperty:
                {
                    retVal = new Outlook::UserDefinedProperty( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olUserProperties:
                {
                    retVal = new Outlook::UserProperties( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olUserProperty:
                {
                    retVal = new Outlook::UserProperty( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olView:
                {
                    retVal = new Outlook::View( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olViewField:
                {
                    retVal = new Outlook::ViewField( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olViewFields:
                {
                    retVal = new Outlook::ViewFields( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olViewFont:
                {
                    retVal = new Outlook::ViewFont( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olViews:
                {
                    retVal = new Outlook::Views( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olLink:
                {
                    retVal = new Outlook::Link( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olLinks:
                {
                    retVal = new Outlook::Links( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olMobile:
                {
                    retVal = new Outlook::MobileItem( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olClassThreadView:
                {
                    retVal = new Outlook::ThreadView( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olPreviewPane:
                {
                    retVal = new Outlook::PreviewPane( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olSensitivityRuleCondition:
                {
                    retVal = new Outlook::SensitivityRuleCondition( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olClassMessageListView:
                {
                    retVal = new Outlook::MessageListView( baseItem );
                    break;
                }
            case Outlook::OlObjectClass::olClassSearchView:
                {
                    retVal = new Outlook::SearchView( baseItem );
                    break;
                }
            default:
                retVal = nullptr;
        }
        return retVal;
    }

    void setObject( QAxObject *object )
    {
        if ( !object )
        {
            fClassType.reset();
            fObject.reset();
            return;
        }

        auto classType = getObjectClass( object );
        if ( !classType.has_value() )
            return;

        if ( object && !isValid( object ) )
        {
            auto newObject = updateObject( object );
            if ( newObject )
            {
                Q_ASSERT( isValid( newObject ) );
                object = newObject;
            }
        }

        if ( object && isValid( object ) )
        {
            fObject.reset( object );
            fClassType = getObjectClass( fObject.get() );
        }

        if ( object )
        {
            Q_ASSERT( isValid() );
            connectToExceptionHandler();
        }
    }
    std::shared_ptr< QAxObject > fObject;
    std::optional< Outlook::OlObjectClass > fClassType;
};

namespace std
{
    template< typename T >
    struct hash< COutlookObj< T > >
    {
        std::size_t operator()( const COutlookObj< T > &value ) const
        {   //
            return std::hash< T * >()( value.get() );
        }
    };
}

#endif