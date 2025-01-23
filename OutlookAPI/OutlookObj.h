#ifndef OUTLOOKOBJ_H
#define OUTLOOKOBJ_H

#include "ExceptionHandler.h"

#include <memory>
#include <QAxObject>
#include <cassert>
#include <utility>
#include <optional>
#include <type_traits>
#include <QString>
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

Outlook::OlObjectClass getObjectClass( IDispatch *item );
Outlook::OlObjectClass getObjectClass( QAxObject * item );


template< typename T, typename = typename std::enable_if< std::is_base_of< QAxObject, T >::value >::type >
class COutlookObj
{
    void initAppl()
    {
        if constexpr ( std::is_same_v< T, Outlook::Application > )
        {
            if ( !isValid() )
            {
                fObject = std::make_shared< Outlook::Application >();
                fClassType = getObjectClass( fObject.get() );
                Q_ASSERT( fClassType == Outlook::OlObjectClass::olApplication );
                connectToExceptionHandler();
                Q_ASSERT( isValid() );
            }
        }
    }

public:
    COutlookObj() { initAppl(); }
    COutlookObj( const std::initializer_list< COutlookObj< T > > & ) { initAppl(); }

    //COutlookObj( std::shared_ptr< QAxObject > object, Outlook::OlObjectClass classType ) :
    //    fObject( object ),
    //    fClassType( classType )
    //{
    //    connectToExceptionHandler();
    //    Q_ASSERT( isValid() );
    //}

    COutlookObj( QAxObject *object ) :
        fObject( object )
    {
        fClassType = getObjectClass( object );
        connectToExceptionHandler();
        Q_ASSERT( isValid() );
    }

    COutlookObj( IDispatch *baseItem )
    {
        constructItem( baseItem );   //
        connectToExceptionHandler();
        Q_ASSERT( isValid() );
    }

    operator bool() const { return isValid(); }

    bool isValid() const { return fClassType.has_value() && ( get() != nullptr ); }
    T *get() const
    {
        if ( !fObject )
            return {};
        auto retVal = std::dynamic_pointer_cast< T >( fObject );
        return retVal.get();
    }
    T *operator->() const { return get(); }

    void reset()
    {
        fObject.reset();
        fClassType.reset();
    }

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
    void connectToExceptionHandler()
    {
        if ( !fObject.get() )
            return;
        CExceptionHandler::instance()->connectToException( fObject.get() );
    }

    void constructItem( IDispatch *baseItem )
    {
        fClassType = getObjectClass( baseItem );
        fObject = nullptr;
        if ( !fClassType.has_value() )
            return;

        switch ( fClassType.value() )
        {
            case Outlook::OlObjectClass::olAccount:
                {
                    fObject.reset( new Outlook::Account( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAccountRuleCondition:
                {
                    fObject.reset( new Outlook::AccountRuleCondition( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAccounts:
                {
                    fObject.reset( new Outlook::Accounts( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAction:
                {
                    fObject.reset( new Outlook::Action( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olActions:
                {
                    fObject.reset( new Outlook::Actions( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAddressEntries:
                {
                    fObject.reset( new Outlook::AddressEntries( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAddressEntry:
                {
                    fObject.reset( new Outlook::AddressEntry( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAddressList:
                {
                    fObject.reset( new Outlook::AddressList( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAddressLists:
                {
                    fObject.reset( new Outlook::AddressLists( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAddressRuleCondition:
                {
                    fObject.reset( new Outlook::AddressRuleCondition( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olApplication:
                {
                    fObject.reset( new Outlook::_Application( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAppointment:
                {
                    fObject.reset( new Outlook::AppointmentItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olMeetingCancellation:
                {
                    fObject.reset( new Outlook::MeetingItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olMeetingForwardNotification:
                {
                    fObject.reset( new Outlook::MeetingItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olMeetingRequest:
                {
                    fObject.reset( new Outlook::MeetingItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olMeetingResponseNegative:
                {
                    fObject.reset( new Outlook::MeetingItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olMeetingResponsePositive:
                {
                    fObject.reset( new Outlook::MeetingItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olMeetingResponseTentative:
                {
                    fObject.reset( new Outlook::MeetingItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAssignToCategoryRuleAction:
                {
                    fObject.reset( new Outlook::AssignToCategoryRuleAction( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAttachment:
                {
                    fObject.reset( new Outlook::Attachment( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAttachments:
                {
                    fObject.reset( new Outlook::Attachments( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAttachmentSelection:
                {
                    fObject.reset( new Outlook::AttachmentSelection( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAutoFormatRule:
                {
                    fObject.reset( new Outlook::AutoFormatRule( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olAutoFormatRules:
                {
                    fObject.reset( new Outlook::AutoFormatRules( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olCalendarModule:
                {
                    fObject.reset( new Outlook::CalendarModule( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olCalendarSharing:
                {
                    fObject.reset( new Outlook::CalendarSharing( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olCategories:
                {
                    fObject.reset( new Outlook::Categories( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olCategory:
                {
                    fObject.reset( new Outlook::Category( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olCategoryRuleCondition:
                {
                    fObject.reset( new Outlook::CategoryRuleCondition( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olClassBusinessCardView:
                {
                    fObject.reset( new Outlook::BusinessCardView( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olClassCalendarView:
                {
                    fObject.reset( new Outlook::CalendarView( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olClassCardView:
                {
                    fObject.reset( new Outlook::CardView( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olClassIconView:
                {
                    fObject.reset( new Outlook::IconView( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olClassNavigationPane:
                {
                    fObject.reset( new Outlook::NavigationPane( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olClassPeopleView:
                {
                    fObject.reset( new Outlook::PeopleView( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olClassTableView:
                {
                    fObject.reset( new Outlook::TableView( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olClassTimeLineView:
                {
                    fObject.reset( new Outlook::TimelineView( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olClassTimeZone:
                {
                    fObject.reset( new Outlook::TimeZone( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olClassTimeZones:
                {
                    fObject.reset( new Outlook::TimeZones( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olColumn:
                {
                    fObject.reset( new Outlook::Column( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olColumnFormat:
                {
                    fObject.reset( new Outlook::ColumnFormat( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olColumns:
                {
                    fObject.reset( new Outlook::Columns( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olConflict:
                {
                    fObject.reset( new Outlook::Conflict( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olConflicts:
                {
                    fObject.reset( new Outlook::Conflicts( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olContact:
                {
                    fObject.reset( new Outlook::ContactItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olContactsModule:
                {
                    fObject.reset( new Outlook::ContactsModule( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olConversation:
                {
                    fObject.reset( new Outlook::Conversation( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olConversationHeader:
                {
                    fObject.reset( new Outlook::ConversationHeader( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olDistributionList:
                {
                    fObject.reset( new Outlook::ExchangeDistributionList( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olDocument:
                {
                    fObject.reset( new Outlook::DocumentItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olException:
                {
                    fObject.reset( new Outlook::Exception( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olExceptions:
                {
                    fObject.reset( new Outlook::Exceptions( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olExchangeDistributionList:
                {
                    fObject.reset( new Outlook::ExchangeDistributionList( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olExchangeUser:
                {
                    fObject.reset( new Outlook::ExchangeUser( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olExplorer:
                {
                    fObject.reset( new Outlook::Explorer( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olExplorers:
                {
                    fObject.reset( new Outlook::Explorers( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olFolder:
                {
                    fObject.reset( new Outlook::MAPIFolder( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olFolders:
                {
                    fObject.reset( new Outlook::Folders( baseItem ) );
                    break;
                }
            //case Outlook::OlObjectClass::olFolderUserProperties:
            //    {
            //        fObject.reset( new Outlook::UserDefinedProperties( baseItem ) );
            //        break;
            //    }
            //case Outlook::OlObjectClass::olFolderUserProperty:
            //    {
            //        fObject.reset( new Outlook::UserDefinedProperty( baseItem ) );
            //        break;
            //    }
            case Outlook::OlObjectClass::olFormDescription:
                {
                    fObject.reset( new Outlook::FormDescription( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olFormNameRuleCondition:
                {
                    fObject.reset( new Outlook::FormNameRuleCondition( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olFormRegion:
                {
                    fObject.reset( new Outlook::FormRegion( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olFromRssFeedRuleCondition:
                {
                    fObject.reset( new Outlook::FromRssFeedRuleCondition( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olFromRuleCondition:
                {
                    fObject.reset( new Outlook::ToOrFromRuleCondition( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olImportanceRuleCondition:
                {
                    fObject.reset( new Outlook::ImportanceRuleCondition( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olInspector:
                {
                    fObject.reset( new Outlook::Inspector( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olInspectors:
                {
                    fObject.reset( new Outlook::Inspectors( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olItemProperties:
                {
                    fObject.reset( new Outlook::ItemProperties( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olItemProperty:
                {
                    fObject.reset( new Outlook::ItemProperty( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olItems:
                {
                    fObject.reset( new Outlook::_Items( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olJournal:
                {
                    fObject.reset( new Outlook::JournalItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olJournalModule:
                {
                    fObject.reset( new Outlook::JournalModule( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olMail:
                {
                    fObject.reset( new Outlook::MailItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olMailModule:
                {
                    fObject.reset( new Outlook::MailModule( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olMarkAsTaskRuleAction:
                {
                    fObject.reset( new Outlook::MarkAsTaskRuleAction( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olMoveOrCopyRuleAction:
                {
                    fObject.reset( new Outlook::MoveOrCopyRuleAction( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olNamespace:
                {
                    fObject.reset( new Outlook::NameSpace( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olNavigationFolder:
                {
                    fObject.reset( new Outlook::NavigationFolder( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olNavigationFolders:
                {
                    fObject.reset( new Outlook::NavigationFolders( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olNavigationGroup:
                {
                    fObject.reset( new Outlook::NavigationGroup( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olNavigationGroups:
                {
                    fObject.reset( new Outlook::NavigationGroups( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olNavigationModule:
                {
                    fObject.reset( new Outlook::NavigationModule( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olNavigationModules:
                {
                    fObject.reset( new Outlook::NavigationModules( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olNewItemAlertRuleAction:
                {
                    fObject.reset( new Outlook::NewItemAlertRuleAction( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olNote:
                {
                    fObject.reset( new Outlook::NoteItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olNotesModule:
                {
                    fObject.reset( new Outlook::NotesModule( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olOrderField:
                {
                    fObject.reset( new Outlook::OrderField( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olOrderFields:
                {
                    fObject.reset( new Outlook::OrderFields( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olOutlookBarGroup:
                {
                    fObject.reset( new Outlook::OutlookBarGroup( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olOutlookBarGroups:
                {
                    fObject.reset( new Outlook::OutlookBarGroups( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olOutlookBarPane:
                {
                    fObject.reset( new Outlook::OutlookBarPane( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olOutlookBarShortcut:
                {
                    fObject.reset( new Outlook::OutlookBarShortcut( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olOutlookBarShortcuts:
                {
                    fObject.reset( new Outlook::OutlookBarShortcuts( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olOutlookBarStorage:
                {
                    fObject.reset( new Outlook::OutlookBarStorage( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olOutspace:
                {
                    fObject.reset( new Outlook::AccountSelector( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olPages:
                {
                    fObject.reset( new Outlook::Pages( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olPanes:
                {
                    fObject.reset( new Outlook::Panes( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olPlaySoundRuleAction:
                {
                    fObject.reset( new Outlook::PlaySoundRuleAction( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olPost:
                {
                    fObject.reset( new Outlook::PostItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olPropertyAccessor:
                {
                    fObject.reset( new Outlook::PropertyAccessor( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olPropertyPages:
                {
                    fObject.reset( new Outlook::PropertyPages( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olPropertyPageSite:
                {
                    fObject.reset( new Outlook::PropertyPageSite( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olRecipient:
                {
                    fObject.reset( new Outlook::Recipient( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olRecipients:
                {
                    fObject.reset( new Outlook::Recipients( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olRecurrencePattern:
                {
                    fObject.reset( new Outlook::RecurrencePattern( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olReminder:
                {
                    fObject.reset( new Outlook::Reminder( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olReminders:
                {
                    fObject.reset( new Outlook::Reminders( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olRemote:
                {
                    fObject.reset( new Outlook::RemoteItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olReport:
                {
                    fObject.reset( new Outlook::ReportItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olResults:
                {
                    fObject.reset( new Outlook::Results( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olRow:
                {
                    fObject.reset( new Outlook::Row( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olRule:
                {
                    fObject.reset( new Outlook::_Rule( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olRuleAction:
                {
                    fObject.reset( new Outlook::RuleAction( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olRuleActions:
                {
                    fObject.reset( new Outlook::RuleActions( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olRuleCondition:
                {
                    fObject.reset( new Outlook::RuleCondition( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olRuleConditions:
                {
                    fObject.reset( new Outlook::RuleConditions( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olRules:
                {
                    fObject.reset( new Outlook::Rules( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olSearch:
                {
                    fObject.reset( new Outlook::Search( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olSelection:
                {
                    fObject.reset( new Outlook::Selection( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olSelectNamesDialog:
                {
                    fObject.reset( new Outlook::SelectNamesDialog( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olSenderInAddressListRuleCondition:
                {
                    fObject.reset( new Outlook::SenderInAddressListRuleCondition( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olSendRuleAction:
                {
                    fObject.reset( new Outlook::SendRuleAction( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olSharing:
                {
                    fObject.reset( new Outlook::SharingItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olSimpleItems:
                {
                    fObject.reset( new Outlook::SimpleItems( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olSolutionsModule:
                {
                    fObject.reset( new Outlook::SolutionsModule( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olStorageItem:
                {
                    fObject.reset( new Outlook::StorageItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olStore:
                {
                    fObject.reset( new Outlook::Store( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olStores:
                {
                    fObject.reset( new Outlook::Stores( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olSyncObject:
                {
                    fObject.reset( new Outlook::SyncObject( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olSyncObjects:
                {
                    fObject.reset( new Outlook::SyncObjects( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olTable:
                {
                    fObject.reset( new Outlook::Table( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olTask:
                {
                    fObject.reset( new Outlook::TaskItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olTaskRequest:
                {
                    fObject.reset( new Outlook::TaskRequestItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olTaskRequestAccept:
                {
                    fObject.reset( new Outlook::TaskRequestAcceptItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olTaskRequestDecline:
                {
                    fObject.reset( new Outlook::TaskRequestDeclineItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olTaskRequestUpdate:
                {
                    fObject.reset( new Outlook::TaskRequestUpdateItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olTasksModule:
                {
                    fObject.reset( new Outlook::TasksModule( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olTextRuleCondition:
                {
                    fObject.reset( new Outlook::TextRuleCondition( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olUserDefinedProperties:
                {
                    fObject.reset( new Outlook::UserDefinedProperties( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olUserDefinedProperty:
                {
                    fObject.reset( new Outlook::UserDefinedProperty( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olUserProperties:
                {
                    fObject.reset( new Outlook::UserProperties( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olUserProperty:
                {
                    fObject.reset( new Outlook::UserProperty( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olView:
                {
                    fObject.reset( new Outlook::View( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olViewField:
                {
                    fObject.reset( new Outlook::ViewField( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olViewFields:
                {
                    fObject.reset( new Outlook::ViewFields( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olViewFont:
                {
                    fObject.reset( new Outlook::ViewFont( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olViews:
                {
                    fObject.reset( new Outlook::Views( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olLink:
                {
                    fObject.reset( new Outlook::Link( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olLinks:
                {
                    fObject.reset( new Outlook::Links( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olMobile:
                {
                    fObject.reset( new Outlook::MobileItem( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olClassThreadView:
                {
                    fObject.reset( new Outlook::ThreadView( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olPreviewPane:
                {
                    fObject.reset( new Outlook::PreviewPane( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olSensitivityRuleCondition:
                {
                    fObject.reset( new Outlook::SensitivityRuleCondition( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olClassMessageListView:
                {
                    fObject.reset( new Outlook::MessageListView( baseItem ) );
                    break;
                }
            case Outlook::OlObjectClass::olClassSearchView:
                {
                    fObject.reset( new Outlook::SearchView( baseItem ) );
                    break;
                }
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