#include "OutlookAPI.h"
#include "OutlookAPI_pri.h"

#include "MSOUTL.h"
//
//template< typename T >
//QString getDescriptionForItem( std::shared_ptr< T > *item )
//{
//    if ( !item )
//        return {};
//    return getDescription( item );
//}
//
//QString getDescriptionForItem( IDispatch *baseItem )
//{
//    auto classType = getObjectClass( baseItem );
//    QString description;
//    switch ( classType )
//    {
//        case Outlook::OlObjectClass::olApplication:
//            {
//                auto item = COutlookObj< Outlook::_Application >( baseItem );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olNamespace:
//            {
//                auto item = connectToException( COutlookObj< Outlook::NameSpace >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olFolder:
//            {
//                auto item = connectToException( COutlookObj< Outlook::MAPIFolder >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olRecipient:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Recipient >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olAttachment:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Attachment >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olAddressList:
//            {
//                auto item = connectToException( COutlookObj< Outlook::AddressList >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olAddressEntry:
//            {
//                auto item = connectToException( COutlookObj< Outlook::AddressEntry >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olAppointment:
//            {
//                auto item = connectToException( COutlookObj< Outlook::AppointmentItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olExplorer:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Explorer >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olInspector:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Inspector >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olFormDescription:
//            {
//                auto item = connectToException( COutlookObj< Outlook::FormDescription >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olUserProperty:
//            {
//                auto item = connectToException( COutlookObj< Outlook::UserProperty >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olContact:
//            {
//                auto item = connectToException( COutlookObj< Outlook::ContactItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olDocument:
//            {
//                auto item = connectToException( COutlookObj< Outlook::DocumentItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olJournal:
//            {
//                auto item = connectToException( COutlookObj< Outlook::JournalItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olMail:
//            {
//                auto item = connectToException( COutlookObj< Outlook::MailItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olNote:
//            {
//                auto item = connectToException( COutlookObj< Outlook::NoteItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olPost:
//            {
//                auto item = connectToException( COutlookObj< Outlook::PostItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olReport:
//            {
//                auto item = connectToException( COutlookObj< Outlook::ReportItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olRemote:
//            {
//                auto item = connectToException( COutlookObj< Outlook::RemoteItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olTask:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TaskItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olTaskRequest:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TaskRequestItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olTaskRequestUpdate:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TaskRequestUpdateItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olTaskRequestAccept:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TaskRequestAcceptItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olTaskRequestDecline:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TaskRequestDeclineItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olOutlookBarPane:
//            {
//                auto item = connectToException( COutlookObj< Outlook::OutlookBarPane >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olOutlookBarGroup:
//            {
//                auto item = connectToException( COutlookObj< Outlook::OutlookBarGroup >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olOutlookBarShortcut:
//            {
//                auto item = connectToException( COutlookObj< Outlook::OutlookBarShortcut >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olDistributionList:
//            {
//                auto item = connectToException( COutlookObj< Outlook::DistListItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olSyncObject:
//            {
//                auto item = connectToException( COutlookObj< Outlook::SyncObject >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olLink:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Link >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olSearch:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Search >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::View >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olItemProperty:
//            {
//                auto item = connectToException( COutlookObj< Outlook::ItemProperty >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olReminder:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Reminder >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olConflict:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Conflict >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olSharing:
//            {
//                auto item = connectToException( COutlookObj< Outlook::SharingItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olAccount:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Account >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olStore:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Store >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olSelectNamesDialog:
//            {
//                auto item = connectToException( COutlookObj< Outlook::SelectNamesDialog >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olExchangeUser:
//            {
//                auto item = connectToException( COutlookObj< Outlook::ExchangeUser >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olExchangeDistributionList:
//            {
//                auto item = connectToException( COutlookObj< Outlook::ExchangeDistributionList >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olStorageItem:
//            {
//                auto item = connectToException( COutlookObj< Outlook::StorageItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olRule:
//            {
//                auto item = connectToException( COutlookObj< Outlook::_Rule >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//
//        case Outlook::OlObjectClass::olFormRegion:
//            {
//                auto item = connectToException( COutlookObj< Outlook::FormRegion >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//
//        case Outlook::OlObjectClass::olClassTableView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TableView >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olClassIconView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::IconView >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olClassCardView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::CardView >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olClassCalendarView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::CalendarView >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olClassTimeLineView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TimelineView >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olViewFont:
//            {
//                auto item = connectToException( COutlookObj< Outlook::ViewFont >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olAutoFormatRule:
//            {
//                auto item = connectToException( COutlookObj< Outlook::AutoFormatRule >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olCategory:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Category >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olColumn:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Column >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olNavigationModule:
//            {
//                auto item = connectToException( COutlookObj< Outlook::NavigationModule >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olMailModule:
//            {
//                auto item = connectToException( COutlookObj< Outlook::MailModule >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olCalendarModule:
//            {
//                auto item = connectToException( COutlookObj< Outlook::CalendarModule >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olContactsModule:
//            {
//                auto item = connectToException( COutlookObj< Outlook::ContactsModule >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olTasksModule:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TasksModule >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olJournalModule:
//            {
//                auto item = connectToException( COutlookObj< Outlook::JournalModule >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olNotesModule:
//            {
//                auto item = connectToException( COutlookObj< Outlook::NotesModule >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olNavigationGroup:
//            {
//                auto item = connectToException( COutlookObj< Outlook::NavigationGroup >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olClassBusinessCardView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::BusinessCardView >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olUserDefinedProperty:
//            {
//                auto item = connectToException( COutlookObj< Outlook::UserDefinedProperty >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olClassTimeZone:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TimeZone >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olMobile:
//            {
//                auto item = connectToException( COutlookObj< Outlook::MobileItem >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olSolutionsModule:
//            {
//                auto item = connectToException( COutlookObj< Outlook::SolutionsModule >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olConversation:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Conversation >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olConversationHeader:
//            {
//                auto item = connectToException( COutlookObj< Outlook::ConversationHeader >( baseItem ) );
//                description = getDescription( item );
//                break;
//            }
//        case Outlook::OlObjectClass::olAccountRuleCondition:
//        case Outlook::OlObjectClass::olAccounts:
//        case Outlook::OlObjectClass::olAction:
//        case Outlook::OlObjectClass::olActions:
//        case Outlook::OlObjectClass::olAddressEntries:
//        case Outlook::OlObjectClass::olAddressLists:
//        case Outlook::OlObjectClass::olAddressRuleCondition:
//        case Outlook::OlObjectClass::olAssignToCategoryRuleAction:
//        case Outlook::OlObjectClass::olAttachmentSelection:
//        case Outlook::OlObjectClass::olAttachments:
//        case Outlook::OlObjectClass::olAutoFormatRules:
//        case Outlook::OlObjectClass::olCalendarSharing:
//        case Outlook::OlObjectClass::olCategories:
//        case Outlook::OlObjectClass::olCategoryRuleCondition:
//        case Outlook::OlObjectClass::olClassMessageListView:
//        case Outlook::OlObjectClass::olClassNavigationPane:
//        case Outlook::OlObjectClass::olClassPeopleView:
//        case Outlook::OlObjectClass::olClassSearchView:
//        case Outlook::OlObjectClass::olClassThreadView:
//        case Outlook::OlObjectClass::olClassTimeZones:
//        case Outlook::OlObjectClass::olColumnFormat:
//        case Outlook::OlObjectClass::olColumns:
//        case Outlook::OlObjectClass::olConflicts:
//        case Outlook::OlObjectClass::olException:
//        case Outlook::OlObjectClass::olExceptions:
//        case Outlook::OlObjectClass::olExplorers:
//        case Outlook::OlObjectClass::olFolders:
//        case Outlook::OlObjectClass::olFormNameRuleCondition:
//        case Outlook::OlObjectClass::olFromRssFeedRuleCondition:
//        case Outlook::OlObjectClass::olFromRuleCondition:
//        case Outlook::OlObjectClass::olImportanceRuleCondition:
//        case Outlook::OlObjectClass::olInspectors:
//        case Outlook::OlObjectClass::olItemProperties:
//        case Outlook::OlObjectClass::olItems:
//        case Outlook::OlObjectClass::olLinks:
//        case Outlook::OlObjectClass::olMarkAsTaskRuleAction:
//        case Outlook::OlObjectClass::olMeetingCancellation:
//        case Outlook::OlObjectClass::olMeetingForwardNotification:
//        case Outlook::OlObjectClass::olMeetingRequest:
//        case Outlook::OlObjectClass::olMeetingResponseNegative:
//        case Outlook::OlObjectClass::olMeetingResponsePositive:
//        case Outlook::OlObjectClass::olMeetingResponseTentative:
//        case Outlook::OlObjectClass::olMoveOrCopyRuleAction:
//        case Outlook::OlObjectClass::olNavigationFolder:
//        case Outlook::OlObjectClass::olNavigationFolders:
//        case Outlook::OlObjectClass::olNavigationGroups:
//        case Outlook::OlObjectClass::olNavigationModules:
//        case Outlook::OlObjectClass::olNewItemAlertRuleAction:
//        case Outlook::OlObjectClass::olOrderField:
//        case Outlook::OlObjectClass::olOrderFields:
//        case Outlook::OlObjectClass::olOutlookBarGroups:
//        case Outlook::OlObjectClass::olOutlookBarShortcuts:
//        case Outlook::OlObjectClass::olOutlookBarStorage:
//        case Outlook::OlObjectClass::olOutspace:
//        case Outlook::OlObjectClass::olPages:
//        case Outlook::OlObjectClass::olPanes:
//        case Outlook::OlObjectClass::olPlaySoundRuleAction:
//        case Outlook::OlObjectClass::olPreviewPane:
//        case Outlook::OlObjectClass::olPropertyAccessor:
//        case Outlook::OlObjectClass::olPropertyPageSite:
//        case Outlook::OlObjectClass::olPropertyPages:
//        case Outlook::OlObjectClass::olRecipients:
//        case Outlook::OlObjectClass::olRecurrencePattern:
//        case Outlook::OlObjectClass::olReminders:
//        case Outlook::OlObjectClass::olResults:
//        case Outlook::OlObjectClass::olRow:
//        case Outlook::OlObjectClass::olRuleAction:
//        case Outlook::OlObjectClass::olRuleActions:
//        case Outlook::OlObjectClass::olRuleCondition:
//        case Outlook::OlObjectClass::olRuleConditions:
//        case Outlook::OlObjectClass::olRules:
//        case Outlook::OlObjectClass::olSelection:
//        case Outlook::OlObjectClass::olSendRuleAction:
//        case Outlook::OlObjectClass::olSenderInAddressListRuleCondition:
//        case Outlook::OlObjectClass::olSensitivityRuleCondition:
//        case Outlook::OlObjectClass::olSimpleItems:
//        case Outlook::OlObjectClass::olStores:
//        case Outlook::OlObjectClass::olSyncObjects:
//        case Outlook::OlObjectClass::olTable:
//        case Outlook::OlObjectClass::olTextRuleCondition:
//        case Outlook::OlObjectClass::olUserDefinedProperties:
//        case Outlook::OlObjectClass::olUserProperties:
//        case Outlook::OlObjectClass::olViewField:
//        case Outlook::OlObjectClass::olViewFields:
//        case Outlook::OlObjectClass::olViews:
//            break;
//    }
//    auto retVal = toString( classType );
//    if ( !description.isEmpty() )
//        retVal += " - " + description;
//    return retVal;
//}
