#include "OutlookAPI.h"
#include "OutlookObj.h"

#include "OutlookAPI_pri.h"
#include "ShowRule.h"
#include <functional>

//#include <QInputDialog>
//#include <QMessageBox>
//#include <QDebug>
//#include <QMetaProperty>
//#include <QTreeView>
//
//#include <cstdlib>
//#include <iostream>
//#include <oaidl.h>
//#include "MSOUTL.h"
//#include <objbase.h>

//template< typename T >
//std::enable_if_t< has_delete< T >::value, void > deleteTheItem( T *obj, const std::function< void( const QString &desc ) > &onDelete )
//{
//    auto descStr = getObjectClass( obj );
//    auto desc = getDescription( obj );
//    if ( !desc.isEmpty() )
//        desc += getDescription( obj );
//    onDelete( desc );
//    obj->Delete();
//}
//
//template< typename T >
//std::enable_if_t< has_delete< T >::value, void > deleteTheItem( std::shared_ptr< T > obj, const std::function< void( const QString &desc ) > &onDelete )
//{
//    deleteTheItem( obj.get(), onDelete );
//}
//
//template< typename T >
//bool deleteIt( T *item )
//{
//    if constexpr ( has_delete< decltype( item ) >::value )
//    {
//        deleteTheItem( item, onDelete );
//        return true;
//    }
//    else
//        return false;
//}

//bool COutlookAPI::deleteItem( IDispatch *baseItem, const std::function< void( const QString &desc ) > &onDelete )
//{
//    auto classType = getObjectClass( baseItem );
//    QString description;
//    switch ( classType )
//    {
//        case Outlook::OlObjectClass::olFolder:
//            {
//                auto item = connectToException( COutlookObj< Outlook::MAPIFolder >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olRecipient:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Recipient >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olAttachment:
//            {
//                auto item = connectToException( COutlookObj< Outlook::Attachment >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olAddressEntry:
//            {
//                auto item = connectToException( COutlookObj< Outlook::AddressEntry >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olAppointment:
//            {
//                auto item = connectToException( COutlookObj< Outlook::AppointmentItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olUserProperty:
//            {
//                auto item = connectToException( COutlookObj< Outlook::UserProperty >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olContact:
//            {
//                auto item = connectToException( COutlookObj< Outlook::ContactItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olDocument:
//            {
//                auto item = connectToException( COutlookObj< Outlook::DocumentItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olJournal:
//            {
//                auto item = connectToException( COutlookObj< Outlook::JournalItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olMail:
//            {
//                auto item = connectToException( COutlookObj< Outlook::MailItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olNote:
//            {
//                auto item = connectToException( COutlookObj< Outlook::NoteItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olPost:
//            {
//                auto item = connectToException( COutlookObj< Outlook::PostItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olReport:
//            {
//                auto item = connectToException( COutlookObj< Outlook::ReportItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olRemote:
//            {
//                auto item = connectToException( COutlookObj< Outlook::RemoteItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olTask:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TaskItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olTaskRequest:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TaskRequestItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olTaskRequestUpdate:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TaskRequestUpdateItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olTaskRequestAccept:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TaskRequestAcceptItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olTaskRequestDecline:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TaskRequestDeclineItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olDistributionList:
//            {
//                auto item = connectToException( COutlookObj< Outlook::DistListItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::View >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olItemProperty:
//            {
//                auto item = connectToException( COutlookObj< Outlook::ItemProperty >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olSharing:
//            {
//                auto item = connectToException( COutlookObj< Outlook::SharingItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//            break;
//        case Outlook::OlObjectClass::olExchangeUser:
//            {
//                auto item = connectToException( COutlookObj< Outlook::ExchangeUser >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olExchangeDistributionList:
//            {
//                auto item = connectToException( COutlookObj< Outlook::ExchangeDistributionList >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olStorageItem:
//            {
//                auto item = connectToException( COutlookObj< Outlook::StorageItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olClassTableView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TableView >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olClassIconView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::IconView >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olClassCardView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::CardView >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olClassCalendarView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::CalendarView >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olClassTimeLineView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::TimelineView >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olClassBusinessCardView:
//            {
//                auto item = connectToException( COutlookObj< Outlook::BusinessCardView >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olUserDefinedProperty:
//            {
//                auto item = connectToException( COutlookObj< Outlook::UserDefinedProperty >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olMobile:
//            {
//                auto item = connectToException( COutlookObj< Outlook::MobileItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olMeetingCancellation:
//        case Outlook::OlObjectClass::olMeetingForwardNotification:
//        case Outlook::OlObjectClass::olMeetingResponseNegative:
//        case Outlook::OlObjectClass::olMeetingResponsePositive:
//        case Outlook::OlObjectClass::olMeetingResponseTentative:
//        case Outlook::OlObjectClass::olMeetingRequest:
//            {
//                auto item = connectToException( COutlookObj< Outlook::MeetingItem >( baseItem ) );
//                deleteTheItem( item, onDelete );
//                break;
//            }
//        case Outlook::OlObjectClass::olApplication:
//        case Outlook::OlObjectClass::olNamespace:
//        case Outlook::OlObjectClass::olAddressList:
//        case Outlook::OlObjectClass::olOutlookBarPane:
//        case Outlook::OlObjectClass::olOutlookBarGroup:
//        case Outlook::OlObjectClass::olOutlookBarShortcut:
//        case Outlook::OlObjectClass::olExplorer:
//        case Outlook::OlObjectClass::olInspector:
//        case Outlook::OlObjectClass::olFormDescription:
//        case Outlook::OlObjectClass::olSyncObject:
//        case Outlook::OlObjectClass::olLink:
//        case Outlook::OlObjectClass::olSearch:
//        case Outlook::OlObjectClass::olReminder:
//        case Outlook::OlObjectClass::olConflict:
//        case Outlook::OlObjectClass::olAccount:
//        case Outlook::OlObjectClass::olStore:
//        case Outlook::OlObjectClass::olSelectNamesDialog:
//        case Outlook::OlObjectClass::olRule:
//        case Outlook::OlObjectClass::olFormRegion:
//        case Outlook::OlObjectClass::olViewFont:
//        case Outlook::OlObjectClass::olAutoFormatRule:
//        case Outlook::OlObjectClass::olCategory:
//        case Outlook::OlObjectClass::olColumn:
//        case Outlook::OlObjectClass::olNavigationModule:
//        case Outlook::OlObjectClass::olMailModule:
//        case Outlook::OlObjectClass::olCalendarModule:
//        case Outlook::OlObjectClass::olContactsModule:
//        case Outlook::OlObjectClass::olTasksModule:
//        case Outlook::OlObjectClass::olJournalModule:
//        case Outlook::OlObjectClass::olNotesModule:
//        case Outlook::OlObjectClass::olNavigationGroup:
//        case Outlook::OlObjectClass::olClassTimeZone:
//        case Outlook::OlObjectClass::olSolutionsModule:
//        case Outlook::OlObjectClass::olConversation:
//        case Outlook::OlObjectClass::olConversationHeader:
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
//            return false;
//            break;
//    }
//    return true;
//}
