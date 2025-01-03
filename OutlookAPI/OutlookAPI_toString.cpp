#include "OutlookAPI.h"
#include <QVariant>

#include "MSOUTL.h"

QString toString( Outlook::OlItemType olItemType )
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

QString toString( Outlook::OlObjectClass objectClass )
{
    switch ( objectClass )
    {
        case Outlook::OlObjectClass::olApplication:
            return "Application";
        case Outlook::OlObjectClass::olNamespace:
            return "Namespace";
        case Outlook::OlObjectClass::olFolder:
            return "Folder";
        case Outlook::OlObjectClass::olRecipient:
            return "Recipient";
        case Outlook::OlObjectClass::olAttachment:
            return "Attachment";
        case Outlook::OlObjectClass::olAddressList:
            return "Address List";
        case Outlook::OlObjectClass::olAddressEntry:
            return "Address Entry";
        case Outlook::OlObjectClass::olFolders:
            return "Folders";
        case Outlook::OlObjectClass::olItems:
            return "Items";
        case Outlook::OlObjectClass::olRecipients:
            return "Recipients";
        case Outlook::OlObjectClass::olAttachments:
            return "Attachments";
        case Outlook::OlObjectClass::olAddressLists:
            return "Address Lists";
        case Outlook::OlObjectClass::olAddressEntries:
            return "Address Entries";
        case Outlook::OlObjectClass::olAppointment:
            return "Appointment";
        case Outlook::OlObjectClass::olMeetingRequest:
            return "Meeting Request";
        case Outlook::OlObjectClass::olMeetingCancellation:
            return "Meeting Cancellation";
        case Outlook::OlObjectClass::olMeetingResponseNegative:
            return "Meeting Response Negative";
        case Outlook::OlObjectClass::olMeetingResponsePositive:
            return "Meeting Response Positive";
        case Outlook::OlObjectClass::olMeetingResponseTentative:
            return "Meeting Response Tentative";
        case Outlook::OlObjectClass::olRecurrencePattern:
            return "Recurrence Pattern";
        case Outlook::OlObjectClass::olExceptions:
            return "Exceptions";
        case Outlook::OlObjectClass::olException:
            return "Exception";
        case Outlook::OlObjectClass::olAction:
            return "Action";
        case Outlook::OlObjectClass::olActions:
            return "Actions";
        case Outlook::OlObjectClass::olExplorer:
            return "Explorer";
        case Outlook::OlObjectClass::olInspector:
            return "Inspector";
        case Outlook::OlObjectClass::olPages:
            return "Pages";
        case Outlook::OlObjectClass::olFormDescription:
            return "Form Description";
        case Outlook::OlObjectClass::olUserProperties:
            return "User Properties";
        case Outlook::OlObjectClass::olUserProperty:
            return "User Property";
        case Outlook::OlObjectClass::olContact:
            return "Contact";
        case Outlook::OlObjectClass::olDocument:
            return "Document";
        case Outlook::OlObjectClass::olJournal:
            return "Journal";
        case Outlook::OlObjectClass::olMail:
            return "Mail";
        case Outlook::OlObjectClass::olNote:
            return "Note";
        case Outlook::OlObjectClass::olPost:
            return "Post";
        case Outlook::OlObjectClass::olReport:
            return "Report";
        case Outlook::OlObjectClass::olRemote:
            return "Remote";
        case Outlook::OlObjectClass::olTask:
            return "Task";
        case Outlook::OlObjectClass::olTaskRequest:
            return "Task Request";
        case Outlook::OlObjectClass::olTaskRequestUpdate:
            return "Task Request Update";
        case Outlook::OlObjectClass::olTaskRequestAccept:
            return "Task Request Accept";
        case Outlook::OlObjectClass::olTaskRequestDecline:
            return "Task Request Decline";
        case Outlook::OlObjectClass::olExplorers:
            return "Explorers";
        case Outlook::OlObjectClass::olInspectors:
            return "Inspectors";
        case Outlook::OlObjectClass::olPanes:
            return "Panes";
        case Outlook::OlObjectClass::olOutlookBarPane:
            return "Outlook Bar Pane";
        case Outlook::OlObjectClass::olOutlookBarStorage:
            return "Outlook Bar Storage";
        case Outlook::OlObjectClass::olOutlookBarGroups:
            return "Outlook Bar Groups";
        case Outlook::OlObjectClass::olOutlookBarGroup:
            return "Outlook Bar Group";
        case Outlook::OlObjectClass::olOutlookBarShortcuts:
            return "Outlook Bar Shortcuts";
        case Outlook::OlObjectClass::olOutlookBarShortcut:
            return "Outlook Bar Shortcut";
        case Outlook::OlObjectClass::olDistributionList:
            return "Distribution List";
        case Outlook::OlObjectClass::olPropertyPageSite:
            return "Property Page Site";
        case Outlook::OlObjectClass::olPropertyPages:
            return "Property Pages";
        case Outlook::OlObjectClass::olSyncObject:
            return "Sync Object";
        case Outlook::OlObjectClass::olSyncObjects:
            return "Sync Objects";
        case Outlook::OlObjectClass::olSelection:
            return "Selection";
        case Outlook::OlObjectClass::olLink:
            return "Link";
        case Outlook::OlObjectClass::olLinks:
            return "Links";
        case Outlook::OlObjectClass::olSearch:
            return "Search";
        case Outlook::OlObjectClass::olResults:
            return "Results";
        case Outlook::OlObjectClass::olViews:
            return "Views";
        case Outlook::OlObjectClass::olView:
            return "View";
        case Outlook::OlObjectClass::olItemProperties:
            return "Item Properties";
        case Outlook::OlObjectClass::olItemProperty:
            return "Item Property";
        case Outlook::OlObjectClass::olReminders:
            return "Reminders";
        case Outlook::OlObjectClass::olReminder:
            return "Reminder";
        case Outlook::OlObjectClass::olConflict:
            return "Conflict";
        case Outlook::OlObjectClass::olConflicts:
            return "Conflicts";
        case Outlook::OlObjectClass::olSharing:
            return "Sharing";
        case Outlook::OlObjectClass::olAccount:
            return "Account";
        case Outlook::OlObjectClass::olAccounts:
            return "Accounts";
        case Outlook::OlObjectClass::olStore:
            return "Store";
        case Outlook::OlObjectClass::olStores:
            return "Stores";
        case Outlook::OlObjectClass::olSelectNamesDialog:
            return "Select Names Dialog";
        case Outlook::OlObjectClass::olExchangeUser:
            return "Exchange User";
        case Outlook::OlObjectClass::olExchangeDistributionList:
            return "Exchange Distribution List";
        case Outlook::OlObjectClass::olPropertyAccessor:
            return "Property Accessor";
        case Outlook::OlObjectClass::olStorageItem:
            return "Storage Item";
        case Outlook::OlObjectClass::olRules:
            return "Rules";
        case Outlook::OlObjectClass::olRule:
            return "Rule";
        case Outlook::OlObjectClass::olRuleActions:
            return "Rule Actions";
        case Outlook::OlObjectClass::olRuleAction:
            return "Rule Action";
        case Outlook::OlObjectClass::olMoveOrCopyRuleAction:
            return "Move Or Copy Rule Action";
        case Outlook::OlObjectClass::olSendRuleAction:
            return "Send Rule Action";
        case Outlook::OlObjectClass::olTable:
            return "Table";
        case Outlook::OlObjectClass::olRow:
            return "Row";
        case Outlook::OlObjectClass::olAssignToCategoryRuleAction:
            return "Assign To Category Rule Action";
        case Outlook::OlObjectClass::olPlaySoundRuleAction:
            return "Play Sound Rule Action";
        case Outlook::OlObjectClass::olMarkAsTaskRuleAction:
            return "Mark As Task Rule Action";
        case Outlook::OlObjectClass::olNewItemAlertRuleAction:
            return "New Item Alert Rule Action";
        case Outlook::OlObjectClass::olRuleConditions:
            return "Rule Conditions";
        case Outlook::OlObjectClass::olRuleCondition:
            return "Rule Condition";
        case Outlook::OlObjectClass::olImportanceRuleCondition:
            return "Importance Rule Condition";
        case Outlook::OlObjectClass::olFormRegion:
            return "Form Region";
        case Outlook::OlObjectClass::olCategoryRuleCondition:
            return "Category Rule Condition";
        case Outlook::OlObjectClass::olFormNameRuleCondition:
            return "Form Name Rule Condition";
        case Outlook::OlObjectClass::olFromRuleCondition:
            return "From Rule Condition";
        case Outlook::OlObjectClass::olSenderInAddressListRuleCondition:
            return "Sender In Address List Rule Condition";
        case Outlook::OlObjectClass::olTextRuleCondition:
            return "Text Rule Condition";
        case Outlook::OlObjectClass::olAccountRuleCondition:
            return "Account Rule Condition";
        case Outlook::OlObjectClass::olClassTableView:
            return "Class Table View";
        case Outlook::OlObjectClass::olClassIconView:
            return "Class Icon View";
        case Outlook::OlObjectClass::olClassCardView:
            return "Class Card View";
        case Outlook::OlObjectClass::olClassCalendarView:
            return "Class Calendar View";
        case Outlook::OlObjectClass::olClassTimeLineView:
            return "Class Time Line View";
        case Outlook::OlObjectClass::olViewFields:
            return "View Fields";
        case Outlook::OlObjectClass::olViewField:
            return "View Field";
        case Outlook::OlObjectClass::olOrderField:
            return "Order Field";
        case Outlook::OlObjectClass::olOrderFields:
            return "Order Fields";
        case Outlook::OlObjectClass::olViewFont:
            return "View Font";
        case Outlook::OlObjectClass::olAutoFormatRule:
            return "Auto Format Rule";
        case Outlook::OlObjectClass::olAutoFormatRules:
            return "Auto Format Rules";
        case Outlook::OlObjectClass::olColumnFormat:
            return "Column Format";
        case Outlook::OlObjectClass::olColumns:
            return "Columns";
        case Outlook::OlObjectClass::olCalendarSharing:
            return "Calendar Sharing";
        case Outlook::OlObjectClass::olCategory:
            return "Category";
        case Outlook::OlObjectClass::olCategories:
            return "Categories";
        case Outlook::OlObjectClass::olColumn:
            return "Column";
        case Outlook::OlObjectClass::olClassNavigationPane:
            return "Class Navigation Pane";
        case Outlook::OlObjectClass::olNavigationModules:
            return "Navigation Modules";
        case Outlook::OlObjectClass::olNavigationModule:
            return "Navigation Module";
        case Outlook::OlObjectClass::olMailModule:
            return "Mail Module";
        case Outlook::OlObjectClass::olCalendarModule:
            return "Calendar Module";
        case Outlook::OlObjectClass::olContactsModule:
            return "Contacts Module";
        case Outlook::OlObjectClass::olTasksModule:
            return "Tasks Module";
        case Outlook::OlObjectClass::olJournalModule:
            return "Journal Module";
        case Outlook::OlObjectClass::olNotesModule:
            return "Notes Module";
        case Outlook::OlObjectClass::olNavigationGroups:
            return "Navigation Groups";
        case Outlook::OlObjectClass::olNavigationGroup:
            return "Navigation Group";
        case Outlook::OlObjectClass::olNavigationFolders:
            return "Navigation Folders";
        case Outlook::OlObjectClass::olNavigationFolder:
            return "Navigation Folder";
        case Outlook::OlObjectClass::olClassBusinessCardView:
            return "Class Business Card View";
        case Outlook::OlObjectClass::olAttachmentSelection:
            return "Attachment Selection";
        case Outlook::OlObjectClass::olAddressRuleCondition:
            return "Address Rule Condition";
        case Outlook::OlObjectClass::olUserDefinedProperty:
            return "User Defined Property";
        case Outlook::OlObjectClass::olUserDefinedProperties:
            return "User Defined Properties";
        case Outlook::OlObjectClass::olFromRssFeedRuleCondition:
            return "From Rss Feed Rule Condition";
        case Outlook::OlObjectClass::olClassTimeZone:
            return "Class Time Zone";
        case Outlook::OlObjectClass::olClassTimeZones:
            return "Class Time Zones";
        case Outlook::OlObjectClass::olMobile:
            return "Mobile";
        case Outlook::OlObjectClass::olSolutionsModule:
            return "Solutions Module";
        case Outlook::OlObjectClass::olConversation:
            return "Conversation";
        case Outlook::OlObjectClass::olSimpleItems:
            return "Simple Items";
        case Outlook::OlObjectClass::olOutspace:
            return "Outspace";
        case Outlook::OlObjectClass::olMeetingForwardNotification:
            return "Meeting Forward Notification";
        case Outlook::OlObjectClass::olConversationHeader:
            return "Conversation Header";
        case Outlook::OlObjectClass::olClassPeopleView:
            return "Class People View";
        case Outlook::OlObjectClass::olClassThreadView:
            return "Class Thread View";
        case Outlook::OlObjectClass::olPreviewPane:
            return "Preview Pane";
        case Outlook::OlObjectClass::olSensitivityRuleCondition:
            return "Sensitivity Rule Condition";
        case Outlook::OlObjectClass::olClassMessageListView:
            return "Class Message List View";
        case Outlook::OlObjectClass::olClassSearchView:
            return "Class Search View";
    }
    return "<UNKNOWN>";
}

QString toString( Outlook::OlDisplayType displayType )
{
    switch ( displayType )
    {
        case Outlook::OlDisplayType::olUser:
            return "User";
        case Outlook::OlDisplayType::olDistList:
            return "Distribution List";
        case Outlook::OlDisplayType::olForum:
            return "Forum";
        case Outlook::OlDisplayType::olAgent:
            return "Agent";
        case Outlook::OlDisplayType::olOrganization:
            return "Organization";
        case Outlook::OlDisplayType::olPrivateDistList:
            return "Private Distribution List";
        case Outlook::OlDisplayType::olRemoteUser:
            return "Remote User";
    }
    return "<UNKNOWN>";
}

QString toString( Outlook::OlRuleConditionType olRuleConditionType )
{
    switch ( olRuleConditionType )
    {
        case Outlook::OlRuleConditionType::olConditionUnknown:
            return "ConditionUnknown";
        case Outlook::OlRuleConditionType::olConditionFrom:
            return "ConditionFrom";
        case Outlook::OlRuleConditionType::olConditionSubject:
            return "ConditionSubject";
        case Outlook::OlRuleConditionType::olConditionAccount:
            return "ConditionAccount";
        case Outlook::OlRuleConditionType::olConditionOnlyToMe:
            return "ConditionOnlyToMe";
        case Outlook::OlRuleConditionType::olConditionTo:
            return "ConditionTo";
        case Outlook::OlRuleConditionType::olConditionImportance:
            return "ConditionImportance";
        case Outlook::OlRuleConditionType::olConditionSensitivity:
            return "ConditionSensitivity";
        case Outlook::OlRuleConditionType::olConditionFlaggedForAction:
            return "ConditionFlaggedForAction";
        case Outlook::OlRuleConditionType::olConditionCc:
            return "ConditionCc";
        case Outlook::OlRuleConditionType::olConditionToOrCc:
            return "ConditionToOrCc";
        case Outlook::OlRuleConditionType::olConditionNotTo:
            return "ConditionNotTo";
        case Outlook::OlRuleConditionType::olConditionSentTo:
            return "ConditionSentTo";
        case Outlook::OlRuleConditionType::olConditionBody:
            return "ConditionBody";
        case Outlook::OlRuleConditionType::olConditionBodyOrSubject:
            return "ConditionBodyOrSubject";
        case Outlook::OlRuleConditionType::olConditionMessageHeader:
            return "ConditionMessageHeader";
        case Outlook::OlRuleConditionType::olConditionRecipientAddress:
            return "ConditionRecipientAddress";
        case Outlook::OlRuleConditionType::olConditionSenderAddress:
            return "ConditionSenderAddress";
        case Outlook::OlRuleConditionType::olConditionCategory:
            return "ConditionCategory";
        case Outlook::OlRuleConditionType::olConditionOOF:
            return "ConditionOOF";
        case Outlook::OlRuleConditionType::olConditionHasAttachment:
            return "ConditionHasAttachment";
        case Outlook::OlRuleConditionType::olConditionSizeRange:
            return "ConditionSizeRange";
        case Outlook::OlRuleConditionType::olConditionDateRange:
            return "ConditionDateRange";
        case Outlook::OlRuleConditionType::olConditionFormName:
            return "ConditionFormName";
        case Outlook::OlRuleConditionType::olConditionProperty:
            return "ConditionProperty";
        case Outlook::OlRuleConditionType::olConditionSenderInAddressBook:
            return "ConditionSenderInAddressBook";
        case Outlook::OlRuleConditionType::olConditionMeetingInviteOrUpdate:
            return "ConditionMeetingInviteOrUpdate";
        case Outlook::OlRuleConditionType::olConditionLocalMachineOnly:
            return "ConditionLocalMachineOnly";
        case Outlook::OlRuleConditionType::olConditionOtherMachine:
            return "ConditionOtherMachine";
        case Outlook::OlRuleConditionType::olConditionAnyCategory:
            return "ConditionAnyCategory";
        case Outlook::OlRuleConditionType::olConditionFromRssFeed:
            return "ConditionFromRssFeed";
        case Outlook::OlRuleConditionType::olConditionFromAnyRssFeed:
            return "ConditionFromAnyRssFeed";
    }
    return "<UNKNOWN>";
}

QString toString( Outlook::OlImportance importance )
{
    switch ( importance )
    {
        case Outlook::OlImportance::olImportanceLow:
            return "Low";
        case Outlook::OlImportance::olImportanceNormal:
            return "Normal";
        case Outlook::OlImportance::olImportanceHigh:
            return "High";
    }

    return "<UNKNOWN>";
}

QString toString( Outlook::OlSensitivity sensitivity )
{
    switch ( sensitivity )
    {
        case Outlook::OlSensitivity::olPersonal:
            return "Personal";
        case Outlook::OlSensitivity::olNormal:
            return "Normal";
        case Outlook::OlSensitivity::olPrivate:
            return "Private";
        case Outlook::OlSensitivity::olConfidential:
            return "Confidential";
    }

    return "<UNKNOWN>";
}

QString toString( Outlook::OlMarkInterval markInterval )
{
    switch ( markInterval )
    {
        case Outlook::OlMarkInterval::olMarkToday:
            return "Mark Today";
        case Outlook::OlMarkInterval::olMarkTomorrow:
            return "Mark Tomorrow";
        case Outlook::OlMarkInterval::olMarkThisWeek:
            return "Mark This Week";
        case Outlook::OlMarkInterval::olMarkNextWeek:
            return "Mark Next Week";
        case Outlook::OlMarkInterval::olMarkNoDate:
            return "Mark No Date";
        case Outlook::OlMarkInterval::olMarkComplete:
            return "Mark Complete";
    }

    return "<UNKNOWN>";
}

QString toString( Outlook::OlAddressEntryUserType entryUserType )
{
    switch ( entryUserType )
    {
        case Outlook::OlAddressEntryUserType::olExchangeUserAddressEntry:
            return "Exchange User Address Entry";
        case Outlook::OlAddressEntryUserType::olExchangeDistributionListAddressEntry:
            return "Exchange Distribution ListAddressEntry";
        case Outlook::OlAddressEntryUserType::olExchangePublicFolderAddressEntry:
            return "Exchange Public Folder AddressEntry";
        case Outlook::OlAddressEntryUserType::olExchangeAgentAddressEntry:
            return "Exchange Agent Address Entry";
        case Outlook::OlAddressEntryUserType::olExchangeOrganizationAddressEntry:
            return "Exchange Organization Address Entry";
        case Outlook::OlAddressEntryUserType::olExchangeRemoteUserAddressEntry:
            return "Exchange Remote User Address Entry";
        case Outlook::OlAddressEntryUserType::olOutlookContactAddressEntry:
            return "Outlook Contact Address Entry";
        case Outlook::OlAddressEntryUserType::olOutlookDistributionListAddressEntry:
            return "Outlook Distribution List Address Entry";
        case Outlook::OlAddressEntryUserType::olLdapAddressEntry:
            return "Ldap Address Entry";
        case Outlook::OlAddressEntryUserType::olSmtpAddressEntry:
            return "Smtp Address Entry";
        case Outlook::OlAddressEntryUserType::olOtherAddressEntry:
            return "Other Address Entry";
    }
    return "<UNKNOWN>";
}

QString toString( const QVariant &variant, const QString &joinSeparator )
{
    auto stringList = toStringList( variant );
    auto retVal = stringList.join( joinSeparator );
    return retVal;
}

QStringList toStringList( const QVariant &variant )
{
    QStringList retVal;
    if ( variant.type() == QVariant::Type::String )
        retVal << variant.toString();
    else if ( variant.type() == QVariant::Type::StringList )
        retVal << variant.toStringList();
    return retVal;
}

QStringList &mergeStringLists( QStringList &lhs, const QStringList &rhs, bool andSort )
{
    lhs << rhs;
    if ( andSort )
        lhs.sort( Qt::CaseInsensitive );

    lhs.removeDuplicates();
    lhs.removeAll( QString() );

    return lhs;
}

QString toString( EFilterType filterType )
{
    switch ( filterType )
    {
        case EFilterType::eByEmailAddress:
            return "Email Address";
        case EFilterType::eByDisplayName:
            return "Display Name";
        case EFilterType::eBySubject:
            return "Subject";
        default:
            break;
    }
    return "<UNKNOWN>";
}