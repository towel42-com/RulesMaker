#ifndef OUTLOOKAPI_H
#define OUTLOOKAPI_H

#include "OutlookObj.h"

#include <QObject>
#include <memory>
#include <list>
#include <vector>
#include <functional>
#include <optional>
#include <unordered_set>
#include <QString>
#include <map>

class QVariant;
class QWidget;
struct IDispatch;
class QStandardItem;
class QTreeView;

enum class EWrapperMode
{
    eAngleAll,
    eParenAll,
    eAngleIndividual,
    eParenIndividual,
    eNone
};

namespace Outlook
{
    class Application;
    class _Application;
    class NameSpace;
    class _NameSpace;
    class Account;
    class _Account;
    class Folder;
    class MAPIFolder;
    class MailItem;
    class _MailItem;
    class AddressEntry;
    class AddressEntries;
    class Recipient;
    class Rules;
    class _Rules;
    class Rule;
    class _Rule;
    class Recipients;
    class AddressList;
    class Items;
    class _Items;

    enum class OlImportance;
    enum class OlItemType;
    enum class OlMailRecipientType;
    enum class OlObjectClass;
    enum class OlRuleConditionType;
    enum class OlSensitivity;
    enum class OlMarkInterval;
    enum class OlDefaultFolders;
    enum class OlAddressEntryUserType;
    enum class OlDisplayType;

    class AccountRuleCondition;
    class AddressRuleCondition;
    class CategoryRuleCondition;
    class FormNameRuleCondition;
    class FromRssFeedRuleCondition;
    class ImportanceRuleCondition;
    class RuleCondition;
    class SenderInAddressListRuleCondition;
    class SensitivityRuleCondition;
    class TextRuleCondition;
    class ToOrFromRuleCondition;

    class AssignToCategoryRuleAction;
    class MarkAsTaskRuleAction;
    class MoveOrCopyRuleAction;
    class NewItemAlertRuleAction;
    class PlaySoundRuleAction;
    class RuleAction;
    class SendRuleAction;
}

enum class EFilterType
{
    eUnknown = 0x00,
    eByEmailAddressContains = 0x01,
    eBySender = 0x02,
    eByDisplayName = 0x04,
    eBySubject = 0x08,
};
bool isFilterType( EFilterType value, EFilterType filter );

QString toString( EFilterType filterType );
class CEmailAddress;
using TEmailAddressList = std::list< std::shared_ptr< CEmailAddress > >;

class COutlookAPI : public QObject
{
    Q_OBJECT;
    struct SPrivate
    {
        explicit SPrivate() = default;
    };

public:
    // neral API in OutlookAPI.cpp
    friend class COutlookSetup;
    COutlookAPI( QWidget *parentWidget, SPrivate pri );

    static std::shared_ptr< COutlookAPI > instance( QWidget *parentWidget = nullptr );
    static std::shared_ptr< COutlookAPI > cliInstance();
    virtual ~COutlookAPI();

    bool canceled() const { return fCanceled; }

    void logout( bool andNotify );

    COutlookObj< Outlook::MAPIFolder > getContacts();
    COutlookObj< Outlook::MAPIFolder > getInbox();
    COutlookObj< Outlook::MAPIFolder > getJunkFolder();
    COutlookObj< Outlook::MAPIFolder > getTrashFolder();

    void setOnlyProcessUnread( bool value, bool update = true );
    bool onlyProcessUnread() const { return fOnlyProcessUnread; }

    void setProcessAllEmailWhenLessThan200Emails( bool value, bool update = true );
    bool processAllEmailWhenLessThan200Emails() const { return fProcessAllEmailWhenLessThan200Emails; }

    void setOnlyProcessTheFirst500Emails( bool value, bool update = true );
    bool onlyProcessTheFirst500Emails() const { return fOnlyProcessTheFirst500Emails; }

    void setIncludeJunkFolderWhenRunningOnAllFolders( bool value, bool update = true );
    bool includeJunkFolderWhenRunningOnAllFolders() { return fIncludeJunkFolderWhenRunningOnAllFolders; }

    void setIncludeDeletedFolderWhenRunningOnAllFolders( bool value, bool update = true );
    bool includeDeletedFolderWhenRunningOnAllFolders() { return fIncludeDeletedFolderWhenRunningOnAllFolders; }

    void setDisableRatherThanDeleteRules( bool value, bool update = true );
    bool disableRatherThanDeleteRules() { return fDisableRatherThanDeleteRules; }

    void setRunRuleOnRootFolderWhenModified( bool value, bool update = true );
    bool runRuleOnRootFolderWhenModified() { return fRunRuleOnRootFolderWhenModified; }

    void setLoadAccountInfo( bool value, bool update = true );
    bool loadAccountInfo() const { return fLoadAccountInfo; }

    void setLastAccountName( const QString &account, bool update = true );
    QString lastAccountName() const { return fLastAccountName; }

    void setRulesToSkip( const QStringList &rulesToSkip, bool update = true );
    QStringList rulesToSkip() const { return fRulesToSkip; }

    void setEmailFilterTypes( const std::list< EFilterType > &value );
    std::list< EFilterType > emailFilterTypes() { return fEmailFilterTypes; }

    QWidget *getParentWidget() const;
    bool showRule( const COutlookObj< Outlook::Rule > & rule );
    bool editRule( const COutlookObj< Outlook::Rule > &rule );

    void sendStatusMesage( const QString &msg ){ emit sigStatusMessage( msg ); }
Q_SIGNALS:
    void sigAccountChanged();
    void sigStatusMessage( const QString &msg );
    void sigInitStatus( const QString &label, int max );
    void sigSetStatus( const QString &label, int curr, int max );
    void sigIncStatusValue( const QString &label );
    void sigStatusFinished( const QString &label );
    void sigOptionChanged();

    void sigRuleChanged( const COutlookObj< Outlook::Rule > &rule );
    void sigRuleAdded( const COutlookObj< Outlook::Rule > &rule );
    void sigRuleDeleted( const COutlookObj< Outlook::Rule > &rule );
public Q_SLOTS:
    void slotCanceled() { fCanceled = true; }
    void slotClearCanceled() { fCanceled = false; }
    void slotHandleRulesSaveException( int, const QString &, const QString &, const QString & );

public:
    // account API in OutlookAPI_account.cpp
    QString defaultProfileName() const;
    std::optional< std::map< QString, COutlookObj< Outlook::_Account > > > getAllAccounts( const QString &profile );

    QString defaultAccountName();
    QString accountName() const;
    bool accountSelected() const;

    COutlookObj< Outlook::_Account > closeAndSelectAccount( bool notifyOnChange );
    bool selectAccount( const QString &accountName, bool notifyOnChange );
    COutlookObj< Outlook::_Account > selectAccount( bool notifyOnChange );

    bool connected();

    // rules API in OutlookAPI_rules.cpp
    std::pair< COutlookObj< Outlook::Rules >, int > getRules();
    COutlookObj< Outlook::Rule > getRule( const COutlookObj< Outlook::Rules > &rules, int num );
    COutlookObj< Outlook::Rule > findRule( const QString &ruleName );

    std::optional< bool > addRule( const COutlookObj< Outlook::MAPIFolder > &folder, const std::list< std::pair< QStringList, EFilterType > > &patterns, QStringList &msgs );
    std::optional< bool > addToRule( const COutlookObj< Outlook::Rule > &rule, const std::list< std::pair< QStringList, EFilterType > > &patterns, QStringList &msg, bool copyFirst );
    COutlookObj< Outlook::Rule > copyRule( const COutlookObj< Outlook::Rule > &rule );

    static QString getDebugName( const COutlookObj< Outlook::Rule > &rule );
    static QString getDebugName( const Outlook::Rule *rule );
    static QString getDisplayName( const COutlookObj< Outlook::Rule > &rule );
    static QString getDisplayName( const Outlook::Rule *rule );

    bool ruleEnabled( const COutlookObj< Outlook::Rule > &rule );
    bool disableRule( const COutlookObj< Outlook::Rule > &rule, bool andSave );
    bool enableRule( const COutlookObj< Outlook::Rule > &rule, bool andSave );
    bool deleteRule( const COutlookObj< Outlook::Rule > &rule, bool forceDisable, bool andSave );

    void loadRuleData( QStandardItem *ruleItem, COutlookObj< Outlook::Rule > rule, bool force = false );

    QString moveTargetFolderForRule( const COutlookObj< Outlook::Rule > &rule ) const;
    std::list< EFilterType > filterTypesForRule( const COutlookObj< Outlook::Rule > &rule ) const;

    static QString ruleNameForRule( const COutlookObj< Outlook::Rule > &rule, bool forDisplay = false );
    static std::optional< QString > getDestFolderNameForRule( const COutlookObj< Outlook::Rule > & rule, bool moveOnly );   // if move only false, checks move than copy

    static QString rawRuleNameForRule( const COutlookObj< Outlook::Rule > & rule );

    static QStringList getActionStrings( const COutlookObj< Outlook::Rule > &rule );

    static std::list< QStringList > getConditionalStringList( const COutlookObj< Outlook::Rule > &rule, bool exceptions, EWrapperMode wrapperMode, bool includeSender = false );
    static std::list< QStringList > getActionStringList( const COutlookObj< Outlook::Rule > &rule );

    static bool isEnabled( const COutlookObj< Outlook::Rule > &rule );

    bool ruleBeenLoaded( const COutlookObj< Outlook::Rule > &rule ) const;
    bool ruleLessThan( const COutlookObj< Outlook::Rule > &lhsRule, const COutlookObj< Outlook::Rule > &rhsRule ) const;

    bool runAllRules( const COutlookObj< Outlook::MAPIFolder > &folder = {} );
    bool runAllRulesOnAllFolders();
    bool runAllRulesOnTrashFolder();
    bool runAllRulesOnJunkFolder();

    bool runRule( const COutlookObj< Outlook::Rule > &rule, const COutlookObj< Outlook::MAPIFolder > &folder = {} );

    // run from command line
    bool runAllRules( COutlookObj< Outlook::MAPIFolder > folder, bool allFolders, bool junk );
    bool runRule( const COutlookObj< Outlook::Rule > &rule, COutlookObj< Outlook::MAPIFolder > folder, bool allFolders, bool junk );

    // tools API in OutlookAPI_tools.cpp
    bool enableAllRules( bool andSave = true, bool *needsSaving = nullptr );
    bool mergeRules( bool andSave = true, bool *needsSaving = nullptr );

    bool moveFromToAddress( bool andSave = true, bool *needsSaving = nullptr );
    bool fixFromMessageHeaderRules( bool andSave = true, bool *needsSaving = nullptr );
    bool renameRules( bool andSave = true, bool *needsSaving = nullptr );
    bool sortRules( bool andSave = true, bool *needsSaving = nullptr );
    bool saveRules();

    // folders API in OutlookAPI_folders.cpp
    using TFolderFunc = std::function< bool( const COutlookObj< Outlook::MAPIFolder > &folder ) >;

    COutlookObj< Outlook::MAPIFolder > rootFolder();
    QString rootFolderName();

    void setRootFolder( const COutlookObj< Outlook::MAPIFolder > &folder, bool update = true );

    QString folderDisplayPath( const COutlookObj< Outlook::MAPIFolder > &folder, bool removeLeadingSlashes = false ) const;
    QString folderDisplayName( const Outlook::MAPIFolder *folder );

    COutlookObj< Outlook::MAPIFolder > findFolder( const QString &folderName, COutlookObj< Outlook::MAPIFolder > parentFolder );
    COutlookObj< Outlook::MAPIFolder > getFolder( const Outlook::MAPIFolder *item );

    int recursiveSubFolderCount( const Outlook::MAPIFolder *parent );

    std::list< COutlookObj< Outlook::MAPIFolder > > getFolders( const COutlookObj< Outlook::MAPIFolder > &parent, bool recursive, const TFolderFunc &acceptFolder = {}, const TFolderFunc &checkChildFolders = {} );
    COutlookObj< Outlook::MAPIFolder > addFolder( const COutlookObj< Outlook::MAPIFolder > &parent, const QString &folderName );
    COutlookObj< Outlook::MAPIFolder > parentFolder( const COutlookObj< Outlook::MAPIFolder > &folder );

    QString nameForFolder( const COutlookObj< Outlook::MAPIFolder > &folder ) const;
    QString rawPathForFolder( const COutlookObj< Outlook::MAPIFolder > &folder ) const;

    QString folderDisplayName( const COutlookObj< Outlook::MAPIFolder > &folder );

    bool emptyJunk();
    bool emptyTrash();

    // email API in OutlookAPI_email.cpp
    enum class EAddressTypes
    {
        eNone = 0x00,
        eOriginator = 0x01,
        eTo = 0x02,
        eCC = 0x04,
        eBCC = 0x08,
        eAllRecipients = eOriginator | eTo | eCC | eBCC,
        eSender = 0x10,
        eAllEmailAddresses = eAllRecipients | eSender
    };
    static bool isAddressType( EAddressTypes value, EAddressTypes filter );
    static bool isAddressType( std::optional< EAddressTypes > value, std::optional< EAddressTypes > filter );
    static bool isAddressType( Outlook::OlMailRecipientType type, std::optional< EAddressTypes > filter );

    enum class EContactTypes
    {
        eNone = 0x00,
        eSMTPContact = 0x01,
        eOutlookContact = 0x02,
        eAllContacts = eSMTPContact | eOutlookContact
    };
    static bool isContactType( EContactTypes value, EContactTypes filter );
    static bool isContactType( bool isExchangeUser, std::optional< EContactTypes > contactTypes );

    static bool isContactType( std::optional< EContactTypes > value, std::optional< EContactTypes > filter );
    static bool isContactType( Outlook::OlAddressEntryUserType type, std::optional< EContactTypes > filter );

    std::pair< COutlookObj< Outlook::_Items >, int > getEmailItemsForRootFolder();

    COutlookObj< Outlook::MailItem > getEmailItem( const COutlookObj< Outlook::_Items > &items, int num );
    COutlookObj< Outlook::MailItem > getEmailItem( IDispatch *item );

    static bool isExchangeUser( Outlook::AddressEntry *address );

    static QString getSubject( const COutlookObj< Outlook::MailItem > & mailItem );
    static QString getSubject( Outlook::MailItem *mailItem );

    void displayEmail( const COutlookObj< Outlook::MailItem > &email ) const;

    // in the Outlook_emailAddress file
    static QStringList getSenderEmailAddresses( Outlook::MailItem *mailItem );

    static TEmailAddressList getEmailAddresses( const COutlookObj< Outlook::MailItem > &mailItem, std::optional< EAddressTypes > addressTypes = {}, std::optional< EContactTypes > contactTypes = {} );
    static TEmailAddressList getEmailAddresses( Outlook::MailItem *mailItem, std::optional< EAddressTypes > addressTypes = {}, std::optional< EContactTypes > contactTypes = {} );
    static TEmailAddressList getEmailAddresses( Outlook::Recipients *recipients, std::optional< EAddressTypes > addressTypes = {}, std::optional< EContactTypes > contactTypes = {} );
    static TEmailAddressList getEmailAddresses( Outlook::Recipient *recipient, std::optional< EAddressTypes > addressTypes = {}, std::optional< EContactTypes > contactTypes = {} );
    static TEmailAddressList getEmailAddresses( Outlook::AddressList *addresses, std::optional< EContactTypes > contactTypes = {} );
    static TEmailAddressList getEmailAddresses( Outlook::AddressEntries *entries, std::optional< EContactTypes > contactTypes = {} );
    static TEmailAddressList getEmailAddresses( Outlook::AddressEntry *address, std::optional< EContactTypes > contactTypes = {} );

    static std::list< Outlook::AddressEntry * > getAddressEntries( Outlook::Recipients *recipients );
    static std::list< Outlook::AddressEntry * > getAddressEntries( Outlook::Recipient *recipients );

    // dump API in OutlookAPI_dump.cpp
    void dumpSession( Outlook::NameSpace &session );
    void dumpMetaMethods( QObject *object );

    std::pair< void *, Outlook::OlObjectClass > createObj( IDispatch *baseItem );

private:
    // general API in OutlookAPI.cpp
    bool showRuleDialog( const COutlookObj< Outlook::Rule > &rule, bool readOnly );

    COutlookObj< Outlook::Application > getApplication();
    COutlookObj< Outlook::Application > outlookApp();

    COutlookObj< Outlook::_Items > getItems( Outlook::_Items *item );
    COutlookObj< Outlook::MAPIFolder > selectContacts();
    COutlookObj< Outlook::MAPIFolder > selectInbox();

    void initSettings();
    void setEmailFilterTypes( EFilterType value );

    QWidget *fParentWidget{ nullptr };
    COutlookObj< Outlook::Application > fOutlookApp;
    COutlookObj< Outlook::_NameSpace > fSession;
    COutlookObj< Outlook::_Account > fAccount;
    COutlookObj< Outlook::MAPIFolder > fInbox;
    COutlookObj< Outlook::MAPIFolder > fRootFolder;   // used for loading emails
    COutlookObj< Outlook::MAPIFolder > fJunkFolder;
    COutlookObj< Outlook::MAPIFolder > fTrashFolder;
    COutlookObj< Outlook::MAPIFolder > fContacts;
    COutlookObj< Outlook::Rules > fRules;

    static std::shared_ptr< COutlookAPI > sInstance;

    bool fLoggedIn{ false };
    bool fCanceled{ false };
    bool fOnlyProcessUnread{ true };
    bool fProcessAllEmailWhenLessThan200Emails{ true };
    bool fOnlyProcessTheFirst500Emails{ true };
    bool fIncludeJunkFolderWhenRunningOnAllFolders{ false };
    bool fIncludeDeletedFolderWhenRunningOnAllFolders{ false };
    bool fDisableRatherThanDeleteRules{ false };
    bool fRunRuleOnRootFolderWhenModified{ true };
    bool fLoadAccountInfo{ true };
    QString fLastAccountName;
    std::list< EFilterType > fEmailFilterTypes{ EFilterType::eByEmailAddressContains };
    QStringList fRulesToSkip;
    bool fSaveRulesSuccess{ true };

    // account API in OutlookAPI_account.cpp
private:
    COutlookObj< Outlook::_Account > getAccount( Outlook::_Account *item );
    COutlookObj< Outlook::_NameSpace > getNamespace( Outlook::_NameSpace *ns );

    // rules api in Outlook_rules.cpp
private:
    // rules API in OutlookAPI_rules.cpp
    COutlookObj< Outlook::Rules > selectRules();
    bool skipRule( const COutlookObj< Outlook::Rule > &rule ) const;

    COutlookObj< Outlook::Rules > getRules( Outlook::Rules *item );

    std::optional< QStringList > getRecipients( Outlook::Rule *rule, QStringList *msgs );

    std::vector< COutlookObj< Outlook::Rule > > getAllRules();

    bool runRules( std::vector< COutlookObj< Outlook::Rule > > rules, COutlookObj< Outlook::MAPIFolder > folder, bool recursive = false, const std::optional< QString > &perFolderMsg = {} );

    bool addRecipientsToRule( Outlook::Rule *rule, const QStringList &recipients, QStringList &msgs );
    bool addRecipientsToRule( Outlook::Rule *rule, const TEmailAddressList &recipients, QStringList &msgs );
    bool addDisplayNamesToRule( Outlook::Rule *rule, const QStringList &displayNames, QStringList &msgs );
    bool addSubjectsToRule( Outlook::Rule *rule, const QStringList &subjects, QStringList &msgs );
    bool addSenderToRule( Outlook::Rule *rule, const TEmailAddressList &senders, QStringList &msgs );
    bool addSenderToRule( Outlook::Rule *rule, const QStringList &senders, QStringList &msgs );

    std::unordered_set< COutlookObj< Outlook::Rule > > fRuleBeenLoaded;

    // tools API in OutlookAPI_tools.cpp
private:
    using TRuleList = std::list< COutlookObj< Outlook::Rule > >;
    using TRulePair = std::pair< COutlookObj< Outlook::Rule >, TRuleList >;
    using TMergeRuleMap = std::multimap< QString, TRulePair >;

    std::optional< QString > mergeKey( const COutlookObj< Outlook::Rule > &rule ) const;
    std::optional< QStringList > mergeRecipients( Outlook::Rule *lhs, Outlook::Rule *rhs, QStringList *msgs );
    std::optional< QStringList > mergeRecipients( Outlook::Rule *lhs, const QStringList &rhs, QStringList *msgs );
    std::optional< QStringList > mergeRecipients( Outlook::Rule *lhs, const TEmailAddressList &rhs, QStringList *msgs );
    std::optional< QStringList > mergeRecipients( const std::list< Outlook::Rule * > &rules, QStringList *msgs );

    static QString stripHeaderStringString( const QString &msg );   // can include the "From" portion of the pattern as well as quotes, returns the raw pattern
    static QStringList getFromMessageHeaderString( const QString &address );   // can include the "From" portion of the pattern as well as quotes, calls stripHeaderStringString first, returns both proper patterns From: "XXX" and From: XXX
    static QStringList getFromMessageHeaderStrings( const QStringList &msgs );   // can include the "From" portion of the pattern as well as quotes, returns the raw pattern, cleans each string first then gets all the proper patterns
    bool canMergeRules( const COutlookObj< Outlook::Rule > &lhs, const COutlookObj< Outlook::Rule > &rhs );

    bool mergeRules( TRulePair &rules );
    bool mergeRule( COutlookObj< Outlook::Rule > &lhs, const COutlookObj< Outlook::Rule > &rhs );
    TMergeRuleMap findMergableRules();

private:
    // folders API in OutlookAPI_folders.cpp
    bool isFolder( const COutlookObj< Outlook::MAPIFolder > &folder, const QString &path ) const;
    bool emptyFolder( const COutlookObj< Outlook::MAPIFolder > &folder );

    void setRootFolder( const QString &folderName, bool update = true );

    COutlookObj< Outlook::MAPIFolder > getDefaultFolder( Outlook::OlDefaultFolders folderType );
    std::pair< COutlookObj< Outlook::MAPIFolder >, bool > getMailFolder( const QString &folderLabel, const QString &fullPath, bool singleOnly );   // full path after \\account
    std::pair< COutlookObj< Outlook::MAPIFolder >, bool > selectFolder( const QString &folderName, const TFolderFunc &acceptFolder, const TFolderFunc &checkChildFolders, bool singleOnly );
    std::pair< COutlookObj< Outlook::MAPIFolder >, bool > selectFolder( const QString &folderName, const std::list< COutlookObj< Outlook::MAPIFolder > > &folders, bool singleOnly );

    std::list< COutlookObj< Outlook::MAPIFolder > > getFolders( bool recursive, const TFolderFunc &acceptFolder = {}, const TFolderFunc &checkChildFolders = {} );

    static QString ruleNameForFolder( Outlook::MAPIFolder *folder );
    static QString ruleNameForFolder( const COutlookObj< Outlook::MAPIFolder > &folder );

    int subFolderCount( const Outlook::MAPIFolder *parent, bool recursive );

    // dump API in OutlookAPI_dump.cpp
    void dumpFolder( Outlook::MAPIFolder *root );
};

QString toString( const QVariant &variant, const QString &joinSeparator );
QStringList toStringList( const QVariant &variant );
[[nodiscard]] QStringList mergeStringLists( const QStringList &lhs, const QStringList &rhs, bool andSort = false );

// general API in OutlookAPI.cpp
COutlookAPI::EAddressTypes operator|( const COutlookAPI::EAddressTypes &lhs, const COutlookAPI::EAddressTypes &rhs );
COutlookAPI::EContactTypes operator|( const COutlookAPI::EContactTypes &lhs, const COutlookAPI::EContactTypes &rhs );

bool equal( const QStringList &lhs, const QStringList &rhs );

enum class EExpandMode
{
    eExpandAll,
    eCollapseAll,
    eExpandAndCollapseAll,
    eNoAction
};

void resizeToContentZero( QTreeView *treeView, EExpandMode expandMode );
#endif