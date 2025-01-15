#ifndef OUTLOOKHELPERS_H
#define OUTLOOKHELPERS_H

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

    std::shared_ptr< Outlook::Folder > getContacts();
    std::shared_ptr< Outlook::Folder > getInbox();
    std::shared_ptr< Outlook::Folder > getJunkFolder();
    std::shared_ptr< Outlook::Folder > getTrashFolder();

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

    void setLoadAccountInfo( bool value, bool update = true );
    bool loadAccountInfo() const { return fLoadAccountInfo; }

    void setLastAccountName( const QString &account, bool update = true );
    QString lastAccountName() const { return fLastAccountName; }

    void setRulesToSkip( const QStringList &rulesToSkip, bool update = true );
    QStringList rulesToSkip() const { return fRulesToSkip; }

    void setEmailFilterTypes( const std::list< EFilterType > &value );
    std::list< EFilterType > emailFilterTypes() { return fEmailFilterTypes; }

    QWidget *getParentWidget() const;
    bool showRule( std::shared_ptr< Outlook::Rule > rule );
    bool editRule( std::shared_ptr< Outlook::Rule > rule );
Q_SIGNALS:
    void sigAccountChanged();
    void sigStatusMessage( const QString &msg );
    void sigInitStatus( const QString &label, int max );
    void sigSetStatus( const QString &label, int curr, int max );
    void sigIncStatusValue( const QString &label );
    void sigStatusFinished( const QString &label );
    void sigOptionChanged();

    void sigRuleChanged( std::shared_ptr< Outlook::Rule > rule );
    void sigRuleAdded( std::shared_ptr< Outlook::Rule > rule );
    void sigRuleDeleted( std::shared_ptr< Outlook::Rule > rule );
public Q_SLOTS:
    void slotCanceled() { fCanceled = true; }
    void slotClearCanceled() { fCanceled = false; }
    void slotHandleException( int code, const QString &source, const QString &desc, const QString &help );
    void slotHandleRulesSaveException( int, const QString &, const QString &, const QString & );

public:
    // account API in OutlookAPI_account.cpp
    QString defaultProfileName() const;
    std::optional< std::map< QString, std::shared_ptr< Outlook::Account > > > getAllAccounts( const QString &profile );

    QString defaultAccountName();
    QString accountName() const;
    bool accountSelected() const;

    std::shared_ptr< Outlook::Account > closeAndSelectAccount( bool notifyOnChange );
    bool selectAccount( const QString &accountName, bool notifyOnChange );
    std::shared_ptr< Outlook::Account > selectAccount( bool notifyOnChange );

    bool connected();

    // rules API in OutlookAPI_rules.cpp
    std::pair< std::shared_ptr< Outlook::Rules >, int > getRules();
    std::shared_ptr< Outlook::Rule > getRule( const std::shared_ptr< Outlook::Rules > &rules, int num );
    std::shared_ptr< Outlook::Rule > findRule( const QString &ruleName );

    std::optional< bool > addRule( const std::shared_ptr< Outlook::Folder > &folder, const std::list< std::pair< QStringList, EFilterType > > &patterns, QStringList &msgs );
    std::optional< bool > addToRule( std::shared_ptr< Outlook::Rule > rule, const std::list< std::pair< QStringList, EFilterType > > &patterns, QStringList &msg, bool copyFirst );
    std::shared_ptr< Outlook::Rule > copyRule( std::shared_ptr< Outlook::Rule > rule );

    static QString getDebugName( const std::shared_ptr< Outlook::Rule > &rule );
    static QString getDebugName( const Outlook::Rule *rule );
    static QString getDebugName( const Outlook::_Rule *rule );
    static QString getDisplayName( const std::shared_ptr< Outlook::Rule > &rule );
    static QString getDisplayName( const Outlook::Rule *rule );
    static QString getDisplayName( const Outlook::_Rule *rule );

    bool ruleEnabled( const std::shared_ptr< Outlook::Rule > &rule );
    bool disableRule( const std::shared_ptr< Outlook::Rule > &rule, bool andSave );
    bool enableRule( const std::shared_ptr< Outlook::Rule > &rule, bool andSave );
    bool deleteRule( std::shared_ptr< Outlook::Rule > rule, bool forceDisable, bool andSave );

    void loadRuleData( QStandardItem *ruleItem, std::shared_ptr< Outlook::Rule > rule, bool force = false );

    QString moveTargetFolderForRule( const std::shared_ptr< Outlook::Rule > &rule ) const;
    std::list< EFilterType > filterTypesForRule( const std::shared_ptr< Outlook::Rule > &rule ) const;

    static QString ruleNameForRule( std::shared_ptr< Outlook::Rule > rule, bool forDisplay = false );
    static std::optional< QString > getDestFolderNameForRule( std::shared_ptr< Outlook::Rule > rule, bool moveOnly );   // if move only false, checks move than copy

    static QString rawRuleNameForRule( std::shared_ptr< Outlook::Rule > rule );

    static QStringList getActionStrings( std::shared_ptr< Outlook::Rule > rule );

    static std::list< QStringList > getConditionalStringList( std::shared_ptr< Outlook::Rule > rule, bool exceptions, EWrapperMode wrapperMode, bool includeSender = false );
    static std::list< QStringList > getActionStringList( std::shared_ptr< Outlook::Rule > rule );

    static bool isEnabled( const std::shared_ptr< Outlook::Rule > &rule );

    bool ruleBeenLoaded( std::shared_ptr< Outlook::Rule > &rule ) const;
    bool ruleLessThan( const std::shared_ptr< Outlook::Rule > &lhsRule, const std::shared_ptr< Outlook::Rule > &rhsRule ) const;

    bool runAllRules( const std::shared_ptr< Outlook::Folder > &folder = {} );
    bool runAllRulesOnAllFolders();
    bool runAllRulesOnTrashFolder();
    bool runAllRulesOnJunkFolder();

    bool runRule( std::shared_ptr< Outlook::Rule > rule, const std::shared_ptr< Outlook::Folder > &folder = {} );

    // run from command line
    bool runAllRules( std::shared_ptr< Outlook::Folder > folder, bool allFolders, bool junk );
    bool runRule( const std::shared_ptr< Outlook::Rule > &rule, std::shared_ptr< Outlook::Folder > folder, bool allFolders, bool junk );

    // tools API in OutlookAPI_tools.cpp
    bool enableAllRules( bool andSave = true, bool *needsSaving = nullptr );
    bool mergeRules( bool andSave = true, bool *needsSaving = nullptr );

    bool moveFromToAddress( bool andSave = true, bool *needsSaving = nullptr );
    bool fixFromMessageHeaderRules( bool andSave = true, bool *needsSaving = nullptr );
    bool renameRules( bool andSave = true, bool *needsSaving = nullptr );
    bool sortRules( bool andSave = true, bool *needsSaving = nullptr );
    bool saveRules();

    // folders API in OutlookAPI_folders.cpp
    using TFolderFunc = std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) >;

    std::shared_ptr< Outlook::Folder > rootFolder();
    QString rootFolderName();

    void setRootFolder( const std::shared_ptr< Outlook::Folder > &folder, bool update = true );

    QString folderDisplayPath( const std::shared_ptr< Outlook::Folder > &folder, bool removeLeadingSlashes = false ) const;
    QString folderDisplayName( const Outlook::Folder *folder );

    std::shared_ptr< Outlook::Folder > findFolder( const QString &folderName, std::shared_ptr< Outlook::Folder > parentFolder );
    std::shared_ptr< Outlook::Folder > getFolder( const Outlook::Folder *item );

    int recursiveSubFolderCount( const Outlook::Folder *parent );

    std::list< std::shared_ptr< Outlook::Folder > > getFolders( const std::shared_ptr< Outlook::Folder > &parent, bool recursive, const TFolderFunc &acceptFolder = {}, const TFolderFunc &checkChildFolders = {} );
    std::shared_ptr< Outlook::Folder > addFolder( const std::shared_ptr< Outlook::Folder > &parent, const QString &folderName );
    std::shared_ptr< Outlook::Folder > parentFolder( const std::shared_ptr< Outlook::Folder > &folder );

    QString nameForFolder( const std::shared_ptr< Outlook::Folder > &folder ) const;
    QString rawPathForFolder( const std::shared_ptr< Outlook::Folder > &folder ) const;

    QString folderDisplayName( const std::shared_ptr< Outlook::Folder > &folder );

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

    std::pair< std::shared_ptr< Outlook::Items >, int > getEmailItemsForRootFolder();

    std::shared_ptr< Outlook::MailItem > getEmailItem( const std::shared_ptr< Outlook::Items > &items, int num );
    std::shared_ptr< Outlook::MailItem > getEmailItem( IDispatch *item );

    static bool isExchangeUser( Outlook::AddressEntry *address );

    static QString getSubject( std::shared_ptr< Outlook::MailItem > mailItem );
    static QString getSubject( Outlook::MailItem *mailItem );

    void displayEmail( const std::shared_ptr< Outlook::MailItem > &email ) const;

    // in the Outlook_emailAddress file
    static QStringList getSenderEmailAddresses( Outlook::MailItem *mailItem );

    static TEmailAddressList getEmailAddresses( std::shared_ptr< Outlook::MailItem > &mailItem, std::optional< EAddressTypes > addressTypes = {}, std::optional< EContactTypes > contactTypes = {} );
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

private:
    // general API in OutlookAPI.cpp
    bool showRuleDialog( std::shared_ptr< Outlook::Rule > rule, bool readOnly );

    std::shared_ptr< Outlook::Application > getApplication();
    std::shared_ptr< Outlook::Application > outlookApp();

    std::shared_ptr< Outlook::Items > getItems( Outlook::_Items *item );
    std::shared_ptr< Outlook::Folder > selectContacts();
    std::shared_ptr< Outlook::Folder > selectInbox();
    template< typename T >
    static Outlook::OlObjectClass getObjectClass( T *item )
    {
        if ( !item )
            return {};
        return item->Class();
    }
    static Outlook::OlObjectClass getObjectClass( IDispatch *item );

    void initSettings();
    void setEmailFilterTypes( EFilterType value );

    template< typename T >
    T connectToException( T obj )
    {
        connect( obj, SIGNAL( exception( int, QString, QString, QString ) ), this, SLOT( slotHandleException( int, const QString &, const QString &, const QString & ) ) );
        return obj;
    }

    template< typename T >
    std::shared_ptr< T > connectToException( std::shared_ptr< T > obj )
    {
        connect( obj.get(), SIGNAL( exception( int, QString, QString, QString ) ), this, SLOT( slotHandleException( int, const QString &, const QString &, const QString & ) ) );
        return obj;
    }

    QWidget *fParentWidget{ nullptr };
    std::shared_ptr< Outlook::Application > fOutlookApp{ nullptr };
    std::shared_ptr< Outlook::NameSpace > fSession{ nullptr };
    std::shared_ptr< Outlook::Account > fAccount{ nullptr };
    std::shared_ptr< Outlook::Folder > fInbox{ nullptr };
    std::shared_ptr< Outlook::Folder > fRootFolder{ nullptr };   // used for loading emails
    std::shared_ptr< Outlook::Folder > fJunkFolder{ nullptr };
    std::shared_ptr< Outlook::Folder > fTrashFolder{ nullptr };
    std::shared_ptr< Outlook::Folder > fContacts{ nullptr };
    std::shared_ptr< Outlook::Rules > fRules{ nullptr };

    static std::shared_ptr< COutlookAPI > sInstance;

    bool fLoggedIn{ false };
    bool fCanceled{ false };
    bool fIgnoreExceptions{ false };
    bool fOnlyProcessUnread{ true };
    bool fProcessAllEmailWhenLessThan200Emails{ true };
    bool fOnlyProcessTheFirst500Emails{ true };
    bool fIncludeJunkFolderWhenRunningOnAllFolders{ false };
    bool fIncludeDeletedFolderWhenRunningOnAllFolders{ false };
    bool fDisableRatherThanDeleteRules{ false };
    bool fLoadAccountInfo{ true };
    QString fLastAccountName;
    std::list< EFilterType > fEmailFilterTypes{ EFilterType::eByEmailAddressContains };
    QStringList fRulesToSkip;
    bool fSaveRulesSuccess{ true };

    // account API in OutlookAPI_account.cpp
private:
    std::shared_ptr< Outlook::Account > getAccount( Outlook::_Account *item );
    std::shared_ptr< Outlook::NameSpace > getNamespace( Outlook::_NameSpace *ns );

    // rules api in Outlook_rules.cpp
private:
    // rules API in OutlookAPI_rules.cpp
    std::shared_ptr< Outlook::Rules > selectRules();
    bool skipRule( const std::shared_ptr< Outlook::Rule > &rule ) const;

    std::shared_ptr< Outlook::Rules > getRules( Outlook::Rules *item );
    std::shared_ptr< Outlook::Rule > getRule( Outlook::_Rule *item );

    std::optional< QStringList > getRecipients( Outlook::Rule *rule, QStringList *msgs );

    std::vector< std::shared_ptr< Outlook::Rule > > getAllRules();

    bool runRules( std::vector< std::shared_ptr< Outlook::Rule > > rules, std::shared_ptr< Outlook::Folder > folder = {}, bool recursive = false, const std::optional< QString > &perFolderMsg = {} );

    bool addRecipientsToRule( Outlook::Rule *rule, const QStringList &recipients, QStringList &msgs );
    bool addRecipientsToRule( Outlook::Rule *rule, const TEmailAddressList &recipients, QStringList &msgs );
    bool addDisplayNamesToRule( Outlook::Rule *rule, const QStringList &displayNames, QStringList &msgs );
    bool addSubjectsToRule( Outlook::Rule *rule, const QStringList &subjects, QStringList &msgs );
    bool addSenderToRule( Outlook::Rule *rule, const TEmailAddressList &senders, QStringList &msgs );
    bool addSenderToRule( Outlook::Rule *rule, const QStringList &senders, QStringList &msgs );

    std::unordered_set< std::shared_ptr< Outlook::Rule > > fRuleBeenLoaded;

    // tools API in OutlookAPI_tools.cpp
private:
    using TRuleList = std::list< std::shared_ptr< Outlook::Rule > >;
    using TRulePair = std::pair< std::shared_ptr< Outlook::Rule >, TRuleList >;
    using TMergeRuleMap = std::multimap< QString, TRulePair >;

    std::optional< QString > mergeKey( const std::shared_ptr< Outlook::Rule > &rule ) const;
    std::optional< QStringList > mergeRecipients( Outlook::Rule *lhs, Outlook::Rule *rhs, QStringList *msgs );
    std::optional< QStringList > mergeRecipients( Outlook::Rule *lhs, const QStringList &rhs, QStringList *msgs );
    std::optional< QStringList > mergeRecipients( Outlook::Rule *lhs, const TEmailAddressList &rhs, QStringList *msgs );
    std::optional< QStringList > mergeRecipients( const std::list< Outlook::Rule * > &rules, QStringList *msgs );

    static QString stripHeaderStringString( const QString &msg );   // can include the "From" portion of the pattern as well as quotes, returns the raw pattern
    static QStringList getFromMessageHeaderString( const QString &address );   // can include the "From" portion of the pattern as well as quotes, calls stripHeaderStringString first, returns both proper patterns From: "XXX" and From: XXX
    static QStringList getFromMessageHeaderStrings( const QStringList &msgs );   // can include the "From" portion of the pattern as well as quotes, returns the raw pattern, cleans each string first then gets all the proper patterns
    bool canMergeRules( std::shared_ptr< Outlook::Rule > lhs, std::shared_ptr< Outlook::Rule > rhs );

    void mergeRules( TRulePair &rules );
    std::shared_ptr< Outlook::Rule > mergeRule( std::shared_ptr< Outlook::Rule > &lhs, std::shared_ptr< Outlook::Rule > &rhs );
    TMergeRuleMap findMergableRules();

private:
    // folders API in OutlookAPI_folders.cpp
    bool isFolder( const std::shared_ptr< Outlook::Folder > &folder, const QString &path ) const;
    bool emptyFolder( std::shared_ptr< Outlook::Folder > &folder );

    std::shared_ptr< Outlook::Folder > getFolder( const Outlook::MAPIFolder *item );

    void setRootFolder( const QString &folderName, bool update = true );

    std::shared_ptr< Outlook::Folder > getDefaultFolder( Outlook::OlDefaultFolders folderType );
    std::pair< std::shared_ptr< Outlook::Folder >, bool > getMailFolder( const QString &folderLabel, const QString &fullPath, bool singleOnly );   // full path after \\account
    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectFolder( const QString &folderName, const TFolderFunc &acceptFolder, const TFolderFunc &checkChildFolders, bool singleOnly );
    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectFolder( const QString &folderName, const std::list< std::shared_ptr< Outlook::Folder > > &folders, bool singleOnly );

    std::list< std::shared_ptr< Outlook::Folder > > getFolders( bool recursive, const TFolderFunc &acceptFolder = {}, const TFolderFunc &checkChildFolders = {} );

    static QString ruleNameForFolder( Outlook::Folder *folder );
    static QString ruleNameForFolder( const std::shared_ptr< Outlook::Folder > &folder );

    int subFolderCount( const Outlook::Folder *parent, bool recursive );

    // dump API in OutlookAPI_dump.cpp
    void dumpFolder( Outlook::Folder *root );
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