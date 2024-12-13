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

class QVariant;
class QWidget;
struct IDispatch;
class QStandardItem;

namespace Outlook
{
    class Application;
    class _Application;
    class NameSpace;
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

class COutlookAPI : public QObject
{
    Q_OBJECT;
    struct SPrivate
    {
        explicit SPrivate() = default;
    };

public:
    // general API in OutlookAPI.cpp
    friend class COutlookSetup;
    COutlookAPI( QWidget *parentWidget, SPrivate pri );

    static std::shared_ptr< COutlookAPI > instance( QWidget *parentWidget = nullptr );
    virtual ~COutlookAPI();

    bool canceled() const { return fCanceled; }

    void logout( bool andNotify );

    std::shared_ptr< Outlook::Folder > getContacts();
    std::shared_ptr< Outlook::Folder > getInbox();
    std::shared_ptr< Outlook::Folder > getJunkFolder();

    void setOnlyProcessUnread( bool value, bool update = true );
    bool onlyProcessUnread() const { return fOnlyProcessUnread; }

    void setProcessAllEmailWhenLessThan200Emails( bool value, bool update = true );
    bool processAllEmailWhenLessThan200Emails() const { return fProcessAllEmailWhenLessThan200Emails; }

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

public:
    // account API in OutlookAPI_account.cpp
    QString accountName() const;
    bool accountSelected() const;
    std::shared_ptr< Outlook::Account > selectAccount( bool notifyOnChange );

    // rules API in OutlookAPI_rules.cpp
    std::pair< std::shared_ptr< Outlook::Rules >, int > getRules();
    std::shared_ptr< Outlook::Rule > getRule( const std::shared_ptr< Outlook::Rules > &rules, int num );

    bool addRule( const std::shared_ptr< Outlook::Folder > &folder, const QStringList &rules, QStringList &msgs );
    bool addToRule( std::shared_ptr< Outlook::Rule > rule, const QStringList &rules, QStringList &msg );

    bool deleteRule( std::shared_ptr< Outlook::Rule > rule );

    void loadRuleData( QStandardItem *ruleItem, std::shared_ptr< Outlook::Rule > rule );

    QString moveTargetFolderForRule( const std::shared_ptr< Outlook::Rule > &rule ) const;
    static QString ruleNameForRule( std::shared_ptr< Outlook::Rule > rule, bool = false );

    bool ruleBeenLoaded( std::shared_ptr< Outlook::Rule > &rule ) const;
    bool ruleLessThan( const std::shared_ptr< Outlook::Rule > &lhsRule, const std::shared_ptr< Outlook::Rule > &rhsRule ) const;

    bool runAllRules( const std::shared_ptr< Outlook::Folder > &folder = {} );
    bool runAllRulesOnAllFolders();
    bool runRule( std::shared_ptr< Outlook::Rule > rule, const std::shared_ptr< Outlook::Folder > &folder = {} );

    // tools
    bool enableAllRules();
    bool mergeRules();
    bool moveFromToAddress();
    bool renameRules();
    bool sortRules();
    void saveRules();

    // folders API in OutlookAPI_folders.cpp
    using TFolderFunc = std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) >;

    std::shared_ptr< Outlook::Folder > rootFolder();
    QString rootFolderName();

    void setRootFolder( const std::shared_ptr< Outlook::Folder > &folder, bool update = true );

    std::shared_ptr< Outlook::Folder > getFolder( const Outlook::Folder *item );

    int recursiveSubFolderCount( const Outlook::Folder *parent );

    std::list< std::shared_ptr< Outlook::Folder > > getFolders( const std::shared_ptr< Outlook::Folder > &parent, bool recursive, const TFolderFunc &acceptFolder = {}, const TFolderFunc &checkChildFolders = {} );
    std::shared_ptr< Outlook::Folder > addFolder( const std::shared_ptr< Outlook::Folder > &parent, const QString &folderName );
    std::shared_ptr< Outlook::Folder > parentFolder( const std::shared_ptr< Outlook::Folder > &folder );

    QString nameForFolder( const std::shared_ptr< Outlook::Folder > &folder ) const;
    QString rawPathForFolder( const std::shared_ptr< Outlook::Folder > &folder ) const;

    QString folderDisplayName( const std::shared_ptr< Outlook::Folder > &folder );

    // email API in OutlookAPI_email.cpp
    enum EAddressTypes
    {
        eNone = 0x00,
        eOriginator = 0x01,
        eTo = 0x02,
        eCC = 0x04,
        eBCC = 0x08,
        eAllRecipients = eOriginator | eTo | eCC | eBCC,
        eSender = 0x10,
        eAllEmailAddresses = eAllRecipients | eSender,
        eSMTPOnly = 0x20
    };
    using TStringListPair = std::pair< QStringList, QStringList >;

    std::pair< std::shared_ptr< Outlook::Items >, int > getEmailItemsForRootFolder();

    std::shared_ptr< Outlook::MailItem > getEmailItem( const std::shared_ptr< Outlook::Items > &items, int num );
    std::shared_ptr< Outlook::MailItem > getEmailItem( IDispatch *item );

    static TStringListPair getEmailAddresses( std::shared_ptr< Outlook::MailItem > &mailItem, EAddressTypes types );   // returns the list of email addresses, display names
    static TStringListPair getEmailAddresses( Outlook::MailItem *mailItem, EAddressTypes types );   // returns the list of email addresses, display names
    static QStringList getEmailAddresses( Outlook::MailItem *mailItem, Outlook::OlMailRecipientType recipientType, bool smtpOnly );
    static QStringList getSenderEmailAddresses( Outlook::MailItem *mailItem );

    static TStringListPair getEmailAddresses( Outlook::Recipients *recipients, EAddressTypes types );
    static QStringList getEmailAddresses( Outlook::Recipients *recipients, std::optional< Outlook::OlMailRecipientType > recipientType, bool smtpOnly );

    static TStringListPair getEmailAddresses( Outlook::Recipient *recipient, EAddressTypes types );

    static TStringListPair getEmailAddresses( Outlook::AddressList *addresses, EAddressTypes types );   // returns the list of email addresses, display names, types is used for SMTP only
    static QStringList getEmailAddresses( Outlook::AddressList *addresses, bool smtpOnly );

    static TStringListPair getEmailAddresses( Outlook::AddressEntries *entries, EAddressTypes types );   // returns the list of email addresses, display names, types is used for SMTP only
    static QStringList getEmailAddresses( Outlook::AddressEntries *entries, bool smtpOnly );

    void displayEmail( const std::shared_ptr< Outlook::MailItem > &email ) const;

    // dump API in OutlookAPI_dump.cpp
    void dumpSession( Outlook::NameSpace &session );
    void dumpMetaMethods( QObject *object );

private:
    // general API in OutlookAPI.cpp
    std::shared_ptr< Outlook::Application > getApplication();
    std::shared_ptr< Outlook::Application > outlookApp();

    std::shared_ptr< Outlook::Items > getItems( Outlook::_Items *item );
    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectContacts( bool singleOnly );
    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectInbox( bool singleOnly );
    template< typename T >
    static Outlook::OlObjectClass getObjectClass( T *item )
    {
        if ( !item )
            return {};
        return item->Class();
    }
    static Outlook::OlObjectClass getObjectClass( IDispatch *item );

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
    std::shared_ptr< Outlook::Account > fAccount{ nullptr };
    std::shared_ptr< Outlook::Folder > fInbox{ nullptr };
    std::shared_ptr< Outlook::Folder > fRootFolder{ nullptr };   // used for loading emails
    std::shared_ptr< Outlook::Folder > fJunkFolder{ nullptr };
    std::shared_ptr< Outlook::Folder > fContacts{ nullptr };
    std::shared_ptr< Outlook::Rules > fRules{ nullptr };

    static std::shared_ptr< COutlookAPI > sInstance;

    bool fLoggedIn{ false };
    bool fCanceled{ false };
    bool fOnlyProcessUnread{ true };
    bool fProcessAllEmailWhenLessThan200Emails{ true };

    // account API in OutlookAPI_account.cpp
    std::shared_ptr< Outlook::Account > getAccount( Outlook::_Account *item );

    // rules api in Outlook_rules.cpp
private:
    // rules API in OutlookAPI_rules.cpp
    std::shared_ptr< Outlook::Rules > selectRules();

    std::shared_ptr< Outlook::Rules > getRules( Outlook::Rules *item );
    std::shared_ptr< Outlook::Rule > getRule( Outlook::_Rule *item );

    std::optional< QStringList > getRecipients( Outlook::Rule *rule, QStringList *msgs );

    std::vector< std::shared_ptr< Outlook::Rule > > getAllRules();

    bool runRules( std::vector< std::shared_ptr< Outlook::Rule > > rules, std::shared_ptr< Outlook::Folder > folder = {}, bool recursive = false, const std::optional< QString > &perFolderMsg = {} );

    bool addRecipientsToRule( Outlook::Rule *rule, const QStringList &recipients, QStringList &msgs );

    std::optional< QStringList > mergeRecipients( Outlook::Rule *lhs, Outlook::Rule *rhs, QStringList *msgs );
    std::optional< QStringList > mergeRecipients( Outlook::Rule *lhs, const QStringList &rhs, QStringList *msgs );

    std::unordered_set< std::shared_ptr< Outlook::Rule > > fRuleBeenLoaded;

private:
    // folders API in OutlookAPI_folders.cpp
    bool isFolder( const std::shared_ptr< Outlook::Folder > &folder, const QString &path ) const;

    std::shared_ptr< Outlook::Folder > getFolder( const Outlook::MAPIFolder *item );

    void setRootFolder( const QString &folderName, bool update = true );

    QString folderDisplayPath( const std::shared_ptr< Outlook::Folder > &folder, bool removeLeadingSlashes = false ) const;
    QString folderDisplayName( const Outlook::Folder *folder );

    std::pair< std::shared_ptr< Outlook::Folder >, bool > getMailFolder( const QString &folderLabel, const QString &fullPath, bool singleOnly );   // full path after \\account
    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectFolder( const QString &folderName, const TFolderFunc &acceptFolder, const TFolderFunc &checkChildFolders, bool singleOnly );
    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectFolder( const QString &folderName, const std::list< std::shared_ptr< Outlook::Folder > > &folders, bool singleOnly );

    std::list< std::shared_ptr< Outlook::Folder > > getFolders( bool recursive, const TFolderFunc &acceptFolder = {}, const TFolderFunc &checkChildFolders = {} );

    static QString ruleNameForFolder( Outlook::Folder *folder );
    static QString ruleNameForFolder( const std::shared_ptr< Outlook::Folder > &folder );

    int subFolderCount( const Outlook::Folder *parent, bool recursive );

    // email API in OutlookAPI_email.cpp
    static TStringListPair getEmailAddresses( Outlook::AddressEntry *address, EAddressTypes types );   // returns the list of email addresses, display names, types is used for SMTP only
    static QStringList getEmailAddresses( Outlook::AddressEntry *address, bool smtpOnly );

    // dump API in OutlookAPI_dump.cpp
    void dumpFolder( Outlook::Folder *root );
};

// toString API in OutlookAPI_toString.cpp
QString toString( const QVariant &variant, const QString &joinSeparator );
QString toString( Outlook::OlItemType olItemType );
QString toString( Outlook::OlRuleConditionType olItemType );
QString toString( Outlook::OlImportance importance );
QString toString( Outlook::OlSensitivity sensitivity );
QString toString( Outlook::OlMarkInterval markInterval );

// general API in OutlookAPI.cpp
COutlookAPI::EAddressTypes operator|( const COutlookAPI::EAddressTypes &lhs, const COutlookAPI::EAddressTypes &rhs );
COutlookAPI::EAddressTypes getAddressTypes( bool smtpOnly );
COutlookAPI::EAddressTypes getAddressTypes( std::optional< Outlook::OlMailRecipientType > recipientType, bool smtpOnly );

#endif