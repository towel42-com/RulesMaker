#ifndef OUTLOOKHELPERS_H
#define OUTLOOKHELPERS_H

#include <QObject>
#include <memory>
#include <list>
#include <functional>
#include <optional>
#include <QString>

class QVariant;
class QWidget;
struct IDispatch;

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
}

class COutlookAPI : public QObject
{
    Q_OBJECT;

public:
    friend class COutlookSetup;
    COutlookAPI();
    static std::shared_ptr< COutlookAPI > getInstance();
    virtual ~COutlookAPI();

    std::shared_ptr< Outlook::Account > selectAccount( bool notifyOnChange, QWidget *parent );
    void logout( bool andNotify );

    QString accountName() const;

    bool accountSelected() const;
    std::shared_ptr< Outlook::Folder > getInbox( QWidget *parent );
    std::shared_ptr< Outlook::Folder > getContacts( QWidget *parent );
    std::shared_ptr< Outlook::Rules > getRules( QWidget *parent );

    std::shared_ptr< Outlook::Folder > rootFolder();
    void setRootFolder( const std::shared_ptr< Outlook::Folder > &folder ) { fRootFolder = folder; }

    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectFolder( QWidget *parent, const QString &folderName, std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) > acceptFolder, std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) > checkChildFolders, bool singleOnly );
    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectFolder( QWidget *parent, const QString &folderName, const std::list< std::shared_ptr< Outlook::Folder > > &folders, bool singleOnly );

    std::list< std::shared_ptr< Outlook::Folder > > getFolders( bool recursive, std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) > acceptFolder = {}, std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) > checkChildFolders = {} );
    std::list< std::shared_ptr< Outlook::Folder > > getFolders( const std::shared_ptr< Outlook::Folder > &parent, bool recursive, std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) > acceptFolder = {}, std::function< bool( const std::shared_ptr< Outlook::Folder > &folder ) > checkChildFolders = {} );

    int subFolderCount( const std::shared_ptr< Outlook::Folder > &parent, std::function< bool( Outlook::Folder *folder ) > acceptFolder = {} );
    std::pair< std::shared_ptr< Outlook::Rule >, bool > addRule( const QString &destFolder, const QStringList &rules, QStringList &msg );
    bool addToRule( std::shared_ptr< Outlook::Rule > rule, const QStringList &rules, QStringList &msg );

    void renameRules();

    void sortRules();
    void moveFromToAddress();
    void mergeRules();

    bool execute( std::shared_ptr< Outlook::Rule > rule );
    bool execute( Outlook::Rule *rule );

    template< typename T >
    static Outlook::OlObjectClass getObjectClass( T *item )
    {
        if ( !item )
            return {};
        return item->Class();
    }
    static Outlook::OlObjectClass getObjectClass( IDispatch *item );
    static QStringList getSenderEmailAddresses( Outlook::MailItem *mailItem );
    static QStringList getRecipientEmails( Outlook::MailItem *mailItem, Outlook::OlMailRecipientType recipientType );
    static QStringList getRecipientEmails( Outlook::Recipients *recipients, std::optional< Outlook::OlMailRecipientType > recipientType );

    static QStringList getEmailAddresses( Outlook::AddressList *addresses );
    static QStringList getEmailAddresses( Outlook::AddressEntry *address );
    static QStringList getEmailAddresses( Outlook::AddressEntries *entries );

    static QString getEmailAddress( Outlook::Recipient *recipient );

    void dumpSession( Outlook::NameSpace &session );
    void dumpFolder( Outlook::Folder *root );

    std::shared_ptr< Outlook::Application > outlookApp() { return fOutlookApp; }

    QString ruleNameForFolder( Outlook::Folder *folder );
    QString ruleNameForFolder( const std::shared_ptr< Outlook::Folder > &folder );

    QString folderName( Outlook::Folder *folder );
    QString folderName( const std::shared_ptr< Outlook::Folder > &folder );

    std::shared_ptr< Outlook::Rule > getRule( Outlook::_Rule *item );
    std::shared_ptr< Outlook::Items > getItems( Outlook::_Items *item );
    std::shared_ptr< Outlook::MailItem > getMailItem( IDispatch *item );
    std::shared_ptr< Outlook::Folder > getFolder( Outlook::Folder *item );

    bool canceled() const { return fCanceled; }
Q_SIGNALS:
    void sigAccountChanged();
    void sigInitStatus( const QString & label, int max );
    void sigSetStatus( const QString &label, int curr, int max );
    void sigIncStatusValue( const QString &label );

public Q_SLOTS:
    void slotHandleException( int code, const QString &source, const QString &desc, const QString &help );
    void slotCanceled() { fCanceled = true; }
    void slotClearCanceled() { fCanceled = false; }

private:
    std::shared_ptr< Outlook::Application > getApplication();
    std::shared_ptr< Outlook::Account > getAccount( Outlook::_Account *item );
    std::shared_ptr< Outlook::Rules > getRules( Outlook::Rules *item );
    std::shared_ptr< Outlook::Folder > getFolder( Outlook::MAPIFolder *item );

    std::optional< QString > ruleNameForRule( std::shared_ptr< Outlook::Rule > rule );
    bool addRecipientsToRule( Outlook::Rule *rule, const QStringList &recipients, QStringList &msgs );

    std::optional< QStringList > mergeRecipients( Outlook::Rule *lhs, Outlook::Rule *rhs, QStringList *msgs );
    std::optional< QStringList > mergeRecipients( Outlook::Rule *lhs, const QStringList &rhs, QStringList *msgs );
    std::optional< QStringList > getRecipients( Outlook::Rule *rule, QStringList *msgs );

    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectInbox( QWidget *parent, bool singleOnly );
    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectContacts( QWidget *parent, bool singleOnly );
    std::shared_ptr< Outlook::Rules > selectRules( QWidget *parent );

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

    std::shared_ptr< Outlook::Application > fOutlookApp{ nullptr };
    std::shared_ptr< Outlook::Account > fAccount{ nullptr };
    std::shared_ptr< Outlook::Folder > fInbox{ nullptr };
    std::shared_ptr< Outlook::Folder > fRootFolder{ nullptr };   // used for loading emails
    std::shared_ptr< Outlook::Folder > fContacts{ nullptr };
    std::shared_ptr< Outlook::Rules > fRules{ nullptr };

    static std::shared_ptr< COutlookAPI > sInstance;

    bool fLoggedIn{ false };
    bool fCanceled{ false };
};

QString toString( Outlook::OlItemType olItemType );
QString toString( Outlook::OlRuleConditionType olItemType );
QString toString( Outlook::OlImportance importance );
QString toString( Outlook::OlSensitivity sensitivity );
QString toString( Outlook::OlMarkInterval markInterval );
QString getValue( const QVariant &variant, const QString &joinSeparator );

void dumpMetaMethods( QObject *object );

#endif