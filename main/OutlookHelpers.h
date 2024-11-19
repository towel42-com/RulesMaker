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
    class Folder;
    class NameSpace;
    class Account;
    class MailItem;
    class AddressEntry;
    class AddressEntries;
    class Recipient;
    class Rules;
    class Rule;
    class Recipients;
    class AddressList;
    enum class OlImportance;
    enum class OlItemType;
    enum class OlMailRecipientType;
    enum class OlObjectClass;
    enum class OlRuleConditionType;
    enum class OlSensitivity;
    enum class OlMarkInterval;
}

class COutlookHelpers : public QObject
{
    Q_OBJECT;

public:
    friend class COutlookSetup;
    COutlookHelpers();
    static std::shared_ptr< COutlookHelpers > getInstance();
    virtual ~COutlookHelpers();

    std::shared_ptr< Outlook::Account > selectAccount( bool notifyOnChange, QWidget *parent );
    void logout( bool andNotify );

    bool accountSelected() const;
    std::shared_ptr< Outlook::Folder > getInbox( QWidget *parent );
    std::shared_ptr< Outlook::Folder > getContacts( QWidget *parent );
    std::shared_ptr< Outlook::Rules > getRules( QWidget *parent );

    std::shared_ptr< Outlook::Folder > rootFolder();
    void setRootFolder( std::shared_ptr< Outlook::Folder > folder ) { fRootFolder = folder; }

    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectFolder( QWidget *parent, const QString &folderName, std::function< bool( std::shared_ptr< Outlook::Folder > folder ) > acceptFolder, std::function< bool( std::shared_ptr< Outlook::Folder > folder ) > checkChildFolders, bool singleOnly );
    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectFolder( QWidget *parent, const QString &folderName, const std::list< std::shared_ptr< Outlook::Folder > > &folders, bool singleOnly );

    std::list< std::shared_ptr< Outlook::Folder > > getFolders( bool recursive, std::function< bool( std::shared_ptr< Outlook::Folder > folder ) > acceptFolder = {}, std::function< bool( std::shared_ptr< Outlook::Folder > folder ) > checkChildFolders = {} );
    std::list< std::shared_ptr< Outlook::Folder > > getFolders( std::shared_ptr< Outlook::Folder > parent, bool recursive, std::function< bool( std::shared_ptr< Outlook::Folder > folder ) > acceptFolder = {}, std::function< bool( std::shared_ptr< Outlook::Folder > folder ) > checkChildFolders = {} );

    bool addRule( const QString &destFolder, const QStringList &rules, QStringList &msg );
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
    QString COutlookHelpers::ruleNameForFolder( std::shared_ptr< Outlook::Folder > folder );

Q_SIGNALS:
    void sigAccountChanged();

private:
    bool addRecipientsToRule( Outlook::Rule *rule, const QStringList &recipients, QStringList &msgs );

    std::optional< QStringList > mergeRecipients( Outlook::Rule *lhs, Outlook::Rule *rhs, QStringList *msgs );
    std::optional< QStringList > mergeRecipients( Outlook::Rule *lhs, const QStringList & rhs, QStringList *msgs );
    std::optional< QStringList > getRecipients( Outlook::Rule *rule, QStringList *msgs );

    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectInbox( QWidget *parent, bool singleOnly );
    std::pair< std::shared_ptr< Outlook::Folder >, bool > selectContacts( QWidget *parent, bool singleOnly );
    std::shared_ptr< Outlook::Rules > selectRules( QWidget *parent );

    std::shared_ptr< Outlook::Application > fOutlookApp;
    std::shared_ptr< Outlook::Account > fAccount;
    std::shared_ptr< Outlook::Folder > fInbox;
    std::shared_ptr< Outlook::Folder > fRootFolder;   // used for loading emails
    std::shared_ptr< Outlook::Folder > fContacts;
    std::shared_ptr< Outlook::Rules > fRules;

    static std::shared_ptr< COutlookHelpers > sInstance;

    bool fLoggedIn{ false };
};

QString toString( Outlook::OlItemType olItemType );
QString toString( Outlook::OlRuleConditionType olItemType );
QString toString( Outlook::OlImportance importance );
QString toString( Outlook::OlSensitivity sensitivity );
QString toString( Outlook::OlMarkInterval markInterval );
QString getValue( const QVariant &variant, const QString &joinSeparator );

void dumpMetaMethods( QObject *object );

#endif