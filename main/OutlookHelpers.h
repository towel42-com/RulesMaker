#ifndef OUTLOOKHELPERS_H
#define OUTLOOKHELPERS_H

#include <QObject>
#include <memory>
#include <list>
#include <functional>
#include <optional>
#include <QString>

class QWidget;
struct IDispatch;

namespace Outlook
{
    class Application;
    class MAPIFolder;
    class NameSpace;
    class Account;
    class MailItem;
    class AddressEntry;
    class AddressEntries;
    class Recipient;
    class Rules;
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
    std::shared_ptr< Outlook::MAPIFolder > getInbox( QWidget *parent );
    std::shared_ptr< Outlook::MAPIFolder > getContacts( QWidget *parent );
    std::shared_ptr< Outlook::Rules > getRules( QWidget *parent );

    std::pair< std::shared_ptr< Outlook::MAPIFolder >, bool > selectFolder( QWidget *parent, const QString &folderName, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > checkChildFolders, bool singleOnly );
    std::pair< std::shared_ptr< Outlook::MAPIFolder >, bool > selectFolder( QWidget *parent, const QString &folderName, const std::list< std::shared_ptr< Outlook::MAPIFolder > > &folders, bool singleOnly );

    std::list< std::shared_ptr< Outlook::MAPIFolder > > getFolders( bool recursive, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder = {}, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > checkChildFolders = {} );
    std::list< std::shared_ptr< Outlook::MAPIFolder > > getFolders( std::shared_ptr< Outlook::MAPIFolder > parent, bool recursive, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder = {}, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > checkChildFolders = {} );

    bool addRule( const QString &destFolder, const QStringList &rules, QStringList &msg );
    template< typename T >
    static Outlook::OlObjectClass getObjectClass( T *item )
    {
        if ( !item )
            return {};
        return item->Class();
    }
    static Outlook::OlObjectClass getObjectClass( IDispatch *item );
    static QString getSenderEmailAddress( Outlook::MailItem *mailItem );
    static QStringList getRecipientEmails( Outlook::MailItem *mailItem, Outlook::OlMailRecipientType recipientType );
    static QStringList getRecipientEmails( Outlook::Recipients *recipients, std::optional< Outlook::OlMailRecipientType > recipientType );

    static QStringList getEmailAddresses( Outlook::AddressList *addresses );
    static QStringList getEmailAddresses( Outlook::AddressEntry *address );
    static QStringList getEmailAddresses( Outlook::AddressEntries *entries );

    static QString getEmailAddress( Outlook::Recipient *recipient );

    void dumpSession( Outlook::NameSpace &session );
    void dumpFolder( Outlook::MAPIFolder *root );

    std::shared_ptr< Outlook::Application > outlookApp() { return fOutlookApp; }

Q_SIGNALS:
    void sigAccountChanged();

private:
    std::pair< std::shared_ptr< Outlook::MAPIFolder >, bool > selectInbox( QWidget *parent, bool singleOnly );
    std::pair< std::shared_ptr< Outlook::MAPIFolder >, bool > selectContacts( QWidget *parent, bool singleOnly );
    std::shared_ptr< Outlook::Rules > selectRules( QWidget *parent );

    std::shared_ptr< Outlook::Application > fOutlookApp;
    std::shared_ptr< Outlook::Account > fAccount;
    std::shared_ptr< Outlook::MAPIFolder > fInbox;
    std::shared_ptr< Outlook::MAPIFolder > fContacts;
    std::shared_ptr< Outlook::Rules > fRules;

    static std::shared_ptr< COutlookHelpers > sInstance;

    bool fLoggedIn{ false };
};

QString toString( Outlook::OlItemType olItemType );
QString toString( Outlook::OlRuleConditionType olItemType );
QString toString( Outlook::OlImportance importance );
QString toString( Outlook::OlSensitivity sensitivity );
QString toString( Outlook::OlMarkInterval markInterval );

void dumpMetaMethods( QObject *object );

#endif