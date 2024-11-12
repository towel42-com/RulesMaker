#ifndef OUTLOOKHELPERS_H
#define OUTLOOKHELPERS_H

//#include "msoutl.h"

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
    class Recipient;
    enum class OlItemType;
    enum class OlMailRecipientType;
    enum class OlObjectClass;
}

class COutlookHelpers
{
public:
    friend class COutlookSetup;
    COutlookHelpers();
    static std::shared_ptr< COutlookHelpers > getInstance();
    virtual ~COutlookHelpers();

    std::shared_ptr< Outlook::MAPIFolder > getInbox( QWidget *parent );
    std::shared_ptr< Outlook::MAPIFolder > getContacts( QWidget *parent );

    std::shared_ptr< Outlook::Account > selectAccount( QWidget *parent );

    std::pair< std::shared_ptr< Outlook::MAPIFolder >, bool > selectFolder( QWidget *parent, const QString &folderName, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder, bool singleOnly );
    std::pair< std::shared_ptr< Outlook::MAPIFolder >, bool > selectFolder( QWidget *parent, const QString &folderName, const std::list< std::shared_ptr< Outlook::MAPIFolder > > &folders, bool singleOnly );

    std::list< std::shared_ptr< Outlook::MAPIFolder > > getFolders( std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder = {} );
    std::list< std::shared_ptr< Outlook::MAPIFolder > > getFolders( std::shared_ptr< Outlook::MAPIFolder > parent, bool recursive, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder = {} );

    template< typename T >
    static Outlook::OlObjectClass getObjectClass( T *item )
    {
        if ( !item )
            return {};
        return item->Class();
    }
    static Outlook::OlObjectClass getObjectClass( IDispatch *item );
    static QString getSenderEmailAddress( Outlook::MailItem *mailItem );
    static QStringList getRecipients( Outlook::MailItem *mailItem, Outlook::OlMailRecipientType recipientType );

    static QString getEmailAddress( Outlook::Recipient *recipient );

    static bool isExchangeUser( Outlook::AddressEntry *address );
    static std::optional< QString > getEmailAddress( Outlook::AddressEntry *address );

    void dumpSession( Outlook::NameSpace &session );
    void dumpFolder( Outlook::MAPIFolder *root );

    std::shared_ptr< Outlook::Application > outlook() { return fOutlook; }

    static QString toString( Outlook::OlItemType olItemType );

private:
    std::pair< std::shared_ptr< Outlook::MAPIFolder >, bool > selectInbox( QWidget *parent, bool singleOnly );
    std::pair< std::shared_ptr< Outlook::MAPIFolder >, bool > selectContacts( QWidget *parent, bool singleOnly );

    std::shared_ptr< Outlook::Application > fOutlook;
    std::shared_ptr< Outlook::Account > fAccount;
    std::shared_ptr< Outlook::MAPIFolder > fInbox;
    std::shared_ptr< Outlook::MAPIFolder > fContacts;

    static std::shared_ptr< COutlookHelpers > sInstance;
};

#endif