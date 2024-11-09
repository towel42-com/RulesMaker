#ifndef OUTLOOKHELPERS_H
#define OUTLOOKHELPERS_H

//#include "msoutl.h"

#include <memory>
#include <list>
#include <functional>
#include <QString>

class QWidget;

namespace Outlook
{
    class Application;
    class MAPIFolder;
    class NameSpace;
    enum class OlItemType;
}

class COutlookHelpers
{
public:
    COutlookHelpers();
    static std::shared_ptr< COutlookHelpers > getInstance();
    virtual ~COutlookHelpers();

    std::shared_ptr< Outlook::MAPIFolder > selectInboxFolder( QWidget *parent );
    std::shared_ptr< Outlook::MAPIFolder > selectContactFolder( QWidget *parent );

    std::shared_ptr< Outlook::MAPIFolder > selectFolder( QWidget *parent, const Outlook::OlItemType &itemType, const QString &folderName, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder );
    std::shared_ptr< Outlook::MAPIFolder > selectFolder( QWidget *parent, const QString &folderName, const std::list< std::shared_ptr< Outlook::MAPIFolder > > &folders );
    std::list< std::shared_ptr< Outlook::MAPIFolder > > getFolders( Outlook::OlItemType itemType, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder = {} );
    std::list< std::shared_ptr< Outlook::MAPIFolder > > getFolders( std::shared_ptr< Outlook::MAPIFolder > parent, bool recursive, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder = {} );
    std::list< std::shared_ptr< Outlook::MAPIFolder > > getFolders( Outlook::OlItemType itemType, std::shared_ptr< Outlook::MAPIFolder > parent, bool recursive, std::function< bool( std::shared_ptr< Outlook::MAPIFolder > folder ) > acceptFolder = {} );

    void dumpSession( Outlook::NameSpace &session );
    void dumpFolder( Outlook::MAPIFolder *root );

    std::shared_ptr< Outlook::Application > outlook() { return fOutlook; }

    static QString toString( Outlook::OlItemType olItemType );

private:
    std::shared_ptr< Outlook::Application > fOutlook;
    static std::shared_ptr< COutlookHelpers > sInstance;
};

#endif