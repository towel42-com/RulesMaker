#include "OutlookAPI.h"

#include <QInputDialog>
#include <QMessageBox>
#include <QDebug>
#include <QMetaProperty>

#include <cstdlib>
#include <iostream>
#include <oaidl.h>
#include "MSOUTL.h"
#include <objbase.h>

std::shared_ptr< COutlookAPI > COutlookAPI::sInstance;

Q_DECLARE_METATYPE( std::shared_ptr< Outlook::Rule > );

COutlookAPI::COutlookAPI( QWidget *parent, COutlookAPI::SPrivate )
{
    getApplication();
    fParentWidget = parent;

    initSettings();

    qRegisterMetaType< std::shared_ptr< Outlook::Rule > >();
    qRegisterMetaType< std::shared_ptr< Outlook::Rule > >( "std::shared_ptr<Outlook::Rule>const&" );
}

std::shared_ptr< COutlookAPI > COutlookAPI::cliInstance()
{
    if ( !sInstance )
    {
        sInstance = std::make_shared< COutlookAPI >( nullptr, SPrivate() );
    }
    return sInstance;
}

std::shared_ptr< COutlookAPI > COutlookAPI::instance( QWidget *parent )
{
    if ( !sInstance )
    {
        Q_ASSERT( parent );
        sInstance = std::make_shared< COutlookAPI >( parent, SPrivate() );
    }
    else
    {
        Q_ASSERT( !parent );
    }
    return sInstance;
}

COutlookAPI::~COutlookAPI()
{
    logout( false );
}

void COutlookAPI::logout( bool andNotify )
{
    fSession.reset();
    fAccount.reset();
    fInbox.reset();
    fRootFolder.reset();
    fJunkFolder.reset();
    fTrashFolder.reset();
    fContacts.reset();
    fRules.reset();

    if ( fLoggedIn && fOutlookApp && !fOutlookApp->isNull() && fOutlookApp->Session() )
    {
        Outlook::NameSpace( fOutlookApp->Session() ).Logoff();
        fLoggedIn = false;
        if ( andNotify )
            emit sigAccountChanged();
    }
}

std::shared_ptr< Outlook::Application > COutlookAPI::getApplication()
{
    static HRESULT comInit = CoInitialize( nullptr );
    Q_UNUSED( comInit );

    if ( !fOutlookApp )
        fOutlookApp = connectToException( std::make_shared< Outlook::Application >() );
    return fOutlookApp;
}

std::shared_ptr< Outlook::Application > COutlookAPI::outlookApp()
{
    return fOutlookApp;
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getContacts()
{
    return selectContacts();
}

std::shared_ptr< Outlook::Folder > COutlookAPI::selectContacts()
{
    if ( !selectAccount( true ) )
        return {};

    if ( fContacts )
        return fContacts;

    return fContacts = getDefaultFolder( Outlook::OlDefaultFolders::olFolderContacts );
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getInbox()
{
    return selectInbox();
}

std::shared_ptr< Outlook::Folder > COutlookAPI::selectInbox()
{
    if ( !selectAccount( true ) )
        return {};

    if ( fInbox )
        return fInbox;

    return fInbox = getDefaultFolder( Outlook::OlDefaultFolders::olFolderInbox );
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getJunkFolder()
{
    if ( !selectAccount( true ) )
        return {};

    if ( fJunkFolder )
        return fJunkFolder;

    return fJunkFolder = getDefaultFolder( Outlook::OlDefaultFolders::olFolderJunk );
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getTrashFolder()
{
    if ( !selectAccount( true ) )
        return {};

    if ( fTrashFolder )
        return fTrashFolder;

    return fTrashFolder = getDefaultFolder( Outlook::OlDefaultFolders::olFolderDeletedItems );
}

void COutlookAPI::slotHandleException( int code, const QString &source, const QString &desc, const QString &help )
{
    if ( fIgnoreExceptions )
        return;

    if ( fParentWidget )
    {
        auto msg = QString( "%1 - %2: %3" ).arg( source ).arg( code );
        auto txt = "<br>" + desc + "</br>";
        if ( !help.isEmpty() )
            txt += "<br>" + help + "</br>";
        msg = msg.arg( txt );

        QMessageBox::critical( nullptr, "Exception Thrown", msg );
    }
    else
    {
        auto msg = QString( "%1 - %2:\n%3" ).arg( source ).arg( code );
        auto txt = desc + "\n";
        if ( !help.isEmpty() )
            txt += help + "\n";
        msg = msg.arg( txt );
        emit sigStatusMessage( msg );
        std::exit( 1 );
    }
}

Outlook::OlObjectClass COutlookAPI::getObjectClass( IDispatch *item )
{
    if ( !item )
        return {};

    IDispatch *pdisp = (IDispatch *)NULL;
    DISPID dispid;
    OLECHAR *szMember = L"Class";
    auto result = item->GetIDsOfNames( IID_NULL, &szMember, 1, LOCALE_SYSTEM_DEFAULT, &dispid );

    if ( result == S_OK )
    {
        VARIANT resultant{};
        DISPPARAMS params{ 0 };
        EXCEPINFO excepInfo{};
        UINT argErr{ 0 };
        result = item->Invoke( dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD | DISPATCH_PROPERTYGET, &params, &resultant, &excepInfo, &argErr );
        if ( result == S_OK )
        {
            return static_cast< Outlook::OlObjectClass >( resultant.lVal );
        }
    }

    auto retVal = QAxObject( item ).property( "Class" );

    return static_cast< Outlook::OlObjectClass >( retVal.toInt() );
}

std::shared_ptr< Outlook::Items > COutlookAPI::getItems( Outlook::_Items *item )
{
    if ( !item )
        return {};
    return connectToException( std::make_shared< Outlook::Items >( item ) );
}

COutlookAPI::EAddressTypes operator|( const COutlookAPI::EAddressTypes &lhs, const COutlookAPI::EAddressTypes &rhs )
{
    auto lhsA = static_cast< int >( lhs );
    auto rhsA = static_cast< int >( rhs );
    return static_cast< COutlookAPI::EAddressTypes >( lhsA | rhsA );
}

COutlookAPI::EAddressTypes getAddressTypes( bool smtpOnly )
{
    return smtpOnly ? COutlookAPI::EAddressTypes::eSMTPOnly : COutlookAPI::EAddressTypes::eNone;
}

COutlookAPI::EAddressTypes getAddressTypes( std::optional< Outlook::OlMailRecipientType > recipientType, bool smtpOnly )
{
    auto types = getAddressTypes( smtpOnly );
    if ( recipientType.has_value() )
    {
        if ( recipientType == Outlook::OlMailRecipientType::olOriginator )
            types = types | COutlookAPI::EAddressTypes::eOriginator;
        if ( recipientType == Outlook::OlMailRecipientType::olTo )
            types = types | COutlookAPI::EAddressTypes::eTo;
        if ( recipientType == Outlook::OlMailRecipientType::olCC )
            types = types | COutlookAPI::EAddressTypes::eCC;
        if ( recipientType == Outlook::OlMailRecipientType::olBCC )
            types = types | COutlookAPI::EAddressTypes::eBCC;
    }
    else
        types = types | COutlookAPI::EAddressTypes::eAllRecipients;

    return types;
}
