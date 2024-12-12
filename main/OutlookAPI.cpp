#include "OutlookAPI.h"

#include <QInputDialog>
#include <QMessageBox>
#include <QDebug>
#include <QMetaProperty>
#include <QSettings>

#include <oaidl.h>
#include "MSOUTL.h"

std::shared_ptr< COutlookAPI > COutlookAPI::sInstance;
Q_DECLARE_METATYPE( std::shared_ptr< Outlook::Rule > );

COutlookAPI::COutlookAPI( QWidget *parent, COutlookAPI::SPrivate )
{
    getApplication();
    fParentWidget = parent;

    QSettings settings;
    setOnlyProcessUnread( settings.value( "OnlyProcessUnread", true ).toBool(), false );
    setProcessAllEmailWhenLessThan200Emails( settings.value( "ProcessAllEmailWhenLessThan200Emails", true ).toBool(), false );
    setRootFolder( settings.value( "RootFolder", R"(\Inbox)" ).toString(), false );

    qRegisterMetaType< std::shared_ptr< Outlook::Rule > >();
    qRegisterMetaType< std::shared_ptr< Outlook::Rule > >( "std::shared_ptr<Outlook::Rule>const&" );
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
    fAccount.reset();
    fInbox.reset();
    fRootFolder.reset();
    fJunkFolder.reset();
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
    if ( !fContacts )
        fContacts = selectContacts( false ).first;
    return fContacts;
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookAPI::selectContacts( bool singleOnly )
{
    if ( !fAccount )
    {
        if ( !selectAccount( true ) )
            return {};
    }

    return selectFolder(
        "Contacts",
        []( const std::shared_ptr< Outlook::Folder > &folder )
        {
            if ( !folder )
                return false;

            if ( folder->DefaultItemType() != Outlook::OlItemType::olContactItem )
                return false;

            auto items = folder->Items();
            if ( items->Count() == 0 )
                return false;

            if ( folder->Name().contains( "meta", Qt::CaseSensitivity::CaseInsensitive ) )
                return false;

            return true;
        },
        {}, singleOnly );
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getInbox()
{
    if ( !fInbox )
        fInbox = selectInbox( false ).first;
    return fInbox;
}

std::pair< std::shared_ptr< Outlook::Folder >, bool > COutlookAPI::selectInbox( bool singleOnly )
{
    if ( !fAccount )
    {
        if ( !selectAccount( true ) )
            return {};
    }

    return getMailFolder( "Inbox", R"(Inbox)", singleOnly );
}

std::shared_ptr< Outlook::Folder > COutlookAPI::getJunkFolder()
{
    if ( !fJunkFolder )
    {
        fJunkFolder = getMailFolder( "Junk Email", R"(Junk Email)", false ).first;
    }
    return fJunkFolder;
}


void COutlookAPI::slotHandleException( int code, const QString &source, const QString &desc, const QString &help )
{
    auto msg = QString( "%1 - %2: %3" ).arg( source ).arg( code );
    auto txt = "<br>" + desc + "</br>";
    if ( !help.isEmpty() )
        txt += "<br>" + help + "</br>";
    msg = msg.arg( txt );

    QMessageBox::critical( nullptr, "Exception Thrown", msg );
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

void COutlookAPI::setOnlyProcessUnread( bool value, bool update )
{
    fOnlyProcessUnread = value;
    QSettings settings;
    settings.setValue( "OnlyProcessUnread", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setProcessAllEmailWhenLessThan200Emails( bool value, bool update )
{
    fProcessAllEmailWhenLessThan200Emails = value;
    QSettings settings;
    settings.setValue( "ProcessAllEmailWhenLessThan200Emails", value );
    if ( update )
        emit sigOptionChanged();
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
