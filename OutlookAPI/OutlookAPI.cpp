#include "OutlookAPI.h"
#include "ShowRule.h"
#include "DelayDlg.h"

#include <QInputDialog>
#include <QMessageBox>
#include <QDebug>
#include <QMetaProperty>
#include <QTreeView>
#include <QFileInfo>

#include "OutlookLib/MSOUTL.h"

#include <iostream>
#include <cstdlib>
#include <chrono>
#include <thread>

#include <qt_windows.h>
#include <oaidl.h>
#include <objbase.h>
#include <psapi.h>

std::shared_ptr< COutlookAPI > COutlookAPI::sInstance;

Q_DECLARE_METATYPE( std::shared_ptr< Outlook::Rule > );

COutlookAPI::COutlookAPI( QWidget *parent, COutlookAPI::SPrivate )
{
    fParentWidget = parent;
    getApplication();

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

QString COutlookAPI::getDebugName( const std::shared_ptr< Outlook::Rule > &rule )
{
    return getDebugName( rule.get() );
}

QString COutlookAPI::getDebugName( const Outlook::Rule *rule )
{
    if ( !rule )
        return {};
    return QString( "%1%3" ).arg( getDisplayName( rule ) ).arg( rule->Enabled() ? QStringLiteral( "" ) : QStringLiteral( " (Disabled)" ) );
}

QString COutlookAPI::getDebugName( const Outlook::_Rule *rule )
{
    if ( !rule )
        return {};
    return QString( "%1%3" ).arg( getDisplayName( rule ) ).arg( rule->Enabled() ? QStringLiteral( "" ): QStringLiteral( " (Disabled)" ) );
}

QString COutlookAPI::getDisplayName( const std::shared_ptr< Outlook::Rule > &rule )
{
    return getDisplayName( rule.get() );
}

QString COutlookAPI::getDisplayName( const Outlook::Rule *rule )
{
    if ( !rule )
        return {};
    return QString( "%1 (%2)" ).arg( rule->Name() ).arg( rule->ExecutionOrder() );
}

QString COutlookAPI::getDisplayName( const Outlook::_Rule *rule )
{
    if ( !rule )
        return {};
    return QString( "%1 (%2)" ).arg( rule->Name() ).arg( rule->ExecutionOrder() );
}

QString COutlookAPI::getSubject( std::shared_ptr< Outlook::MailItem > mailItem )
{
    return getSubject( mailItem.get() );
}

QString COutlookAPI::getSubject( Outlook::MailItem *mailItem )
{
    return mailItem ? mailItem->Subject() : QString();
}

bool COutlookAPI::outlookProcessRunning()
{
    if ( !fOutlookProcessRunning.has_value() )
    {
        std::size_t baseSize = 256;
        std::size_t incrSize = 256;
        std::vector< DWORD > processIDs;
        processIDs.resize( baseSize );
        bool aOK = false;
        while ( !aOK )
        {
            DWORD size = static_cast< DWORD >( processIDs.capacity() * sizeof( DWORD ) );
            DWORD bytesNeeded{ 0 };
            if ( !EnumProcesses( processIDs.data(), size, &bytesNeeded ) )
            {
                return false;
            }
            aOK = size > bytesNeeded;
            if ( !aOK )
                processIDs.resize( processIDs.size() + incrSize );
        }

        bool found = false;
        for ( auto &&processID : processIDs )
        {
            //    // Get a handle to the process.

            auto hProcess = OpenProcess( PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processID );
            if ( hProcess )
            {
                HMODULE hMod{ 0 };
                DWORD cbNeeded{ 0 };
                if ( EnumProcessModules( hProcess, &hMod, sizeof( hMod ), &cbNeeded ) )
                {
                    TCHAR szProcessName[ MAX_PATH ] = TEXT( "<unknown>" );
                    GetModuleBaseName( hProcess, hMod, szProcessName, sizeof( szProcessName ) / sizeof( TCHAR ) );
                    auto nm = QString::fromWCharArray( szProcessName );
                    if ( QFileInfo( QString::fromWCharArray( szProcessName ) ).fileName().toLower() == "outlook.exe" )
                    {
                        found = true;
                    }
                }
                CloseHandle( hProcess );
            }
            if ( found )
                break;
        }

        fOutlookProcessRunning = found;
    }
    return fOutlookProcessRunning.value();
}

std::shared_ptr< Outlook::Application > COutlookAPI::getApplication()
{
    if ( !fOutlookApp )
    {
        (void)CoInitialize( nullptr );

        auto outlookRunning = outlookProcessRunning();
        fOutlookApp = std::make_shared< Outlook::Application >();
        int numAttempts = 0;
        if ( !outlookRunning )
        {
            if ( getParentWidget() )
            {
                CDelayDlg dlg( [ = ]() { return outlookFullySetup( true ); }, getParentWidget() );
                dlg.exec();
                if ( dlg.cancelled() || dlg.timedOut() )
                {
                    fOutlookApp.reset();
                }
            }
            else
            {
                while ( !outlookRunning && !outlookFullySetup( false ) && ( numAttempts < 5 ) )
                {
                    std::cerr << "WARNING: Outlook was not previously running, waiting up to 5 seconds to allow the system to initialize (Remaining: " << ( 5 - numAttempts ) << ")." << std::endl;

                    using namespace std::chrono_literals;
                    std::this_thread::sleep_for( 1000ms );

                    resetApplication();
                    numAttempts++;
                }
                if ( !outlookFullySetup( false ) )
                {
                    fOutlookApp.reset();
                }
            }
        }
        if ( fOutlookApp )
            fOutlookApp = connectToException( fOutlookApp );
    }
    return fOutlookApp;
}

void COutlookAPI::resetApplication()
{
    fOutlookApp.reset();
    fOutlookApp = std::make_shared< Outlook::Application >();
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

QWidget *COutlookAPI::getParentWidget() const
{
    return fParentWidget;
}

bool COutlookAPI::showRule( std::shared_ptr< Outlook::Rule > rule )
{
    return showRuleDialog( rule, true );
}

bool COutlookAPI::editRule( std::shared_ptr< Outlook::Rule > rule )
{
    return showRuleDialog( rule, false );
}

bool COutlookAPI::showRuleDialog( std::shared_ptr< Outlook::Rule > rule, bool readOnly )
{
    CShowRule ruleDlg( rule, readOnly, getParentWidget() );

    return ruleDlg.exec() == QDialog::Accepted;
}

void COutlookAPI::slotHandleException( int code, const QString &source, const QString &desc, const QString &help )
{
    if ( fIgnoreExceptions )
        return;

    if ( getParentWidget() )
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
    LPOLESTR szMember = const_cast< LPOLESTR >( L"Class" );
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

bool isFilterType( EFilterType value, EFilterType filter )
{
    return ( static_cast< int >( filter ) & static_cast< int >( value ) ) != 0;
}

bool COutlookAPI::isAddressType( EAddressTypes value, EAddressTypes filter )
{
    return ( static_cast< int >( filter ) & static_cast< int >( value ) ) != 0;
}

bool COutlookAPI::isAddressType( std::optional< EAddressTypes > value, std::optional< EAddressTypes > filter )
{
    if ( !value.has_value() || !filter.has_value() )
        return true;
    return isAddressType( value.value(), filter.value() );
}

bool COutlookAPI::isAddressType( Outlook::OlMailRecipientType recipientType, std::optional< EAddressTypes > filter )
{
    if ( !filter.has_value() )
        return true;

    bool retVal = false;
    switch ( recipientType )
    {
        case Outlook::OlMailRecipientType::olOriginator:
            retVal = isAddressType( filter, EAddressTypes::eOriginator );
            break;
        case Outlook::OlMailRecipientType::olTo:
            retVal = isAddressType( filter, EAddressTypes::eTo );
            break;
        case Outlook::OlMailRecipientType::olCC:
            retVal = isAddressType( filter, EAddressTypes::eCC );
            break;
        case Outlook::OlMailRecipientType::olBCC:
            retVal = isAddressType( filter, EAddressTypes::eBCC );
            break;
        default:
            break;
    }
    return retVal;
}

bool COutlookAPI::isContactType( EContactTypes value, EContactTypes filter )
{
    return ( static_cast< int >( filter ) & static_cast< int >( value ) ) != 0;
}

bool COutlookAPI::isContactType( bool isExchangeUser, std::optional< EContactTypes > contactTypes )
{
    if ( !contactTypes.has_value() )
        return true;
    switch ( contactTypes.value() )
    {
        case EContactTypes::eNone:
            return false;
        case EContactTypes::eAllContacts:
            return true;
        case EContactTypes::eSMTPContact:
            return !isExchangeUser;
        case EContactTypes::eOutlookContact:
            return isExchangeUser;
    }
    return false;
}

bool COutlookAPI::isContactType( Outlook::OlAddressEntryUserType contactType, std::optional< EContactTypes > filter )
{
    if ( !filter.has_value() )
        return true;

    bool retVal = false;
    switch ( contactType )
    {
        case Outlook::OlAddressEntryUserType::olExchangeUserAddressEntry:
        case Outlook::OlAddressEntryUserType::olExchangeDistributionListAddressEntry:
        case Outlook::OlAddressEntryUserType::olExchangePublicFolderAddressEntry:
        case Outlook::OlAddressEntryUserType::olExchangeAgentAddressEntry:
        case Outlook::OlAddressEntryUserType::olExchangeOrganizationAddressEntry:
        case Outlook::OlAddressEntryUserType::olExchangeRemoteUserAddressEntry:
        case Outlook::OlAddressEntryUserType::olOutlookContactAddressEntry:
        case Outlook::OlAddressEntryUserType::olOutlookDistributionListAddressEntry:
            retVal = isContactType( true, filter );
            break;
        case Outlook::OlAddressEntryUserType::olLdapAddressEntry:
        case Outlook::OlAddressEntryUserType::olSmtpAddressEntry:
            retVal = isContactType( false, filter );
            break;
        case Outlook::OlAddressEntryUserType::olOtherAddressEntry:
        default:
            break;
    }
    return retVal;
}

bool COutlookAPI::isContactType( std::optional< EContactTypes > value, std::optional< EContactTypes > filter )
{
    if ( !value.has_value() || !filter.has_value() )
        return true;
    return isContactType( value.value(), filter.value() );
}

COutlookAPI::EAddressTypes operator|( const COutlookAPI::EAddressTypes &lhs, const COutlookAPI::EAddressTypes &rhs )
{
    auto lhsA = static_cast< int >( lhs );
    auto rhsA = static_cast< int >( rhs );
    return static_cast< COutlookAPI::EAddressTypes >( lhsA | rhsA );
}

COutlookAPI::EContactTypes operator|( const COutlookAPI::EContactTypes &lhs, const COutlookAPI::EContactTypes &rhs )
{
    auto lhsA = static_cast< int >( lhs );
    auto rhsA = static_cast< int >( rhs );
    return static_cast< COutlookAPI::EContactTypes >( lhsA | rhsA );
}

//COutlookAPI::EAddressTypes getAddressTypes( bool smtpOnly )
//{
//    return smtpOnly ? COutlookAPI::EAddressTypes::eSMTPOnly : COutlookAPI::EAddressTypes::eNone;
//}

//COutlookAPI::EAddressTypes getAddressTypes( std::optional< Outlook::OlMailRecipientType > recipientType, bool smtpOnly )
//{
//    auto types = getAddressTypes( smtpOnly );
//    if ( recipientType.has_value() )
//    {
//        if ( recipientType == Outlook::OlMailRecipientType::olOriginator )
//            types = types | COutlookAPI::EAddressTypes::eOriginator;
//        if ( recipientType == Outlook::OlMailRecipientType::olTo )
//            types = types | COutlookAPI::EAddressTypes::eTo;
//        if ( recipientType == Outlook::OlMailRecipientType::olCC )
//            types = types | COutlookAPI::EAddressTypes::eCC;
//        if ( recipientType == Outlook::OlMailRecipientType::olBCC )
//            types = types | COutlookAPI::EAddressTypes::eBCC;
//    }
//    else
//        types = types | COutlookAPI::EAddressTypes::eAllRecipients;
//
//    return types;
//}

bool equal( const QStringList &lhs, const QStringList &rhs )
{
    auto retVal = lhs.count() == rhs.count();

    auto cnt = lhs.count() < rhs.count() ? lhs.count() : rhs.count();

    for ( auto ii = 0; retVal && ( ii < cnt ); ++ii )
    {
        retVal = retVal && ( lhs[ ii ] == rhs[ ii ] );
    }
    return retVal;
}

void resizeToContentZero( QTreeView *treeView, EExpandMode expandMode )
{
    if ( !treeView )
        return;
    if ( ( expandMode == EExpandMode::eExpandAll ) || ( expandMode == EExpandMode::eExpandAndCollapseAll ) )
        treeView->expandAll();
    treeView->resizeColumnToContents( 0 );
    if ( treeView->columnWidth( 0 ) > 300 )
        treeView->setColumnWidth( 0, 300 );
    if ( ( expandMode == EExpandMode::eCollapseAll ) || ( expandMode == EExpandMode::eExpandAndCollapseAll ) )
        treeView->collapseAll();
}
