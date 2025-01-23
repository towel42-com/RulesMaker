#include "OutlookAPI.h"
#include "ShowRule.h"
#include "ExceptionHandler.h"

#include <QInputDialog>
#include <QMessageBox>
#include <QDebug>
#include <QMetaProperty>
#include <QTreeView>

#include <cstdlib>
#include <iostream>
#include "MSOUTL.h"
#include <oaidl.h>
#include <objbase.h>

std::shared_ptr< COutlookAPI > COutlookAPI::sInstance;

Q_DECLARE_METATYPE( COutlookObj< Outlook::Rule > );

COutlookAPI::COutlookAPI( QWidget *parent, COutlookAPI::SPrivate )
{
    getApplication();
    fParentWidget = parent;

    initSettings();

    qRegisterMetaType< COutlookObj< Outlook::Rule > >();
    qRegisterMetaType< COutlookObj< Outlook::Rule > >( "COutlookObj< Outlook::Rule >const&" );
}

std::shared_ptr< COutlookAPI > COutlookAPI::cliInstance()
{
    if ( !sInstance )
    {
        CExceptionHandler::cliInstance();
        sInstance = std::make_shared< COutlookAPI >( nullptr, SPrivate() );
    }
    return sInstance;
}

std::shared_ptr< COutlookAPI > COutlookAPI::instance( QWidget *parent )
{
    if ( !sInstance )
    {
        CExceptionHandler::instance( parent );
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

QString COutlookAPI::getDebugName( const COutlookObj< Outlook::Rule > &rule )
{
    return getDebugName( rule.get() );
}

QString COutlookAPI::getDebugName( const Outlook::Rule *rule )
{
    if ( !rule )
        return {};
    return QString( "%1%3" ).arg( getDisplayName( rule ) ).arg( rule->Enabled() ? "" : " (Disabled)" );
}

QString COutlookAPI::getDisplayName( const COutlookObj< Outlook::Rule > &rule )
{
    return getDisplayName( rule.get() );
}

QString COutlookAPI::getDisplayName( const Outlook::Rule *rule )
{
    if ( !rule )
        return {};
    return QString( "%1 (%2)" ).arg( rule->Name() ).arg( rule->ExecutionOrder() );
}

QString COutlookAPI::getSubject( const COutlookObj< Outlook::MailItem > &mailItem )
{
    return getSubject( mailItem.get() );
}

QString COutlookAPI::getSubject( Outlook::MailItem *mailItem )
{
    return mailItem ? mailItem->Subject() : QString();
}

COutlookObj< Outlook::Application > COutlookAPI::getApplication()
{
    if ( !fOutlookApp )
        fOutlookApp = COutlookObj< Outlook::Application >();
    return fOutlookApp;
}

COutlookObj< Outlook::Application > COutlookAPI::outlookApp()
{
    return fOutlookApp;
}

COutlookObj< Outlook::MAPIFolder > COutlookAPI::getContacts()
{
    return selectContacts();
}

COutlookObj< Outlook::MAPIFolder > COutlookAPI::selectContacts()
{
    if ( !selectAccount( true ) )
        return {};

    if ( fContacts )
        return fContacts;

    return fContacts = getDefaultFolder( Outlook::OlDefaultFolders::olFolderContacts );
}

COutlookObj< Outlook::MAPIFolder > COutlookAPI::getInbox()
{
    return selectInbox();
}

COutlookObj< Outlook::MAPIFolder > COutlookAPI::selectInbox()
{
    if ( !selectAccount( true ) )
        return {};

    if ( fInbox )
        return fInbox;

    return fInbox = getDefaultFolder( Outlook::OlDefaultFolders::olFolderInbox );
}

COutlookObj< Outlook::MAPIFolder > COutlookAPI::getJunkFolder()
{
    if ( !selectAccount( true ) )
        return {};

    if ( fJunkFolder )
        return fJunkFolder;

    return fJunkFolder = getDefaultFolder( Outlook::OlDefaultFolders::olFolderJunk );
}

COutlookObj< Outlook::MAPIFolder > COutlookAPI::getTrashFolder()
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

bool COutlookAPI::showRule( const COutlookObj< Outlook::Rule > &rule )
{
    return showRuleDialog( rule, true );
}

bool COutlookAPI::editRule( const COutlookObj< Outlook::Rule > &rule )
{
    return showRuleDialog( rule, false );
}

bool COutlookAPI::showRuleDialog( const COutlookObj< Outlook::Rule > &rule, bool readOnly )
{
    CShowRule ruleDlg( rule, readOnly, fParentWidget );

    return ruleDlg.exec() == QDialog::Accepted;
}

COutlookObj< Outlook::_Items > COutlookAPI::getItems( Outlook::_Items *items )
{
    if ( !items )
        return {};
    return COutlookObj< Outlook::_Items >( items );
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
