#include "ShowRule.h"
#include "OutlookAPI/OutlookAPI.h"
#include "OutlookLib/MSOUTL.h"
#include "ui_ShowRule.h"

#include <QInputDialog>
#include <QSpinBox>

CShowRule::CShowRule( std::shared_ptr< Outlook::Rule > rule, bool readOnly, QWidget *parent ) :
    QDialog( parent ),
    fImpl( new Ui::CShowRule ),
    fRule( rule ),
    fReadOnly( readOnly )
{
    init();
}

QString getHtml( QStringList list, const QString &separator )
{
    list.removeAll( QString() );
    if ( list.isEmpty() )
        return {};

    list.sort();

    if ( list.count() == 1 )
        return list.first();

    QString retVal;
    retVal += "<table>";
    bool first = true;
    for ( auto &str : list )
    {
        retVal += QString( "<tr><td style=\"white-space:nowrap\">%1</td><td style=\"white-space:nowrap\">%2</td>" ).arg( first ? "&nbsp;" : separator, str );
        first = false;
    }
    retVal += "</table>";
    return retVal;
}

QString getHtml( const std::list< QStringList > &list )
{
    int count = 0;
    for ( auto &&ii : list )
    {
        if ( !ii.isEmpty() )
        {
            count++;
        }
    }
    if ( count == 0 )
        return {};

    if ( count == 1 )
    {
        for ( auto &&ii : list )
        {
            if ( !ii.isEmpty() )
            {
                return getHtml( ii, "or" );
            }
        }
    }

    QString retVal;

    retVal += "<table>";
    bool first = true;
    for ( auto &curr : list )
    {
        if ( curr.isEmpty() )
            continue;
        retVal += QString( "<tr><td style=\"white-space:nowrap\">%1</td><td style=\"white-space:nowrap\">%2</td>" ).arg( first ? "&nbsp;" : "and", getHtml( curr, "or" ) );
        first = false;
    }
    retVal += "</table>";
    return retVal;
}

void CShowRule::init()
{
    fImpl->setupUi( this );

    if ( fReadOnly )
    {
        connect( fImpl->enabled, &QGroupBox::toggled, this, [ this ]() { fImpl->enabled->setChecked( fRule->Enabled() ); } );
        connect( fImpl->localRule, &QCheckBox::toggled, this, [ this ]() { fImpl->localRule->setChecked( fRule->IsLocalRule() ); } );
        connect( fImpl->executionOrder, qOverload< int >( &QSpinBox::valueChanged ), this, [ this ]() { fImpl->executionOrder->setValue( fRule->ExecutionOrder() ); } );
        fImpl->name->setReadOnly( true );
        fImpl->autoRename->setEnabled( false );
    }
    connect(
        fImpl->autoRename, &QToolButton::clicked, this,
        [ this ]()
        {
            auto newName = COutlookAPI::ruleNameForRule( fRule, false );
            fImpl->name->setText( newName );
        } );
    fImpl->enabled->setChecked( fRule->Enabled() );
    fImpl->name->setText( fRule->Name() );
    fImpl->executionOrder->setValue( fRule->ExecutionOrder() );
    fImpl->localRule->setChecked( fRule->IsLocalRule() );
    fImpl->when->setHtml( getHtml( COutlookAPI::getConditionalStringList( fRule, false, EWrapperMode::eNone, true ) ) );
    fImpl->except->setHtml( getHtml( COutlookAPI::getConditionalStringList( fRule, true, EWrapperMode::eNone, true ) ) );
    fImpl->actions->setHtml( getHtml( COutlookAPI::getActionStrings( fRule ), "&#x21AA;" ) );

    fImpl->executionOrder->setValue( fRule->ExecutionOrder() );
    fImpl->localRule->setChecked( fRule->IsLocalRule() );
}

CShowRule::~CShowRule()
{
}

void CShowRule::accept()
{
    if ( !fReadOnly )
    {
        //if ( !COutlookAPI::instance()->saveRules() )
        // return;
    }
    QDialog::accept();
}

bool CShowRule::changed() const
{
    if ( fReadOnly )
        return false;

    return false;
}
