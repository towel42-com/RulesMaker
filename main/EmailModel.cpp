#include "EmailModel.h"
#include "EmailGroupingModel.h"
#include "OutlookHelpers.h"

#include <QProgressDialog>
#include <QTimer>
#include <QProgressBar>

#include "MSOUTL.h"

CEmailModel::CEmailModel( QObject *parent ) :
    QAbstractListModel( parent )
{
    auto folder = COutlookHelpers::getInstance()->getInbox( dynamic_cast< QWidget * >( parent ) );
    if ( !folder )
        return;

    fItems = std::make_shared< Outlook::Items >( folder->Items() );
    if ( fItems )
        fCountCache = fItems->Count();
    connect( fItems.get(), SIGNAL( ItemAdd( IDispatch * ) ), parent, SLOT( updateOutlook() ) );
    connect( fItems.get(), SIGNAL( ItemChange( IDispatch * ) ), parent, SLOT( updateOutlook() ) );
    connect( fItems.get(), SIGNAL( ItemRemove() ), parent, SLOT( updateOutlook() ) );
}

CEmailModel::~CEmailModel()
{
}

int CEmailModel::rowCount( const QModelIndex & ) const
{
    if ( fItems && fCountCache.has_value() )
        return fCountCache.value();
    return fItems ? fItems->Count() : 0;
}

int CEmailModel::columnCount( const QModelIndex & /*parent*/ ) const
{
    return 4;
}

QVariant CEmailModel::headerData( int section, Qt::Orientation /*orientation*/, int role ) const
{
    if ( role != Qt::DisplayRole )
        return QVariant();

    switch ( section )
    {
        case 0:
            return tr( "From" );
        case 1:
            return tr( "To" );
        case 2:
            return tr( "CC" );
        case 3:
            return tr( "Subject" );
        default:
            break;
    }

    return QVariant();
}

QVariant CEmailModel::data( const QModelIndex &index, int role ) const
{
    if ( !index.isValid() || role != Qt::DisplayRole )
        return QVariant();

    QStringList data;
    if ( fCache.contains( index.row() ) )
    {
        data = fCache.value( index.row() );
    }
    else if ( fItems )
    {
        Outlook::MailItem mail( fItems->Item( index.row() + 1 ) );
        if ( COutlookHelpers::getObjectClass( &mail ) == Outlook::OlObjectClass::olMail )
        {
            auto from = COutlookHelpers::getSenderEmailAddress( &mail );
            auto to = COutlookHelpers::getRecipients( &mail, Outlook::OlMailRecipientType::olTo );
            auto cc = COutlookHelpers::getRecipients( &mail, Outlook::OlMailRecipientType::olCC );

            data << from << to.join( ";" ) << cc.join( ";" ) << mail.Subject();
            fCache.insert( index.row(), data );
        }
    }

    if ( index.column() < data.count() )
        return data.at( index.column() );

    return QVariant();
}

std::tuple< CEmailGroupingModel *, CEmailGroupingModel *, CEmailGroupingModel *, QStandardItemModel * > CEmailModel::getGroupedEmailModels( QWidget *parent )
{
    delete fGroupedFrom;
    delete fGroupedTo;
    delete fGroupedCC;
    delete fUniqueSubjects;
    fSubjectMap.clear();

    fGroupedFrom = new CEmailGroupingModel( this );
    fGroupedTo = new CEmailGroupingModel( this );
    fGroupedCC = new CEmailGroupingModel( this );
    fUniqueSubjects = new QStandardItemModel( this );
    fUniqueSubjects->setHorizontalHeaderLabels( { "Subject" } );

    auto retVal = std::make_tuple( fGroupedFrom, fGroupedTo, fGroupedCC, fUniqueSubjects );

    QTimer::singleShot( 10, [ = ]() { addMailItems( parent ); } );

    return retVal;
}

void CEmailModel::addMailItems( QWidget *parent )
{
    if ( !fItems )
        return;

    auto itemCount = fItems->Count();
    QProgressDialog dlg( parent );
    auto bar = new QProgressBar;
    bar->setFormat( "(%v of %m - %p%)" );
    dlg.setBar( bar );
    dlg.setMinimum( 0 );
    dlg.setMaximum( itemCount );
    dlg.setLabelText( "Grouping Emails" );
    dlg.setWindowModality( Qt::WindowModal );

    for ( int ii = 1; ii <= itemCount; ++ii )
    {
        dlg.setValue( ii );
        if ( dlg.wasCanceled() )
        {
            fGroupedFrom->clear();
            fGroupedTo->clear();
            fGroupedCC->clear();
            fSubjectMap.clear();
            fUniqueSubjects->clear();
            break;
        }

        auto item = fItems->Item( ii );
        if ( !item )
            continue;

        Outlook::MailItem mail( item );
        if ( COutlookHelpers::getObjectClass( &mail ) == Outlook::OlObjectClass::olMail )
            addMailItem( &mail );
    }
    emit sigFinishedGroupingEmails();
}

void CEmailModel::addMailItem( Outlook::MailItem *mailItem )
{
    if ( !mailItem )
        return;

    fGroupedFrom->addEmailAddress( COutlookHelpers::getSenderEmailAddress( mailItem ) );

    auto emailList = COutlookHelpers::getRecipients( mailItem, Outlook::OlMailRecipientType::olTo );
    for ( auto &&ii : emailList )
        fGroupedTo->addEmailAddress( ii );

    emailList = COutlookHelpers::getRecipients( mailItem, Outlook::OlMailRecipientType::olCC );
    for ( auto &&ii : emailList )
        fGroupedCC->addEmailAddress( ii );

    auto subject = mailItem->Subject();
    auto pos = fSubjectMap.find( subject );
    if ( pos == fSubjectMap.end() )
    {
        fSubjectMap[ subject ] = true;
        fUniqueSubjects->appendRow( new QStandardItem( subject ) );
    }
}
