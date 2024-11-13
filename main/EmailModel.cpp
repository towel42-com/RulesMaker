#include "EmailModel.h"
#include "EmailGroupingModel.h"
#include "OutlookHelpers.h"

#include <QTimer>
#include <QProgressDialog>
#include <QProgressBar>
#include <QDateTime>

#include <QDebug>
#include "MSOUTL.h"

//QTimer::singleShot( 10, [ = ]() { addMailItems( parent ); } );

//return retVal;

CEmailModel::CEmailModel( QObject *parent ) :
    QAbstractListModel( parent )
{
}

void CEmailModel::reload()
{
    beginResetModel();
    clear();
    auto folder = COutlookHelpers::getInstance()->getInbox( dynamic_cast< QWidget * >( parent() ) );
    if ( !folder )
        return;

    fItems = std::make_shared< Outlook::Items >( folder->Items() );
    if ( fItems )
        fCountCache = fItems->Count();

    endResetModel();
    emit sigFinishedLoading();
    QTimer::singleShot( 0, [ = ]() { groupMailItemsBySender( dynamic_cast< QWidget * >( parent() ) ); } );
}

void CEmailModel::clear()
{
    beginResetModel();
    fGroupedFrom->clear();
    fCache.clear();
    fItems.reset();
    fCountCache.reset();
    endResetModel();
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
            auto to = COutlookHelpers::getRecipientEmails( &mail, Outlook::OlMailRecipientType::olTo );
            auto cc = COutlookHelpers::getRecipientEmails( &mail, Outlook::OlMailRecipientType::olCC );

            data << from << to.join( ";" ) << cc.join( ";" ) << mail.Subject();
            fCache.insert( index.row(), data );
        }
    }

    if ( index.column() < data.count() )
        return data.at( index.column() );

    return QVariant();
}

CEmailGroupingModel *CEmailModel::getGroupedEmailModel()
{
    if ( !fGroupedFrom )
        fGroupedFrom = new CEmailGroupingModel( this );
    return fGroupedFrom;
}

void CEmailModel::groupMailItemsBySender( QWidget *parent )
{
    if ( !fItems )
        return;

    auto start = QDateTime::currentDateTime();

    auto itemCount = fItems->Count();
    QProgressDialog dlg( parent );
    auto bar = new QProgressBar;
    bar->setFormat( "(%v of %m - %p%)" );
    dlg.setBar( bar );
    dlg.setMinimum( 0 );
    dlg.setMaximum( itemCount );
    dlg.setLabelText( "Grouping Emails" );
    dlg.setMinimumDuration( 0 );
    dlg.setWindowModality( Qt::WindowModal );

    for ( int ii = 1; ii <= itemCount; ++ii )
    {
        dlg.setValue( ii );
        if ( dlg.wasCanceled() )
        {
            fGroupedFrom->clear();
            break;
        }

        auto item = fItems->Item( ii );
        if ( !item )
            continue;

        auto mail = std::make_shared< Outlook::MailItem >( item );
        if ( COutlookHelpers::getObjectClass( mail.get() ) == Outlook::OlObjectClass::olMail )
            fGroupedFrom->addEmailAddress( mail, COutlookHelpers::getSenderEmailAddress( mail.get() ) );
#ifdef _DEBUG
        if ( ii >= 1000 )
            break;
#endif
    }
    emit sigFinishedGrouping();
    auto end = QDateTime::currentDateTime();
    auto diff = start.secsTo( end );
    qDebug() << "It took " << diff << " seconds to group emails.";
}
