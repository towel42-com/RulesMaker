#ifndef EMAILMODEL_H
#define EMAILMODEL_H

#include <QString>
#include <QAbstractListModel>
#include <optional>
#include <memory>
#include <tuple>
namespace Outlook
{
    class Items;
    class MailItem;
}

class QStandardItemModel;
class CEmailGroupingModel;

class CEmailModel : public QAbstractListModel
{
    Q_OBJECT;

public:
    explicit CEmailModel( QObject *parent );
    virtual ~CEmailModel();

    int rowCount( const QModelIndex &parent = QModelIndex() ) const;
    int columnCount( const QModelIndex &parent ) const;
    QVariant headerData( int section, Qt::Orientation orientation, int role ) const;
    QVariant data( const QModelIndex &index, int role ) const;

    std::tuple< CEmailGroupingModel *, CEmailGroupingModel *, CEmailGroupingModel *, QStandardItemModel * > CEmailModel::getGroupedEmailModels( QWidget *parent );

Q_SIGNALS:
    void sigFinishedGroupingEmails();

private:
    void addMailItems( QWidget *parent );
    void addMailItem( Outlook::MailItem *mailItem );

    std::shared_ptr< Outlook::Items > fItems{ nullptr };
    mutable QHash< int, QStringList > fCache;
    mutable std::optional< int > fCountCache;

    CEmailGroupingModel *fGroupedFrom{ nullptr };
    CEmailGroupingModel *fGroupedTo{ nullptr };
    CEmailGroupingModel *fGroupedCC{ nullptr };
    QStandardItemModel *fUniqueSubjects{ nullptr };
    QHash< QString, bool > fSubjectMap;
};

#endif
