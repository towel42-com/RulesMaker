#ifndef LISTFILTERMODEL_H
#define LISTFILTERMODEL_H

#include <QString>
#include <QSortFilterProxyModel>

#include <memory>

class CListFilterModel : public QSortFilterProxyModel
{
    Q_OBJECT;

public:
    explicit CListFilterModel( QObject *parent );
    virtual ~CListFilterModel();

    void setOnlyFilterParent( bool value ) { fOnlyFilterParent = value; }
    bool onlyFilterParent()const { return fOnlyFilterParent ; }

    virtual bool filterAcceptsColumn( int source_column, const QModelIndex &source_parent ) const;
    virtual bool filterAcceptsRow( int source_row, const QModelIndex &source_parent ) const;
Q_SIGNALS:

public Q_SLOTS:
    void slotSetFilter( const QString &filter );

private:
    bool fOnlyFilterParent{ false };
};

#endif
