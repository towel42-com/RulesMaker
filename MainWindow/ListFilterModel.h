#ifndef LISTFILTERMODEL_H
#define LISTFILTERMODEL_H

#include <QString>
#include <QSortFilterProxyModel>
#include <functional>
#include <memory>

class CListFilterModel : public QSortFilterProxyModel
{
    Q_OBJECT;

public:
    explicit CListFilterModel( QObject *parent );
    virtual ~CListFilterModel();

    void setOnlyFilterParent( bool value ) { fOnlyFilterParent = value; }
    bool onlyFilterParent() const { return fOnlyFilterParent; }

    void setLessThanOp( const std::function< bool( const QModelIndex &source_left, const QModelIndex &source_right ) > &lessThanOp ) { fLessThanOp = lessThanOp; }

    virtual bool filterAcceptsColumn( int source_column, const QModelIndex &source_parent ) const override;
    virtual bool filterAcceptsRow( int source_row, const QModelIndex &source_parent ) const override;
    virtual bool lessThan( const QModelIndex &source_left, const QModelIndex &source_right ) const override;

Q_SIGNALS:

public Q_SLOTS:
    void slotSetFilter( const QString &filter );

private:
    bool fOnlyFilterParent{ false };
    std::function< bool( const QModelIndex &source_left, const QModelIndex &source_right ) > fLessThanOp;
};

#endif
