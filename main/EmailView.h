#ifndef EMAILVIEW_H
#define EMAILVIEW_H

#include <QWidget>
#include <memory>
namespace Ui
{
    class CEmailView;
}

class QModelIndex;
class CEmailModel;
class CEmailGroupingModel;

class CEmailView : public QWidget
{
    Q_OBJECT

public:
    explicit CEmailView( QWidget *parent = nullptr );

    void init();

    ~CEmailView();

    void clear();
    void reload();

Q_SIGNALS:
    void sigFinishedLoading();
    void sigFinishedGrouping();

protected slots:
    void slotItemSelected( const QModelIndex &index );
    void slotGroupedItemDoubleClicked( const QModelIndex &idx );
    void slotLoadGrouped();

protected:
    std::shared_ptr< CEmailModel > fModel;
    CEmailGroupingModel *fGroupedModel{ nullptr }; // owned by fModel
    std::unique_ptr< Ui::CEmailView > fImpl;
};

#endif   // CONTACTSVIEW_H
