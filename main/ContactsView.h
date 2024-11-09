#ifndef CONTACTSVIEW_H
#define CONTACTSVIEW_H

#include <QWidget>
#include <memory>
namespace Ui
{
    class CContactsView;
}

class QModelIndex;
class CContactsModel;

class CContactsView : public QWidget
{
    Q_OBJECT

public:
    explicit CContactsView( QWidget *parent = nullptr );
    ~CContactsView();

protected slots:
    void addEntry();
    void changeEntry();
    void itemSelected( const QModelIndex &index );

    void updateOutlook();

protected:
    std::shared_ptr< CContactsModel > fModel;
    std::unique_ptr< Ui::CContactsView > fImpl;
};

#endif   // CONTACTSVIEW_H
