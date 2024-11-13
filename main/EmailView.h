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

class CEmailView : public QWidget
{
    Q_OBJECT

public:
    explicit CEmailView( QWidget *parent = nullptr );
    ~CEmailView();

    void reload();

Q_SIGNALS:
    void sigFinishedLoading();
    void sigFinishedGrouping();

protected slots:
    void itemSelected( const QModelIndex &index );
    void slotLoadGrouped();

protected:
    std::shared_ptr< CEmailModel > fModel;
    std::unique_ptr< Ui::CEmailView > fImpl;
};

#endif   // CONTACTSVIEW_H
