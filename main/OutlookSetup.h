#ifndef OUTLOOKSETUP_H
#define OUTLOOKSETUP_H

#include <QDialog>
#include <memory>
namespace Ui
{
    class COutlookSetup;
}

class QModelIndex;
class CFoldersModel;

class COutlookSetup : public QDialog
{
    Q_OBJECT

public:
    explicit COutlookSetup( QWidget *parent = nullptr );
    ~COutlookSetup();

    virtual void reject() override;
protected Q_SLOTS:
    void slotSelectAccount( bool useInbox = true );
    void slotSelectFolder( bool useInbox = true );

protected:
    std::unique_ptr< Ui::COutlookSetup > fImpl;
};

#endif   // OUTLOOKSETUP_H
