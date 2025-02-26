#ifndef DelayDlg_H
#define DelayDlg_H

#include <memory>

#include <QDateTime>
#include <QDialog>
#include <functional>

class QTimer;

namespace Outlook
{
    class Account;
}

namespace Ui
{
    class CDelayDlg;
}

class CDelayDlg : public QDialog
{
    Q_OBJECT

public:
    explicit CDelayDlg( const std::function< bool() > &okToEnd, QWidget *parent = nullptr );
    ~CDelayDlg();

    virtual int exec();
    void setTimeout( int64_t ms ) { fTimeOut = ms; }

    virtual void reject() override
    {
        fCancelled = true;
        QDialog::reject();
    }
    bool timedOut() { return fTimedOut; }
    bool cancelled() { return fCancelled; }
Q_SIGNALS:

protected Q_SLOTS:
    void slotUpdateTime();

protected:
    QTimer *fTimer{ nullptr };
    QDateTime fStartTime;
    int64_t fTimeOut{ 5000 };
    bool fTimedOut{ false };
    bool fCancelled{ false };
    std::function< bool() > fOKToEnd;
    std::unique_ptr< Ui::CDelayDlg > fImpl;
};

#endif   // CONTACTSVIEW_H
