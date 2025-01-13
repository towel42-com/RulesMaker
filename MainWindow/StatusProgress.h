#ifndef STATUSPROGRESS_H
#define STATUSPROGRESS_H

#include <QWidget>
class QLabel;
class QProgressBar;

namespace Ui
{
    class CStatusProgress;
}

class QModelIndex;
class CFoldersModel;

class CStatusProgress : public QWidget
{
    Q_OBJECT

public:
    explicit CStatusProgress( QWidget *parent = nullptr ) :
        CStatusProgress( {}, parent )
    {
    }
    explicit CStatusProgress( const QString &label, QWidget *parent = nullptr );
    ~CStatusProgress();

    void setRange( int min, int max );
    void finished();
public Q_SLOTS:
    void slotSetStatus( int curr, int max );
    void slotIncValue();

Q_SIGNALS:
    void sigShow();
    void sigFinished();

protected:
    QLabel *fLabel{ nullptr };
    QProgressBar *fProgressBar{ nullptr };
};

#endif   // STATUSPROGRESS_H
