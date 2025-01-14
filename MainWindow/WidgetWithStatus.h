#ifndef OBJECTWITHSTATUS_H
#define OBJECTWITHSTATUS_H

#include <QString>
#include <QWidget>

class CWidgetWithStatus : public QWidget
{
    Q_OBJECT
public:
    CWidgetWithStatus( QWidget *parent = nullptr ) :
        QWidget( parent ) {};
    ~CWidgetWithStatus() = default;

    QString statusLabel() const { return fStatusLabel; }
    void setStatusLabel( const QString &label ) { fStatusLabel = label; }
Q_SIGNALS:
    void sigStatusMessage( const QString &msg );
    void sigInitStatus( const QString &label, int max );
    void sigSetStatus( const QString &label, int curr, int max );
    void sigIncStatusValue( const QString &label );
    void sigStatusFinished( const QString &label );

protected:
    QString fStatusLabel;
};

#endif
