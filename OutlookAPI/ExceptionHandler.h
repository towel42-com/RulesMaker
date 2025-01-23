#ifndef EXCEPTION_HANDLER
#define EXCEPTION_HANDLER

#include <QObject>
#include <QWidget>
#include <memory>

class QString;

class CExceptionHandler : public QObject
{
    Q_OBJECT;

    struct SPrivate
    {
        explicit SPrivate() = default;
    };

public:
    CExceptionHandler() = delete;
    CExceptionHandler( QWidget *parent, SPrivate );
    static std::shared_ptr< CExceptionHandler > instance( QWidget *parentWidget = nullptr );
    static std::shared_ptr< CExceptionHandler > cliInstance();

    void connectToException( QObject * obj );
    void setIgnoreExceptions( bool value ) { fIgnoreExceptions = value; }
Q_SIGNALS:
    void sigStatusMessage( const QString &msg );

public Q_SLOTS:
    void slotHandleException( int code, const QString &source, const QString &desc, const QString &help );

private:
    QWidget *fParentWidget{ nullptr };
    bool fIgnoreExceptions{ false };

    static std::shared_ptr< CExceptionHandler > sInstance;
};
#endif