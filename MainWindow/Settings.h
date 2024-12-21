#ifndef Settings_H
#define Settings_H

#include <QDialog>
#include <memory>
namespace Ui
{
    class CSettings;
}

class CEmailModel;

class CSettings : public QDialog
{
    Q_OBJECT

public:
    explicit CSettings( QWidget *parent = nullptr );
    ~CSettings();

    virtual void accept() override;

Q_SIGNALS:

protected Q_SLOTS:

protected:
    void init();
    std::unique_ptr< Ui::CSettings > fImpl;
};

#endif   // CONTACTSVIEW_H
