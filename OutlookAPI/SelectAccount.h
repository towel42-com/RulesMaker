#ifndef SelectAccount_H
#define SelectAccount_H

#include <memory>
#include <optional>

#include <QDialog>
#include <QString>

namespace Outlook
{
    class Account;
}

namespace Ui
{
    class CSelectAccount;
}

class CEmailModel;

class CSelectAccount : public QDialog
{
    Q_OBJECT

public:
    explicit CSelectAccount( QWidget *parent = nullptr );
    virtual void accept() override;

    enum class EInitResult
    {
        eError,
        eNoAccounts,
        eSingleAccount,
        eSuccess
    };

    [[nodiscard]] EInitResult initResult() { return fInitResult; }

    ~CSelectAccount();

    std::pair< QString, std::shared_ptr< Outlook::Account > > account() const;
    bool loadAccountInfo() const;

Q_SIGNALS:

protected Q_SLOTS:

protected:
    void init();
    std::unique_ptr< Ui::CSelectAccount > fImpl;
    EInitResult fInitResult;
    std::pair< QString, std::shared_ptr< Outlook::Account > > fAccount;
    std::optional< std::map< QString, std::shared_ptr< Outlook::Account > > > fAllAccounts;
};

#endif   // CONTACTSVIEW_H
