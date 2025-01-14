#ifndef ShowRule_H
#define ShowRule_H

#include <QDialog>
#include <memory>
namespace Ui
{
    class CShowRule;
}

namespace Outlook
{
    class Rule;
}

class CShowRule : public QDialog
{
    Q_OBJECT

public:
    explicit CShowRule( std::shared_ptr< Outlook::Rule > rule, bool readOnly, QWidget *parent = nullptr );
    ~CShowRule();

    virtual void accept() override;

    bool changed() const;
Q_SIGNALS:

protected Q_SLOTS:

protected:
    void init();
    std::shared_ptr< Outlook::Rule > fRule;
    std::unique_ptr< Ui::CShowRule > fImpl;
    bool fReadOnly{ false };
};

#endif   // CONTACTSVIEW_H
