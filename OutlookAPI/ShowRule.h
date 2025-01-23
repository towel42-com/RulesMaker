#ifndef ShowRule_H
#define ShowRule_H

#include "OutlookObj.h"

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
    explicit CShowRule( const COutlookObj< Outlook::Rule > &rule, bool readOnly, QWidget *parent = nullptr );
    ~CShowRule();

    virtual void accept() override;

    bool changed() const;
Q_SIGNALS:

protected Q_SLOTS:

protected:
    void init();
    COutlookObj< Outlook::Rule > fRule;
    std::unique_ptr< Ui::CShowRule > fImpl;
    bool fReadOnly{ false };
};

#endif   // CONTACTSVIEW_H
