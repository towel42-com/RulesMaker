#ifndef RULESVIEW_H
#define RULESVIEW_H

#include <QWidget>
#include <memory>
namespace Ui
{
    class CRulesView;
}

class QModelIndex;
class CRulesModel;

class CRulesView : public QWidget
{
    Q_OBJECT

public:
    explicit CRulesView( QWidget *parent = nullptr );
    ~CRulesView();

protected slots:
    void addEntry();
    void changeEntry();
    void itemSelected( const QModelIndex &index );

    void updateOutlook();

protected:
    std::shared_ptr< CRulesModel > fModel;
    std::unique_ptr< Ui::CRulesView > fImpl;
};

#endif   // CRulesView_H
