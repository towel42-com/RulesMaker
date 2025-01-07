#ifndef EMAILVIEW_H
#define EMAILVIEW_H

#include "WidgetWithStatus.h"
#include <memory>
#include <QModelIndexList>
namespace Ui
{
    class CEmailView;
}

class CEmailModel;
enum class EFilterType;

class CEmailView : public CWidgetWithStatus
{
    Q_OBJECT

public:
    explicit CEmailView( QWidget *parent = nullptr );
    ~CEmailView();

    void init();

    void clear();
    void clearSelection();
    void reload( bool notifyOnFinished );

    std::pair< QStringList, EFilterType > getPatternsForSelection() const;   // the patterns, by emails or display names

    bool selectionHasDisplayName() const;
    QString getDisplayNameForSingleSelection() const;
    QString getDisplayNamePatternForSelection() const;
    QString getEmailPatternForSelection() const;
    QString getSubjectPatternForSelection() const;

    EFilterType getFilterType() const;

Q_SIGNALS:
    void sigFinishedLoading();
    void sigFinishedGrouping();
    void sigEmailSelected();
    void sigFilterTypeChanged();

public Q_SLOTS:
    void slotRunningStateChanged( bool running );

protected Q_SLOTS:
    void slotSelectionChanged();
    void slotItemDoubleClicked( const QModelIndex &idx );

protected:
    void setFilterType( EFilterType filterType );
    void updateEditFields();
    void updateEditFields( EFilterType filterType );

    QStringList getDisplayNamesForSelection() const;
    QStringList getEmailsForSelection() const;
    QStringList getSubjectsForSelection() const;

    QModelIndexList getSelectedRows() const;

    CEmailModel *fGroupedModel{ nullptr };
    std::unique_ptr< Ui::CEmailView > fImpl;
    bool fNotifyOnFinish{ true };
};

#endif   // CONTACTSVIEW_H
