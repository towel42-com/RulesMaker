#ifndef EMAILVIEW_H
#define EMAILVIEW_H

#include "WidgetWithStatus.h"
#include <memory>
#include <QModelIndexList>
namespace Ui
{
    class CFilterFromEmailView;
}

class CFilterFromEmailModel;
enum class EFilterType;

class CFilterFromEmailView : public CWidgetWithStatus
{
    Q_OBJECT

public:
    explicit CFilterFromEmailView( QWidget *parent = nullptr );
    ~CFilterFromEmailView();

    void init();

    void initFilterTypes();

    void clear();
    void clearSelection();
    void reload( bool notifyOnFinished );

    std::list< std::pair< QStringList, EFilterType > > getPatternsForSelection() const;   // the patterns, by emails or display names

    bool selectionHasSender() const;
    bool selectionHasDisplayName() const;
    bool selectionHasPattern() const;

    QString getDisplayNameForSingleSelection() const;
    QString getDisplayNamePatternForSelection() const;
    QString getEmailPatternForSelection() const;
    QString getSubjectPatternForSelection() const;
    QString getSenderPatternForSelection() const;

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
    void updateEditFields();

    QStringList getDisplayNamesForSelection() const;
    QStringList getEmailsForSelection() const;
    QStringList getSubjectsForSelection() const;
    QStringList getSendersForSelection() const;

    QModelIndexList getSelectedRows() const;

    CFilterFromEmailModel *fGroupedModel{ nullptr };
    std::unique_ptr< Ui::CFilterFromEmailView > fImpl;
    bool fNotifyOnFinish{ true };
};

#endif   // CONTACTSVIEW_H
