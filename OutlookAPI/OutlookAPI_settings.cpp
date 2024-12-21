#include "OutlookAPI.h"
#include <QSettings>

void COutlookAPI::initSettings()
{
    QSettings settings;

    setOnlyProcessUnread( settings.value( "OnlyProcessUnread", true ).toBool(), false );
    setProcessAllEmailWhenLessThan200Emails( settings.value( "ProcessAllEmailWhenLessThan200Emails", true ).toBool(), false );
    setOnlyProcessTheFirst500Emails( settings.value( "OnlyProcessTheFirst500Emails", true ).toBool(), false );
    setIncludeJunkFolderWhenRunningOnAllFolders( settings.value( "IncludeJunkFolderWhenRunningOnAllFolders", false ).toBool(), false );
    setIncludeDeletedFolderWhenRunningOnAllFolders( settings.value( "IncludeJunkDeletedWhenRunningOnAllFolders", false ).toBool(), false );
    setDisableRatherThanDeleteRules( settings.value( "DisableRatherThanDeleteRules", true ).toBool(), false );
    setRulesToSkip( settings.value( "RulesToSkip", true ).toStringList(), false );

    setRootFolder( settings.value( "RootFolder", R"(\Inbox)" ).toString(), false );
}

void COutlookAPI::setOnlyProcessUnread( bool value, bool update )
{
    fOnlyProcessUnread = value;
    QSettings settings;
    settings.setValue( "OnlyProcessUnread", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setIncludeJunkFolderWhenRunningOnAllFolders( bool value, bool update )
{
    fIncludeJunkFolderWhenRunningOnAllFolders = value;
    QSettings settings;
    settings.setValue( "IncludeJunkFolderWhenRunningOnAllFolders ", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setIncludeDeletedFolderWhenRunningOnAllFolders( bool value, bool update )
{
    fIncludeDeletedFolderWhenRunningOnAllFolders = value;
    QSettings settings;
    settings.setValue( "IncludeDeletedFolderWhenRunningOnAllFolders ", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setProcessAllEmailWhenLessThan200Emails( bool value, bool update )
{
    fProcessAllEmailWhenLessThan200Emails = value;
    QSettings settings;
    settings.setValue( "ProcessAllEmailWhenLessThan200Emails", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setOnlyProcessTheFirst500Emails( bool value, bool update )
{
    fOnlyProcessTheFirst500Emails = value;
    QSettings settings;
    settings.setValue( "OnlyProcessTheFirst500Emails", value );
    if ( update )
        emit sigOptionChanged();
}


void COutlookAPI::setDisableRatherThanDeleteRules( bool value, bool update )
{
    fDisableRatherThanDeleteRules = value;
    QSettings settings;
    settings.setValue( "DisableRatherThanDeleteRules", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setRulesToSkip( const QStringList & value, bool update )
{
    fRulesToSkip = value;
    QSettings settings;
    settings.setValue( "RulesToSkip", value );
    if ( update )
        emit sigOptionChanged();
}
