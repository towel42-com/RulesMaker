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
    setEmailFilterTypes( static_cast< EFilterType >( settings.value( "EmailFilterTypes", 1 ).toInt() ) );
    setRulesToSkip( settings.value( "RulesToSkip", true ).toStringList(), false );

    setRootFolder( settings.value( "RootFolder", R"(\Inbox)" ).toString(), false );
}

void COutlookAPI::setOnlyProcessUnread( bool value, bool update )
{
    update = update && ( fOnlyProcessUnread != value );

    fOnlyProcessUnread = value;
    QSettings settings;
    settings.setValue( "OnlyProcessUnread", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setIncludeJunkFolderWhenRunningOnAllFolders( bool value, bool update )
{
    update = update && ( fIncludeJunkFolderWhenRunningOnAllFolders != value );

    fIncludeJunkFolderWhenRunningOnAllFolders = value;
    QSettings settings;
    settings.setValue( "IncludeJunkFolderWhenRunningOnAllFolders ", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setIncludeDeletedFolderWhenRunningOnAllFolders( bool value, bool update )
{
    update = update || ( fIncludeDeletedFolderWhenRunningOnAllFolders != value );

    fIncludeDeletedFolderWhenRunningOnAllFolders = value;
    QSettings settings;
    settings.setValue( "IncludeDeletedFolderWhenRunningOnAllFolders ", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setProcessAllEmailWhenLessThan200Emails( bool value, bool update )
{
    update = update || ( fProcessAllEmailWhenLessThan200Emails != value );

    fProcessAllEmailWhenLessThan200Emails = value;
    QSettings settings;
    settings.setValue( "ProcessAllEmailWhenLessThan200Emails", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setOnlyProcessTheFirst500Emails( bool value, bool update )
{
    update = update || ( fOnlyProcessTheFirst500Emails != value );

    fOnlyProcessTheFirst500Emails = value;
    QSettings settings;
    settings.setValue( "OnlyProcessTheFirst500Emails", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setDisableRatherThanDeleteRules( bool value, bool update )
{
    update = update || ( fDisableRatherThanDeleteRules != value );

    fDisableRatherThanDeleteRules = value;
    QSettings settings;
    settings.setValue( "DisableRatherThanDeleteRules", value );
    if ( update )
        emit sigOptionChanged();
}

void COutlookAPI::setEmailFilterTypes( const std::list< EFilterType > &value )
{
    int tmp = 0;
    for ( auto &&ii : value )
        tmp |= static_cast< int >( ii );
    setEmailFilterTypes( static_cast< EFilterType >( tmp ) );
}

void COutlookAPI::setEmailFilterTypes( EFilterType value )
{
    fEmailFilterTypes.clear();
    if ( ( static_cast< int >( value ) & static_cast< int >( EFilterType::eByDisplayName ) ) != 0 )
        fEmailFilterTypes.push_back( EFilterType::eByDisplayName );
    if ( ( static_cast< int >( value ) & static_cast< int >( EFilterType::eByEmailAddress ) ) != 0 )
        fEmailFilterTypes.push_back( EFilterType::eByEmailAddress );
    if ( ( static_cast< int >( value ) & static_cast< int >( EFilterType::eBySubject ) ) != 0 )
        fEmailFilterTypes.push_back( EFilterType::eBySubject );
    QSettings settings;
    settings.setValue( "EmailFilterTypes", static_cast< int >( value ) );
}

void COutlookAPI::setRulesToSkip( const QStringList &value, bool update )
{
    update = update && ( value.length() != fRulesToSkip.length() );
    auto count = ( std::min( fRulesToSkip.length(), value.length() ) );
    for ( int ii = 0; update && ( ii < count ); ++ii )
        update = update || ( fRulesToSkip[ ii ] != value[ ii ] );

    fRulesToSkip = value;
    QSettings settings;
    settings.setValue( "RulesToSkip", value );
    if ( update )
        emit sigOptionChanged();
}
