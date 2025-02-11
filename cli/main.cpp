#include <QCoreApplication>
#include <QCommandLineParser>
#include <QSettings>

#include "OutlookAPI/OutlookAPI.h"

#include "Version.h"
#include <iostream>

enum class EParseResult
{
    eSuccess,
    eError,
    eHelp,
    eVersion,
};

enum EOperation
{
    eNone = 0x0000,
    eRun = 0x0001,
    eRename = 0x0002,
    eSort = 0x0004,
    eMerge = 0x0010,
    eEnableAll = 0x0020,
    eEmptyJunk = 0x0040,
    eEmptyTrash = 0x0080,
    eRunOnJunk = 0x0100
};

EOperation operator|( const EOperation &lhs, const EOperation &rhs )
{
    return static_cast< EOperation >( static_cast< int >( lhs ) | static_cast< int >( rhs ) );
}

bool isOperation( EOperation value, EOperation filter )
{
    return ( static_cast< int >( filter ) & static_cast< int >( value ) ) != 0;
}

QCommandLineParser sParser;

std::pair< EParseResult, EOperation > parseCommandLine( QString &errorMsg )
{
    auto helpOpt = sParser.addHelpOption();
    auto versionOpt = sParser.addVersionOption();

    EOperation operation = EOperation::eNone;
    if ( !sParser.parse( QCoreApplication::arguments() ) )
    {
        errorMsg = sParser.errorText();
        return { EParseResult::eError, operation };
    }

    if ( sParser.isSet( helpOpt ) )
        return { EParseResult::eHelp, operation };

    if ( sParser.isSet( versionOpt ) )
        return { EParseResult::eVersion, operation };

    auto api = COutlookAPI::cliInstance();

    QSettings settings;
    QString profileName;
    if ( sParser.isSet( "profile" ) )
    {
        profileName = sParser.value( "profile" );
        settings.setValue( "Profile", profileName );
    }
    else
        profileName = api->defaultProfileName();

    if ( profileName.isEmpty() )
    {
        errorMsg = "-profile unset and no default profile found.";
        return { EParseResult::eError, operation };
    }

    QString accountName;
    if ( sParser.isSet( "account" ) )
    {
        accountName = sParser.value( "account" );
        settings.setValue( "Account", accountName );
    }
    else
        accountName = api->defaultAccountName();

    if ( accountName.isEmpty() )
    {
        errorMsg = QString( R"(-account unset and no default account found for profile "%1".)" ).arg( profileName );
        return { EParseResult::eError, operation };
    }

    if ( !api->selectAccount( accountName, true ) )
    {
        errorMsg = QString( R"(Failed to select account "%1".)" ).arg( accountName );
        return { EParseResult::eError, operation };
    }

    bool needsRules = false;

    if ( sParser.isSet( "run" ) )
    {
        operation = operation | EOperation::eRun;
        needsRules = true;
    }
    if ( sParser.isSet( "run_on_junk" ) )
    {
        operation = operation | EOperation::eRunOnJunk;
        needsRules = true;
    }

    if ( sParser.isSet( "rename" ) )
    {
        operation = operation | EOperation::eRename;
        needsRules = true;
    }
    if ( sParser.isSet( "sort" ) )
    {
        operation = operation | EOperation::eSort;
        needsRules = true;
    }
    if ( sParser.isSet( "merge" ) )
    {
        operation = operation | EOperation::eMerge;
        needsRules = true;
    }
    if ( sParser.isSet( "enable_all" ) )
    {
        operation = operation | EOperation::eEnableAll;
        needsRules = true;
    }
    if ( sParser.isSet( "empty_junk" ) )
        operation = operation | EOperation::eEmptyJunk;
    if ( sParser.isSet( "empty_trash" ) )
        operation = operation | EOperation::eEmptyTrash;

    if ( operation == EOperation::eNone )
    {
        operation = EOperation::eRun;
        needsRules = true;
    }

    if ( needsRules )
        api->getRules();

    return { EParseResult::eSuccess, operation };
}

bool runRules( bool onJunkIfFolderNotSet )
{
    auto api = COutlookAPI::instance();
    std::shared_ptr< Outlook::Rule > rule;
    if ( sParser.isSet( "rule" ) )
    {
        auto ruleName = sParser.value( "rule" );
        rule = api->findRule( ruleName );
        if ( !rule )
        {
            std::cerr << "Could not find rule named: " << qPrintable( ruleName ) << std::endl;
            return false;
        }
    }

    std::shared_ptr< Outlook::Folder > folder;
    if ( sParser.isSet( "folder" ) )
    {
        auto folderName = sParser.value( "folder" );
        folder = api->findFolder( folderName, {} );
        if ( !folder )
        {
            std::cerr << "Could not find folder named: " << qPrintable( folderName ) << std::endl;
            return false;
        }
    }

    bool aOK = false;
    if ( !rule )
    {
        if ( onJunkIfFolderNotSet )
            aOK = api->runAllRulesOnJunkFolder();
        else
            aOK = api->runAllRules( folder );
    }
    else
        aOK = api->runRule( rule, folder );

    if ( !aOK )
        std::cerr << "Failed to run rule(s)." << std::endl;
    return aOK;
}

void msgHandler( QtMsgType /*type*/, const QMessageLogContext & /*context*/, const QString & /*msg*/ )
{
    //Logger::instance()->handleMessage( type, msg );
}

int main( int argc, char *argv[] )
{
    QCoreApplication appl( argc, argv );
    NVersion::setupApplication( appl, true );
    qInstallMessageHandler( msgHandler );

    std::map< QString, std::pair< int, int > > statusCounter;

    auto api = COutlookAPI::cliInstance();
    QObject::connect(
        api.get(), &COutlookAPI::sigInitStatus,
        [ & ]( const QString &label, int max )
        {
            statusCounter[ label ] = std::make_pair( 0, max );
            std::cout << qPrintable( label ) << std::endl;
        } );
    QObject::connect(
        api.get(), &COutlookAPI::sigSetStatus,
        [ & ]( const QString &label, int curr, int max )
        {
            statusCounter[ label ] = { curr, max };
            std::cout << qPrintable( label ) << " " << curr << " of " << max << std::endl;
        } );
    QObject::connect(
        api.get(), &COutlookAPI::sigIncStatusValue,
        [ & ]( const QString &label )
        {
            auto pos = statusCounter.find( label );
            if ( pos == statusCounter.end() )
                pos = statusCounter.insert( { label, { 0, 0 } } ).first;
            else
                ( *pos ).second.first = ( *pos ).second.first + 1;
            std::cout << qPrintable( label ) << " " << ( *pos ).second.first << " of " << ( *pos ).second.second << std::endl;
        } );
    QObject::connect( api.get(), &COutlookAPI::sigStatusMessage, [ = ]( const QString &msg ) { std::cout << qPrintable( msg ) << std::endl; } );
    QObject::connect( api.get(), &COutlookAPI::sigStatusFinished, [ = ]( const QString &label ) { std::cout << "Finished - " << qPrintable( label ) << std::endl; } );

    sParser.setApplicationDescription( NVersion::APP_NAME + " is a tool to help create and maintain Outlook Rules to keep your inbox clean." );

    QSettings settings;
    auto lastProfile = settings.value( "Profile", QString() ).toString();
    sParser.addOption( QCommandLineOption( { "p", "profile" }, "The Outlook <profile> to use. If there is only one profile, it is selected.  If unset, use the default profile if one exists.", "profile", lastProfile ) );
    auto lastAccount = settings.value( "Account", QString() ).toString();
    sParser.addOption( QCommandLineOption( { "a", "account" }, "The <account> to use. If there is only 1 account it is automatically selected, and this option is ignored.", "account", lastAccount ) );

    sParser.addOption( QCommandLineOption( "rule", "The <rule> to run (if not set all rules are run).", "rule" ) );
    sParser.addOption( QCommandLineOption( { "f", "folder" }, "The <folder> to run the rule on (if not set the rule is run on the inbox).", "folder" ) );

    sParser.addOption( QCommandLineOption( "run", "Run the rule(s) on the folder(s) (if nothing else set, defaults to true)." ) );
    sParser.addOption( QCommandLineOption( "run_on_junk", "Run all the rule(s) on the junk folder." ) );

    sParser.addOption( QCommandLineOption( "rename", "Rename all rules based on their settings." ) );
    sParser.addOption( QCommandLineOption( "sort", "Sort rules based on their names." ) );
    sParser.addOption( QCommandLineOption( "merge", R"(Merge rules based on destination folder and conditions)" ) );
    sParser.addOption( QCommandLineOption( "enable_all", R"(Enable all rules)" ) );
    sParser.addOption( QCommandLineOption( "empty_junk", R"(Empty Junk Folder)" ) );
    sParser.addOption( QCommandLineOption( "empty_trash", R"(Empty Trash Folder)" ) );

    QString msg;
    auto &&[ result, operation ] = parseCommandLine( msg );

    switch ( result )
    {
        case EParseResult::eSuccess:
            break;
        case EParseResult::eError:
            std::cerr << qPrintable( msg ) << std::endl;
            sParser.showHelp( 1 );
            return 1;
        case EParseResult::eVersion:
            sParser.showVersion();
            return 0;
        case EParseResult::eHelp:
            sParser.showHelp();
            return 0;
    }

    bool aOK = true;
    bool needsSaving = false;
    bool currNeedsSaving = false;
    if ( aOK && isOperation( operation, EOperation::eRename ) )
    {
        aOK = aOK && api->renameRules( false, &currNeedsSaving );
        if ( !aOK )
            std::cerr << "Failed to rename rule(s)." << std::endl;
        else
            needsSaving = needsSaving || currNeedsSaving;
    }

    if ( aOK && isOperation( operation, EOperation::eSort ) )
    {
        aOK = aOK && api->sortRules( false, &currNeedsSaving );
        if ( !aOK )
            std::cerr << "Failed to sort rule(s)." << std::endl;
        else
            needsSaving = needsSaving || currNeedsSaving;
    }

    if ( aOK && isOperation( operation, EOperation::eMerge ) )
    {
        aOK = aOK && api->mergeRules( false, &currNeedsSaving );
        if ( !aOK )
            std::cerr << "Failed to merge rule(s)." << std::endl;
        else
            needsSaving = needsSaving || currNeedsSaving;
    }

    if ( aOK && isOperation( operation, EOperation::eEnableAll ) )
    {
        aOK = aOK && api->enableAllRules( false, &currNeedsSaving );
        if ( !aOK )
            std::cerr << "Failed to enable all rules rule(s)." << std::endl;
        else
            needsSaving = needsSaving || currNeedsSaving;
    }

    if ( needsSaving )
        api->saveRules();

    if ( aOK && isOperation( operation, EOperation::eRun ) )
    {
        aOK = runRules( false );
    }

    if ( aOK && isOperation( operation, EOperation::eRunOnJunk ) )
    {
        aOK = runRules( true );
    }

    if ( aOK && isOperation( operation, EOperation::eEmptyJunk ) )
    {
        aOK = api->emptyJunk();
        if ( !aOK )
            std::cerr << "Failed to empty Junk." << std::endl;
    }

    if ( aOK && isOperation( operation, EOperation::eEmptyTrash ) )
    {
        aOK = api->emptyTrash();
        if ( !aOK )
            std::cerr << "Failed to empty Trash." << std::endl;
    }
    return aOK ? 0 : 1;
}
