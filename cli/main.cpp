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
    eNone = 0x00,
    eRun = 0x01,
    eRename = 0x02,
    eSort = 0x04,
    eFixRules = 0x08,
    eMerge = 0x10,
    eEnableAll = 0x20,
    eEmptyJunk = 0x40,
    eEmptyTrash = 0x80
};

std::pair< EParseResult, EOperation > parseCommandLine( QCommandLineParser &parser, QString &errorMsg )
{
    auto helpOpt = parser.addHelpOption();
    auto versionOpt = parser.addVersionOption();

    EOperation operation = EOperation::eNone;
    if ( !parser.parse( QCoreApplication::arguments() ) )
    {
        errorMsg = parser.errorText();
        return { EParseResult::eError, operation };
    }

    if ( parser.isSet( helpOpt ) )
        return { EParseResult::eHelp, operation };

    if ( parser.isSet( versionOpt ) )
        return { EParseResult::eVersion, operation };

    auto api = COutlookAPI::cliInstance();

    QSettings settings;
    QString profileName;
    if ( parser.isSet( "profile" ) )
    {
        profileName = parser.value( "profile" );
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
    if ( parser.isSet( "account" ) )
    {
        accountName = parser.value( "account" );
        settings.setValue( "Account", accountName );
    }
    else
        accountName = api->defaultAccountName( profileName );

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

    if ( parser.isSet( "run" ) )
    {
        operation = static_cast< EOperation >( operation | EOperation::eRun );
        needsRules = true;
    }
    if ( parser.isSet( "rename" ) )
    {
        operation = static_cast< EOperation >( operation | EOperation::eRename );
        needsRules = true;
    }
    if ( parser.isSet( "sort" ) )
    {
        operation = static_cast< EOperation >( operation | EOperation::eSort );
        needsRules = true;
    }
    if ( parser.isSet( "fix_rules" ) )
    {
        operation = static_cast< EOperation >( operation | EOperation::eFixRules );
        needsRules = true;
    }
    if ( parser.isSet( "merge" ) )
    {
        operation = static_cast< EOperation >( operation | EOperation::eMerge );
        needsRules = true;
    }
    if ( parser.isSet( "enable_all" ) )
    {
        operation = static_cast< EOperation >( operation | EOperation::eEnableAll );
        needsRules = true;
    }
    if ( parser.isSet( "empty_junk" ) )
        operation = static_cast< EOperation >( operation | EOperation::eEmptyJunk );
    if ( parser.isSet( "empty_trash" ) )
        operation = static_cast< EOperation >( operation | EOperation::eEmptyTrash );

    if ( operation == EOperation::eNone )
    {
        operation = EOperation::eRun;
        needsRules = true;
    }

    if ( needsRules )
        api->getRules();

    return { EParseResult::eSuccess, operation };
}

bool runRules( QCommandLineParser &parser )
{
    auto api = COutlookAPI::instance();
    std::shared_ptr< Outlook::Rule > rule;
    if ( parser.isSet( "rule" ) )
    {
        auto ruleName = parser.value( "rule" );
        rule = api->findRule( ruleName );
        if ( !rule )
        {
            std::cerr << "Could not find rule named: " << qPrintable( ruleName ) << std::endl;
            return false;
        }
    }

    std::shared_ptr< Outlook::Folder > folder;
    auto allFolders = parser.isSet( "all_folders" );
    if ( parser.isSet( "folder" ) )
    {
        auto folderName = parser.value( "folder" );
        folder = api->findFolder( folderName, {} );
        if ( !folder )
        {
            std::cerr << "Could not find folder named: " << qPrintable( folderName ) << std::endl;
            return false;
        }
    }

    auto junk = parser.isSet( "junk" );

    bool aOK = false;
    if ( !rule )
        aOK = api->runAllRules( folder, allFolders, junk );
    else
        aOK = api->runRule( rule, folder, allFolders, junk );

    if ( !aOK )
        std::cerr << "Failed to run rule(s)." << std::endl;
    return aOK;
}

void msgHandler( QtMsgType /*type*/, const QMessageLogContext &/*context*/, const QString &/*msg*/ )
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

    QCommandLineParser parser;
    parser.setApplicationDescription( NVersion::APP_NAME + " is a tool to help create and maintain Outlook Rules to keep your inbox clean." );

    QSettings settings;
    auto lastProfile = settings.value( "Profile", QString() ).toString();
    parser.addOption( QCommandLineOption( { "p", "profile" }, "The Outlook <profile> to use. If there is only one profile, it is selected.  If unset, use the default profile if one exists.", "profile", lastProfile ) );
    auto lastAccount = settings.value( "Account", QString() ).toString();
    parser.addOption( QCommandLineOption( { "a", "account" }, "The <account> to use. If there is only 1 account it is automatically selected, and this option is ignored.", "account", lastAccount ) );

    parser.addOption( QCommandLineOption( { "r", "rule" }, "The <rule> to run (if not set all rules are run).", "rule" ) );
    parser.addOption( QCommandLineOption( { "f", "folder" }, "The <folder> to run the rule on (if not set the rule is run on the inbox).", "folder" ) );

    parser.addOption( QCommandLineOption( "all_folders", "Run rule(s) on all folders, overwrites the -f <folder> option." ) );
    parser.addOption( QCommandLineOption( "junk", "Include junk in all_folders." ) );

    parser.addOption( QCommandLineOption( "run", "Run the rule(s) on the folder(s) (if nothing else set, defaults to true)." ) );
    parser.addOption( QCommandLineOption( "rename", "Rename all rules based on their settings." ) );
    parser.addOption( QCommandLineOption( "sort", "Sort rules based on their names." ) );
    parser.addOption( QCommandLineOption( "fix_rules", R"(Move "From" conditional to "Address" in rules settings)" ) );
    parser.addOption( QCommandLineOption( "merge", R"(Merge rules based on destination folder)" ) );
    parser.addOption( QCommandLineOption( "enable_all", R"(Enable all rules)" ) );
    parser.addOption( QCommandLineOption( "empty_junk", R"(Empty Junk Folder)" ) );
    parser.addOption( QCommandLineOption( "empty_trash", R"(Empty Trash Folder)" ) );

    QString msg;
    auto &&[ result, operation ] = parseCommandLine( parser, msg );

    switch ( result )
    {
        case EParseResult::eSuccess:
            break;
        case EParseResult::eError:
            std::cerr << qPrintable( msg ) << std::endl;
            parser.showHelp( 1 );
            return 1;
        case EParseResult::eVersion:
            parser.showVersion();
            return 0;
        case EParseResult::eHelp:
            parser.showHelp();
            return 0;
    }

    bool aOK = true;
    bool needsSaving = false;
    if ( aOK && ( operation & EOperation::eRename ) != 0 )
    {
        aOK = aOK && api->renameRules( false, &needsSaving );
        if ( !aOK )
            std::cerr << "Failed to rename rule(s)." << std::endl;
    }

    if ( aOK && ( operation & EOperation::eSort ) != 0 )
    {
        aOK = aOK && api->sortRules( false, &needsSaving );
        if ( !aOK )
            std::cerr << "Failed to sort rule(s)." << std::endl;
    }

    if ( aOK && ( operation & EOperation::eFixRules ) != 0 )
    {
        aOK = aOK && api->moveFromToAddress( false, &needsSaving );
        if ( !aOK )
            std::cerr << "Failed to fix rule(s)." << std::endl;
    }

    if ( aOK && ( operation & EOperation::eMerge ) != 0 )
    {
        aOK = aOK && api->mergeRules( false, &needsSaving );
        if ( !aOK )
            std::cerr << "Failed to merge rule(s)." << std::endl;
    }

    if ( aOK && ( operation & EOperation::eEnableAll ) != 0 )
    {
        aOK = aOK && api->enableAllRules( false, &needsSaving );
        if ( !aOK )
            std::cerr << "Failed to enable all rules rule(s)." << std::endl;
    }

    if ( needsSaving )
        api->saveRules();

    if ( aOK && ( operation & EOperation::eRun ) != 0 )
    {
        aOK = runRules( parser );
    }

    if ( aOK && ( operation & EOperation::eEmptyJunk ) != 0 )
    {
        aOK = aOK && api->emptyJunk();
        if ( !aOK )
            std::cerr << "Failed to empty Junk." << std::endl;
    }

    if ( aOK && ( operation & EOperation::eEmptyTrash ) != 0 )
    {
        aOK = aOK && api->emptyTrash();
        if ( !aOK )
            std::cerr << "Failed to empty Trash." << std::endl;
    }
    return aOK ? 0 : 1;
}
