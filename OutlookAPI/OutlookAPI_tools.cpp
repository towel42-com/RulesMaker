#include "OutlookAPI.h"
#include "EmailAddress.h"
#include "SelectFolders.h"

#include <QMessageBox>

#include <map>
#include <QRegularExpression>
#include <QDebug>
#include <QTextStream>
#include <QMetaMethod>

#include "MSOUTL.h"
#include <tuple>

bool COutlookAPI::deleteAllDisabledRules( bool andSave /*= true*/, bool *needsSaving /*= nullptr*/ )
{
    if ( needsSaving )
        *needsSaving = false;

    if ( !fRules )
        return false;

    slotClearCanceled();

    auto numRules = fRules->Count();
    emit sigInitStatus( "Deleting Disabled Rules:", numRules );

    std::list< Outlook::_Rule * > rules;
    int numChanged = 0;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;
        auto rule = getRule( fRules->Item( ii ) );
        emit sigIncStatusValue( "Deleting Disabled Rules:" );
        if ( !rule )
            continue;
        if ( rule->Enabled() )
            continue;
        emit sigStatusMessage( QString( "Deleting rule '%1'" ).arg( getDisplayName( rule ) ) );
        deleteRule( rule, false, false );
        numChanged++;
        numRules--;
        --ii;
    }
    if ( canceled() )
        return false;

    if ( fParentWidget )
        QMessageBox::information( fParentWidget, R"(Delete Disabled Rules)", QString( "%1 rules deleted" ).arg( numChanged ) );
    else
        emit sigStatusMessage( QString( "%1 rules deleted" ).arg( numChanged ) );

    if ( needsSaving )
        *needsSaving = numChanged != 0;

    if ( andSave && ( numChanged != 0 ) )
        saveRules();
    return true;
}

bool COutlookAPI::enableAllRules( bool andSave /*= true*/, bool *needsSaving /*= nullptr*/ )
{
    if ( needsSaving )
        *needsSaving = false;

    if ( !fRules )
        return false;

    slotClearCanceled();

    auto numRules = fRules->Count();
    emit sigInitStatus( "Enabling Rules:", numRules );

    std::list< Outlook::_Rule * > rules;
    int numChanged = 0;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;
        auto rule = fRules->Item( ii );
        emit sigIncStatusValue( "Enabling Rules:" );
        if ( !rule )
            continue;
        if ( rule->Enabled() )
            continue;
        emit sigStatusMessage( QString( "Enabling rule '%1'" ).arg( getDisplayName( rule ) ) );
        rule->SetEnabled( true );
        numChanged++;
    }
    if ( canceled() )
        return false;

    if ( fParentWidget )
        QMessageBox::information( fParentWidget, R"(Enable All Rules)", QString( "%1 rules enabled" ).arg( numChanged ) );
    else
        emit sigStatusMessage( QString( "%1 rules enabled" ).arg( numChanged ) );

    if ( needsSaving )
        *needsSaving = numChanged != 0;

    if ( andSave && ( numChanged != 0 ) )
        saveRules();
    return true;
}

std::optional< QStringList > COutlookAPI::mergeRecipients( Outlook::Rule *lhs, const QStringList &rhs, QStringList *msgs )
{
    auto lhsRecipients = getRecipients( lhs, msgs );
    if ( !lhsRecipients.has_value() )
        return rhs;
    lhsRecipients.value() = mergeStringLists( lhsRecipients.value(), rhs, false );
    return lhsRecipients;
}

std::optional< QStringList > COutlookAPI::mergeRecipients( Outlook::Rule *lhs, const TEmailAddressList &rhs, QStringList *msgs )
{
    return mergeRecipients( lhs, getAddresses( rhs ), msgs );
}

std::optional< QStringList > COutlookAPI::mergeRecipients( Outlook::Rule *lhs, Outlook::Rule *rhs, QStringList *msgs )
{
    auto tmpRhsRecipients = getRecipients( rhs, msgs );
    QStringList rhsRecipients;
    if ( tmpRhsRecipients.has_value() )
        rhsRecipients = tmpRhsRecipients.value();

    return mergeRecipients( lhs, rhsRecipients, msgs );
}

QString COutlookAPI::stripHeaderStringString( const QString &msg )
{
    auto regEx = QRegularExpression( R"((From:)|")", QRegularExpression::CaseInsensitiveOption );
    auto retVal = msg;
    retVal = retVal.remove( regEx ).trimmed();
    return retVal;
}

QStringList COutlookAPI::getFromMessageHeaderString( const QString &address )
{
    auto stripped = stripHeaderStringString( address );
    if ( stripped.isEmpty() )
        return {};

    auto retVal = QStringList()   //
                  << QString( R"(From: %1)" ).arg( stripped )   //
                  << QString( R"(From: "%1")" ).arg( stripped );
    return retVal;
}

QStringList COutlookAPI::getFromMessageHeaderStrings( const QStringList &addresses )
{
    QStringList retVal;
    for ( auto &&address : addresses )
    {
        auto cleandedAddresses = getFromMessageHeaderString( address );
        retVal << cleandedAddresses;
    }
    retVal.removeDuplicates();
    retVal.removeAll( QString() );
    retVal.sort( Qt::CaseInsensitive );

    return retVal;
}

QString getULForList( const QStringList &list )
{
    QString retVal;
    for ( auto &&ii : list )
        retVal += QString( "<li style=\"white-space:nowrap\">%1</li>\n" ).arg( ii.toHtmlEscaped() );
    retVal = QString( "<ul>\n%1</ul>" ).arg( retVal );
    return retVal;
};

struct SMessage
{
    SMessage() = default;
    SMessage( const QString &msg, const QStringList &params ) :
        fMessage( msg ),
        fParams( params )
    {
    }

    QString toString( bool toHtml, bool bold = false )
    {
        auto retVal = parameterize( toHtml );
        if ( toHtml && bold )
            retVal = "<b>" + retVal + "</b>";
        return retVal;
    }

private:
    QString parameterize( bool toHtml )
    {
        for ( auto &&param : fParams )
        {
            if ( toHtml )
                param = param.toHtmlEscaped();
        }
        for ( auto &&param : fParams )
            fMessage = fMessage.arg( param );
        fParams.clear();
        return fMessage;
    }

    QString fMessage;
    QStringList fParams;
};

struct SDisplayMessage
{
    SDisplayMessage(){};

    QString toString( bool toHtml, int level = 0 )
    {
        QString retVal;
        QTextStream ts( &retVal );

        bool isListItem = level != 0;
        if ( isListItem )
        {
            ts << indent( level );
            if ( toHtml )
                ts << "<li>\n";
            level++;
        }

        for ( auto &&ii = fTitle.begin(); ii != fTitle.end(); ++ii )
        {
            ts << indent( level ) << ( *ii ).toString( toHtml, true );
            if ( toHtml && ( std::next( ii ) == fTitle.end() ) )
                ts << R"(<br>)";
            ts << "\n";
        }

        auto isList = isListItem || ( fChildren.size() + fMessages.size() ) > 1;
        if ( toHtml )
        {
            if ( isList )
                ts << indent( level ) << "<ul>\n";
        }
        for ( auto &&msg : fMessages )
        {
            if ( toHtml && isList )
                ts << indent( level + 1 ) << QString( "<li style=\"white-space:nowrap\">%1</li>\n" ).arg( msg.toString( toHtml ) );
            else
                ts << indent( level + 1 ) << msg.toString( toHtml ) << "\n";
        }
        for ( auto &&child : fChildren )
        {
            ts << child.toString( toHtml, level + 1 );
        }
        if ( toHtml && isList )
            ts << indent( level ) << "</ul>\n";

        for ( auto &&ii = fFooter.begin(); ii != fFooter.end(); ++ii )
        {
            ts << indent( level ) << ( *ii ).toString( toHtml, true );
            if ( toHtml && ( std::next( ii ) != fFooter.end() ) )
                ts << R"(<br>)";
            ts << "\n";
        }

        if ( isListItem )
        {
            level--;
            ts << indent( level );
            if ( toHtml )
                ts << "</li>\n";
        }

        if ( toHtml && ( level == 0 ) )
            retVal = retVal.remove( "\n" );
        return retVal;
    }

    QString indent( int level ) { return QString( level * 4, ' ' ); }

    void addTitle( const QString &msg, const QStringList &parameters ) { fTitle.emplace_back( msg, parameters ); }
    void addTitle( const QString &msg ) { addTitle( msg, {} ); }

    void addMessage( const QString &msg, const QStringList &parameters ) { fMessages.emplace_back( msg, parameters ); }
    void addMessage( const QString &msg ) { addMessage( msg, {} ); }

    void addFooter( const QString &msg, const QStringList &parameters ) { fFooter.emplace_back( msg, parameters ); }
    void addFooter( const QString &msg ) { addFooter( msg, {} ); }

    void addChild( const SDisplayMessage &child ) { fChildren.push_back( child ); }

private:
    std::list< SMessage > fTitle;
    std::list< SMessage > fMessages;
    std::list< SMessage > fFooter;

    std::list< SDisplayMessage > fChildren;
};

bool COutlookAPI::mergeRules( bool andSave /*= true*/, bool *needsSaving /*= nullptr*/ )
{
    if ( needsSaving )
        *needsSaving = false;

    if ( !fRules )
        return false;

    slotClearCanceled();

    auto rules = findMergableRules();
    if ( canceled() )
        return false;

    bool forMsgBox = fParentWidget != nullptr;
    SDisplayMessage msg;
    msg.addTitle( QString( "%1 merge(s) found" ), { QString::number( rules.size() ) } );
    if ( !rules.empty() )
        msg.addTitle( QString( "Merging the following rules:" ) );

    for ( auto &&[ key, matches ] : rules )
    {
        auto primaryRule = matches.first;
        auto &&matchedRules = matches.second;

        SDisplayMessage ruleMessage;
        ruleMessage.addTitle( QString( "Primary Rule: %1" ), { getDisplayName( primaryRule ) } );
        for ( auto &&currRule : matchedRules )
        {
            ruleMessage.addMessage( QString( "Rule: %1" ), { getDisplayName( currRule ) } );
        }
        msg.addChild( ruleMessage );
    }

    if ( !forMsgBox )
        emit sigStatusMessage( msg.toString( false ) );
    else
    {
        if ( !rules.empty() )
        {
            msg.addFooter( "Do you wish to continue?" );

            auto process = QMessageBox::information( fParentWidget, R"(Merge Rules by Target Folder)", msg.toString( true ), QMessageBox::Yes | QMessageBox::No );
            if ( process == QMessageBox::No )
                return false;
        }
        else
        {
            QMessageBox::information( fParentWidget, R"(Merge Rules by Target Folder)", msg.toString( true ) );
            return false;
        }
    }

    if ( canceled() )
        return false;

    emit sigInitStatus( "Merging Rules:", static_cast< int >( rules.size() ) );

    for ( auto &&ii : rules )
    {
        if ( canceled() )
            return false;

        mergeRules( ii.second );

        emit sigIncStatusValue( "Merging Rules:" );
    }

    if ( needsSaving )
        *needsSaving = !rules.empty();

    if ( andSave && !rules.empty() )
        saveRules();

    return true;
}

bool COutlookAPI::fixFromMessageHeaderRules( bool andSave /*= true*/, bool *needsSaving /*= nullptr*/ )
{
    if ( needsSaving )
        *needsSaving = false;

    if ( !fRules )
        return false;

    slotClearCanceled();

    auto numRules = fRules->Count();
    emit sigInitStatus( "Fixing From Message Header Rules:", numRules );
    std::list< std::pair< std::shared_ptr< Outlook::Rule >, std::pair< QStringList, QStringList > > > changes;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;

        auto rule = getRule( fRules->Item( ii ) );
        if ( !rule )
            continue;

        emit sigIncStatusValue( "Fixing From Message Header Rules:" );

        auto conditions = rule->Conditions();
        if ( !conditions )
            continue;

        auto header = rule->Conditions()->MessageHeader();
        if ( !header || !header->Enabled() )
            continue;

        emit sigStatusMessage( QString( "Checking Message Header on rule '%1'" ).arg( getDisplayName( rule ) ) );

        auto headerText = toStringList( header->Text() );
        auto newHeaderText = getFromMessageHeaderStrings( headerText );
        bool currChanged = !equal( headerText, newHeaderText );

        if ( currChanged )
        {
            changes.emplace_back( rule, std::make_pair( headerText, newHeaderText ) );
        }
    }

    if ( changes.empty() )
    {
        if ( fParentWidget )
            QMessageBox::information( fParentWidget, "Fixing From Message Header Rules", QString( "No rules needed fixing" ) );
        else
            emit sigStatusMessage( QString( "No rules needed fixing" ) );
        return false;
    }

    if ( fParentWidget )
    {
        QStringList tmp;
        for ( auto &&ii : changes )
        {
            auto curr = QString( "<li style=\"white-space:nowrap\">Rule: %1 will change in the following manner:<br>" ).arg( getDisplayName( ii.first ).toHtmlEscaped() );

            curr += "\nFrom:" + getULForList( ii.second.first );
            curr += "\nTo:" + getULForList( ii.second.second );
            curr += "\n</li>";

            tmp << curr;
        }
        auto msg = QString( "Rules to be changed:<ul>\n%1</ul>\nDo you wish to continue?" ).arg( tmp.join( "\n" ) );
        auto process = QMessageBox::information( fParentWidget, "Renamed Rules", msg, QMessageBox::Yes | QMessageBox::No );
        if ( process == QMessageBox::No )
            return false;
    }
    else
    {
        QStringList tmp;
        for ( auto &&ii : changes )
        {
            emit sigStatusMessage( QString( "Rule: %1 will change in the following manner:\n" ).arg( ii.first->Name() ) );
            emit sigStatusMessage( "    From: \n" + ii.second.first.join( "    \n" ) );
            emit sigStatusMessage( "    To: \n" + ii.second.second.join( "    \n" ) );
        }
    }

    for ( auto &&change : changes )
    {
        auto rule = change.first;
        if ( !rule )
            continue;

        auto conditions = rule->Conditions();
        if ( !conditions )
            continue;

        auto header = rule->Conditions()->MessageHeader();
        if ( !header || !header->Enabled() )
            continue;
        header->SetText( change.second.second );
        header->SetEnabled( true );
    }

    if ( needsSaving )
        *needsSaving = true;
    if ( andSave )
        saveRules();

    return true;
}

template< typename T >
bool folderIsEmpty( const T *folder )
{
    if ( !folder )
        return false;

    auto hasFolders = folder->Folders() && ( folder->Folders()->Count() > 0 );
    auto hasItems = folder->Items() && ( folder->Items()->Count() > 0 );
    bool isEmpty = !hasFolders && !hasItems;
    if ( isEmpty )
        return true;

    if ( !hasItems && hasFolders )
    {
        for ( int ii = 1; ii < folder->Folders()->Count(); ++ii )
        {
            auto subFolder = folder->Folders()->Item( ii );
            if ( !folderIsEmpty( subFolder ) )
                return false;
        }
        return true;
    }
    return false;
}

bool COutlookAPI::folderIsEmpty( const Outlook::Folder *folder )
{
    return ::folderIsEmpty( folder );
}

bool COutlookAPI::folderIsEmpty( const Outlook::MAPIFolder *folder )
{
    return ::folderIsEmpty( folder );
}

bool COutlookAPI::folderIsEmpty( const std::shared_ptr< Outlook::Folder > &folder )
{
    return ::folderIsEmpty( folder.get() );
}

bool COutlookAPI::folderIsEmpty( const std::shared_ptr< Outlook::MAPIFolder > &folder )
{
    return ::folderIsEmpty( folder.get() );
}

void COutlookAPI::deleteFolderAndParentsIfEmpty( const std::shared_ptr< Outlook::Folder > &folder )
{
    if ( !folder )
        return;
    if ( !folderIsEmpty( folder ) )
        return;

    auto currFolder = folder;
    while ( currFolder && folderIsEmpty( currFolder ) )
    {
        auto mapiFolder = reinterpret_cast< Outlook::MAPIFolder * >( currFolder.get() );
        auto parentFolder = this->parentFolder( currFolder );
        mapiFolder->Delete();
        currFolder = parentFolder;
    }
}

bool COutlookAPI::findEmptyFolders()
{
    slotClearCanceled();
    auto allFolders = getFolders(
        getInbox(), true, true,
        [ = ]( const std::shared_ptr< Outlook::Folder > &folder )
        {
            return folderIsEmpty( folder );   //
        } );
    auto numFolders = allFolders.size();

    allFolders.sort(
        []( const std::shared_ptr< const Outlook::Folder > &lhs, const std::shared_ptr< Outlook::Folder > &rhs )
        {
            if ( !lhs )
                return false;
            if ( !rhs )
                return true;
            auto lhsName = lhs->FolderPath();
            auto rhsName = rhs->FolderPath();
            if ( lhsName.startsWith( rhsName ) && ( lhsName != rhsName ) )
                return true;
            else if ( rhsName.startsWith( lhsName ) && ( lhsName != rhsName ) )
                return false;
            else
                return lhsName < rhsName;
        } );

    if ( allFolders.empty() )
    {
        if ( fParentWidget )
            QMessageBox::information( fParentWidget, "Find Empty Folders", QString( "No empty folders found" ) );
        else
            emit sigStatusMessage( QString( "No empty folders found" ) );
        return false;
    }

    if ( fParentWidget )
    {
        CSelectFolders dlg( fParentWidget );
        dlg.setWindowTitle( "Select Folders to Delete:" );
        dlg.setFolders( allFolders );
        if ( dlg.exec() == QDialog::Accepted )
        {
            auto selectedFolders = dlg.selectedFolders();
            QStringList folderNames;
            for ( auto &&ii : selectedFolders )
            {
                if ( folderIsEmpty( ii ) )
                {
                    folderNames.push_back( ii->FolderPath() );
                }
            }

            auto msg = QString( "%1 folders will be deleted<ul>" ).arg( folderNames.size() );
            msg += "<ul>";
            folderNames.sort();
            msg += getULForList( folderNames ) + "</ul";
            auto process = QMessageBox::information( fParentWidget, "Delete Empty Folders", msg, QMessageBox::Yes | QMessageBox::No );
            if ( process == QMessageBox::No )
                return false;
            for ( auto &&ii : selectedFolders )
            {
                if ( folderIsEmpty( ii ) )
                {
                    deleteFolderAndParentsIfEmpty( ii );
                }
            }
        }
        else
            return false;
    }
    else
    {
        emit sigStatusMessage( QString( "The Following Folders are empty:\n" ) );
        for ( auto &&ii : allFolders )
        {
            emit sigStatusMessage( QString( "    %1:\n" ).arg( ii->FolderPath() ) );
        }
    }
    return true;
}

bool COutlookAPI::moveFromToAddress( bool andSave /*= true*/, bool *needsSaving /*= nullptr*/ )
{
    if ( needsSaving )
        *needsSaving = false;

    if ( !fRules )
        return false;

    slotClearCanceled();

    auto numRules = fRules->Count();
    emit sigInitStatus( "Transforming From to Address Rules:", numRules );
    std::list< std::pair< std::shared_ptr< Outlook::Rule >, TEmailAddressList > > changes;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;

        auto rule = getRule( fRules->Item( ii ) );
        if ( !rule )
            continue;

        emit sigIncStatusValue( "Transforming From to Address Rules:" );

        auto conditions = rule->Conditions();
        if ( !conditions )
            continue;

        auto from = conditions->From();
        if ( !from->Enabled() )
            continue;

        emit sigStatusMessage( QString( "Checking from email addresses on rule '%1'" ).arg( getDisplayName( rule ) ) );
        auto fromEmails = getEmailAddresses( from->Recipients(), {}, EContactTypes::eSMTPContact );
        if ( fromEmails.empty() )
            continue;

        changes.emplace_back( rule, fromEmails );
    }

    if ( changes.empty() )
    {
        if ( fParentWidget )
            QMessageBox::information( fParentWidget, "Transforming From to Address", QString( "No rules needed fixing" ) );
        else
            emit sigStatusMessage( QString( "No rules needed fixing" ) );
        return false;
    }

    if ( fParentWidget )
    {
        QStringList tmp;
        for ( auto &&ii : changes )
        {
            auto curr = QString( "<li style=\"white-space:nowrap\">Rule: %1 will have the following Address(es) now:<br>" ).arg( getDisplayName( ii.first ).toHtmlEscaped() );
            curr += "From:" + getULForList( toStringList( ii.second ) );
            curr += "</li>";

            tmp << curr;
        }

        auto msg = QString( "Rules to be changed:<ul>%1</ul>Do you wish to continue?" ).arg( tmp.join( "\n" ) );
        auto process = QMessageBox::information( fParentWidget, "Renamed Rules", msg, QMessageBox::Yes | QMessageBox::No );
        if ( process == QMessageBox::No )
            return false;
    }
    else
    {
        QStringList tmp;
        for ( auto &&ii : changes )
        {
            emit sigStatusMessage( QString( "Rule: %1 will have the following Address(es):\n" ).arg( getDisplayName( ii.first ) ) );
            emit sigStatusMessage( "    From: \n" + toStringList( ii.second ).join( "    \n" ) );
        }
    }

    QStringList msgs;
    for ( auto &&change : changes )
    {
        auto rule = change.first;
        if ( !rule )
            continue;

        auto conditions = rule->Conditions();
        if ( !conditions )
            continue;

        auto from = conditions->From();
        if ( !from->Enabled() )
            continue;

        auto senderAddress = rule->Conditions()->SenderAddress();
        if ( !senderAddress )
            continue;

        from->SetEnabled( false );
        senderAddress->SetEnabled( false );
        if ( !addRecipientsToRule( rule.get(), change.second, msgs ) )
            return false;
    }

    if ( needsSaving )
        *needsSaving = true;
    if ( andSave )
        saveRules();

    return true;
}

bool COutlookAPI::renameRules( bool andSave /*= true*/, bool *needsSaving /*= nullptr*/ )
{
    if ( needsSaving )
        *needsSaving = false;

    if ( !fRules )
        return false;

    slotClearCanceled();
    auto numRules = fRules->Count();

    emit sigInitStatus( "Analyzing Rule Names:", numRules );

    std::list< std::pair< std::shared_ptr< Outlook::Rule >, QString > > changes;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;

        emit sigIncStatusValue( "Analyzing Rule Names:" );
        auto rule = getRule( fRules->Item( ii ) );
        if ( !rule )
            continue;

        auto ruleName = ruleNameForRule( rule );
        auto currName = rule->Name();
        if ( ruleName != currName )
        {
            changes.emplace_back( rule, ruleName );
        }
    }
    if ( canceled() )
        return false;

    if ( changes.empty() )
    {
        if ( fParentWidget )
            QMessageBox::information( fParentWidget, "Renamed Rules", QString( "No rules needed renaming" ) );
        else
            emit sigStatusMessage( QString( "No rules needed renaming" ) );
        return false;
    }

    if ( fParentWidget )
    {
        QStringList tmp;
        for ( auto &&ii : changes )
        {
            tmp << "<li style=\"white-space:nowrap\">" + getDisplayName( ii.first ).toHtmlEscaped() + " => " + ii.second.toHtmlEscaped() + "</li>";
        }
        auto msg = QString( "Rules to be changed:<ul>%1</ul>Do you wish to continue?" ).arg( tmp.join( "\n" ) );
        auto process = QMessageBox::information( fParentWidget, "Renamed Rules", msg, QMessageBox::Yes | QMessageBox::No );
        if ( process == QMessageBox::No )
            return false;
    }
    else
    {
        for ( auto &&ii : changes )
        {
            emit sigStatusMessage( QString( "Rule '%1' will be renamed to '%2'" ).arg( getDisplayName( ii.first ), ii.second ) );
        }
    }

    emit sigInitStatus( "Renaming Rules:", static_cast< int >( changes.size() ) );
    for ( auto &&ii : changes )
    {
        ii.first->SetName( ii.second );
        emit sigIncStatusValue( "Renaming Rules:" );
    }

    if ( needsSaving )
        *needsSaving = !changes.empty();

    if ( andSave && !changes.empty() )
        saveRules();

    return true;
}

bool COutlookAPI::sortRules( bool andSave /*= true*/, bool *needsSaving /*= nullptr*/ )
{
    if ( needsSaving )
        *needsSaving = false;

    if ( !fRules )
        return false;

    slotClearCanceled();

    auto numRules = fRules->Count();
    emit sigInitStatus( "Sorting Rules:", numRules );

    std::list< Outlook::_Rule * > rules;
    for ( int ii = 1; ii <= numRules; ++ii )
    {
        if ( canceled() )
            return false;
        auto rule = fRules->Item( ii );
        emit sigIncStatusValue( "Sorting Rules:" );
        if ( !rule )
            continue;
        rules.push_back( rule );
    }
    if ( canceled() )
        return false;

    rules.sort(
        []( Outlook::_Rule *lhs, Outlook::_Rule *rhs )
        {
            if ( !lhs )
                return false;
            if ( !rhs )
                return true;
            auto lhsName = lhs->Name();
            auto rhsName = rhs->Name();
            if ( lhsName.startsWith( rhsName ) && ( lhsName != rhsName ) )
                return true;
            else if ( rhsName.startsWith( lhsName ) && ( lhsName != rhsName ) )
                return false;
            else
                return lhsName < rhsName;
        } );

    if ( canceled() )
        return false;
    auto pos = 1;
    emit sigInitStatus( "Recomputing Execution Order:", numRules );

    std::list< std::tuple< QString, Outlook::_Rule *, int > > rulesChanged;
    for ( auto &&ii : rules )
    {
        if ( canceled() )
            return false;

        if ( ii->ExecutionOrder() != pos )
        {
            auto msg = QString( "%1 -> %3" ).arg( getDisplayName( ii ) ).arg( pos );
            if ( !fParentWidget )
                emit sigStatusMessage( msg );
            rulesChanged.emplace_back( msg.toHtmlEscaped(), ii, pos );
        }
        pos++;
        emit sigIncStatusValue( "Recomputing Execution Order:" );
    }

    auto msg = QString( "%1 rules needed re-ordering" ).arg( rulesChanged.size() );
    if ( fParentWidget )
    {
        if ( rulesChanged.empty() )
            QMessageBox::information( fParentWidget, R"(Sorting Rules by Name)", msg );
        else
        {
            msg += "<ul>";
            int cnt = 0;
            for ( auto &&ii : rulesChanged )
            {
                if ( cnt >= 5 )
                {
                    msg += "And More...";
                    break;
                }
                msg += "<li style=\"white-space:nowrap\">" + std::get< 0 >( ii ) + "</li>";
                cnt++;
            }
            msg += "</ul>";
            msg += "Do you wish to Continue?";
            auto process = QMessageBox::information( fParentWidget, R"(Sorting Rules by Name)", msg, QMessageBox::Yes | QMessageBox::No );
            if ( process == QMessageBox::No )
                return false;
        }
    }
    else
        emit sigStatusMessage( msg );

    for ( auto &&ii : rulesChanged )
    {
        auto rule = std::get< 1 >( ii );
        rule->SetExecutionOrder( std::get< 2 >( ii ) );
    }

    if ( needsSaving )
        *needsSaving = rulesChanged.size() != 0;

    if ( andSave && ( rulesChanged.size() != 0 ) )
        saveRules();

    return true;
}

void COutlookAPI::slotHandleRulesSaveException( int, const QString &, const QString &, const QString & )
{
    fSaveRulesSuccess = false;
}

bool COutlookAPI::saveRules()
{
    emit sigStatusMessage( QString( "Saving Rules" ) );
    fSaveRulesSuccess = true;
    connect( fRules.get(), SIGNAL( exception( int, QString, QString, QString ) ), this, SLOT( slotHandleRulesSaveException( int, const QString &, const QString &, const QString & ) ) );
    fRules->Save( fParentWidget ? true : false );
    disconnect( fRules.get(), SIGNAL( exception( int, QString, QString, QString ) ), this, SLOT( slotHandleRulesSaveException( int, const QString &, const QString &, const QString & ) ) );
    return fSaveRulesSuccess;
}
