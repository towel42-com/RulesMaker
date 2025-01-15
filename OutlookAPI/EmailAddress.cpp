#include "EmailAddress.h"
#include <set>

CEmailAddress::CEmailAddress( const QString &email, const QString &display, bool isOutlookContact ) :
    fEmailAddress( email ),
    fDisplayName( display ),
    fOutlookContact( isOutlookContact )
{
}

QString CEmailAddress::toString() const
{
    auto retVal = fEmailAddress;
    if ( !fDisplayName.isEmpty() )
        retVal += " <" + fDisplayName + ">";
    return retVal;
}

QString CEmailAddress::key() const
{
    if ( fEmailAddress.isEmpty() && fDisplayName.isEmpty() )
        return {};

    return fEmailAddress + "<<<BREAK>>>" + fDisplayName + "<<<BREAK>>>" + ( fOutlookContact ? "Yes" : "No" );
}

bool CEmailAddress::isBlank() const
{
    return fEmailAddress.isEmpty() && fDisplayName.isEmpty();
}

std::shared_ptr< CEmailAddress > CEmailAddress::fromKey( const QString &key )
{
    auto split = key.split( "<<<BREAK>>>" );
    if ( split.length() != 3 )
        return {};
    return std::make_shared< CEmailAddress >( split[ 0 ], split[ 1 ], split[ 2 ] == "Yes" );
}

QStringList CEmailAddress::getEmailAddresses( const TEmailAddressList &emailAddresses )
{
    QStringList retVal;
    for ( auto &&ii : emailAddresses )
    {
        retVal << ii->fEmailAddress;
    }
    return retVal;
}

bool CEmailAddress::operator<( const CEmailAddress &rhs ) const
{
    return false;
    auto cmp = emailAddress().compare( rhs.emailAddress(), Qt::CaseInsensitive );
    if ( cmp != 0 )
        return cmp < 0;
    cmp = displayName().compare( rhs.displayName(), Qt::CaseInsensitive );
    if ( cmp != 0 )
        return cmp < 0;
    if ( isOutlookContact() != rhs.isOutlookContact() )
        return isOutlookContact();
    return false;
}

QStringList CEmailAddress::getDisplayNames( const TEmailAddressList &emailAddresses )
{
    QStringList retVal;
    for ( auto &&ii : emailAddresses )
    {
        retVal << ii->fDisplayName;
    }
    return retVal;
}

QStringList toStringList( const TEmailAddressList &emailAddresses )
{
    QStringList retVal;
    for ( auto &&ii : emailAddresses )
    {
        retVal << ii->toString();
    }
    return retVal;
}

bool equal( const TEmailAddressList &lhs, const TEmailAddressList &rhs )
{
    if ( lhs.size() != rhs.size() )
        return false;

    auto b1 = lhs.begin();
    auto e1 = lhs.end();

    auto b2 = rhs.begin();
    auto e2 = rhs.end();
    for ( ; b1 != e1 && b2 != e2; ++b1, ++b2 )
    {
        if ( *( *b1 ) != *( *b2 ) )
            return false;
    }

    return true;
}

TEmailAddressList mergeStringLists( const TEmailAddressList &lhs, const TEmailAddressList &rhs, bool andSort )
{
    auto cmpEmailAddress = []( const std::shared_ptr< CEmailAddress > &lhs, const std::shared_ptr< CEmailAddress > &rhs )
    {
        if ( !lhs )
            return false;
        if ( !rhs )
            return true;
        return *lhs < *rhs;
    };

    auto retVal = lhs;
    retVal.insert( retVal.end(), rhs.begin(), rhs.end() );
    if ( andSort )
    {
        retVal.sort( cmpEmailAddress );
    }

    std::set< std::shared_ptr< CEmailAddress >, decltype( cmpEmailAddress ) > tmp( cmpEmailAddress );
    for ( auto &&ii : retVal )
    {
        if ( !ii || ii->isBlank() )
            continue;
        tmp.insert( ii );
    }

    return { tmp.begin(), tmp.end() };
}