#ifndef EMAILADDRESS_H
#define EMAILADDRESS_H

#include <QString>
#include <QStringList>
#include <optional>
#include <memory>

class CEmailAddress;
using TEmailAddressList = std::list< std::shared_ptr< CEmailAddress > >;

class CEmailAddress
{
public:
    CEmailAddress() = default;
    CEmailAddress( const QString &email, const QString &display, bool isOutlookContact );
    ;

    QString toString() const;
    QString key() const;
    bool isBlank() const;

    static std::shared_ptr< CEmailAddress > fromKey( const QString &key );
    static QStringList getDisplayNames( const TEmailAddressList &emailAddresses );
    static QStringList getEmailAddresses( const TEmailAddressList &emailAddresses );

    bool operator==( const CEmailAddress &rhs ) const { return fEmailAddress == rhs.fEmailAddress && fDisplayName == rhs.fDisplayName && fOutlookContact == rhs.fOutlookContact; }
    bool operator!=( const CEmailAddress &rhs ) const { return !operator==( rhs ); }
    bool operator<( const CEmailAddress &rhs ) const;
    
    QString emailAddress() const { return fEmailAddress; }
    QString displayName() const { return fDisplayName; }
    bool isOutlookContact() const{ return fOutlookContact;}
private:
    QString fEmailAddress;
    QString fDisplayName;
    bool fOutlookContact{false};
};

bool equal( const TEmailAddressList &lhs, const TEmailAddressList &rhs );
[[nodiscard]] TEmailAddressList mergeStringLists( const TEmailAddressList &lhs, const TEmailAddressList &rhs, bool andSort = false );
[[nodiscard]] QStringList toStringList( const TEmailAddressList &values );

#endif
