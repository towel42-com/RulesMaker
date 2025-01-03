#include "OutlookAPI.h"
#include <QVariant>

#include "MSOUTL.h"

QString toString( const QVariant &variant, const QString &joinSeparator )
{
    auto stringList = toStringList( variant );
    auto retVal = stringList.join( joinSeparator );
    return retVal;
}

QStringList toStringList( const QVariant &variant )
{
    QStringList retVal;
    if ( variant.type() == QVariant::Type::String )
        retVal << variant.toString();
    else if ( variant.type() == QVariant::Type::StringList )
        retVal << variant.toStringList();
    return retVal;
}

QStringList mergeStringLists( const QStringList &lhs, const QStringList &rhs, bool andSort )
{
    auto retVal = QStringList() << lhs << rhs;
    if ( andSort )
        retVal.sort( Qt::CaseInsensitive );

    retVal.removeDuplicates();
    retVal.removeAll( QString() );

    return retVal;
}

QString toString( EFilterType filterType )
{
    switch ( filterType )
    {
        case EFilterType::eByEmailAddress:
            return "Email Address";
        case EFilterType::eByDisplayName:
            return "Display Name";
        case EFilterType::eBySubject:
            return "Subject";
        default:
            break;
    }
    return "<UNKNOWN>";
}