#include "utils.h"

#include <QUuid>
#include <QSettings>
#include <QWidget>
#include <QHash>
#include <QRegularExpression>
#include <QMetaEnum>
#include <unordered_set>

#include <qt_windows.h>
#include <ocidl.h>
#include <map>


Options gOptions;
static QHash< QString, bool > sToStringDeclMap;
static QHash< QString, bool > sToStringImplMap;
static QList< std::pair< QString, QString > > sToStringImplList;
static std::map< QString, QString > sEnumMap;

extern QMetaObject *qax_readClassInfo( ITypeLib *typeLib, ITypeInfo *typeInfo, const QMetaObject *parentObject );
extern QMetaObject *qax_readInterfaceInfo( ITypeLib *typeLib, ITypeInfo *typeInfo, const QMetaObject *parentObject );
extern void qax_deleteMetaObject( QMetaObject *mo );

QString stripPrefix( const QString &enumName )
{
    if ( enumName.isEmpty() || gOptions.enumPrefix.isEmpty() )
        return enumName;
    static std::map< QString, QString > sEnumNameMap;
    auto pos = sEnumNameMap.find( enumName );
    if ( pos != sEnumNameMap.end() )
        return ( *pos ).second;

    auto retVal = enumName;
    while ( retVal.startsWith( QString::fromUtf8( "_" ) ) )
        retVal = retVal.mid( 1 );
    if ( retVal.startsWith( gOptions.enumPrefix, Qt::CaseInsensitive ) )
        retVal = retVal.mid( gOptions.enumPrefix.length() );
    while ( retVal.startsWith( QString::fromUtf8( "_" ) ) )
        retVal = retVal.mid( 1 );
    sEnumNameMap[ enumName ] = retVal;
    return retVal;
}

bool writeToFromStringImpl( QTextStream &implOut )
{
    for ( auto &&ii : sToStringImplList )
    {
        implOut << ii.first << Qt::endl;
        implOut << ii.second << Qt::endl;
    }
    return true;
}

QString getValueNameForEnum( QString valueName )
{
    valueName = stripPrefix( valueName );
    if ( valueName.contains( QString::fromUtf8( "_" ) ) )
    {
        valueName = valueName.toLower();
        auto pos = valueName.indexOf( QString::fromUtf8( "_" ) );
        while ( pos != -1 )
        {
            valueName = valueName.remove( pos, 1 );
            valueName[ pos ] = valueName[ pos ].toUpper();
            pos = valueName.indexOf( QChar::fromLatin1( '_' ), pos );
        }
    }

    if ( !valueName.isEmpty() )
    {
        valueName[ 0 ] = valueName[ 0 ].toLower();
    }
    return valueName;
}

QString getEnumDescriptiveString( const QString &enumName )
{
    if ( enumName.isEmpty() )
        return {};
    auto retVal = getValueNameForEnum( enumName );
    if ( retVal.isEmpty() )
        return {};

    retVal[ 0 ] = retVal[ 0 ].toUpper();

    QStringList words;
    auto regEx = QRegularExpression( QStringLiteral( "[A-Z][a-z]+" ) );
    auto iter = regEx.globalMatch( retVal );
    auto prevEnd = 0;
    while ( iter.hasNext() )
    {
        auto match = iter.next();
        if ( match.capturedStart( 0 ) != prevEnd )
        {
            words << retVal.mid( prevEnd, match.capturedStart( 0 ) - prevEnd );
        }
        auto word = match.captured( 0 );
        words << word;
        prevEnd = match.capturedEnd( 0 );
    }

    retVal = words.join( QStringLiteral( " " ) );

    return retVal;
}

QString getToStringDecl( const QString &enumName, const std::optional< QString > &nameSpace = {} )
{
    QString prefix;
    if ( nameSpace.has_value() )
        prefix = nameSpace.value() + QLatin1String( "::" );

    return QStringLiteral( "QString %1toString( %2 %3 )" ).arg( prefix, enumName, getValueNameForEnum( enumName ) );
}

QString getFromStringDecl( const QString &enumName, const std::optional< QString > &nameSpace = {} )
{
    QString prefix;
    QString templPrefix;
    QString templDecl;
    if ( nameSpace.has_value() )
    {
        prefix = nameSpace.value() + QLatin1String( "::" );
        templDecl = QStringLiteral( "< %1%2 >" ).arg( prefix ).arg( enumName );
    }
    else
        templPrefix = QStringLiteral( "template<> " );

    return QStringLiteral( "%1std::optional< %3%2 > %3fromString%4( const QString & %5Str )" ).arg( templPrefix, enumName, prefix, templDecl, getValueNameForEnum( enumName ) );
}

bool generateToString( QTextStream &out, ITypeLib *typelib, ObjectCategories category )
{
    if ( !typelib )
        return false;

    bool metaObjectFound = false;
    UINT typeCount = typelib->GetTypeInfoCount();
    for ( UINT index = 0; index < typeCount; ++index )
    {
        ITypeInfo *typeinfo = nullptr;
        typelib->GetTypeInfo( index, &typeinfo );
        if ( !typeinfo )
            continue;

        TYPEATTR *typeattr;
        typeinfo->GetTypeAttr( &typeattr );
        if ( !typeattr )
        {
            typeinfo->Release();
            continue;
        }

        TYPEKIND typekind;
        typelib->GetTypeInfoType( index, &typekind );

        ObjectCategories object_category = category;
        if ( !( typeattr->wTypeFlags & TYPEFLAG_FCANCREATE ) )
            object_category |= SubObject;
        else if ( typeattr->wTypeFlags & TYPEFLAG_FCONTROL )
            object_category |= ActiveX;

        QMetaObject *metaObject = 0;
        QUuid guid( typeattr->guid );

        if ( !( object_category & ActiveX ) )
        {
            QSettings settings( QLatin1String( "HKEY_LOCAL_MACHINE\\Software\\Classes\\CLSID\\" ) + guid.toString(), QSettings::NativeFormat );
            if ( settings.childGroups().contains( QLatin1String( "Control" ) ) )
            {
                object_category |= ActiveX;
                object_category &= ~SubObject;
            }
        }

        switch ( typekind )
        {
            case TKIND_COCLASS:
                if ( object_category & ActiveX )
                    metaObject = qax_readClassInfo( typelib, typeinfo, &QWidget::staticMetaObject );
                else
                    metaObject = qax_readClassInfo( typelib, typeinfo, &QObject::staticMetaObject );
                break;
            case TKIND_DISPATCH:
                if ( object_category & ActiveX )
                    metaObject = qax_readInterfaceInfo( typelib, typeinfo, &QWidget::staticMetaObject );
                else
                    metaObject = qax_readInterfaceInfo( typelib, typeinfo, &QObject::staticMetaObject );
                break;
            default:
                break;
        }

        if ( !metaObject )
        {
            typeinfo->ReleaseTypeAttr( typeattr );
            typeinfo->Release();
            continue;
        }

        auto getEnumString = []( QString enumName ) -> QString
        {
            if ( enumName.isEmpty() )
                return {};
            if ( enumName.startsWith( QStringLiteral( "ol" ), Qt::CaseInsensitive ) )
                enumName = enumName.mid( 2 );

            QStringList words;
            auto regEx = QRegularExpression( QStringLiteral( "[A-Z][a-z]+" ) );
            auto iter = regEx.globalMatch( enumName );
            auto prevPos = 0;
            while ( iter.hasNext() )
            {
                auto match = iter.next();
                auto word = match.captured( 0 );
                words << word;
            }

            enumName = words.join( QStringLiteral( " " ) );
            return enumName;
        };

        int allEnumCount = metaObject->enumeratorCount();
        int thisEnumCount = allEnumCount - metaObject->enumeratorOffset();
        if ( thisEnumCount )
        {
            for ( auto ii = metaObject->enumeratorOffset(); ii < allEnumCount; ++ii )
            {
                auto mo = metaObject->enumerator( ii );
                if ( mo.isValid() )
                {
                    auto enumName = QString::fromLocal8Bit( mo.name() );
                    //qDebug() << "    Generating toString for enum " << enumName;
                    out << Qt::endl   //
                        << "QString toString( " << enumName << " value )" << Qt::endl
                        << "{" << Qt::endl
                        << "    switch( value )" << Qt::endl
                        << "    {" << Qt::endl;

                    for ( int j = 0; j < mo.keyCount(); ++j )
                    {
                        out << QStringLiteral( "        case " ) << mo.key( j ) << QStringLiteral( ": return \"" ) << getEnumString( QString::fromLocal8Bit( mo.key( j ) ) ) << QStringLiteral( "\";" ) << Qt::endl;
                    }
                    out << QString::fromLocal8Bit( R"(        default: return "<UNKNOWN-%1>";)" ).arg( enumName ) << Qt::endl   //
                        << "    };" << Qt::endl
                        << "};" << Qt::endl;
                }
            }
            metaObjectFound = true;
            out.flush();
        }
        qax_deleteMetaObject( metaObject );
        typeinfo->ReleaseTypeAttr( typeattr );
        typeinfo->Release();
        break;
    }
    return metaObjectFound;
}

std::optional< QString > generateToString( const QMetaEnum &metaEnum, const QString &nameSpace )
{
    if ( !metaEnum.isValid() )
        return {};

    QString retVal;
    QTextStream out( &retVal );

    //qDebug() << "    Generating toString for enum " << enumName;

    auto enumName = QString::fromLocal8Bit( metaEnum.name() );
    out << getToStringDecl( enumName, nameSpace ) << Qt::endl << "{" << Qt::endl << "    switch( " << getValueNameForEnum( enumName ) << " )" << Qt::endl << "    {" << Qt::endl;

    std::unordered_set< int > valuesUsed;

    for ( int ii = 0; ii < metaEnum.keyCount(); ++ii )
    {
        auto enumString = getEnumDescriptiveString( QString::fromLocal8Bit( metaEnum.key( ii ) ) );
        auto key = enumName + QLatin1String( "::" ) + QString::fromLocal8Bit( metaEnum.key( ii ) );
        auto value = metaEnum.value( ii );

        out << "        ";
        if ( valuesUsed.find( value ) != valuesUsed.end() )
            out << "//";
        valuesUsed.insert( value );
        out << "case " << key << ": return \"" << enumString << "\";" << Qt::endl;
    }

    out << QStringLiteral( R"(        default: return "<UNKNOWN-%1>";)" ).arg( enumName ) << Qt::endl   //
        << "    };" << Qt::endl
        << "};" << Qt::endl;

    return retVal;
}

std::optional< QString > generateFromString( const QMetaEnum &metaEnum, const QString &nameSpace )
{
    QString retVal;
    QTextStream out( &retVal );

    auto enumName = metaEnum.name();
    out << getFromStringDecl( QString::fromLocal8Bit( enumName ), nameSpace ) << Qt::endl   //
        << "{" << Qt::endl
        << "    static QHash< QString, " << enumName << " > sEnumMap;" << Qt::endl
        << "    if ( sEnumMap.isEmpty() )" << Qt::endl   //
        << "    {" << Qt::endl;

    for ( int ii = 0; ii < metaEnum.keyCount(); ++ii )
    {
        auto enumKey = QString::fromLocal8Bit( metaEnum.key( ii ) );
        auto strippedKey = stripPrefix( enumKey );
        auto descriptiveKey = getEnumDescriptiveString( enumKey );

        // dont use a set for uniquifying since we want consistant order based on the enum
        auto keys = QStringList() << enumKey.toLower() << strippedKey.toLower() << descriptiveKey.toLower();
        keys.removeDuplicates();

        auto enumValue = QString::fromLocal8Bit( enumName ) + QLatin1String( "::" ) + QString::fromLocal8Bit( metaEnum.key( ii ) );

        if ( ii )
            out << Qt::endl;
        for ( auto &&ii : keys )
        {
            out << "        "
                << "sEnumMap[\"" << ii << "\"] = " << enumValue << ";" << Qt::endl;
        }
    }

    out << "    }" << Qt::endl;   //

    out << Qt::endl
        << "    auto pos = sEnumMap.find( " << getValueNameForEnum( QString::fromLocal8Bit( enumName ) ) << "Str.toLower() );" << Qt::endl   //
        << "    if ( pos != sEnumMap.end() )" << Qt::endl   //
        << "        return pos.value();" << Qt::endl   //
        << "    return {};" << Qt::endl   //
        << "};" << Qt::endl;

    return retVal;
}

void generateToFromCppEnum( QTextStream &out, const QMetaEnum &metaEnum, const QString &nameSpace )
{
    auto metaEnumName = QString::fromLocal8Bit( metaEnum.name() );
    if ( !sToStringDeclMap.contains( metaEnumName ) )
    {
        sToStringDeclMap[ metaEnumName ] = true;
        out << "    " << getToStringDecl( metaEnumName ) << ";" << Qt::endl;
        out << "    " << getFromStringDecl( metaEnumName ) << ";" << Qt::endl;
    }

    if ( !sToStringImplMap.contains( metaEnumName ) )
    {
        auto toStringStr = generateToString( metaEnum, nameSpace );
        auto fromStringStr = generateFromString( metaEnum, nameSpace );
        if ( toStringStr.has_value() && fromStringStr.has_value() )
        {
            sToStringImplMap[ metaEnumName ] = true;
            sToStringImplList << std::make_pair( toStringStr.value(), fromStringStr.value() );
        }
    }
    out << Qt::endl;
}

void updateEnumMap( const QByteArray &key, const QByteArray &metaEnumName )
{
    auto enumMapKey = gOptions.nameSpace + QStringLiteral( "::" ) + QString::fromLatin1( key );
    auto enumMapValue = gOptions.nameSpace;
    if ( gOptions.enumClass )
    {
        enumMapValue += QStringLiteral( "::" ) + QString::fromLatin1( metaEnumName );
    }
    enumMapValue += QStringLiteral( "::" ) + QString::fromLatin1( key );

    sEnumMap[ enumMapKey ] = enumMapValue;
}