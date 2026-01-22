#include "MetaUtils.h"

#include <QString>
#include <QMetaObject>
#include <QMetaEnum>
#include <QMetaProperty>
#include <QMetaMethod>
#include <QMetaClassInfo>
#include <QMetaType>

inline QString indentText( int indent )
{
    return QString( 4 * indent, QChar::fromLatin1( ' ' ) );
}

void dumpProperty( const QMetaProperty &metaProperty, std::optional< int > num, const QString & type, int indent )
{
    if ( !metaProperty.isValid() )
        return;

    qDebug().nospace().noquote()   //
        << indentText( indent ) << "QMetaProperty: " << type << " " << ( num.has_value() ? std::to_string( num.value() ) : std::string() ) << "\n"   //
        << indentText( indent + 1 ) << "name() = " << metaProperty.name() << "\n"   //
        << indentText( indent + 1 ) << "typeId() = " << metaProperty.typeId() << "\n"   //
#pragma warning( push )
#pragma warning( disable:4996 )
        << indentText( indent + 1 ) << "type() = " << metaProperty.type() << "\n"   //
#pragma warning( pop )
        << indentText( indent + 1 ) << "typeName() = " << metaProperty.typeName() << "\n"   //
        << indentText( indent + 1 ) << "userType() = " << metaProperty.userType() << "\n"   //
        << indentText( indent + 1 ) << "isEnumType() = " << metaProperty.isEnumType() << "\n"   //
        << indentText( indent + 1 ) << "revision() = " << metaProperty.revision() << "\n"   //
        << indentText( indent + 1 ) << "hasNotifySignal() = " << metaProperty.hasNotifySignal() << "\n"   //
        << indentText( indent + 1 ) << "isBindable() = " << metaProperty.isBindable() << "\n"   //
        << indentText( indent + 1 ) << "isConstant() = " << metaProperty.isConstant() << "\n"   //
        << indentText( indent + 1 ) << "isDesignable() = " << metaProperty.isDesignable() << "\n"   //
        << indentText( indent + 1 ) << "isEnumType() = " << metaProperty.isEnumType() << "\n"   //
        << indentText( indent + 1 ) << "isFinal() = " << metaProperty.isFinal() << "\n"   //
        << indentText( indent + 1 ) << "isFlagType() = " << metaProperty.isFlagType() << "\n"   //
        << indentText( indent + 1 ) << "isReadable() = " << metaProperty.isReadable() << "\n"   //
        << indentText( indent + 1 ) << "isRequired() = " << metaProperty.isRequired() << "\n"   //
        << indentText( indent + 1 ) << "isResettable() = " << metaProperty.isResettable() << "\n"   //
        << indentText( indent + 1 ) << "isScriptable() = " << metaProperty.isScriptable() << "\n"   //
        << indentText( indent + 1 ) << "isStored() = " << metaProperty.isStored() << "\n"   //
        << indentText( indent + 1 ) << "isUser() = " << metaProperty.isUser() << "\n"   //
        << indentText( indent + 1 ) << "isValid() = " << metaProperty.isValid() << "\n"   //
        << indentText( indent + 1 ) << "isWritable() = " << metaProperty.isWritable() << "\n"   //
        << indentText( indent + 1 ) << "metaType() = " << metaProperty.metaType() << "\n"   //

        << indentText( indent + 1 ) << "notifySignalIndex() = " << metaProperty.notifySignalIndex() << "\n"   //
        << indentText( indent + 1 ) << "propertyIndex() = " << metaProperty.propertyIndex() << "\n"   //
        << indentText( indent + 1 ) << "relativePropertyIndex() = " << metaProperty.relativePropertyIndex() << "\n"   //
        ;
    dumpMetaType( metaProperty.metaType(), {}, indent + 1 );
}

void dumpMetaType( const QMetaType &metaType, std::optional< int > num, int indent )
{
    if ( !metaType.isValid() )
        return;

    qDebug().nospace().noquote()   //
        << indentText( indent ) << "QMetaType: " << ( num.has_value() ? std::to_string( num.value() ) : std::string() ) << "\n"   //
        << indentText( indent + 1 ) << "name() = " << metaType.name() << "\n"   //
        << indentText( indent + 1 ) << "id() = " << metaType.id() << "\n"   //
        << indentText( indent + 1 ) << "alignOf() = " << metaType.alignOf() << "\n"   //
        << indentText( indent + 1 ) << "flags() = " << metaType.flags() << "\n"   //
        << indentText( indent + 1 ) << "hasRegisteredDataStreamOperators() = " << metaType.hasRegisteredDataStreamOperators() << "\n"   //
        << indentText( indent + 1 ) << "isCopyConstructible() = " << metaType.isCopyConstructible() << "\n"   //
        << indentText( indent + 1 ) << "isDefaultConstructible() = " << metaType.isDefaultConstructible() << "\n"   //
        << indentText( indent + 1 ) << "isDestructible() = " << metaType.isDestructible() << "\n"   //
        << indentText( indent + 1 ) << "isEqualityComparable() = " << metaType.isEqualityComparable() << "\n"   //
        << indentText( indent + 1 ) << "isMoveConstructible() = " << metaType.isMoveConstructible() << "\n"   //
        << indentText( indent + 1 ) << "isOrdered() = " << metaType.isOrdered() << "\n"   //
        << indentText( indent + 1 ) << "isRegistered() = " << metaType.isRegistered() << "\n"   //
        << indentText( indent + 1 ) << "isValid() = " << metaType.isValid() << "\n"   //
        << indentText( indent + 1 ) << "isOrdered() = " << metaType.isOrdered() << "\n"   //
        << indentText( indent + 1 ) << "sizeOf() = " << metaType.sizeOf() << "\n"   //
        ;
    //if ( metaType.metaObject() )
    //    dumpMetaObject( *metaType.metaObject(), indent + 1 );

    if ( ( metaType.underlyingType() != metaType ) )
    {
        dumpMetaType( metaType.underlyingType(), {}, indent + 1 );
    }
}

void dumpMethod( const QMetaMethod &metaMethod, std::optional< int > num, const QString &type, int indent )
{
    if ( !metaMethod.isValid() )
        return;

    qDebug().nospace().noquote()   //
        << indentText( indent ) << "QMetaMethod: " << type << " " << ( num.has_value() ? std::to_string( num.value() ) : std::string() ) << "\n"
        << indentText( indent + 1 ) << "name() = " << metaMethod.name() << "\n"   //
        << indentText( indent + 1 ) << "isConst() = " << metaMethod.isConst() << "\n"   //
        << indentText( indent + 1 ) << "isValid() = " << metaMethod.isValid() << "\n"   //
        << indentText( indent + 1 ) << "methodIndex() = " << metaMethod.methodIndex() << "\n"   //
        << indentText( indent + 1 ) << "methodSignature() = " << metaMethod.methodSignature() << "\n"   //
        << indentText( indent + 1 ) << "methodType() = " << metaMethod.methodType() << "\n"   //
        << indentText( indent + 1 ) << "methodSignature() = " << metaMethod.methodSignature() << "\n"   //
        << indentText( indent + 1 ) << "parameterNames() = " << metaMethod.parameterNames() << "\n"   //
        << indentText( indent + 1 ) << "parameterTypes() = " << metaMethod.parameterTypes() << "\n"   //
        << indentText( indent + 1 ) << "relativeMethodIndex() = " << metaMethod.relativeMethodIndex() << "\n"   //
        << indentText( indent + 1 ) << "returnMetaType() = " << metaMethod.returnMetaType() << "\n"   //
        << indentText( indent + 1 ) << "returnType() = " << metaMethod.returnType() << "\n"   //
        << indentText( indent + 1 ) << "revision() = " << metaMethod.revision() << "\n"   //
        << indentText( indent + 1 ) << "tag() = " << metaMethod.tag() << "\n"   //
        << indentText( indent + 1 ) << "typeName() = " << metaMethod.typeName() << "\n"   //
        << indentText( indent + 1 ) << "tag() = " << metaMethod.tag() << "\n"   //
        << indentText( indent + 1 ) << "parameterCount() = " << metaMethod.parameterCount() << "\n"   //
        << indentText( indent + 1 ) << "parameterTypes() = " << metaMethod.parameterTypes() << "\n"   //
        ;
    for ( int ii = 0; ii < metaMethod.parameterCount(); ++ii )
    {
        qDebug().nospace().noquote()   //
            << indentText( indent + 1 ) << "parameterType(" << ii << ") = " << metaMethod.parameterType( ii ) << "\n"   //
            << indentText( indent + 1 ) << "parameterTypeName(" << ii << ") = " << metaMethod.parameterTypeName( ii ) << "\n"   //
            ;

        dumpMetaType( metaMethod.parameterMetaType( ii ), ii, indent + 1 );
    }
}

void dumpClassInfo( const QMetaClassInfo &metaClassInfo, std::optional< int > num, int indent )
{
    qDebug().nospace().noquote()   //
        << indentText( indent ) << "QMetaClassInfo: " << ( num.has_value() ? std::to_string( num.value() ) : std::string() ) << "\n"   //
        << indentText( indent + 1 ) << "name() = " << metaClassInfo.name() << "\n"   //
        << indentText( indent + 1 ) << "value() = " << metaClassInfo.value() << "\n"   //
        ;
}

void dumpMetaEnum( const QMetaEnum &metaEnum, std::optional< int > num, int indent )
{
    if ( !metaEnum.isValid() )
        return;

    qDebug().nospace().noquote()   //
        << indentText( indent ) << "QMetaEnum: " << ( num.has_value() ? std::to_string( num.value() ) : std::string() ) << "\n"   //
        << indentText( indent + 1 ) << "name() = " << metaEnum.name() << "\n"   //
        << indentText( indent + 1 ) << "enumName() = " << metaEnum.enumName() << "\n"   //
        << indentText( indent + 1 ) << "scope() = " << metaEnum.scope() << "\n"   //
        << indentText( indent + 1 ) << "isFlag() = " << metaEnum.isFlag() << "\n"   //
        << indentText( indent + 1 ) << "isScoped() = " << metaEnum.isScoped() << "\n"   //
        << indentText( indent + 1 ) << "isValid() = " << metaEnum.isValid() << "\n"   //
        << indentText( indent + 1 ) << "enumName() = " << metaEnum.enumName() << "\n"   //
        << indentText( indent + 1 ) << "keyCount() = " << metaEnum.keyCount() << "\n"   //
        ;

    for ( int ii = 0; ii < metaEnum.keyCount(); ++ii )
    {
        qDebug().nospace().noquote()   //
            << indentText( indent + 1 ) << "key(" << ii << ") = " << metaEnum.key( ii ) << "\n"   //
            ;
    }
    dumpMetaType( metaEnum.metaType(), {}, indent + 1 );
}

void dumpMetaObject( const QMetaObject &metaObject, int indent )
{
    qDebug().nospace().noquote()   //
        << indentText( indent ) << "QMetaObject: \n"   //
        << indentText( indent + 1 ) << "className() = " << metaObject.className() << "\n"   //
        ;
    dumpMetaType( metaObject.metaType(), {}, indent + 1 );

    qDebug().nospace().noquote()   //
        << indentText( indent + 1 ) << "constructorCount()  = " << metaObject.constructorCount() << "\n"   //
        ;
    for ( int ii = 0; ii < metaObject.constructorCount(); ++ii )
    {
        dumpMethod( metaObject.constructor( ii ), ii, QStringLiteral( "Constructor" ), indent + 1 );
    }

    qDebug().nospace().noquote()   //
        << indentText( indent + 1 ) << "enumeratorCount()  = " << metaObject.enumeratorCount() << "\n"   //
        << indentText( indent + 1 ) << "enumeratorOffset()  = " << metaObject.enumeratorOffset() << "\n"   //
        ;
    for ( int ii = 0; ii < metaObject.enumeratorCount(); ++ii )
    {
        dumpMetaEnum( metaObject.enumerator( ii ), ii, indent + 1 );
    }

    qDebug().nospace().noquote()   //
        << indentText( indent + 1 ) << "methodCount()  = " << metaObject.methodCount() << "\n"   //
        << indentText( indent + 1 ) << "methodOffset()  = " << metaObject.methodOffset() << "\n"   //
        ;
    for ( int ii = 0; ii < metaObject.methodCount(); ++ii )
    {
        dumpMethod( metaObject.method( ii ), ii, QStringLiteral( "Method" ), indent + 1 );
    }

    qDebug().nospace().noquote()   //
        << indentText( indent + 1 ) << "propertyCount()  = " << metaObject.propertyCount() << "\n"   //
        << indentText( indent + 1 ) << "propertyOffset()  = " << metaObject.propertyOffset() << "\n"   //
        ;
    for ( int ii = 0; ii < metaObject.propertyCount(); ++ii )
    {
        dumpProperty( metaObject.property( ii ), ii, QStringLiteral( "Object Property" ), indent + 1 );
    }

    qDebug().nospace().noquote()   //
        << indentText( indent + 1 ) << "classInfoCount() = " << metaObject.classInfoCount() << "\n"   //
        << indentText( indent + 1 ) << "classInfoOffset() = " << metaObject.classInfoOffset() << "\n"   //
        ;
    for ( int ii = 0; ii < metaObject.classInfoCount(); ++ii )
    {
        dumpClassInfo( metaObject.classInfo( ii ), ii, indent + 1 );
    }

    dumpProperty( metaObject.userProperty(), {}, QStringLiteral( "User Property" ), indent + 1 );
}
