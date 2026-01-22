#ifndef __METAUTILS_H
#define __METAUTILS_H

#include <optional>
struct QMetaObject;
class QMetaEnum;
class QMetaProperty;
class QMetaMethod;
class QMetaClassInfo;
class QMetaType;
class QString;

void dumpMetaObject( const QMetaObject &metaObject, int indent  );
void dumpMetaEnum( const QMetaEnum &metaEnum, std::optional< int > num, int indent  );
void dumpProperty( const QMetaProperty &metaProperty, std::optional< int > num, const QString & type, int indent  );
void dumpMethod( const QMetaMethod &metaMethod, std::optional< int > num, const QString &type, int indent  );
void dumpClassInfo( const QMetaClassInfo &metaClassInfo, std::optional< int > num, int indent  );
void dumpMetaType( const QMetaType &metaClassInfo, std::optional< int > num, int indent  );

#endif
