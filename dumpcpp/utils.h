#ifndef __UTILS_H
#define __UTILS_H

#include <QString>

class QTextStream;
class QMetaEnum;
class QByteArray;

QString stripPrefix( const QString &enumName );
void generateToFromCppEnum( QTextStream &out, const QMetaEnum &metaEnum, const QString &nameSpace );
bool writeToFromStringImpl( QTextStream &implOut );
void updateEnumMap( const QByteArray &key, const QByteArray &metaEnumName );

enum ObjectCategory
{
    DefaultObject = 0x00,
    SubObject = 0x001,
    ActiveX = 0x002,
    NoMetaObject = 0x004,
    NoImplementation = 0x008,
    NoDeclaration = 0x010,
    NoInlines = 0x020,
    OnlyInlines = 0x040,
    Licensed = 0x100,
};

Q_DECLARE_FLAGS( ObjectCategories, ObjectCategory );
Q_DECLARE_OPERATORS_FOR_FLAGS( ObjectCategories );

enum class ProgramMode
{
    GenerateMode,
    TypeLibID
};

struct Options
{
    Options() = default;

    ProgramMode mode{ ProgramMode::GenerateMode };
    ObjectCategories category = DefaultObject;
    bool dispatchEqualsIDispatch = false;
    bool useControlName = false;

    QString outname;
    QString typeLib;
    QString nameSpace;
    QString enumPrefix;
    QString mocExecPath{ QStringLiteral( "moc.exe" ) };
    bool enumClass{ false };
    QByteArray enumToken() const { return enumClass ? "enum class" : "enum"; }
    bool disableClangFormat{ false };
    bool generateToFromEnum{ false };
};

extern Options gOptions;

#endif
