// xlsxglobal.h

#ifndef XLSXGLOBAL_H
#define XLSXGLOBAL_H

#include <cstdio>
#include <string>
#include <iostream>

#include <QtGlobal>
#include <QObject>
#include <QString>
#include <QVariant>
#include <QIODevice>
#include <QByteArray>
#include <QStringList>

#if defined(QXlsx_SHAREDLIB)
#if defined(QXlsx_EXPORTS)
#  define QXLSX_EXPORT Q_DECL_EXPORT
#else
#  define QXLSX_EXPORT Q_DECL_IMPORT
#endif
#else
#  define QXLSX_EXPORT
#endif

#define QT_BEGIN_NAMESPACE_XLSX namespace QXlsx {
#define QT_END_NAMESPACE_XLSX }

#define QXLSX_USE_NAMESPACE using namespace QXlsx;

//#include <magic_enum_all.hpp>

//using namespace magic_enum::istream_operators;
//using namespace magic_enum::ostream_operators;

//#if MAGIC_ENUM_SUPPORTED
//template<typename T>
//std::optional<T> magic_cast(QStringRef val)
//{
//    return magic_enum::enum_cast<T>(val.toString().toStdString());
//}

//template<typename T>
//QString magic_name(std::optional<T> val)
//{
//    if (val.has_value())
//        return QString::fromStdString(std::string(magic_enum::enum_name(val.value())));
//    return {};
//}

//template<typename T>
//QString magic_name(T val)
//{
//    return QString::fromStdString(std::string(magic_enum::enum_name(val)));
//}
//#else

//template<typename T>
//constexpr std::string_view enum_name(T value) noexcept
//{
//    (void)value;
//    return "";
//}

//template<typename T>
//std::optional<T> magic_cast(QStringRef val)
//{
//    //return magic_enum::enum_cast<T>(val.toString().toStdString());
//    return {};
//}

//template<typename T>
//QString magic_name(std::optional<T> val)
//{
//    if (val.has_value())
//        return QString::fromStdString(std::string(enum_name<T>(val.value())));
//    return {};
//}

//template<typename T>
//QString magic_name(T val)
//{
//    return QString::fromStdString(std::string(enum_name<T>(val)));
//}
//#endif


#endif // XLSXGLOBAL_H
