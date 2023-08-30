// xlsxabstractsheet.cpp

#include <QtGlobal>

#include "xlsxabstractsheet.h"
#include "xlsxabstractsheet_p.h"
#include "xlsxworkbook.h"

namespace QXlsx {

AbstractSheetPrivate::AbstractSheetPrivate(AbstractSheet *p, AbstractSheet::CreateFlag flag)
    : AbstractOOXmlFilePrivate(p, flag)
{
    type = AbstractSheet::Type::Worksheet;
    sheetState = AbstractSheet::Visibility::Visible;
}

AbstractSheetPrivate::~AbstractSheetPrivate()
{
}



/*!
 * \internal
 */
AbstractSheet::AbstractSheet(const QString &name, int id, Workbook *workbook, AbstractSheetPrivate *d) :
    AbstractOOXmlFile(d)
{
    d_func()->name = name;
    d_func()->id = id;
    d_func()->workbook = workbook;
}


QString AbstractSheet::name() const
{
    Q_D(const AbstractSheet);
    return d->name;
}

/*!
 * \internal
 */
void AbstractSheet::setName(const QString &sheetName)
{
    Q_D(AbstractSheet);
    d->name = sheetName;
}

/*!
 * Returns the type of the sheet.
 */
AbstractSheet::Type AbstractSheet::type() const
{
    Q_D(const AbstractSheet);
    return d->type;
}

/*!
 * \internal
 */
void AbstractSheet::setType(Type type)
{
    Q_D(AbstractSheet);
    d->type = type;
}

AbstractSheet::Visibility AbstractSheet::visibility() const
{
    Q_D(const AbstractSheet);
    return d->sheetState;
}

void AbstractSheet::setVisibility(Visibility visibility)
{
    Q_D(AbstractSheet);
    d->sheetState = visibility;
}

bool AbstractSheet::isHidden() const
{
    Q_D(const AbstractSheet);
    return d->sheetState != Visibility::Visible;
}

bool AbstractSheet::isVisible() const
{
    return !isHidden();
}

void AbstractSheet::setHidden(bool hidden)
{
    Q_D(AbstractSheet);
    if (!hidden) d->sheetState = Visibility::Visible;
    else if (d->sheetState == Visibility::Visible) d->sheetState = Visibility::Hidden;
}

void AbstractSheet::setVisible(bool visible)
{
    Q_D(AbstractSheet);
    if (visible) d->sheetState = Visibility::Visible;
    else if (d->sheetState == Visibility::Visible) d->sheetState = Visibility::Hidden;
}

/*!
 * \internal
 */
int AbstractSheet::id() const
{
    Q_D(const AbstractSheet);
    return d->id;
}

/*!
 * \internal
 */
Drawing *AbstractSheet::drawing() const
{
    Q_D(const AbstractSheet);
    return d->drawing.get();
}

/*!
 * Return the workbook
 */
Workbook *AbstractSheet::workbook() const
{
    Q_D(const AbstractSheet);
    return d->workbook;
}

}
