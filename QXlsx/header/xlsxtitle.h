#ifndef XLSXTITLE_H
#define XLSXTITLE_H

#include <QFont>
#include <QColor>
#include <QByteArray>
#include <QList>
#include <QVariant>
#include <QXmlStreamWriter>
#include <QSharedDataPointer>

#include "xlsxglobal.h"
#include "xlsxcolor.h"
#include "xlsxfillformat.h"
#include "xlsxlayout.h"
#include "xlsxshapeformat.h"
#include "xlsxtext.h"
#include "xlsxcellrange.h"

QT_BEGIN_NAMESPACE_XLSX

class TitlePrivate;
class AbstractSheet;

/**
 * @brief The Title class is used to specify the chart title or the axis title properties.
 *
 * The class is used to specify the title's:
 * - Text, either as a reference to the other cell or as a formatted text.
 * - Layout (positioning of a chart / axis title on screen)
 * - overlay (positioning of a chart / axis title on screen)
 * - ShapeFormat properties (f.e. line and fill)
 * - text and paragraph properties.
 *
 * Title can be used in three different scenarios:
 *
 * 1. Set the rich-formatted title using setHtml().
 *
 * @code
 * Title title;
 * title.setHtml("<b>Graph</b> of <i>temperatures</i>");
 * //and optionally set text, paragraph and character properties
 *
 * @endcode
 *
 * 2. Set the plain string and default paragraph and character properties
 *
 * @code
 * Title title("Graph of temperatures");
 * title.defaultParagraphProperties().leftMargin = Coordinate(6.0); // 6 pt
 * title.defaultCharacterProperties().fontSize = 20.0; // 20 pt
 * @endcode
 *
 * 3. Set the string reference and default paragraph and character properties
 *
 * @code
 * Title title;
 * title.setStringReference("Sheet1!$C$2");
 * title.defaultParagraphProperties().leftMargin = Coordinate(6.0); // 6 pt
 * title.defaultCharacterProperties().fontSize = 20.0; // 20 pt
 * title.defaultCharacterProperties().bold = true; // 20 pt
 * @endcode
 *
 */
class QXLSX_EXPORT Title
{
public:
    Title();
    /**
     * @brief Title creates new Title and sets plain text to it.
     * @param text
     */
    Title(const QString &text);
    Title(const Title &other);
    Title &operator=(const Title &other);
    ~Title();

    QString toPlainText() const;
    void setPlainText(const QString &text);

    QString toHtml() const;
    void setHtml(const QString &formattedText);

    QString stringReference() const;
    /**
     * @brief setStringReference sets title as string reference.
     * @note the method does not check reference for validity.
     * @param reference a string like "Sheet1!A1".
     */
    void setStringReference(const QString &reference);
    /**
     * @brief setStringReference sets title as string reference.
     * @param range reference range
     * @param sheet pointer to a sheet that contains the range.
     */
    void setStringReference(const CellRange &range, AbstractSheet *sheet);



    bool isRichString() const; //only checks type, ignores actual formatting
    bool isPlainString() const;
    bool isStringReference() const;

    Text &text();
    Text text() const;
    void setText(const Text &text);

    TextProperties textProperties() const;
    TextProperties &textProperties();
    void setTextProperties(const TextProperties &textProperties);

    ParagraphProperties defaultParagraphProperties() const;
    ParagraphProperties &defaultParagraphProperties();
    void setDefaultParagraphProperties(const ParagraphProperties &defaultParagraphProperties);

    CharacterProperties defaultCharacterProperties() const;
    CharacterProperties &defaultCharacterProperties();
    void setDefaultCharacterProperties(const CharacterProperties &defaultCharacterProperties);

    /**
     * @brief layout returns the title positioning layout.
     * @return shallow copy of the title Layout.
     */
    Layout layout() const;
    Layout &layout();
    void setLayout(const Layout &layout);
    /**
     * @brief moveTo moves the title layout to the specified point.
     * @param point (0,0) is the top left corner, (1,1) is the bottom right corner.
     */
    void moveTo(const QPointF &point);

    /**
     * @brief overlay This element specifies that other chart elements shall be allowed to overlap this chart element.
     * @return true if this chart element can be overlapped by other chart elements, false otherwise.
     */
    std::optional<bool> overlay() const;
    /**
     * @brief setOverlay specifies that other chart elements shall be allowed to overlap this chart element.
     * @param overlay true if this chart element can be overlapped by other chart elements, false otherwise.
     */
    void setOverlay(bool overlay);

    ShapeFormat shape() const;
    ShapeFormat &shape();
    void setShape(const ShapeFormat &shape);

    bool isValid() const;

    void write(QXmlStreamWriter &writer) const;
    void read(QXmlStreamReader &reader);

    bool operator ==(const Title &other) const;
    bool operator !=(const Title &other) const;

    operator QVariant() const;

private:
    QSharedDataPointer<TitlePrivate> d;
    friend QDebug operator<<(QDebug dbg, const Title &f);
};

QDebug operator<<(QDebug dbg, const Title &f);

//TODO: add all documentation

QT_END_NAMESPACE_XLSX

Q_DECLARE_METATYPE(QXlsx::Title)

#endif // XLSXTITLE_H
