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

namespace QXlsx {

class TitlePrivate;
class AbstractSheet;

/**
 * @brief The Title class represents the chart title, the displayed units label,
 * the trend line label or the axis title.
 *
 * The class is used to specify the title's:
 * - Text, either as a reference to other cell(s) or as a formatted text.
 * - Layout (positioning of a chart / axis title on screen)
 * - overlay (laying the title over a chart)
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
 * You can also create a title piece-by-piece using the low-level methods of Text class.
 *
 * Title is _implicitly shareable_: the deep copy occurs only in the non-const methods.
 */
class QXLSX_EXPORT Title
{
public:
    /**
     * @brief creates an invalid title
     */
    Title();
    /**
     * @brief creates new Title and sets plain text to it.
     * @param text a plain string.
     */
    Title(const QString &text);
    Title(const Title &other);
    Title &operator=(const Title &other);
    ~Title();

    /**
     * @brief returns the title's text as plain string
     * @return
     */
    QString toPlainString() const;
    /**
     * @brief sets the title's text as plain string and _clears_ the formatting.
     * @param text a plain string.
     */
    void setPlainString(const QString &text);

    /**
     * @brief returns the title's text as an html string (not implemented yet).
     * @return
     */
    QString toHtml() const;
    /**
     * @brief sets the title's text as formatted text.
     * @param formattedText
     */
    void setHtml(const QString &formattedText);

    /**
     * @brief returns the title's text as a string reference.
     * @return valid string if the title is reference text, empty string otherwise.
     */
    QString stringReference() const;
    /**
     * @brief sets the title's text as a string reference.
     * @note the method does not check reference for validity.
     * @param reference a string like "Sheet1!A1".
     */
    void setStringReference(const QString &reference);
    /**
     * @brief sets title's text as string reference.
     * @param range reference range
     * @param sheet pointer to a sheet that contains the range.
     */
    void setStringReference(const CellRange &range, AbstractSheet *sheet);


    /**
     * @brief returns whether the title's text is a formatted text.
     * @return true if the title's text is set as an html text, false otherwise.
     * @note this method does not check if the formatting is actually present in the text. It merely
     * checks the Text::type().
     */
    bool isRichString() const; //only checks type, ignores actual formatting
    /**
     * @brief returns whether the title's text is a plain text without formatting.
     * @return true if the title's text was set as a plain text.
     * @note this method does not check if the formatting is actually absent in the text. It merely
     * checks the Text::type().
     */
    bool isPlainString() const;
    /**
     * @brief returns whether the title's text is a string reference.
     * @return true if the title's text was set as a string reference.
     * @note this method merely checks the Text::type().
     */
    bool isStringReference() const;

    /**
     * @brief returns a reference to the title's text.
     * This reference allows to use the low-level methods of the Text class to add, remove or modify
     * text fragments with specific formatting.
     * @return
     */
    Text &text();
    /**
     * @brief returns a copy of the title's text.
     * @return
     */
    Text text() const;
    /**
     * @brief sets the title's text.
     * @param text
     */
    void setText(const Text &text);

    /**
     * @brief returns a copy of the title's text properties.
     * Text properties define the overall layout of the text box within its bounding rectangle.
     * @return
     */
    TextProperties textProperties() const;
    /**
     * @brief returns a reference to the title's text properties.
     * Text properties define the overall layout of the text box within its bounding rectangle.
     * @return
     */
    TextProperties &textProperties();
    /**
     * @brief sets the title's text properties.
     * @param textProperties
     */
    void setTextProperties(const TextProperties &textProperties);

    TextFormat textFormat() const;
    TextFormat &textFormat();
    void setTextFormat(const TextFormat &textFormat);

    /**
     * @brief returns the default paragraph properties of the title.
     * If you want to add several paragraphs with different properties, use the Text class methods.
     * @return
     */
    ParagraphProperties defaultParagraphProperties() const;
    /**
     * @brief returns the default paragraph properties of the title.
     * If you want to add several paragraphs with different properties, use the Text class methods.
     * @return
     */
    ParagraphProperties &defaultParagraphProperties();
    /**
     * @brief sets the default paragraph properties of the title.
     */
    void setDefaultParagraphProperties(const ParagraphProperties &defaultParagraphProperties);

    /**
     * @brief returns the default character properties of the title.
     * If you want to add several fragments with different properties, use the Text class methods.
     * @return
     */
    CharacterProperties defaultCharacterProperties() const;
    /**
     * @brief returns the default character properties of the title.
     * If you want to add several fragments with different properties, use the Text class methods.
     * @return
     */
    CharacterProperties &defaultCharacterProperties();
    void setDefaultCharacterProperties(const CharacterProperties &defaultCharacterProperties);

    /**
     * @brief returns the title's positioning layout.
     * @return a copy of the title Layout.
     */
    Layout layout() const;
    /**
     * @brief returns a reference to the title's positioning layout.
     * @return
     */
    Layout &layout();
    /**
     * @brief sets the title's positioning layout.
     * @param layout
     */
    void setLayout(const Layout &layout);
    /**
     * @brief moves the title to the specified point.
     * @param point (0,0) is the top left corner, (1,1) is the bottom right corner.
     *
     * To fine-tune the title position use the Layout methods.
     */
    void moveTo(const QPointF &point);

    /**
     * @brief returns whether other chart elements shall be allowed to overlap this chart element.
     * @return true if this chart element can be overlapped by other chart elements, false otherwise.
     */
    std::optional<bool> overlay() const;
    /**
     * @brief specifies whether other chart elements shall be allowed to overlap this chart element.
     * @param overlay true if this chart element can be overlapped by other chart elements, false otherwise.
     */
    void setOverlay(bool overlay);
    /**
     * @brief returns the title's shape parameters.
     * @return a copy of the title's shape.
     */
    ShapeFormat shape() const;
    /**
     * @brief returns a reference to the title's shape parameters.
     * @return
     */
    ShapeFormat &shape();
    /**
     * @brief sets the title's shape.
     * @param shape
     */
    void setShape(const ShapeFormat &shape);
    /**
     * @brief returns whether the title has any parameters set.
     * @return true if any of the title's parameters are valid (set).
     */
    bool isValid() const;

    void write(QXmlStreamWriter &writer, const QString &name) const;
    void read(QXmlStreamReader &reader);

    bool operator ==(const Title &other) const;
    bool operator !=(const Title &other) const;

    operator QVariant() const;

private:
    QSharedDataPointer<TitlePrivate> d;
    friend QDebug operator<<(QDebug dbg, const Title &f);
};

QDebug operator<<(QDebug dbg, const Title &f);

}

Q_DECLARE_METATYPE(QXlsx::Title)
Q_DECLARE_TYPEINFO(QXlsx::Title, Q_MOVABLE_TYPE);

#endif // XLSXTITLE_H
