// main.cpp

#include <iostream>

#include <QtGlobal>
#include <QObject>
#include <QVector>
#include <QList>

#include "recurse.hpp"

#include "xlsxdocument.h"
#include "xlsxcellrange.h"
#include "xlsxworkbook.h"
#include "xlsxabstractsheet.h"

using namespace QXlsx;

QString getHtml(QString strFilename);
bool loadXlsx(QString fileName, QString& strHtml);
QString g_htmlDoc;

int main(int argc, char *argv[])
{
    g_htmlDoc = getHtml(":/test.xlsx"); // convert from xlsx to html

    Recurse::Application app(argc, argv);

    app.use([](auto &ctx)
    {
        ctx.response.send(g_htmlDoc);
    });

    quint16 listenPort = 3001; // default port
    auto result = app.listen( listenPort );
    if ( result.error() )
    {
        qDebug() << "error upon listening:" << result.lastError();
        return (-1);
    }
    std::cout << " listening port: " << listenPort << std::endl;

    return 0;
}

QString getHtml(QString strFilename)
{
    QString ret;

    ret = ret + QString("<html>\n");

    ret = ret + QString("<head>\n");
    ret = ret + QString("<title>") + strFilename + QString("</title>\n");
    ret = ret + QString("<meta http-equiv='Content-Type' content='text/html;charset=UTF-8'>\n" );
    ret = ret + QString("</head>\n");

    ret = ret + QString("<body>\n");

    QString strTableStyle = \
        "<style>\n"\
        " table { border-collapse: collapse; } \n"\
        " td, th { border: 1px solid black; } \n"\
        "</style>\n";
    ret = ret + strTableStyle;

    if (!loadXlsx(strFilename, ret))
        return QString("");

    ret = ret + QString("</body>\n");

    ret = ret + QString("</html>\n");

    qDebug() <<  ret << "\n";

    return ret;
}

bool loadXlsx(QString fileName, QString& strHtml)
{
    // tried to load xlsx using temporary document
    QXlsx::Document xlsxTmp( fileName );
    if ( !xlsxTmp.isLoaded() )
    {
        return false; // failed to load
    }

    // load new xlsx using new document
    QXlsx::Document xlsxDoc( fileName );
    xlsxDoc.isLoaded();

    int sheetIndexNumber = 0;
    foreach( QString curretnSheetName, xlsxDoc.sheetNames() )
    {
        QXlsx::AbstractSheet* currentSheet = xlsxDoc.sheet( curretnSheetName );
        if ( NULL == currentSheet )
            continue;

        // get full cells of sheet
        currentSheet->workbook()->setActiveSheet( sheetIndexNumber );
        Worksheet* wsheet = (Worksheet*) currentSheet->workbook()->activeSheet();
        if ( NULL == wsheet )
            continue;

        QString strSheetName = wsheet->name(); // sheet name
        strHtml = strHtml + QString("<b>") + strSheetName + QString("</b><br>\n"); // UTF-8
        strHtml = strHtml + QString("<table>");

        int maxRow = wsheet->dimension().lastRow() - wsheet->dimension().firstRow();
        int maxCol = wsheet->dimension().lastColumn() - wsheet->dimension().firstColumn();

        QString strTableRecord;
        for (int rc = 0; rc < maxRow; rc++)
        {
            strTableRecord = strTableRecord + QString("<tr>");
            for (int cc = 0; cc < maxCol; cc++)
            {
                QString strTemp = wsheet->read(rc + wsheet->dimension().firstRow(),
                                               cc + wsheet->dimension().firstColumn()).toString();
                if (strTemp.isEmpty()) strTemp = "&nbsp;";
                strTableRecord = strTableRecord + QString("<td>");
                strTableRecord = strTableRecord + strTemp; // UTF-8
                strTableRecord = strTableRecord + QString("</td>");
            }
            strTableRecord = strTableRecord + QString("</tr>\n");
        }
        strHtml = strHtml + strTableRecord;

        strHtml = strHtml + QString("</table>\n");

        sheetIndexNumber++;
    }

    return true;
}
