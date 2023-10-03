// main.cpp

#include <QCoreApplication>
#include <QVector>
#include <QVariant>
#include <QDir>
#include <QDebug>

#include "xlsxdocument.h"

int main()
{
    qDebug() << "[debug] current path : " << QDir::currentPath();

    using namespace QXlsx;

    qDebug() << "\n\ndoc1\n";

    Document doc;
    auto sheet = doc.activeWorksheet();

    sheet->write( "A1", QVariant(QDateTime::currentDateTimeUtc()) );
    sheet->write( "A2", QVariant(double(10.5)) );
    sheet->write( "A3", QVariant(QDate(2019, 10, 9)) );
    sheet->write( "A4", QVariant(QTime(10, 9, 5)) );
    sheet->write( "A5", QVariant((int) 40000) );

    // read() method returns values with respect to their types.
    qDebug() << "doc.read()";
    qDebug() << sheet->read( 1, 1 ).type() << sheet->read( 1, 1 );
    qDebug() << sheet->read( 2, 1 ).type() << sheet->read( 2, 1 );
    qDebug() << sheet->read( 3, 1 ).type() << sheet->read( 3, 1 );
    qDebug() << sheet->read( 4, 1 ).type() << sheet->read( 4, 1 );
    qDebug() << sheet->read( 5, 1 ).type() << sheet->read( 5, 1 );
    qDebug() << "\n";

    // cellAt()->value() method returns values as they are stored in cells (numerics usually as double).
    qDebug() << "sheet->cell()->value()";
    qDebug() << sheet->cell( 1, 1 )->value().type() << sheet->cell( 1, 1 )->value();
    qDebug() << sheet->cell( 2, 1 )->value().type() << sheet->cell( 2, 1 )->value();
    qDebug() << sheet->cell( 3, 1 )->value().type() << sheet->cell( 3, 1 )->value();
    qDebug() << sheet->cell( 4, 1 )->value().type() << sheet->cell( 4, 1 )->value();
    qDebug() << sheet->cell( 5, 1 )->value().type() << sheet->cell( 5, 1 )->value();
    qDebug() << "\n";

    doc.saveAs("types1.xlsx");

    qDebug() << "\n\ndoc2\n";

    Document doc2("types1.xlsx");
    if ( !doc2.isLoaded() )
    {
        qWarning() << "failed to load types1.xlsx" ;
        return (-1);
    }
    qDebug() << "\n\n";
    sheet = doc2.activeWorksheet();

    sheet->write( "A6", QVariant(QDateTime::currentDateTimeUtc()) );
    sheet->write( "A7", QVariant(double(10.5)) );
    sheet->write( "A8", QVariant(QDate(2019, 10, 9)) );
    sheet->write( "A9", QVariant(QTime(10, 9, 5)) );
    sheet->write( "A10", QVariant((int) 40000) );

    qDebug() << "doc2.read()";
    qDebug() << sheet->read( 1, 1 ).type() << sheet->read( 1, 1 );
    qDebug() << sheet->read( 2, 1 ).type() << sheet->read( 2, 1 );
    qDebug() << sheet->read( 3, 1 ).type() << sheet->read( 3, 1 );
    qDebug() << sheet->read( 4, 1 ).type() << sheet->read( 4, 1 );
    qDebug() << sheet->read( 5, 1 ).type() << sheet->read( 5, 1 );
    qDebug() << sheet->read( 6, 1 ).type() << sheet->read( 6, 1 );
    qDebug() << sheet->read( 7, 1 ).type() << sheet->read( 7, 1 );
    qDebug() << sheet->read( 8, 1 ).type() << sheet->read( 8, 1 );
    qDebug() << sheet->read( 9, 1 ).type() << sheet->read( 9, 1 );
    qDebug() << sheet->read(10, 1 ).type() << sheet->read(10, 1 );
    qDebug() << "\n";

    qDebug() << "sheet->cellAt()->value()";
    qDebug() << sheet->cell( 1, 1 )->value().type() << sheet->cell( 1, 1 )->value();
    qDebug() << sheet->cell( 2, 1 )->value().type() << sheet->cell( 2, 1 )->value();
    qDebug() << sheet->cell( 3, 1 )->value().type() << sheet->cell( 3, 1 )->value();
    qDebug() << sheet->cell( 4, 1 )->value().type() << sheet->cell( 4, 1 )->value();
    qDebug() << sheet->cell( 5, 1 )->value().type() << sheet->cell( 5, 1 )->value();
    qDebug() << sheet->cell( 6, 1 )->value().type() << sheet->cell( 6, 1 )->value();
    qDebug() << sheet->cell( 7, 1 )->value().type() << sheet->cell( 7, 1 )->value();
    qDebug() << sheet->cell( 8, 1 )->value().type() << sheet->cell( 8, 1 )->value();
    qDebug() << sheet->cell( 9, 1 )->value().type() << sheet->cell( 9, 1 )->value();
    qDebug() << sheet->cell(10, 1 )->value().type() << sheet->cell(10, 1 )->value();
    doc2.saveAs("types2.xlsx");

    qDebug() << "\n\ndoc3\n";

    Document doc3("types2.xlsx");
    if (!doc3.isLoaded()) {
        qWarning() << "failed to load types2.xlsx" ;
        return -1;
    }
    qDebug() << "\n\n";
    sheet = doc3.activeWorksheet();

    qDebug() << "sheet->read()";
    qDebug() << sheet->read( 1, 1 ).type() << sheet->read( 1, 1 );
    qDebug() << sheet->read( 2, 1 ).type() << sheet->read( 2, 1 );
    qDebug() << sheet->read( 3, 1 ).type() << sheet->read( 3, 1 );
    qDebug() << sheet->read( 4, 1 ).type() << sheet->read( 4, 1 );
    qDebug() << sheet->read( 5, 1 ).type() << sheet->read( 5, 1 );
    qDebug() << sheet->read( 6, 1 ).type() << sheet->read( 6, 1 );
    qDebug() << sheet->read( 7, 1 ).type() << sheet->read( 7, 1 );
    qDebug() << sheet->read( 8, 1 ).type() << sheet->read( 8, 1 );
    qDebug() << sheet->read( 9, 1 ).type() << sheet->read( 9, 1 );
    qDebug() << sheet->read(10, 1 ).type() << sheet->read(10, 1 );
    qDebug() << "\n";

    qDebug() << "sheet->cellAt()->value()";
    qDebug() << sheet->cell( 1, 1 )->value().type() << sheet->cell( 1, 1 )->value();
    qDebug() << sheet->cell( 2, 1 )->value().type() << sheet->cell( 2, 1 )->value();
    qDebug() << sheet->cell( 3, 1 )->value().type() << sheet->cell( 3, 1 )->value();
    qDebug() << sheet->cell( 4, 1 )->value().type() << sheet->cell( 4, 1 )->value();
    qDebug() << sheet->cell( 5, 1 )->value().type() << sheet->cell( 5, 1 )->value();
    qDebug() << sheet->cell( 6, 1 )->value().type() << sheet->cell( 6, 1 )->value();
    qDebug() << sheet->cell( 7, 1 )->value().type() << sheet->cell( 7, 1 )->value();
    qDebug() << sheet->cell( 8, 1 )->value().type() << sheet->cell( 8, 1 )->value();
    qDebug() << sheet->cell( 9, 1 )->value().type() << sheet->cell( 9, 1 )->value();
    qDebug() << sheet->cell(10, 1 )->value().type() << sheet->cell(10, 1 )->value();

	return 0; 
}

