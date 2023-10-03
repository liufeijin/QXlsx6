// image.cpp

#include <QtGlobal>
#include <QtCore>
#include <QDebug>
#include <QtGui>
#include <QVector>
#include <QVariant>
#include <QDebug> 
#include <QDir>
#include <QColor>
#include <QImage>
#include <QRgb>

#if QT_VERSION >= 0x060000 // Qt 6.0 or over
#include <QRandomGenerator>
#endif

#include "xlsxdocument.h"

int image()
{
    using namespace QXlsx;

    Document xlsx;
    auto sheet = xlsx.activeWorksheet();
    if (!sheet) return -1;

    for (int i=0; i<10; ++i)
    {
        QImage image(40, 30, QImage::Format_RGB32);

#if QT_VERSION >= 0x060000 // Qt 6.0 or over
        QRandomGenerator rgen;
        image.fill( uint( rgen.generate()  % 16581375 ) );
#else
        image.fill( uint(qrand() % 16581375) );
#endif
        int index = sheet->insertImage( 10*i, 5, image );

       QImage img = sheet->image(index);
       if (!img.isNull()) {
           QString filename;
           filename = QString("image %1.png").arg( index+1 );
           img.save( filename );

            qDebug() << " [image index] " << index;
       }
    }
    //testing background images
    sheet->setBackgroundImage(":/background1.jpg");
    xlsx.saveAs("image1.xlsx");



    QXlsx::Document xlsx2("image1.xlsx");
    sheet = xlsx2.activeWorksheet();
    //testing remove images by index and by coords
    //remove 2nd image
    Q_ASSERT_X(sheet->removeImage(1), "image()", "couldn't remove image 2");
    Q_ASSERT_X(sheet->removeImage(30,5), "image()", "couldn't remove image at (30,5)");
    Q_ASSERT_X(sheet->removeBackgroundImage(), "image()", "couldn't remove background image");


    xlsx2.saveAs("image2.xlsx");

    return 0;
}
