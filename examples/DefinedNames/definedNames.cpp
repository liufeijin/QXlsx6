// definedNames.cpp

#include <QtGlobal>
#include <QtCore>
#include <QDebug>

#include "xlsxdocument.h"
#include "xlsxformat.h"
#include "xlsxworkbook.h"


QXLSX_USE_NAMESPACE

int main()
{
    Document xlsx;
    for (int i=1; i<=10; ++i) {
        xlsx.write(i, 1, i);
        xlsx.write(i, 2, i*i);
        xlsx.write(i, 3, i*i*i);
    }

    xlsx.addDefinedName("MyCol_1", "=Sheet1!$A$1:$A$10");
    xlsx.addDefinedName("MyCol_2", "=Sheet1!$B$1:$B$10");
    auto definedName = xlsx.definedName("MyCol_2");
    if (definedName)
        definedName->comment = "This is a comment";
    xlsx.addDefinedName("MyCol_3", "=Sheet1!$C$1:$C$10", "Sheet1");
    definedName = xlsx.definedName("MyCol_3");
    if (definedName)
        definedName->description = "This is a description";
    xlsx.addDefinedName("Factor", "=0.5");
    auto factor = xlsx.definedName("Factor");
    if (factor) {
        factor->hidden = true; //hide this name from the list of defined names.
        factor->workbookParameter = true;
    }
    //test adding duplicate defined name
    Q_ASSERT_X(!xlsx.addDefinedName("MyCol_1", "=Sheet1!$A$1:$A$20"), "", "MyCol_1 is already defined");
    qDebug() << "Defined names:" << xlsx.definedNames();

    xlsx.write(11, 1, "=SUM(MyCol_1)");
    xlsx.write(11, 2, "=SUM(MyCol_2)");
    xlsx.write(11, 3, "=SUM(MyCol_3)");
    xlsx.write(12, 1, "=SUM(MyCol_1)*Factor");
    xlsx.write(12, 2, "=SUM(MyCol_2)*Factor");
    xlsx.write(12, 3, "=SUM(MyCol_3)*Factor");

    xlsx.saveAs("DefinedNames1.xlsx");

    Document xlsx2("DefinedNames1.xlsx");
    if ( xlsx2.isLoaded() )
    {
        xlsx2.saveAs("DefinedNames2.xlsx");
    }

    qDebug()<<"**** end of main() ****";
    return 0;
}
