// datavalidation.cpp

#include <QtGlobal>

#include "xlsxdocument.h"
#include "xlsxdatavalidation.h"

QXLSX_USE_NAMESPACE

int datavalidation()
{
    Document xlsx;

    // Integer validation

    xlsx.write("A1", "A2 and A3:E5 only accept the number between 33 and 99");
    Format f;
    f.setPatternBackgroundColor(Qt::yellow);
    xlsx.setFormat("A2", f);
    xlsx.setFormat("A3:E5", f);

    DataValidation validation(DataValidation::Type::Whole, "33", DataValidation::Operator::Between,  "99");
    validation.addRange("A2");
    validation.addRange("A3:E5");
    validation.setPromptMessage("Please input integer between 33 and 99", "Range validation");
    validation.setPromptMessageVisible(true);
    validation.setDropDownVisible(true);
    validation.setErrorMessage("Your input should be in the range [33, 99]", "Invalid input");
    validation.setErrorMessageVisible(true);
    xlsx.addDataValidation(validation);

    // List validation

    xlsx.write("G1", "G2:J5 only accept one of the following values: red, green, blue");
    xlsx.write("H6","red");
    xlsx.write("I6","green");
    xlsx.write("J6","blue");
    xlsx.setFormat("G2:J5",f);

    xlsx.addDataValidation("G2:J5","H6:J6");

    xlsx.saveAs("datavalidation.xlsx"); 

    return 0;
}
