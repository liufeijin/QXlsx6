// datavalidation.cpp

#include <QtGlobal>

#include "xlsxdocument.h"
#include "xlsxdatavalidation.h"

QXLSX_USE_NAMESPACE

int main()
{
    Document xlsx;
    auto sheet = xlsx.currentWorksheet();

    // Integer validation

    xlsx.write("A1", "A2 and A3:E5 only accept the number between 33 and 99");
    Format f;
    f.setPatternBackgroundColor(Qt::yellow);
    sheet->setFormat("A2", f);
    sheet->setFormat("A3:E5", f);

    // No specialized constructor, use low-level
    DataValidation validation(DataValidation::Type::Whole, "33", DataValidation::Predicate::Between,  "99");
    validation.addRange("A2");
    validation.addRange("A3:E5");
    validation.setPromptMessage("Please input integer between 33 and 99", "Range validation");
    validation.setPromptMessageVisible(true);
    validation.setDropDownVisible(true);
    validation.setErrorMessage("Your input should be in the range [33, 99]", "Invalid input");
    validation.setErrorMessageVisible(true);
    sheet->addDataValidation(validation);

    // List validation

    xlsx.write("G1", "G2:J5 only accept one of the following values: red, green, blue");
    xlsx.write("H6","red");
    xlsx.write("I6","green");
    xlsx.write("J6","blue");
    sheet->setFormat("G2:J5",f);
    sheet->addDataValidation("G2:J5","H6:J6");

    // Dynamic list validation example
    //dropdown shows values *relative* to the current cell: column G shows all
    //three values (from cells H12:J12),
    //column H shows only two values (from cells I12:K12),
    //column I shows only one value (from cells J12:L12),
    //column J shows no values.

    xlsx.write("G8", "G9:J11 only accept one of the following values: red, green, blue");
    xlsx.write("H12","red");
    xlsx.write("I12","green");
    xlsx.write("J12","blue");
    f.setPatternBackgroundColor(Qt::red);
    sheet->setFormat("G9:J11", f);
    DataValidation validation1(DataValidation::Type::List, "H$12:J$12");
    validation1.addRange("G9:J11");
    validation1.setErrorMessageVisible(true);
    sheet->addDataValidation(validation1);

    xlsx.saveAs("datavalidation.xlsx"); 

    return 0;
}
