// datavalidation.cpp

#include <QtGlobal>

#include "xlsxdocument.h"
#include "xlsxdatavalidation.h"

QXLSX_USE_NAMESPACE

int datavalidation()
{
    Document xlsx;
    xlsx.write("A1", "A2 and A3:E5 only accept the number between 33 and 99");

    //![1]
    DataValidation validation(DataValidation::Type::Whole, DataValidation::Operator::Between, "33", "99");
    validation.addRange("A2");
    validation.addRange("A3:E5");
    validation.setPromptMessage("Please input integer between 33 and 99", "Validation");
    validation.setPromptMessageVisible(true);
    validation.setDropDownVisible(true);
    validation.setErrorMessage("Your input should be in the range [33, 99]", "Invalid input");
    validation.setErrorMessageVisible(true);

    xlsx.addDataValidation(validation);
    //![1]

    xlsx.saveAs("datavalidation.xlsx"); 

    return 0;
}
