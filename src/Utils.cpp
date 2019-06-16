#include "Utils.h"
#include "Dates.h"

namespace readxlsb {

// Convert column name (ie column 28) to an Excel Label (column AB)
std::string Utils::ColumnToExcelColumnLabel(int col, bool zero_based) {
    if (col < 0) return "";
    if (!zero_based) col--;

    std::string label;
    int value = col;

    while (value >= 0) {
        label.insert(0, 1, (char)('A' + (value % 26)));
        value = (value / 26) - 1;
    }

    return label;
}

double Utils::ToDouble(const Rcpp::String &str) {
    try {
        return std::stod((std::string)str);
    }
    
    catch (const std::invalid_argument&) {
        return NA_REAL;
    }
}

int Utils::ToInt(const Rcpp::String &str) {
    try {
        return std::stoi((std::string)str);
    }
    
    catch (const std::invalid_argument&) {
        return NA_INTEGER;
    }
}

double Utils::ToDateTime(const Rcpp::String &str) {
    double result;
    
    if (!SerialDate::ParseDateTimeString(str, result)) {
        return NA_REAL;
    } else {
        return result;
    }
}


} // namespace readxlsb