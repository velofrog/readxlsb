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
    std::istringstream ss((std::string)str);
    std::tm datetime = {0};
    
    if (ss >> std::get_time(&datetime, "%Y-%m-%dT%H:%M:%S")) {
        if (datetime.tm_mday == 0) datetime.tm_mday++; // Handle strings like 2019 or 2019-01
        return SerialDate::JulianDate(datetime.tm_year + 1900, datetime.tm_mon + 1, datetime.tm_mday,
                                      datetime.tm_hour, datetime.tm_min, datetime.tm_sec) - BASE_JD;
    } else
        return NA_REAL;
}


} // namespace readxlsb