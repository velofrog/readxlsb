#ifndef READXLSB_UTILS
#define READXLSB_UTILS

#include <Rcpp.h>
#include <string>

namespace readxlsb {

class Utils {
public:
    static std::string ColumnToExcelColumnLabel(int col, bool zero_based = true);
    static double ToDouble(const Rcpp::String &str);
    static int ToInt(const Rcpp::String &str);
    static double ToDateTime(const Rcpp::String &str);
};

}

#endif