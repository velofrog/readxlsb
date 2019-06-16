#ifndef READXLSB_DATES
#define READXLSB_DATES

#include <Rcpp.h>

namespace readxlsb {

constexpr double BASE_EXCEL_DATE = 25569; // 25569 => Excel's serial date for 1970-01-01
constexpr double BASE_JD = 2440588;       // 2440588 => JulianDate(1970, 1, 1)

class SerialDate {
public:
    static double JulianDate(int year, int month, int day, int hour = 0, int minute = 0, int second = 0);
    static double JulianDateToEpoch(double jd);
    static double ExcelToBase(double excel_date);
    static void BaseTotm(double serial, std::tm &datetime);
    static Rcpp::String BaseToString(double serial);
    static bool ParseDateTimeString(const std::string &str, double &serial);
    static bool ParseDateTimeString(const Rcpp::String &str, double &serial);
};

}

#endif