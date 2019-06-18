#include "Dates.h"
#include <string>
#if defined(__MINGW32__) || defined(__MINGW64__) // R's Windows compiler does not have std::get_time (yet). We'll have to wait for GCC 5
#include "get_time.h"
#endif

namespace readxlsb {

double SerialDate::ExcelToBase(double excel_date) {
    return (excel_date - BASE_EXCEL_DATE);
}

double SerialDate::JulianDate(int year, int month, int day, int hour, int minute, int second) {
    // Should we be adding 0.5 to the result? Possibly, but shouldn't make a difference 
    // for how we intend to use this
    return (1461 * (year + 4800 + (month - 14) / 12)) / 4 +
           (367 * (month - 2 - 12 * ((month - 14) / 12))) / 12 - 
           (3 * ((year + 4900 + (month - 14) / 12) / 100)) / 4 + day - 32075 +
           (double)hour / 24.0 + (double)minute / 24.0 / 60.0 + (double)second / 24.0 / 60.0 / 60.0;
}

double SerialDate::JulianDateToEpoch(double jd) {
    return (jd - BASE_JD) * 24.0 * 60.0 * 60.0;
}

void SerialDate::BaseTotm(double serial, std::tm &datetime) {
    datetime = {0};

    if (std::isnan(serial)) serial = 0;

    // Straight out of wikipedia
    // https://en.wikipedia.org/wiki/Julian_day#Julian_or_Gregorian_calendar_from_Julian_day_number
    int jd = (int)serial - BASE_EXCEL_DATE + BASE_JD;
    int f = jd + 1401 + (((4 * jd + 274277) / 146097) * 3) / 4 - 38;
    int e = 4 * f + 3;
    int g = (e % 1461) / 4;
    int h = (5 * g) + 2;

    datetime.tm_mday = (h % 153) / 5 + 1;
    datetime.tm_mon = ((h / 153) + 2) % 12;
    datetime.tm_year = (e / 1461) - 4716 + (14 - (datetime.tm_mon + 1)) / 12 - 1900;

    double time = std::modf(serial, &serial);
    datetime.tm_hour = time * 24;
    datetime.tm_min = (time * 24 - datetime.tm_hour) * 60;
    datetime.tm_sec = ((time * 24 - datetime.tm_hour) * 60 - datetime.tm_min) * 60;
}

Rcpp::String SerialDate::BaseToString(double serial) {
    std::tm datetime;
    SerialDate::BaseTotm(serial, datetime);

    std::ostringstream ss;
    ss << std::setfill('0') << std::setw(4) << (datetime.tm_year + 1900) << "-"
        << std::setw(2) << (datetime.tm_mon + 1) << "-" << std::setw(2) << datetime.tm_mday;

    if (!((datetime.tm_hour == 0) && (datetime.tm_min == 0) && (datetime.tm_sec == 0)))
    ss << " " << std::setw(2) << datetime.tm_hour << ":" 
        << std::setw(2) << datetime.tm_min << ":" << std::setw(2) << datetime.tm_sec;

    return Rcpp::String(ss.str());
}

bool SerialDate::ParseDateTimeString(const std::string &str, double &serial) {
    std::istringstream ss((std::string)str);
    std::tm datetime = {0};
    bool success;
    
    #if defined(__MINGW32__) || defined(__MINGW64__)
    success = (bool)(ss >> std_backport::get_time(&datetime, "%Y-%m-%dT%H:%M:%S"));
    #else
    success = (bool)(ss >> std::get_time(&datetime, "%Y-%m-%dT%H:%M:%S"));
    #endif  
    
    if (success) {
        if (datetime.tm_mday == 0) datetime.tm_mday++; // Handle strings like 2019 or 2019-01
        serial = SerialDate::JulianDate(datetime.tm_year + 1900, datetime.tm_mon + 1, datetime.tm_mday,
                                    datetime.tm_hour, datetime.tm_min, datetime.tm_sec) - BASE_JD;
        return true;
    } else {
        serial = 0;
        return false;
    }
} 

bool SerialDate::ParseDateTimeString(const Rcpp::String &str, double &serial) {
    return SerialDate::ParseDateTimeString((std::string)str, serial);
}

} // namespace readxlsb
