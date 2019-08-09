#ifndef ALTSTD_GET_TIME
#define ALTSTD_GET_TIME

#include <regex>
#include <time.h>
#include "Dates.h"

using namespace readxlsb;

namespace alt_std {

// Parse string assuming format of %Y-%m-%dT%H:%M:%S
bool get_time(std::string src, std::tm *tmb) {    
    try {
        std::regex re("(\\d+)[\\-/\\.\\s]?(\\d+)?[\\-/\\.\\s]?(\\d+)?[\\sTt]*(\\d+)?[\\-\\.:\\s]?(\\d+)?[\\-\\.:\\s]?(\\d+)?");
        std::smatch match;
        
        if (std::regex_search(src, match, re)) {
            tmb->tm_isdst = -1;
            
            std::size_t matches = match.size();
            if ((matches > 1) && (match.length(1) > 0))
                tmb->tm_year = std::stoi(match.str(1)) - 1900;                
            
            if ((matches > 2) && (match.length(2) > 0))
                tmb->tm_mon = std::stoi(match.str(2)) - 1;
            
            if ((matches > 3) && (match.length(3) > 0))
                tmb->tm_mday = std::stoi(match.str(3));
            
            if ((matches > 4) && (match.length(4) > 0))
                tmb->tm_hour = std::stoi(match.str(4));
            
            if ((matches > 5) && (match.length(5) > 0))
                tmb->tm_min = std::stoi(match.str(5));
            
            if ((matches > 6) && (match.length(6) > 0))
                tmb->tm_sec = std::stoi(match.str(6));  
            
            // Convert tmb to JulianDate and then back to handle intentional overflows
            // eg: month == 12 => January following year. day == 0 => last day of previous month, etc
            double serial = SerialDate::JulianDate(tmb->tm_year + 1900, tmb->tm_mon + 1, tmb->tm_mday,
                                                   tmb->tm_hour, tmb->tm_min, tmb->tm_sec) - BASE_JD;
            SerialDate::BaseTotm(serial, *tmb);
            
            return true;
        } else {
            return false;
        }
    }
    
    catch (const std::exception& ex) {
        return false;
    }
}

}

#endif
