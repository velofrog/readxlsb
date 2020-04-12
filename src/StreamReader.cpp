#include "StreamReader.h"
#include <vector>
#if defined(__MINGW32__) || defined(__MINGW64__)
    #include <Windows.h>
#else
    #include <codecvt>
    #include <locale>
#endif
#include <algorithm>

namespace readxlsb {

bool StreamReader::IsDateTimeFormat(std::string &str) {
    size_t non_matching;
    
    // We're assuming that valid date/time formats are only constructed using the
    // characters y,m,d,h,s and the symbols '\', '-', ':', ' ' and '.' (The '\' appears in some formats preceeding '-' etc)
    non_matching = str.find_first_not_of("YMDHSymdhs\\-: .");
    return (non_matching == std::string::npos);
}

bool StreamReader::IsBuiltInDateTimeFormat(uint16_t fmt) {
    std::vector<uint16_t> datetime_formats{14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47, 48}; // See header file
    return (std::find(datetime_formats.begin(), datetime_formats.end(), fmt) != datetime_formats.end());
}

bool StreamReader::Skip(uint8_t *&content, int &max_length, int skip_bytes) {
    if (max_length < skip_bytes) return false;
    content += skip_bytes;
    max_length -= skip_bytes;
    return true;
}

bool StreamReader::Cell(uint8_t *&content, int &max_length, int32_t &column, uint32_t &style_ref) {
    int32_t signed_column;
    uint32_t _style_ref;
    
    if (!StreamReader::Int32_t(content, max_length, signed_column)) return false;
    if (!StreamReader::Uint32_t(content, max_length, _style_ref)) return false;
    // iStyleRef is first 24 bits
    style_ref = _style_ref & ((1<<24)-1);  
    column = signed_column;
    return true;
}

bool StreamReader::RkNumber(uint8_t *&content, int &max_length, RkNumeric &result) {
    uint32_t data;
    if (max_length < 4) return false;
    data = content[0] | (content[1] << 8) | (content[2] << 16) | (content[3] << 24);
    
    bool fx100 = (data & 0b1) != 0;
    bool f_int = (data & 0b10) != 0;
    data = data >> 2;

    if (f_int) {
        if (fx100) {
            result.type = RkNumeric::DOUBLE;
            result.double_value = ((double)data) / 100.0;
        } else {
            result.type = RkNumeric::INT;
            result.int_value = (int)data;
        }
    } else {
        // data holds the 30 most significant bits of a floating point number
        uint64_t value = data;
        value = value << 34;
        double *dbl = (double  *)&value;
        result.type = RkNumeric::DOUBLE;
        result.double_value = (fx100 ? *dbl/100.0 : *dbl);
    }
    
    return true;
}

bool StreamReader::Double(uint8_t *&content, int &max_length, double &result) {
    if (max_length < 8) return false;
    result = *(double *)content;
    max_length -= 8;
    content += 8;
    return true;
}

bool StreamReader::Uint8_t(uint8_t *&content, int &max_length, uint8_t &result) {
    if (max_length < 1) return false;
    result = content[0];
    max_length--;
    content++;
    return true;
}

bool StreamReader::Uint32_t(uint8_t *&content, int &max_length, uint32_t &result) {
    if (max_length < 4) return false;
    result = 
        (uint32_t)(content[3] << 24) |
        (uint32_t)(content[2] << 16) |
        (uint32_t)(content[1] << 8) |
        (uint32_t)(content[0]);
    
    max_length -= 4;
    content += 4;
    
    return true;
}

bool StreamReader::Int32_t(uint8_t *&content, int &max_length, int32_t &result) {
    uint32_t unsigned_result;
    if (StreamReader::Uint32_t(content, max_length, unsigned_result)) {
        result = (int32_t)unsigned_result;
        return true;
    } else return false;
}

bool StreamReader::Char16_t(uint8_t *&content, int &max_length, char16_t &result) {
    if (max_length < 2) return false;
    result = (char16_t)(content[1] << 8) | (char16_t)(content[0]);
    max_length -= 2;
    content += 2;
    return true;
}

bool StreamReader::Int16_t(uint8_t *&content, int &max_length, int16_t &result) {
    if (max_length < 2) return false;
    result = (uint16_t)(content[1] << 8) | (uint16_t)(content[0]);
    max_length -= 2;
    content += 2;
    return true;
}

bool StreamReader::Uint16_t(uint8_t *&content, int &max_length, uint16_t &result) {
    if (max_length < 2) return false;
    result = (uint16_t)(content[1] << 8) | (uint16_t)(content[0]);
    max_length -= 2;
    content += 2;
    return true;
}

void StreamReader::UTF16toUTF8(const std::u16string &src, std::string &dest) {
#if defined(__WIN32) || defined(__WIN64)
    // Windows version
    // Windows UNICODE conversion functions require 
    // wide string (wchar_t), which is always 2 bytes on Windows
    std::wstring wstr(src.begin(), src.end());
    std::size_t req_size = WideCharToMultiByte(CP_UTF8, 0,
        wstr.c_str(), wstr.length(), nullptr, 0, nullptr, nullptr);
    
    dest.resize(req_size, '\0');
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), wstr.length(),
        (char *)dest.data(), req_size, nullptr, nullptr);
#else
    // Linux and Mac version can use C++11 codecvt
    std::wstring_convert<std::codecvt_utf8_utf16<char16_t>, char16_t> converter;
    dest = converter.to_bytes(src);
#endif    
}
    
bool StreamReader::XLNullableWideString(uint8_t *&content, int &max_length, std::string &result) {
    uint32_t cch_characters;
    
    if (!StreamReader::Uint32_t(content, max_length, cch_characters)) return false;

    if (cch_characters == 0xFFFFFFFF) {
        // identifies nullable string.
        // Let's return empty string for now
        result.clear();
        return true;
    }
    
    std::u16string source;
    char16_t c;
    
    for (uint32_t len=0; len<cch_characters; len++) {
        if (!StreamReader::Char16_t(content, max_length, c)) return false; // wide character set
        source += c;
    }
    
    StreamReader::UTF16toUTF8(source, result);
    return true;
}

bool StreamReader::XLWideString(uint8_t *&content, int &max_length, std::string &result) {
    uint32_t cch_characters;
    
    if (!StreamReader::Uint32_t(content, max_length, cch_characters)) return false;
    
    if (cch_characters == 0) {
        result.clear();
        return true;
    }
    
    std::u16string source;
    char16_t c;
    
    for (uint32_t len=0; len<cch_characters; len++) {
        if (!StreamReader::Char16_t(content, max_length, c)) return false;
        source += c;
    }

    StreamReader::UTF16toUTF8(source, result);
    return true;
}

} // namespace readxlsb
