#ifndef READXLSB_UTF
#define READXLSB_UTF

#include <string_view>
#include <string>

// minimal set of functions to replace deprecated codecvt

namespace utf {

std::string utf16le_utf8(const std::u16string_view &str);

}

#endif
