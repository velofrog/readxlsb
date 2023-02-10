#include <string_view>
#include <cstddef>
#include <string>
#include <exception>
#include <utility>
#include "UTF.h"

namespace utf {

enum endianness { big, little };

// Is R supported on any big endian processors?
// This may very well be redundant
endianness machine_endianness() {
  const uint16_t test_value = 0x0100;
  if (*reinterpret_cast<const uint8_t *>(&test_value)) {
    return endianness::big;
  } else {
    return endianness::little;
  }
}

char16_t byteswap(char16_t c) {
  return ((c & 0xFF) << 8) | (c >> 8);
}

char16_t utf16le_peek(std::u16string_view::iterator &iter, std::u16string_view::iterator end, endianness endian = endianness::little) {
  if (iter < end) {
    char16_t c = *iter;
    if (endian == endianness::big) c = byteswap(c);
    return c;
  } else {
    throw std::out_of_range("Attempt to read beyond length of string");
  }
}

char16_t utf16le_read(std::u16string_view::iterator &iter, std::u16string_view::iterator end, endianness endian = endianness::little) {
  char16_t c = utf16le_peek(iter, end, endian);
  iter++;
  return c;
}

std::string utf16le_utf8(const std::u16string_view &str) {
  std::string result;
  
  std::u16string_view::iterator it = str.begin();
  std::u16string_view::iterator end = str.end();
  endianness endian = machine_endianness();
  
  // check for byte-order-mark (FEFF).
  if (str.length() > 1) {
    if (utf16le_peek(it, end, endian) == 0xFEFF) {
      it++;
    }
  }
  
  while (it < end) {
    uint32_t cp = 0;
    char16_t c = utf16le_read(it, end, endian);
    
    if ((c < 0xD800) || (c > 0xDFFF)) {
      // single utf-16 character
      cp = c;
    } else {
      // https://en.wikipedia.org/wiki/UTF-16
      // U' = yyyyyyyyyyxxxxxxxxxx  // U - 0x10000
      // W1 = 110110yyyyyyyyyy      // 0xD800 + yyyyyyyyyy
      // W2 = 110111xxxxxxxxxx      // 0xDC00 + xxxxxxxxxx
      if (it < end) {
        uint32_t high_surrogate = c;
        uint32_t low_surrogage = utf16le_read(it, end, endian);
        if (endian == endianness::big) std::swap(high_surrogate, low_surrogage);
        
        if ((high_surrogate & 0xD800) == 0xD800 && (low_surrogage & 0xDC00) == 0xDC00) {
          high_surrogate &= 0b1111111111;
          low_surrogage &= 0b1111111111;
          cp = (high_surrogate << 10) + low_surrogage + 0x10000;
        } else {
          // invalid: surrogate pairs not encoded correctly
          cp = 0xfffd;  // question mark in diamond
        }
      } else {
        // invalid: string not long enough
        cp = 0xfffd;   // question mark in diamond
      }
    }
    
    // https://en.wikipedia.org/wiki/UTF-8
    if (cp < 0x80) {
      result.push_back((char)cp);
    } else if (cp < 0x800) {
      result.push_back((char)(0b11000000 + (cp >> 6)));
      result.push_back((char)(0b10000000 + (cp & 0b111111)));
    } else if (cp < 0x10000) {
      result.push_back((char)(0b11100000 + (cp >> 12)));
      result.push_back((char)(0b10000000 + ((cp >> 6) & 0b111111)));
      result.push_back((char)(0b10000000 + (cp & 0b111111)));
    } else if (cp < 0x110000) {
      result.push_back((char)(0b11110000 + (cp >> 18)));
      result.push_back((char)(0b10000000 + ((cp >> 12) & 0b111111)));
      result.push_back((char)(0b10000000 + ((cp >> 6) & 0b111111)));
      result.push_back((char)(0b10000000 + (cp & 0b111111)));
    } else {
      // invalid codepoint
    }
  }
  
  return result;
}

} // utf
