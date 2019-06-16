#include <string>

namespace readxlsb {

/*
 * Documented built-in number formats (BrtXF:iFmt)
 * Date formats are localised.
 * 
 * Index	Index	Format String
 * 0	0x00	General
 * 1	0x01	0
 * 2	0x02	0.00
 * 3	0x03	#,##0
 * 4	0x04	#,##0.00
 * 5	0x05	($#,##0_);($#,##0)
 * 6	0x06	($#,##0_);[Red]($#,##0)
 * 7	0x07	($#,##0.00_);($#,##0.00)
 * 8	0x08	($#,##0.00_);[Red]($#,##0.00)
 * 9	0x09	0%
 * 10	0x0a	0.00%
 * 11	0x0b	0.00E+00
 * 12	0x0c	# ?/?
 * 13	0x0d	# ??/??
 * 14	0x0e	m/d/yy
 * 15	0x0f	d-mmm-yy
 * 16	0x10	d-mmm
 * 17	0x11	mmm-yy
 * 18	0x12	h:mm AM/PM
 * 19	0x13	h:mm:ss AM/PM
 * 20	0x14	h:mm
 * 21	0x15	h:mm:ss
 * 22	0x16	m/d/yy h:mm
 * …	…	…   Undocumented
 * 37	0x25	(#,##0_);(#,##0)
 * 38	0x26	(#,##0_);[Red](#,##0)
 * 39	0x27	(#,##0.00_);(#,##0.00)
 * 40	0x28	(#,##0.00_);[Red](#,##0.00)
 * 41	0x29	_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)
 * 42	0x2a	_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)
 * 43	0x2b	_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)
 * 44	0x2c	_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)
 * 45	0x2d	mm:ss
 * 46	0x2e	[h]:mm:ss
 * 47	0x2f	mm:ss.0
 * 48	0x30	##0.0E+0
 * 49	0x31	@
 * 
 */

struct RkNumeric {
  enum {INT, DOUBLE} type;
  union {
    int int_value;
    double double_value;
  };
};

class StreamReader {
public:
    static bool Skip(uint8_t *&content, int &max_length, int skip_bytes);
    static bool Uint8_t(uint8_t *&content, int &max_length, uint8_t &result);
    static bool Char16_t(uint8_t *&content, int &max_length, char16_t &result);
    static bool Int16_t(uint8_t *&content, int &max_length, int16_t &result);
    static bool Uint16_t(uint8_t *&content, int &max_length, uint16_t &result);
    static bool Double(uint8_t *&content, int &max_length, double &result);
    static bool Uint32_t(uint8_t *&content, int &max_length, uint32_t &result);
    static bool Int32_t(uint8_t *&content, int &max_length, int32_t &result);
    static bool Cell(uint8_t *&content, int &max_length, int32_t &column, uint32_t &style_ref);
    static bool XLNullableWideString(uint8_t *&content, int &max_length, std::string &result);
    static bool XLWideString(uint8_t *&content, int &max_length, std::string &result);
    static bool RkNumber(uint8_t *&content, int &max_length, RkNumeric &result);
    static bool IsDateTimeFormat(std::string &str);
    static bool IsBuiltInDateTimeFormat(uint16_t fmt);
};

}