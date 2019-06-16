#ifndef READXLSB_BINARYREADER
#define READXLSB_BINARYREADER

#include <vector>
#include <string>
#include <Rcpp.h>
#include "StreamReader.h"

namespace readxlsb {

// Workbook record types
constexpr uint32_t BRT_BEGINBOOK = 131;
constexpr uint32_t BRT_ENDBOOK = 132;
//constexpr uint32_t BRT_FILEVERSION = 128;
//constexpr uint32_t BRT_WBPROP = 153;
//constexpr uint32_t BRT_BEGINBOOKVIEWS = 135;
//constexpr uint32_t BRT_BOOKVIEW = 158;
//constexpr uint32_t BRT_ENDBOOKVIEWS = 136;
//constexpr uint32_t BRT_BEGINSHEETS = 143;
constexpr uint32_t BRT_SHEET = 156;
constexpr uint32_t BRT_BEGINEXTERNALS = 353;
constexpr uint32_t BRT_SUPSELF = 357;
constexpr uint32_t BRT_SUPSAME = 358;
constexpr uint32_t BRT_SUPADDIN = 667;
constexpr uint32_t BRT_SUPBOOKSRC = 355;
constexpr uint32_t BRT_EXTERNSHEET = 362;
constexpr uint32_t BRT_ENDEXTERNALS = 354;
//constexpr uint32_t BRT_ENDSHEETS = 144;
constexpr uint32_t BRT_NAME = 39;
constexpr uint32_t BRT_BEGINSHEET = 129;
//constexpr uint32_t BRT_ENDSHEET = 130;
constexpr uint32_t BRT_BEGINSHEETDATA = 145;
constexpr uint32_t BRT_ENDSHEETDATA = 146;
constexpr uint32_t BRT_ROWHEADER = 0;
constexpr uint32_t BRT_BEGINSST = 159;
constexpr uint32_t BRT_ENDSST = 160;
constexpr uint32_t BRT_SSTITEM = 19;

// Processed cell types
constexpr uint32_t BRT_CELLBLANK = 1;
constexpr uint32_t BRT_CELLRK = 2;
constexpr uint32_t BRT_CELLERROR = 3;
constexpr uint32_t BRT_CELLBOOL = 4;
constexpr uint32_t BRT_CELLREAL = 5;
constexpr uint32_t BRT_CELLST = 6;
constexpr uint32_t BRT_CELLISST = 7;
constexpr uint32_t BRT_CELLFMLASTRING = 8;
constexpr uint32_t BRT_CELLFMLANUM = 9;
constexpr uint32_t BRT_CELLFMLABOOL = 10;
constexpr uint32_t BRT_CELLFMLAERROR = 11;

// Style records
constexpr uint32_t BRT_FMT = 44;
constexpr uint32_t BRT_XF = 47;
constexpr uint32_t BRT_BEGINCELLXFS = 617;
constexpr uint32_t BRT_ENDCELLXFS = 618;

// Internal constants
constexpr uint32_t BRT_INVALID = -1;

// Maximum column and row, zero-based
constexpr uint32_t MAX_COL = 0x4000-1;
constexpr uint32_t MAX_ROW = 0x100000-1;

typedef enum {TYPE_IGNORE = -2, TYPE_IMPLY = -1, 
              TYPE_BLANK = 0, TYPE_ERROR, TYPE_DOUBLE, TYPE_LOGICAL, TYPE_INTEGER, TYPE_STRING,
              TYPE_DATETIME} CONTENT_TYPE;

typedef enum {ROW = 0, COLUMN, EXCEL_TYPE, MAPPED_TYPE, FORMAT_REF, SHARED_STRING_INDEX, 
              CONTENT_OFFSET, LOGICAL_VALUE, INT_VALUE, DOUBLE_VALUE} CELLINFO;

class File;

struct CellAddress {
  uint32_t row;
  uint32_t column;
};

class BinRecord {
public:
    int32_t id;
    uint32_t length;

private:
    File *_file;
    int _content_ptr;
  
public:
    BinRecord();
    BinRecord(File *file);

    bool Next();
    bool IsCell();

    int32_t GetRecordID();
    uint32_t GetRecordLength();
    uint8_t *ContentPtr();
};

class File {
public:
    int offset;
    int content_length;
    uint8_t *content;
    BinRecord record;
  
private:
    Rcpp::Environment _env;
    Rcpp::RawVector _raw_content;

public:
    File(Rcpp::Environment xlsb_env);
    bool NextRecord();
    bool IsEOF();
};

class BinContentRecord {
public:
    bool valid;
    int length;
  
protected:
    uint8_t *_content;
    uint32_t const _record_id;
  
public:
    BinContentRecord(BinRecord &r, uint32_t record_id);
    BinContentRecord(uint8_t *content_ptr, int content_length);
    virtual ~BinContentRecord() = default;
    virtual void Parse() = 0;
};

class SheetRecord : public BinContentRecord {
public:
    uint32_t hs_state;
    uint32_t tab_id;
    std::string str_rel_id;
    std::string str_name;

public:
    SheetRecord(BinRecord &r) : BinContentRecord(r, BRT_SHEET) {Parse();}
    void Parse();
};

struct Xti {
    uint32_t external_link;
    bool is_internal_ref;
    int32_t first_sheet;
    int32_t last_sheet;
};

class ExternalSheets : public BinContentRecord {
public:
    uint32_t c_xti;
    std::vector<Xti> rg_xti;
  
public:
    ExternalSheets(BinRecord &r) : BinContentRecord(r, BRT_EXTERNSHEET) {Parse();}
    void Parse();
};

class DefinedNameRecord : public BinContentRecord {
public:
    bool is_hidden;
    bool is_proc;
    bool is_built_in;
    uint32_t tab;
    std::string name;
    bool valid_cell_ref;
    uint16_t ixti;
    int32_t row_first, row_last;
    int16_t col_first, col_last;

public:
    DefinedNameRecord(BinRecord &r) : BinContentRecord(r, BRT_NAME) {Parse();}

    void Parse();
    bool IsSingleRef();
};

class NameParsedFormula : public BinContentRecord {
public:
    uint16_t ixti;
    int32_t row_first, row_last;
    int16_t col_first, col_last;
    bool col_first_relative, col_last_relative;
    bool row_first_relative, row_last_relative;
  
    NameParsedFormula(uint8_t *content_ptr, int content_length): 
        BinContentRecord(content_ptr, content_length) {Parse();}
    void Parse();
};

class RowHeaderRecord : public BinContentRecord {
public:
    int rw;

    RowHeaderRecord(BinRecord &r) : BinContentRecord(r, BRT_ROWHEADER) {Parse();};
    void Parse();
};

class CellRecord : public BinContentRecord {
public:
    int column;
    uint32_t style_ref;

public:
    CellRecord(BinRecord &r) : BinContentRecord(r, r.id) {Parse();}
    CellRecord(BinRecord &r, uint32_t id) : BinContentRecord(r, id) {Parse();}
    void Parse();
    uint32_t CellType();
};

class CellSharedStringRecord : public CellRecord {
public:
    uint32_t iSST_item;

    CellSharedStringRecord(BinRecord &r) : CellRecord(r, BRT_CELLISST) {Parse();};
    void Parse();
};

class CellStringRecord : public CellRecord {
public:
    std::string value;

    CellStringRecord(BinRecord &r) : CellRecord(r, BRT_CELLST) {Parse();}
    CellStringRecord(BinRecord &r, uint32_t id) : CellRecord(r, id) {Parse();}
    void Parse();
};

class CellFormulaStringRecord : public CellStringRecord {
public:
    CellFormulaStringRecord(BinRecord &r) : CellStringRecord(r, BRT_CELLFMLASTRING) {Parse();}
};

class CellFormulaNumberRecord : public CellRecord {
public:
    double value;

    CellFormulaNumberRecord(BinRecord &r) : CellRecord(r, BRT_CELLFMLANUM) {Parse();}
    void Parse();
};

class CellFormulaBoolRecord : public CellRecord {
public:
    bool value;

    CellFormulaBoolRecord(BinRecord &r) : CellRecord(r, BRT_CELLFMLABOOL) {Parse();}
    void Parse();
};

class CellRkRecord : public CellRecord {
public:
    RkNumeric value;

    CellRkRecord(BinRecord &r) : CellRecord(r, BRT_CELLRK) {Parse();}
    void Parse();
};

class CellBoolRecord : public CellRecord {
public:
    bool value;

    CellBoolRecord(BinRecord &r) : CellRecord(r, BRT_CELLBOOL) {Parse();}
    void Parse();
};

class CellRealRecord : public CellRecord {
public:
    double value;

    CellRealRecord(BinRecord &r) : CellRecord(r, BRT_CELLREAL) {Parse();}
    void Parse();
};

class SharedStringItem : public BinContentRecord {
public:
    std::string value;

    SharedStringItem(BinRecord &r) : BinContentRecord(r, BRT_SSTITEM) {Parse();}
    void Parse();
};

class NumberFormat : public BinContentRecord {
public:
    uint16_t fmt;
    std::string stfmt_code;
    bool is_datetime_format;

    NumberFormat(BinRecord &r) : BinContentRecord(r, BRT_FMT) {Parse();}
    void Parse();
};

class CellFormat : public BinContentRecord {
public:
    uint16_t ixfe_parent;
    uint16_t fmt;

    CellFormat(BinRecord &r) : BinContentRecord(r, BRT_XF) {Parse();}
    void Parse();

    bool IsBuiltInFormat();
    bool IsBuiltInDateTimeFormat();
};

}

#endif