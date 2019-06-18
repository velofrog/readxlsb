#include "BinaryReader.h"
#include "StreamReader.h"

namespace readxlsb {

File::File(Rcpp::Environment xlsb_env) {
    _env = xlsb_env;
    _raw_content = xlsb_env["stream"];
    content_length = _raw_content.size();
    content = (uint8_t *)&_raw_content[0];
    offset = 0;
    record = BinRecord(this);
}

bool File::NextRecord() {
    if (offset < content_length) return record.Next();
    return false;
}

bool File::IsEOF() {
    return (offset >= content_length);
}

BinContentRecord::BinContentRecord(BinRecord &r, uint32_t record_id) : _record_id(record_id) {
    if ((r.id == BRT_INVALID) || ((uint32_t)r.id != _record_id)) {
        valid = false;
    } else {
        valid = true;
        length = r.length;
        _content = r.ContentPtr();
        if ((_content == NULL) && (length > 0)) valid = false;
    } 
}

BinContentRecord::BinContentRecord(uint8_t *content_ptr, int content_length) : _record_id(0) {
    valid = true;
    _content = content_ptr;
    length = content_length;
    if ((_content == NULL) && (length > 0)) valid = false;
}

void RowHeaderRecord::Parse() {
    // Offset (bytes)   Length (bytes)    Name         Description
    // 0                4                 rw           Row index
    // 4                4                 ixfe         BrtXF record ID specifying style
    // 13               4                 ccolspan     Number of BrtColSpan variables
    // 17               variable          BrtColSpan   Permissible locations for cells within the row
    if (!valid) return;
    uint32_t row_index;
    StreamReader::Uint32_t(_content, length, row_index);

    rw = (int)row_index;
}

void SheetRecord::Parse() {
    // Offset (bytes)   Length (bytes)    Name         Description
    // 0                4                 hsState      ST_SheetState specifies visibility
    // 4                4                 iTabID       Sheet unique identifier
    // 8                variable          strRelID     RelID specifying relationship: XLNullableWideString
    // variable         variable          strName      Sheet name: XLWideString
    if (!valid) return;
    StreamReader::Uint32_t(_content, length, hs_state);
    StreamReader::Uint32_t(_content, length, tab_id);
    StreamReader::XLNullableWideString(_content, length, str_rel_id);
    StreamReader::XLWideString(_content, length, str_name);
}

void ExternalSheets::Parse() {
    // Offset (bytes)   Length (bytes)    Name         Description
    // 0                4                 cXti         Item count in rgXti array
    // 4                variable          rgXti        Array of Xti
    c_xti = 0;
    StreamReader::Uint32_t(_content, length, c_xti);
    if (valid && (c_xti > 0)) {
        for (uint32_t i = 0; i < c_xti; i++) {
            Xti xti;
            xti.is_internal_ref = false;
            StreamReader::Uint32_t(_content, length, xti.external_link);
            StreamReader::Int32_t(_content, length, xti.first_sheet);
            StreamReader::Int32_t(_content, length, xti.last_sheet);
            rg_xti.push_back(xti);
        }
    }
}

void DefinedNameRecord::Parse() {
    // Offset (bytes)   Length (bytes)    Name         Description
    // 0                4                 various      Flags
    //                                                   Bit 0 - fHidden. Is name visible or hidden
    //                                                   Bit 1 - fFunc. Does name represent a XLM macro
    //                                                   Bit 2 - fOB. Does name represent a VBA macro
    //                                                   Bit 3 - fProc. Does name represent a macro
    //                                                   Bit 4 - fCalcExp. Does formula return array
    //                                                   Bit 5 - fBuiltIn. Does name represent built in name
    // 4                1                 chKey        Shortcut key for macro
    // 5                4                 iTab         Name scope. 0xFFFFFFFF is the entire workbook,
    //                                                   otherwise offset into BrtBundleSh collection
    // 9                variable          name         XLNameWideString representing name
    // ...              variable          formula      NamedParsedFormula representing formula

    if (!valid) {
        valid_cell_ref = false;
        return;
    }

    uint32_t key;
    StreamReader::Uint32_t(_content, length, key);
    is_hidden = (key & 0x01) != 0;
    is_proc = (key & 0x08) != 0;
    is_built_in = (key & 0x20) != 0;
    StreamReader::Skip(_content, length, 1);
    StreamReader::Uint32_t(_content, length, tab);
    StreamReader::XLWideString(_content, length, name);
    NameParsedFormula formula(_content, length);

    // As we're only interested in defined/named ranges, we're only after absolute references
    // -> of the form SheetName!$A$1 or SheetName!$A$10:$D$24
    if (valid && formula.valid && !formula.col_first_relative && !formula.row_first_relative &&
        !formula.col_last_relative && !formula.row_last_relative) {
        ixti = formula.ixti;
        row_first = formula.row_first;
        row_last = formula.row_last;
        col_first = formula.col_first;
        col_last = formula.col_last;
        valid_cell_ref = true;
    } else {
        valid_cell_ref = false;
    }
}

bool DefinedNameRecord::IsSingleRef() {
    return ((col_first == col_last) && (row_first == row_last));
}

// NameParsedFormula only handles 3d single cell and area references
// SheetName!A1 or SheetName!A1:D10 style
void NameParsedFormula::Parse() {
    // Offset (bytes)   Length (bytes)    Name         Description
    // 0                4                 cce          Length of rgce in bytes
    // 4                variable          rgce         A sequence of Ptg structures for the formula
    uint32_t size;
    uint8_t ptg;
    uint16_t value;
    int32_t lvalue;

    if (!valid) return;
    StreamReader::Uint32_t(_content, length, size);
    StreamReader::Uint8_t(_content, length, ptg);
    
    switch (ptg) {
    case 0x3a: //PtgRef3d with PtgDataType == REFERENCE
        StreamReader::Uint16_t(_content, length, ixti);
        StreamReader::Int32_t(_content, length, lvalue);
        row_first = row_last = lvalue;
        StreamReader::Uint16_t(_content, length, value);
        col_first = col_last = (int16_t)(value & 0x3fff);
        col_first_relative = col_last_relative = ((value & 0x4000) != 0);
        row_first_relative = row_last_relative = ((value & 0x8000) != 0);
        break;
        
    case 0x3b: //PtgArea3d with PtgDataType == REFERENCE
        StreamReader::Uint16_t(_content, length, ixti);
        StreamReader::Int32_t(_content, length, row_first);
        StreamReader::Int32_t(_content, length, row_last);
        StreamReader::Uint16_t(_content, length, value);
        col_first = (int16_t)(value & 0x3fff);
        col_first_relative = ((value & 0x4000) != 0);
        row_first_relative = ((value & 0x4000) != 0);
        StreamReader::Uint16_t(_content, length, value);
        col_last = (int16_t)(value & 0x3fff);
        col_last_relative = ((value & 0x4000) != 0);
        row_last_relative = ((value & 0x4000) != 0);
        break; 
        
    default:
        // Unhandled ptg
        valid = false;
        break;
    }
}


void CellRecord::Parse() {
    StreamReader::Cell(_content, length, column, style_ref);
}

uint32_t CellRecord::CellType() {
    return _record_id;
}

void CellSharedStringRecord::Parse() {
    StreamReader::Uint32_t(_content, length, iSST_item);
}

void CellStringRecord::Parse() {
    StreamReader::XLWideString(_content, length, value);
}

void CellBoolRecord::Parse() {
    uint8_t result;
    StreamReader::Uint8_t(_content, length, result);
    result = result & 0x1;
    value = result;
}

void CellRkRecord::Parse() {
    StreamReader::RkNumber(_content, length, value);
}

void CellFormulaNumberRecord::Parse() {
    StreamReader::Double(_content, length, value);
}

void CellFormulaBoolRecord::Parse() {
    uint8_t result;
    StreamReader::Uint8_t(_content, length, result);
    result = result & 0x1;
    value = result;
}

void CellRealRecord::Parse() {
    StreamReader::Double(_content, length, value);
}

void SharedStringItem::Parse() {
    StreamReader::Skip(_content, length, 1);
    StreamReader::XLWideString(_content, length, value);
}

void NumberFormat::Parse() {
    is_datetime_format = false;
    
    StreamReader::Uint16_t(_content, length, fmt);
    if (StreamReader::XLWideString(_content, length, stfmt_code))
        is_datetime_format = StreamReader::IsDateTimeFormat(stfmt_code);
}

void CellFormat::Parse() {
    StreamReader::Uint16_t(_content, length, ixfe_parent);
    StreamReader::Uint16_t(_content, length, fmt);
}

bool CellFormat::IsBuiltInFormat() {
    // See list of standard Excel formats in StreamReader header file
    if (fmt <= 22) return true;
    if ((fmt >= 37) && (fmt <= 49)) return true;
    return false;
}

bool CellFormat::IsBuiltInDateTimeFormat() {
    if (!IsBuiltInFormat()) return false;
    return StreamReader::IsBuiltInDateTimeFormat(fmt);
}

BinRecord::BinRecord() {
    _file = NULL;
}

BinRecord::BinRecord(File *file) {
    _file = file;
    Next();
}

bool BinRecord::Next() {
    id = GetRecordID();

    length = 0;
    _content_ptr = -1;
    
    if (id == BRT_INVALID) return false;
    
    length = GetRecordLength();
    if (length > 0) _content_ptr = _file->offset;
    _file->offset += (int)length;

    return true;
}

bool BinRecord::IsCell() {
    // Is the record id a known cell type
    return ((id >= BRT_CELLBLANK) && (id <= BRT_CELLFMLAERROR));  
}

int32_t BinRecord::GetRecordID() {
    // Record identifier is at most 2 bytes. 
    // If high bit in first byte is set, then read second byte. 
    int read_byte = 0;
    uint32_t rec_id = 0;
    uint8_t b;

    // Check high bit of first byte to see if second byte required  
    while (read_byte < 2) {
        if (_file->offset >= _file->content_length) return BRT_INVALID;

        b = _file->content[_file->offset++];
        rec_id = ((uint32_t)(b & 0x7f) << (7 * read_byte)) | rec_id;
        read_byte++;
        if ((b & 0x80) == 0) break;
    }
    
    return (int32_t)rec_id;
}

uint32_t BinRecord::GetRecordLength() {
    // At most 4 bytes. As with GetRecordID, if high bit is 1 then use next byte in stream
    int read_byte = 0;
    uint32_t rec_len = 0;
    uint8_t b;
    
    // Record length is at most 4 bytes long  
    while (read_byte < 4) {
        if (_file->offset >= _file->content_length) return 0;
        
        b = _file->content[_file->offset++];
        rec_len = ((uint32_t)(b & 0x7f) << (7*read_byte)) | rec_len;
        read_byte++;
        if ((b & 0x80) == 0) break;
    }
    
    return rec_len;
}

uint8_t *BinRecord::ContentPtr() {
  if (_content_ptr < 0) return NULL;
  return &_file->content[_content_ptr];
}

} // namespace readxlsb