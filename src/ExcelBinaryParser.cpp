#include <Rcpp.h>
#include "BinaryReader.h"
#include "Dates.h"
#include "Utils.h"
#include <vector>
#include <algorithm>
#include <regex>

using Rcpp::as;

namespace readxlsb {

// Helper templates
template<int RTYPE, typename T> void find_all(Rcpp::Vector<RTYPE> &src, T match, std::vector<int> &indices) {
    indices.clear();
    
    typename Rcpp::Vector<RTYPE>::iterator it;
    
    for (it = std::find(src.begin(), src.end(), match); it != src.end(); it = std::find(it + 1, src.end(), match)) {
        indices.push_back(std::distance(src.begin(), it));
    }
}

template<int RTYPE, typename T> bool find_any(Rcpp::Vector<RTYPE> &src, T match) {
    return (std::find(src.begin(), src.end(), match) != src.end());
}

// Given our vectors (row / column/ mapped.type / etc) are all the same length, we can
// do away with out-of-bounds checks 
template<int RTYPE, typename T> bool find_any(Rcpp::Vector<RTYPE> &src, std::vector<int> &indices, T match) {
    for (int idx : indices) {
        if (src[idx] == match) return true;
    }  
    
    return false;
}

// Case insensitive string comparison
bool iequals(const std::string &str1, const std::string &str2) {
    return ((str1.size() == str2.size()) && 
            std::equal(str1.begin(), str1.end(), str2.begin(), 
                        [](const char &c1, const char &c2) { 
                        return (c1 == c2 || std::toupper(c1) == std::toupper(c2)); }));
}

bool iequals(const Rcpp::String &str1, const std::string &str2) {
    return iequals((std::string)str1, str2);
}

// [[Rcpp::export]]
void ParseWorkbook(Rcpp::Environment xlsb_env) {
    File workbook(xlsb_env);
    
    if (workbook.record.id != BRT_BEGINBOOK) {
        Rcpp::Rcerr << "Expecting to find BrtBeginBook token" << std::endl;
        return;
    }
    
    Rcpp::NumericVector sheet_tabIDs;
    Rcpp::CharacterVector sheet_relIDs;
    Rcpp::CharacterVector sheet_names;
    Rcpp::NumericVector supporting_link_records;
    std::vector<Xti> extsheet_refs;
    std::vector<DefinedNameRecord> defined_names;

    while (workbook.record.id != BRT_ENDBOOK) {
        if (workbook.record.id == BRT_BEGINEXTERNALS) {
            while (workbook.record.id != BRT_ENDEXTERNALS) {
                if (!workbook.NextRecord()) {
                    Rcpp::Rcerr << "Encountered unexpected end in workbook" << std::endl;
                    break;
                }
                
                if ((workbook.record.id == BRT_SUPSELF) || (workbook.record.id == BRT_SUPSAME) ||
                    (workbook.record.id == BRT_SUPADDIN) || (workbook.record.id == BRT_SUPBOOKSRC)) {
                    supporting_link_records.push_back(workbook.record.id);
                } else if (workbook.record.id == BRT_EXTERNSHEET) {
                    ExternalSheets externsheets(workbook.record);
                    for (int i=0; i<externsheets.c_xti; i++) {
                        Xti x;
                        x = externsheets.rg_xti[i];
                        x.is_internal_ref = ((x.external_link < supporting_link_records.size()) &&
                                            (supporting_link_records[x.external_link] == BRT_SUPSELF));
                        extsheet_refs.push_back(x);
                    }
                }
            }
        } else if (workbook.record.id == BRT_NAME) {
            DefinedNameRecord name(workbook.record);
            if (name.valid_cell_ref) defined_names.push_back(name);
        } else if (workbook.record.id == BRT_SHEET) {
            SheetRecord sheet(workbook.record);
            if (sheet.valid) {
                sheet_tabIDs.push_back(sheet.tab_id);
                sheet_relIDs.push_back(sheet.str_rel_id);
                sheet_names.push_back(sheet.str_name);
            }
        } else { // Something else
            // Encountered some other token
        }
        if (!workbook.NextRecord()) break;
    }
    
    Rcpp::DataFrame result = Rcpp::DataFrame::create(
        Rcpp::Named("Id") = sheet_relIDs,
        Rcpp::Named("Name") = sheet_names,
        Rcpp::Named("stringsAsFactors") = false);
    xlsb_env.assign("sheets", result);
    
    // named_ranges: name | range | sheet_idx | first_column | first_row | last_column | last_row
    Rcpp::List named_ranges(7);
    named_ranges[0] = Rcpp::CharacterVector(defined_names.size());
    named_ranges[1] = Rcpp::CharacterVector(defined_names.size());
    named_ranges[2] = Rcpp::IntegerVector(defined_names.size());
    named_ranges[3] = Rcpp::IntegerVector(defined_names.size());
    named_ranges[4] = Rcpp::IntegerVector(defined_names.size());
    named_ranges[5] = Rcpp::IntegerVector(defined_names.size());
    named_ranges[6] = Rcpp::IntegerVector(defined_names.size());
    for (int i=0; i<defined_names.size(); i++) {
        as<Rcpp::CharacterVector>(named_ranges[0])[i] = defined_names[i].name;
        as<Rcpp::IntegerVector>(named_ranges[2])[i] = extsheet_refs[defined_names[i].ixti].first_sheet;
        as<Rcpp::IntegerVector>(named_ranges[3])[i] = defined_names[i].col_first + 1;
        as<Rcpp::IntegerVector>(named_ranges[4])[i] = defined_names[i].row_first + 1;
        as<Rcpp::IntegerVector>(named_ranges[5])[i] = defined_names[i].col_last + 1;
        as<Rcpp::IntegerVector>(named_ranges[6])[i] = defined_names[i].row_last + 1;
        std::string display_name;
        
        if ((defined_names[i].ixti < extsheet_refs.size()) && 
            (extsheet_refs[defined_names[i].ixti].is_internal_ref)) {
            if (extsheet_refs[defined_names[i].ixti].first_sheet < 0) {
                display_name = "#REF";
            } else {
                // Excel escapes quotes (') in the sheet name with two single quotes ('')
                std::string sheet_name = std::regex_replace(
                    as<std::string>(sheet_names[extsheet_refs[defined_names[i].ixti].first_sheet]), 
                    std::regex("'"), "''");
                
                if (sheet_name.find_first_of(" \t") == std::string::npos) {
                    display_name = sheet_name + "!";
                } else {
                    display_name = "'" + sheet_name + "'!";
                }

                std::string col_first_str = Utils::ColumnToExcelColumnLabel(defined_names[i].col_first);
                std::string col_last_str = Utils::ColumnToExcelColumnLabel(defined_names[i].col_last);
                if (defined_names[i].IsSingleRef()) {
                    display_name = display_name + "$" + col_first_str + "$" + std::to_string(defined_names[i].row_first+1);
                } else if ((defined_names[i].col_first == 0) && (defined_names[i].col_last >= MAX_COL)) {
                    display_name = display_name + "$" + std::to_string(defined_names[i].row_first+1) + 
                        ":$" + std::to_string(defined_names[i].row_last+1);
                } else if ((defined_names[i].row_first == 0) && (defined_names[i].row_last >= MAX_ROW)) {
                    display_name = display_name + "$" + col_first_str + ":$" + col_last_str;
                } else {
                    display_name = display_name + "$" + col_first_str + "$" + std::to_string(defined_names[i].row_first+1) +
                        ":$" + col_last_str + "$" + std::to_string(defined_names[i].row_last+1);
                }
            }
        } else {
            display_name = "#EXTERN";
        }
        as<Rcpp::CharacterVector>(named_ranges[1])[i] = display_name;
    }

    named_ranges.attr("names") = Rcpp::CharacterVector::create("name", "range", "sheet_idx", "first_column", 
        "first_row", "last_column", "last_row");
    named_ranges.attr("class") = "data.frame";
    named_ranges.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, -1 * defined_names.size());
    
    xlsb_env.assign("named_ranges", named_ranges);
}

// [[Rcpp::export]]
void ParseWorksheet(Rcpp::Environment xlsb_env, int start_row = 0, int end_row = 0, int start_col = 0, int end_col = 0) {
    File worksheet(xlsb_env);

    if (worksheet.record.id != BRT_BEGINSHEET) {
        Rcpp::Rcerr << "Expecting to find BrtBeginSheet token at start" << std::endl;
        return;
    }
    
    while (worksheet.record.id != BRT_BEGINSHEETDATA) {
        if (!worksheet.NextRecord()) {
            Rcpp::Rcerr << "Worksheet ends unexpectedly (no BrtBeginSheetData token found)" << std::endl;
            return;
        }
    }
    
    int saved_offset = worksheet.offset;
    int count_cells = 0;
    int active_row = 0;
    
    // change column and row indicies that are NA, to zeros
    if (start_row == NA_INTEGER) start_row = 0;
    if (end_row == NA_INTEGER) end_row = 0;
    if (start_col == NA_INTEGER) start_col = 0;
    if (end_col == NA_INTEGER) end_col = 0;
    
    // make zero-based
    start_row--;
    end_row--;
    start_col--;
    end_col--;
    
    // count cells in range
    while (worksheet.record.id != BRT_ENDSHEETDATA) {
        if (!worksheet.NextRecord()) {
            Rcpp::Rcerr << "Worksheet ends unexpectedly (no BrtEndSheetData token found)" << std::endl;
            return;
        }
        
        if (worksheet.record.id == BRT_ROWHEADER) {
            RowHeaderRecord row(worksheet.record);
            active_row = row.rw;
        } else if (worksheet.record.IsCell()) {
            CellRecord cell(worksheet.record);
            // check within range
            if ((active_row >= start_row) && ((end_row < 0) || (active_row <= end_row)) &&
                (cell.column >= start_col) && ((end_col < 0) || (cell.column <= end_col))) {
                count_cells++;
            }
        }
    }
    
    if (count_cells == 0) {
        xlsb_env["content"] = Rcpp::DataFrame();
        return; // no matching cells      
    }
    
    // [row | column | excel type | mapped type | style ref | shared string index | content offset]
    Rcpp::IntegerMatrix cellinfo = Rcpp::IntegerMatrix(count_cells, 7); 
    colnames(cellinfo) = Rcpp::CharacterVector::create("row", "column", "excel.type", "mapped.type", 
            "style.ref", "shared.string.index", "content.offset");
    // Store strings
    Rcpp::CharacterVector cellstrings = Rcpp::CharacterVector(count_cells, NA_STRING);
    Rcpp::IntegerVector cellbools = Rcpp::IntegerVector(count_cells, NA_INTEGER);
    Rcpp::IntegerVector cellints = Rcpp::IntegerVector(count_cells, NA_INTEGER);
    Rcpp::NumericVector celldoubles = Rcpp::NumericVector(count_cells, NA_REAL);
    worksheet.offset = saved_offset; // go back and populate matrix
    worksheet.record.id = 0xFFFF;
    int matrix_info_row = 0;

    while (worksheet.record.id != BRT_ENDSHEETDATA) {
        int cur_offset = worksheet.offset;
        if (!worksheet.NextRecord()) {
            Rcpp::Rcerr << "Worksheet ends unexpectedly (no BrtEndSheetData token found)" << std::endl;
            return;
        }

        if (worksheet.record.id == BRT_ROWHEADER) {
            RowHeaderRecord row(worksheet.record);
            active_row = row.rw;
        } else if (worksheet.record.IsCell()) {
            CellRecord cell(worksheet.record);
            // check within range
            if ((active_row >= start_row) && ((end_row < 0) || (active_row <= end_row)) &&
                (cell.column >= start_col) && ((end_col < 0) || (cell.column <= end_col))) {
                
                cellinfo(matrix_info_row, CELLINFO::ROW) = active_row;
                cellinfo(matrix_info_row, CELLINFO::COLUMN) = cell.column;
                cellinfo(matrix_info_row, CELLINFO::EXCEL_TYPE) = cell.CellType();
                cellinfo(matrix_info_row, CELLINFO::MAPPED_TYPE) = -1;
                cellinfo(matrix_info_row, CELLINFO::FORMAT_REF) = cell.style_ref;
                cellinfo(matrix_info_row, CELLINFO::SHARED_STRING_INDEX) = NA_INTEGER;
                cellinfo(matrix_info_row, CELLINFO::CONTENT_OFFSET) = cur_offset;
                
                worksheet.offset = cur_offset;
                worksheet.NextRecord(); // Refetch
                
                if (cell.CellType() == BRT_CELLISST) {
                    CellSharedStringRecord cell(worksheet.record);
                    cellinfo(matrix_info_row, CELLINFO::SHARED_STRING_INDEX) = cell.iSST_item;
                    cellinfo(matrix_info_row, CELLINFO::MAPPED_TYPE) = CONTENT_TYPE::TYPE_STRING;
                } else if (cell.CellType() == BRT_CELLST) {
                    CellStringRecord cell(worksheet.record);
                    cellstrings(matrix_info_row) = cell.value;
                    cellinfo(matrix_info_row, CELLINFO::MAPPED_TYPE) = CONTENT_TYPE::TYPE_STRING;
                } else if (cell.CellType() == BRT_CELLFMLASTRING) {
                    CellFormulaStringRecord cell(worksheet.record);
                    cellstrings(matrix_info_row) = cell.value;
                    cellinfo(matrix_info_row, CELLINFO::MAPPED_TYPE) = CONTENT_TYPE::TYPE_STRING;
                } else if (cell.CellType() == BRT_CELLBOOL) {
                    CellBoolRecord cell(worksheet.record);
                    cellbools(matrix_info_row) = cell.value;
                    cellinfo(matrix_info_row, CELLINFO::MAPPED_TYPE) = CONTENT_TYPE::TYPE_LOGICAL;
                } else if (cell.CellType() == BRT_CELLRK) {
                    CellRkRecord cell(worksheet.record);
                    if (cell.value.type == RkNumeric::INT) {
                        cellints(matrix_info_row) = cell.value.int_value;
                        cellinfo(matrix_info_row, CELLINFO::MAPPED_TYPE) = CONTENT_TYPE::TYPE_INTEGER;
                    } else if (cell.value.type == RkNumeric::DOUBLE) {
                        celldoubles(matrix_info_row) = cell.value.double_value;
                        cellinfo(matrix_info_row, CELLINFO::MAPPED_TYPE) = CONTENT_TYPE::TYPE_DOUBLE;
                    }
                } else if (cell.CellType() == BRT_CELLFMLANUM) {
                    CellFormulaNumberRecord cell(worksheet.record);
                    if (cell.valid) celldoubles(matrix_info_row) = cell.value;
                    cellinfo(matrix_info_row, CELLINFO::MAPPED_TYPE) = CONTENT_TYPE::TYPE_DOUBLE;
                } else if (cell.CellType() == BRT_CELLFMLABOOL) {
                    CellFormulaBoolRecord cell(worksheet.record);
                    if (cell.valid) cellbools(matrix_info_row) = cell.value;
                    cellinfo(matrix_info_row, CELLINFO::MAPPED_TYPE) = CONTENT_TYPE::TYPE_LOGICAL;
                } else if (cell.CellType() == BRT_CELLBLANK) {
                   cellinfo(matrix_info_row, CELLINFO::MAPPED_TYPE) = CONTENT_TYPE::TYPE_BLANK;
                } else if (cell.CellType() == BRT_CELLERROR) {
                    cellinfo(matrix_info_row, CELLINFO::MAPPED_TYPE) = CONTENT_TYPE::TYPE_ERROR;
                } else if (cell.CellType() == BRT_CELLREAL) {
                    CellRealRecord cell(worksheet.record);
                if (cell.valid) celldoubles(matrix_info_row) = cell.value;
                    cellinfo(matrix_info_row, CELLINFO::MAPPED_TYPE) = CONTENT_TYPE::TYPE_DOUBLE;
                }
                
                matrix_info_row++;
            }
        }
    }

    Rcpp::DataFrame result = Rcpp::DataFrame::create(
        Rcpp::Named("row") = cellinfo.column(CELLINFO::ROW) + 1,
        Rcpp::Named("column") = cellinfo.column(CELLINFO::COLUMN) + 1,
        Rcpp::Named("excel.type") = cellinfo.column(CELLINFO::EXCEL_TYPE),
        Rcpp::Named("mapped.type") = cellinfo.column(CELLINFO::MAPPED_TYPE),
        Rcpp::Named("format.ref") = cellinfo.column(CELLINFO::FORMAT_REF),
        Rcpp::Named("shared.string.index") = cellinfo.column(CELLINFO::SHARED_STRING_INDEX),
        Rcpp::Named("content.offset") = cellinfo.column(CELLINFO::CONTENT_OFFSET),
        Rcpp::Named("logical.value") = cellbools,
        Rcpp::Named("int.value") = cellints,
        Rcpp::Named("double.value") = celldoubles,
        Rcpp::Named("str.value") = cellstrings,
        Rcpp::Named("stringsAsFactors") = false);

    xlsb_env.assign("content", result);
}

// [[Rcpp::export]]
void ParseSharedStrings(Rcpp::Environment xlsb_env) {
    File sharedstrings(xlsb_env);
    Rcpp::DataFrame content = as<Rcpp::DataFrame>(xlsb_env["content"]);
    if (content.nrow() == 0) return;
    
    Rcpp::CharacterVector str_values = content["str.value"];
    Rcpp::IntegerVector shared_string_index = content["shared.string.index"];

    while (sharedstrings.record.id != BRT_ENDSST) {
        if (sharedstrings.record.id == BRT_BEGINSST) {
            int iSST_item = 0;
            while (sharedstrings.record.id != BRT_ENDSST) {
                sharedstrings.NextRecord();
                if (sharedstrings.record.id == BRT_SSTITEM) {
                    SharedStringItem ssi(sharedstrings.record);
                    Rcpp::IntegerVector::iterator it = std::find(shared_string_index.begin(), 
                        shared_string_index.end(), iSST_item);
                    while (it != shared_string_index.end()) {
                        int index = std::distance(shared_string_index.begin(), it);
                        str_values[index] = ssi.value;
                        it = std::find(it + 1, shared_string_index.end(), iSST_item);
                    }
                    iSST_item++;
                }
            }
            break;
        }
        
        if (!sharedstrings.NextRecord()) break;
    }
}

// [[Rcpp::export]]
void ParseStyles(Rcpp::Environment xlsb_env) {
    File styles(xlsb_env);
    
    Rcpp::DataFrame content = as<Rcpp::DataFrame>(xlsb_env["content"]);
    if (content.nrow() == 0) return;
    
    Rcpp::IntegerVector format_ref = content["format.ref"];
    Rcpp::IntegerVector mapped_type = content["mapped.type"];

    if (format_ref.size() == 0) {
        Rcpp::Rcerr << "style.ref missing from content dataframe" << std::endl;
        return;
    }
    
    if (mapped_type.size() == 0) {
        Rcpp::Rcerr << "mapped.type missing from content dataframe" << std::endl;
        return;
    }
    
    // First pass. Build table of custom formats
    std::map<uint16_t, NumberFormat> custom_formats;
    styles.offset = 0;
    while (styles.NextRecord()) {
        if (styles.record.id == BRT_FMT) {
            NumberFormat fmt(styles.record);
            custom_formats.insert(std::pair<uint16_t, NumberFormat>(fmt.fmt, fmt));
        }
    }
    
    // Second pass. Traverse cell formats, identifying date/time formatting
    styles.offset = 0;
    styles.NextRecord();
    while (styles.record.id != BRT_BEGINCELLXFS) {
        if (!styles.NextRecord()) {
            Rcpp::Rcerr << "Could not find cell XF table in style sheet" << std::endl;
            return;
        }  
    }
    
    uint32_t icell_XF = 0; // iCellXF counter
    std::vector<int> idxs;
    
    while (styles.record.id != BRT_ENDCELLXFS) {
        if (styles.record.id == BRT_XF) {
            CellFormat fmt(styles.record);
            if (fmt.IsBuiltInFormat() && fmt.IsBuiltInDateTimeFormat()) {
                find_all(format_ref, icell_XF, idxs);
                for (int idx : idxs)
                    if (mapped_type[idx] == CONTENT_TYPE::TYPE_INTEGER || mapped_type[idx] == CONTENT_TYPE::TYPE_DOUBLE)
                        mapped_type[idx] = CONTENT_TYPE::TYPE_DATETIME;  
            } else {
                std::map<uint16_t, NumberFormat>::iterator mapIt = custom_formats.find(fmt.fmt);
                if (mapIt != custom_formats.end()) {
                    if (mapIt->second.is_datetime_format) {
                        find_all(format_ref, icell_XF, idxs);
                        for (int idx : idxs)
                            if (mapped_type[idx] == CONTENT_TYPE::TYPE_INTEGER || mapped_type[idx] == CONTENT_TYPE::TYPE_DOUBLE)
                                mapped_type[idx] = CONTENT_TYPE::TYPE_DATETIME;  
                    }
                }
            }
            icell_XF++;
        }

        if (!styles.NextRecord()) break;
    }
}

void PopulateLogicalVector(int start_row, int end_row, Rcpp::IntegerVector &rows, 
                           std::vector<int> &indices, Rcpp::IntegerVector &mapped_types, 
                           Rcpp::LogicalVector &bools, Rcpp::IntegerVector &ints, Rcpp::DoubleVector &doubles, 
                           Rcpp::CharacterVector &strs, Rcpp::LogicalVector &result) {
    
    std::vector<int>::iterator it = indices.begin();
    
    for (int row = start_row; row <= end_row; row++) {
        if (it == indices.end()) {
            result.push_back(NA_LOGICAL);
        } else {
            while ((it != indices.end()) && (rows[*it] < row)) it++;
        
            if ((it != indices.end()) && (rows[*it] == row)) {
                switch (mapped_types[*it]) {
                case CONTENT_TYPE::TYPE_BLANK:
                    result.push_back(NA_LOGICAL);
                    break;
                case CONTENT_TYPE::TYPE_ERROR:
                    result.push_back(NA_LOGICAL);
                    break;
                case CONTENT_TYPE::TYPE_DOUBLE:
                    result.push_back(doubles[*it] == 0);
                    break;
                case CONTENT_TYPE::TYPE_LOGICAL:
                    result.push_back(bools[*it]);
                    break;
                case CONTENT_TYPE::TYPE_INTEGER:
                    result.push_back(ints[*it] == 0);
                    break;
                case CONTENT_TYPE::TYPE_STRING:
                    if (iequals(strs[*it], "true") || iequals(strs[*it], "t")) result.push_back(true); else
                    if (iequals(strs[*it], "false") || iequals(strs[*it], "f")) result.push_back(false); else
                        result.push_back(NA_LOGICAL);
                    break;
                case CONTENT_TYPE::TYPE_DATETIME:
                    result.push_back(NA_LOGICAL);
                    break;
                default:
                    result.push_back(NA_LOGICAL);
                }
                it++;
            } else result.push_back(NA_LOGICAL);
        }
    }
}

void PopulateDoubleVector(int start_row, int end_row, Rcpp::IntegerVector &rows, 
                           std::vector<int> &indices, Rcpp::IntegerVector &mapped_types, 
                           Rcpp::LogicalVector &bools, Rcpp::IntegerVector &ints, Rcpp::DoubleVector &doubles, 
                           Rcpp::CharacterVector &strs, Rcpp::DoubleVector &result) {
  
    std::vector<int>::iterator it = indices.begin();
    
    for (int row = start_row; row <= end_row; row++) {
        if (it == indices.end()) {
            result.push_back(NA_REAL);
        } else {
            while ((it != indices.end()) && (rows[*it] < row)) it++;
            
            if ((it != indices.end()) && (rows[*it] == row)) {
                switch (mapped_types[*it]) {
                case CONTENT_TYPE::TYPE_BLANK:
                    result.push_back(NA_REAL);
                    break;
                case CONTENT_TYPE::TYPE_ERROR:
                    result.push_back(NA_REAL);
                    break;
                case CONTENT_TYPE::TYPE_DOUBLE:
                    result.push_back(doubles[*it]);
                    break;
                case CONTENT_TYPE::TYPE_LOGICAL:
                    result.push_back((bools[*it] == NA_LOGICAL ? NA_REAL : (double)bools[*it]));
                    break;
                case CONTENT_TYPE::TYPE_INTEGER:
                    result.push_back((double)ints[*it]);
                    break;
                case CONTENT_TYPE::TYPE_STRING:
                    result.push_back(Utils::ToDouble(strs[*it]));
                    break;
                case CONTENT_TYPE::TYPE_DATETIME:
                    result.push_back(doubles[*it]);
                    break;
                default:
                    result.push_back(NA_REAL);
                }
                it++;
            } else result.push_back(NA_REAL);
        }
    }
}

void PopulateIntegerVector(int start_row, int end_row, Rcpp::IntegerVector &rows, 
                           std::vector<int> &indices, Rcpp::IntegerVector &mapped_types, 
                           Rcpp::LogicalVector &bools, Rcpp::IntegerVector &ints, Rcpp::DoubleVector &doubles, 
                           Rcpp::CharacterVector &strs, Rcpp::IntegerVector &result) {
  
    std::vector<int>::iterator it = indices.begin();
    
    for (int row = start_row; row <= end_row; row++) {
        if (it == indices.end()) {
            result.push_back(NA_INTEGER);
        } else {
            while ((it != indices.end()) && (rows[*it] < row)) it++;
            
            if ((it != indices.end()) && (rows[*it] == row)) {
                switch (mapped_types[*it]) {
                case CONTENT_TYPE::TYPE_BLANK:
                    result.push_back(NA_INTEGER);
                    break;
                case CONTENT_TYPE::TYPE_ERROR:
                    result.push_back(NA_INTEGER);
                    break;
                case CONTENT_TYPE::TYPE_DOUBLE:
                    result.push_back((int)doubles[*it]);
                    break;
                case CONTENT_TYPE::TYPE_LOGICAL:
                    result.push_back((bools[*it] == NA_LOGICAL ? NA_INTEGER : (int)bools[*it]));
                    break;
                case CONTENT_TYPE::TYPE_INTEGER:
                    result.push_back(ints[*it]);
                    break;
                case CONTENT_TYPE::TYPE_STRING:
                    result.push_back(Utils::ToInt(strs[*it]));
                    break;
                case CONTENT_TYPE::TYPE_DATETIME:
                    result.push_back((int)doubles[*it]);
                    break;
                default:
                    result.push_back(NA_INTEGER);
                }
                it++;
            } else result.push_back(NA_INTEGER);
        }
    }
}

void PopulateDateTimeVector(int start_row, int end_row, Rcpp::IntegerVector &rows, 
                            std::vector<int> &indices, Rcpp::IntegerVector &mapped_types, 
                            Rcpp::LogicalVector &bools, Rcpp::IntegerVector &ints, Rcpp::DoubleVector &doubles, 
                            Rcpp::CharacterVector &strs, Rcpp::DoubleVector &result) {
  
    std::vector<int>::iterator it = indices.begin();
    
    for (int row = start_row; row <= end_row; row++) {
        if (it == indices.end()) {
            result.push_back(NA_REAL);
        } else {
            while ((it != indices.end()) && (rows[*it] < row)) it++;
        
            if ((it != indices.end()) && (rows[*it] == row)) {
                switch (mapped_types[*it]) {
                case CONTENT_TYPE::TYPE_BLANK:
                    result.push_back(NA_REAL);
                    break;
                case CONTENT_TYPE::TYPE_ERROR:
                    result.push_back(NA_REAL);
                    break;
                case CONTENT_TYPE::TYPE_DOUBLE:
                    result.push_back(SerialDate::ExcelToBase(doubles[*it]));
                    break;
                case CONTENT_TYPE::TYPE_LOGICAL:
                    result.push_back(NA_REAL);
                    break;
                case CONTENT_TYPE::TYPE_INTEGER:
                    result.push_back(SerialDate::ExcelToBase((double)ints[*it]));
                    break;
                case CONTENT_TYPE::TYPE_STRING:
                    result.push_back(Utils::ToDateTime(strs[*it]));
                    break;
                case CONTENT_TYPE::TYPE_DATETIME:
                    result.push_back(SerialDate::ExcelToBase(doubles[*it]));
                    break;
                default:
                    result.push_back(NA_REAL);
                }
                it++;
            } else result.push_back(NA_REAL);
        }
    }
    
    bool requires_POSIXct = false;
    double tmp;

    // Try to cast to Date, but if fraction of day (ie time encoded into value), then fall back to POSIXct
    for (auto d : result) {
        if (!R_IsNA(d) && (std::fabs(modf(d, &tmp) * 24 * 60 * 60) >= 0.5)) requires_POSIXct = true; // 84600 = 24*60*60
    }
    
    if (requires_POSIXct) {
        result.attr("class") = "POSIXct";
        result.attr("tzone") = "UTC";
        for (int i=0; i<result.size(); i++)
            if (!R_IsNA(result[i])) result[i] = result[i] * 86400.0f;
        
    } else {
        result.attr("class") = "Date";
    }
}

void PopulateCharacterVector(int start_row, int end_row, Rcpp::IntegerVector &rows, 
                             std::vector<int> &indices, Rcpp::IntegerVector &mapped_types, 
                             Rcpp::LogicalVector &bools, Rcpp::IntegerVector &ints, Rcpp::DoubleVector &doubles, 
                             Rcpp::CharacterVector &strs, Rcpp::CharacterVector &result) {
  
    std::vector<int>::iterator it = indices.begin();
    
    for (int row = start_row; row <= end_row; row++) {
        if (it == indices.end()) {
            result.push_back("");
        } else {
            while ((it != indices.end()) && (rows[*it] < row)) it++;

            if ((it != indices.end()) && (rows[*it] == row)) {
                switch (mapped_types[*it]) {
                case CONTENT_TYPE::TYPE_BLANK:
                    result.push_back("");
                    break;
                case CONTENT_TYPE::TYPE_ERROR:
                    result.push_back(NA_STRING);
                    break;
                case CONTENT_TYPE::TYPE_DOUBLE:
                    result.push_back(Rcpp::String(doubles[*it]));
                    break;
                case CONTENT_TYPE::TYPE_LOGICAL:
                    result.push_back((bools[*it] == NA_LOGICAL ? NA_STRING : (bools[*it] ? Rcpp::String("TRUE") : Rcpp::String("FALSE"))));
                    break;
                case CONTENT_TYPE::TYPE_INTEGER:
                    result.push_back(Rcpp::String(ints[*it]));
                    break;
                case CONTENT_TYPE::TYPE_STRING:
                    result.push_back(strs[*it]);
                    break;
                case CONTENT_TYPE::TYPE_DATETIME:
                    result.push_back(SerialDate::BaseToString(doubles[*it]));
                    break;
                default:
                    result.push_back(NA_STRING);
                }
                it++;
            } else result.push_back("");
        }
    }
}

// [[Rcpp::export]]
Rcpp::DataFrame TransformContents(Rcpp::Environment xlsb_env, int start_row, int end_row, int start_col, int end_col,
                       Rcpp::IntegerVector col_int_types, Rcpp::CharacterVector col_names) {
  
    if (end_row == NA_INTEGER) end_row = MAX_ROW;
    if (end_col == NA_INTEGER) end_col = MAX_COL;
    if (col_int_types.size() != col_names.size()) {
        Rcpp::Rcerr << "Expecting column types to be the same length as column names" << std::endl;
        return R_NilValue;
    }
    
    Rcpp::List result;
    Rcpp::CharacterVector final_col_names;
    Rcpp::DataFrame content = as<Rcpp::DataFrame>(xlsb_env["content"]);
    Rcpp::IntegerVector rows = content["row"];
    Rcpp::IntegerVector columns = content["column"];
    Rcpp::IntegerVector mapped_types = content["mapped.type"];
    Rcpp::LogicalVector col_bools = content["logical.value"];
    Rcpp::IntegerVector col_ints = content["int.value"];
    Rcpp::DoubleVector col_doubles = content["double.value"];
    Rcpp::CharacterVector col_strs = content["str.value"];
    
    std::vector<int> column_idxs;
    
    for (int column = start_col; column <= end_col; column++) {
        // Only process columns that are not ignored
        if (col_int_types[column - start_col] != CONTENT_TYPE::TYPE_IGNORE) {
            int col_type;
            find_all(columns, column, column_idxs);
            
            int first_read_row = 0;
            while ((column_idxs.size() > first_read_row) && (rows[column_idxs[first_read_row]] < start_row))
                first_read_row++;
            
            if (first_read_row > 0) column_idxs.erase(column_idxs.begin(), 
                                                      column_idxs.begin() + first_read_row);
        
            // If the contents dataframe doesn't contain this column, handle separately
            if (column_idxs.size() == 0) {
                col_type = (col_int_types[column - start_col] == CONTENT_TYPE::TYPE_IMPLY ? 
                            CONTENT_TYPE::TYPE_INTEGER : col_int_types[column - start_col]);
            } else {
                if (col_int_types[column - start_col] == CONTENT_TYPE::TYPE_IMPLY) {
                    col_type = CONTENT_TYPE::TYPE_INTEGER; // Default type
                    // From most to least fragile data types
                    // logical -- datetime -- integer -- double -- string
                    if (find_any(mapped_types, column_idxs, CONTENT_TYPE::TYPE_STRING))
                        col_type = CONTENT_TYPE::TYPE_STRING; else
                    if (find_any(mapped_types, column_idxs, CONTENT_TYPE::TYPE_DOUBLE))
                        col_type = CONTENT_TYPE::TYPE_DOUBLE; else
                    if (find_any(mapped_types, column_idxs, CONTENT_TYPE::TYPE_INTEGER))
                        col_type = CONTENT_TYPE::TYPE_INTEGER; else
                    if (find_any(mapped_types, column_idxs, CONTENT_TYPE::TYPE_DATETIME))
                        col_type = CONTENT_TYPE::TYPE_DATETIME; else
                    if (find_any(mapped_types, column_idxs, CONTENT_TYPE::TYPE_LOGICAL))
                        col_type = CONTENT_TYPE::TYPE_LOGICAL;
                } else col_type = col_int_types[column - start_col];
            }
        
            switch (col_type) {
            case CONTENT_TYPE::TYPE_LOGICAL: {
                Rcpp::LogicalVector col_res;
                PopulateLogicalVector(start_row, end_row, rows, column_idxs, mapped_types,
                                        col_bools, col_ints, col_doubles, col_strs, col_res);
                result.push_back(col_res);
                break;                   
            }
            case CONTENT_TYPE::TYPE_INTEGER: {
                Rcpp::IntegerVector col_res;
                PopulateIntegerVector(start_row, end_row, rows, column_idxs, mapped_types,
                                        col_bools, col_ints, col_doubles, col_strs, col_res);
                result.push_back(col_res);
                break;
            }
            case CONTENT_TYPE::TYPE_DOUBLE: {
                Rcpp::DoubleVector col_res;
                PopulateDoubleVector(start_row, end_row, rows, column_idxs, mapped_types,
                                    col_bools, col_ints, col_doubles, col_strs, col_res);
                result.push_back(col_res);
                break;
            }
            case CONTENT_TYPE::TYPE_STRING: {
                Rcpp::CharacterVector col_res;
                PopulateCharacterVector(start_row, end_row, rows, column_idxs, mapped_types,
                                        col_bools, col_ints, col_doubles, col_strs, col_res);
                result.push_back(col_res);
                break;
            }
            case CONTENT_TYPE::TYPE_DATETIME: {
                Rcpp::DoubleVector col_res;
                PopulateDateTimeVector(start_row, end_row, rows, column_idxs, mapped_types,
                                        col_bools, col_ints, col_doubles, col_strs, col_res);
                result.push_back(col_res);
                break;
            }
            default:
                Rcpp::Rcerr << "An unexpected column type was encountered (type=" << col_type << ")" << std::endl;
                return R_NilValue;
            }
            
            final_col_names.push_back(col_names[column - start_col]);
        }
    }

    result.attr("class") = "data.frame";
    result.attr("names") = final_col_names;
    result.attr("row.names") = Rcpp::IntegerVector::create(NA_INTEGER, 
                                (final_col_names.size() == 0 ? 0 : -(end_row - start_row + 1)));
    
    return result; 
}

} // namespace readxlsb
