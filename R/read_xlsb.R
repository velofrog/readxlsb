#' @useDynLib readxlsb, .registration = TRUE
#' @importFrom Rcpp evalCpp
#' @importFrom utils unzip
NULL

## internal constants
excel_MAX_COLS = 0x4000
excel_MAX_ROWS = 0x100000

## mapped.type enums
CONTENT_TYPE = list(TYPE_BLANK = 0, TYPE_ERROR = 1, TYPE_DOUBLE = 2, TYPE_LOGICAL = 3,
                    TYPE_INTEGER = 4, TYPE_STRING = 5, TYPE_DATETIME = 6)


NA_if_larger = function(value, boundary) {
  return(ifelse(value >= boundary, NA, value))
}

stop_if_not_defined = function(value, msg) {
  if (is.null(value) || is.na(value))
    stop(msg, call. = FALSE)
  
  if (is.character(value) && (value == ""))
    stop(msg, call. = FALSE)
}

## Converting Excel's serial date to an R date
## const double BASE_EXCEL_DATE = 25569 // 25569 => Excel's serial date for 1970-01-01
excel_to_R_datetime = function(excel_datetime) {
  if ((abs(excel_datetime)-floor(excel_datetime)) < 1e-8) {
    # return Date object
    return(structure(excel_datetime - 25569, class = "Date"))
  } else {
    # return POSIXct object
    return(structure((excel_datetime - 25569) * 24 * 60 * 60, class = "POSIXct", "tzone" = "UTC"))
  }
}

excel_date_to_string = function(excel_datetime) {
  return (format(excel_to_R_datetime(excel_datetime)))
}

#' Read xlsb workbook
#'
#' Imports a region from an xlsb workbook 
#' @usage read_xlsb(path, sheet = NULL, range = NULL, col_names = TRUE, 
#'  col_types = NULL, na = "", trim_ws = TRUE, skip = 0, ...)
#' @param path Path to the xlsb workbook
#' @param sheet Name or index of the sheet to read. If the sheet name is specified as part of the range, this parameter is ignored
#' @param range A named range or a string representing an excel range (of the form Sheet!A1:D10) or a cellranger object
#' @param col_names TRUE uses the first row as the column name, FALSE sets names to column.#, or a character vector
#' @param col_types NULL to imply type from spreadsheet or one of ignore/logical/numeric/date/string per column
#' @param na String to interpret as missing
#' @param trim_ws Trim whitespace from strings
#' @param skip Number of rows to skip before reading data
#' @param ... Additional options. Pass debug = TRUE to return xlsb environment
#' @examples
#' read_xlsb(path = system.file("extdata", "TestBook.xlsb", package = "readxlsb"), 
#'  range = "PORTFOLIO")
#'
#' @export
read_xlsb = function(path, sheet = NULL, range = NULL, col_names = TRUE, col_types = NULL,
                     na = "", trim_ws = TRUE, skip = 0, ...) {
  
  ## Parameter checks
  stop_if_not_defined(path, "path missing")
  if (!is.numeric(skip) || (skip < 0)) stop("Expecting non-negative value for skip")
  
  ## Store intermediate objects in an environment
  xlsb_env = new.env()

  xlsb_contents = unzip(zipfile = path, list = TRUE)
  idx = which(grepl("xl/workbook.bin", xlsb_contents$Name))
  if (length(idx) == 0) stop("Failed to find xl/workbook.bin in file")
  cn = unz(path, filename = xlsb_contents[idx, "Name"], open = "rb")
  xlsb_env$stream = readBin(cn, "raw", xlsb_contents[idx, "Length"])
  close(cn)
  
  ## Parse workbook.bin
  ParseWorkbook(xlsb_env)
  
  ## Use workbook.bin.rels XML file to extract Type and Target fields
  idx = which(grepl("xl/_rels/workbook.bin.rels", xlsb_contents$Name))
  if (length(idx) == 0) stop("Failed to find xl/_rels/workbook.bin.rels in file")
  cn = unz(path, filename = xlsb_contents[idx, "Name"], open = "rb")
  xlsb_env$stream = readBin(cn, "raw", xlsb_contents[idx, "Length"])
  close(cn)
  
  doc = xml2::read_xml(xlsb_env$stream)
  rels = as.data.frame(do.call(rbind, lapply(xml2::xml_children(doc), xml2::xml_attrs)), 
                       stringsAsFactors = FALSE)
  xlsb_env$sheets = cbind(xlsb_env$sheets, rels[match(xlsb_env$sheets$Id, rels$Id), -1])
  rownames(xlsb_env$sheets) = NULL
  
  ## If sheet index specified, replace with name
  if (is.numeric(sheet)) {
    if ((sheet <= 0) || (sheet > nrow(xlsb_env$sheets))) {
      stop("sheet index out-of-bounds")
    }
    sheet = xlsb_env$sheets$Name[sheet]
  }
  
  ## If range is not specified, but sheet is, then set range to the entire sheet
  if ((is.null(range) || is.na(range)) && (!is.null(sheet) && !is.na(sheet))) {
    range = cellranger::cell_limits(ul = c(1, 1), lr = c(NA, NA), sheet = sheet)
  }
  
  ## If range is a Named Range in the workbook, convert to cellranger range
  if (is.character(range)) {
    idx = match(tolower(range), tolower(xlsb_env$named_ranges$name))
    if (!is.na(idx)) {
      sheet = xlsb_env$sheets$Name[xlsb_env$named_ranges$sheet_idx[idx] + 1]
      range = cellranger::cell_limits(ul = c(xlsb_env$named_ranges$first_row[idx],
                                             xlsb_env$named_ranges$first_column[idx]),
                                      lr = c(NA_if_larger(xlsb_env$named_ranges$last_row[idx], excel_MAX_ROWS),
                                             NA_if_larger(xlsb_env$named_ranges$last_column[idx], excel_MAX_COLS)),
                                      sheet = xlsb_env$sheets$Name[xlsb_env$named_ranges$sheet_idx[idx] + 1])
    } else {
      range = cellranger::as.cell_limits(range)
      if (is.na(range$sheet)) range$sheet = sheet else sheet = range$sheet
    }
  }
  
  if (!inherits(range, "cell_limits")) {
    stop("range: expecting named or Excel-style A1/R1C1 reference, or a cellranger cell_limits object")
  }
  
  ## Fixed upper left coordinates of range if either are NA
  if (is.na(range$ul[1])) range$ul[1] = 1L
  if (is.na(range$ul[2])) range$ul[2] = 1L
  
  ## If sheet unspecified, extract from range
  if (is.null(sheet) || (is.na(sheet)) || (sheet=="")) {
    sheet = range$sheet
  }
  
  if (is.na(sheet)) stop("sheet: missing or not specified in range object")
  
  ## Extract worksheet binary
  idx = match(tolower(sheet), tolower(xlsb_env$sheets$Name))
  if (is.na(idx) || (length(idx) != 1)) {
    stop(paste0("Cannot find worksheet '", sheet, "' in workbook"))
  }
  idx = which(grepl(xlsb_env$sheets[idx, "Target"], xlsb_contents[, "Name"]))
  if (is.na(idx)) {
    stop(paste0("Cannot find binary worksheet '", sheet, "' in xlsb file"))
  }
  cn = unz(path, filename = xlsb_contents[idx, "Name"], open = "rb")
  xlsb_env$stream = readBin(cn, "raw", xlsb_contents[idx, "Length"])
  close(cn)
  
  ## Parse worksheet, extract raw contents
  ParseWorksheet(xlsb_env, start_row = range$ul[1], end_row = range$lr[1], start_col = range$ul[2],
                 end_col = range$lr[2])

  ## Populate strings from the sharedStrings binary  
  idx = which(grepl("sharedStrings.bin", xlsb_contents$Name))
  if (length(idx) == 1) {
    cn = unz(path, filename = xlsb_contents[idx, "Name"], open = "rb")
    xlsb_env$stream = readBin(cn, "raw", xlsb_contents[idx, "Length"])
    close(cn)
    ParseSharedStrings(xlsb_env)
  }

  ## Check for dates by parsing the styles binary
  idx = which(grepl("styles.bin", xlsb_contents$Name))
  if (length(idx) == 1) {
    cn = unz(path, filename = xlsb_contents[idx, "Name"], open = "rb")
    xlsb_env$stream = readBin(cn, "raw", xlsb_contents[idx, "Length"])
    close(cn)
    ParseStyles(xlsb_env)
  }
  
  ## Entirely blank rows and columns won't come through in ParseWorksheet
  ## So this range should be an actual (shrunk) data area
  row_range = range(xlsb_env$content$row)
  col_range = range(xlsb_env$content$column)
  
  ## Apply skip, if necessary taking into account rows that have been automatically
  ## skipped already
  row_range[1] = row_range[1] + max(0, skip - (row_range[1] - range$ul[1]))
  if (row_range[1] > row_range[2]) stop("No data found")
  
  ## for strings that match the na parameter, set type to doule and content to NA
  idx = intersect(which(xlsb_env$content$mapped.type == CONTENT_TYPE$TYPE_STRING),
                  which(xlsb_env$content$str.value == na))
  if (length(idx) > 0) {
    xlsb_env$content$mapped.type[idx] = CONTENT_TYPE$TYPE_DOUBLE
    xlsb_env$content$double.value[idx] = NA
    xlsb_env$content$str.value[idx] = NA
  }
  
  ## trim.ws if set
  if (trim_ws) {
    idx = which(xlsb_env$content$mapped.type == CONTENT_TYPE$TYPE_STRING)
    
    if (length(idx) > 0) {
      xlsb_env$content$str.value[idx] = trimws(xlsb_env$content$str.value[idx])
    }
  }
  
  ## Set up col_types object. User can specify column types 
  ## (logical, double, integer, date, character) or ignore to skip the column
  ## blank or non-handled descriptions implies type should be guessed from underlying data
  if (!is.null(col_types) && !is.na(col_types)) {
    col_types = rep(col_types, length.out = (col_range[2] - col_range[1] + 1))
  } else {
    col_types = rep("", length.out = (col_range[2] - col_range[1] + 1))
  }
  
  col_int_types = rep(-1, length = (col_range[2] - col_range[1] + 1))
  col_int_types[grepl("logi", col_types, ignore.case = TRUE)] = CONTENT_TYPE$TYPE_LOGICAL
  col_int_types[grepl("bool", col_types, ignore.case = TRUE)] = CONTENT_TYPE$TYPE_LOGICAL
  col_int_types[grepl("num", col_types, ignore.case = TRUE)] = CONTENT_TYPE$TYPE_DOUBLE
  col_int_types[grepl("double", col_types, ignore.case = TRUE)] = CONTENT_TYPE$TYPE_DOUBLE
  col_int_types[grepl("int", col_types, ignore.case = TRUE)] = CONTENT_TYPE$TYPE_INTEGER
  col_int_types[grepl("date", col_types, ignore.case = TRUE)] = CONTENT_TYPE$TYPE_DATETIME
  col_int_types[grepl("char", col_types, ignore.case = TRUE)] = CONTENT_TYPE$TYPE_STRING
  col_int_types[grepl("string", col_types, ignore.case = TRUE)] = CONTENT_TYPE$TYPE_STRING
  col_int_types[grepl("skip", col_types, ignore.case = TRUE)] = -2
  col_int_types[grepl("ignore", col_types, ignore.case = TRUE)] = -2
  
  ## Set up column names
  ## If TRUE, this is taken as the first row (header row). 
  if (is.logical(col_names)) {
    if (col_names) {
      col_names = paste0("column.", 1:(col_range[2] - col_range[1] + 1))
      hdr_row = xlsb_env$content[xlsb_env$content$row == row_range[1], ]
      if (nrow(hdr_row) > 0) {
        for (i in 1:nrow(hdr_row)) {
          str_value = with(hdr_row[i, ], 
                           if (mapped.type == CONTENT_TYPE$TYPE_BLANK) "" else
                           if (mapped.type == CONTENT_TYPE$TYPE_ERROR) "NA" else
                           if (mapped.type == CONTENT_TYPE$TYPE_DOUBLE) as.character(double.value) else
                           if (mapped.type == CONTENT_TYPE$TYPE_LOGICAL) as.character(logical.value) else
                           if (mapped.type == CONTENT_TYPE$TYPE_INTEGER) as.character(int.value) else
                           if (mapped.type == CONTENT_TYPE$TYPE_STRING) str.value else
                           if (mapped.type == CONTENT_TYPE$TYPE_DATETIME) excel_date_to_string(double.value) else ""
                      )
          col_names[hdr_row[i, "column"] - col_range[1] + 1] = str_value
        }
      }
      row_range[1] = row_range[1] + 1
      col_names = make.names(col_names, unique = TRUE)
    } else {
      col_names = paste0("column.", 1:(col_range[2] - col_range[1] + 1))
    }
  } else {
    ## col_names could be listing only non-ignored columns (type == -2), so handle that case too
    if (length(col_names) == length(which(col_int_types != -2))) {
      tmp_names = rep("ignored", length = (col_range[2] - col_range[1] + 1))
      tmp_names[which(col_int_types != -2)] = col_names
      col_names = tmp_names
    }
    
    col_names = make.names(rep(col_names, length = (col_range[2] - col_range[1] + 1)), 
                           unique = TRUE)
  }
  
  result = TransformContents(xlsb_env, row_range[1], row_range[2], col_range[1], col_range[2], 
                             col_int_types, col_names)

  ## check if 'debug' specified in dots. If so, return environment, else return result
  args = list(...)
  if ("debug" %in% names(args)) {
    if (args[["debug"]]) return(list(result = result, env = xlsb_env))
  }
  
  return(result)
}
