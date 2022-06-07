#' List all sheets in a .xlsb workbook
#'
#' Analogous to the \code{excel_sheets} function from \code{readxl}. Efficiently gets the sheet names without needing to read any surplus data.
#' @usage xlsb_sheets(path)
#' @param path Path to the xlsb workbook
#' @examples
#' xlsb_sheets(path = system.file("extdata", "TestBook.xlsb", package = "readxlsb"))
#'
#' @export
xlsb_sheets = function(path) {
  ## Parameter checks
  stop_if_not_defined(path, "path missing")
  
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
  xlsb_env$sheets$Name
}