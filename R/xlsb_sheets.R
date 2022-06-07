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
  parse_xlsb$sheets$Name
}
