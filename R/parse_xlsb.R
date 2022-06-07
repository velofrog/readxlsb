## Backend function for superior modularity
parse_xlsb <- function(path) {
  ## Parameter checks
  stop_if_not_defined(path, "path missing")
  
  xlsb_env = new.env()
  
  xlsb_contents = unzip(zipfile = path, list = TRUE)
  idx = which(grepl("xl/workbook.bin", xlsb_contents$Name))
  if (length(idx) == 0) stop("Failed to find xl/workbook.bin in file")
  cn = unz(path, filename = xlsb_contents[idx, "Name"], open = "rb")
  xlsb_env$stream = readBin(cn, "raw", xlsb_contents[idx, "Length"])
  close(cn)
  
  ## Parse workbook.bin
  ParseWorkbook(xlsb_env)
  xlsb_env
}
