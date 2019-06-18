# readxlsb
Import Excel binary (.xlsb) spreadsheets into R

Not on CRAN... yet... but give it a go 
```
library(devtools)
install_github("velofrog/readxlsb")

library(readxlsb)
read_xlsb(path = system.file("extdata", "TestBook.xlsb", package="readxlsb"), range="PORTFOLIO")
```
