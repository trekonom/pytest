library(tidyverse)
library(reticulate)

reticulate::use_condaenv("pytest")

## generate some time series data for month & year
tbl <- expand_grid(Year = 2017:2020, Month = month.name) |>
  mutate(N = sample(100, size = n(), replace = TRUE))

## ggplot2 plot of the data so we know what to expect
fig <-
  ggplot(data = tbl) +
  geom_line(aes(
    x = Month, y = N, group = Year,
    colour = factor(Year)
  ), linewidth = 1) +
  theme_minimal() +
  NULL
print(fig) # see a ggplot2 version of same plot

# convert data to wide format to put in excel
tbl_wide_format <- tbl |>
  pivot_wider(names_from = Month, values_from = N)

# convert wide format data to pandas dataframe, to pass to python script
tbl_pandas <- r_to_py(tbl_wide_format)

## import python script
source_python("py/write_xlsx_and_chart_to_file.py")

## save chart using python script
save_time_series_as_xlsx_with_chart(
  tbl_pandas,
  "reticulate_pandas_writexlsx_excel_line_chart2.xlsx"
)
