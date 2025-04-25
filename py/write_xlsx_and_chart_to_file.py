import pandas as pd
import xlsxwriter as xw

# The skeleton of below function based on example from: https://xlsxwriter.readthedocs.io/example_pandas_chart.html#ex-pandas-chart
# We pass the function a pandas dataframe;
# The dataframe is inserted in an .xslx spreadsheet
# We take note of the number of rows and columns, and use those to position the chart below the data
# We then iterate over the rows of the data and insert each row as a separate line (series) in the line chart

def save_time_series_as_xlsx_with_chart(pandas_df, filename):
  if not(filename.endswith('.xlsx')):
    print("Warning: added .xlsx to filename")
    filename = filename + '.xlsx'
  # Create a Pandas dataframe from the data.
  # pandas_df = pd.DataFrame({'Data': [10, 20, 30, 20, 15, 30, 45]})

  ## get dimensions of data frame to use for positioning the chart later
  pandas_df_nrow, pandas_df_ncol = pandas_df.shape

  # Create a Pandas Excel writer using XlsxWriter as the engine.
  writer = pd.ExcelWriter(filename, engine='xlsxwriter')

  # Convert the dataframe to an XlsxWriter Excel object.
  pandas_df.to_excel(writer, sheet_name='Sheet1', index=False)

  # Get the xlsxwriter workbook and worksheet objects.
  workbook  = writer.book
  worksheet = writer.sheets['Sheet1']

  # Create a chart object.
  chart = workbook.add_chart({'type': 'line'})

  # Configure the series of the chart from the dataframe data
  # THe coordinates of each series in the line chart are the positions of the data in the excel file
  # Note that data starts at row 2, column 1, so the row/col values need to be adjusted accordingly
  # However, python counts rows & columns from 0
  colors = {
    'funblue': '#155091', 'viking': '#70B7E1', 
    'anthrazit': '#666F77', 'hellgrau': '#B7BCBF',
    'bondiblue': '#0096B1', 'jetstream': '#A6CFC8'
  }
  colors = [
    '#155091', '#70B7E1', 
    '#666F77', '#B7BCBF',
    '#0096B1', '#A6CFC8'
  ]
  for row_in_data in range(0,pandas_df_nrow):
    row_in_sheet = row_in_data+1  # data starts on 2nd row
    last_col_in_sheet = pandas_df_ncol-1 # number of columns minus one in 0-notation
    first_col_with_data = 1  # 2nd column in 0-notation
    range_of_series = xw.utility.xl_range(
      first_row=row_in_sheet,  # read from the current row in loop only
      first_col=first_col_with_data, # data starts in 2nd column, i.e. 1 in 0-notation
      last_row=row_in_sheet,
      last_col=last_col_in_sheet
      )
    range_of_categories = xw.utility.xl_range(
      first_row=0, # read from 1st row only - header
      first_col=first_col_with_data,  # read from 2nd column for month headers
      last_row=0, 
      last_col=last_col_in_sheet
      )
    formula_for_series = '=Sheet1!' + range_of_series
    col_with_series_name = 0  # first column
    name_of_series = '=Sheet1!' + xw.utility.xl_rowcol_to_cell(row=row_in_sheet, col=col_with_series_name)
    formula_for_categories = 'Sheet1!' + range_of_categories
    chart.add_series({
      'values': formula_for_series, 
      'name': name_of_series, 
      'categories': formula_for_categories,
      'line': {'width': 1.5, 'color': colors[row_in_data]}
      })
    chart.set_plotarea({
      'fill': {'color': '#ECECED'},
      'layout': {
        'x': 0.055,
        'y': 0.06,
        'width': 0.93,
        'height': 0.725
      }
    })
    chart.set_chartarea({
      'fill': {'color': '#ECECED'},
      'border': {'none': True}
    })
    chart.set_x_axis({
      'name_font': {
        'size': 8, 
        'name': 'Franklin Gothic Book', 
        'color': '#3F464A'},
      'num_font': {
        'size': 8, 
        'name': 'Franklin Gothic Book', 
        'color': '#3F464A'},
      'line': {
        'color': '#B7BCBF', 
        'size': 0.75}
    })
    chart.set_y_axis({
       'name_font': {
        'size': 8, 
        'name': 'Franklin Gothic Book', 
        'color': '#3F464A'},
      'num_font': {
        'size': 8, 
        'name': 'Franklin Gothic Book', 
        'color': '#3F464A'},
      'line': {
        'color': '#B7BCBF', 
        'size': 0.75,
        'none': True},
      'major_gridlines': {
        'visible': True, 
        'line': {
          'color': '#B7BCBF', 
          'width': 0.75}},
      'major_tick_mark': 'none',
      'visible': True
    })
    chart.set_legend({
      'font': {
        'size': 8, 
        'name': 'Franklin Gothic Book', 
        'color': '#3F464A'},
      'position': 'bottom',
      'layout': {
        'x':      0.002,
        'y':      0.94,
        'width':  0.6,
        'height': 0.09,
      }
    })
    chart.set_size({
      'width': 604.8,
      'height': 453.4
    })
    
  # Insert the chart into the worksheet.
  worksheet.insert_chart(pandas_df_nrow+2, 2, chart)

  # Close the Pandas Excel writer and output the Excel file.
  writer.close()
