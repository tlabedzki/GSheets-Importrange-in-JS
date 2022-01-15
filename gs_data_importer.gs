/*
    Google Sheets Data Importer
    Describe: Can be used in case of problems with importrange.
    Author: Tomasz Łabędzki

    Script Structure:
    1. Main function with settings.
    2. Standard importrange function replacement (data_importer) with smaller methods for remove headers from souce data & text formatting.
*/

function data_importer_run_with_settings() {
  /* Settings: addresses, worksheets names & other settings */
  
  source_file = "1uSskgHmR48ASlC-JNqATQJd9-zgAJwD5tTEaddress1"
  target_file = "1E1ivL84pVWeBD20MDk-EBnuIrmTGbRvW6oEaddress2"

  source_sheet = "source_sheet"
  target_sheet = "target_sheet"

  headers = 1
  font_size_of_pasted_data = 8
  
  data_importer(source_file, target_file, source_sheet, target_sheet, headers);
}

function data_importer(source_file, target_file, source_sheet, target_sheet, headers) {
  Logger.log('Connecting to source file.')
  ss = SpreadsheetApp.openById(source_file);

  Logger.log('Connected, downloading the data from source file.')
  source_data = ss.getSheetByName(source_sheet).getDataRange().getValues();
  filtered_source_data = remove_headers(source_data,1);

  Logger.log('Data downloaded, connecting to target file.')
  target_ss = SpreadsheetApp.openById(target_file);

  Logger.log('Connected, pasting data to target file.')
  rows_in_target_range = target_ss.getSheetByName(target_sheet).getDataRange().getLastRow();
  columns_in_target_range = target_ss.getSheetByName(target_sheet).getDataRange().getNumColumns();
  rows_qty = rows_in_target_range - headers
  start_row = 1 + headers
  
  try{
    range_clear = target_ss.getSheetByName(target_sheet).getRange(start_row, 1, rows_qty, columns_in_target_range).clear();
  } catch(e) {
    Logger.log('No data to clear, starting of paste the data.')
  }
  
  target_range = target_ss.getSheetByName(target_sheet).getRange(start_row, 1, filtered_source_data.length, columns_in_target_range).setValues(filtered_source_data);
  text_formatting_set(target_ss, target_sheet, font_size_of_pasted_data);
  Logger.log('Data pasted and formatted.')
}

function remove_headers(array_to_clean, number_of_lines) {
  for(i = 0; i<number_of_lines; i++){
    array_to_clean.shift();
  }
  return array_to_clean
}

function text_formatting_set(target_ss, target_sheet, font_size_of_pasted_data) {
  set_font_size = target_ss.getSheetByName(target_sheet).getDataRange().setFontSize(font_size_of_pasted_data);
}
