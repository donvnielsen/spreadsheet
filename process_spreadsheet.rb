require 'roo'

workbook = Roo::Spreadsheet.open './sample_excel_files/xlsx_500_rows.xlsx'
worksheets = workbook.sheets
puts "Found #{worksheets.count} worksheets"

worksheets.each do |worksheet|
  puts "Reading: #{worksheet}"
  num_rows = 0
  workbook.sheet(worksheet).each_row_streaming do |row|
    row_cells = row.map { |cell| cell.value }
    num_rows += 1
  end
  puts "Read #{num_rows} rows"
end