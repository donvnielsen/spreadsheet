require 'roo'
require 'pp'

TBLNAME = 'tbCMDBApplication'
DEVELOPERS = 9..10

def create_sql_update(cols)
  [
     "UPDATE #{TBLNAME} ",
     "SET #{} = #{}, #{} = #{}",
     "WHERE id = #{}"
  ].join('\n')
end

workbook = Roo::Spreadsheet.open './data/cmdb_master_v3.xlsx'
worksheets = workbook.sheets
puts "Found #{worksheets.count} worksheets"

pp worksheets

num_rows = 0
workbook.sheet(TBLNAME).each_row_streaming do |row|
  pp row;break
  # row_cells = row.map { |cell| cell.value }
  # pp row_cells
  # num_rows += 1
end
puts "Read #{num_rows} rows"
