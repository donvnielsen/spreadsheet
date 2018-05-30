require 'roo'
require 'pp'

TBLNAME = 'tbCMDBApplication'
COLUMNS = [8,9]
RGXID = /\{(?<guid>[A-Za-z0-9-]*)\}/

def create_sql_update(cols,id)
  ary = ["UPDATE #{TBLNAME}"]
  i = 0
  cols.each {|k,v|
    ary << "#{(i+=1) == 1 ? 'SET ' : ', '}#{k} = '#{v}'"
  }
  ary << "WHERE id = '#{id}';"
  ary.join(' ')
end

workbook = Roo::Spreadsheet.open './data/cmdb_master_v3.xlsx'
worksheets = workbook.sheets
puts "Found #{worksheets.count} worksheets:"
puts worksheets,'-'*70

num_rows = 0
col_names = nil
workbook.sheet(TBLNAME).each_row_streaming do |row|
  row_cells = row.map { |cell| cell.value }

  if num_rows == 0
    col_names = row_cells
    pp col_names
  else
    cols = {}
    COLUMNS.each {|col|
      cols[col_names[col]] = row_cells[col].nil? ? '' : row_cells[col]
    }
    id = RGXID.match(row_cells[0])
    puts create_sql_update(cols,id[:guid])
  end
  num_rows += 1

end
puts '-'*70,"Read #{num_rows} rows"
