require 'roo'
require 'pp'

TBLNAME = 'tbCMDBApplication'
PRIMARY = 8
SECONDARY = 9

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
pri = nil
sec = nil
workbook.sheet(TBLNAME).each_row_streaming do |row|
  row_cells = row.map { |cell| cell.value }

  if num_rows == 0
    # pp row_cells;break
    pri = row_cells[PRIMARY]
    sec = row_cells[SECONDARY]
  else
    cols = {}
    cols[pri] = row_cells[PRIMARY].nil? ? '' : row_cells[PRIMARY]
    cols[sec] = row_cells[SECONDARY].nil? ? '' : row_cells[SECONDARY]
    id = /\{(?<guid>[A-Za-z0-9-]*)\}/.match(row_cells[0])
    puts create_sql_update(cols,id[:guid]) unless num_rows == 0
  end
  num_rows += 1

end
puts '-'*70,"Read #{num_rows} rows"
