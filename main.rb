require 'roo'

workbook = Roo::Spreadsheet.open './test.xlsx'

if workbook.sheets.count > 1
    raise "Found #{workbook.sheets.count} worksheets, I accept only 1"
end
workbook.default_sheet = workbook.sheets.first
header = workbook.row(1)

puts "Reading file"
num_rows = 0
2.upto(workbook.last_row) do |row|
    row_data = Hash[header.zip workbook.row(row)]
    num_rows += 1
end
puts "Read #{num_rows} rows"