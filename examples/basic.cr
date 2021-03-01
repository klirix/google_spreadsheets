require "../src/google_spreadsheets.cr"
require "./credentials.cr"

spreadsheet = GoogleSpreadsheets::Spreadsheet.new(
  id: ID,
  api_key: API_KEY
)
sheet = spreadsheet.sheet("Sheet1")

sheet.get("a1:b10", GoogleSpreadsheets::MajorDimensions::Rows).each do |row|
  row.each do |el|
    print el
  end
end

pp sheet.append({"name" => "Joe", "age" => 10})
