require "../src/google_spreadsheets.cr"
require "./credentials.cr"
include GoogleSpreadsheets

spreadsheet = Spreadsheet.new(
  id: ID,
  api_key: API_KEY,
  token: ACCESS_TOKEN
)

sheet = spreadsheet.sheet("Sheet1")

pp sheet.get("a1:b10")

pp sheet.get(["a1:a10", "b1:b10"])

pp sheet.get([RangeDataFilter.new(a1Range: "a1:b20")])

sheet.append({"name" => "Joe", "age" => 10})

sheet.update([ValueRange.new("b6", [["JORG"]])])
sheet.update("a6", [[12]])
sheet.clear("a7:b7")
sheet.clear(["a8:b8"])
sheet.clear([RangeDataFilter.new(a1Range: "a9:b9")])
sheet.update([DataFilterValueRange.new(RangeDataFilter.new("a8:b9"), [[12, "JORG"]])])
