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

pp sheet.batchGet(["a1:a10", "b1:b10"])

pp sheet.append({"name" => "Joe", "age" => 10})
sheet.update("a6", [[12]])
sheet.clear("a6:b6")
