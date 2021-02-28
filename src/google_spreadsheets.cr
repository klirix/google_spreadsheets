require "http/client"
require "uri"
require "json"
# TODO: Write documentation for `GoogleSpreadsheets`
# module GoogleSpreadsheets
#   VERSION = "0.1.0"

#   # TODO: Put your code here

# end

record GridProperties, rows : Int32, columns : Int32

class Worksheet

  property spreadsheet : Spreadsheet

  property id : Int64
  property title : String
  property index : Int32
  property sheet_type : String
  @properties : GridProperties

  def initialize(@id, @title, @index, @sheet_type, @spreadsheet, @properties)

  end

  def self.from_json(spreadsheet, data : JSON::Any)
    gridData = data["gridProperties"]
    Worksheet.new(spreadsheet: spreadsheet,
      id: data["sheetId"].as_i64,
      title: data["title"].as_s,
      index: data["index"].as_i,
      properties: GridProperties.new(gridData["rowCount"].as_i, gridData["rowCount"].as_i),
      sheet_type: data["sheetType"].as_s)
  end

  def append()

  end

  def get(range)
    sheeted_range = "#{title}!#{range}"
    res = @spreadsheet.client.get "/v4/spreadsheets/#{@spreadsheet.id}/values/#{sheeted_range}?key=#{@spreadsheet.api_key}"
    parsed_res = JSON.parse(res.body)
    parsed_res["values"]
  end
end

class Spreadsheet
  property id : String
  property worksheets : Array(Worksheet) = Array(Worksheet).new
  getter client : HTTP::Client
  getter api_key : String
  def initialize(@id, @api_key)
    @client = HTTP::Client.new URI.parse("https://sheets.googleapis.com")
    @worksheets = load_worksheets
  end

  def load_worksheets
    res = @client.get "/v4/spreadsheets/#{@id}?key=#{@api_key}"
    sheets = if sheets_json = JSON.parse(res.body)["sheets"].as_a?
      sheets_json.map {|x| Worksheet.from_json(self, x["properties"])}
    else
      [] of Worksheet
    end
  end

  def sheet(name)
    @worksheets.select {|w| w.title == name }[0]?
  end
end

spreadsheet = Spreadsheet.new("1M6DSm-pfjWnuyHrk2tq2YQrFHHnxvlwn1KzmsesyRbs", "AIzaSyDfQVG-JprzWnbiL28LFv_Hf73vUjssXJw")
if worksheet = spreadsheet.sheet("Sheet1")
  pp worksheet.get("A1:b2")
end
