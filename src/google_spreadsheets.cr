require "http/client"
require "uri"
require "json"
# TODO: Write documentation for `GoogleSpreadsheets`
module GoogleSpreadsheets
  VERSION = "0.1.0"

  # TODO: Put your code here
  record GridProperties, rows : Int32, columns : Int32

  record SheetRow, data : Array(String) do
    include Enumerable(String)

    def each
      data.each {|x| yield x}
    end

    def <<(x)
      data << x.to_s
    end
  end

  alias CellInput = String | Int64 | Int32 | Bool | Nil

  abstract struct ValueRangeData
    include JSON::Serializable
    property values : Array(Array(CellInput))
  end

  struct ValueRange < ValueRangeData
    property range : String
    property majorDimension : MajorDimension

    def initialize(@range, values, @majorDimension = MajorDimension::ROWS)
      @values = values.map &.map &.as(CellInput)
    end
  end

  struct DataFilterValueRange < ValueRangeData
    property dataFilter : DataFilter
    property majorDimension : MajorDimension

    def initialize(@dataFilter, values, @majorDimension = MajorDimension::ROWS)
      @values = values.map &.map &.as(CellInput)
    end
  end

  abstract struct DataFilter
    include JSON::Serializable
  end

  struct RangeDataFilter < DataFilter
    property a1Range : String
    def initialize(@a1Range)
    end
  end

  struct GridRangeDataFilter < DataFilter
    property gridRange : GridRange
    def initialize(@gridRange)
    end
  end

  struct GridRange
    include JSON::Serializable
    @sheetId : Int64
    @startRowIndex : Int64
    @endRowIndex : Int64
    @startColumnIndex : Int64
    @endColumnIndex : Int64

    def initialize(@sheetId, @startRowIndex, @endRowIndex, @startColumnIndex, @endColumnIndex)
    end
  end

  struct MatchedValueRange
    include JSON::Serializable

    property valueRange : ValueRange
  end

  struct UpdateValuesByDataFilterResponse
    include JSON::Serializable

    getter updatedRange : String
    getter updatedRows : Int64
    getter updatedColumns : Int64
    getter updatedCells : Int64
    getter updatedData : ValueRange
  end

  struct UpdateValuesResponse
    include JSON::Serializable

    getter spreadsheetId : String
    getter updatedRange : String
    getter updatedRows : Int64
    getter updatedColumns : Int64
    getter updatedCells : Int64
    getter updatedData : ValueRange
  end

  enum MajorDimension
    ROWS
    COLUMNS
  end

  enum ValueInputOption
    RAW
    USER_ENTERED
  end

  class AuthException < Exception

  end

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
      grid = data["gridProperties"]
      Worksheet.new(spreadsheet: spreadsheet,
        id: data["sheetId"].as_i64,
        title: data["title"].as_s,
        index: data["index"].as_i,
        properties: GridProperties.new(
          rows: grid["rowCount"].as_i,
          columns: grid["rowCount"].as_i),
        sheet_type: data["sheetType"].as_s)
    end

    def before_updating
      raise AuthException.new("OAuth 2.0 token needed to edit the document") if @spreadsheet.oauth_token.nil?
    end

    def before_reading
      if @spreadsheet.oauth_token.nil? && @spreadsheet.api_key.nil?
        raise AuthException.new("OAuth 2.0 token or API Key are needed to edit the document")
      end
    end

    def append(data, range : String = "A1", major_dimension = MajorDimension::ROWS)
      before_updating
      sheeted_range = "#{title}!#{range}"
      url = "/v4/spreadsheets/#{@spreadsheet.id}" \
            "/values/#{sheeted_range}:append" \
            "?valueInputOption=USER_ENTERED"
      payload = ValueRange.new(sheeted_range, [data], major_dimension).to_json
      res = @spreadsheet.client.post url, body: payload, headers: @spreadsheet.auth_headers
    end

    def append(data : Hash(String, CellInput), range : String = "A1")
      before_reading
      keys = titles(range)
      newData = keys.map { |k| data[k] }
      append(newData, range, MajorDimension::ROWS)
    end

    def titles(range : String)
      before_reading
      title_range = if range.includes? ":"
        range
      else
        "#{range}:#{column_to_letters(@properties.columns)}#{range[1]}"
      end
      get(title_range).values[0]
    end

    def column_to_letters(column : Int32)
      letters = ""
      while column != 0
        letters = "#{letters}#{'A' + (column % 26)}"
        column = column // 26
      end
    end

    def clear(range : String)
      before_updating
      sheeted_range = range.includes?("!") ? range : "#{title}!#{range}"
      url = "/v4/spreadsheets/#{@spreadsheet.id}/values/#{sheeted_range}:clear"
      res = @spreadsheet.client.post url, headers: @spreadsheet.auth_headers
    end

    def clear(ranges : Array(String))
      before_updating
      payload = {
        ranges: ranges.map do |range|
          range.includes?("!") ? range : "#{title}!#{range}"
        end
      }.to_json
      url = "/v4/spreadsheets/#{@spreadsheet.id}/values:batchClear"
      res = @spreadsheet.client.post url, body: payload, headers: @spreadsheet.auth_headers
    end

    def clear(data_filters : Array(DataFilter))
      before_updating
      payload = { dataFilters: data_filters }.to_json
      url = "/v4/spreadsheets/#{@spreadsheet.id}/values:batchClearByDataFilter"
      res = @spreadsheet.client.post url, body: payload, headers: @spreadsheet.auth_headers
    end

    def update(range, data, *,
        major_dimension = MajorDimension::ROWS,
        value_input_option = ValueInputOption::USER_ENTERED)
      sheeted_range = range.includes?("!") ? range : "#{title}!#{range}"
      update(ValueRange.new(sheeted_range, data, major_dimension),
        major_dimension: major_dimension,
        value_input_option: value_input_option
      )
    end

    def update(value_range : ValueRange, *,
      major_dimension = MajorDimension::ROWS,
      value_input_option = ValueInputOption::USER_ENTERED)
      before_updating
      url = "/v4/spreadsheets/" \
            "#{@spreadsheet.id}/values/#{value_range.range}" \
            "?key=#{@spreadsheet.api_key}&valueInputOption=#{value_input_option.to_s}&includeValuesInResponse=true"
      payload = value_range.to_json
      res = @spreadsheet.client.put url, body: payload, headers: @spreadsheet.auth_headers
      UpdateValuesResponse.from_json(res.body)
    end

    def update(value_ranges : Array(ValueRange), *,
      major_dimension = MajorDimension::ROWS,
      value_input_option = ValueInputOption::USER_ENTERED)
      before_updating
      url = "/v4/spreadsheets/#{@spreadsheet.id}/values:batchUpdate"
      payload = {
        valueInputOption: value_input_option,
        data: value_ranges,
        includeValuesInResponse: true,
      }.to_json
      res = @spreadsheet.client.post url, body: payload, headers: @spreadsheet.auth_headers
      Array(UpdateValuesResponse).from_json(res.body, "responses")
    end

    def update(value_ranges : Array(DataFilterValueRange), *,
      major_dimension = MajorDimension::ROWS,
      value_input_option = ValueInputOption::USER_ENTERED)
      before_updating
      url = "/v4/spreadsheets/#{@spreadsheet.id}/values:batchUpdateByDataFilter"
      payload = {
        valueInputOption: value_input_option,
        data: value_ranges,
        includeValuesInResponse: true,
      }.to_json
      res = @spreadsheet.client.post url, body: payload, headers: @spreadsheet.auth_headers
      Array(UpdateValuesByDataFilterResponse).from_json(res.body, "responses")
    end

    def get(range, major_dimension = MajorDimension::ROWS)
      before_reading
      sheeted_range = "#{title}!#{range}"
      url = "/v4/spreadsheets/" \
            "#{@spreadsheet.id}/values/#{sheeted_range}" \
            "?key=#{@spreadsheet.api_key}&majorDimension=#{major_dimension.to_s.upcase}"
      res = @spreadsheet.client.get url
      ValueRange.from_json(res.body)
    end

    def get(ranges : Array(String), major_dimension = MajorDimension::ROWS)
      before_reading
      queriable_ranges = ranges.map do |range|
        "ranges=#{
          range.includes?("!") ? range : "#{title}!#{range}"
        }"
      end.join("&")
      url = "/v4/spreadsheets/" \
            "#{@spreadsheet.id}/values:batchGet" \
            "?#{queriable_ranges}" \
            "&key=#{@spreadsheet.api_key}&majorDimension=#{major_dimension.to_s}"
      res = @spreadsheet.client.get url
      Array(ValueRange).from_json(res.body, "valueRanges")
    end

    def get(filters : Array(DataFilter), major_dimension = MajorDimension::ROWS)
      before_reading
      url = "/v4/spreadsheets/#{@spreadsheet.id}/values:batchGetByDataFilter"
      payload = {
        dataFilters: filters,
        majorDimension: major_dimension.to_s
      }.to_json
      res = @spreadsheet.client.post url, body: payload, headers: @spreadsheet.auth_headers
      pp res
      Array(MatchedValueRange).from_json(res.body, "valueRanges")
    end
  end

  class Spreadsheet
    property id : String
    property worksheets : Array(Worksheet) = Array(Worksheet).new
    getter client : HTTP::Client
    getter api_key : String
    getter oauth_token : String | Nil
    def oauth_token=(token)
      @oauth_token = token.includes?("Bearer") ? token.split(" ")[1] : token
    end

    def auth_headers
      headers = HTTP::Headers.new
      headers["Authorization"] = "Bearer #{oauth_token}"
      headers
    end

    def initialize(@id, @api_key, token : String? = nil)
      @client = HTTP::Client.new URI.parse("https://sheets.googleapis.com")
      @worksheets = load_worksheets
      @oauth_token = token.includes?("Bearer") ? token.split(" ")[1] : token unless token.nil?
    end

    def load_worksheets
      if api_key.nil?
        raise AuthException.new("OAuth 2.0 token or API Key are needed to edit the document")
      end
      res = @client.get "/v4/spreadsheets/#{@id}?key=#{@api_key}"
      sheets = if sheets_json = JSON.parse(res.body)["sheets"].as_a?
                sheets_json.map { |x| Worksheet.from_json(self, x["properties"]) }
              else
                [] of Worksheet
              end
    end

    def sheet?(name)
      @worksheets.select { |w| w.title == name }[0]?
    end

    def sheet(name)
      @worksheets.select { |w| w.title == name }[0]
    end
  end

end

