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

  struct ValueRange
    include JSON::Serializable
    property range : String
    @[JSON::Field(key: "majorDimension")]
    property major_dimension : MajorDimension
    property values : Array(Array(CellInput))

    def initialize(@range, values, @major_dimension = MajorDimension::ROWS)
      @values = values.map &.map &.as(CellInput)
    end
  end

  enum MajorDimension
    ROWS
    COLUMNS

    def to_json(io)
      io << "\""
      io << to_s(self)
      io << "\""
    end
  end

  enum ValueInputOption
    RAW
    USER_ENTERED

    def to_json(io)
      io << "\""
      io << to_s(self)
      io << "\""
    end
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
      headers = HTTP::Headers.new
      headers["Authorization"] = "Bearer #{@spreadsheet.oauth_token}"
      res = @spreadsheet.client.post url, body: payload, headers: headers
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

    def clear(range)
      before_updating
      sheeted_range = range.includes?("!") ? range : "#{title}!#{range}"
      url = "/v4/spreadsheets/#{@spreadsheet.id}/values/#{sheeted_range}:clear?key=#{@spreadsheet.api_key}"
      res = @spreadsheet.client.post url
    end

    def update(range, data, *,
        major_dimension = MajorDimension::ROWS,
        value_input_option = ValueInputOption::USER_ENTERED)
      update(ValueRange.new(sheeted_range, data, major_dimension),
        major_dimension: major_dimension,
        value_input_option: value_input_option
      )
    end

    def update(value_range : ValueRange, *,
      major_dimension = MajorDimension::ROWS,
      value_input_option = ValueInputOption::USER_ENTERED)
      before_updating
      sheeted_range = range.includes?("!") ? range : "#{title}!#{range}"
      url = "/v4/spreadsheets/" \
            "#{@spreadsheet.id}/values/#{sheeted_range}" \
            "?key=#{@spreadsheet.api_key}&valueInputOption=#{value_input_option.to_s}"
      payload = value_range.to_json
      headers = HTTP::Headers.new
      headers["Authorization"] = "Bearer #{@spreadsheet.oauth_token}"
      res = @spreadsheet.client.put url, body: payload, headers: headers
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

    def batchGet(ranges, major_dimension = MajorDimension::ROWS)
      before_reading
      queriable_ranges = ranges.map do |range|
        "ranges=#{
          range.includes?("!") ? range : "#{title}!#{range}"
        }"
      end.join("&")
      url = "/v4/spreadsheets/" \
            "#{@spreadsheet.id}/values:batchGet" \
            "?#{queriable_ranges}" \
            "&key=#{@spreadsheet.api_key}&majorDimension=#{major_dimension.to_s.upcase}"
      res = @spreadsheet.client.get url
      Array(ValueRange).from_json(res.body, "valueRanges")
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

