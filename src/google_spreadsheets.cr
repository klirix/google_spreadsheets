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
  end

  enum MajorDimensions
    Rows
    Columns
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
      if @spreadsheet.oauth_token.nil? || @spreadsheet.api_key.nil?
        raise AuthException.new("OAuth 2.0 token or API Key are needed to edit the document")
      end
    end

    def append(data : Array(String | Int), range : String = "A1", major_dimension = MajorDimensions::Rows)
      before_updating
      sheeted_range = "#{title}!#{range}"
      url = "/v4/spreadsheets/#{@spreadsheet.id}" \
            "/values/#{sheeted_range}:append" \
            "?valueInputOption=USER_ENTERED"
      payload = {
        range:          sheeted_range,
        values:         [data],
        majorDimension: major_dimension.to_s,
      }.to_json
      headers = HTTP::Headers.new
      headers["Authorization"] = "Bearer #{@spreadsheet.oauth_token}"
      res = @spreadsheet.client.post url, body: payload, headers: headers
    end

    def append(data : Hash(String, String | Int), range : String = "A1")
      before_reading
      keys = titles(range)
      newData = keys.map { |k| data[k] }
      append(newData, range, MajorDimensions::Rows)
    end

    def titles(range : String)
      before_reading
      title_range = if range.includes? ":"
        range
      else
        "#{range}:#{column_to_letters(@properties.columns)}#{range[1]}"
      end
      get(title_range)[0].data
    end

    def column_to_letters(column : Int32)
      letters = ""
      while column != 0
        letters = "#{letters}#{'A' + (column % 26)}"
        column = column // 26
      end
    end

    def get(range, major_dimension = MajorDimensions::Rows)
      before_reading
      sheeted_range = "#{title}!#{range}"
      url = "/v4/spreadsheets/" \
            "#{@spreadsheet.id}/values/#{sheeted_range}" \
            "?key=#{@spreadsheet.api_key}&majorDimension=#{major_dimension.to_s.upcase}"
      res = @spreadsheet.client.get url
      lines = JSON.parse(res.body)["values"]
      lines.as_a.map { |line|
        SheetRow.new(line.as_a.map { |v| v.to_s })
      }
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

    def initialize(@id, @api_key, token : String? = nil )
      @client = HTTP::Client.new URI.parse("https://sheets.googleapis.com")
      @worksheets = load_worksheets
      @oauth_token = token.includes?("Bearer") ? token.split(" ")[1] : token unless token.nil?
    end

    def load_worksheets
      if oauth_token.nil? || api_key.nil?
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

