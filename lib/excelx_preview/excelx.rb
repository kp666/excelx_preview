class NoRowError < StandardError
end

module ExcelX
  module Previewer

    @preview ={}
    STANDARD_FORMATS = {
        0 => 'General',
        1 => '0',
        2 => '0.00',
        3 => '#,##0',
        4 => '#,##0.00',
        9 => '0%',
        10 => '0.00%',
        11 => '0.00E+00',
        12 => '# ?/?',
        13 => '# ??/??',
        14 => 'mm-dd-yy',
        15 => 'd-mmm-yy',
        16 => 'd-mmm',
        17 => 'mmm-yy',
        18 => 'h:mm AM/PM',
        19 => 'h:mm:ss AM/PM',
        20 => 'h:mm',
        21 => 'h:mm:ss',
        22 => 'm/d/yy h:mm',
        37 => '#,##0 ;(#,##0)',
        38 => '#,##0 ;[Red](#,##0)',
        39 => '#,##0.00;(#,##0.00)',
        40 => '#,##0.00;[Red](#,##0.00)',
        45 => 'mm:ss',
        46 => '[h]:mm:ss',
        47 => 'mmss.0',
        48 => '##0.0E+0',
        49 => '@',
    }
    FORMATS = {
        "general"=>:float,
        "0"=>:float,
        "0.00"=>:float,
        "#,##0"=>:float,
        "#,##0.00"=>:float,
        "0%"=>:percentage,
        "0.00%"=>:percentage,
        "0.00E+00"=>:float,
        "# ?/?"=>:float,
        "# ??/??"=>:float,
        "mm-dd-yy"=>:date,
        "d-mmm-yy"=>:date,
        "d-mmm"=>:date,
        "mmm-yy"=>:date,
        "h:mm AM/PM"=>:time,
        "h:mm:ss AM/PM"=>:time,
        "hh:mm:ss AM/PM"=>:time,
        "h:mm"=>:time,
        "h:mm:ss"=>:time,
        "m/d/yy h:mm"=>:datetime,
        "#,##0 ;(#,##0)"=>:float,
        "#,##0 ;[Red](#,##0)"=>:float,
        "#,##0.00;(#,##0.00)"=>:float,
        "#,##0.00;[Red](#,##0.00)"=>:float,
        "mm:ss"=>:time,
        "[h]:mm:ss"=>:time,
        "mmss.0"=>:time,
        "##0.0E+0"=>:float,
        "@"=>:float,
        "yyyy\\-mm\\-dd"=>:date,
        "dd/mm/yy"=>:date,
        "hh:mm:ss"=>:time,
        "dd/mm/yy hh:mm"=>:datetime,
        "dd/mmm/yy"=>:date,
        "yyyy-mm-dd"=>:date,
        "hh:mm:ss am/pm" => :time,
        "mm/dd/yy hh:mm am/pm" =>:datetime,
        "mm/dd/yy" => :date,
        "h:mm am/pm" => :time,
        "m/d/yyyy" => :date,
        "m/d/yyyy h:mm" => :datetime,
        "hh:mm am/pm" => :time,
        "dd/mm/yyyy" => :date,
    }
    DATE_TIME_FORMAT={#add more formats when found
                      "mm/dd/yy" => "%y,%m,%d", #coz excelx stores in this format
                      "m/d/yyyy" =>"%Y,%m,%d",
                      "h:mm am/pm" => "%I:%M:%S %p"
    }

    def self.styles() #taken from roo gem
      @numFmts =[]
      @cellXfs =[]
      style= Nokogiri::XML(File.open("#{@tmp_folder}/xl/styles.xml"))
      style.xpath("//*[local-name()='numFmt']").each do |numFmt|
        numFmtId = numFmt.attributes['numFmtId']
        formatCode = numFmt.attributes['formatCode']
        @numFmts << [numFmtId, formatCode]
      end
      style.xpath("//*[local-name()='cellXfs']").each do |xfs|
        xfs.children.each do |xf|
          numFmtId = xf['numFmtId']
          @cellXfs << [numFmtId]
        end
      end
    end

    def self.attribute2format(s) #taken from roo gem
      result = nil
      @numFmts.each { |nf|
        # to_s weil das eine Nokogiri::XML::Attr und das
        # andere ein String ist
        if nf.first.to_s == @cellXfs[s.to_i].first
          result = nf[1]
          break
        end
      }
      unless result
        id = @cellXfs[s.to_i].first.to_i
        if STANDARD_FORMATS.has_key? id
          result = STANDARD_FORMATS[id]
        end
      end
      result
    end

    def self.format2type(format) #taken from roo gem
      if FORMATS.has_key? format
        FORMATS[format]
      else
        :float
      end
    end

    def self.datetimeformat(format, type)
      if DATE_TIME_FORMAT.has_key? format
        DATE_TIME_FORMAT[format]
      else
        if type ==:date
          "%Y,%m,%d"
        elsif type ==:time
          "%I:%M:%S %p"
        else
          ""
        end
      end
    end

    def self.datetime(value, type)
      seconds = (value.to_f - 25569) * 86400.0
      if type == :time
        (Time.at seconds).utc.strftime("%I:%M:%S %p") rescue value
      elsif type == :datetime
        (Time.at seconds).utc.strftime("%m/%d/%Y %I:%M:%S %p") rescue value
      elsif type == :date
        (Time.at seconds).utc.strftime("%m/%d/%Y") rescue value
      elsif type == :float
        value
      end
    end

    def self.datetime_whenis(value, type, format)

      date_or_time_format = datetimeformat(format, type)
      # pp "#{value} #{format} #{date_or_time_format} #{type}"
      if  type == :date
        #Date.new("#{value}".delete('date(').chop.split(",").map(&:to_i)).to_s  rescue ""
        Date.strptime(value, "date(#{date_or_time_format}").to_s rescue value # TODO do same formatting as in the sheet
      elsif type == :time
        Time.parse("#{value}".delete('time(').chop.split(",").map(&:to_i).join(":")).strftime(date_or_time_format) rescue value
      end

    end

    def self.content_from_link(link)
      if  link.children.first.name=="f"
        content = link.children.last.children.last.text
      else
        content = link.content
      end
    end

    def self.is?(c)
      c.children.each do |f_or_is|
        if f_or_is.name=="is"
          return true
        end
      end
      return false
    end

    def self.get_content(link, is=false)
      content = nil
      s_value = link["s"].to_i
      format = attribute2format(s_value).to_s.downcase.gsub(/\\/, "").gsub("-", "/")
      type = format2type(format)

      if is?(link)
        value = link.content.downcase
        datetime_whenis(value, type, format)
        return datetime_whenis(value, type, format)
      end
      if s_value == 0
        if link['t']=="s"
          content = @shared_strings[link.content.to_i]
        else
          content = content_from_link(link)
        end
      elsif s_value >0 && s_value <48
        content = datetime(content_from_link(link), type)
      else
        content = content_from_link(link)
      end
      content
    end

    def self.first_10_from_sheet()
      @preview[@sheet_name]= {}
      return if @row_count < 1
      @doc.xpath("//*[local-name()='row']")[0..10].each do |row|
        sheet = {}
        row.children.each do |c|
          content = get_content(c)
          sheet[c['r']] = content
        end
        @preview[@sheet_name]["Row#{row['r']}"]= sheet
      end
    end

    def self.preview(filename, sheets=false)

      @filename = filename.chomp(File.extname(filename)) rescue filename
      @tmp_folder = UUID.new.generate
      unzip
      @sheet_list = extracted_sheets unless sheets
      @sheet_list = [sheets] if sheets
      shared_strings
      @sheet_list.each do |sheet_name_in_xml, sheet_name_in_excelx|
        parse_xml(sheet_name_in_xml, sheet_name_in_excelx)
        fetch_headers
        styles
        first_10_from_sheet
      end
      cleanup()
      @preview
    end

    def self.unzip()
      `unzip -o  #{@filename}.xlsx -d #{@tmp_folder}`
    end

    def self.cleanup()
      `rm -rf #{@tmp_folder}`
    end

    def self.extracted_sheets()
      sheet_list_xml= Nokogiri::XML(File.open("#{@tmp_folder}/xl/workbook.xml"))
      @sheet_list = {}
      sheet_list_xml.xpath("//*[local-name()='sheet']").each do |sheet|
        sheet_name_in_xml = "sheet#{sheet.attributes["sheetId"].value}"
        sheet_name_in_excelx = sheet.attributes["name"].value.downcase
        @sheet_list[sheet_name_in_xml] = sheet_name_in_excelx
      end
      @sheet_list
    end

    def self.shared_strings()
      @shared_strings =[]
      begin
        shared = Nokogiri::XML(File.open ("#{@tmp_folder}/xl/sharedStrings.xml"))
      rescue
        "no shared file"
        return @shared_strings
      end
      shared.xpath("//*[local-name()='si']").each do |shared_strings|
        @shared_strings << shared_strings.content
      end


      @shared_strings
    end

    def self.parse_xml(sheet_name_in_xml, sheet_name_in_excelx)
      @sheet_name = sheet_name_in_excelx
      sheet_xml_file = "#{@tmp_folder}/xl/worksheets/#{sheet_name_in_xml}.xml"
      xml_string =""
      sheets =[]
      counter = 0
      File.open(sheet_xml_file, "r") do |f|
        while counter<10 do
          buffer = f.read(1024)
          occurrence = 0
          occurrence = buffer.scan("</row>").size rescue break
          counter += occurrence
          xml_string = "#{xml_string}#{buffer}"
        end
      end
      rows = xml_string.split("</row>")[0..-2].join("</row>")
      @doc = Nokogiri.XML(rows)

    end

    def self.fetch_headers()
      @headers =[]
      @row_count = @doc.xpath("//*[local-name()='row']").count
      return if @row_count <1
      @doc.xpath("//*[local-name()='row']")[0].children.each do |first_row|
        if  first_row['t']=="s"
          @headers<< @shared_strings[first_row.content.to_i]
        else
          @headers<< first_row.content
        end
      end
      @headers
    end

  end
end



