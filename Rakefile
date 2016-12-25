class Document
  attr_accessor :lines

  def initialize
    @lines = []
  end
end

class LineContainer
  attr_accessor :items

  def initialize
    @items = []
  end
end

class LineItem
  attr_accessor :field_id
  attr_accessor :section
  attr_accessor :data_stream
  attr_accessor :position
  attr_accessor :length
  attr_accessor :occurences
  attr_accessor :field_info
  attr_accessor :values
  attr_accessor :required_du
  attr_accessor :required_government
  attr_accessor :required_credit_request
  attr_accessor :required_early_check
  attr_accessor :mismo_element
end

task :parse_1003 do
  require 'rubyXL'
  require 'yaml'
  require 'json'
  require 'active_support/all'

  document = Document.new
  current_line = LineContainer.new

  book = RubyXL::Parser.parse('data/residential-loan-data-1003-integration-guide.xlsx')
  sheet = book['RLD 1003 3.2 Table']
  sheet.each_with_index do |row, i|
    next if i == 0 || i == 1
    if !row[3].nil? && row[3].value == '<N/L>'
      document.lines << current_line
      current_line = LineContainer.new
      next
    end

    item = LineItem.new
    item.field_id = row[1].try(:value)
    item.section = row[2].try(:value)
    item.data_stream = row[3].try(:value).try(:strip)
    item.position = row[4].try(:value).to_i
    item.length = row[5].try(:value).to_i
    item.occurences = row[6].try(:value)
    item.field_info = row[7].try(:value)
    item.values = row[8].try(:value)
    item.required_du = row[9].try(:value)
    item.required_government = row[10].try(:value)
    item.required_credit_request = row[11].try(:value)
    item.required_early_check = row[12].try(:value)
    item.mismo_element = row[13].try(:value)
    current_line.items << item
  end

  File.open('data/1003.json', 'w') do |f|
    f.write document.to_json
  end

  File.open('data/1003.yaml', 'w') do |f|
    f.write document.to_yaml
  end
end

task default: :parse_1003
