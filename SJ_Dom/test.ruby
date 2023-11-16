require "google_drive"
session = GoogleDrive::Session.from_config("config.json")

ws = session.spreadsheet_by_key("1Cf8XJWsoFWCXIYnwd5arUyzobThEDeLau7GauOMYwC0").worksheets[0]

GoogleDrive::Worksheet.class_eval do
  include Enumerable
  alias_method :old_s, :[]
  def each
    self.rows.filter.each do |row|
      row.each do |elem|
        unless elem.empty?
          yield elem
        end
      end
    end
  end
  def [] (*args)
    if args.length == 1
      headers = self.rows[0]
      if headers.include?(args[0])
        header_ind = headers.index(args[0])
        return self.rows.map {|row| row[header_ind] unless row[header_ind].empty? || row.include?("total") || row.include?("subtotal") }.compact
      end
    else 
      self.old_s(*args)
    end
  end

  def method_missing(method_name, *args, &block)
    column_name = method_name.to_s.downcase
    return self[column_name] if self.rows[0].include?(column_name)
    super
  end
end

Array.class_eval do
  def methode_missing(method)
    method = method[0..-1]
    return self.index(method)+1
  end
  def sum
    rez = 0
    self.each do |element|
      if element.match(/\A-?\d+(\.\d+)?\z/)
        rez += element.to_i
      end
    end
    return rez
  end
  def avg
    rez = 0
    i = 0
    self.each do |element|
      if element.match(/\A-?\d+(\.\d+)?\z/)
        rez += element.to_i
        i += 1
      end
    end
    return rez/(i)
  end
end



def add_method(c, m, &b)
  c.class_eval {
    define_method(m, &b)
  }
end

add_method(GoogleDrive::Worksheet, :table){
  table = []
  self.rows.filter.each do |row|
    unless row.include?("total") || row.include?("subtotal")
      row_arr = []
      row.each do |elem|
        unless elem.empty?
          row_arr << elem
        end
      end
      table << row_arr
    end
  end
  return table
 }
add_method(GoogleDrive::Worksheet, :row) {|index| return self.filter[index] }
add_method(GoogleDrive::Worksheet, :filter) {return ws.rows.reject { |sublist| sublist.include?("subtotal") || sublist.include?("total") }}

# p ws.filter
p ws.row(3)
p ws["prvakolona"][3]

 p ws.prvakolona
 p ws.prvakolona.sum
 p ws.table
