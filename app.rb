require 'docx'
require 'sablon'
require 'yaml'
require 'money'
require './lib/client'
require 'io/console'

def write_invoice(client_name, output_filename, invoice)
  template_path = "#{@template_location}/#{client_name.downcase.gsub(' ', '_')}_template.docx"
  unless File.exists? template_path
    puts "Warning: Could not find template for #{client_name}"
    return
  end

  template = Sablon.template(File.expand_path(template_path))
  context = { invoice: invoice }
  output_path = @output_location + "/" + output_filename
  puts "Creating invoice #{output_path}"
  template.render_to_file File.expand_path(@output_location + "/" + output_filename), context
end

def add_line_items(line_items, matter)
  line_items << {
    file: matter.external_file_info,
    date: matter.work.first.date,
    description: matter.work.first.description,
    hours: matter.work.first.time_spent,
    rate: "$#{matter.work.first.rate}/hr",
    amount: matter.work.first.amount.format,
  }

  matter.work[1..-1].each {|work_item|
    line_items << {
      date: work_item.date,
      description: work_item.description,
      hours: work_item.time_spent,
      amount: work_item.amount.format,
    }
  }
  line_items.last[:subtotal] =  {
    amount: matter.subtotal.format,
    hst: matter.hst.format,
    amount_with_hst: matter.total.format
  }
end

I18n.config.available_locales = :en

settings = YAML.load(File.read('config/settings.yaml'))
docket_location = settings['docket_location']
@output_location = settings['output_location']
@template_location = settings['template_location']

docket_month = Time.now.strftime("%B %Y").upcase
print "Create invoices for which month? [#{docket_month}] "
selected_month = gets.strip.upcase
docket_month = selected_month unless selected_month.size.zero?

docket_file = docket_location + "/DAILY TIMESHEET (#{docket_month}).docx"
puts
puts "Reading docket file #{docket_file}"
doc = nil
begin
  doc = Docx::Document.open(docket_file)
rescue
  puts
  puts "Could not open file -- please check the filename"
end

puts

unless doc.nil?
  table = doc.tables[0]
  rows = table.rows
  rows.shift # header row
  date = nil
  clients = {}
  rows.each do |row|
    row_date = row.cells.first.paragraphs.first.to_s
    date = Date.parse(row_date).strftime("%b %-d/%y") if !row_date.empty?
    client_info = row.cells[1]
    internal_file_number = client_info.paragraphs.first.to_s
    client_id = internal_file_number.split('-').first

    begin
      clients[client_id] ||= Client.new(client_id)
    rescue Exception => e
      puts "Warning: #{e}"
      next
    end

    clients[client_id].add_work(date: date,
      internal_file_number: internal_file_number,
      external_file_info: client_info.paragraphs[1].to_s,
      description: row.cells[2],
      time_spent: row.cells[3].to_s.to_f)
  end

  puts

  clients.values.each do |client|
    if client.invoice_per == 'client'
      line_items = []
      client.matters.values.each {|matter|
        add_line_items(line_items, matter)
      }

      invoice = {
        date: Time.now.strftime("%B %-d, %Y"),
        number: "#{client.id}#{Date.parse(docket_month).strftime('%m%y')}0",
        line_items: line_items,
        total_fees: client.total_fees.format,
        total_hst: client.total_hst.format,
        total_payable: client.total_payable.format
      }
      write_invoice(client.name, "#{client.name} #{docket_month} Invoice.docx", invoice)
    else # invoice per matter
      client.matters.values.each_with_index {|matter, i|
        line_items = []
        add_line_items(line_items, matter)

        invoice = {
          date: Time.now.strftime("%B %-d, %Y"),
          number: "#{client.id}#{Date.parse(docket_month).strftime('%m%y')}#{i + 1}",
          line_items: line_items,
          total_fees: matter.subtotal.format,
          total_hst: matter.hst.format,
          total_payable: matter.total.format
        }
        write_invoice(client.name, "#{client.name} #{docket_month} Invoice ##{i + 1}.docx", invoice)
      }
    end
  end

  puts <<~DONE_MESSAGE

  DONE!

  Don't forget to:
   - adjust invoice numbers if required
   - adjust the invoice dates if required
   - adjust the File descriptions if required
   - any manual formatting you require

  DONE_MESSAGE
end

if /mingw/ =~ RUBY_PLATFORM
  `%SystemRoot%\\explorer.exe #{@output_location}`
end

puts "Press any key to exit"
STDIN.getch
