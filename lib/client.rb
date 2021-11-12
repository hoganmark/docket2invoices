require 'yaml'
require './lib/matter.rb'

class Client
  attr_accessor :id, :name, :invoice_per, :rate
  attr_accessor :matters

  def self.init_configs
    @@configs = {}
    settings = YAML.load(File.read('config/settings.yaml'))
    clients_location = settings['clients_location']
    doc = Docx::Document.open(clients_location)
    table = doc.tables[0]
    rows = table.rows
    rows.shift # header row
    rows.each do |row|
      @@configs[row.cells[0].to_s] = {
        name: row.cells[1].to_s,
        rate: row.cells[2].to_s.to_i,
        invoice_per: row.cells[3].to_s.downcase
      }
    end
  end

  init_configs

  def initialize(id)
    self.id = id

    config = @@configs[id]
    raise "Unknown client: #{id}" unless config
    self.name = config[:name]
    self.invoice_per = config[:invoice_per]
    self.rate = config[:rate]
  end

  def add_work(date:, internal_file_number:, external_file_info:, description:, time_spent:)
    self.matters ||= {}
    matters[internal_file_number] ||= Matter.new(internal_file_number, external_file_info)
    # todo: override external_file_info with file_info.docx or show warning
    matters[internal_file_number].add_work(date, description, time_spent, rate)
  end

  def total_fees
    matters.values.inject(Money.new(0)) {|t, m| t = t + m.subtotal}
  end

  def total_hst
    matters.values.inject(Money.new(0)) {|t, m| t = t + m.hst}
  end

  def total_payable
    matters.values.inject(Money.new(0)) {|t, m| t = t + m.total}
  end

  def to_s
    <<~STRING
      #{name} (rate: #{rate}):
      ====================================
      #{matters.values.inject ("") {|s, matter| s = s + matter.to_s}}
      TOTAL FEES: #{total_fees.format}
      TOTAL HST: #{total_hst.format}
      TOTAL PAYABLE: #{total_payable.format}
    STRING
  end
end
