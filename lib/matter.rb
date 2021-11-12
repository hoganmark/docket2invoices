require './lib/work_item.rb'

class Matter
  attr_accessor :internal_file_number, :external_file_info
  attr_accessor :work

  def initialize(internal_file_number, external_file_info)
    self.internal_file_number = internal_file_number
    self.external_file_info = external_file_info
  end

  def add_work(date, description, time_spent, rate)
    self.work ||= []
    work << WorkItem.new(date, description, time_spent, rate)
  end

  def subtotal
    work.inject (Money.new(0)) {|s, item| s = s + item.amount}
  end

  def hst
    subtotal * 0.13
  end

  def total
    subtotal + hst
  end

  def to_s
    <<~STRING
      #{external_file_info}:
      --------------------------------------
      #{work.inject ("") {|s, item| s = s + item.to_s + "\n"}}
      SUBTOTAL: #{subtotal.format}
      HST: #{hst.format}
      TOTAL: #{total.format}

    STRING
  end
end
