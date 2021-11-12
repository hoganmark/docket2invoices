class WorkItem
  attr_accessor :date, :description, :time_spent, :rate

  def initialize(date, description, time_spent, rate)
    self.date = date
    self.description = description
    self.time_spent = time_spent
    self.rate = rate
  end

  def amount
    Money.new(time_spent * rate * 100)
  end

  def to_s
    "#{date}: #{description} (#{time_spent} hours, #{amount.format})"
  end
end
