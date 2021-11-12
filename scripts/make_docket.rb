require 'docx'
require 'faker'
require 'sablon'

doc = Docx::Document.open("./docket.docx")
table = doc.tables[0]
docket = []
table.rows.each_with_index do |row, i|
  if i == 0
    entry = {
      date: row.cells[0].to_s,
      file: row.cells[1].to_s,
      desc: row.cells[2].to_s,
      time: row.cells[3].to_s
    }
  else
    entry = {
      date: row.cells[0].to_s,
      file_internal: row.cells[1].paragraphs.first.to_s,
      file_external: "Work for " + Faker::Name.name,
      desc: Faker::Lorem.words((Random.rand * 20).to_i + 1).join(' '),
      time: row.cells[3].to_s
    }
  end
  docket << entry
end

template = Sablon.template(File.expand_path("./template.docx"))
context = { docket: docket }
template.render_to_file File.expand_path("./out.docx"), context
