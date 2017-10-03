# encoding: UTF-8
# frozen_string_literal: true

require "benchmark"
Bundler.require :default, :development

def path_to_excel_file
  ARGV[0]
end

def try_creek
  Creek::Book.new(path_to_excel_file).sheets[0].rows.map { |row| row.values.to_a }
end

def try_ooxl
  doc = OOXL.open(path_to_excel_file)
  doc.sheet(doc.sheets[0]).rows.map do |row|
    row.cells.map { |c| c.value }
  end
end

def try_rubyxl
  RubyXL::Parser.parse(path_to_excel_file)[0].sheet_data.rows.each_with_object [] do |row, memo|
    if row
      memo << row.cells.map { |c| c.respond_to?(:value) ? c.value : nil }
    end
  end
end

def try_roo
  doc   = Roo::Spreadsheet.open(path_to_excel_file)
  sheet = doc.sheet(0)
  (sheet.first_row..sheet.last_row).each_with_object [] do |i, memo|
    memo << sheet.row(i)
  end
end

def try_simple_xlsx_reader
  SimpleXlsxReader.open(path_to_excel_file).sheets[0].rows
end

def try_simple_spreadsheet
  doc                = SimpleSpreadsheet::Workbook.read(path_to_excel_file)
  doc.selected_sheet = doc.sheets[0]
  [].tap do |memo|
    doc.first_row.upto doc.last_row do |i|
      memo << doc.first_column.upto(doc.last_column).map { |j| doc.cell(i, j) }
    end
  end
end

def available_gems
  [:creek, :ooxl, :rubyxl, :roo, :simple_xlsx_reader, :simple_spreadsheet]
end

Benchmark.bmbm available_gems.map(&:to_s).map(&:size).max do |bm|
  available_gems.shuffle.each do |gem|
    bm.report(gem) { send(:"try_#{gem}") }
  end
end
