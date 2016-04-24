#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

#```ruby
require 'axlsx'
examples = []
examples << :panes

p = Axlsx::Package.new
wb = p.workbook

## Frozen/Split panes
## ``` ruby
if examples.include? :panes
  wb.add_worksheet(:name => 'panes') do |sheet|
    sheet.add_row(['',  (0..5).map { |i| "column header #{i}" }].flatten )
    #sheet.add_row(['Jan total', "=SUM(B3:B5)", "=SUM(c3:c5)", "=SUM(d3:d5)", "=SUM(e3:e5)", "=SUM(f3:f5)"])
    sheet.add_row(['Jan', (1..6).map { |i| "#{i}" }].flatten)
    sheet.add_row(['Jan', (7..12).map { |i| "#{i}" }].flatten)
    sheet.add_row(['Jan', (12..17).map { |i| "#{i}" }].flatten)
    sheet.add_row(['Jan total', "=SUM(B2:B4)", "=SUM(c2:c4)", "=SUM(d2:d4)", "=SUM(e2:e4)", "=SUM(f2:f4)"])
    #sheet.add_row(['Feb total', "=SUM(B7:B8)", "=SUM(c7:c8)", "=SUM(d7:d8)", "=SUM(e7:e8)", "=SUM(f7:f8)"])
    sheet.add_row(['Feb', (23..28).map { |i| "#{i}" }].flatten)
    sheet.add_row(['Feb', (28..32).map { |i| "#{i}" }].flatten)
    sheet.add_row(['Feb total', "=SUM(B6:B7)", "=SUM(c6:c7)", "=SUM(d6:d7)", "=SUM(e6:e7)", "=SUM(f6:f7)"])
    sheet.sheet_view.pane do |pane|
      pane.top_left_cell = "B2"
      pane.state = :frozen_split
      #pane.y_split = 1
      pane.x_split = 2
      pane.active_pane = :bottom_right
    end
    sheet.rows[1..3].each do |row|
    #sheet.rows[2..4].each do |row|
      row.outline_level = 1
      row.hidden = true
    end
    sheet.rows[5..6].each do |row|
    #sheet.rows[6..7].each do |row|
      row.outline_level = 1
      row.hidden = true
    end
    sheet.sheet_view do |view|
      view.show_outline_symbols = true
    end
  end
end

p.serialize("pane.xlsx")