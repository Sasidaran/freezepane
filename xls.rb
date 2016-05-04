#!/usr/bin/env ruby -w -s
# -*- coding: utf-8 -*-
$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

#```ruby
require 'axlsx'
examples = []
examples << :panes



class Array
  def add_style(style)
    return unless map{ |e| e.kind_of? Axlsx::Cell }.uniq.first
    each { |cell| cell.add_style(style) }
  end
end

class Axlsx::Workbook
  attr_accessor :styled_cells

  def add_styled_cell(cell)
    self.styled_cells ||= []
    self.styled_cells << cell
  end

  def apply_styles
    return unless styled_cells
    styled_cells.each do |cell|
      cell.style = styles.add_style(cell.raw_style)
    end
  end
end

class Axlsx::Cell
  attr_accessor :raw_style

  def workbook
    row.worksheet.workbook
  end

  def add_style(style)
    self.raw_style ||= {}
    self.raw_style = raw_style.merge(style)
    workbook.add_styled_cell(self)
  end
end



header = ["2015","mars","02","650922281","530182885","145915994","-25176598","144983684","963552887","-578353686","21045393","124870601","-3313714","-21862884","-25437976","201850722","235508019","14666558","28387790","37385805","-342617420","-4759814","46174160","-18081019","-31508378","-53236842","5948841","16443828","83284122","478919231","175817413","11120941","74817621","173852969","-","-589950983","11597297","-","-","282144","9952688","6949926","-9647237","13507872","-53798828","-9848280","331350906","-347488191","-250813267","349547181","274705217","-2629346159","-11008454","10855981","2276256630","4933177","723735","21503438","277828","-12810527","5518403","-2179551","-27859425","-26511461","220862248","-3313714","-6524155","-3414952","-1382740","-","1088057","-4821136","-3447782","-3360176","-"]

row_1 = ["2015","mars","02","650922281","530182885","145915994","-25176598","144983684","963552887","-578353686","21045393","124870601","-3313714","-21862884","-25437976","201850722","235508019","14666558","28387790","37385805","-342617420","-4759814","46174160","-18081019","-31508378","-53236842","5948841","16443828","83284122","478919231","175817413","11120941","74817621","173852969","-","-589950983","11597297","-","-","282144","9952688","6949926","-9647237","13507872","-53798828","-9848280","331350906","-347488191","-250813267","349547181","274705217","-2629346159","-11008454","10855981","2276256630","4933177","723735","21503438","277828","-12810527","5518403","-2179551","-27859425","-26511461","220862248","-3313714","-6524155","-3414952","-1382740","-","1088057","-4821136","-3447782","-3360176","-"]
sub_row = ["","Jan Total","","650922281","530182885","145915994","-25176598","144983684","963552887","-578353686","21045393","124870601","-3313714","-21862884","-25437976","201850722","235508019","14666558","28387790","37385805","-342617420","-4759814","46174160","-18081019","-31508378","-53236842","5948841","16443828","83284122","478919231","175817413","11120941","74817621","173852969","-","-589950983","11597297","-","-","282144","9952688","6949926","-9647237","13507872","-53798828","-9848280","331350906","-347488191","-250813267","349547181","274705217","-2629346159","-11008454","10855981","2276256630","4933177","723735","21503438","277828","-12810527","5518403","-2179551","-27859425","-26511461","220862248","-3313714","-6524155","-3414952","-1382740","-","1088057","-4821136","-3447782","-3360176","-"]
sub_row_1 = ["2015","Jan","03","650922281","530182885","145915994","-25176598","144983684","963552887","-578353686","21045393","124870601","-3313714","-21862884","-25437976","201850722","235508019","14666558","28387790","37385805","-342617420","-4759814","46174160","-18081019","-31508378","-53236842","5948841","16443828","83284122","478919231","175817413","11120941","74817621","173852969","-","-589950983","11597297","-","-","282144","9952688","6949926","-9647237","13507872","-53798828","-9848280","331350906","-347488191","-250813267","349547181","274705217","-2629346159","-11008454","10855981","2276256630","4933177","723735","21503438","277828","-12810527","5518403","-2179551","-27859425","-26511461","220862248","-3313714","-6524155","-3414952","-1382740","-","1088057","-4821136","-3447782","-3360176","-"]
p = Axlsx::Package.new
wb = p.workbook


months  = ["Jan", "Feb", "Mar", "Apr", "May" , "june", "July", "Aug", "Sep", "Oct", "Nov", "Dec"]


## Frozen/Split panes
## ``` ruby
if examples.include? :panes
  wb.add_worksheet(:name => 'panes') do |sheet|
    title = wb.styles.add_style( :bg_color => "CCE5FF0",  :sz=>12,  :border=> {:style => :thin, :color => "444444"},
                                 :alignment => { :horizontal => :center,:vertical => :center , :wrap_text => true})

    sheet.add_row(header, :style=>title, :height => 30)
    #sheet.column_info.first.width = 30

    #sheet.add_row(['Jan total', "=SUM(B3:B5)", "=SUM(c3:c5)", "=SUM(d3:d5)", "=SUM(e3:e5)", "=SUM(f3:f5)"])
   
    3.times do
     sheet.add_row(row_1)
    end


    cal =  ("A".."BW").to_a.collect{|x| "=SUM(#{x}6" ":" "#{x}7)" }
    cal[0],cal[1],cal[2] = "2015 Total","",""
    sheet.add_row(cal)
    
    months.each do |x|
     sub_row[1] = x
     sheet.add_row(sub_row)
     count = 0
     20.times do
        count += 1
        sub_row_1[0],sub_row_1[1], sub_row_1[1] = "","", count
	      sheet.add_row(sub_row_1)
     end
    end


    cal =  ("A".."BW").to_a.collect{|x| "=SUM(#{x}2" ":" "#{x}4)" }
    cal[0],cal[1],cal[2] = "Grand Total","",""
    sheet.add_row(cal)

  sheet["A1:BW1"].add_style(b: true)
  sheet["A1:BW1"].add_style(:alignment => { :horizontal => :center,:vertical => :center , :wrap_text => true})
  sheet["A1:D1"].add_style(bg_color: "FF9999")
  sheet["D1:F1"].add_style(bg_color: "CCE5FF0")
  sheet["G1:I1"].add_style(bg_color: "3399FF")
  sheet["J1:L1"].add_style(bg_color: "FFFF99")
  sheet["M1:O1"].add_style(bg_color: "66FF66")
  sheet["P1:BW1"].add_style(bg_color: "FF66FF")


    sheet.sheet_view.pane do |pane|
       pane.top_left_cell = "C2"
       pane.state = :frozen_split
       pane.y_split = 1
       pane.x_split = 2
       pane.active_pane = :bottom_right
     end

    sheet.rows[1..257].each do |row|
      #sheet.rows[2].hidden = true
      row.outline_level = 1
      #row.hidden = true
    end


    sheet.rows[5..257].each do |row|
      row.outline_level = 2
      #row.hidden = true
    end

    sheet.rows[6..254].each do |row|
      row.outline_level = 3
      row.hidden = true
    end

    sheet.rows[27..46].each do |row|
      row.outline_level =4
      row.hidden = true
    end 
    
   sheet.rows[6..26].each do |row|
      row.outline_level =4
      row.hidden = true
    end 

   sheet.rows[48..67].each do |row|
      row.outline_level = 4
      row.hidden = true
    end 

   sheet.rows[69..88].each do |row|
      row.outline_level = 4
      row.hidden = true
   end 

   sheet.rows[90..109].each do |row|
      row.outline_level = 4
      row.hidden = true
   end 

   sheet.rows[111..130].each do |row|
      row.outline_level = 4
      row.hidden = true
   end 

   sheet.rows[132..151].each do |row|
      row.outline_level = 4
      row.hidden = true
   end 
   sheet.rows[153..172].each do |row|
      row.outline_level = 4
      row.hidden = true
   end 

   sheet.rows[174..193].each do |row|
      row.outline_level = 4
      row.hidden = true
   end 


   sheet.rows[195..214].each do |row|
      row.outline_level = 4
      row.hidden = true
   end 

   sheet.rows[216..235].each do |row|
      row.outline_level = 4
      row.hidden = true
   end 

   sheet.rows[237..256].each do |row|
      row.outline_level = 4
      row.hidden = true
   end 
    
   sheet.rows[1].hidden = true
   
  #sheet.column_info[4].hidden = true

    sheet.sheet_view do |view|
      view.show_outline_symbols = true
    end
  end
end
wb.apply_styles

p.serialize("pane_latest.xlsx")