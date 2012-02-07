test_hash = {"sheet1"=>
                 {"Row1"=>
                      {"A1"=>"01:25:30 PM",
                       "B1"=>"0.154251449887584",
                       "C1"=>"150",
                       "D1"=>"2012-01-28",
                       "E1"=>"12/25/2012 12:12:30 PM",
                       "F1"=>"12/03/2012",
                       "G1"=>"01:00:45 AM"},
                  "Row2"=>{"A2"=>"1", "B2"=>"2", "C2"=>"3", "E2"=>"4", "F2"=>"5", "G2"=>"6"}
                 },
             "sheet2"=>{"Row1"=>{"A1"=>nil, "D1"=>nil, "E1"=>nil, "F1"=>nil, "G1"=>nil}},
             "sheet3"=>{}}

require "rspec"

require_relative 'spec_helper'
require_relative '../lib/excelx_preview'
describe ExcelX do
  describe "#preview" do
    it "it should preview first 10 lines of each sheets" do
      ExcelX::Previewer.preview("data/formulas.xlsx").should == test_hash
    end
  end
end
