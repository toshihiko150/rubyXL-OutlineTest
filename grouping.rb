require "rubyXL"
require "rubyXL/convenience_methods"

# テストデータを作成
workbook = RubyXL::Workbook.new
worksheet = workbook[0]
0.upto(10).each_with_index do |col_index|
  ("A".."J").each_with_index do |v, row_index|
    worksheet.add_cell(row_index, col_index, v)
  end
end

# 1行目に2行目と3行目をまとめる
worksheet[1].outline_level = 1
worksheet[2].outline_level = 1

# 4行目に5行目と6行目をまとめる
worksheet[4].outline_level = 1
worksheet[5].outline_level = 1

# 7行目に8行目から10行目をまとめる
worksheet[7].outline_level = 1
worksheet[8].outline_level = 1
worksheet[9].outline_level = 1

# 1列目に2列目と3列目をまとめる
cr1 = RubyXL::ColumnRange.new(min: 1, max: 1, width: RubyXL::ColumnRange::DEFAULT_WIDTH)
cr2 = RubyXL::ColumnRange.new(min: 2, max: 3, width: RubyXL::ColumnRange::DEFAULT_WIDTH, outline_level: 1)

# 4列目に5列目から7列目をまとめる
cr3 = RubyXL::ColumnRange.new(min: 4, max: 4, width: RubyXL::ColumnRange::DEFAULT_WIDTH)
cr4 = RubyXL::ColumnRange.new(min: 5, max: 7, width: RubyXL::ColumnRange::DEFAULT_WIDTH, outline_level: 1)
worksheet.cols.push cr1, cr2, cr3, cr4

# 8列目に9列目から11列目をまとめる
cr5 = RubyXL::ColumnRange.new(min: 8, max: 8, width: RubyXL::ColumnRange::DEFAULT_WIDTH)
cr6 = RubyXL::ColumnRange.new(min: 9, max: 11, width: RubyXL::ColumnRange::DEFAULT_WIDTH, outline_level: 1)
worksheet.cols.push cr1, cr2, cr3, cr4, cr5, cr6

# まとめる列をどちら側にするかを指定する（行は上下、列は左右）
# デフォルトは下側と右側
worksheet.sheet_pr ||= RubyXL::WorksheetProperties.new
worksheet.sheet_pr.outline_pr ||= RubyXL::OutlineProperties.new
worksheet.sheet_pr.outline_pr.summary_below = false # 上側でまとめる
worksheet.sheet_pr.outline_pr.summary_right = false # 左側でまとめる

# データ出力
workbook.write("output.xlsx")
