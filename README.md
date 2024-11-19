# rubyXL-OutlineTest

rubyXLを使ってExcelのGroupingをしたファイルを出力する実例

## 方法

rubyXLではgithubのissueにある[Outline (Grouping) #272](https://github.com/weshatheleopard/rubyXL/issues/272)に行のグルーピングの実例があるので、
行のグルーピングはその方法で行える。ただ、列のグルーピングは実例がないので調査をする必要がある。

Excelの行と列のGroupingはOffice Open XML(OOXML)ではOutlineという形の実装がされています。
rubyXLではグルーピングはドキュメント化されていませんが、OOXMLに完全準拠しているため、
適切なプロパティを与えることによって、グルーピングを実装することができます。

### 行のグルーピング

```ruby
require 'rubyXL'

workbook = RubyXL::Workbook.new
worksheet = workbook[0]

# 1行目を2行目にまとめる
row1 = worksheet[0]
row1.outline_level = 1
```

デフォルトでは下の行にまとめられるので、上の行にまとめるには以下の設定が追加で必要になる。

```ruby
worksheet.sheet_pr ||= RubyXL::WorksheetProperties.new
worksheet.sheet_pr.outline_pr ||= RubyXL::OutlineProperties.new
worksheet.sheet_pr.outline_pr.summary_below = false # 上側でまとめる
```

### 列のグルーピング

rubyXL、というより、OOXMLでは列は範囲単位でのプロパティのみを持ち、データを持たない。

OOXMLの列プロパティの仕様[ssml:col](http://www.datypic.com/sc/ooxml/e-ssml_col-1.html)には、
列プロパティの一つに、`outlineLevel`というのがあるので、これを設定してやれば良い。

```ruby
# 2列目と3列目をまとめる
cr1 = RubyXL::ColumnRange.new(min: 1, max: 1, width: RubyXL::ColumnRange::DEFAULT_WIDTH)
cr2 = RubyXL::ColumnRange.new(min: 2, max: 3, width: RubyXL::ColumnRange::DEFAULT_WIDTH, outline_level: 1)

# 列プロパティを設定
worksheet.cols.push cr1, cr2
```

列についても、行と同じようにデフォルトでは右側にまとめられるので、左側にまとめるためには次の設定が必要がある。

```ruby
worksheet.sheet_pr ||= RubyXL::WorksheetProperties.new
worksheet.sheet_pr.outline_pr ||= RubyXL::OutlineProperties.new
worksheet.sheet_pr.outline_pr.summary_below = false # 上側でまとめる
```

これらを組み合わせたものが、`grouping.rb`となる。

# Copyright

MIT License
