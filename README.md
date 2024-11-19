# rubyXL-OutlineTest

rubyXLを使ってExcelのGroupingをしたファイルを出力する実例

## 方法

rubyXLではgithubのissueにある[Outline (Grouping) #272](https://github.com/weshatheleopard/rubyXL/issues/272)に行のグルーピングの実例があるので、
行のグルーピングはその方法で行える。ただ、列のグルーピングは実例がないので調査をする必要がある。

Excelの行と列のGroupingはOffice Open XML(OOXML)ではOutlineという形の実装がされています。
rubyXLではグルーピングはドキュメント化されていませんが、OOXMLに準拠しているため、
適切なプロパティを与えることによって、グルーピングを実装することができます。

### 行のグルーピング

```ruby
require 'rubyXL'

workbook = RubyXL::Workbook.new
worksheet = workbook[0]

# 2行目をまとめる
row2 = worksheet[1]
row2.outline_level = 1
```

デフォルトでは下の行にまとめられるので、上記の例だと2行目が3行目にまとめられる。  
上の行にまとめるには以下の設定が追加で必要になる。

```ruby
worksheet.sheet_pr ||= RubyXL::WorksheetProperties.new
worksheet.sheet_pr.outline_pr ||= RubyXL::OutlineProperties.new
worksheet.sheet_pr.outline_pr.summary_below = false # 上側でまとめる
```

### 列のグルーピング

rubyXL、というより、OOXMLでは列はプロパティのみを持ち、データを持たない。  
列のプロパティは範囲でプロパティを定義される。

OOXMLの列プロパティの仕様[ssml:col](http://www.datypic.com/sc/ooxml/e-ssml_col-1.html)には、
列プロパティの一つに、`outlineLevel`というのがあるので、これを設定してやれば良い。

```ruby
# 2列目と3列目をまとめる
cr = RubyXL::ColumnRange.new(min: 2, max: 3, width: RubyXL::ColumnRange::DEFAULT_WIDTH, outline_level: 1)

# 列プロパティを設定
worksheet.cols.push cr
```

列についても、行と同じようにデフォルトでは右側にまとめられるので、左側にまとめるためには次の設定が必要がある。

```ruby
worksheet.sheet_pr ||= RubyXL::WorksheetProperties.new
worksheet.sheet_pr.outline_pr ||= RubyXL::OutlineProperties.new
worksheet.sheet_pr.outline_pr.summary_below = false # 上側でまとめる
```

これらを組み合わせたものが、`grouping.rb`となる。

# License

MIT License
