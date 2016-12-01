# coding:utf-8
require 'rubyXL'

# 新しいworkbookの作成
workbook = RubyXL::Workbook.new

# worksheetを取得
# デフォルトで1つ生成されている
worksheet = workbook[0]
# フォントのデフォルト指定
workbook.fonts[0].set_name('梅明朝')
workbook.fonts[0].set_size(14)

# シートの名前を変更
worksheet.sheet_name = '新しいシート名'

# セルに文字を追加
cell = worksheet.add_cell(0, 1, 12345)

# 幅を広げる
worksheet.change_column_width(1, 20)

# セルの下に罫線を引く
cell.change_border(:bottom, 'medium')

# 細い罫線を引く
# チェインメソッドな指定も可能
#cell = worksheet.add_cell(0, 3, '')
#                .change_border(:bottom, 'thin')
# 書式設定
# see https://support.office.com/en-us/article/5026bbd6-04bc-48cd-bf33-80f18b4eae68
cell.set_number_format '#,##0'

#### スタイル関連 ####
# フォントを変更
cell.change_font_name '梅明朝'

# フォントサイズを変更
cell.change_font_size 20

# フォントの色を変更
cell.change_font_color 'ff0000'

# セルの色を変更
cell.change_fill '00ff00'
######################

# 2シート目を追加
worksheet = workbook.add_worksheet('次のシート')

## OUTPUT
# sample.xlsxという名前で保存
workbook.write('sample.xlsx')
