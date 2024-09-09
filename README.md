日本語の説明は下にあります。（Japanese explanation is the below.）

# Adminer Plugin - Dump XLSX (Excel)

![Adminer Plugin - Dump XLSX (Excel)](img/adminer_xlsx.png "Adminer Plugin - Dump XLSX (Excel)")

Adminer plugin to download selected result as XLSX (Excel).

It is available in the pages of default results and SQL results.  
If multileple selected results are displayed, they are packed in the single file and each result is saved in sheet.

## How to intall plugin

1. Install Adminer PHP plugin.
2. Place `dumpxlsx.js` to the same directory of plugin PHP file.
3. You can change paths to Sheet JS, FileSaver.js and dumpxlsx.js with arguments of constructor of plugin class.

# Adminerプラグイン - XLSX(Excel)ダウンロード

![Adminer Plugin - Dump XLSX (Excel)](img/adminer_xlsx.png "Adminer Plugin - Dump XLSX (Excel)")

SELECT結果をXLSX(Excel)にダウンロードするAdminerプラグインです。

デフォルトの結果ページ、SQLの結果ページで利用可能です。  
複数のSELECT結果がある場合は、それぞれシートに分けて1つのファイルにまとめます。

自分の記事で恐縮ですが、以下の内容を反映したものです。  
https://qiita.com/tomgoodsun/items/0107e5d778b803935fc0

## インストール方法

1. PHPファイルのプラグインをインストールしてください。
2. `dumpxlsx.js`をプラグインと同じディレクトリに配置してください。
3. プラグインクラスのコンストラクタの引数でJSファイルのパスを変更することができます。
