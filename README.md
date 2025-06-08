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

## Comptible with Admine Neo

1. Place `dumpxlsx.js` to the web directory, for example, the same directory of the `adminneo.php`.
2. Place the `adminneo-config.php` in the same directory of `adminneo.php` and write the following.

```php
<?php
return [
    'jsUrls' => [
        'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.16.9/xlsx.full.min.js',
        'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js',
        'dumpxlsx.js',
    ],
];
```

For more details, check [Admine Neo Configuration page](https://www.adminneo.org/configuration#jsUrls).

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

## Admine Neo互換

1. `dumpxlsx.js`をウェブディレクトリに配置してください。例えば、`adminneo.php`と同じディレクトリなどに配置してください。
2. `adminneo-config.php`を作成し、以下のように記述してください。

```php
<?php
return [
    'jsUrls' => [
        'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.16.9/xlsx.full.min.js',
        'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js',
        'dumpxlsx.js',
    ],
];
```

詳細は[Admine Neo Configurationページ](https://www.adminneo.org/configuration#jsUrls)を参照してください。

