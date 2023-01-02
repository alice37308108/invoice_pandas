## やりたいこと
Excelの取引先一覧シートにこのようなデータがあります。
![company](img/company.png)  

登録されている取引先についてWeb-APIを使用して、インボイス登録事業者として登録されていたら、登録番号等をExcelに出力します。
![invoice](img/invoice.png)  

## 動作の流れ
* Excelの取引先一覧シートにあるデータをPandasデータフレームとして取得
* 日本郵便から最新の郵便番号をダウンロードして郵便番号と市区町村コードが入ったデータフレームを作成
* 取引先と郵便番号のデータフレームについて、郵便番号をキーとしてマージする
* 市区町村コードが存在しない（郵便番号が日本郵便の郵便番号と一致しない）ときはエラー.csvに出力
* マージしたデータフレームをもとに法人番号システム Web-APIを使用して、法人番号等を取得
* 法人番号システム Web-APIから取得した法人番号をもとに適格請求書発行事業者公表システムWeb-APIからインボイス登録番号を取得
* 結果をExcelに出力する

## 注意事項
* 法人番号システム Web-APIおよび適格請求書発行事業者公表システムWeb-APIを使用するにはアプリケーションIDが必要です。
* Excelの取引先データの法人名は全角で登録してください。

## 参考
* [法人番号システム Web-API](https://www.houjin-bangou.nta.go.jp/webapi/)
* [適格請求書発行事業者公表システムWeb-API](https://www.invoice-kohyo.nta.go.jp/)
