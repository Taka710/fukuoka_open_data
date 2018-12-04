## 福岡オープンデータの可視化

福岡のオープンデータをTableau Publicに表示するための
元データを作成するためのプログラムです。

完成したTableau Publicは以下のアドレスとなります。
https://public.tableau.com/profile/takahiro.hama8517#!/vizhome/populationofFukuokaprefecture/fukuokapopulation


### フォルダについて
xlsx: 福岡県のオープンデータサイトから取得したxlsファイルを配置する
json: google spreadに登録するためのjsonファイルを配置

### 定数について

GOOGLE_SPREAD_NAME: googleスプレッド名を設定
POPULAR_SPREAD: 人口情報を作成するシート名を設定
AREA_SPREAD: 面積情報を作成するシート名を設定

AREA_URL: 面積データを取得するURLを設定
今回は以下のサイトからデータを取得させていただきました。
https://uub.jp

※google spreadに書き込みを行うためのAPI設定が別途必要です。


以下のコマンドを実行すれば、google spreadにデータを作成できます。

```
python main.py
```