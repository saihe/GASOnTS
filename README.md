# GASをTSで書く第一弾

勤務表の初期化をGASで記載したがTSで書きたくなったため、こちらで作業。

## 開発手順

1. claspを使用可能にする

    ``` bash
      npm install -g @google/clasp
    ```

2. .clasp_template.jsonを.clasp.jsonという名前でコピーする
3. .clasp.jsonのscriptIdにGASのIDを転記する。ただし、当該GAS_IDは非公開。
4. コードを書く
5. clasp pushして反映する
  ※TSは勝手にGASにトランスパイルされる。

## 休日に関して

`src/holiday.js` というファイルを用意してるが、もともとのファイル有(<https://github.com/osamutake/japanese-holidays-js/blob/master/lib/japanese-holidays.min.js>)年度代わりなどリリース状況を確認してファイルを最新化する必要がある。
