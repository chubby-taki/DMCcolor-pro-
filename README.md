# DMC Color Matcher Pro

写真から、いちばん近いDMC刺繍糸の色番号を調べるツール。
リアルペット刺繍の色選びのために作りました。

https://chubby-taki.github.io/DMCcolor-pro-/

どなたでも無料で使えます。登録も入力も要りません。

## できること

- 写真を読み込んで、拡大しながら色を拾う（毛の1本まで狙える）
- 拾った色に近いDMC糸を5色まで提示（CIEDE2000で色差を計算）
- 選んだ糸を一覧に貯める（Color Legend）
- DMC色番号の収納位置を表示（列-行）
- 一覧をCSV／PDFで書き出す

## 使い方

1. 写真を選ぶ。うまく撮れていなくても大丈夫
2. 色を知りたいところをタップ（またはクリック）
3. 近い糸が5色出るので、使うものを選ぶ
4. 一覧に貯まったら、CSVかPDFで持ち出す

## プライバシー

写真はブラウザの中だけで処理されます。サーバーには送信されません。
外部への通信は、同梱の `dmc_master_data.json` と `dmc_position_data.json` の読み込みのみです。

## 技術

HTML / CSS / Vanilla JavaScript のみ。ビルド不要。
`main` ブランチを GitHub Pages がそのまま公開しています。

- HTML5 Canvas
- CIEDE2000 色差計算（`ciede2000.js`）
- jsPDF・FileSaver（CDN読み込み。失敗しても本体は動く）

## ライセンス

Private - Real Pet Embroidery / takis
