---
name: proposal-deck
description: |
  既存のPowerPoint提案書の体裁（配色・フォント・ヘッダ/フッタ・レイアウト・Appendix）をそのまま引き継いで、
  新しい案件の提案書デッキを作る。既定はVRAIN Solutionの営業提案書（AI外観検査など）。
  任意のテンプレpptxを渡された場合は、そのデッキから設計トークンを抽出して流用する汎用モードで動く。
  Use this skill whenever the user asks for 提案書・ご提案資料・スライド・デッキ・プレゼン資料 to be created or updated,
  whenever they paste 案件情報・ヒアリングメモ・議事録 and want it turned into slides,
  whenever they say「他依頼みたいに作って」「前回と同じ体裁で」「この資料をベースに」,
  and whenever a .pptx is attached as a reference or template — even if they never say the word "skill" or "template".
  営業提案書・見積前提資料・顧客向け説明資料はすべてこのスキルの対象。
  Do NOT use for: 単発のグラフ作成のみ、Googleスライドの編集、Word/PDF成果物。
---

# 提案書デッキ作成

既存デッキを**設計システムの供給源**として扱い、中身だけ入れ替える。ゼロから作らない理由は単純で、
顧客に出す資料は体裁が1pxでもブレると「別の会社の資料」に見えるから。マスター・レイアウト・
Appendixはテンプレのものを原子単位で再利用し、こちらが書くのは本編スライドのXMLだけにする。

## スクリプトの場所について

以下 `$SK` は**このSKILL.mdがあるディレクトリ**（読み込み時に表示される "Base directory for this skill"）。
リポジトリ同梱なら `.claude/skills/proposal-deck`、プロフィールに保存した場合は別の場所になるので、
最初に `SK=<そのパス>` を決めてから以降のコマンドを打つ。

`add_slide` `clean` `validate` `thumbnail` は同梱の `pptx` スキルのものを使うが、その置き場所は
セッションごとにUUID入りのパスに変わる。**直接パスを書かず、必ず `scripts/pptx_tool.py` 経由で呼ぶ。**

## 0. 環境を整える（最初に必ず）

このコンテナには **LibreOffice Impress・poppler・CJKフォント・markitdown・python-pptx が入っていない**。
入れずに進めると「source file could not be loaded」で詰まるので、最初に流す:

```bash
bash $SK/scripts/setup_env.sh
```

数分かかる。待っている間にテンプレの中身を読む作業を進めてよい。

## 1. テンプレを手に入れる

優先順:

1. `assets/` に同梱のテンプレがあればそれを使う（`ls $SK/assets/*.pptx`）
2. ユーザーが添付したpptx
3. どちらも無ければ**聞く**。「直近の提案書pptxを添付してください」。推測で自作しない

## 2. テンプレを読む

作業用に展開しておく（以降のスライド操作はすべてこの `unpacked/` に対して行う）:

```python
import sys; sys.path.insert(0, "<$SK>/scripts")
import deck_lib as d
d.unpack("<template.pptx>", "unpacked")
```

```bash
python $SK/scripts/inspect_template.py <template.pptx>
```

スライドごとの「使用レイアウト・図形名・座標(EMU)・塗り色・テキスト」と、本文プレースホルダの
idx対応が出る。加えて `markitdown <template.pptx>` で全文、
`python $SK/scripts/pptx_tool.py thumbnail <template.pptx> tpl` でサムネイルを見る。

VRAIN以外のテンプレ（スライドサイズが違う）なら、生成前に必ず:

```python
d.configure_for("unpacked")   # L/W/FOOT をスライドサイズから再計算
```

`from deck_lib import *` は使わない。star importした `L` `W` は値のコピーなので、
`configure_for` で更新されず古い座標のまま生成してしまう。常に `d.L` のようにモジュール経由で参照する。

ここで決めるのは3つ:

- **どのレイアウトが本編スライドか**（VRAINテンプレでは `slideLayout5`。タイトル=idx12 / リード文=idx13 のプレースホルダを供給）
- **どのスライドが流用対象か**（会社概要・導入事例・プロダクト紹介などのAppendixは**触らず丸ごと残す**）
- **どのスライドが案件固有か**（前案件の検証結果・写真つきスライドは捨てる）

設計トークン（色・フォント・座標）は `references/design-system.md` に既出のVRAIN値がある。
別テンプレなら `inspect_template.py` の出力から同じ形で起こす。手順は `references/generic-template.md`。

## 3. 章立てを決める

`references/deck-structure.md` に標準構成と各スライドの型（2カード型・フロー型・表型・ガント型・
ステップ型・効果型）がある。案件情報をそこに流し込む。

原則として **1スライド1メッセージ**。リード文（idx13）は「■〜いたします。」の一文で、
そのスライドで言いたいことを言い切る。

## 4. 本編スライドを生成する

`scripts/deck_lib.py` がテンプレの作法そのままのXMLを吐く。使い方:

```python
i = d.ids()
(x1, w1), (x2, w2) = d.cols(2)        # 座標は直書きせず cols() から取る
b  = d.card(i, x1, d.TOP, w1, 2926080, "＜案件のご状況＞", ["・…", "・…"])
b += d.card(i, x2, d.TOP, w2, 2926080, "＜課題＞", ["・…"], header_fill=d.RED)
b += d.banner(i, d.L, 4892040, d.W, 868680, [d.R("⇒ …", sz=1300, b=1, color=d.NAVY)])
d.build("unpacked/ppt/slides/slide4.xml", "現状のご状況と課題", "■…いたします。", b)
```

揃っているもの: `card` `banner` `point_row` `flow` `gantt` `table` `header_row` `zebra` `tcell` `shape` `textbox` `R` `P` `ids` `cols` `build`、
色定数 `NAVY/RED/BLUE/LTBLUE/LTGRAY/GRAY/BORDER/MIDGRAY/WHITE`、
版面定数 `L`（左端）`W`（本文幅）`TOP`（本文上端）。詳細は `references/design-system.md`。

表紙・目次・中扉はゼロから作らず、テンプレの該当スライドの**文字だけ差し替える**:

```python
d.edit_slide("unpacked", "slide1",   # 段落を1runに潰して置換
             lambda x: d.set_para(x, "ヤマキ株式会社　御中", "〇〇株式会社　御中"))
d.edit_slide("unpacked", "slide2",   # run単位の入れ替え（同時置換なので A→B, B→C も安全）
             lambda x: d.swap_runs(x, [("検証結果（再掲）", "ご検討にあたっての比較観点")]))
```

スライドの書き換えは `edit_slide` を通す。`open(path, "w")` を先に開いてから置換すると、
置換が失敗した瞬間に中身が空のスライドが残り、デッキ全体が開けなくなる。

ページ番号は `<a:fld type="slidenum">` なので触らなくていい。並べ替えれば自動で振り直る。

## 5. 組み立てる

```bash
# 本編スライドが足りなければ複製（package bookkeeping込み）
python $SK/scripts/pptx_tool.py add_slide unpacked/ slide4.xml
```

並び順と不要スライドの削除は `<p:sldIdLst>` の書き換えで行う:

```python
d.reorder_slides("unpacked", ["slide1","slide2","slide3","slide4", ..., "slide39"])
```

その後 `clean.py` で孤児（消したスライドが抱えていた画像）を回収してから梱包:

```bash
python $SK/scripts/pptx_tool.py clean unpacked/
```
```python
d.pack("unpacked", "out.pptx")   # 既存のout.pptxは消してから詰める（消したパーツが残らないように）
```

**順序が大事**: 構造操作（複製・削除・並べ替え）を全部終えてから中身を書く。
`add_slide.py` は複製元をそのままコピーするし、`clean.py` は `sldIdLst` に無いスライドを消す。

## 6. 検証する

```python
d.check_bounds("unpacked")      # 何も出なければOK
```
```bash
python $SK/scripts/pptx_tool.py validate out.pptx --original <template.pptx>
# 前案件の残骸チェック。OLD にはテンプレ表紙の社名・件名・品名を入れる（例: "ヤマキ|花かつお|かつお節"）
OLD="ヤマキ|花かつお"
markitdown out.pptx > out.md || { echo "!! markitdown failed — デッキが壊れている"; exit 1; }
grep -inE "lorem|ipsum|TODO|\[insert|$OLD" out.md && echo "!! 残骸あり" || echo "clean"
```

markitdownを先にファイルへ書き出すのは、パイプにすると変換失敗が `grep` の「該当なし」と区別できず、
壊れたデッキを「clean」と誤判定するから。`check_bounds` は何も出なければOK。スライド外にはみ出した図形はエラーにならず黙って見切れるので、
ここで機械的に潰しておく。`--original` を付けるのは、テンプレ自体が持っているスキーマ違反を
自分のエラーと取り違えないため。

## 7. 目で見る（省略不可）

```bash
bash $SK/scripts/render_qa.sh out.pptx outimg
```

出た画像を**全部**見る。初回は必ず数枚おかしい。特に:

- **文字のはみ出し・見切れ**（最頻出。日本語は全角1文字≒フォントpt。11ptなら幅4,000,000EMUに約28文字）
- カードの下が間延びしている（箇条書きを足すか、箱を縮める）
- 図形内の2行ラベルが枠を割っている（段落を分ける／文言を短くする）
- 隣の要素との距離が0.3インチ未満
- テンプレ由来の装飾がテキスト差し替えでズレた

直したら**再梱包→再レンダリング**（PDFを作り直さないと画像は変わらない）。

## 8. 渡す

ファイル名は `YYYYMMDD_<件名>_ご提案書.pptx`。`SendUserFile` で渡し、本文には章立ての表と
**確認が必要な箇所**を明記する。

## 数字と固有名詞の扱い

提案書は顧客に出る。だから:

- **もらっていない数値を断定で書かない。** 処理時間の内訳のように推定が必要なら、スライド上に
  「当社想定」「実機構成確定後に実測にて再提示」と明記し、返信でも推定値だと伝える
- **社名・担当者名が未確定なら `〇〇株式会社　御中` のままにして、返信の冒頭で指摘する。**
  推測で埋めると表紙が間違ったまま客先に出る
- 金額は原則スライドに載せない（見積書側の領域）。載せる方針なら明示的に確認する
- 前案件の社名・品名・写真が残っていないか、§6のgrepで必ず確認する

## 参照ファイル

| ファイル | 読むとき |
|---|---|
| `references/design-system.md` | 色・座標・フォント・レイアウト対応の実値が要るとき |
| `references/deck-structure.md` | 章立てとスライドの型を決めるとき |
| `references/generic-template.md` | VRAIN以外のテンプレを渡されたとき |
