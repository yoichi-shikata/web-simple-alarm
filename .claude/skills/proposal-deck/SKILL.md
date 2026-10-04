---
name: proposal-deck
description: |
  既存のPowerPoint提案書の体裁（配色・フォント・ヘッダ/フッタ・レイアウト・Appendix）をそのまま引き継いで、
  新しい案件の提案書デッキ(.pptx)を作る。既定はVRAIN Solutionの営業提案書（AI外観検査など）で、
  同梱の正解例と同じ構成・同じ書き方・同じ座標で組む。別のテンプレpptxを渡された場合は汎用モードで流用する。
  Use this skill whenever the user asks for 提案書・ご提案資料・スライド・デッキ・プレゼン資料 to be created or updated,
  whenever they paste 案件情報・ヒアリングメモ・議事録 and want it turned into slides,
  whenever they say「他依頼みたいに作って」「前回と同じ体裁で」「この資料をベースに」,
  and whenever a .pptx is attached as a reference or template — even if they never say "skill" or "template".
  Do NOT use for: 単発のグラフ作成のみ、Googleスライドの編集、Word/PDF成果物。
---

# 提案書デッキ作成（Claude Code / Codex 共通）

既存デッキを**設計システムの供給源**として扱い、中身だけ入れ替える。顧客に出す資料は体裁が少しでも
ブレると「別の会社の資料」に見えるので、マスター・レイアウト・Appendixはテンプレのものをそのまま使い、
こちらが書くのは本編スライドのXMLだけにする。

このスキルは **Claude Code と Codex で同じ出力になる**ことを目標にしている。そのために:

- ツールは全部このフォルダ内のスクリプト（標準ライブラリ＋Pillow/PyMuPDF）。エージェント固有の機能に依存しない
- 構成・文言・座標は `examples/two-station-line/build_deck.py`（**正解例**）で固定している。
  自分で一からレイアウトを考えず、正解例をコピーして `CONTENT` を書き換えるのが基本

## パスの決め方

以下 `$SK` は**このSKILL.mdがあるディレクトリ**。最初に確定させてから進める
（例: Claude Code のリポジトリ同梱なら `.claude/skills/proposal-deck`、Codex なら `~/.agents/skills/proposal-deck` や
`.agents/skills/proposal-deck`）。

Windows（PowerShell）では `~` が展開されないことがあるので、変数に入れてから使う:

```powershell
$SK = "$env:USERPROFILE\.agents\skills\proposal-deck"
python "$SK\scripts\setup_env.py"
```

コマンド例は `/` 区切りで書いているが、Windows でもそのまま `python $SK/scripts/...` で動く。
未インストールなら `INSTALL.md` を案内する。

## 0. 環境を整える（最初に1回）

```bash
python $SK/scripts/setup_env.py      # 不足を自動インストール（できない分はコマンドを表示）
python $SK/tests/run_e2e.py          # このマシンでスキルが動くことの自己診断。ALL PASSED を確認
```

LibreOffice（Impress込み）が無いとスライドを画像化して目視確認できない。Linuxで `soffice` はあるのに
「source file could not be loaded」になるのは Impress 抜けが原因で、デッキの破損ではない。

## 1. テンプレと案件情報をそろえる

- テンプレ: `$SK/assets/*.pptx` があればそれ、なければユーザーが添付した直近の提案書pptx。
  **どちらも無ければ聞く。** 推測で自作しない
- 案件情報: ユーザーが貼ったメモ。`examples/two-station-line/request.md` が依頼文の典型形

## 2. テンプレを読む

```bash
python $SK/scripts/inspect_template.py <template.pptx> --layouts --tokens
python $SK/scripts/pptx_tool.py text <template.pptx>
python $SK/scripts/pptx_tool.py thumbnail <template.pptx> tpl      # tpl-1.jpg … を目で見る
```

確認すること:

- **スライド割り当てが正解例の `TEMPLATE` と一致するか**（VRAINテンプレなら 表紙=1、目次=2、中扉=3/5/8/11、
  本編=4/6/7/9/10、前案件の検証結果=12〜25、Appendix=26〜39）。違えば `TEMPLATE` を合わせる
- `media` 列が付いているスライド（画像あり）は**書き直さず丸ごと流用**するか捨てるかのどちらか
- スライドサイズがA4（9906000×6858000）でなければ汎用モード → `references/generic-template.md`

## 3. 正解例をコピーして CONTENT を書く

```bash
cp $SK/examples/two-station-line/build_deck.py ./build_deck.py
```

`build_deck.py` の `CONTENT` だけを案件情報で書き換える。**スライドの型関数（`s_genjou` など）と
`TEMPLATE` 以外の座標・色・文言の作法は変えない。** そこを変えると Claude Code と Codex で出力がずれる。

コピー先からでもスキル本体は自動で見つかる（`.agents/skills` `~/.agents/skills` `~/.codex/skills`
`~/.claude/skills` などを探す）。見つからないと言われたら `PROPOSAL_DECK_SKILL=$SK` を付けて実行する。

書き方のルール（正解例がすべてこの通りになっている）:

- 各スライドのリード文は「■〜いたします。」の一文で、そのスライドの結論を言い切る
- 課題カードの最後の1項目は赤太字 `(text, 1, d.RED)` にして「だから検査機が要る」につなげる
- ステーションは `stations` に1件ずつ。構成表は**見積書の品目と同じ並び**（区分/仕様/数量）
- 案件に無い章は消してよいが、残す章の型は正解例のまま使う
- `CONTENT["sections"]` は中扉と目次のⅠ〜Ⅳ。テンプレの前案件の章名は実行時に自動で置き換わる

## 4. 組む

```bash
python build_deck.py <template.pptx> YYYYMMDD_<件名>_ご提案書.pptx
```

スクリプトは、構造操作（複製・並べ替え・削除）→ 中身の書き込み → 表紙差し替え → 章名差し替え →
はみ出しチェック → 構造検証 → 前案件の残骸チェックの順に実行し、1つでも問題があれば終了コード1で止まる。
最後の行が `off-slide: 0 | validate: PASS | leftovers: none` になっていることを確認する。

## 5. 目で見る（省略不可）

```bash
python $SK/scripts/pptx_tool.py render YYYYMMDD_<件名>_ご提案書.pptx qa
```

`qa/p-*.jpg` を**全ページ**見る。自動チェックは座標のはみ出ししか拾わず、**文字の見切れは拾えない**。

- 文字が枠からはみ出す・2行のラベルが枠を割る → 文言を短くする（型は変えない）
- カードの下が大きく空く → 箇条書きを足す、または正解例と同じ数にそろえる
- 前案件の写真・社名が残っていない（Appendix以外）

日本語は全角1文字 ≒ フォントのpt数。11ptなら幅4,000,000EMUに約28文字、が目安。
直したら `build_deck.py` を再実行 → 再レンダリング（PDFを作り直さないと画像は変わらない）。

## 6. 渡す

ファイル名は `YYYYMMDD_<件名>_ご提案書.pptx`。作業フォルダに保存し、ファイル送信ツールがあればそれで渡す。
返信には **章立ての表** と **確認が必要な箇所** を必ず書く。

## 数字と固有名詞の扱い

提案書はそのまま客先に出るので:

- **もらっていない数値を断定で書かない。** 処理時間の内訳など推定が要るものは、スライドに「当社想定」
  「実機構成確定後に実測にて再提示」と書き、返信でも推定値だと伝える
- **社名・担当者名が未確定なら `〇〇株式会社　御中` のままにし、返信の冒頭で指摘する**
- 金額は載せない（見積書の領域）。載せる方針なら確認する
- 案件メモの社内向け情報（予算感・競合状況・支払条件・社内の担当経緯など）はスライドに出さない

## 部品だけ使いたいとき

正解例に無いスライドが要るときは `scripts/deck_lib.py` の部品で組む。座標・色の実値は
`references/design-system.md`。

```python
import deck_lib as d          # from deck_lib import * は使わない（configure_for の更新が反映されない）
i = d.ids(); (x1, w1), (x2, w2) = d.cols(2)
b = d.card(i, x1, d.TOP, w1, 2926080, "＜左＞", ["・…"])
d.build(path, "タイトル", "■…いたします。", b)
```

部品: `card` `banner` `point_row` `flow` `gantt` `table` `header_row` `zebra` `tcell` `shape` `textbox` `R` `P`、
ファイル操作: `unpack` `edit_slide` `set_para` `swap_runs` `paragraph_texts` `run_texts` `reorder_slides` `pack`
`check_bounds` `configure_for`、パッケージ操作: `scripts/pptx_tool.py add_slide|clean|validate|text|render|thumbnail`。

スライドの書き換えは必ず `d.edit_slide()` を通す。`open(path, "w")` を先に開いてから置換すると、
置換が失敗した瞬間に空のスライドが残り、デッキ全体が開けなくなる。

## 参照ファイル

| ファイル | 読むとき |
|---|---|
| `examples/two-station-line/build_deck.py` | **毎回**。正解例。コピーして使う |
| `examples/two-station-line/README.md` | 正解例の各スライドの役割と、他の案件への当てはめ方 |
| `references/design-system.md` | 色・座標・フォント・レイアウトの実値が要るとき |
| `references/deck-structure.md` | 章立てを増減するとき、各スライドの考え方 |
| `references/generic-template.md` | VRAIN以外のテンプレを渡されたとき |
| `PROMPT.md` | ユーザーに依頼文の書き方を聞かれたとき |
| `INSTALL.md` | 導入方法を聞かれたとき、スキルが見つからないと言われたとき |
