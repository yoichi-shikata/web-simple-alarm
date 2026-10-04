# 汎用モード：VRAIN以外のテンプレートを使う

手順は同じで、設計トークンだけ差し替える。

## 1. 抽出する

```bash
python scripts/inspect_template.py <template.pptx> --layouts
```

出力から次を拾う:

- **スライドサイズ** → `deck_lib.L / W / TOP` を再計算。
  左右マージンは本文図形のxの最小値、本文上端は最初の図形のyから取る
- **本編レイアウト** → 最も多くの中身スライドが参照しているレイアウト。
  その `<p:ph idx="…">` の番号を `build(..., idx_title=, idx_msg=)` に渡す
- **色** → 本文図形の `<a:srgbClr val>` の出現頻度上位。
  濃色1・警告色1・淡色2・本文グレーの5つが揃えば足りる
- **フォント** → `<a:latin typeface>` の最頻値を `deck_lib.FONT` に入れる

## 2. 上書きする

```python
import deck_lib as d
d.configure_for("unpacked")          # スライドサイズから L / W / FOOT を再計算
d.FONT = "Yu Gothic"
d.NAVY, d.RED, d.LTBLUE, d.LTGRAY, d.GRAY = "1F3864", "C00000", "DEEBF7", "F2F2F2", "404040"
(x1, w1), (x2, w2) = d.cols(2)       # 列の座標はすべて cols() から取る
```

モジュール変数なので代入すれば以降の呼び出しに効く。ただし `card` などの
**デフォルト引数は定義時に束縛済み**なので、`header_fill=d.NAVY` のように明示で渡す。
`design-system.md` に書いてある `5074920` 等の座標はA4前提の実測値なので、
別サイズでは使わず `cols()` で出し直す。生成後は `d.check_bounds("unpacked")` で
はみ出しが0件であることを確認する。

## 3. プレースホルダが無いテンプレの場合

タイトル枠がプレースホルダでなく素のテキストボックスのテンプレもある。その場合は
`build(path, "", "", body, with_placeholders=False)` でプレースホルダを出さずに、
タイトルとリード文も `textbox()` で実測座標に置く。ヘッダの罫線やロゴがレイアウトではなく
スライドに直接描かれているときは、そのスライドを `add_slide.py` で複製して中身だけ差し替える。

## 4. 変わらないこと

- 構造操作（複製・削除・並べ替え）を全部終えてから中身を書く
- Appendix・事例など画像を抱えるスライドは書き直さず丸ごと流用する
- `validate.py --original` → `render_qa.sh` → 全ページ目視、は省略しない
