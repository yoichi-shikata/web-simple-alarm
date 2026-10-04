#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""正解例：2ステーション構成の新設ライン向け 外観検査提案書（VRAINテンプレート）

これは「理想形」として採用されたデッキと同じ構成・同じ書き方・同じ座標で組む
完全なビルドスクリプト。内容は架空のサンプル案件（公開リポジトリのため）。

新しい案件では、このファイルを作業ディレクトリにコピーして **CONTENT だけ**
書き換える。スライドの型・配置・色・文言の作法はここで固定されているので、
Claude Code でも Codex でも同じ見た目のデッキになる。

  python build_deck.py <VRAINテンプレ.pptx> <出力.pptx>

前提のテンプレ構造（VRAIN提案書。違う場合は TEMPLATE を合わせる）:
  slide1=表紙 / slide2=目次 / slide3,5,8,11=中扉Ⅰ〜Ⅳ / slide4,6,7,9,10=本編
  slide12-25=前案件の検証結果（捨てる） / slide26=中扉Ⅴ Appendix / slide27-39=Appendix
"""
import datetime
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def _skill_scripts():
    """Find the skill's scripts/ whether this file runs in place or as a copy
    in a work folder. Set PROPOSAL_DECK_SKILL=<skill dir> to override."""
    cands = [os.environ.get("PROPOSAL_DECK_SKILL", ""), os.path.join(HERE, "..", ".."),
             ".agents/skills/proposal-deck", ".claude/skills/proposal-deck", ".codex/skills/proposal-deck",
             "~/.agents/skills/proposal-deck", "~/.codex/skills/proposal-deck", "~/.claude/skills/proposal-deck"]
    for c in cands:
        if c and os.path.exists(os.path.join(os.path.expanduser(c), "scripts", "deck_lib.py")):
            return os.path.abspath(os.path.join(os.path.expanduser(c), "scripts"))
    sys.exit("proposal-deck skill not found; set PROPOSAL_DECK_SKILL=<skill dir>")


sys.path.insert(0, _skill_scripts())
import deck_lib as d  # noqa: E402
import pptx_pkg as pk  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):  # Windows console encodings must not crash a run
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# =====================================================================
# CONTENT — 案件ごとにここだけ書き換える
# =====================================================================
today = datetime.date.today()
CONTENT = {
    "customer": "〇〇株式会社　御中",          # 未確定なら必ずこのまま。推測で埋めない
    "title": "PETボトル製品外観検査（新設ライン）",
    "date": "%d年%d月%d日" % (today.year, today.month, today.day),

    # ---- Ⅰ. 現状のご状況と課題（2カード型）
    "genjou": {
        "lead": "■新設ラインのご状況と、外観検査における課題を整理いたします。",
        "status": [
            "・増産に伴う新設ラインの構築を計画されており、充填機・キャッパー等を含むプロジェクト予算として始動",
            "・2028年6月　新ライン本稼働（2028年3月より試運転を開始）",
            "・生産量：3,600本/時（1.0秒/個）",
            "・検査対象：PETボトル製品の外観（液面・異物・キャップ・ラベル）",
            "・充填直後と、キャップ・ラベル装着後とで、工程条件が大きく異なる",
        ],
        "issues": [
            "・1.0秒/個のタクトに対し、目視検査では速度・精度ともに対応できない",
            "・液面／異物／キャップ／ラベルと検査項目が多岐にわたる",
            "・異物（樹脂片・繊維くず 等）は微小かつ不定形で、判定基準の統一が難しい",
            "・充填直後は泡立ちがあり、撮像タイミングの考慮が必要",
            ("・新設ラインのため、稼働開始時点から全数検査体制を確立する必要がある", 1, d.RED),
        ],
        "banner": [("⇒ 新設ラインの立ち上げに合わせ、", d.NAVY),
                   ("AI外観検査による全数自動検査", d.RED), ("を実現します。", d.NAVY)],
    },

    # ---- Ⅱ. ご提案｜全体構成（フロー型 + ポイント3行）
    "overview": {
        "title": "ご提案｜2ステーション構成による全数外観検査",
        "lead": "■ライン構成に合わせ、上流・下流の2ステーションで全数外観検査を行う構成をご提案いたします。",
        # (表示テキスト, 種別)  種別: "ext"=対象外の既設/他社設備, "ours"=当社納入, "out"=出口
        "flow": [("充填機", "ext"), ("ST1\n液面・異物検査", "ours"), ("キャッパー・\nラベラー", "ext"),
                 ("ST2\nキャップ・\nラベル検査", "ours"), ("後工程へ\n（NG：排出）", "out")],
        "flow_note": ["※ST1は充填直後、ST2はキャップ・ラベル装着後に設置します。",
                      "　 各STともNG品はプッシャーで排出し、サイドコンベア（バッファコンベア）へ移送します。"],
        "points": [
            ("工程条件に合わせた2ST構成",
             ["充填直後に液面・異物検査（ST1）、キャップ・ラベル装着後に外観検査（ST2）を配置。",
              "工程ごとに異なる撮像条件それぞれに最適な撮像環境を構築します。"]),
            ("AIによる高速・高精度な判定",
             ["1.0秒/個のタクトに対応。AIの領域抽出により、ルールベースでは分離が難しい",
              "微小・不定形な異物（樹脂片・繊維くず 等）を特徴単位で判別します。"]),
            ("搬送・治具・排出・制御まで一括対応",
             ["カメラ・照明からコンベア、治具、プッシャー排出、制御盤までを当社にて一括対応。",
              "分離発注による取り合い調整の手間なく、新ラインの立ち上げに対応いたします。"]),
        ],
    },

    # ---- Ⅱ. ステーション詳細（2カード型：左＝検査概要、右＝構成表）
    "stations": [
        {
            "title": "ST1｜液面・異物検査（上流）",
            "lead": "■充填直後の工程にて、液面高さと液中異物の検査をカメラ2台で行います。",
            "summary": [
                "・トップチェーンコンベア上を搬送される製品を、バックライト透過で撮像します",
                "・液面高さの過不足と、液中に混入した異物を検出します",
                "・泡立ちの影響を避けるため、撮像位置を充填機から一定距離下流に設定します",
                "・NG品はプッシャーで排出し、サイドコンベア（バッファコンベア）へ移送します",
            ],
            "spec": [("AIプロダクト", "Phoenix Vision ＋ Phoenix Edge", "1式"),
                     ("カメラ", "1,200万画素 モノクロカメラ", "2台"),
                     ("照明", "バックライト照明", "2台"),
                     ("搬送", "トップチェーンコンベア", "1式"),
                     ("治具", "ワーク保持・位置決め治具", "1式"),
                     ("排出", "プッシャー排出＋サイドコンベア", "1式"),
                     ("制御", "制御盤・シーケンス制御", "1式")],
            "banner": [("タクト：1.0秒/個（生産量 3,600本/時）　／　検出対象：", d.NAVY),
                       ("液面過不足・異物（樹脂片・繊維くず 等）", d.RED)],
        },
        {
            "title": "ST2｜キャップ・ラベル検査（下流）",
            "lead": "■キャップ・ラベル装着後の工程にて、外観検査をカメラ3台で行います。",
            "summary": [
                "・サイドベルトにて挟持した状態のワークを搬送しながら検査します",
                "・キャップの浮き・傾き・欠けを、上方カメラ1台で検査します",
                "・ラベルのずれ・しわ・破れを、側方カメラ2台で全周検査します",
                "・NG品はコンベア移載後にプッシャーで排出し、サイドコンベアへ移送します",
            ],
            "spec": [("AIプロダクト", "Phoenix Vision ＋ Phoenix Edge\n（3カメラライセンス）", "1式"),
                     ("カメラ①", "500万画素 カラーカメラ\n（キャップ）", "1台"),
                     ("カメラ②", "500万画素 カラーカメラ（ラベル）", "2台"),
                     ("照明", "リング照明・バー照明", "3台"),
                     ("搬送", "サイドベルトコンベア", "1式"),
                     ("治具", "ワーク保持・位置決め治具", "1式"),
                     ("排出", "プッシャー排出＋サイドコンベア", "1式"),
                     ("制御", "制御盤・シーケンス制御", "1式")],
            "banner": [("検査項目：", d.NAVY), ("キャップ＝浮き・傾き・欠け　／　ラベル＝ずれ・しわ・破れ", d.RED),
                       ("　（タクト：1.0秒/個）", d.NAVY)],
        },
    ],

    # ---- Ⅱ. ご提供範囲（表型）  ●=当社 △=協議 ―=貴社手配
    "scope": {
        "lead": "■上流ST1・下流ST2ともに、撮像から搬送・排出・制御まで当社にて一括対応いたします。",
        "rows": [
            ("AIプロダクト", "Phoenix Vision（AI外観検査ソフト）＋ Phoenix Edge", "●", "ST2は3カメラライセンス"),
            ("撮像機器", "産業用カメラ・レンズ", "●", "ST1：2台／ST2：3台"),
            ("照明", "バックライト照明・リング照明・バー照明", "●", "ST1：透過／ST2：反射"),
            ("検査架台", "架台、遮光カバー", "●", "各STごとに製作"),
            ("搬送装置", "トップチェーンコンベア／サイドベルトコンベア", "●", "ST1・ST2それぞれに対応"),
            ("ワーク治具", "製品保持・位置決め治具", "●", "―"),
            ("NG排出機構", "プッシャー排出、サイドコンベア", "●", "バッファコンベアへ移送"),
            ("制御", "制御盤、シーケンス制御、上位設備との信号連携", "●", "―"),
            ("設置・立ち上げ", "現地据付、調整、試運転立会", "●", "2028年3月〜（試運転期間）"),
            ("ユーティリティ", "電源・エア一次側配管、設置スペースのご手配", "―", "貴社にてご手配をお願いします"),
            ("既設設備との調整", "充填機・キャッパー・ラベラーとの機械的取り合い", "△", "貴社・設備メーカ様と協議"),
        ],
    },

    # ---- Ⅱ. タクト適合（帯 + 内訳表）。実測が無ければ必ず「当社想定」
    "tact": {
        "lead": "■1.0秒/個のタクトに対し、判定結果の出力まで余裕をもって処理できる見込みです。",
        "total_ms": 1000, "label": "1.0秒/個（1,000ms）", "rate": "3,600本/時",
        "bar": [("撮像・画像取込", 300), ("AI処理・判定出力", 300)],
        "rows": [("①ワーク検出・撮像トリガ出力", "約50ms", "当社想定"),
                 ("②撮像・画像転送（ST1：2台／ST2：3台）", "約250ms", "当社想定（同時撮像・並列転送）"),
                 ("③AI処理（領域抽出）", "約250ms", "当社想定（Phoenix Edge での処理）"),
                 ("④判定・排出信号出力", "約50ms", "当社想定")],
    },

    # ---- Ⅲ. 導入スケジュール（ガント型）
    "schedule": {
        "lead": "■2028年6月の新ライン本稼働から逆算した、導入スケジュールをご提示いたします。",
        "months": [6, 7, 8, 9, 10, 11, 12, 1, 2, 3, 4, 5, 6, 7],
        "years": [(0, 7, "2027年"), (7, 7, "2028年")],
        # (行ラベル, 開始列, 月数, 種別, バー文言)  種別: "main"/"order"/"work"
        "rows": [("検証・仕様のご検討", 0, 3, "main", "検証・仕様検討"),
                 ("最終お見積り・ご発注", 3, 2, "order", "ご発注"),
                 ("設計・製作・社内調整", 5, 4, "work", "設計・製作"),
                 ("現地据付・試運転", 9, 3, "work", "据付〜試運転・追加学習"),
                 ("新ライン本稼働", 12, 2, "main", "本稼働")],
        "banner": [("⇒ 2028年3月からの試運転開始に間に合わせるため、", d.NAVY),
                   ("2027年9〜10月中のご発注", d.RED), ("をお願いいたします。", d.NAVY)],
    },

    # ---- Ⅲ. 今後の進め方（ステップ型 + 2カード）
    "next": {
        "title": "今後の進め方｜検証について",
        "lead": "■実サンプルによる事前検証を実施し、検出性能をご確認いただいたうえで仕様を確定いたします。",
        "steps": [("STEP 1", "サンプルご提供・\n撮像条件のすり合わせ"),
                  ("STEP 2", "撮像・学習・判定検証\n（2027年6月 実施予定）"),
                  ("STEP 3", "検証結果のご報告・\n仕様確定")],
        "focus": [("① 検出性能の確認（最優先）", 1, d.NAVY),
                  "　微小・不定形な異物の検出可否、および過検知の発生有無を確認します。",
                  ("② 搬送状態での撮像可否（次点）", 1, d.NAVY),
                  "　ワークが流れた状態でも必要な画質で撮像できるかを確認します。"],
        "items": ["・液面：高さの過不足", "・液中：異物（樹脂片・繊維くず 等）",
                  "・キャップ：浮き・傾き・欠け", "・ラベル：ずれ・しわ・破れ", "・良品に対する過検知の有無"],
        "banner": "⇒ 検証結果は、2027年7月上旬のご訪問時にご報告いたします。",
    },

    # ---- Ⅳ. 比較観点（表型）
    "compare": {
        "lead": "■複数社をご比較いただく際は、下記5つの観点でご評価いただくことを推奨いたします。",
        "head": ["観点", "ご確認いただきたいポイント", "当社\n（Phoenix Vision＋Edge）",
                 "A社タイプ\n汎用AI検査ソフト", "B社タイプ\nルールベース検査機"],
        "rows": [
            ("費用", "・初期費用だけでなく、追加学習・モデル更新・年間ライセンス費用を含めた総額で比較できるか\n"
                     "・搬送・治具・排出・制御まで含めた一括見積りか（分離発注時の取り合い調整工数）",
             "初期費用のみで完結\n一括対応", "機材費用\n＋年間ソフト\nライセンス費用", "装置は別途\n手配が必要"),
            ("精度\n（見逃し・過検知）", "・自社の実サンプルで事前検証した結果が提示されるか\n"
                                    "・見逃し率／過検知率を、両者提示の条件として合意できるか\n"
                                    "・不良傾向の変化に追加学習で対応できるか",
             "事前検証\n＋追加学習", "汎用モデル依存", "微小・不定形\n異物が苦手"),
            ("処理速度", "・タクト（1.0秒/個）内に判定出力まで収まるか\n"
                        "・多カメラ構成（ST1：2台／ST2：3台）での実測値が提示されるか",
             "約600ms見込み\n（当社想定）", "PC構成に依存", "高速"),
            ("UIの\n使いやすさ", "・現場担当者だけで追加学習・調整まで完結できるか\n"
                               "・専門知識やベンダー依頼が毎回必要にならないか\n"
                               "・判定結果・画像の蓄積と振り返りができるか",
             "現場で自走可能", "専門知識が必要", "設定項目は単純"),
            ("対応範囲", "・カメラ・照明から搬送・治具・排出・制御まで一括で対応できるか\n"
                        "・新設ラインの立ち上げ（試運転期間）に伴走できるか",
             "ワンストップ対応\n試運転立会あり", "ソフトのみ", "装置主体"),
        ],
    },

    # ---- Ⅳ. 導入効果（Before→After + 3カード）
    "effect": {
        "lead": "■新ラインの稼働開始時点から、全数自動検査による安定した品質保証体制を構築できます。",
        "before": ("＜目視検査の場合＞", "1.0秒/個", "人手では対応不可"),
        "after": ("＜AI外観検査 導入後＞", "3,600本/時", "全数を自動検査"),
        "note": "※ 上記は効果イメージです。実際の検出性能は事前検証の結果によりご提示いたします。",
        "cards": [("検査基準の一定化", ["作業者ごとの判断のばらつきをなくし、", "同一基準での判定を継続できます。"]),
                  ("採用・教育負担の軽減", ["検査要員の採用・教育にかかる", "負担とリスクを低減できます。"]),
                  ("データの蓄積・活用", ["判定結果と画像を蓄積し、傾向分析や", "トレーサビリティに活用できます。"])],
    },

    # 中扉・目次の章名（Ⅰ〜Ⅳ。ⅤのAppendixはテンプレのまま）
    "sections": ["現状のご状況と課題", "ご提案", "導入スケジュール", "ご検討にあたっての比較観点"],
}

# =====================================================================
# TEMPLATE — VRAINテンプレのスライド割り当て
# =====================================================================
TEMPLATE = {
    "cover": "slide1", "toc": "slide2",
    "dividers": ["slide3", "slide5", "slide8", "slide11"],
    "body": ["slide4", "slide6", "slide7", "slide9", "slide10"],   # layout5、画像なし → 書き直してよい
    "appendix": ["slide26"] + ["slide%d" % i for i in range(27, 40)],  # 画像あり → 触らない
    # 表紙の段落の並び（paragraph_texts の添字）: [0]=御中 [1]=件名 [2]=ご提案書 [3]=日付
    "cover_idx": {"customer": 0, "title": 1, "date": 3},
}

FLOW_KIND = {"ext": (d.MIDGRAY, d.WHITE), "ours": (d.NAVY, d.WHITE), "out": (d.BLUE, d.WHITE)}
BAR_KIND = {"main": d.NAVY, "order": d.RED, "work": d.BLUE}


def runs(spec, sz):
    return [d.R(t, sz=sz, b=1, color=c) for t, c in spec]


# =====================================================================
# スライドの型（ここは変えない）
# =====================================================================
def s_genjou(path, c):
    i = d.ids(); (x1, w1), (x2, w2) = d.cols(2)
    b = d.card(i, x1, d.TOP, w1, 2926080, "＜案件のご状況＞", c["status"])
    b += d.card(i, x2, d.TOP, w2, 2926080, "＜課題＞", c["issues"], header_fill=d.RED)
    b += d.banner(i, d.L, 4892040, d.W, 868680, runs(c["banner"], 1300))
    d.build(path, "現状のご状況と課題", c["lead"], b)


def s_overview(path, c):
    i = d.ids()
    labels = [(t, *FLOW_KIND[k]) for t, k in c["flow"]]
    b = d.flow(i, labels)
    b += d.textbox(next(i), "FlowNote", 548640, 2788920, 8869680, 420000,
                   [d.P(d.R(t, sz=950, color=d.GRAY)) for t in c["flow_note"]])
    b += d.textbox(next(i), "PointsHdr", d.L, 3200400, d.W, 320040,
                   [d.P(d.R("＜本構成のポイント＞", sz=1300, b=1, color=d.NAVY))], anchor="ctr")
    y = 3584448
    for k, (label, lines) in enumerate(c["points"]):
        b += d.point_row(i, d.L, y, d.W, 749808, label, lines, fill=d.LTGRAY if k % 2 == 0 else d.WHITE)
        y += 841248
    d.build(path, c["title"], c["lead"], b)


def s_station(path, c):
    i = d.ids(); (x1, w1), (x2, w2) = d.cols(2)
    n = len(c["spec"])
    row_h = 380000 if n <= 7 else 370000
    ch = 400000 + n * row_h
    b = d.card(i, x1, d.TOP, w1, ch, "＜検査概要＞", c["summary"], body_sz=1000,
               spc_aft=700 if len(c["summary"]) <= 4 else 600)
    rows = [d.header_row(["区分", "仕様", "数量"])] + d.zebra(
        c["spec"], [{"sz": 950, "b": 1, "color": d.NAVY}, {"sz": 950}, {"sz": 950, "algn": "ctr"}])
    b += d.table(next(i), x2, d.TOP, [1150000, w2 - 2050000, 900000], rows, [400000] + [row_h] * n)
    by = d.TOP + ch + 230000
    b += d.banner(i, d.L, by, d.W, 731520 if n > 7 else 868680, runs(c["banner"], 1200))
    d.build(path, c["title"], c["lead"], b)


def s_scope(path, c):
    i = d.ids()
    rows = [d.header_row(["区分", "主な内容", "当社\nご提供", "備考"], algns=["ctr", "l", "ctr", "l"])]
    for k, (a, bb, mark, note) in enumerate(c["rows"]):
        fill = d.LTGRAY if k % 2 else d.WHITE
        rows.append([d.tcell(a, sz=950, b=1, color=d.NAVY, fill=fill),
                     d.tcell(bb, sz=950, fill=fill),
                     d.tcell(mark, sz=1100, b=1, color=d.NAVY if mark == "●" else d.GRAY, fill=fill, algn="ctr"),
                     d.tcell(note, sz=950, fill=fill)])
    b = d.table(next(i), d.L, 1600200, [1550000, 3350000, 1000000, d.W - 5900000], rows,
                [420000] + [335000] * len(c["rows"]))
    b += d.textbox(next(i), "Note", d.L, 5989320, d.W, 260000,
                   [d.P(d.R("※ ●＝当社ご提供範囲、△＝貴社・設備メーカ様との協議事項、―＝貴社ご手配範囲。"
                            "最終的な範囲は仕様確定時のお見積書にて確定いたします。", sz=900, color=d.GRAY))])
    d.build(path, "ご提供範囲（納入スコープ）", c["lead"], b)


def s_tact(path, c):
    i = d.ids()
    total = c["total_ms"]
    used = sum(ms for _, ms in c["bar"])
    scale = d.W / float(total)
    b = d.textbox(next(i), "TactLabel", d.L, 1737360, d.W, 320040,
                  [d.P([d.R("タクトタイム　", sz=1200, b=1, color=d.NAVY),
                        d.R(c["label"], sz=1200, b=1, color=d.RED),
                        d.R("　＜生産量 %s＞" % c["rate"], sz=1200, b=1, color=d.NAVY)])], anchor="ctr")
    x = d.L
    segs = [(n, ms, f, d.WHITE) for (n, ms), f in zip(c["bar"], (d.NAVY, d.BLUE))]
    segs.append(("バッファ", total - used, d.LTGRAY, d.GRAY))
    for name, ms, fill, col in segs:
        w = int(ms * scale)
        b += d.shape(next(i), "Seg", x, 2103120, w, 548640, fill=fill, line=d.WHITE,
                     paras=[d.P(d.R(name, sz=950, b=1, color=col), algn="ctr"),
                            d.P(d.R("{:,}ms".format(ms), sz=1050, b=1, color=col), algn="ctr")])
        x += w
    b += d.textbox(next(i), "BarNote", d.L, 2743200, d.W, 260000,
                   [d.P(d.R("⇒ 判定結果の出力まで約{:,}ms。{}のタクトに対し、約{:,}msの余裕を見込んでいます。".format(
                       used, c["label"].split("（")[0], total - used), sz=1050, b=1, color=d.NAVY))])
    rows = [d.header_row(["工程", "所要時間", "出典・備考"], algns=["l", "ctr", "l"])]
    for k, (a, t, note) in enumerate(c["rows"]):
        fill = d.LTGRAY if k % 2 else d.WHITE
        rows.append([d.tcell(a, sz=950, fill=fill), d.tcell(t, sz=950, fill=fill, algn="ctr"),
                     d.tcell(note, sz=950, fill=fill)])
    rows.append([d.tcell("合計", sz=950, b=1, color=d.NAVY, fill=d.LTBLUE),
                 d.tcell("約{:,}ms".format(used), sz=950, b=1, color=d.RED, fill=d.LTBLUE, algn="ctr"),
                 d.tcell("{}に対し約{:,}msの余裕".format(c["label"].split("（")[0], total - used),
                         sz=950, b=1, color=d.NAVY, fill=d.LTBLUE)])
    b += d.table(next(i), d.L, 3200400, [3200000, 1800000, d.W - 5000000], rows,
                 [370000] + [340000] * (len(c["rows"]) + 1))
    b += d.textbox(next(i), "Disc", d.L, 5303520, d.W, 500000,
                   [d.P(d.R("※ 上記はいずれも当社想定値です。現地条件（画像サイズ・カメラ構成・ネットワーク構成・"
                            "判定モデル）により変動いたします。", sz=900, color=d.GRAY)),
                    d.P(d.R("　 実機構成の確定後、実測値にて改めてご提示いたします。", sz=900, color=d.GRAY))])
    d.build(path, "処理時間｜タクトタイムへの適合", c["lead"], b)


def s_schedule(path, c):
    i = d.ids()
    rows = [(lab, s, n, BAR_KIND[k], txt) for lab, s, n, k, txt in c["rows"]]
    b = d.gantt(i, rows, c["months"], c["years"])
    b += d.banner(i, d.L, 5181600, d.W, 731520, runs(c["banner"], 1250))
    d.build(path, "導入スケジュール", c["lead"], b)


def s_next(path, c):
    i = d.ids()
    (sx1, sw), (sx2, _), (sx3, _) = d.cols(3, gap=228600)
    b = ""
    for x, (num, label), col in zip((sx1, sx2, sx3), c["steps"], (d.NAVY, d.RED, d.BLUE)):
        b += d.shape(next(i), "Step", x, 1737360, sw, 1005840, fill=col, prst="roundRect",
                     paras=[d.P(d.R(num, sz=1000, b=1, color=d.WHITE), algn="ctr")] +
                           [d.P(d.R(t, sz=1050, b=1, color=d.WHITE), algn="ctr") for t in label.split("\n")])
    for gx in (sx1 + sw, sx2 + sw):
        b += d.shape(next(i), "StepArrow", gx + 27432, 2148840, 173736, 182880, fill=d.BORDER, prst="rightArrow")
    (x1, w1), (x2, w2) = d.cols(2)
    b += d.card(i, x1, 3017520, w1, 2011680, "＜重点的にご確認いただく点＞", c["focus"], body_sz=1000, spc_aft=400)
    b += d.card(i, x2, 3017520, w2, 2011680, "＜検証項目＞", c["items"], body_sz=1000, spc_aft=400,
                header_fill=d.BLUE)
    b += d.banner(i, d.L, 5181600, d.W, 640080, [d.R(c["banner"], sz=1250, b=1, color=d.NAVY)])
    d.build(path, c["title"], c["lead"], b)


def s_compare(path, c):
    i = d.ids()
    rows = [[d.tcell(h, sz=1050, b=1, color=d.WHITE, fill=d.NAVY, algn="l" if k == 1 else "ctr")
             for k, h in enumerate(c["head"])]]
    for k, (a, pts, ours, A, B) in enumerate(c["rows"]):
        fill = d.LTGRAY if k % 2 else d.WHITE
        rows.append([d.tcell(a, sz=1000, b=1, color=d.NAVY, fill=fill, algn="ctr"),
                     d.tcell(pts, sz=900, fill=fill),
                     d.tcell(ours, sz=950, b=1, color=d.NAVY, fill=d.LTBLUE, algn="ctr"),
                     d.tcell(A, sz=950, fill=fill, algn="ctr"),
                     d.tcell(B, sz=950, fill=fill, algn="ctr")])
    b = d.table(next(i), d.L, 1627632, [1143000, 3520440, 1417320, 1371600, 1508760], rows,
                [620000] + [742851] * len(c["rows"]))
    d.build(path, "ご検討にあたっての比較観点", c["lead"], b)


def s_effect(path, c):
    i = d.ids()
    bh, bv, bs = c["before"]; ah, av, as_ = c["after"]
    b = d.shape(next(i), "Before", 822960, 1783080, 3383280, 1691640, fill=d.LTGRAY)
    b += d.textbox(next(i), "BeforeH", 822960, 1847088, 3383280, 365760,
                   [d.P(d.R(bh, sz=1200, b=1, color=d.GRAY), algn="ctr")], anchor="ctr")
    b += d.textbox(next(i), "BeforeB", 822960, 2240280, 3383280, 1097280,
                   [d.P(d.R(bv, sz=2400, b=1, color=d.GRAY), algn="ctr"),
                    d.P(d.R(bs, sz=1300, b=1, color=d.GRAY), algn="ctr")], anchor="ctr")
    b += d.shape(next(i), "Arrow", 4434840, 2331720, 822960, 594360, fill=d.NAVY, prst="rightArrow")
    b += d.shape(next(i), "After", 5486400, 1783080, 3383280, 1691640, fill=d.LTBLUE, line=d.NAVY)
    b += d.textbox(next(i), "AfterH", 5486400, 1847088, 3383280, 365760,
                   [d.P(d.R(ah, sz=1200, b=1, color=d.NAVY), algn="ctr")], anchor="ctr")
    b += d.textbox(next(i), "AfterB", 5486400, 2240280, 3383280, 1097280,
                   [d.P(d.R(av, sz=2400, b=1, color=d.NAVY), algn="ctr"),
                    d.P(d.R(as_, sz=1300, b=1, color=d.NAVY), algn="ctr")], anchor="ctr")
    b += d.textbox(next(i), "Note", 822960, 3520440, 8046720, 274320,
                   [d.P(d.R(c["note"], sz=900, color=d.GRAY), algn="ctr")])
    b += d.textbox(next(i), "SubHdr", d.L, 3931920, d.W, 320040,
                   [d.P(d.R("＜定量以外の効果＞", sz=1300, b=1, color=d.NAVY))], anchor="ctr")
    for (x, w), (title, lines) in zip(d.cols(3, gap=183520), c["cards"]):
        b += d.shape(next(i), "C", x, 4315968, w, 1417320, fill=d.WHITE, line=d.BORDER)
        b += d.shape(next(i), "CH", x, 4315968, w, 384048, fill=d.NAVY)
        b += d.textbox(next(i), "CHT", x, 4315968, w, 384048,
                       [d.P(d.R(title, sz=1150, b=1, color=d.WHITE), algn="ctr")], anchor="ctr")
        b += d.textbox(next(i), "CB", x + 109728, 4754880, w - 219456, 914400,
                       [d.P(d.R(t, sz=1000, color=d.GRAY)) for t in lines])
    d.build(path, "導入効果", c["lead"], b)


# =====================================================================
def main(template, out):
    work = os.path.splitext(os.path.abspath(out))[0] + "_work"
    import shutil
    shutil.rmtree(work, ignore_errors=True)
    d.unpack(template, work)
    S = lambda n: os.path.join(work, "ppt", "slides", n + ".xml")  # noqa: E731
    C = CONTENT

    # 表紙の現在の文言を記録（前案件の残骸チェックに使う）
    cover_before = d.paragraph_texts(open(S(TEMPLATE["cover"]), encoding="utf-8").read())
    old_marks = [cover_before[TEMPLATE["cover_idx"][k]] for k in ("customer", "title")]
    print("template cover:", cover_before)

    # ---- 構造操作を先に全部終える
    n_body = 4 + len(C["stations"]) + 4          # 現状 + 全体構成 + ST×n + 範囲 + タクト + 日程 + 進め方 + 比較 + 効果
    body = list(TEMPLATE["body"])
    while len(body) < n_body:
        body.append(pk.add_slide(work, body[0] + ".xml")[:-4])
    it = iter(body)
    plan = {"genjou": next(it), "overview": next(it), "stations": [next(it) for _ in C["stations"]],
            "scope": next(it), "tact": next(it), "schedule": next(it), "next": next(it),
            "compare": next(it), "effect": next(it)}
    dv = TEMPLATE["dividers"]
    order = ([TEMPLATE["cover"], TEMPLATE["toc"], dv[0], plan["genjou"], dv[1], plan["overview"]]
             + plan["stations"] + [plan["scope"], plan["tact"], dv[2], plan["schedule"], plan["next"],
                                   dv[3], plan["compare"], plan["effect"]] + TEMPLATE["appendix"])
    d.reorder_slides(work, order)
    pk.clean(work)

    # ---- 中身を書く
    s_genjou(S(plan["genjou"]), C["genjou"])
    s_overview(S(plan["overview"]), C["overview"])
    for name, st in zip(plan["stations"], C["stations"]):
        s_station(S(name), st)
    s_scope(S(plan["scope"]), C["scope"])
    s_tact(S(plan["tact"]), C["tact"])
    s_schedule(S(plan["schedule"]), C["schedule"])
    s_next(S(plan["next"]), C["next"])
    s_compare(S(plan["compare"]), C["compare"])
    s_effect(S(plan["effect"]), C["effect"])

    # ---- 表紙：テンプレの現在の文言を見つけて置換
    ci = TEMPLATE["cover_idx"]
    def cover(x):
        cur = d.paragraph_texts(x)
        for key in ("customer", "title", "date"):
            x = d.set_para(x, cur[ci[key]], C[key])
        return x
    d.edit_slide(work, TEMPLATE["cover"], cover)

    # ---- 目次・中扉：前案件の章名 → 今回の章名（同時置換）
    old_secs = [d.run_texts(open(S(s), encoding="utf-8").read())[-1] for s in dv]
    pairs = [(o, n) for o, n in zip(old_secs, C["sections"]) if o != n]
    for s in [TEMPLATE["toc"]] + dv:
        d.edit_slide(work, s, lambda x: d.swap_runs(x, pairs))

    # ---- 仕上げと検証
    bad = d.check_bounds(work)
    d.pack(work, out)
    errs = pk.validate(out)
    txt = pk.text(out)
    left = [m for m in old_marks if m and m in txt]
    print("slides:", len(order), "| off-slide:", len(bad), "| validate:", "PASS" if not errs else errs,
          "| leftovers:", left or "none")
    if bad or errs or left:
        sys.exit(1)
    return out


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("usage: build_deck.py <template.pptx> <out.pptx>")
    main(sys.argv[1], sys.argv[2])
