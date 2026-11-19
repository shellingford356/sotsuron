import re, ast, collections, datetime as dt
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

META = "sherry/ted_main_V2.xlsx"
TRANSCRIPT = "sherry/transcript.xlsx"
OUT_XLSX = "sherry/mergedTEDdata.xlsx"
OUT_CSV = "sherry/transcript_V2.csv"

# ---- 分析サンプルの除外条件 ----
MIN_WORDS = 300
MIN_WPM = 60
MIN_VOTES = 100

TAGS = ["Beautiful", "Confusing", "Courageous", "Fascinating", "Funny", "Informative", "Ingenious",
        "Inspiring", "Jaw-dropping", "Longwinded", "OK", "Obnoxious", "Persuasive", "Unconvincing"]
def tagcol(t): return t.lower().replace("-", "_")

# ---------------- テキスト整形 ----------------
WORD = re.compile(r"[^\W_]+(?:['’\-][^\W_]+)*")
PAREN = re.compile(r"\([^()]*\)")
UNCLEAR = re.compile(r"\[\s*unclear\s*\]", re.I)
SQUARE = re.compile(r"\[([^\[\]]*)\]")
LAUGH = re.compile(r"\(\s*laughter\s*\)", re.I)
APPL = re.compile(r"\(\s*applause\s*\)", re.I)
NOTE = re.compile(r"[♫♪]")

def clean(text):
    s = str(text).replace("\xa0", " ")
    n_paren = len(PAREN.findall(s))
    s = PAREN.sub(" ", s)
    s = UNCLEAR.sub(" ", s)
    s = SQUARE.sub(r" \1 ", s)
    s = NOTE.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s, n_paren

def n_words(s): return len(WORD.findall(s))

# ---------------- 話者ラベル検出 ----------------
KNOWN = {"Chris Anderson", "Bruno Giussani", "CA", "BG", "Audience", "Narrator", "Video", "Interviewer",
         "Host", "Moderator", "Voice", "Man", "Woman", "Boy", "Girl"}
LABEL = re.compile(r"(?:^|(?<=[\.\!\?\)\"”\u2019]))\s*((?:[A-Z]{2,4})|Chris Anderson|Bruno Giussani|Audience|"
                   r"Narrator|Man|Woman|Boy|Girl|Video|Interviewer|Host|Moderator|Voice|Q|A):\s")

def initials(name):
    parts = [p for p in re.split(r"[ \-]+", re.sub(r"[^A-Za-zÀ-ÿ \-]", "", str(name))) if p]
    return "".join(p[0].upper() for p in parts)

def dialogue_labels(raw, main_speaker):
    """採用するラベル: (a)既知のホスト/聴衆ラベル (b)主話者のイニシャル (c)同一トーク内で2回以上現れる大文字2-4字。
    DNA: のような略語の単発出現を誤検出しないための条件。"""
    c = collections.Counter(LABEL.findall(str(raw)))
    ini = initials(main_speaker)
    keep = {l: n for l, n in c.items() if (l in KNOWN) or (l == ini) or (re.fullmatch(r"[A-Z]{2,4}", l) and n >= 2)}
    return sum(keep.values()), keep

# ---------------- 読み込みと結合 ----------------
meta = pd.read_excel(META, sheet_name="ted_main")
tr = pd.read_excel(TRANS)
n_meta = len(meta)
meta["url"] = meta["url"].astype(str).str.strip()
tr["url"] = tr["url"].astype(str).str.strip()
assert meta["url"].is_unique, "metadata url not unique"

dups = tr[tr["url"].duplicated(keep=False)]
assert (dups.groupby("url")["transcript"].nunique() == 1).all(), "duplicate urls with different text"
tr = tr.drop_duplicates("url")
tr_null = tr[tr["transcript"].isna()]["url"].tolist()
tr = tr[tr["transcript"].notna()]
assert set(tr["url"]) <= set(meta["url"]), "transcript url not in metadata"

df = meta.merge(tr, on="url", how="left")
assert len(df) == n_meta

# ---------------- 評価 (ratings) ----------------
def parse_ratings(s):
    L = ast.literal_eval(s)
    d = {x["name"]: x["count"] for x in L}
    assert sorted(d) == sorted(TAGS), "unexpected tag set"
    return d
rat = df["ratings"].apply(parse_ratings)
for t in TAGS:
    df["votes_" + tagcol(t)] = rat.apply(lambda d: d[t])
vcols = ["votes_" + tagcol(t) for t in TAGS]
df["votes_total"] = df[vcols].sum(axis=1)
for t in TAGS:
    df["pct_" + tagcol(t)] = df["votes_" + tagcol(t)] / df["votes_total"]
assert (df["votes_total"] > 0).all()
assert (df[["pct_" + tagcol(t) for t in TAGS]].sum(axis=1) - 1).abs().max() < 1e-9

# ---------------- 日付 ----------------
for c, new in [("film_date", "film"), ("published_date", "published")]:
    d_ = pd.to_datetime(df[c], unit="s", utc=True)
    df[new + "_date_dt"] = d_.dt.date
    df[new + "_year"] = d_.dt.year

# ---------------- スクリプト整形と指標 ----------------
clean_text, rec = {}, []
for i, r in df.iterrows():
    if pd.isna(r["transcript"]):
        rec.append(dict(i=i)); continue
    raw = str(r["transcript"])
    c, n_par = clean(raw)
    nl, keep = dialogue_labels(raw, str(r["main_speaker"]).strip())
    clean_text[i] = c
    rec.append(dict(i=i, n_words_raw=n_words(raw), n_words_clean=n_words(c), n_stage_directions_removed=n_par,
                    n_laughter=len(LAUGH.findall(raw)), n_applause=len(APPL.findall(raw)),
                    n_music_symbols=len(NOTE.findall(raw)), n_dialogue_labels=nl,
                    dialogue_labels_found="; ".join(f"{k}x{v}" for k, v in sorted(keep.items(), key=lambda kv: -kv[1]))))
X = pd.DataFrame(rec).set_index("i")
df = df.join(X)
df["has_transcript"] = df["transcript"].notna().astype(int)
df["words_per_min"] = df["n_words_clean"] / (df["duration"] / 60)
assert all("(Laughter)" not in c and "(Applause)" not in c for c in clean_text.values())

# ---------------- 分析サンプルの除外条件 (先に当てはまった理由を記録) ----------------
conds = [
    ("no_transcript",        lambda d: d["has_transcript"] == 1),
    ("not_single_individual", lambda d: (d["num_speaker"] == 1) & (d["speaker_type"] == "individual")),
    ("gender_not_male_female", lambda d: d["gender"].isin(["Male", "Female"])),
    ("too_few_words",        lambda d: d["n_words_clean"] >= MIN_WORDS),
    ("low_speech_density",   lambda d: d["words_per_min"] >= MIN_WPM),
    ("music_lyrics",         lambda d: d["n_music_symbols"] < MAX_MUSIC_SYMBOLS),
    ("low_votes",            lambda d: d["votes_total"] >= MIN_VOTES),
    ("dialogue_labels",      lambda d: d["n_dialogue_labels"] < MAX_DIALOGUE_LABELS),
]
alive = pd.Series(True, index=df.index)
reason = pd.Series("", index=df.index)
flow = []
for name, f in conds:
    ok = f(df).fillna(False)
    drop = alive & ~ok
    reason[drop] = name
    alive = alive & ok
    g = df.loc[alive, "gender"]
    flow.append((name, int(drop.sum()), int(alive.sum()), int((g == "Male").sum()), int((g == "Female").sum())))
df["exclusion_reason"] = reason
df["in_sample_strict"] = (reason == "").astype(int)
df["in_sample_wide"] = reason.isin(["", "dialogue_labels"]).astype(int)

# ---------------- 列の並び ----------------
orig_cols = [c for c in meta.columns]
derived = (["film_date_dt", "film_year", "published_date_dt", "published_year", "has_transcript",
            "n_words_raw", "n_words_clean", "words_per_min", "n_stage_directions_removed", "n_laughter",
            "n_applause", "n_music_symbols", "n_dialogue_labels", "dialogue_labels_found"]
           + vcols + ["votes_total"] + ["pct_" + tagcol(t) for t in TAGS]
           + ["exclusion_reason", "in_sample_strict", "in_sample_wide"])
out = df[orig_cols + derived].copy()

# ---------------- 整形済み本文をCSVへ ----------------
csvdf = pd.DataFrame({"url": df.loc[list(clean_text), "url"].values,
                      "main_speaker": df.loc[list(clean_text), "main_speaker"].astype(str).str.strip().values,
                      "title": df.loc[list(clean_text), "title"].values,
                      "transcript_clean": list(clean_text.values())})
csvdf.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")
chk = pd.read_csv(OUT_CSV, encoding="utf-8-sig", keep_default_na=False)   # 整形後が空(演奏のみ)の行は空文字のまま読む
assert len(chk) == len(csvdf) and (chk["transcript_clean"] == csvdf["transcript_clean"]).all(), "CSV round-trip mismatch"
assert (chk["url"] == csvdf["url"]).all()

# ---------------- Excel 出力 ----------------
F = Font(name="Arial", size=10); FB = Font(name="Arial", size=10, bold=True)
FILL = PatternFill("solid", fgColor="D9E1F2")
wb = Workbook()

# README
ws0 = wb.active; ws0.title = "README"

# ted_merged
ws = wb.create_sheet("ted_merged")
cols = list(out.columns)
ws.append(cols)
for c in ws[1]: c.font = FB; c.fill = FILL; c.alignment = Alignment(wrap_text=True, vertical="top")
ci = {c: j + 1 for j, c in enumerate(cols)}
v0, v1 = get_column_letter(ci[vcols[0]]), get_column_letter(ci[vcols[-1]])
tot_col = get_column_letter(ci["votes_total"])
def py(v):
    if v is None: return None
    if isinstance(v, float) and pd.isna(v): return None
    if hasattr(v, "item") and not isinstance(v, (dt.date,)): return v.item()
    return v
for r_i, (_, row) in enumerate(out.iterrows(), start=2):
    vals = []
    for c in cols:
        if c == "votes_total": vals.append(f"=SUM({v0}{r_i}:{v1}{r_i})")
        elif c.startswith("pct_"):
            vc = get_column_letter(ci["votes_" + c[4:]])
            vals.append(f"={vc}{r_i}/${tot_col}{r_i}")
        else: vals.append(py(row[c]))
    ws.append(vals)
for row in ws.iter_rows(min_row=2):
    for c in row: c.font = F
for c in ("film_date_dt", "published_date_dt"):
    for cell in ws[get_column_letter(ci[c])][1:]: cell.number_format = "yyyy-mm-dd"
for c in cols:
    if c.startswith("pct_"):
        for cell in ws[get_column_letter(ci[c])][1:]: cell.number_format = "0.0%"
for c in ("words_per_min",):
    for cell in ws[get_column_letter(ci[c])][1:]: cell.number_format = "0.0"
for j, c in enumerate(cols, 1):
    ws.column_dimensions[get_column_letter(j)].width = 14
for c, w in [("main_speaker", 24), ("title", 40), ("description", 40), ("url", 50), ("gender_note", 30),
             ("dialogue_labels_found", 24), ("exclusion_reason", 22)]:
    ws.column_dimensions[get_column_letter(ci[c])].width = w
ws.freeze_panes = ws.cell(2, ci["main_speaker"] + 1)
ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"

# sample_flow
ws1 = wb.create_sheet("sample_flow")
ws1.append(["step", "exclusion_reason (first failing)", "criterion", "n_excluded", "n_remaining", "n_male", "n_female"])
desc = {"no_transcript": "スクリプトがない (結合できず/空)",
        "not_single_individual": "num_speaker!=1 または 団体・複数話者・人間以外",
        "gender_not_male_female": "gender が Male/Female 以外 (Unknown/Mixed)",
        "too_few_words": f"整形後の語数 < {MIN_WORDS}",
        "low_speech_density": f"words_per_min < {MIN_WPM}",
        "music_lyrics": f"♫♪ の個数 >= {MAX_MUSIC_SYMBOLS}",
        "low_votes": f"評価の総票数 < {MIN_VOTES}",
        "dialogue_labels": f"話者ラベル数 >= {MAX_DIALOGUE_LABELS} (厳密版のみ除外。広い版には含む)"}
ws1.append([0, "(start)", "ted_main の全トーク", None, n_meta, None, None])
for k, (name, nd, nr, nm, nf) in enumerate(flow, 1):
    ws1.append([k, name, desc[name], nd, nr, nm if k >= 3 else None, nf if k >= 3 else None])
for c in ws1[1]: c.font = FB; c.fill = FILL
for row in ws1.iter_rows(min_row=2):
    for c in row: c.font = F
for j, w in enumerate([6, 30, 60, 12, 13, 10, 10], 1): ws1.column_dimensions[get_column_letter(j)].width = w
ws1.append([]); ws1.append(["", "in_sample_strict (主分析用)", "上の全条件を満たす", "", int(df.in_sample_strict.sum())])
ws1.append(["", "in_sample_wide (感度分析用)", "dialogue_labels 以外の条件を満たす", "", int(df.in_sample_wide.sum())])
for r in (ws1.max_row - 1, ws1.max_row):
    for c in ws1[r]: c.font = FB

# no_transcript_list
ws2 = wb.create_sheet("no_transcript_list")
nt = df[df["has_transcript"] == 0][["main_speaker", "title", "event", "num_speaker", "gender", "duration", "url"]]
ws2.append(["main_speaker", "title", "event", "num_speaker", "gender", "duration(sec)", "url", "reason"])
for _, r in nt.iterrows():
    ws2.append([str(r.main_speaker).strip(), r.title, r.event, int(r.num_speaker), r.gender, int(r.duration), r.url,
                "スクリプトが空(NaN)" if r.url in tr_null else "transcripts_V2 に該当urlなし"])
for c in ws2[1]: c.font = FB; c.fill = FILL
for row in ws2.iter_rows(min_row=2):
    for c in row: c.font = F
for j, w in enumerate([26, 50, 22, 11, 10, 13, 60, 28], 1): ws2.column_dimensions[get_column_letter(j)].width = w
ws2.freeze_panes = "A2"

# dialogue_flagged
ws3 = wb.create_sheet("dialogue_flagged")
dg = df[(df["n_dialogue_labels"] >= MAX_DIALOGUE_LABELS)].sort_values("n_dialogue_labels", ascending=False)
ws3.append(["main_speaker", "title", "n_dialogue_labels", "dialogue_labels_found", "n_words_clean", "gender", "exclusion_reason", "url"])
for _, r in dg.iterrows():
    ws3.append([str(r.main_speaker).strip(), r.title, int(r.n_dialogue_labels), r.dialogue_labels_found,
                int(r.n_words_clean), r.gender, r.exclusion_reason, r.url])
for c in ws3[1]: c.font = FB; c.fill = FILL
for row in ws3.iter_rows(min_row=2):
    for c in row: c.font = F
for j, w in enumerate([26, 50, 12, 30, 12, 10, 24, 60], 1): ws3.column_dimensions[get_column_letter(j)].width = w
ws3.freeze_panes = "A2"

# README 本文
n_f_dlg = int((dg.gender == "Female").sum()); n_m_dlg = int((dg.gender == "Male").sum())
R = [
 ("このファイルについて", None),
 ("ted_merged", "1行=1トーク(2,550行)。ted_main_V2_gender の全列 + 下記の追加列。スクリプト本文は含まない(Excelの1セル=32,767文字の上限を超えるトークが2件あるため)。本文は ted_transcripts_clean_V2.csv(url で結合)。"),
 ("sample_flow", "2,550件から分析サンプルまでの絞り込みの流れ(論文のサンプル選択の記述に使える)。"),
 ("no_transcript_list", "スクリプトを結合できなかった87件の一覧と理由。"),
 ("dialogue_flagged", "話者ラベルが閾値以上のトーク(対談・質疑形式の疑い)の一覧。目視確認用。"),
 ("", None),
 ("元データの修正", None),
 ("url", "両ファイルとも全行の末尾に改行が付いていたため除去して結合。"),
 ("transcripts の重複", "url重複3件(Brené Brown / Jonathan Haidt / Rob Reid)は本文が完全一致のため1件に集約。"),
 ("transcripts の欠損", "Yuval Noah Harari のトーク1件はスクリプトが空(NaN)。"),
 ("", None),
 ("追加列", None),
 ("film_date_dt / published_date_dt / *_year", "Unix時間(秒,UTC)を日付と年に変換。"),
 ("has_transcript", "スクリプトあり=1。"),
 ("n_words_raw / n_words_clean", "単語数(英数字の連続、アポストロフィ・ハイフン内包を1語)。raw=原文、clean=整形後。"),
 ("words_per_min", "n_words_clean / (duration/60)。"),
 ("n_stage_directions_removed", "除去した丸括弧(入れ子でないもの)の個数。"),
 ("n_laughter / n_applause", "(Laughter) / (Applause) の個数(原文)。Funnyなど評価の統制変数に使える。"),
 ("n_music_symbols", "♫♪ の個数(原文)。歌詞混入の目安。"),
 ("n_dialogue_labels / dialogue_labels_found", "話者ラベルの検出数と内訳(例 CAx38; EMx40)。検出対象は既知のホスト・聴衆ラベル(Chris Anderson, CA, Bruno Giussani, BG, Audience, Narrator 等)、主話者のイニシャル、同一トーク内で2回以上現れる大文字2-4字。ヒューリスティックなので誤検出・見逃しがあり得る。"),
 ("votes_<tag> / votes_total", "14種類の評価タグの票数と総票数(Excel数式 =SUM)。"),
 ("pct_<tag>", "各タグの票数 / 総票数(Excel数式)。従属変数の候補。14種の割合の合計は1。"),
 ("exclusion_reason", "分析サンプルから除外される場合、最初に当てはまった理由(空欄=除外なし)。順序はsample_flow参照。"),
 ("in_sample_strict", "主分析用サンプル=1(全条件を満たす)。"),
 ("in_sample_wide", "感度分析用サンプル=1(対談ラベル以外の全条件を満たす。対談形式のトークを含む)。"),
 ("", None),
 ("整形ルール(本文)", None),
 ("1", "丸括弧(...)を除去: (Laughter),(Applause),(Music) 等。ただし括弧内に実際の発話が入る例外が少数ある(6語以上で約90個/全18,244個)ため、完全には区別していない。"),
 ("2", "[unclear] は除去。[the] のような補足語は括弧だけ外して語を残す。"),
 ("3", "♫♪ の記号を除去(歌詞の語は残る)。"),
 ("4", "話者ラベルと対談相手の発話は除去していない(誤って主話者の発話を削るリスクを避けるため)。フラグで識別する。"),
 ("", None),
 ("注意点", None),
 ("対談形式", f"n_dialogue_labels>={MAX_DIALOGUE_LABELS} のトークは{len(dg)}件。性別の内訳は男性{n_m_dlg}・女性{n_f_dlg}。除外が性別に偏っていないかは分析前に確認し、広いサンプル(in_sample_wide)での感度分析を推奨。"),
 ("閾値", f"語数>={MIN_WORDS}, 話速>={MIN_WPM}語/分, ♫<{MAX_MUSIC_SYMBOLS}, 総票数>={MIN_VOTES}, 話者ラベル<{MAX_DIALOGUE_LABELS} は暫定値。変更する場合は preprocess_ted.py の定数を直して再実行。"),
 ("評価割合", "割合は14タグ全体の票数に対する比。タグ間で合計が1になるため、1つの評価が高いと他が低くなる構造に注意。"),
]
for i, (a, b) in enumerate(R, 1):
    ws0.cell(i, 1, a)
    if b is None: ws0.cell(i, 1).font = FB
    else:
        ws0.cell(i, 1).font = FB; ws0.cell(i, 2, b).font = F
        ws0.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top")
ws0.column_dimensions["A"].width = 38; ws0.column_dimensions["B"].width = 120
wb.save(OUT_XLSX)

# ---------------- ログ ----------------
print("rows:", len(out), "| with transcript:", int(df.has_transcript.sum()))
print("flow:"); [print("  ", f) for f in flow]
print("strict:", int(df.in_sample_strict.sum()), " wide:", int(df.in_sample_wide.sum()))
print("dialogue-flagged (>=%d): %d  M=%d F=%d" % (MAX_DIALOGUE_LABELS, len(dg), n_m_dlg, n_f_dlg))
pool = df[(df.exclusion_reason.isin(["", "dialogue_labels"]))]
print("female share (wide pool): %.1f%% | in dialogue: %.1f%% | not dialogue: %.1f%%" % (
    100 * (pool.gender == "Female").mean(),
    100 * (pool[pool.exclusion_reason == "dialogue_labels"].gender == "Female").mean(),
    100 * (pool[pool.exclusion_reason == ""].gender == "Female").mean()))
