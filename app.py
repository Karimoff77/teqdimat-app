"""
APP.PY — Manual rejim (Streamlit). Məqsəd: testerlərə verilə bilən ilk işlək forma.
İşə sal:  streamlit run app.py
Test üçün rəy linki (istəyə bağlı):  FEEDBACK_URL=https://... streamlit run app.py
"""
import hashlib, json, os, re, sys, time
import pandas as pd
import streamlit as st

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(ROOT, "src"))
import kataloq, manual, engine, onizleme

def _rey_linki():
    v = os.environ.get("FEEDBACK_URL", "")
    if not v:
        try:                                   # Streamlit Cloud: Settings -> Secrets
            v = str(st.secrets.get("FEEDBACK_URL", ""))
        except Exception:
            v = ""
    return v if v.startswith("https://") else ""


FEEDBACK_URL = _rey_linki()
LOG = os.path.join(ROOT, "logs", "istifade.jsonl")

st.set_page_config(page_title="Təqdimat — Manual rejim", page_icon="🧩", layout="wide")


# ------------------------------------------------------------------ köməkçilər
def img(path, **kw):
    if os.path.exists(path):
        st.image(path, **kw)


def log_hadise(d):
    """Yalnız say və kateqoriya yazılır; istifadəçinin mətni YAZILMIR."""
    try:
        os.makedirs(os.path.dirname(LOG), exist_ok=True)
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(json.dumps({"ts": int(time.time()), **d}, ensure_ascii=False) + "\n")
    except OSError:
        pass


def df_from(rows, cols, n_min=0):
    rows = [list(r) + [""] * (len(cols) - len(r)) for r in (rows or [])]
    rows = [[str(c) for c in r[:len(cols)]] for r in rows]
    while len(rows) < n_min:
        rows.append([""] * len(cols))
    return pd.DataFrame(rows, columns=cols, dtype="object")


def rows_from(df, ncols):
    out = []
    for _, r in df.iterrows():
        v = ["" if (c is None or (isinstance(c, float) and pd.isna(c))) else str(c).strip() for c in list(r)[:ncols]]
        if any(v):
            out.append(v)
    return out


def init_df(sid, ad, make):
    """data_editor-a həmişə EYNİ başlanğıc cədvəl verilir (yoxsa redaktələr təkrarlanır)."""
    k = f"init:{sid}:{ad}"
    if k not in st.session_state:
        st.session_state[k] = make()
    return st.session_state[k]


def lbl(f):
    return f["etiket"] + (" — istəyə bağlı" if f.get("ixtiyari") else "")


# ------------------------------------------------------------------ vəziyyət
if "slaydlar" not in st.session_state:
    st.session_state.slaydlar = [manual.yeni("giris", "titul"), manual.yeni("sonluq", "tesekkur")]
    st.session_state.acik = None
    st.session_state.cixis = None

S = st.session_state.slaydlar
for _k, _v in (("movzu", "Təqdimatın mövzusu"), ("muellif", ""), ("tema", "auto"), ("qarisiq", False)):
    st.session_state.setdefault(_k, _v)


def _idx(sid):
    return next(i for i, r in enumerate(S) if r["id"] == sid)


def cb_sil(sid):
    del S[_idx(sid)]


def cb_yuxari(sid):
    i = _idx(sid)
    if i > 0:
        S[i - 1], S[i] = S[i], S[i - 1]


def cb_asagi(sid):
    i = _idx(sid)
    if i < len(S) - 1:
        S[i + 1], S[i] = S[i], S[i + 1]


def cb_elave(kid, tip, variant):
    rec = manual.yeni(kid, tip, variant)
    # Sonluq slaydı sonda qalsın: yeni slayd ondan əvvələ düşür
    if S and S[-1]["kat"] == "sonluq" and kid != "sonluq":
        S.insert(len(S) - 1, rec)
    else:
        S.append(rec)
    st.session_state.acik = rec["id"]


def cb_yukle():
    """Layihə faylı yüklənəndə: slaydları və ayarları əvəz edir (widget-lər yenidən qurulmazdan əvvəl)."""
    f = st.session_state.get("yukle")
    if f is None:
        return
    try:
        if f.size > 2_000_000:
            raise ValueError("fayl çox böyükdür")
        ay, recs, atilan = manual.layihe_yukle(f.getvalue().decode("utf-8"))
    except (ValueError, UnicodeDecodeError) as e:
        st.session_state.yukle_mesaj = ("error", f"Yüklənmədi: {e}")
        return
    st.session_state.slaydlar[:] = recs
    for k in [k for k in st.session_state if str(k).startswith("init:")]:
        del st.session_state[k]
    for k in ("movzu", "muellif"):
        if isinstance(ay.get(k), str):
            st.session_state[k] = ay[k]
    if ay.get("tema") in ["auto"] + sorted(engine.THEMES):
        st.session_state["tema"] = ay["tema"]
    if isinstance(ay.get("qarisiq"), bool):
        st.session_state["qarisiq"] = ay["qarisiq"]
    st.session_state.cixis = None
    st.session_state.acik = None
    st.session_state.yukle_mesaj = ("success", f"Yükləndi: {len(recs)} slayd" + (f" ({atilan} slayd oxunmadı)" if atilan else ""))


# ------------------------------------------------------------------ yan panel: ayarlar
with st.sidebar:
    st.header("Təqdimat ayarları")
    movzu = st.text_input("Mövzu (hər slaydın altında görünür)", key="movzu")
    muellif = st.text_input("Fənn / müəllif (istəyə bağlı)", key="muellif")
    temalar = ["auto"] + sorted(engine.THEMES)
    tema = st.selectbox("Rəng teması", temalar, key="tema", format_func=lambda t: "Təsadüfi" if t == "auto" else t.capitalize())
    if tema != "auto":
        img(os.path.join(ROOT, "assets", "tema", f"{tema}.png"))
    qarisiq = st.checkbox("Üslubu qarışdır (künc, başlıq, dekor)", key="qarisiq",
                          help="Söndürülüdürsə, temanın öz üslubu işləyir. Qarışdıranda hər təqdimat başqa görünür.")
    st.divider()
    st.subheader("Layihə")
    ad_layihe = re.sub(r"[^\w\-]+", "_", movzu.strip(), flags=re.UNICODE).strip("_")[:40] or "layihe"
    st.download_button("💾 Layihəni saxla (.json)",
                       manual.layihe_json(S, {"movzu": movzu, "muellif": muellif, "tema": tema, "qarisiq": qarisiq}),
                       file_name=f"{ad_layihe}.json", mime="application/json",
                       help="Sonra davam etmək üçün. Brauzeri bağlayanda iş itir, bu fayl isə qalır.")
    st.file_uploader("Saxlanmış layihəni yüklə", type=["json"], key="yukle", on_change=cb_yukle)
    _m = st.session_state.pop("yukle_mesaj", None)
    if _m:
        (st.error if _m[0] == "error" else st.success)(_m[1])

# ------------------------------------------------------------------ başlıq
st.title("🧩 Təqdimat hazırla")
st.caption("Slaydları kateqoriya üzrə seç, məzmunu yaz, hazır `.pptx` faylını yüklə. Fayl PowerPoint-də tam redaktə olunur.")

sol, sag = st.columns([2, 3], gap="large")

# ------------------------------------------------------------------ sol: yeni slayd
with sol:
    st.subheader("➕ Yeni slayd")
    kats = kataloq.KATEQORIYALAR
    kat = st.selectbox("Kateqoriya", kats, format_func=lambda k: k["ad"])
    st.caption(kat["izah"])
    if len(kat["altlar"]) > 1:
        alt = st.selectbox("Növ", kat["altlar"], format_func=lambda a: a["ad"])
    else:
        alt = kat["altlar"][0]
    variants = alt["variantlar"]
    if len(variants) > 1:
        v = st.radio("Düzülüş", variants, format_func=lambda x: kataloq.variant_ad(alt["tip"], x), horizontal=True)
    else:
        v = variants[0]
    img(os.path.join(ROOT, "assets", "variant", f"{alt['tip']}_{v}.png"))
    st.caption("Nümunə görünüş. Rəng seçdiyin temaya görə dəyişəcək.")
    st.button("Slayd əlavə et", type="primary", on_click=cb_elave, args=(kat["id"], alt["tip"], v), use_container_width=True)

# ------------------------------------------------------------------ sağ: slayd siyahısı
with sag:
    st.subheader(f"📑 Təqdimat ({len(S)} slayd)")
    if not S:
        st.info("Hələ slayd yoxdur. Soldan əlavə et.")
    for i, rec in enumerate(S):
        sid = rec["id"]
        alt_r = manual.alt_tap(rec["kat"], rec["tip"])
        ad_k = manual.kat_tap(rec["kat"])["ad"]
        nun = manual.numunedir(rec)
        baslik = f"{i + 1}. {ad_k} — {kataloq.variant_ad(rec['tip'], rec['variant'])}" + ("   ⚠ nümunə mətn" if nun else "")
        with st.expander(baslik, expanded=(sid == st.session_state.acik)):
            c1, c2, c3, c4 = st.columns([1, 1, 1, 4])
            c1.button("↑", key=f"up{sid}", on_click=cb_yuxari, args=(sid,), help="Yuxarı")
            c2.button("↓", key=f"dn{sid}", on_click=cb_asagi, args=(sid,), help="Aşağı")
            c3.button("🗑", key=f"del{sid}", on_click=cb_sil, args=(sid,), help="Sil")
            if len(alt_r["variantlar"]) > 1:
                rec["variant"] = c4.selectbox("Düzülüş", alt_r["variantlar"], key=f"var{sid}",
                                              index=alt_r["variantlar"].index(rec["variant"]),
                                              format_func=lambda x, t=rec["tip"]: kataloq.variant_ad(t, x))
            d = rec["data"]
            for f in alt_r["sahələr"]:
                ad, key, nv = f["ad"], f"{sid}:{f['ad']}", f["növ"]
                if nv == "str":
                    d[ad] = st.text_input(f"{lbl(f)}  (≤{f['max']})", d.get(ad, ""), key=key)
                elif nv == "sətirlər":
                    if ad == "kateqoriyalar":      # qrafik widget-ində idarə olunur
                        continue
                    txt = st.text_area(f"{lbl(f)}  (hər sətir ayrı; ən çox {f['say']}, hər biri ≤{f['max']})",
                                       "\n".join(d.get(ad, [])), key=key, height=40 + 24 * min(f["say"], 5))
                    d[ad] = [x.strip() for x in txt.splitlines() if x.strip()]
                elif nv in ("cütlər", "üçlüklər"):
                    cols = list(f["sütunlar"])
                    st.markdown(f"**{lbl(f)}** — ən çox {f['say']} sətir")
                    df = st.data_editor(init_df(sid, ad, lambda: df_from(d.get(ad), cols)), key=key, num_rows="dynamic",
                                        hide_index=True)
                    d[ad] = rows_from(df, len(cols))
                elif nv == "cədvəl":
                    cols = ["Sütun 1", "Sütun 2", "Sütun 3"]
                    st.markdown(f"**{lbl(f)}** — ən çox {f['say']} sətir; başlıq sayı qədər sütun işlənir")
                    df = st.data_editor(init_df(sid, ad, lambda: df_from(d.get(ad), cols)), key=key, num_rows="dynamic",
                                        hide_index=True)
                    d[ad] = rows_from(df, 3)
                elif nv == "seriyalar":
                    st.markdown(f"**{lbl(f)}**")
                    ser = d.get("setirler", [])
                    ad_cols = st.columns(3)
                    adlar = [ad_cols[j].text_input(f"Seriya {j + 1} adı", ser[j]["ad"] if j < len(ser) else "", key=f"{key}:ad{j}")
                             for j in range(3)]
                    cats0 = d.get("kateqoriyalar", [])

                    def mk():
                        rows = []
                        for r, c in enumerate(cats0):
                            row = [c] + [str(ser[j]["qiymetler"][r]) if j < len(ser) and r < len(ser[j]["qiymetler"]) else "" for j in range(3)]
                            rows.append(row)
                        return df_from(rows, ["X dəyəri", "Seriya 1", "Seriya 2", "Seriya 3"])
                    df = st.data_editor(init_df(sid, ad, mk), key=key, num_rows="dynamic", hide_index=True)
                    rows = rows_from(df, 4)
                    cats = [r[0] for r in rows if r[0]]
                    d["kateqoriyalar"] = cats
                    d["setirler"] = [{"ad": adlar[j], "qiymetler": [r[j + 1] for r in rows if r[0]]}
                                     for j in range(3) if adlar[j].strip()]
            # xəbərdarlıqlar (dogrula.py ilə eyni yoxlama)
            if nun:
                st.warning("Bu slaydda hələ nümunə mətn var. Öz məzmununla əvəz et.")
            for w in manual.xeberdarliqlar(rec):
                st.warning(w)

# ------------------------------------------------------------------ yarat və yüklə
st.divider()
spec = manual.spec_qur(S, tema, movzu, muellif, qarisiq, seed=movzu)
h = hashlib.md5(json.dumps(spec, sort_keys=True, ensure_ascii=False, default=str).encode()).hexdigest()
n_nun = sum(manual.numunedir(r) for r in S)

if n_nun:
    st.warning(f"{n_nun} slaydda nümunə mətn dəyişdirilməyib. Faylı yaratmazdan əvvəl onları yenilə.")

ok_onizle, sebeb = onizleme.mumkundur()
if st.button("📥 Hazırla və önizlə", type="primary", disabled=not S):
    with st.spinner("Təqdimat hazırlanır..."):
        b, rep, warn = manual.pptx_bytes(spec)
        png, png_xeta = None, ""
        if ok_onizle:
            try:
                png = onizleme.pptx_png(b)
            except onizleme.OnizlemeXetasi as e:
                png_xeta = str(e)
    st.session_state.cixis = {"b": b, "rep": rep, "warn": warn, "h": h, "n": len(S), "png": png, "png_xeta": png_xeta,
                              "adlar": [f"{r_i + 1}. {manual.kat_tap(r['kat'])['ad']}" for r_i, r in enumerate(S)]}

cx = st.session_state.cixis
if cx:
    if cx["h"] != h:
        st.info("Təqdimatda dəyişiklik olub. Yenilənmiş fayl və önizləmə üçün «Hazırla və önizlə» düyməsini yenidən bas.")
    ad_fayl = re.sub(r"[^\w\-]+", "_", movzu.strip(), flags=re.UNICODE).strip("_")[:40] or "teqdimat"
    st.success(f"Hazırdır: {cx['n']} slayd.")
    st.download_button("⬇️ .pptx yüklə", cx["b"], file_name=f"{ad_fayl}.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                       on_click=log_hadise,
                       args=({"slayd": cx["n"], "kateqoriyalar": [r["kat"] for r in S], "tema": tema,
                              "temir": len(cx["rep"]), "numune_qalan": n_nun, "qarisiq": qarisiq,
                              "onizleme": bool(cx["png"])},))
    if cx["rep"]:
        with st.expander(f"{len(cx['rep'])} avtomatik düzəliş edildi (uzun mətn və s.)"):
            for r in cx["rep"]:
                st.write("•", r)
    for w in cx["warn"]:
        st.caption("! " + w)
    if cx["png"]:
        st.subheader("👁 Önizləmə")
        st.caption("Önizləmə LibreOffice ilə çəkilir. PowerPoint-də şriftlər və aralıqlar bir qədər fərqli ola bilər. Böyütmək üçün şəklin üstünə gəl.")
        for r0 in range(0, len(cx["png"]), 3):
            cols = st.columns(3)
            for c, j in zip(cols, range(r0, min(r0 + 3, len(cx["png"])))):
                c.image(cx["png"][j], caption=cx["adlar"][j] if j < len(cx["adlar"]) else None)
    elif ok_onizle and cx["png_xeta"]:
        st.warning(f"Önizləmə alınmadı ({cx['png_xeta']}). Fayl isə hazırdır, yükləyə bilərsən.")
    elif not ok_onizle:
        st.info("Bu versiyada slayd önizləməsi yoxdur. Fayl hazırdır: yüklə və PowerPoint-də aç, orada redaktə edə bilərsən.")
    if FEEDBACK_URL:
        st.link_button("2 dəqiqəlik rəy ver", FEEDBACK_URL)
