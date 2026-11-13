"""
APP.PY — Manual rejim (Streamlit). Məqsəd: testerlərə verilə bilən ilk işlək forma.
İşə sal:  streamlit run app.py
Test üçün rəy linki (istəyə bağlı):  FEEDBACK_URL=https://... streamlit run app.py
"""
import colorsys, hashlib, json, os, re, sys, time
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

st.set_page_config(page_title="Təqdimat — Manual rejim", page_icon="🧩", layout="wide", initial_sidebar_state="collapsed")

YAZI = {"Standart": 1.0, "Böyük": 1.15, "Çox böyük": 1.3}       # yazı ölçüsü çarpanı


def tema_emoji(t):
    """Temanın əsas rənginə uyğun dairə (seçim düymələrində rəngi göstərir)."""
    r, g, b = [c / 255 for c in engine.THEMES[t]["accent"]]
    h, l, sat = colorsys.rgb_to_hls(r, g, b)
    if sat < 0.15:
        return "⚫" if l < 0.5 else "⚪"
    h *= 360
    return "🔴" if (h < 20 or h >= 345) else "🟠" if h < 45 else "🟡" if h < 70 else "🟢" if h < 170 else "🔵" if h < 260 else "🟣"


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
    st.session_state.acik_set = set()
    st.session_state.cixis = None

S = st.session_state.slaydlar
for _k, _v in (("movzu", "Təqdimatın mövzusu"), ("muellif", ""), ("tema", "auto"), ("qarisiq", False), ("yazi", "Standart"), ("kat_sec", "giris")):
    st.session_state.setdefault(_k, _v)


def _idx(sid):
    return next(i for i, r in enumerate(S) if r["id"] == sid)


def _unut(sid):
    """Slayd bağlananda/silinəndə cədvəl redaktorlarının başlanğıc şəkli təmizlənir: açılanda cari məzmundan yenidən qurulur."""
    for k in [k for k in st.session_state if str(k).startswith(f"init:{sid}:")]:
        del st.session_state[k]


def cb_toggle(sid):
    a = st.session_state.acik_set
    if sid in a:
        a.discard(sid)
        _unut(sid)
    else:
        a.add(sid)


def cb_sil(sid):
    del S[_idx(sid)]
    st.session_state.acik_set.discard(sid)
    _unut(sid)


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
    st.session_state.acik_set.add(rec["id"])


def cb_variant(sid):
    """Düzülüş çipi dəyişəndə: dərhal slaydın qeydinə yazılır (başlıq və şəkil eyni dəqiqədə yenilənir)."""
    rec = st.session_state.slaydlar[_idx(sid)]
    v = st.session_state.get(f"var{sid}")
    if v in manual.alt_tap(rec["kat"], rec["tip"])["variantlar"]:
        rec["variant"] = v
    else:                                   # çip söndürülübsə əvvəlki seçim qalsın
        st.session_state[f"var{sid}"] = rec["variant"]


def cb_dlg(delta):
    n = len(st.session_state.cixis["png"])
    st.session_state.dlg_i = (st.session_state.dlg_i + delta) % n


@st.dialog("Slayd önizləməsi", width="large")
def boyut():
    cx = st.session_state.cixis
    i = st.session_state.dlg_i
    st.image(cx["png"][i], caption=f"{cx['adlar'][i]}  ({i + 1}/{len(cx['png'])})")
    a, b = st.columns(2)
    a.button("◀ Əvvəlki", key="dlg_prev", on_click=cb_dlg, args=(-1,), use_container_width=True)
    b.button("Növbəti ▶", key="dlg_next", on_click=cb_dlg, args=(1,), use_container_width=True)


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
    _yz = [k for k, v in YAZI.items() if v == ay.get("yazi")]
    if _yz:
        st.session_state["yazi"] = _yz[0]
    st.session_state.cixis = None
    st.session_state.acik_set = set()
    st.session_state.yukle_mesaj = ("success", f"Yükləndi: {len(recs)} slayd" + (f" ({atilan} slayd oxunmadı)" if atilan else ""))


# ------------------------------------------------------------------ başlıq
st.markdown("""<style>
/* şəkillərin üstündəki kiçik «tam ekran» düyməsi lazım deyil: böyütmək üçün «🔍 Böyüt» düyməsi var */
[data-testid="stElementContainer"]:has([data-testid="stImage"]) [data-testid="stElementToolbar"] { display: none !important; }
</style>""", unsafe_allow_html=True)
st.title("🧩 Təqdimat hazırla")
st.caption("Slaydları seç, məzmunu yaz, hazır `.pptx` faylını yüklə. Fayl PowerPoint-də tam redaktə olunur.")

# ------------------------------------------------------------------ 1. ayarlar
with st.container(border=True):
    st.markdown("### 1. Təqdimatın ayarları")
    c1, c2 = st.columns(2)
    movzu = c1.text_input("Mövzu (hər slaydın altında görünür)", key="movzu")
    muellif = c2.text_input("Fənn / müəllif (istəyə bağlı)", key="muellif")
    temalar = ["auto"] + sorted(engine.THEMES)
    tema = st.pills("Rəng teması", temalar, key="tema", selection_mode="single", wrap=True,
                    format_func=lambda t: "🎲 Təsadüfi" if t == "auto" else f"{tema_emoji(t)} {t.capitalize()}") or "auto"
    if tema != "auto":
        img(os.path.join(ROOT, "assets", "tema", f"{tema}.png"), width=340)
    c3, c4 = st.columns(2)
    yazi_ad = c3.pills("Yazı ölçüsü", list(YAZI), key="yazi", selection_mode="single", wrap=True) or "Standart"
    qarisiq = c4.checkbox("Üslubu qarışdır (künc, başlıq, dekor)", key="qarisiq",
                          help="Söndürülüdürsə, temanın öz üslubu işləyir. Qarışdıranda hər təqdimat başqa görünür.")

# ------------------------------------------------------------------ 2. yeni slayd
sol, sag = st.columns([2, 3], gap="large")
with sol:
    st.markdown("### 2. Slayd əlavə et")
    kats = kataloq.KATEQORIYALAR
    ids = [k["id"] for k in kats]
    sec = st.pills("Slayd növü", ids, key="kat_sec", selection_mode="single", wrap=True,
                   format_func=lambda i: f"{kataloq.ikon(i)} {manual.kat_tap(i)['ad']}")
    kat = manual.kat_tap(sec or st.session_state.get("kat_son") or ids[0])
    st.session_state["kat_son"] = kat["id"]
    st.caption(kat["izah"])
    if len(kat["altlar"]) > 1:
        alt_tip = st.pills("Növ", [a["tip"] for a in kat["altlar"]], key=f"alt_{kat['id']}", selection_mode="single", wrap=True,
                           default=kat["altlar"][0]["tip"], format_func=lambda t, k=kat: next(a["ad"] for a in k["altlar"] if a["tip"] == t))
        alt = next((a for a in kat["altlar"] if a["tip"] == alt_tip), kat["altlar"][0])
    else:
        alt = kat["altlar"][0]
    variants = alt["variantlar"]
    v = variants[0]
    img(os.path.join(ROOT, "assets", "variant", f"{alt['tip']}_{v}.png"))
    st.caption("Nümunə görünüş. Rəng seçdiyin temaya görə dəyişəcək." +
               (" Düzülüşü slayd əlavə olunandan sonra slaydın içində dəyişə bilərsən." if len(variants) > 1 else ""))
    st.button("➕ Slayd əlavə et", type="primary", on_click=cb_elave, args=(kat["id"], alt["tip"], v), use_container_width=True)

# ------------------------------------------------------------------ 3. slaydlar
with sag:
    st.markdown(f"### 3. Slaydların ({len(S)})")
    if not S:
        st.info("Hələ slayd yoxdur. Soldan əlavə et.")
    for i, rec in enumerate(S):
        sid = rec["id"]
        alt_r = manual.alt_tap(rec["kat"], rec["tip"])
        ad_k = manual.kat_tap(rec["kat"])["ad"]
        nun = manual.numunedir(rec)
        baslik = f"{i + 1}. {ad_k} — {kataloq.variant_ad(rec['tip'], rec['variant'])}" + ("   ⚠ nümunə mətn" if nun else "")
        acik = sid in st.session_state.acik_set
        with st.container(border=True):
            st.button(f"{'▾' if acik else '▸'} {baslik}", key=f"hdr{sid}", on_click=cb_toggle, args=(sid,), use_container_width=True)
            if not acik:
                continue
            c1, c2, c3, _ = st.columns([1, 1, 1, 4])
            c1.button("↑", key=f"up{sid}", on_click=cb_yuxari, args=(sid,), help="Yuxarı")
            c2.button("↓", key=f"dn{sid}", on_click=cb_asagi, args=(sid,), help="Aşağı")
            c3.button("🗑", key=f"del{sid}", on_click=cb_sil, args=(sid,), help="Sil")
            if len(alt_r["variantlar"]) > 1:
                st.pills("Düzülüş", alt_r["variantlar"], key=f"var{sid}", selection_mode="single", wrap=True,
                         default=rec["variant"], on_change=cb_variant, args=(sid,),
                         format_func=lambda x, t=rec["tip"]: kataloq.variant_ad(t, x))
            img(os.path.join(ROOT, "assets", "variant", f"{rec['tip']}_{rec['variant']}.png"), width=300)
            d = rec["data"]
            for f in alt_r["sahələr"]:
                ad, key, nv = f["ad"], f"{sid}:{f['ad']}", f["növ"]
                if nv == "str" and f.get("çoxsətir"):
                    d[ad] = st.text_area(f"{lbl(f)}  (hər sətir ≤{f['max']})", d.get(ad, ""), key=key, height=90)
                elif nv == "str":
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
spec = manual.spec_qur(S, tema, movzu, muellif, qarisiq, seed=movzu, yazi=YAZI[yazi_ad])
h = hashlib.md5(json.dumps(spec, sort_keys=True, ensure_ascii=False, default=str).encode()).hexdigest()
n_nun = sum(manual.numunedir(r) for r in S)

if n_nun:
    st.warning(f"{n_nun} slaydda nümunə mətn dəyişdirilməyib. Faylı yaratmazdan əvvəl onları yenilə.")

ok_onizle, sebeb = onizleme.mumkundur()

if st.query_params.get("diaq"):          # yalnız linkin sonuna ?diaq=1 yazanda görünür (texniki yoxlama)
    import platform, shutil
    with st.expander("🔧 Diaqnostika (önizləmə)", expanded=True):
        st.write({"önizləmə mümkündür": ok_onizle, "səbəb": sebeb, "soffice_yolu()": onizleme.soffice_yolu(),
                  "which soffice": shutil.which("soffice"), "which libreoffice": shutil.which("libreoffice"),
                  "python": platform.python_version(), "sistem": platform.platform()})
        try:
            import pymupdf
            st.write("pymupdf:", pymupdf.__version__)
        except Exception as _e:
            st.write("pymupdf xətası:", repr(_e))
        st.write({y: os.path.exists(y) for y in ("/usr/bin/soffice", "/usr/bin/libreoffice", "/usr/lib/libreoffice/program/soffice")})
if st.button("📥 Hazırla və önizlə", type="primary", disabled=not S):
    with st.spinner("Təqdimat hazırlanır..."):
        b, rep, warn = manual.pptx_bytes(spec)
        png, png_xeta = None, ""
        if ok_onizle:
            try:
                png = onizleme.pptx_png(b, dpi=100)
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
                              "onizleme": bool(cx["png"]), "yazi": yazi_ad},))
    if cx["rep"]:
        with st.expander(f"{len(cx['rep'])} avtomatik düzəliş edildi (uzun mətn və s.)"):
            for r in cx["rep"]:
                st.write("•", r)
    for w in cx["warn"]:
        st.caption("! " + w)
    if cx["png"]:
        st.subheader("👁 Önizləmə")
        st.caption("Önizləmə LibreOffice ilə çəkilir. PowerPoint-də şriftlər və aralıqlar bir qədər fərqli ola bilər. "
                   "Böyütmək üçün şəklin altındakı «🔍 Böyüt» düyməsinə bas.")
        _ac = None
        for r0 in range(0, len(cx["png"]), 3):
            cols = st.columns(3)
            for c, j in zip(cols, range(r0, min(r0 + 3, len(cx["png"])))):
                c.image(cx["png"][j], caption=cx["adlar"][j] if j < len(cx["adlar"]) else None)
                if c.button("🔍 Böyüt", key=f"boyut{j}", use_container_width=True):
                    _ac = j
        if _ac is not None:
            st.session_state.dlg_i = _ac
            boyut()
    elif ok_onizle and cx["png_xeta"]:
        st.warning(f"Önizləmə alınmadı ({cx['png_xeta']}). Fayl isə hazırdır, yükləyə bilərsən.")
    elif not ok_onizle:
        st.info("Bu versiyada slayd önizləməsi yoxdur. Fayl hazırdır: yüklə və PowerPoint-də aç, orada redaktə edə bilərsən.")
    if FEEDBACK_URL:
        st.link_button("2 dəqiqəlik rəy ver", FEEDBACK_URL)

# ------------------------------------------------------------------ layihəni saxla / yüklə
st.divider()
with st.expander("💾 Layihəni saxla / yüklə (sonra davam etmək üçün)"):
    ad_layihe = re.sub(r"[^\w\-]+", "_", movzu.strip(), flags=re.UNICODE).strip("_")[:40] or "layihe"
    st.download_button("💾 Layihəni saxla (.json)",
                       manual.layihe_json(S, {"movzu": movzu, "muellif": muellif, "tema": tema, "qarisiq": qarisiq, "yazi": YAZI[yazi_ad]}),
                       file_name=f"{ad_layihe}.json", mime="application/json",
                       help="Brauzeri bağlayanda iş itir, bu fayl isə qalır.")
    st.file_uploader("Saxlanmış layihəni yüklə", type=["json"], key="yukle", on_change=cb_yukle)
    _m = st.session_state.pop("yukle_mesaj", None)
    if _m:
        (st.error if _m[0] == "error" else st.success)(_m[1])
