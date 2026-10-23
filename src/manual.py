"""
MANUAL.PY — Manual rejimin məntiqi (UI-dan asılı olmayan hissə). app.py bunu çağırır.
Slayd qeydi:  {"id": int, "kat": "metn", "tip": "maddeler", "variant": "list", "data": {...sahələr...}}
`data` JSON slaydın sahələridir ("tip" və variant sahəsi olmadan).
"""
import copy, io, itertools
import kataloq, qur
from dogrula import dogrula

_ids = itertools.count(1)


def kat_tap(kid):
    return next(k for k in kataloq.KATEQORIYALAR if k["id"] == kid)


def alt_tap(kid, tip):
    return next(a for a in kat_tap(kid)["altlar"] if a["tip"] == tip)


def yeni(kid, tip, variant=None):
    """Nümunə məzmunla yeni slayd qeydi."""
    alt = alt_tap(kid, tip)
    variant = variant or alt["variantlar"][0]
    data = copy.deepcopy(alt["nümunə"])
    data.pop("tip", None)
    data.pop(alt.get("variant_sahəsi", "duzulus"), None)
    return {"id": next(_ids), "kat": kid, "tip": tip, "variant": variant, "data": data}


def slayd_json(rec):
    alt = alt_tap(rec["kat"], rec["tip"])
    s = copy.deepcopy(rec["data"])
    s["tip"] = rec["tip"]
    if rec["variant"] != "standart":
        s[alt.get("variant_sahəsi", "duzulus")] = rec["variant"]
    return s


def numunedir(rec):
    """İstifadəçi nümunə mətni dəyişməyibsə True (təqdimata nümunə mətn düşməsin)."""
    alt = alt_tap(rec["kat"], rec["tip"])
    n = copy.deepcopy(alt["nümunə"])
    n.pop("tip", None); n.pop(alt.get("variant_sahəsi", "duzulus"), None)
    return _norm(n) == _norm(rec["data"])


def _norm(x):
    """Cədvəl widget-indən keçən rəqəmlər mətnə çevrilir (10 -> "10"); müqayisədə fərq sayılmasın."""
    if isinstance(x, dict):
        return {k: _norm(v) for k, v in x.items()}
    if isinstance(x, (list, tuple)):
        return [_norm(v) for v in x]
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    return str(x).strip() if isinstance(x, (int, float, str)) else x


def xeberdarliqlar(rec):
    """Bir slayd üçün: boş məcburi sahələr + dogrula.py-nin təmir qeydləri (eyni funksiya qurma zamanı da işləyir)."""
    alt = alt_tap(rec["kat"], rec["tip"])
    out = []
    for f in alt["sahələr"]:
        v = rec["data"].get(f["ad"])
        if f["növ"] == "str" and not f.get("ixtiyari") and not (v or "").strip():
            out.append(f"«{f['etiket']}» boşdur")
    sp, log = dogrula({"slaydlar": [slayd_json(rec)]}, ai=False)
    if not sp["slaydlar"]:
        out.append("Məzmun çox azdır: bu slayd təqdimata düşməyəcək")
    for l in log:
        out.append(l.split(": ", 1)[-1] if ": " in l else l)
    return out


def spec_qur(recs, tema="corporate", movzu="Təqdimat", muellif="", qarisiq=False, seed=""):
    return {"tema": tema, "movzu": movzu or "Təqdimat", "muellif": muellif, "seed": seed or movzu or "t",
            "stil": "auto" if qarisiq else "tema", "slaydlar": [slayd_json(r) for r in recs]}


def pptx_bytes(spec):
    """spec -> (bytes, təmir_qeydləri, qurma_xəbərdarlıqları). Çökməməlidir."""
    spec, rep = dogrula(spec, ai=False)
    qur.WARN.clear()
    d = qur.build(spec)
    bio = io.BytesIO()
    d.prs.save(bio)
    return bio.getvalue(), rep, list(qur.WARN)


# ------------------------------------------------------------------ layihəni saxla / yüklə
import json

MAX_SLAYD_MANUAL = 60


def _s(v, ehtiyat=""):
    return str(v) if isinstance(v, (str, int, float)) and not isinstance(v, bool) else ehtiyat


def sanitize(kid, tip, data):
    """Yüklənmiş fayldan gələn data-nı sahə növünə görə təhlükəsiz formaya salır (səhv tipli dəyər forma sahəsini sındırmasın)."""
    alt = alt_tap(kid, tip)
    nun = copy.deepcopy(alt["nümunə"])
    data = data if isinstance(data, dict) else {}
    out = {}
    for f in alt["sahələr"]:
        ad, nv, v = f["ad"], f["növ"], data.get(f["ad"])
        bos = "" if nv == "str" else []
        ehtiyat = bos if f.get("ixtiyari") else nun.get(ad, bos)
        if nv == "str":
            out[ad] = _s(v, ehtiyat) if ad in data else ehtiyat
        elif nv == "sətirlər":
            out[ad] = [_s(x) for x in v] if isinstance(v, list) else ehtiyat
        elif nv in ("cütlər", "üçlüklər", "cədvəl"):
            out[ad] = [[_s(c) for c in r] if isinstance(r, (list, tuple)) else [_s(r)] for r in v] if isinstance(v, list) else ehtiyat
        elif nv == "seriyalar":
            out[ad] = [{"ad": _s(z.get("ad")), "qiymetler": [_s(q) for q in z.get("qiymetler", [])] if isinstance(z.get("qiymetler"), list) else []}
                       for z in v if isinstance(z, dict)] if isinstance(v, list) else ehtiyat
    return out


def layihe_json(recs, ayarlar):
    return json.dumps({"versiya": 1, "ayarlar": ayarlar,
                       "slaydlar": [{"kat": r["kat"], "tip": r["tip"], "variant": r["variant"], "data": r["data"]} for r in recs]},
                      ensure_ascii=False, indent=1)


def layihe_yukle(metn):
    """JSON mətni -> (ayarlar, [slayd qeydləri], atılan_say). Etibarsız fayl üçün ValueError atır."""
    o = json.loads(metn)
    if not isinstance(o, dict) or o.get("versiya") != 1 or not isinstance(o.get("slaydlar"), list):
        raise ValueError("bu fayl bu proqramın layihə faylı deyil")
    recs, atilan = [], 0
    for s in o["slaydlar"][:MAX_SLAYD_MANUAL]:
        try:
            alt = alt_tap(s["kat"], s["tip"])
        except (StopIteration, KeyError, TypeError):
            atilan += 1
            continue
        v = s.get("variant") if s.get("variant") in alt["variantlar"] else alt["variantlar"][0]
        recs.append({"id": next(_ids), "kat": s["kat"], "tip": s["tip"], "variant": v, "data": sanitize(s["kat"], s["tip"], s.get("data"))})
    atilan += max(0, len(o["slaydlar"]) - MAX_SLAYD_MANUAL)
    if not recs:
        raise ValueError("faylda oxuna bilən slayd yoxdur")
    ay = o.get("ayarlar") if isinstance(o.get("ayarlar"), dict) else {}
    return ay, recs, atilan
