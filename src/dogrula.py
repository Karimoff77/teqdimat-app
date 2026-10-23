"""
DOGRULA.PY — AI-dan gələn JSON-u yoxlayır və təmir edir.
Məqsəd: zəif və ya səhv format verən model də olsa, renderer heç vaxt çökməsin,
tələbə həmişə işlək fayl alsın. Hər düzəliş `log`-a yazılır (keyfiyyət ölçmək üçün).
"""
import re

MAX_SLAYD = 20
ALIAS = {"reqemler": "rəqəmlər", "rakamlar": "rəqəmlər", "rəqəmlər": "rəqəmlər"}


def _s(x):
    """Hər şeyi sətrə çevir."""
    if x is None: return ""
    if isinstance(x, float) and x.is_integer(): x = int(x)
    return str(x).strip()


def cut(txt, mx, log, yer):
    txt = _s(txt)
    if len(txt) <= mx:
        return txt
    c = txt[:mx - 1]
    sp = c.rfind(" ")
    if sp > mx * 0.6: c = c[:sp]
    log.append(f"{yer}: {len(txt)}→{mx} simvola kəsildi")
    return c.rstrip(" ,;:-") + "…"


def pair(x, log, yer):
    """Hər forma-nı [başlıq, izah] cütünə çevir."""
    if isinstance(x, dict):
        v = list(x.values())
        return [_s(v[0]) if v else "", _s(v[1]) if len(v) > 1 else ""]
    if isinstance(x, (list, tuple)):
        v = [_s(i) for i in x]
        return [v[0] if v else "", v[1] if len(v) > 1 else ""]
    t = _s(x)
    for sep in (":", " — ", " - "):
        if sep in t:
            a, b = t.split(sep, 1); log.append(f"{yer}: sətir cütə bölündü"); return [a.strip(), b.strip()]
    log.append(f"{yer}: izah yoxdur"); return [t, ""]


def strs(lst, mx, n, log, yer):
    if not isinstance(lst, (list, tuple)): lst = [lst] if lst else []
    out = []
    for i in lst:
        if isinstance(i, dict): i = " ".join(_s(v) for v in i.values())
        t = cut(i, mx, log, yer)
        if t: out.append(t)
    if len(out) > n:
        log.append(f"{yer}: {len(out)}→{n} elementə azaldıldı"); out = out[:n]
    return out


def pairs(lst, m1, m2, n, log, yer):
    if not isinstance(lst, (list, tuple)): lst = []
    out = []
    for i, x in enumerate(lst):
        a, b = pair(x, log, yer)
        if a: out.append([cut(a, m1, log, yer), cut(b, m2, log, yer)])
    if len(out) > n:
        log.append(f"{yer}: {len(out)}→{n} elementə azaldıldı"); out = out[:n]
    return out


def head(s, log, yer, kick="MÖVZU"):
    s["kicker"] = cut(s.get("kicker") or kick, 30, log, yer + ".kicker").upper()
    s["baslik"] = cut(s.get("baslik") or s.get("kicker", "") or "Başlıq", 48, log, yer + ".baslik")


def fix_slide(s, i, log):
    yer = f"s{i}.{s.get('tip')}"
    t = s.get("tip")
    if t == "titul":
        s["kicker"] = cut(s.get("kicker") or "TƏQDİMAT", 40, log, yer)
        lines = [cut(l, 33, log, yer + ".baslik") for l in _s(s.get("baslik") or "Təqdimat").split("\n")][:2]
        s["baslik"] = "\n".join(lines)
        s["alt_baslik"] = cut(s.get("alt_baslik"), 110, log, yer)
        s["info"] = strs(s.get("info", []), 60, 3, log, yer)
        s["tarix_yer"] = cut(s.get("tarix_yer"), 40, log, yer)
    elif t == "plan":
        s["maddeler"] = pairs(s.get("maddeler"), 24, 70, 7, log, yer)
        if len(s["maddeler"]) < 2: return None
    elif t in ("maddeler", "sekilli"):
        head(s, log, yer)
        s["meqamlar"] = strs(s.get("meqamlar"), 80 if t == "maddeler" else 75, 6 if t == "maddeler" else 4, log, yer)
        if not s["meqamlar"]: return None
        s["sekil_axtaris"] = cut(s.get("sekil_axtaris"), 60, log, yer)
    elif t == "kartlar":
        head(s, log, yer)
        s["kartlar"] = pairs(s.get("kartlar"), 30, 80, 6, log, yer)
        if len(s["kartlar"]) < 2: return None
    elif t == "muqayise":
        head(s, log, yer)
        s["sol_baslik"] = cut(s.get("sol_baslik") or "Birinci", 34, log, yer)
        s["sag_baslik"] = cut(s.get("sag_baslik") or "İkinci", 34, log, yer)
        s["sol"] = strs(s.get("sol"), 80, 4, log, yer); s["sag"] = strs(s.get("sag"), 80, 4, log, yer)
        if not s["sol"] or not s["sag"]:
            log.append(yer + ": bir tərəf boşdur → maddələr slaydına çevrildi")
            s2 = {"tip": "maddeler", "kicker": s["kicker"], "baslik": s["baslik"], "meqamlar": (s["sol"] + s["sag"])[:6]}
            return fix_slide(s2, i, log) if s2["meqamlar"] else None
        s["izah"] = cut(s.get("izah"), 110, log, yer)
    elif t == "proses":
        head(s, log, yer)
        s["addimlar"] = pairs(s.get("addimlar"), 36, 90, 4, log, yer)
        if len(s["addimlar"]) < 2: return None
        s["netice"] = cut(s.get("netice"), 70, log, yer)
    elif t == "xett":
        head(s, log, yer)
        ev = []
        for x in (s.get("hadiseler") or [])[:4]:
            v = [_s(a) for a in (x if isinstance(x, (list, tuple)) else [x])] + ["", "", ""]
            ev.append([cut(v[0], 18, log, yer), cut(v[1], 28, log, yer), cut(v[2], 40, log, yer)])
        if len(ev) < 2: return None
        s["hadiseler"] = ev; s["esas_fikir"] = cut(s.get("esas_fikir"), 110, log, yer)
    elif t == "statistika":
        head(s, log, yer)
        raw = None
        for k in list(s.keys()):
            if k in ALIAS: raw = s.pop(k); break
        raw = raw or s.get("rəqəmlər") or []
        pr = [pair(x, log, yer) for x in raw][:4]
        mx = {1: 12, 2: 9, 3: 6, 4: 5}.get(len(pr), 6)
        keep = []
        for a, b in pr:
            if not a: continue
            if len(a) > mx:
                log.append(f"{yer}: rəqəm '{a}' çox uzundur → element atıldı (kəsmək məlumatı pozardı)"); continue
            keep.append([a, cut(b, 60, log, yer)])
        pr = keep
        if len(pr) < 2: return None
        s["rəqəmlər"] = pr
    elif t == "cedvel":
        head(s, log, yer)
        hd = [cut(h, 24, log, yer) for h in (s.get("basliqlar") or [])][:3]
        if len(hd) < 2: return None
        rows = []
        for r in (s.get("setirler") or [])[:6]:
            r = [cut(c, 48, log, yer) for c in (r if isinstance(r, (list, tuple)) else [r])]
            if len(r) != len(hd): log.append(f"{yer}: sətir uzunluğu {len(r)}→{len(hd)}")
            rows.append((r + [""] * len(hd))[:len(hd)])
        if not rows: return None
        s["basliqlar"], s["setirler"] = hd, rows
    elif t == "iddia":
        head(s, log, yer)
        s["simvol"] = cut(s.get("simvol"), 60, log, yer)
        s["iddia"] = cut(s.get("iddia") or s.get("baslik"), 80, log, yer)
        s["izah"] = cut(s.get("izah"), 150, log, yer)
    elif t == "formula":
        head(s, log, yer)
        s["formula"] = cut(s.get("formula"), 45, log, yer)
        s["aciqlamalar"] = pairs(s.get("aciqlamalar"), 22, 50, 3, log, yer)
        if not s["formula"] or not s["aciqlamalar"]: return None
    elif t == "qrafik":
        head(s, log, yer)
        cats = [_s(c) for c in (s.get("kateqoriyalar") or [])][:10]
        ser = []
        for z in (s.get("setirler") or [])[:3]:
            vals = []
            for v in (z.get("qiymetler") or []):
                try: vals.append(float(str(v).replace(",", ".").replace("%", "")))
                except ValueError:
                    log.append(f"{yer}: '{v}' rəqəm deyil → slayd atıldı (uydurma qiymət yazılmır)"); return None
            if len(vals) != len(cats):
                log.append(f"{yer}: qiymət sayı ({len(vals)}) kateqoriya sayına ({len(cats)}) uyğun deyil → slayd atıldı")
                return None
            ser.append({"ad": cut(z.get("ad") or "Seriya", 24, log, yer), "qiymetler": vals})
        if len(cats) < 2 or not ser: return None
        s["kateqoriyalar"], s["setirler"] = cats, ser
        s["novu"] = s.get("novu") if s.get("novu") in ("xett", "sutun") else "xett"
        s["neticeler"] = strs(s.get("neticeler"), 80, 3, log, yer)
        s["x_oxu"] = cut(s.get("x_oxu"), 24, log, yer); s["y_oxu"] = cut(s.get("y_oxu"), 24, log, yer)
    elif t == "tesekkur":
        s["alt"] = cut(s.get("alt"), 60, log, yer)
    else:
        log.append(f"{yer}: naməlum tip → maddələr slaydına çevrildi")
        txt = [_s(v) for k, v in s.items() if k not in ("tip", "kicker", "baslik") and isinstance(v, (str, list))]
        flat = []
        for v in txt:
            flat += [x for x in (v if isinstance(v, list) else [v]) if isinstance(x, str)]
        s2 = {"tip": "maddeler", "kicker": s.get("kicker", "MÖVZU"), "baslik": s.get("baslik", "Məlumat"), "meqamlar": flat}
        return fix_slide(s2, i, log) if flat else None
    return s


def dogrula(spec, ai=True):
    """spec -> (təmiz_spec, log). Heç vaxt xəta atmır.
    ai=True: AI çıxışı üçün (slayd sayı limiti + titul/təşəkkür məcburi əlavə olunur).
    ai=False: Manual rejim (slayd sayı limitsiz, quruluşu istifadəçi qurur; yalnız hər slayd yoxlanılır)."""
    log = []
    spec = dict(spec)
    spec["movzu"] = cut(spec.get("movzu") or "Təqdimat", 70, log, "movzu")
    spec["muellif"] = cut(spec.get("muellif"), 50, log, "muellif")
    out = []
    for i, s in enumerate(spec.get("slaydlar") or [], 1):
        if not isinstance(s, dict) or not s.get("tip"):
            log.append(f"s{i}: tip yoxdur → atıldı"); continue
        r = fix_slide(dict(s), i, log)
        if r is None: log.append(f"s{i}.{s.get('tip')}: kifayət qədər məzmun yoxdur → atıldı"); continue
        out.append(r)
    if ai and len(out) > MAX_SLAYD:
        log.append(f"slayd sayı {len(out)}→{MAX_SLAYD}"); out = out[:MAX_SLAYD - 1] + [out[-1]]
    if ai and (not out or out[0]["tip"] != "titul"):
        log.append("titul slaydı yox idi → əlavə edildi")
        out.insert(0, {"tip": "titul", "kicker": "TƏQDİMAT", "baslik": spec["movzu"][:33], "alt_baslik": "", "info": [], "tarix_yer": ""})
    if ai and out[-1]["tip"] != "tesekkur":
        log.append("təşəkkür slaydı yox idi → əlavə edildi")
        out.append({"tip": "tesekkur", "alt": ""})
    spec["slaydlar"] = out
    return spec, log
