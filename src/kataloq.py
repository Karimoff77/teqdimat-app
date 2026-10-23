"""
KATALOQ.PY — Manual rejimin kateqoriya kataloqu.

Bir kateqoriya = istifadəçinin gördüyü ad (Giriş, Plan, Müqayisə...).
Hər kateqoriyanın içində bir və ya bir neçə "alt" var; alt = mühərrikdəki slayd tipi (JSON "tip").
Hər altın: düzülüş variantları, forma sahələri (etiket + limit) və dolu nümunəsi var.

Bu fayl gələcək formanın (Manual rejim) məlumat mənbəyidir: forma bu kataloqdan çəkilir.
Limitlər dogrula.py / qur.py / prompt_az.md ilə eyni olmalıdır; tests/yoxla.py hər nümunəni
dogrula-dan keçirir və 0 təmir / 0 xəbərdarlıq tələb edir (limit uyğunsuzluğu dərhal görünür).

mərhələ: 1 = v1 (indi), 2 = sonra (kod artıq işləyir, amma ikinci sıradadır)
"""
import duzulus

# ---- sahə növləri ----
# str        : bir sətir mətn, max simvol
# sətirlər   : siyahı, hər biri max simvol, ən çox "say"
# cütlər     : [başlıq, izah] cütləri, (max1, max2), ən çox "say"
# cədvəl     : başlıqlar + sətirlər
S = lambda ad, et, mx, ixtiyari=False: {"ad": ad, "etiket": et, "növ": "str", "max": mx, "ixtiyari": ixtiyari}
L = lambda ad, et, mx, say, ixtiyari=False: {"ad": ad, "etiket": et, "növ": "sətirlər", "max": mx, "say": say, "ixtiyari": ixtiyari}
P = lambda ad, et, mx, say, col=("Başlıq", "İzah"): {"ad": ad, "etiket": et, "növ": "cütlər", "max": mx, "say": say, "sütunlar": col}

KICKER = S("kicker", "Üst etiket (böyük hərflə)", 30)
BASLIK = S("baslik", "Başlıq", 48)


def _v(tip):
    return duzulus.VARIANTS.get(tip, ["standart"])


KATEQORIYALAR = [
    # ------------------------------------------------------------ 1
    {"id": "giris", "ad": "Giriş", "mərhələ": 1, "izah": "Açılış (titul) slaydı: mövzu, müəllif, tarix",
     "altlar": [{
         "tip": "titul", "ad": "Titul slaydı", "variantlar": ["left", "center", "band"], "variant_sahəsi": "stil_titul",
         "sahələr": [S("kicker", "Üst etiket", 40), S("baslik", "Başlıq (2 sətrə qədər, \\n ilə)", 33),
                     S("alt_baslik", "Alt başlıq", 110), L("info", "Məlumat sətirləri (fənn, müəllim, qrup)", 60, 3, True),
                     S("tarix_yer", "Tarix • yer", 40, True)],
         "nümunə": {"tip": "titul", "kicker": "FƏNN ADI • TƏQDİMAT",
                    "baslik": "Mövzunun adı burada\nikinci sətirdə davam edir",
                    "alt_baslik": "Mövzunun bir cümləlik izahı: nədən bəhs edir və nə üçün vacibdir.",
                    "info": ["Hazırladı: Ad Soyad", "Qrup: 123", "Müəllim: Ad Soyad"], "tarix_yer": "2026 • Bakı"}}]},
    # ------------------------------------------------------------ 2
    {"id": "plan", "ad": "Plan", "mərhələ": 1, "izah": "Təqdimatın məzmunu (2–7 maddə)",
     "altlar": [{
         "tip": "plan", "ad": "Plan slaydı", "variantlar": _v("plan"),
         "sahələr": [P("maddeler", "Plan maddələri (qısa ad + izah)", (24, 70), 7, ("Qısa ad", "İzah"))],
         "nümunə": {"tip": "plan", "maddeler": [
             ["Giriş", "Mövzunun mahiyyəti və əhəmiyyəti"], ["Əsas anlayışlar", "Terminlər və təriflər"],
             ["Metodlar", "Əsas yanaşmalar və qaydalar"], ["Nümunələr", "Praktik tətbiq halları"],
             ["Nəticə", "Əsas fikirlərin yekunu"]]}}]},
    # ------------------------------------------------------------ 3
    {"id": "muqayise", "ad": "Müqayisə", "mərhələ": 1, "izah": "İki anlayışı, metodu və ya yanaşmanı yan-yana qoymaq",
     "altlar": [{
         "tip": "muqayise", "ad": "İki tərəfli müqayisə", "variantlar": _v("muqayise"),
         "sahələr": [KICKER, BASLIK, S("sol_baslik", "Sol tərəfin adı", 34), L("sol", "Sol tərəf maddələri", 80, 4),
                     S("sag_baslik", "Sağ tərəfin adı", 34), L("sag", "Sağ tərəf maddələri", 80, 4),
                     S("izah", "Yekun izah", 110, True)],
         "nümunə": {"tip": "muqayise", "kicker": "MÜQAYİSƏ", "baslik": "Birinci yanaşma və ikinci yanaşma",
                    "sol_baslik": "Birinci yanaşma", "sol": ["Birinci xüsusiyyət haqqında qısa cümlə", "İkinci xüsusiyyət haqqında qısa cümlə", "Üçüncü xüsusiyyət haqqında qısa cümlə"],
                    "sag_baslik": "İkinci yanaşma", "sag": ["Birinci xüsusiyyət haqqında qısa cümlə", "İkinci xüsusiyyət haqqında qısa cümlə", "Üçüncü xüsusiyyət haqqında qısa cümlə"],
                    "izah": "Seçim məqsəddən və şəraitdən asılıdır."}}]},
    # ------------------------------------------------------------ 4
    {"id": "statistika", "ad": "Statistika", "mərhələ": 1, "izah": "Böyük rəqəmlər və ya qrafik (xətt / sütun)",
     "altlar": [
         {"tip": "statistika", "ad": "Böyük rəqəmlər", "variantlar": _v("statistika"),
          "sahələr": [KICKER, BASLIK, {"ad": "rəqəmlər", "etiket": "Rəqəm + izah (3 rəqəmdə rəqəm ≤6 simvol)", "növ": "cütlər", "max": (6, 60), "say": 4, "sütunlar": ("Rəqəm", "İzah")}],
          "nümunə": {"tip": "statistika", "kicker": "RƏQƏMLƏRLƏ", "baslik": "Əsas göstəricilər (nümunə)",
                     "rəqəmlər": [["45%", "Nümunə göstərici: birinci faiz"], ["1:3", "Nümunə nisbət: ikinci göstərici"], ["2026", "Nümunə il: üçüncü göstərici"]]}},
         {"tip": "qrafik", "ad": "Qrafik (redaktə olunan)", "variantlar": ["xett", "sutun"], "variant_sahəsi": "novu",
          "sahələr": [KICKER, BASLIK, L("kateqoriyalar", "X oxu dəyərləri", 20, 10),
                      {"ad": "setirler", "etiket": "Seriyalar (ad + rəqəmlər, ≤3 seriya)", "növ": "seriyalar", "max": 24, "say": 3},
                      S("x_oxu", "X oxu adı", 24, True), S("y_oxu", "Y oxu adı", 24, True),
                      L("neticeler", "Əsas fikirlər (qrafikin yanında)", 80, 3, True)],
          "nümunə": {"tip": "qrafik", "kicker": "QRAFİK", "baslik": "Dəyişmə dinamikası (illüstrativ)", "novu": "xett",
                     "kateqoriyalar": ["2021", "2022", "2023", "2024", "2025"],
                     "setirler": [{"ad": "Göstərici A", "qiymetler": [10, 14, 19, 25, 33]}, {"ad": "Göstərici B", "qiymetler": [8, 9, 12, 13, 15]}],
                     "x_oxu": "İl", "y_oxu": "Dəyər",
                     "neticeler": ["A göstəricisi daha sürətlə artır", "B göstəricisi sabit tempdə dəyişir", "Rəqəmlər illüstrativ nümunədir"]}}]},
    # ------------------------------------------------------------ 5
    {"id": "metn", "ad": "Mətn", "mərhələ": 1, "izah": "Əsas məzmun: siyahı və ya kartlar",
     "altlar": [
         {"tip": "maddeler", "ad": "Maddələr siyahısı", "variantlar": _v("maddeler"),
          "sahələr": [KICKER, BASLIK, L("meqamlar", "Maddələr", 80, 6)],
          "nümunə": {"tip": "maddeler", "kicker": "ƏSAS MƏZMUN", "baslik": "Mövzunun əsas məqamları",
                     "meqamlar": ["Birinci məqam: anlayışın qısa və aydın izahı", "İkinci məqam: onun əsas xüsusiyyəti haqqında cümlə",
                                  "Üçüncü məqam: tətbiq sahəsi və ya nümunə", "Dördüncü məqam: vacib qayda və ya istisna"]}},
         {"tip": "kartlar", "ad": "Kartlar", "variantlar": _v("kartlar"),
          "sahələr": [KICKER, BASLIK, P("kartlar", "Kartlar (başlıq + izah)", (30, 80), 6)],
          "nümunə": {"tip": "kartlar", "kicker": "ANLAYIŞLAR", "baslik": "Əsas anlayışlar",
                     "kartlar": [["Anlayış 1", "Birinci anlayışın qısa izahı"], ["Anlayış 2", "İkinci anlayışın qısa izahı"],
                                 ["Anlayış 3", "Üçüncü anlayışın qısa izahı"], ["Anlayış 4", "Dördüncü anlayışın qısa izahı"],
                                 ["Anlayış 5", "Beşinci anlayışın qısa izahı"], ["Anlayış 6", "Altıncı anlayışın qısa izahı"]]}}]},
    # ------------------------------------------------------------ 6
    {"id": "cedvel", "ad": "Cədvəl", "mərhələ": 1, "izah": "Məlumatı sətir və sütunlarla (2–3 sütun, ≤6 sətir)",
     "altlar": [{
         "tip": "cedvel", "ad": "Cədvəl", "variantlar": _v("cedvel"),
         "sahələr": [KICKER, BASLIK, L("basliqlar", "Sütun başlıqları", 24, 3),
                     {"ad": "setirler", "etiket": "Sətirlər (hər xana ≤48)", "növ": "cədvəl", "max": 48, "say": 6}],
         "nümunə": {"tip": "cedvel", "kicker": "CƏDVƏL", "baslik": "Əsas növlərin müqayisəsi",
                    "basliqlar": ["Növ", "Xüsusiyyət", "Nümunə"],
                    "setirler": [["Birinci növ", "Əsas xüsusiyyəti", "Nümunə 1"], ["İkinci növ", "Əsas xüsusiyyəti", "Nümunə 2"],
                                 ["Üçüncü növ", "Əsas xüsusiyyəti", "Nümunə 3"], ["Dördüncü növ", "Əsas xüsusiyyəti", "Nümunə 4"]]}}]},
    # ------------------------------------------------------------ 7
    {"id": "sekilli", "ad": "Şəkilli mətn", "mərhələ": 1, "izah": "Mətn + şəkil yeri (PowerPoint-də klik edib şəkil seçilir)",
     "altlar": [{
         "tip": "sekilli", "ad": "Mətn + şəkil yeri", "variantlar": _v("sekilli"),
         "sahələr": [KICKER, BASLIK, L("meqamlar", "Mətn maddələri", 75, 4), S("sekil_axtaris", "Şəkil üçün axtarış sözü (ingiliscə)", 60, True)],
         "nümunə": {"tip": "sekilli", "kicker": "ŞƏKİLLƏ İZAH", "baslik": "Obyekt şəkildə necə görünür?",
                    "meqamlar": ["Şəkildə əsas hissə qeyd olunur", "İkinci hissənin vəzifəsi izah edilir", "Ümumi işləmə qaydası göstərilir"],
                    "sekil_axtaris": "example diagram"}}]},
    # ------------------------------------------------------------ 8
    {"id": "sonluq", "ad": "Sonluq", "mərhələ": 1, "izah": "Yekun slaydı: təşəkkür və ya suallar",
     "altlar": [{
         "tip": "tesekkur", "ad": "Təşəkkür / suallar", "variantlar": _v("tesekkur"),
         "sahələr": [S("alt", "Yekun cümlə", 60, True)],
         "nümunə": {"tip": "tesekkur", "alt": "Fənn adı • Mövzunun adı"}}]},
    # ------------------------------------------------------------ əlavələr (mənim təklifim)
    {"id": "proses", "ad": "Addımlar", "mərhələ": 1, "izah": "Ardıcıl proses: necə edilir, necə işləyir (3–4 addım)",
     "altlar": [{
         "tip": "proses", "ad": "Addım-addım proses", "variantlar": _v("proses"),
         "sahələr": [KICKER, BASLIK, P("addimlar", "Addımlar (başlıq + izah)", (36, 90), 4, ("Addım", "İzah")), S("netice", "Qısa yekun xətt", 70, True)],
         "nümunə": {"tip": "proses", "kicker": "ADDIMLAR", "baslik": "Proses necə həyata keçirilir?",
                    "addimlar": [["Hazırlıq", "İlkin məlumatlar toplanır və məqsəd müəyyən edilir"], ["İcra", "Əsas əməliyyatlar ardıcıl yerinə yetirilir"],
                                 ["Yoxlama", "Alınan nəticə qaydalara uyğun yoxlanılır"], ["Yekun", "Nəticə sənədləşdirilir və təqdim edilir"]],
                    "netice": "Hazırla → İcra et → Yoxla → Yekunlaşdır"}}]},
    {"id": "dustur", "ad": "Düstur", "mərhələ": 1, "izah": "Riyazi və ya iqtisadi düstur + hissələrin izahı (düz mətn, redaktor deyil)",
     "altlar": [{
         "tip": "formula", "ad": "Düstur", "variantlar": ["standart"],
         "sahələr": [KICKER, BASLIK, S("formula", "Düstur (düz mətn, məs. ax + by = c)", 45),
                     P("aciqlamalar", "Hissələrin izahı (işarə + izah)", (22, 50), 3, ("İşarə", "İzah"))],
         "nümunə": {"tip": "formula", "kicker": "DÜSTUR", "baslik": "Düz xəttin ümumi tənliyi",
                    "formula": "ax + by + c = 0",
                    "aciqlamalar": [["a, b", "x və y-in əmsalları"], ["c", "sərbəst hədd"], ["x, y", "dəyişənlər"]]}}]},
    # ------------------------------------------------------------ mərhələ 2 (kod hazırdır, ikinci sıra)
    {"id": "xronologiya", "ad": "Xronologiya", "mərhələ": 2, "izah": "Tarixi və ya mərhələli inkişaf (tək düzülüş)",
     "altlar": [{
         "tip": "xett", "ad": "Zaman xətti", "variantlar": ["standart"],
         "sahələr": [KICKER, BASLIK, {"ad": "hadiseler", "etiket": "Hadisələr (tarix, başlıq, qısa izah)", "növ": "üçlüklər", "max": (18, 28, 40), "say": 4, "sütunlar": ("Tarix", "Başlıq", "İzah")},
                     S("esas_fikir", "Əsas fikir", 110, True)],
         "nümunə": {"tip": "xett", "kicker": "TARİXÇƏ", "baslik": "İnkişaf mərhələləri",
                    "hadiseler": [["1-ci mərhələ", "Başlanğıc", "İlkin ideya yaranır"], ["2-ci mərhələ", "Formalaşma", "Əsas qaydalar müəyyənləşir"],
                                  ["3-ci mərhələ", "Yayılma", "Geniş tətbiq başlayır"], ["Bu gün", "Müasir vəziyyət", "Yeni yanaşmalar işlənir"]],
                    "esas_fikir": "Hər mərhələ növbətisinin əsasını qoyur."}}]},
    {"id": "esas_fikir", "ad": "Əsas fikir", "mərhələ": 2, "izah": "Bir böyük iddia və ya nəticə (tək düzülüş)",
     "altlar": [{
         "tip": "iddia", "ad": "Əsas iddia", "variantlar": ["standart"],
         "sahələr": [KICKER, BASLIK, S("simvol", "Böyük simvol (məs. ?, !, ≈)", 60, True), S("iddia", "İddia", 80), S("izah", "İzah", 150, True)],
         "nümunə": {"tip": "iddia", "kicker": "ƏSAS FİKİR", "baslik": "Yadda saxlamalı olduğumuz",
                    "simvol": "!", "iddia": "Düzgün başlanğıc işin yarısıdır",
                    "izah": "Məsələni əvvəl düzgün qurmaq sonrakı bütün addımları asanlaşdırır və səhv ehtimalını azaldır."}}]},
]


def hamisi(merhele=None):
    """(kateqoriya, alt, variant) üçlüklərini qaytarır. merhele=1 yalnız v1."""
    for k in KATEQORIYALAR:
        if merhele and k["mərhələ"] > merhele:
            continue
        for a in k["altlar"]:
            for v in a["variantlar"]:
                yield k, a, v


def numune_slayd(alt, variant):
    """Alt-ın nümunəsini verilmiş variantla JSON slayd kimi qaytarır."""
    import copy
    s = copy.deepcopy(alt["nümunə"])
    sahe = alt.get("variant_sahəsi", "duzulus")
    if variant != "standart":
        s[sahe] = variant
    return s


# ---- variantların istifadəçiyə göstərilən adları ----
VARIANT_AD = {"left": "Sol", "center": "Mərkəz", "band": "Zolaqlı", "list": "Siyahı", "cards": "Kartlar", "line": "Xətt",
              "versus": "Qarşı-qarşıya", "plain": "Sadə", "xett": "Xətt qrafiki", "sutun": "Sütun qrafiki",
              "numbered": "Nömrəli", "split": "Bölünmüş", "cards2": "2×2 kart", "grid": "Şəbəkə", "rows": "Sətirlər",
              "icons": "Dairəli nömrə", "table": "Cədvəl", "horizontal": "Üfüqi", "vertical": "Şaquli", "chevron": "Oxlar",
              "question": "Suallar", "standart": "Standart"}
VARIANT_AD_TIP = {("sekilli", "right"): "Şəkil sağda", ("sekilli", "left"): "Şəkil solda", ("sekilli", "top"): "Şəkil yuxarıda",
                  ("titul", "left"): "Sol", ("muqayise", "cards"): "Kartlar", ("tesekkur", "split"): "Bölünmüş"}


def variant_ad(tip, v):
    return VARIANT_AD_TIP.get((tip, v)) or VARIANT_AD.get(v, v)
