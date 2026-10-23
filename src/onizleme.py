"""
ONIZLEME.PY — .pptx baytlarından slayd şəkilləri (PNG) çıxarır: LibreOffice (pptx->pdf) + PyMuPDF (pdf->png).
Önizləmə İSTƏYƏ BAĞLIDIR: LibreOffice və ya pymupdf yoxdursa mumkundur() izah qaytarır, tətbiq önizləmə olmadan işləyir.
Diqqət: önizləmə LibreOffice ilə çəkilir, PowerPoint ilə piksel-piksel eyni olmaya bilər (şrift əvəzləməsi).
"""
import os, shutil, subprocess, tempfile, pathlib


class OnizlemeXetasi(Exception):
    pass


_YOLLAR = [r"C:\Program Files\LibreOffice\program\soffice.exe",
           r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
           "/Applications/LibreOffice.app/Contents/MacOS/soffice"]


def soffice_yolu():
    for ad in ("soffice", "libreoffice", "soffice.exe"):
        p = shutil.which(ad)
        if p:
            return p
    return next((p for p in _YOLLAR if os.path.exists(p)), None)


def _pymupdf():
    try:
        import pymupdf            # yeni ad
        return pymupdf
    except ImportError:
        import fitz               # köhnə ad (köhnə versiyalar)
        return fitz


def mumkundur():
    """(True, "") və ya (False, səbəb)."""
    if not soffice_yolu():
        return False, "LibreOffice tapılmadı"
    try:
        _pymupdf()
    except ImportError:
        return False, "pymupdf quraşdırılmayıb"
    return True, ""


def pptx_png(pptx_bytes, dpi=70, timeout=180):
    """pptx baytları -> [png baytları] (hər slayd üçün bir). Xəta olsa OnizlemeXetasi atır."""
    ok, sebeb = mumkundur()
    if not ok:
        raise OnizlemeXetasi(sebeb)
    fitz = _pymupdf()
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
        src = os.path.join(tmp, "deck.pptx")
        with open(src, "wb") as f:
            f.write(pptx_bytes)
        profil = pathlib.Path(tmp, "lo_profil").as_uri()      # ayrı profil: paralel işləmə və kilid problemi olmasın
        kw = {"creationflags": 0x08000000} if os.name == "nt" else {}     # Windows-da qara pəncərə açılmasın
        try:
            subprocess.run([soffice_yolu(), f"-env:UserInstallation={profil}", "--headless", "--convert-to", "pdf",
                            "--outdir", tmp, src], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           timeout=timeout, check=False, **kw)
        except subprocess.TimeoutExpired:
            raise OnizlemeXetasi("Önizləmə çox uzun çəkdi")
        pdf = os.path.join(tmp, "deck.pdf")
        if not os.path.exists(pdf):
            raise OnizlemeXetasi("LibreOffice faylı çevirə bilmədi")
        doc = fitz.open(pdf)
        try:
            return [p.get_pixmap(dpi=dpi).tobytes("png") for p in doc]
        finally:
            doc.close()
