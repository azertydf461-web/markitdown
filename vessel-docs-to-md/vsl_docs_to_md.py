#!/usr/bin/env python3
"""Пакетная конвертация судовой документации в Markdown.

Обходит папки-источники рекурсивно и для каждого файла создаёт <имя файла>.md
в зеркальной структуре папок в выходном каталоге. Исходные файлы не изменяются.
Всё выполняется локально: без LLM и облачных сервисов, в интернет ничего не отправляется.

    python vsl_docs_to_md.py           конвертация с распознаванием сканов (повторный запуск — только новые/изменённые)
    python vsl_docs_to_md.py --scan    инвентаризация без конвертации

Чертежи и схемы (DWG/DXF, крупноформатные PDF/TIFF, файлы в папках «Чертежи», «Схемы», «Drawings»)
оформляются карточкой: ссылка на оригинал, превью листа и все надписи с чертежа (штамп, позиции, примечания).

Подробности в README.md.
"""

import argparse
import csv
import datetime as dt
import email
import email.policy
import glob
import io
import multiprocessing as mp
import os
import queue
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time

# ---------------------------------------------------------------------------
# Настройки по умолчанию (можно переопределить параметрами --source и --out)
# ---------------------------------------------------------------------------
DEFAULT_SOURCES = [
    ("Avacha", r"E:\000 Avacha\1. VSLdocuments"),
    ("Chaivo", r"E:\000 Chaivo"),
]
DEFAULT_OUTPUT = r"E:\000 MD"

IS_WIN = os.name == "nt"

# Форматы, которые MarkItDown читает сам
NATIVE_EXT = {
    ".pdf",
    ".docx",
    ".pptx",
    ".xlsx",
    ".xls",
    ".csv",
    ".msg",
    ".html",
    ".htm",
    ".txt",
    ".text",
    ".md",
    ".markdown",
    ".json",
    ".jsonl",
    ".epub",
    ".zip",
    ".ipynb",
}
# Родственные OOXML-форматы, которые читаются тем же конвертером
EXT_ALIAS = {
    ".docm": ".docx",
    ".dotx": ".docx",
    ".dotm": ".docx",
    ".xlsm": ".xlsx",
    ".xltx": ".xlsx",
    ".xltm": ".xlsx",
    ".pptm": ".pptx",
    ".ppsx": ".pptx",
    ".ppsm": ".pptx",
    ".potx": ".pptx",
}
# Старые/чужие офисные форматы: сначала пересохраняются в OOXML через MS Office или LibreOffice
LEGACY_EXT = {
    ".doc": ".docx",
    ".dot": ".docx",
    ".rtf": ".docx",
    ".odt": ".docx",
    ".wps": ".docx",
    ".ppt": ".pptx",
    ".pps": ".pptx",
    ".pot": ".pptx",
    ".odp": ".pptx",
    ".xlsb": ".xlsx",
    ".ods": ".xlsx",
}
IMAGE_EXT = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif"}
CAD_EXT = {".dwg", ".dxf"}
CODE_EXT = {".xml": "xml", ".log": "", ".ini": "ini", ".cfg": "", ".conf": ""}
SKIP_NAMES = {"thumbs.db", "desktop.ini", ".ds_store"}

# Признаки чертежа: слово в пути или лист формата A2 и крупнее (длинная сторона ≥ ~565 мм)
DRAWING_WORDS = ("чертеж", "чертёж", "схем", "drawing", "dwg", "p&id", "diagram")
LARGE_SHEET_PT = 1600
PREVIEW_PX = 1600
OCR_MAX_PIXELS = 40e6  # ограничение для OCR листов A1/A0, иначе медленно и много памяти

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TESSDATA_URLS = (
    "https://github.com/tesseract-ocr/tessdata_best/raw/main/{}.traineddata",
    "https://github.com/tesseract-ocr/tessdata_fast/raw/main/{}.traineddata",
)

NO_PASSWORD = "vsl2md-no-password"  # подставляется, чтобы защищённые паролем файлы не вешали Office диалогом

STATUS_HELP = {
    "OK": "сконвертировано",
    "OCR": "сконвертировано с распознаванием сканов",
    "DRAWING": "чертёж/схема: карточка с превью и надписями",
    "DRAWING_NO_TEXT": "DWG без ODA File Converter: карточка без надписей",
    "LOW_TEXT": "сканы не распознаны: OCR был недоступен (см. начало вывода)",
    "NO_TEXT": "изображение без распознаваемого текста",
    "UNCHANGED": "уже сконвертировано ранее, исходник не менялся",
    "NEEDS_OCR": "изображение не обработано: OCR был недоступен",
    "NO_CONVERTER": "старый формат Office: нужен MS Office (pywin32) или LibreOffice",
    "UNSUPPORTED": "формат не поддерживается (архивы rar/7z, видео, аудио и т.п.)",
    "TOO_LARGE": "файл больше --max-size-mb",
    "ERROR": "ошибка конвертации (см. колонку message)",
    "TIMEOUT": "превышен лимит времени --timeout",
    "CRASH": "аварийное завершение конвертера на этом файле",
    "ACCESS": "нет доступа к папке",
}


# ---------------------------------------------------------------------------
# Пути: префикс \\?\ снимает ограничение Windows MAX_PATH = 260 символов
# ---------------------------------------------------------------------------
def long_path(p):
    if not IS_WIN:
        return os.path.abspath(p)
    p = os.path.abspath(p)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def plain_path(p):
    if p.startswith("\\\\?\\UNC\\"):
        return "\\\\" + p[8:]
    if p.startswith("\\\\?\\"):
        return p[4:]
    return p


def page_ranges(pages):
    """[0,1,2,5] -> '1-3, 6' (нумерация с 1)."""
    out, start, prev = [], None, None
    for n in sorted(pages):
        if start is None:
            start = prev = n
        elif n == prev + 1:
            prev = n
        else:
            out.append((start, prev))
            start = prev = n
    if start is not None:
        out.append((start, prev))
    return ", ".join(str(a + 1) if a == b else f"{a + 1}-{b + 1}" for a, b in out)


def yaml_str(s):
    return "'" + str(s).replace("'", "''") + "'"


def fmt_time(ts):
    return dt.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")


def read_front_matter(path):
    """Поля front matter ранее созданного .md (только простые key: 'value')."""
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            head = f.read(4096)
    except OSError:
        return {}
    if not head.startswith("---"):
        return {}
    fields = {}
    for line in head.splitlines()[1:]:
        if line.strip() == "---":
            break
        m = re.match(r"^(\w+): (.*)$", line)
        if m:
            v = m.group(2).strip()
            if len(v) >= 2 and v[0] == v[-1] == "'":
                v = v[1:-1].replace("''", "'")
            fields[m.group(1)] = v
    return fields


# ---------------------------------------------------------------------------
# Поиск внешних средств
# ---------------------------------------------------------------------------
def find_soffice():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for c in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if os.path.isfile(c):
            return c
    return None


def has_pywin32():
    if not IS_WIN:
        return False
    try:
        import win32com.client  # noqa: F401

        return True
    except ImportError:
        return False


def find_oda():
    p = shutil.which("ODAFileConverter")
    if p:
        return p
    for base in (r"C:\Program Files\ODA", r"C:\Program Files (x86)\ODA"):
        found = sorted(
            glob.glob(os.path.join(base, "ODAFileConverter*", "ODAFileConverter.exe"))
        )
        if found:
            return found[-1]
    return None


def ensure_tessdata(user_value, langs):
    """Папка с языковыми файлами OCR; если их нет — скачиваются рядом со скриптом."""
    d = find_tessdata(user_value, langs)
    if d:
        return d, ""
    import urllib.request

    target = os.path.join(SCRIPT_DIR, "tessdata")
    os.makedirs(target, exist_ok=True)
    for lang in langs:
        dst = os.path.join(target, f"{lang}.traineddata")
        if os.path.isfile(dst):
            continue
        err = ""
        for url in TESSDATA_URLS:
            print(f"Скачиваю языковой файл OCR {lang}.traineddata ...", flush=True)
            try:
                with urllib.request.urlopen(url.format(lang), timeout=180) as r, open(
                    dst + ".part", "wb"
                ) as f:
                    shutil.copyfileobj(r, f)
                os.replace(dst + ".part", dst)
                break
            except Exception as e:
                err = f"{type(e).__name__}: {e}"
                try:
                    os.remove(dst + ".part")
                except OSError:
                    pass
        else:
            return None, f"не удалось скачать {lang}.traineddata ({err})"
    return target, ""


def find_tessdata(user_value, langs):
    candidates = [user_value, os.environ.get("TESSDATA_PREFIX")]
    candidates += [
        os.path.join(SCRIPT_DIR, "tessdata"),
        r"C:\Program Files\Tesseract-OCR\tessdata",
        r"C:\Program Files (x86)\Tesseract-OCR\tessdata",
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tessdata"),
        "/usr/share/tesseract-ocr/5/tessdata",
        "/usr/share/tesseract-ocr/4.00/tessdata",
        "/usr/share/tessdata",
        "/opt/homebrew/share/tessdata",
        "/usr/local/share/tessdata",
    ]
    for c in candidates:
        if c and all(
            os.path.isfile(os.path.join(c, f"{l}.traineddata")) for l in langs
        ):
            return c
    return None


# ---------------------------------------------------------------------------
# Пересохранение старых форматов в OOXML (выполняется внутри рабочего процесса)
# ---------------------------------------------------------------------------
class LegacyConverter:
    def __init__(self, mode, tmpdir, soffice):
        self.mode = mode  # auto | office | libreoffice | off
        self.tmpdir = tmpdir
        self.soffice = soffice
        self.office_ok = mode in ("auto", "office") and has_pywin32()
        self._apps = {}
        self._com_ready = False
        self._n = 0

    def convert(self, src, target_ext):
        """Возвращает (путь к временному OOXML-файлу, описание метода)."""
        self._n += 1
        ext = os.path.splitext(src)[1].lower()
        # Короткое ASCII-имя: Office и LibreOffice плохо переносят длинные пути
        tmp_in = os.path.join(self.tmpdir, f"in{self._n}{ext}")
        shutil.copyfile(src, tmp_in)
        errors = []
        try:
            if self.office_ok:
                dst = os.path.join(self.tmpdir, f"in{self._n}{target_ext}")
                try:
                    self._office(tmp_in, dst, target_ext)
                    return dst, f"MS Office -> {target_ext}"
                except Exception as e:
                    errors.append(f"Office: {type(e).__name__}: {e}")
                    if self.mode == "office":
                        raise RuntimeError("; ".join(errors))
            if self.mode in ("auto", "libreoffice") and self.soffice:
                try:
                    return (
                        self._libreoffice(tmp_in, target_ext),
                        f"LibreOffice -> {target_ext}",
                    )
                except Exception as e:
                    errors.append(f"LibreOffice: {type(e).__name__}: {e}")
            raise RuntimeError(
                "; ".join(errors) or "нет доступного конвертера старых форматов"
            )
        finally:
            try:
                os.remove(tmp_in)
            except OSError:
                pass

    def _libreoffice(self, tmp_in, target_ext):
        profile = os.path.join(self.tmpdir, "lo_profile")
        from pathlib import Path

        cmd = [
            self.soffice,
            "--headless",
            "--norestore",
            "--nolockcheck",
            "--nodefault",
            f"-env:UserInstallation={Path(profile).as_uri()}",
            "--convert-to",
            target_ext.lstrip("."),
            "--outdir",
            self.tmpdir,
            tmp_in,
        ]
        flags = subprocess.CREATE_NO_WINDOW if IS_WIN else 0
        proc = subprocess.run(
            cmd, capture_output=True, timeout=600, creationflags=flags
        )
        dst = os.path.splitext(tmp_in)[0] + target_ext
        if not os.path.isfile(dst):
            msg = (proc.stderr or proc.stdout or b"").decode("utf-8", "replace").strip()
            raise RuntimeError(msg[:300] or f"код возврата {proc.returncode}")
        return dst

    def _app(self, progid):
        import win32com.client

        app = self._apps.get(progid)
        if app is None:
            app = win32com.client.DispatchEx(progid)
            try:
                app.AutomationSecurity = (
                    3  # msoAutomationSecurityForceDisable: макросы не запускать
                )
            except Exception:
                pass
            if progid == "Word.Application":
                app.Visible = False
                app.DisplayAlerts = 0
            elif progid == "Excel.Application":
                app.Visible = False
                app.DisplayAlerts = False
                app.AskToUpdateLinks = False
            else:
                app.DisplayAlerts = 1  # ppAlertsNone
            self._apps[progid] = app
        return app

    def _office(self, src, dst, target_ext):
        import pythoncom

        if not self._com_ready:
            pythoncom.CoInitialize()
            self._com_ready = True
        src, dst = os.path.abspath(src), os.path.abspath(dst)
        if target_ext == ".docx":
            doc = self._app("Word.Application").Documents.Open(
                FileName=src,
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                PasswordDocument=NO_PASSWORD,
                Visible=False,
                NoEncodingDialog=True,
            )
            try:
                try:
                    doc.SaveAs2(FileName=dst, FileFormat=16)  # wdFormatDocumentDefault
                except AttributeError:  # Word 2007
                    doc.SaveAs(FileName=dst, FileFormat=16)
            finally:
                doc.Close(SaveChanges=0)
        elif target_ext == ".xlsx":
            wb = self._app("Excel.Application").Workbooks.Open(
                Filename=src,
                UpdateLinks=0,
                ReadOnly=True,
                Password=NO_PASSWORD,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
            try:
                wb.SaveAs(Filename=dst, FileFormat=51)  # xlOpenXMLWorkbook
            finally:
                wb.Close(SaveChanges=False)
        else:
            pres = self._app("PowerPoint.Application").Presentations.Open(
                FileName=src,
                ReadOnly=True,
                Untitled=False,
                WithWindow=False,
            )
            try:
                pres.SaveAs(FileName=dst, FileFormat=24)  # ppSaveAsOpenXMLPresentation
            finally:
                pres.Close()

    def close(self):
        for app in self._apps.values():
            try:
                app.Quit()
            except Exception:
                pass
        self._apps.clear()


# ---------------------------------------------------------------------------
# Конвертер одного файла (живёт внутри рабочего процесса)
# ---------------------------------------------------------------------------
class FileConverter:
    def __init__(self, opts):
        self.opts = opts
        self.tmpdir = tempfile.mkdtemp(prefix="w", dir=opts["tmp_root"])
        self.legacy = LegacyConverter(opts["legacy"], self.tmpdir, opts["soffice"])
        self._md = None

    @property
    def md(self):
        if self._md is None:
            from markitdown import MarkItDown

            self._md = MarkItDown(enable_plugins=False)
        return self._md

    def close(self):
        self.legacy.close()
        shutil.rmtree(self.tmpdir, ignore_errors=True)

    # -- диспетчер ----------------------------------------------------------
    def run(self, job):
        t0 = time.time()
        try:
            body, info = self.convert(job)
            self.write(job, body, info)
            return {
                "status": info["status"],
                "chars": len(body),
                "pages": info.get("pages", ""),
                "method": info["method"],
                "message": info.get("note", ""),
                "seconds": round(time.time() - t0, 1),
            }
        except Exception as e:
            return {
                "status": "ERROR",
                "message": f"{type(e).__name__}: {e}"[:500],
                "seconds": round(time.time() - t0, 1),
            }

    def convert(self, job):
        kind, src, ext = job["kind"], job["src"], job["ext"]
        if kind == "native" and ext == ".pdf":
            drawing = self.drawing_pdf(job)
            if drawing is not None:
                return drawing
        if kind == "native":
            try:
                body = self.markitdown(src, EXT_ALIAS.get(ext))
                method = "markitdown"
            except Exception:
                if ext != ".xls" or not job["legacy_available"]:
                    raise
                # Часть «.xls» на деле HTML/XML-выгрузки или повреждены: пробуем пересохранить
                body, method = self.via_legacy(src, ".xlsx")
            if ext == ".pdf":
                return self.pdf_postprocess(src, body)
            return body, {"status": "OK", "method": method}
        if kind == "legacy":
            body, method = self.via_legacy(src, LEGACY_EXT[ext])
            return body, {"status": "OK", "method": method}
        if kind == "image":
            if is_drawing_path(job["rel"]):
                return self.drawing_image(job)
            return self.ocr_image(src)
        if kind == "cad":
            return self.cad(job)
        if kind == "eml":
            return self.eml(src), {"status": "OK", "method": "email"}
        if kind == "code":
            return self.code(src, CODE_EXT[ext]), {"status": "OK", "method": "text"}
        raise ValueError(f"неизвестный тип задания {kind}")

    def markitdown(self, path, ext_override=None):
        from markitdown import StreamInfo

        si = StreamInfo(extension=ext_override) if ext_override else None
        return self.md.convert_local(path, stream_info=si).markdown or ""

    def via_legacy(self, src, target_ext):
        tmp, method = self.legacy.convert(src, target_ext)
        try:
            return self.markitdown(tmp), method + " + markitdown"
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass

    # -- PDF: поиск страниц-сканов и OCR ----------------------------------
    def pdf_postprocess(self, src, body):
        try:
            import pymupdf
        except ImportError:
            status = "LOW_TEXT" if len(body.strip()) < 200 else "OK"
            note = (
                "мало текста, возможно скан (pymupdf не установлен)"
                if status == "LOW_TEXT"
                else ""
            )
            return body, {"status": status, "method": "markitdown", "note": note}

        with open_doc(pymupdf, src) as doc:
            if doc.needs_pass:
                return body, {
                    "status": "OK",
                    "method": "markitdown",
                    "note": "PDF защищён паролем",
                }
            pages = doc.page_count
            low = [
                i
                for i, page in enumerate(doc)
                if len(page.get_text("text").strip()) < self.opts["min_chars"]
                and page.get_images()
            ]
            info = {"status": "OK", "method": "markitdown", "pages": pages}
            if not low:
                return body, info
            if not self.opts["ocr"]:
                info.update(
                    status="LOW_TEXT",
                    note=f"без текстового слоя стр. {page_ranges(low)} из {pages}",
                )
                return body, info
            texts = self.ocr_pages(doc, low)
        info.update(
            status="OCR",
            method="markitdown + OCR",
            note=f"OCR стр. {page_ranges(low)} из {pages}",
        )
        sections = [
            f"## Стр. {i + 1} (OCR)\n\n{t or '_(текст не распознан)_'}"
            for i, t in texts
        ]
        if len(low) == pages:
            return "\n\n".join(sections), info
        return (
            body.rstrip()
            + "\n\n---\n\n# Текст страниц-сканов (OCR)\n\n"
            + "\n\n".join(sections),
            info,
        )

    def ocr_page(self, page):
        # Большие листы распознаются с меньшим dpi, чтобы уложиться в OCR_MAX_PIXELS
        area_in2 = (page.rect.width / 72) * (page.rect.height / 72)
        dpi = min(self.opts["ocr_dpi"], int((OCR_MAX_PIXELS / area_in2) ** 0.5))
        tp = page.get_textpage_ocr(
            language=self.opts["ocr_lang"],
            dpi=max(dpi, 150),
            full=True,
            tessdata=self.opts["tessdata"],
        )
        return fix_homoglyphs(page_text(page, tp))

    def ocr_pages(self, doc, indexes):
        return [(i, self.ocr_page(doc[i])) for i in indexes]

    # -- чертежи и схемы ----------------------------------------------------
    def drawing_pdf(self, job):
        """Карточка чертежа для PDF, либо None, если это обычный документ."""
        try:
            import pymupdf
        except ImportError:
            return None
        with open_doc(pymupdf, job["src"]) as doc:
            if doc.needs_pass or doc.page_count == 0:
                return None
            large = sum(
                1 for p in doc if max(p.rect.width, p.rect.height) >= LARGE_SHEET_PT
            )
            if not (is_drawing_path(job["rel"]) or large * 2 >= doc.page_count):
                return None
            return self.drawing_card(doc, job, "PDF")

    def drawing_image(self, job):
        import pymupdf

        with open_doc(pymupdf, job["src"]) as img:
            pdf_bytes = img.convert_to_pdf()
        with pymupdf.open("pdf", pdf_bytes) as doc:
            return self.drawing_card(doc, job, "изображение")

    def drawing_card(self, doc, job, source_kind):
        sheets, ocr_done, ocr_missing = [], [], []
        for i, page in enumerate(doc):
            text = page_text(page)
            if len(text) < self.opts["min_chars"] and page.get_images():
                if self.opts["ocr"]:
                    text = self.ocr_page(page)
                    ocr_done.append(i)
                else:
                    ocr_missing.append(i)
            sheets.append(text)
        preview = self.save_preview(doc[0], job)
        lines = [drawing_notice(job), ""]
        if preview:
            title = "Превью листа 1" if len(sheets) > 1 else "Превью"
            lines += [f"![{title}](<{preview}>)", ""]
        for i, text in enumerate(sheets):
            head = "## Надписи на чертеже" if len(sheets) == 1 else f"## Лист {i + 1}"
            if i in ocr_done:
                head += " (OCR)"
            lines += [head, "", text or "_(надписи не найдены)_", ""]
        info = {
            "status": "LOW_TEXT" if ocr_missing else "DRAWING",
            "method": f"чертёж, {source_kind}"
            + (", OCR" if ocr_done else ", текстовый слой"),
            "pages": len(sheets),
        }
        if ocr_done:
            info["note"] = f"OCR листов {page_ranges(ocr_done)}"
        if ocr_missing:
            info["note"] = f"не распознаны листы {page_ranges(ocr_missing)}"
        return "\n".join(lines), info

    def save_preview(self, page, job):
        import pymupdf

        try:
            zoom = PREVIEW_PX / max(page.rect.width, page.rect.height)
            pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
            name = os.path.basename(job["dst"])[: -len(".md")] + ".preview.jpg"
            os.makedirs(os.path.dirname(job["dst"]), exist_ok=True)
            with open(os.path.join(os.path.dirname(job["dst"]), name), "wb") as f:
                f.write(pix.tobytes("jpg", jpg_quality=70))
            return name
        except Exception:
            return None

    def cad(self, job):
        src, ext = job["src"], job["ext"]
        oda = self.opts["oda"]
        if ext == ".dwg" and not oda:
            body = (
                drawing_notice(job)
                + "\n\n_Надписи из DWG не извлечены: на компьютере не установлен"
                " бесплатный ODA File Converter._"
            )
            return body, {"status": "DRAWING_NO_TEXT", "method": "чертёж, DWG"}
        import ezdxf

        tmp = os.path.join(self.tmpdir, "cad" + ext)
        shutil.copyfile(src, tmp)  # короткий путь для ODA и ezdxf
        try:
            if ext == ".dwg":
                from ezdxf.addons import odafc

                key = "win_exec_path" if IS_WIN else "unix_exec_path"
                ezdxf.options.set("odafc-addon", key, oda)
                doc = odafc.readfile(tmp)
            else:
                from ezdxf import recover

                doc, _ = recover.readfile(tmp)
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass
        texts, seen = [], set()
        for layout in [doc.modelspace()] + list(doc.layouts):
            for e in layout:
                for t in cad_entity_texts(e):
                    t = " ".join(t.split())
                    if t and t not in seen:
                        seen.add(t)
                        texts.append(t)
        body = drawing_notice(job) + "\n\n## Надписи на чертеже\n\n"
        body += "\n".join(texts) if texts else "_(надписи не найдены)_"
        return body, {
            "status": "DRAWING",
            "method": f"чертёж, {ext.upper().lstrip('.')}",
            "note": f"надписей: {len(texts)}",
        }

    def ocr_image(self, src):
        import pymupdf

        with open_doc(pymupdf, src) as img:
            pdf_bytes = img.convert_to_pdf()
        with pymupdf.open("pdf", pdf_bytes) as doc:
            texts = self.ocr_pages(doc, range(doc.page_count))
        pages = len(texts)
        if not any(t for _, t in texts):
            return "_(текст на изображении не обнаружен)_", {
                "status": "NO_TEXT",
                "method": "OCR",
                "pages": pages,
            }
        if pages == 1:
            body = texts[0][1]
        else:
            body = "\n\n".join(f"## Стр. {i + 1}\n\n{t}" for i, t in texts)
        return body, {"status": "OCR", "method": "OCR", "pages": pages}

    # -- прочие форматы -----------------------------------------------------
    def eml(self, src):
        from markitdown import StreamInfo

        with open(src, "rb") as f:
            msg = email.message_from_binary_file(f, policy=email.policy.default)
        lines = [f"## {msg['subject'] or '(без темы)'}", ""]
        for h in ("From", "To", "Cc", "Date"):
            if msg[h]:
                lines.append(f"**{h}:** {msg[h]}  ")
        part = msg.get_body(preferencelist=("plain", "html"))
        text = ""
        if part is not None:
            text = part.get_content()
            if part.get_content_type() == "text/html":
                text = self.md.convert_stream(
                    io.BytesIO(text.encode("utf-8")),
                    stream_info=StreamInfo(extension=".html", charset="utf-8"),
                ).markdown
        lines += ["", text.strip()]
        names = [p.get_filename() for p in msg.iter_attachments() if p.get_filename()]
        if names:
            lines += ["", "**Вложения (не конвертированы):** " + "; ".join(names)]
        return "\n".join(lines)

    def code(self, src, lang):
        from charset_normalizer import from_bytes

        with open(src, "rb") as f:
            raw = f.read()
        best = from_bytes(raw).best()
        text = str(best) if best is not None else raw.decode("utf-8", "replace")
        fence = "````" if "```" in text else "```"
        return f"{fence}{lang}\n{text.rstrip()}\n{fence}"

    # -- запись результата -------------------------------------------------
    def write(self, job, body, info):
        fm = [
            "---",
            f"source: {yaml_str(job['display_src'])}",
            f"vessel: {yaml_str(job['vessel'])}",
            f"relative_path: {yaml_str(job['rel'])}",
            f"size_kb: {job['size'] // 1024}",
            f"modified: {yaml_str(fmt_time(job['mtime']))}",
            f"converted: {yaml_str(dt.datetime.now().strftime('%Y-%m-%d %H:%M'))}",
            f"method: {yaml_str(info['method'])}",
        ]
        if info.get("pages"):
            fm.append(f"pages: {info['pages']}")
        fm.append(f"status: {yaml_str(info['status'].lower())}")
        if info.get("note"):
            fm.append(f"note: {yaml_str(info['note'])}")
        fm.append("---")
        text = (
            "\n".join(fm)
            + f"\n\n# {os.path.basename(job['rel'])}\n\n"
            + body.strip()
            + "\n"
        )
        os.makedirs(os.path.dirname(job["dst"]), exist_ok=True)
        tmp = job["dst"] + ".tmp"
        with open(tmp, "w", encoding="utf-8", newline="\n") as f:
            f.write(text)
        os.replace(tmp, job["dst"])


_CYR2LAT = str.maketrans("АВЕКМНОРСТХУаеорсух", "ABEKMHOPCTXYaeopcyx")
_LAT2CYR = str.maketrans("ABEKMHOPCTXYaeopcyx", "АВЕКМНОРСТХУаеорсух")
_CYR_ONLY = re.compile(r"[а-яёА-ЯЁ]")
_LAT = re.compile(r"[A-Za-z]")


def _fix_token(tok):
    cyr = _CYR_ONLY.findall(tok)
    if not cyr:
        return tok
    pure_cyr = [
        c for c in cyr if c.translate(_CYR2LAT) == c
    ]  # буквы без латинского двойника
    if not _LAT.search(tok) and not any(ch.isdigit() for ch in tok):
        return tok  # обычное русское слово
    if pure_cyr:
        return tok.translate(_LAT2CYR)  # русское слово с латинскими «двойниками»
    return tok.translate(_CYR2LAT)  # код/номер: «СН-310-002» -> «CH-310-002»


def fix_homoglyphs(text):
    """OCR rus+eng путает похожие буквы в номерах чертежей и тегах; приводим к одной письменности."""
    return re.sub(r"\S+", lambda m: _fix_token(m.group(0)), text)


def page_text(page, textpage=None):
    """Текст страницы по блокам, без позиционных пробелов."""
    blocks = page.get_text("blocks", sort=True, textpage=textpage)
    out = []
    for b in blocks:
        if b[6] != 0:
            continue
        lines = [" ".join(l.split()) for l in b[4].splitlines()]
        lines = [l for l in lines if l]
        if lines:
            out.append("\n".join(lines))
    return "\n\n".join(out)


def is_drawing_path(rel):
    low = rel.lower()
    return any(w in low for w in DRAWING_WORDS)


def drawing_notice(job):
    return (
        "> **Чертёж / схема.** Размеры, допуски и технические параметры брать только"
        f" из оригинала: `{job['display_src']}`"
    )


def cad_entity_texts(e):
    kind = e.dxftype()
    try:
        if kind in ("TEXT", "ATTRIB", "ATTDEF"):
            yield e.dxf.text
        elif kind == "MTEXT":
            yield e.plain_text()
        elif kind == "INSERT":
            for a in e.attribs:
                yield a.dxf.text
    except Exception:
        return


def open_doc(pymupdf, path):
    try:
        return pymupdf.open(path)
    except Exception:
        # MuPDF может не принять длинный путь \\?\ — читаем через поток
        with open(path, "rb") as f:
            return pymupdf.open(
                stream=f.read(), filetype=os.path.splitext(path)[1].lstrip(".").lower()
            )


def worker_main(conn, opts):
    signal.signal(signal.SIGINT, signal.SIG_IGN)  # Ctrl+C обрабатывает главный процесс
    conv = FileConverter(opts)
    try:
        while True:
            try:
                job = conn.recv()
            except EOFError:
                break
            if job is None:
                break
            conn.send(conv.run(job))
    finally:
        conv.close()


class Worker:
    """Отдельный процесс-конвертер: зависший или упавший файл не останавливает весь пакет."""

    def __init__(self, opts):
        self.opts = opts
        self.proc = None
        self.conn = None

    def _start(self):
        parent, child = mp.Pipe()
        self.proc = mp.Process(target=worker_main, args=(child, self.opts), daemon=True)
        self.proc.start()
        child.close()
        self.conn = parent

    def run(self, job, timeout):
        if self.proc is None or not self.proc.is_alive():
            self._start()
        try:
            self.conn.send(job)
            if self.conn.poll(timeout):
                return self.conn.recv()
            result = {
                "status": "TIMEOUT",
                "message": f"дольше {timeout} с",
                "seconds": timeout,
            }
        except (EOFError, OSError, BrokenPipeError):
            result = {
                "status": "CRASH",
                "message": "процесс конвертации аварийно завершился",
            }
        self.kill()
        return result

    def kill(self):
        if self.proc is not None and self.proc.is_alive():
            self.proc.kill()
            self.proc.join(5)
        self.proc = None

    def close(self):
        if self.proc is not None and self.proc.is_alive():
            try:
                self.conn.send(None)
                self.proc.join(30)
            except OSError:
                pass
        self.kill()


# ---------------------------------------------------------------------------
# Главный процесс
# ---------------------------------------------------------------------------
def classify(ext):
    if ext in NATIVE_EXT or ext in EXT_ALIAS:
        return "native"
    if ext in LEGACY_EXT:
        return "legacy"
    if ext in IMAGE_EXT:
        return "image"
    if ext in CAD_EXT:
        return "cad"
    if ext == ".eml":
        return "eml"
    if ext in CODE_EXT:
        return "code"
    return None


def walk_sources(sources, out_root, errors):
    for vessel, root in sources:
        root_lp = long_path(root)

        def onerror(e, vessel=vessel, root=root):
            errors.append(
                {
                    "vessel": vessel,
                    "source": plain_path(getattr(e, "filename", "") or root),
                    "status": "ACCESS",
                    "message": str(e),
                }
            )

        for dirpath, dirnames, filenames in os.walk(root_lp, onerror=onerror):
            # не заходить в выходной каталог, если он вдруг лежит внутри источника
            dirnames[:] = sorted(
                d
                for d in dirnames
                if os.path.normcase(os.path.join(dirpath, d))
                != os.path.normcase(out_root)
            )
            for name in sorted(filenames):
                if name.lower() in SKIP_NAMES or name.startswith("~$"):
                    continue
                full = os.path.join(dirpath, name)
                rel = full[len(root_lp) :].lstrip("\\/")
                try:
                    st = os.stat(full)
                except OSError as e:
                    errors.append(
                        {
                            "vessel": vessel,
                            "source": plain_path(full),
                            "status": "ACCESS",
                            "message": str(e),
                        }
                    )
                    continue
                yield {
                    "vessel": vessel,
                    "src": full,
                    "rel": rel,
                    "display_src": os.path.join(root, rel),
                    "ext": os.path.splitext(name)[1].lower(),
                    "size": st.st_size,
                    "mtime": st.st_mtime,
                }


def write_csv(path, rows, fields):
    with open(
        path, "w", encoding="utf-8-sig", newline=""
    ) as f:  # BOM + ';' — открывается в русском Excel
        w = csv.DictWriter(f, fieldnames=fields, delimiter=";", extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def do_scan(files, out_root, legacy_available, ocr, max_size):
    stats = {}
    labels = {
        "native": "MarkItDown",
        "legacy": "через Office/LibreOffice",
        "image": "OCR изображения",
        "cad": "чертёж DWG/DXF",
        "eml": "письмо .eml",
        "code": "текст/XML",
        None: "не поддерживается",
    }
    for f in files:
        kind = classify(f["ext"])
        label = labels[kind]
        if kind == "legacy" and not legacy_available:
            label = "НЕТ конвертера .doc/.ppt"
        if kind == "image" and not ocr:
            label = "изображение: OCR недоступен"
        if kind == "image" and is_drawing_path(f["rel"]):
            label = "чертёж (изображение)"
        if kind == "native" and f["ext"] == ".pdf" and is_drawing_path(f["rel"]):
            label = "чертёж (PDF)"
        if f["size"] > max_size:
            label = "больше --max-size-mb"
        key = (f["vessel"], f["ext"] or "(без расширения)", label)
        s = stats.setdefault(key, {"count": 0, "size": 0})
        s["count"] += 1
        s["size"] += f["size"]
    rows = [
        {
            "vessel": v,
            "ext": e,
            "handling": k,
            "count": s["count"],
            "size_mb": round(s["size"] / 2**20, 1),
        }
        for (v, e, k), s in sorted(
            stats.items(), key=lambda kv: (kv[0][0], -kv[1]["count"])
        )
    ]
    path = os.path.join(out_root, "_inventory.csv")
    write_csv(path, rows, ["vessel", "ext", "handling", "count", "size_mb"])
    print(
        f"\n{'Судно':<10} {'Расширение':<14} {'Обработка':<26} {'Файлов':>8} {'МБ':>10}"
    )
    for r in rows:
        print(
            f"{r['vessel']:<10} {r['ext']:<14} {r['handling']:<26} {r['count']:>8} {r['size_mb']:>10}"
        )
    total = sum(r["count"] for r in rows)
    print(f"\nВсего файлов: {total}. Таблица сохранена: {plain_path(path)}")


def build_index(out_root):
    entries, counts = [], {}
    for dirpath, dirnames, filenames in os.walk(out_root):
        dirnames[:] = sorted(dirnames)
        for name in sorted(filenames):
            if not name.endswith(".md") or (
                dirpath == out_root and name.startswith("_")
            ):
                continue
            full = os.path.join(dirpath, name)
            fm = read_front_matter(full)
            status = fm.get("status", "?")
            counts[status] = counts.get(status, 0) + 1
            if status == "no_text":
                continue
            rel = full[len(out_root) :].lstrip("\\/").replace("\\", "/")
            entries.append((rel, fm))
    lines = [
        "# Индекс судовой документации в Markdown",
        "",
        f"Сформирован {dt.datetime.now().strftime('%Y-%m-%d %H:%M')}. Файлов .md: {sum(counts.values())}.",
        "",
        "| Статус | Файлов |",
        "|---|---|",
    ]
    lines += [f"| {k} | {v} |" for k, v in sorted(counts.items())]
    if counts.get("no_text"):
        lines += [
            "",
            "Изображения без распознанного текста (no_text) в список ниже не включены.",
        ]
    last_vessel = last_dir = None
    for rel, fm in entries:
        parts = rel.split("/")
        vessel, folder = parts[0], "/".join(parts[1:-1]) or "(корень)"
        if vessel != last_vessel:
            lines += ["", f"## {vessel}"]
            last_vessel, last_dir = vessel, None
        if folder != last_dir:
            lines += ["", f"### {folder}", ""]
            last_dir = folder
        extra = []
        if fm.get("pages"):
            extra.append(f"{fm['pages']} стр.")
        if fm.get("status", "").startswith("drawing"):
            extra.insert(0, "чертёж")
        if fm.get("status") == "low_text":
            extra.append("сканы не распознаны: " + fm.get("note", ""))
        suffix = f" — {'; '.join(extra)}" if extra else ""
        lines.append(f"- [{parts[-1][:-3]}](<{rel}>){suffix}")
    with open(
        os.path.join(out_root, "_INDEX.md"), "w", encoding="utf-8", newline="\n"
    ) as f:
        f.write("\n".join(lines) + "\n")


def parse_args():
    ap = argparse.ArgumentParser(
        description="Конвертация судовой документации в Markdown"
    )
    ap.add_argument(
        "--source",
        action="append",
        metavar="ИМЯ=ПУТЬ",
        help="папка-источник; можно указать несколько раз. По умолчанию: "
        + "; ".join(f"{n}={p}" for n, p in DEFAULT_SOURCES),
    )
    ap.add_argument(
        "--out",
        default=DEFAULT_OUTPUT,
        help=f"выходная папка (по умолчанию {DEFAULT_OUTPUT})",
    )
    ap.add_argument(
        "--scan",
        action="store_true",
        help="только инвентаризация: что и сколько будет обработано",
    )
    ap.add_argument(
        "--no-ocr",
        action="store_true",
        help="не распознавать сканы (по умолчанию распознаются)",
    )
    ap.add_argument("--ocr", action="store_true", help=argparse.SUPPRESS)
    ap.add_argument("--ocr-lang", default="rus+eng")
    ap.add_argument("--ocr-dpi", type=int, default=300)
    ap.add_argument("--tessdata", help="папка с rus.traineddata и eng.traineddata")
    ap.add_argument(
        "--legacy",
        choices=["auto", "office", "libreoffice", "off"],
        default="auto",
        help="чем пересохранять .doc/.ppt/.rtf/.xlsb/.od*: auto = MS Office, затем LibreOffice",
    )
    ap.add_argument(
        "--workers", type=int, default=max(1, min(4, (os.cpu_count() or 2) - 1))
    )
    ap.add_argument(
        "--timeout",
        type=int,
        default=1800,
        help="лимит на один файл, с (по умолчанию 1800)",
    )
    ap.add_argument("--max-size-mb", type=int, default=500)
    ap.add_argument(
        "--min-chars",
        type=int,
        default=30,
        help="страница PDF с картинкой и меньшим числом символов считается сканом",
    )
    ap.add_argument(
        "--force",
        action="store_true",
        help="переконвертировать всё, даже неизменённые файлы",
    )
    return ap.parse_args()


def main():
    try:
        sys.stdout.reconfigure(errors="replace")
    except AttributeError:
        pass
    args = parse_args()

    sources = DEFAULT_SOURCES
    if args.source:
        sources = []
        for s in args.source:
            name, sep, path = s.partition("=")
            if not sep:
                sys.exit(f"--source должен быть в виде ИМЯ=ПУТЬ, получено: {s}")
            sources.append((name.strip(), path.strip()))
    missing = [(n, p) for n, p in sources if not os.path.isdir(p)]
    for n, p in missing:
        print(f"ВНИМАНИЕ: папка {n} не найдена: {p}")
    sources = [s for s in sources if s not in missing]
    if not sources:
        sys.exit("Нет доступных папок-источников.")

    out_root = long_path(args.out)
    for n, p in sources:
        src_lp = long_path(p)
        o, sp = os.path.normcase(out_root), os.path.normcase(src_lp.rstrip("\\/"))
        if o == sp or o.startswith(sp + os.sep):
            print(
                f"Выходная папка находится внутри источника {n}: она будет пропущена при обходе."
            )
    os.makedirs(out_root, exist_ok=True)

    soffice = find_soffice() if args.legacy in ("auto", "libreoffice") else None
    office = args.legacy in ("auto", "office") and has_pywin32()
    legacy_available = bool(soffice or office)

    oda = find_oda()
    args.ocr, tessdata, ocr_problem = not args.no_ocr, None, ""
    if args.ocr:
        try:
            import pymupdf  # noqa: F401

            tessdata, ocr_problem = ensure_tessdata(
                args.tessdata, args.ocr_lang.split("+")
            )
        except ImportError:
            ocr_problem = "не установлен пакет pymupdf (запустите install.bat)"
        if ocr_problem:
            args.ocr = False
            print(
                f"\n!!! OCR НЕДОСТУПЕН: {ocr_problem}.\n"
                "!!! Сканы будут помечены LOW_TEXT/NEEDS_OCR; после устранения причины"
                " просто запустите конвертацию ещё раз — будут обработаны только они.\n"
            )

    print("Источники:")
    for n, p in sources:
        print(f"  {n}: {p}")
    print(f"Результат: {plain_path(out_root)}")
    print(
        f"Старые форматы Office: {'MS Office' if office else ''}"
        f"{' + ' if office and soffice else ''}{'LibreOffice' if soffice else ''}"
        f"{'' if legacy_available else 'НЕТ конвертера (.doc/.ppt/.rtf будут пропущены)'}"
    )
    print(
        f"OCR: {'включён, ' + args.ocr_lang + ', ' + tessdata if args.ocr else 'выключен'}"
    )
    print(
        f"DWG: {'ODA File Converter ' + oda if oda else 'нет ODA File Converter (только карточки)'}"
    )

    errors = []
    print("\nСбор списка файлов...")
    files = list(walk_sources(sources, out_root, errors))
    print(f"Найдено файлов: {len(files)}")

    if args.scan:
        do_scan(files, out_root, legacy_available, args.ocr, args.max_size_mb * 2**20)
        return

    rows, jobs = list(errors), []
    for f in files:
        f["dst"] = os.path.join(out_root, f["vessel"], f["rel"] + ".md")
        f["kind"] = kind = classify(f["ext"])
        f["legacy_available"] = legacy_available
        status = None
        if kind is None:
            status = "UNSUPPORTED"
        elif kind == "image" and not args.ocr:
            status = "NEEDS_OCR"
        elif kind == "legacy" and not legacy_available:
            status = "NO_CONVERTER"
        elif f["size"] > args.max_size_mb * 2**20:
            status = "TOO_LARGE"
        elif (
            not args.force
            and os.path.isfile(f["dst"])
            and os.path.getmtime(f["dst"]) >= f["mtime"]
        ):
            prev = read_front_matter(f["dst"]).get("status", "")
            redo = (args.ocr and prev == "low_text") or (
                oda and prev == "drawing_no_text"
            )
            if not redo:
                status = "UNCHANGED"
        if status:
            rows.append(
                {
                    **f,
                    "source": f["display_src"],
                    "status": status,
                    "output": plain_path(f["dst"]) if status == "UNCHANGED" else "",
                }
            )
        else:
            jobs.append(f)

    jobs.sort(
        key=lambda j: -j["size"]
    )  # крупные файлы первыми — ровнее загрузка процессов
    total = len(jobs)
    print(f"К конвертации: {total}, пропущено: {len(rows)} (детали в _report.csv)\n")

    tmp_root = tempfile.mkdtemp(prefix="vsl2md_")
    opts = {
        "tmp_root": tmp_root,
        "legacy": args.legacy,
        "soffice": soffice,
        "ocr": args.ocr,
        "ocr_lang": args.ocr_lang,
        "ocr_dpi": args.ocr_dpi,
        "tessdata": tessdata,
        "oda": oda,
        "min_chars": args.min_chars,
    }
    job_q, res_q = queue.Queue(), queue.Queue()
    for j in jobs:
        job_q.put(j)
    stop = threading.Event()
    workers = [Worker(opts) for _ in range(max(1, args.workers))]

    def thread_main(worker):
        while not stop.is_set():
            try:
                job = job_q.get_nowait()
            except queue.Empty:
                break
            res_q.put((job, worker.run(job, args.timeout)))

    threads = [
        threading.Thread(target=thread_main, args=(w,), daemon=True) for w in workers
    ]
    for t in threads:
        t.start()

    done, t_start = 0, time.time()
    interrupted = False
    try:
        while done < total:
            try:
                job, res = res_q.get(timeout=0.5)
            except queue.Empty:
                if not any(t.is_alive() for t in threads) and res_q.empty():
                    break
                continue
            done += 1
            rows.append(
                {
                    **job,
                    **res,
                    "source": job["display_src"],
                    "output": plain_path(job["dst"])
                    if res["status"] not in ("ERROR", "TIMEOUT", "CRASH")
                    else "",
                }
            )
            eta = (time.time() - t_start) / done * (total - done)
            print(
                f"[{done:>6}/{total}] {res['status']:<8} {res.get('seconds', ''):>6}s  "
                f"ост. ~{int(eta // 60)} мин  {os.path.join(job['vessel'], job['rel'])}",
                flush=True,
            )
    except KeyboardInterrupt:
        interrupted = True
        print(
            "\nПрервано. Уже созданные .md сохранены; повторный запуск продолжит с оставшихся файлов."
        )
        stop.set()
        for w in workers:
            w.kill()
    finally:
        stop.set()
        for t in threads:
            t.join(timeout=0 if interrupted else 60)
        for w in workers:
            w.close()
        shutil.rmtree(tmp_root, ignore_errors=True)

    for r in rows:
        r["size_kb"] = r.get("size", 0) // 1024
        r.setdefault("source", r.get("display_src", ""))
    report = os.path.join(out_root, "_report.csv")
    write_csv(
        report,
        rows,
        [
            "vessel",
            "status",
            "source",
            "output",
            "ext",
            "size_kb",
            "pages",
            "chars",
            "method",
            "message",
            "seconds",
        ],
    )
    build_index(out_root)

    counts = {}
    for r in rows:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    print("\nИтог:")
    for k, v in sorted(counts.items(), key=lambda kv: -kv[1]):
        print(f"  {k:<13} {v:>7}  {STATUS_HELP.get(k, '')}")
    print(f"\nОтчёт по каждому файлу: {plain_path(report)}")
    print(f"Оглавление: {plain_path(os.path.join(out_root, '_INDEX.md'))}")
    if interrupted:
        sys.exit(130)


if __name__ == "__main__":
    mp.freeze_support()
    main()
