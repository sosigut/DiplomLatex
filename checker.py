import os
import shutil
from dataclasses import dataclass, field
from typing import List

try:
    import win32com.client
    from win32com.client import dynamic
except Exception as e:
    win32com = None
    dynamic = None
    WIN32_IMPORT_ERROR = e
else:
    WIN32_IMPORT_ERROR = None

CM_TO_PT = 28.3464567

WD_ALIGN_CENTER = 1
WD_WITHIN_TABLE = 12
WD_HEADER_FOOTER_PRIMARY = 1
WD_LINE_SPACE_SINGLE = 0
WD_STATISTIC_PAGES = 2
WD_ACTIVE_END_PAGE_NUMBER = 3

WD_COLOR_RED = 255
WD_FIELD_PAGE = 33


def pt_to_cm(pt):
    return pt / CM_TO_PT

@dataclass
class CheckIssue:
    rule: str
    severity: str
    location: str
    message: str
    priority: int = 3


@dataclass
class CheckReport:
    issues: List[CheckIssue] = field(default_factory=list)

    def to_text(self):
        if not self.issues:
            return "✅ Ошибок не найдено."

        self.issues.sort(key=lambda x: (x.priority, x.location))

        lines = []
        for i, issue in enumerate(self.issues, 1):
            lines.append(f"❌ Ошибка #{i}")
            lines.append(f"   📌 Раздел: {issue.rule}")
            lines.append(f"   📍 Где: {issue.location}")
            lines.append(f"   💬 Что нужно исправить: {issue.message}")
            lines.append("")
        return "\n".join(lines)

class WordMethodicalChecker:
    def __init__(self, visible: bool = False, mark_document: bool = True):
        if dynamic is None:
            raise RuntimeError(f"pywin32 не установлен: {WIN32_IMPORT_ERROR}")
        self.word = None
        self.visible = visible
        self.mark_document = mark_document

    def __enter__(self):
        try:
            self.word = dynamic.Dispatch("Word.Application")

            self.word.DisplayAlerts = 0

            try:
                self.word.Visible = self.visible
            except Exception as e:
                print(f"Warning: Не удалось установить Word.Visible: {e}")

        except Exception as e:
            raise RuntimeError(f"Не удалось запустить Microsoft Word через COM: {e}")

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.word:
            self.word.Quit(False)

    def open_document(self, path, readonly=True):
        return self.word.Documents.Open(os.path.abspath(path), ReadOnly=readonly)

    def check(self, path):
        report = CheckReport()

        checked_path = None
        open_path = path
        readonly = True

        if self.mark_document:
            base, ext = os.path.splitext(path)
            checked_path = base + "_checked" + ext
            shutil.copy(path, checked_path)
            open_path = checked_path
            readonly = False

        doc = self.open_document(open_path, readonly=readonly)

        try:
            self._check_paragraphs(doc, report)
            self._check_margins(doc, report)
            self._check_page_numbers(doc, report)

            if self.mark_document:
                doc.Save()
        finally:
            doc.Close(False)

        return report, checked_path


    def _mark_range(self, doc, start_par, end_par, message):
        if not self.mark_document:
            return

        try:
            start = doc.Paragraphs(start_par).Range.Start
            end = doc.Paragraphs(end_par).Range.End
            rng = doc.Range(start, end)

            rng.Font.Color = WD_COLOR_RED
            doc.Comments.Add(rng, message)
        except Exception:
            pass

    def _mark_page(self, doc, page, message):
        if not self.mark_document:
            return

        try:
            start = doc.GoTo(1, 1, page).Start
            try:
                end = doc.GoTo(1, 1, page + 1).Start
            except Exception:
                end = doc.Content.End

            rng = doc.Range(start, end)
            rng.Font.Color = WD_COLOR_RED
            doc.Comments.Add(rng, message)
        except Exception:
            pass

    def _group_pages(self, doc, report, pages, message, rule="3", priority=1):
        pages = sorted(set(pages))
        if not pages:
            return

        start = prev = pages[0]

        for p in pages[1:] + [None]:
            if p != prev + 1:
                loc = f"Стр. {start} - Стр. {prev}"

                report.issues.append(
                    CheckIssue(
                        rule=rule,
                        severity="ERROR",
                        location=loc,
                        message=message,
                        priority=priority,
                    )
                )

                for pg in range(start, prev + 1):
                    self._mark_page(doc, pg, message)

                if p is not None:
                    start = p
            prev = p

    def _check_paragraphs(self, doc, report):
        font_pages = []
        size_pages = []
        indent_pages = []
        spacing_pages = []

        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            text = (p.Range.Text or "").strip()

            if not text or len(text) < 3:
                continue

            try:
                if p.Range.Information(WD_WITHIN_TABLE):
                    continue
            except Exception:
                pass

            try:
                style_name = str(p.Range.Style.NameLocal).lower()
                if "heading" in style_name or "заголов" in style_name:
                    continue
            except Exception:
                pass

            page = p.Range.Information(WD_ACTIVE_END_PAGE_NUMBER)

            indent = p.Range.ParagraphFormat.FirstLineIndent
            if abs(indent) < 0.01:
                try:
                    indent = p.Range.Style.ParagraphFormat.FirstLineIndent
                except Exception:
                    pass

            indent_cm = pt_to_cm(indent)
            if abs(indent_cm - 1.25) > 0.05:
                indent_pages.append(page)

            font = str(p.Range.Font.Name).strip()
            if font != "Times New Roman":
                font_pages.append(page)

            try:
                size = float(p.Range.Font.Size)
            except Exception:
                size = 16.0

            if size == 9999999.0:
                real_bad_size_found = False

                try:
                    words = p.Range.Words
                    for j in range(1, words.Count + 1):
                        w = words(j)
                        w_text = (w.Text or "").strip()

                        if not w_text:
                            continue

                        if w_text in {".", ",", ";", ":", "-", "–", "—", "(", ")", "[", "]", "{", "}", "/", "\\"}:
                            continue

                        try:
                            w_size = float(w.Font.Size)
                        except Exception:
                            continue

                        if w_size == 9999999.0:
                            continue

                        if abs(w_size - 16) > 0.2:
                            real_bad_size_found = True
                            break

                    if real_bad_size_found:
                        size_pages.append(page)

                except Exception:
                    pass
            else:
                if abs(size - 16) > 0.2:
                    size_pages.append(page)

            spacing = int(p.Range.ParagraphFormat.LineSpacingRule)
            if spacing != WD_LINE_SPACE_SINGLE:
                spacing_pages.append(page)

        self._group_pages(
            doc,
            report,
            font_pages,
            (
                "В основном тексте используется неверный шрифт. "
                "Необходимо установить шрифт Times New Roman для всего основного текста документа."
            ),
            rule="3",
            priority=1,
        )

        self._group_pages(
            doc,
            report,
            size_pages,
            (
                "В основном тексте используется неверный размер шрифта. "
                "Необходимо установить размер шрифта 16 pt для всего основного текста документа."
            ),
            rule="3",
            priority=1,
        )

        self._group_pages(
            doc,
            report,
            indent_pages,
            (
                "Обнаружен неверный абзацный отступ первой строки. "
                "Необходимо установить отступ первой строки 1,25 см."
            ),
            rule="3",
            priority=1,
        )

        self._group_pages(
            doc,
            report,
            spacing_pages,
            (
                "Обнаружен неверный межстрочный интервал. "
                "Необходимо установить одинарный межстрочный интервал для основного текста документа."
            ),
            rule="3",
            priority=1,
        )

    def _check_margins(self, doc, report):
        pages = []

        for i in range(1, doc.Sections.Count + 1):
            sec = doc.Sections(i)
            ps = sec.PageSetup

            if (
                abs(pt_to_cm(ps.TopMargin) - 3.0) > 0.05 or
                abs(pt_to_cm(ps.BottomMargin) - 2.25) > 0.05 or
                abs(pt_to_cm(ps.LeftMargin) - 2.2) > 0.05 or
                abs(pt_to_cm(ps.RightMargin) - 2.3) > 0.05
            ):
                pages.append(i)

        self._group_pages(
            doc,
            report,
            pages,
            (
                "Обнаружены неверные размеры полей страницы. "
                "Необходимо установить поля документа: верхнее — 3 см, "
                "нижнее — 2,25 см, левое — 2,2 см, правое — 2,3 см."
            ),
            rule="3",
            priority=1,
        )

    def _check_page_numbers(self, doc, report):
        total = doc.ComputeStatistics(WD_STATISTIC_PAGES)

        found_top = False
        found_bottom = False
        issues = []

        for i in range(1, doc.Sections.Count + 1):
            sec = doc.Sections(i)

            header = sec.Headers(WD_HEADER_FOOTER_PRIMARY)
            footer = sec.Footers(WD_HEADER_FOOTER_PRIMARY)

            for f in header.Range.Fields:
                if f.Type == WD_FIELD_PAGE:
                    found_top = True
                    try:
                        para = f.Result.Paragraphs(1)
                        if int(para.Alignment) != WD_ALIGN_CENTER:
                            issues.append(
                                "Номер страницы расположен в верхнем колонтитуле, но выровнен не по центру. "
                                "Необходимо выровнять номер страницы строго по центру."
                            )
                    except Exception:
                        issues.append(
                            "Не удалось определить выравнивание номера страницы. "
                            "Проверьте, что номер страницы расположен сверху и строго по центру."
                        )

            if footer.PageNumbers.Count > 0:
                found_bottom = True

        loc = f"Стр. 1 - Стр. {total}"

        if not found_top and not found_bottom:
            message = (
                "Нумерация страниц отсутствует. "
                "Необходимо добавить номер страницы в верхней части страницы и выровнять его по центру."
            )
            report.issues.append(CheckIssue("4", "ERROR", loc, message, 1))
            self._mark_page(doc, 1, message)
            return

        if not found_top and found_bottom:
            message = (
                "Номер страницы расположен внизу страницы. "
                "По требованиям номер страницы должен находиться вверху страницы и быть выровнен по центру."
            )
            report.issues.append(CheckIssue("4", "ERROR", loc, message, 1))
            self._mark_page(doc, 1, message)
            return

        if issues:
            message = " ".join(issues)
            report.issues.append(CheckIssue("4", "ERROR", loc, message, 1))
            self._mark_page(doc, 1, message)