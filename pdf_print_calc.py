# -*- coding: utf-8 -*-
"""حاسبة طباعة PDF - تحسب عدد الأوراق وتكلفة الطباعة لكل ملف (واجهة بأسلوب Windows 10)."""
import sys
import csv
import html
import json
import math
from datetime import datetime
from pathlib import Path

from pypdf import PdfReader
from PySide6.QtCore import Qt, QMarginsF, QUrl
from PySide6.QtGui import (QFont, QColor, QTextDocument, QPdfWriter, QPageLayout,
                           QPageSize, QDesktopServices)
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame,
    QPushButton, QDoubleSpinBox, QTableWidget, QTableWidgetItem, QFileDialog,
    QMessageBox, QHeaderView, QDialog, QTextEdit, QAbstractItemView,
)

SETTINGS = Path.home() / ".pdf_print_calc.json"

# (الاسم الكامل، الاسم المختصر، عدد الصفحات في الورقة وجه وظهر، لون التمييز، لون الخلفية الفاتحة)
TYPES = [
    ("A4 - صفحتين في الوجه (وجه وظهر)", "A4 ‹ 2 في الوجه", 4, "#0078D7", "#EAF3FC"),
    ("B5 - وجه وظهر", "B5 ‹ وجه وظهر", 2, "#107C10", "#E9F5E9"),
    ("A4 - وجه وظهر", "A4 ‹ وجه وظهر", 2, "#D83B01", "#FDEEE8"),
]
NEUTRAL = "#605E5C"

QSS = """
* { font-family: 'Segoe UI', 'Tahoma', sans-serif; font-size: 10pt; color: #1F1F1F; }
QWidget#root, QDialog { background: #F3F3F3; }
QFrame#titlebar { background: #0078D7; }
QFrame#titlebar QLabel { color: white; background: transparent; }
QLabel#section { font-weight: 600; font-size: 11pt; color: #323130; background: transparent; }
QLabel#hint { color: #8A8886; background: transparent; }

QPushButton {
    background: #E1E1E1; border: 1px solid #E1E1E1; padding: 7px 18px; min-height: 20px;
}
QPushButton:hover { background: #E5F1FB; border-color: #0078D7; }
QPushButton:pressed { background: #CCE4F7; }
QPushButton#primary { background: #0078D7; border-color: #0078D7; color: white; font-weight: 600; }
QPushButton#primary:hover { background: #106EBE; border-color: #106EBE; }
QPushButton#primary:pressed { background: #005A9E; }
QPushButton#danger:hover { background: #FDE7E9; border-color: #C42B1C; color: #C42B1C; }

QDoubleSpinBox {
    background: white; border: 1px solid #8A8886; padding: 5px 8px; min-height: 22px;
    font-size: 12pt; font-weight: 600; selection-background-color: #0078D7;
}
QDoubleSpinBox:hover { border-color: #323130; }
QDoubleSpinBox:focus { border: 2px solid #0078D7; }

QTableWidget {
    background: white; border: 1px solid #E1E1E1; gridline-color: #EDEBE9;
    alternate-background-color: #FAFAFA; outline: 0;
}
QTableWidget::item { padding: 4px 8px; }
QTableWidget::item:selected { background: #CCE4F7; color: #000000; }

QTextEdit { background: white; border: 1px solid #E1E1E1; padding: 8px; }

QScrollBar:vertical { background: transparent; width: 12px; margin: 0; }
QScrollBar::handle:vertical { background: #C1C1C1; border-radius: 6px; min-height: 30px; margin: 2px; }
QScrollBar::handle:vertical:hover { background: #8A8886; }
QScrollBar:horizontal { background: transparent; height: 12px; margin: 0; }
QScrollBar::handle:horizontal { background: #C1C1C1; border-radius: 6px; min-width: 30px; margin: 2px; }
QScrollBar::handle:horizontal:hover { background: #8A8886; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; }
QScrollBar::add-page, QScrollBar::sub-page { background: transparent; }
"""


def count_pages(path):
    reader = PdfReader(str(path))
    if reader.is_encrypted:
        reader.decrypt("")
    return len(reader.pages)


def fmt(x):
    return f"{x:,.2f}"


class ColorHeader(QHeaderView):
    """رأس جدول بألوان مخصصة لكل عمود."""

    def __init__(self, parent=None):
        super().__init__(Qt.Horizontal, parent)
        self.colors = {}
        self.setHighlightSections(False)
        self.setSectionsClickable(False)

    def paintSection(self, painter, rect, logicalIndex):
        painter.save()
        painter.fillRect(rect, QColor(self.colors.get(logicalIndex, NEUTRAL)))
        painter.setPen(QColor("white"))
        f = QFont("Segoe UI", 9)
        f.setBold(True)
        painter.setFont(f)
        text = self.model().headerData(logicalIndex, Qt.Horizontal, Qt.DisplayRole)
        painter.drawText(rect.adjusted(6, 0, -6, 0), Qt.AlignCenter | Qt.TextWordWrap, str(text))
        painter.setPen(QColor(255, 255, 255, 90))
        painter.drawLine(rect.topLeft(), rect.bottomLeft())
        painter.restore()


def make_card(accent):
    card = QFrame()
    card.setObjectName("card")
    card.setStyleSheet(
        f"QFrame#card {{ background: white; border: 1px solid #E1E1E1; border-top: 4px solid {accent}; }}"
        "QFrame#card QLabel { background: transparent; border: none; }"
    )
    lay = QVBoxLayout(card)
    lay.setContentsMargins(14, 10, 14, 12)
    lay.setSpacing(4)
    return card, lay


class ReportDialog(QDialog):
    def __init__(self, text, rows, html_doc, parent=None):
        super().__init__(parent)
        self.setWindowTitle("قائمة الأسعار")
        self.html_doc = html_doc
        self.resize(680, 600)
        self.rows = rows
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 16)
        self.box = QTextEdit()
        self.box.setReadOnly(True)
        self.box.setFont(QFont("Segoe UI", 11))
        self.box.setPlainText(text)
        lay.addWidget(self.box)
        btns = QHBoxLayout()
        b1 = QPushButton("نسخ")
        b0 = QPushButton("حفظ PDF")
        b0.setObjectName("primary")
        b0.clicked.connect(self.save_pdf)
        b2 = QPushButton("حفظ Excel (CSV)")
        b3 = QPushButton("إغلاق")
        b1.clicked.connect(self.copy)
        b2.clicked.connect(self.save_csv)
        b3.clicked.connect(self.accept)
        btns.addWidget(b0)
        btns.addWidget(b2)
        btns.addWidget(b1)
        btns.addStretch()
        btns.addWidget(b3)
        lay.addLayout(btns)

    def save_pdf(self):
        name = f"قائمة_الأسعار_{datetime.now():%Y-%m-%d}.pdf"
        p, _ = QFileDialog.getSaveFileName(self, "حفظ قائمة الأسعار PDF", name, "PDF (*.pdf)")
        if not p:
            return
        try:
            writer = QPdfWriter(p)
            writer.setPageLayout(QPageLayout(QPageSize(QPageSize.A4), QPageLayout.Landscape,
                                             QMarginsF(12, 12, 12, 12), QPageLayout.Millimeter))
            doc = QTextDocument()
            doc.setDefaultFont(QFont("Segoe UI", 10))
            doc.setHtml(self.html_doc)
            doc.print_(writer)
        except Exception as e:
            QMessageBox.warning(self, "خطأ", f"تعذر حفظ الملف:\n{e}")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(p))

    def copy(self):
        QApplication.clipboard().setText(self.box.toPlainText())

    def save_csv(self):
        p, _ = QFileDialog.getSaveFileName(self, "حفظ", "الاسعار.csv", "CSV (*.csv)")
        if not p:
            return
        with open(p, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows(self.rows)


class Main(QWidget):
    def __init__(self):
        super().__init__()
        self.setObjectName("root")
        self.setWindowTitle("حاسبة طباعة ملفات PDF")
        self.resize(1200, 780)
        self.setAcceptDrops(True)
        self.files = []  # [(path, pages)]

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # ---- الشريط العلوي ----
        bar = QFrame()
        bar.setObjectName("titlebar")
        bl = QVBoxLayout(bar)
        bl.setContentsMargins(24, 14, 24, 14)
        bl.setSpacing(0)
        t = QLabel("حاسبة طباعة ملفات PDF")
        t.setStyleSheet("font-size: 18pt; font-weight: 600; color: white;")
        s = QLabel("احسب عدد الأوراق وتكلفة الطباعة لكل ملف")
        s.setStyleSheet("font-size: 10pt; color: #DCEBFA;")
        bl.addWidget(t)
        bl.addWidget(s)
        outer.addWidget(bar)

        body = QVBoxLayout()
        body.setContentsMargins(20, 16, 20, 16)
        body.setSpacing(12)
        outer.addLayout(body)

        # ---- بطاقات الأسعار ----
        lbl = QLabel("سعر الورقة الواحدة لكل نوع طباعة")
        lbl.setObjectName("section")
        body.addWidget(lbl)
        prow = QHBoxLayout()
        prow.setSpacing(12)
        self.spins = []
        saved = self.load_prices()
        for i, (name, _, per, accent, _) in enumerate(TYPES):
            card, cl = make_card(accent)
            n = QLabel(name)
            n.setStyleSheet(f"color: {accent}; font-weight: 600;")
            sp = QDoubleSpinBox()
            sp.setRange(0, 100000)
            sp.setDecimals(2)
            sp.setSingleStep(0.25)
            sp.setButtonSymbols(QDoubleSpinBox.NoButtons)
            sp.setValue(saved[i])
            sp.valueChanged.connect(self.on_price)
            self.spins.append(sp)
            cl.addWidget(n)
            cl.addWidget(sp)
            prow.addWidget(card)
        body.addLayout(prow)

        # ---- الأزرار ----
        brow = QHBoxLayout()
        brow.setSpacing(8)
        for text, fn, oid in [
            ("＋  إضافة ملفات PDF", self.pick_files, "primary"),
            ("إضافة مجلد", self.pick_folder, ""),
            ("حذف المحدد", self.remove_selected, "danger"),
            ("مسح الكل", self.clear_all, "danger"),
        ]:
            b = QPushButton(text)
            if oid:
                b.setObjectName(oid)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(fn)
            brow.addWidget(b)
        brow.addStretch()
        rb = QPushButton("عرض قائمة الأسعار")
        rb.setObjectName("primary")
        rb.setCursor(Qt.PointingHandCursor)
        rb.clicked.connect(self.show_report)
        brow.addWidget(rb)
        body.addLayout(brow)

        # ---- الجدول ----
        headers = ["اسم الملف", "عدد الصفحات"]
        for _, short, *_ in TYPES:
            headers += [short + "\nالأوراق", short + "\nالسعر"]
        self.table = QTableWidget(0, len(headers))
        hdr = ColorHeader(self.table)
        self.table.setHorizontalHeader(hdr)
        self.table.setHorizontalHeaderLabels(headers)
        for i, (_, _, _, accent, _) in enumerate(TYPES):
            hdr.colors[2 + 2 * i] = accent
            hdr.colors[3 + 2 * i] = accent
        hdr.setFixedHeight(52)
        hdr.setMinimumSectionSize(90)
        hdr.setSectionResizeMode(QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(38)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(False)
        self.table.setFrameShape(QFrame.NoFrame)
        body.addWidget(self.table, 1)

        self.hint = QLabel("اسحب وأفلت ملفات PDF أو مجلدات داخل النافذة لإضافتها.")
        self.hint.setObjectName("hint")
        body.addWidget(self.hint)

        # ---- بطاقات الإجمالي ----
        lbl2 = QLabel("الإجمالي")
        lbl2.setObjectName("section")
        body.addWidget(lbl2)
        trow = QHBoxLayout()
        trow.setSpacing(12)
        self.tot_labels = []
        for name, _, per, accent, _ in TYPES:
            card, cl = make_card(accent)
            n = QLabel(name)
            n.setStyleSheet(f"color: {accent}; font-weight: 600;")
            price = QLabel("0.00")
            price.setStyleSheet("font-size: 20pt; font-weight: 600;")
            sheets = QLabel("0 ورقة")
            sheets.setStyleSheet("color: #605E5C;")
            cl.addWidget(n)
            cl.addWidget(price)
            cl.addWidget(sheets)
            trow.addWidget(card)
            self.tot_labels.append((price, sheets))
        card, cl = make_card(NEUTRAL)
        n = QLabel("الملفات المضافة")
        n.setStyleSheet(f"color: {NEUTRAL}; font-weight: 600;")
        self.files_lbl = QLabel("0 ملف")
        self.files_lbl.setStyleSheet("font-size: 20pt; font-weight: 600;")
        self.pages_lbl = QLabel("0 صفحة")
        self.pages_lbl.setStyleSheet("color: #605E5C;")
        cl.addWidget(n)
        cl.addWidget(self.files_lbl)
        cl.addWidget(self.pages_lbl)
        trow.addWidget(card)
        body.addLayout(trow)

        self.refresh()

    # ---------- الإعدادات ----------
    def load_prices(self):
        try:
            d = json.loads(SETTINGS.read_text(encoding="utf-8"))
            return ([float(x) for x in d["prices"]] + [0.0] * 3)[:3]
        except Exception:
            return [0.0, 0.0, 0.0]

    def prices(self):
        return [s.value() for s in self.spins]

    def on_price(self):
        try:
            SETTINGS.write_text(json.dumps({"prices": self.prices()}), encoding="utf-8")
        except Exception:
            pass
        self.refresh()

    # ---------- الملفات ----------
    def pick_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "اختر ملفات PDF", "", "PDF (*.pdf)")
        self.add_files(paths)

    def pick_folder(self):
        d = QFileDialog.getExistingDirectory(self, "اختر مجلد")
        if d:
            self.add_files([str(p) for p in sorted(Path(d).rglob("*"))
                            if p.suffix.lower() == ".pdf"])

    def add_files(self, paths):
        existing = {p for p, _ in self.files}
        failed = []
        for p in paths:
            p = str(p)
            if p in existing or Path(p).suffix.lower() != ".pdf":
                continue
            try:
                self.files.append((p, count_pages(p)))
            except Exception:
                failed.append(Path(p).name)
        self.refresh()
        if failed:
            QMessageBox.warning(self, "تعذر قراءة بعض الملفات", "\n".join(failed))

    def remove_selected(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        for r in rows:
            if r < len(self.files):
                del self.files[r]
        self.refresh()

    def clear_all(self):
        self.files.clear()
        self.refresh()

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        paths = []
        for u in e.mimeData().urls():
            p = Path(u.toLocalFile())
            if p.is_dir():
                paths += [str(x) for x in sorted(p.rglob("*")) if x.suffix.lower() == ".pdf"]
            else:
                paths.append(str(p))
        self.add_files(paths)

    # ---------- الحساب ----------
    def compute(self):
        prices = self.prices()
        rows = []
        for path, pages in self.files:
            cells = []
            for (_, _, per, _, _), price in zip(TYPES, prices):
                sheets = math.ceil(pages / per)
                cells.append((sheets, sheets * price))
            rows.append((Path(path).name, pages, cells))
        tot_pages = sum(r[1] for r in rows)
        tot = [(sum(r[2][i][0] for r in rows), sum(r[2][i][1] for r in rows))
               for i in range(len(TYPES))]
        return rows, tot_pages, tot

    def refresh(self):
        rows, tot_pages, tot = self.compute()
        self.table.setRowCount(len(rows))

        def put(r, c, text, bg=None, color=None, bold=False):
            it = QTableWidgetItem(str(text))
            it.setTextAlignment(Qt.AlignCenter if c else (Qt.AlignVCenter | Qt.AlignRight))
            if bg:
                it.setBackground(QColor(bg))
            if color:
                it.setForeground(QColor(color))
            if bold:
                f = it.font()
                f.setBold(True)
                it.setFont(f)
            self.table.setItem(r, c, it)

        for r, (name, pages, cells) in enumerate(rows):
            put(r, 0, name)
            put(r, 1, pages, bold=True)
            for i, (s, p) in enumerate(cells):
                _, _, _, accent, tint = TYPES[i]
                put(r, 2 + 2 * i, s, bg=tint)
                put(r, 3 + 2 * i, fmt(p), bg=tint, color=accent, bold=True)

        for (price, sheets), (s, p) in zip(self.tot_labels, tot):
            price.setText(fmt(p))
            sheets.setText(f"{s} ورقة")
        self.files_lbl.setText(f"{len(rows)} ملف")
        self.pages_lbl.setText(f"{tot_pages} صفحة")
        self.hint.setVisible(not rows)

    def show_report(self):
        if not self.files:
            QMessageBox.information(self, "تنبيه", "أضف ملفات PDF أولاً.")
            return
        rows, tot_pages, tot = self.compute()
        prices = self.prices()
        line = "=" * 52
        out = [line, "قائمة الأسعار", line, ""]
        csv_rows = [["الملف", "الصفحات"] + [f"{t[0]} - أوراق" for t in TYPES]
                    + [f"{t[0]} - السعر" for t in TYPES]]
        for name, pages, cells in rows:
            out.append(f"الملف: {name}   (الصفحات: {pages})")
            for t, (s, p) in zip(TYPES, cells):
                out.append(f"   {t[0]}: {s} ورقة = {fmt(p)}")
            out.append("")
            csv_rows.append([name, pages] + [c[0] for c in cells] + [round(c[1], 2) for c in cells])
        out += [line, f"الإجمالي  (الملفات: {len(rows)} | الصفحات: {tot_pages})", line]
        for t, (s, p), pr in zip(TYPES, tot, prices):
            out.append(f"{t[0]}\n   {s} ورقة × {fmt(pr)} = {fmt(p)}")
        csv_rows.append(["الإجمالي", tot_pages] + [x[0] for x in tot] + [round(x[1], 2) for x in tot])
        ReportDialog("\n".join(out), csv_rows, self.build_html(rows, tot_pages, tot, prices), self).exec()

    def build_html(self, rows, tot_pages, tot, prices):
        e = html.escape
        th = lambda t, c: f'<td bgcolor="{c}" align="center"><font color="white"><b>{t}</b></font></td>'
        h = ['<html><body dir="rtl" style="font-family:Segoe UI;">',
             '<h2 style="color:#0078D7;">قائمة أسعار الطباعة</h2>',
             f'<p style="color:#605E5C;">التاريخ: {datetime.now():%Y-%m-%d %H:%M} &nbsp;|&nbsp; '
             f'عدد الملفات: {len(rows)} &nbsp;|&nbsp; إجمالي الصفحات: {tot_pages}</p>',
             '<table width="100%" border="1" cellspacing="0" cellpadding="6" '
             'style="border-collapse:collapse; border-color:#C8C6C4;">',
             '<tr>' + th("اسم الملف", NEUTRAL) + th("الصفحات", NEUTRAL)]
        for t in TYPES:
            h.append(th(e(t[1]) + " - أوراق", t[3]) + th(e(t[1]) + " - السعر", t[3]))
        h.append('</tr>')
        for name, pages, cells in rows:
            h.append(f'<tr><td>{e(name)}</td><td align="center">{pages}</td>')
            for t, (sh, pr) in zip(TYPES, cells):
                h.append(f'<td align="center" bgcolor="{t[4]}">{sh}</td>'
                         f'<td align="center" bgcolor="{t[4]}"><b>{fmt(pr)}</b></td>')
            h.append('</tr>')
        h.append('<tr bgcolor="#FFF4CE"><td><b>الإجمالي</b></td>'
                 f'<td align="center"><b>{tot_pages}</b></td>')
        for sh, pr in tot:
            h.append(f'<td align="center"><b>{sh}</b></td><td align="center"><b>{fmt(pr)}</b></td>')
        h.append('</tr></table>')

        h.append('<h3 style="color:#323130;">ملخص الإجمالي حسب نوع الطباعة</h3>')
        h.append('<table width="60%" border="1" cellspacing="0" cellpadding="8" '
                 'style="border-collapse:collapse; border-color:#C8C6C4;">')
        h.append('<tr>' + th("نوع الطباعة", NEUTRAL) + th("عدد الأوراق", NEUTRAL)
                 + th("سعر الورقة", NEUTRAL) + th("الإجمالي", NEUTRAL) + '</tr>')
        for t, (sh, pr), price in zip(TYPES, tot, prices):
            h.append(f'<tr bgcolor="{t[4]}"><td><font color="{t[3]}"><b>{e(t[0])}</b></font></td>'
                     f'<td align="center">{sh}</td><td align="center">{fmt(price)}</td>'
                     f'<td align="center"><b>{fmt(pr)}</b></td></tr>')
        h.append('</table></body></html>')
        return "".join(h)


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setLayoutDirection(Qt.RightToLeft)
    app.setFont(QFont("Segoe UI", 10))
    app.setStyleSheet(QSS)
    w = Main()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
