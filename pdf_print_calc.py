# -*- coding: utf-8 -*-
"""حاسبة طباعة PDF - تحسب عدد الأوراق وتكلفة الطباعة لكل ملف."""
import sys
import csv
import json
import math
from pathlib import Path

from pypdf import PdfReader
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QColor
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
    QPushButton, QDoubleSpinBox, QTableWidget, QTableWidgetItem, QFileDialog,
    QMessageBox, QHeaderView, QGroupBox, QDialog, QTextEdit, QAbstractItemView,
)

SETTINGS = Path.home() / ".pdf_print_calc.json"

# (الاسم، عدد الصفحات في الورقة الواحدة وجه وظهر)
TYPES = [
    ("A4 - صفحتين في الوجه (وجه وظهر)", 4),
    ("B5 - وجه وظهر", 2),
    ("A4 - وجه وظهر", 2),
]


def count_pages(path):
    reader = PdfReader(str(path))
    if reader.is_encrypted:
        reader.decrypt("")
    return len(reader.pages)


def fmt(x):
    return f"{x:,.2f}"


class ReportDialog(QDialog):
    def __init__(self, text, rows, parent=None):
        super().__init__(parent)
        self.setWindowTitle("قائمة الأسعار")
        self.resize(640, 560)
        self.rows = rows
        lay = QVBoxLayout(self)
        self.box = QTextEdit()
        self.box.setReadOnly(True)
        self.box.setFont(QFont("Consolas", 11))
        self.box.setPlainText(text)
        lay.addWidget(self.box)
        btns = QHBoxLayout()
        b1 = QPushButton("نسخ")
        b2 = QPushButton("حفظ Excel (CSV)")
        b3 = QPushButton("إغلاق")
        b1.clicked.connect(self.copy)
        b2.clicked.connect(self.save_csv)
        b3.clicked.connect(self.accept)
        for b in (b1, b2, b3):
            btns.addWidget(b)
        lay.addLayout(btns)

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
        self.setWindowTitle("حاسبة طباعة ملفات PDF")
        self.resize(1150, 700)
        self.setAcceptDrops(True)
        self.files = []  # [(path, pages)]

        root = QVBoxLayout(self)

        # الأسعار
        gb = QGroupBox("سعر الورقة الواحدة لكل نوع طباعة")
        grid = QGridLayout(gb)
        self.spins = []
        saved = self.load_prices()
        for i, (name, _) in enumerate(TYPES):
            grid.addWidget(QLabel(name), 0, i)
            sp = QDoubleSpinBox()
            sp.setRange(0, 100000)
            sp.setDecimals(2)
            sp.setSingleStep(0.25)
            sp.setValue(saved[i])
            sp.valueChanged.connect(self.on_price)
            self.spins.append(sp)
            grid.addWidget(sp, 1, i)
        root.addWidget(gb)

        # الأزرار
        bar = QHBoxLayout()
        for text, fn in [
            ("إضافة ملفات PDF", self.pick_files),
            ("إضافة مجلد", self.pick_folder),
            ("حذف المحدد", self.remove_selected),
            ("مسح الكل", self.clear_all),
        ]:
            b = QPushButton(text)
            b.clicked.connect(fn)
            bar.addWidget(b)
        bar.addStretch()
        b = QPushButton("عرض قائمة الأسعار")
        b.setStyleSheet("font-weight:bold; padding:6px 14px;")
        b.clicked.connect(self.show_report)
        bar.addWidget(b)
        root.addLayout(bar)

        # الجدول
        headers = ["اسم الملف", "عدد الصفحات"]
        for name, _ in TYPES:
            headers += [name + "\nعدد الأوراق", name + "\nالسعر"]
        self.table = QTableWidget(0, len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(0, QHeaderView.Stretch)
        root.addWidget(self.table)

        hint = QLabel("يمكنك سحب وإفلات ملفات PDF أو مجلدات داخل النافذة.")
        hint.setStyleSheet("color:gray;")
        root.addWidget(hint)

        self.refresh()

    # ---------- الإعدادات ----------
    def load_prices(self):
        try:
            d = json.loads(SETTINGS.read_text(encoding="utf-8"))
            return [float(x) for x in d["prices"]][:3] + [0.0] * 3
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
            for (_, per), price in zip(TYPES, prices):
                sheets = math.ceil(pages / per)
                cells.append((sheets, sheets * price))
            rows.append((Path(path).name, pages, cells))
        tot_pages = sum(r[1] for r in rows)
        tot = [(sum(r[2][i][0] for r in rows), sum(r[2][i][1] for r in rows))
               for i in range(len(TYPES))]
        return rows, tot_pages, tot

    def refresh(self):
        rows, tot_pages, tot = self.compute()
        self.table.setRowCount(len(rows) + 1)

        def put(r, c, text, bold=False, bg=None):
            it = QTableWidgetItem(str(text))
            it.setTextAlignment(Qt.AlignCenter if c else Qt.AlignVCenter | Qt.AlignRight)
            if bold:
                f = it.font()
                f.setBold(True)
                it.setFont(f)
            if bg:
                it.setBackground(QColor(bg))
                it.setForeground(QColor("#000000"))
            self.table.setItem(r, c, it)

        for r, (name, pages, cells) in enumerate(rows):
            put(r, 0, name)
            put(r, 1, pages)
            for i, (s, p) in enumerate(cells):
                put(r, 2 + 2 * i, s)
                put(r, 3 + 2 * i, fmt(p))

        r = len(rows)
        put(r, 0, "الإجمالي", True, "#FFE9A8")
        put(r, 1, tot_pages, True, "#FFE9A8")
        for i, (s, p) in enumerate(tot):
            put(r, 2 + 2 * i, s, True, "#FFE9A8")
            put(r, 3 + 2 * i, fmt(p), True, "#FFE9A8")

    def show_report(self):
        if not self.files:
            QMessageBox.information(self, "تنبيه", "أضف ملفات PDF أولاً.")
            return
        rows, tot_pages, tot = self.compute()
        prices = self.prices()
        line = "=" * 56
        out = [line, "قائمة الأسعار", line, ""]
        csv_rows = [["الملف", "الصفحات"] + [f"{n} - أوراق" for n, _ in TYPES]
                    + [f"{n} - السعر" for n, _ in TYPES]]
        for name, pages, cells in rows:
            out.append(f"الملف: {name}   (الصفحات: {pages})")
            for (tname, _), (s, p) in zip(TYPES, cells):
                out.append(f"   {tname}: {s} ورقة = {fmt(p)}")
            out.append("")
            csv_rows.append([name, pages] + [c[0] for c in cells] + [round(c[1], 2) for c in cells])
        out += [line, f"الإجمالي  (عدد الملفات: {len(rows)} | الصفحات: {tot_pages})", line]
        for (tname, _), (s, p), pr in zip(TYPES, tot, prices):
            out.append(f"{tname}\n   {s} ورقة × {fmt(pr)} = {fmt(p)}")
        csv_rows.append(["الإجمالي", tot_pages] + [t[0] for t in tot] + [round(t[1], 2) for t in tot])
        ReportDialog("\n".join(out), csv_rows, self).exec()


def main():
    app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.RightToLeft)
    app.setFont(QFont("Segoe UI", 10))
    w = Main()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
