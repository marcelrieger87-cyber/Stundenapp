from __future__ import annotations
import sys
import os
from dataclasses import dataclass
from datetime import date, timedelta

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QGridLayout, QMessageBox, QDialog, QScrollArea, QFrame, QLineEdit
)

from excel_io import ExcelIO, H1, H2


EXCEL_PATH = r"P:\10245_Dateiupload-Qliksense-VGSG-K5\03_Service GW und int Produktbetreuung\01_MOIA Technik\Stundennachweis DLV 2026 1.0.xlsm"
# Tipp: Falls P: mal nicht gemappt ist, wäre ein UNC-Pfad (\\server\share\...) zuverlässiger.


@dataclass
class State:
    emp: str = ""
    mode: str = ""      # "PROJ" / "ABS"
    proj: str = ""
    hrs: float | None = None
    abs_type: str = ""
    d_from: date | None = None
    d_to: date | None = None
    month: date = date.today().replace(day=1)


class RestDialog(QDialog):
    def __init__(self, parent, projects: list[str], exclude: str):
        super().__init__(parent)
        self.setWindowTitle("Reststunden buchen")
        self.setModal(True)
        self.pick: str | None = None

        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(f"Rest {H1}h pro Tag auf welches Projekt buchen?"))

        grid = QGridLayout()
        r = c = 0
        for p in projects:
            if p.strip().lower() == exclude.strip().lower():
                continue
            btn = QPushButton(p)
            btn.clicked.connect(lambda _, x=p: self._select(x))
            grid.addWidget(btn, r, c)
            c += 1
            if c >= 3:
                c = 0
                r += 1
            if r >= 4:
                break

        box = QWidget()
        box.setLayout(grid)
        lay.addWidget(box)

        row = QHBoxLayout()
        skip = QPushButton("Überspringen")
        skip.clicked.connect(self.reject)
        row.addStretch(1)
        row.addWidget(skip)
        lay.addLayout(row)

    def _select(self, p: str):
        self.pick = p
        self.accept()


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("StundenApp (Desktop)")

        self.io = ExcelIO(EXCEL_PATH)
        self.state = State()

        # Daten aus Excel
        try:
            self.emps, self.projs, self.abss = self.io.load_lists()
        except Exception as e:
            QMessageBox.critical(self, "Fehler", str(e))
            self.emps, self.projs, self.abss = ["Muster"], ["MOIA", "DiE"], ["Urlaub", "Krank"]

        self._build_ui()
        self._render_info()
        self._render_calendar()

    # ---------------- UI ----------------
    def _build_ui(self):
        root = QVBoxLayout(self)

        self.info = QLabel("")
        self.info.setFrameShape(QFrame.StyledPanel)
        self.info.setStyleSheet("padding:8px;")
        root.addWidget(self.info)

        main = QHBoxLayout()
        root.addLayout(main)

        # Left panel
        left = QVBoxLayout()
        main.addLayout(left, 1)

        left.addWidget(QLabel("Mitarbeiter"))
        left.addWidget(self._tile_area(self.emps, self._pick_emp))

        row = QHBoxLayout()
        btn_p = QPushButton("Projekt")
        btn_a = QPushButton("Abwesenheit")
        btn_p.clicked.connect(lambda: self._set_mode("PROJ"))
        btn_a.clicked.connect(lambda: self._set_mode("ABS"))
        row.addWidget(btn_p)
        row.addWidget(btn_a)
        left.addLayout(row)

        left.addWidget(QLabel("Projekte"))
        left.addWidget(self._tile_area(self.projs, self._pick_proj))

        row2 = QHBoxLayout()
        b1 = QPushButton(str(H1).replace(".", ","))
        b2 = QPushButton(str(H2).replace(".", ","))
        b1.clicked.connect(lambda: self._pick_hours(H1))
        b2.clicked.connect(lambda: self._pick_hours(H2))
        row2.addWidget(QLabel("Stunden"))
        row2.addWidget(b1)
        row2.addWidget(b2)
        left.addLayout(row2)

        left.addWidget(QLabel("Abwesenheit"))
        left.addWidget(self._tile_area(self.abss, self._pick_abs))

        # Right panel (calendar + actions)
        right = QVBoxLayout()
        main.addLayout(right, 1)

        nav = QHBoxLayout()
        prev = QPushButton("◀")
        nxt = QPushButton("▶")
        self.month_label = QLabel("")
        self.month_label.setAlignment(Qt.AlignCenter)
        prev.clicked.connect(self._prev_month)
        nxt.clicked.connect(self._next_month)
        nav.addWidget(prev)
        nav.addWidget(self.month_label, 1)
        nav.addWidget(nxt)
        right.addLayout(nav)

        self.cal_grid = QGridLayout()
        right.addLayout(self.cal_grid)

        actions = QHBoxLayout()
        save = QPushButton("SPEICHERN ✅")
        reset = QPushButton("AUSWAHL LÖSCHEN")
        save.clicked.connect(self._save)
        reset.clicked.connect(self._reset)
        actions.addWidget(save, 1)
        actions.addWidget(reset, 1)
        right.addLayout(actions)

    def _tile_area(self, items: list[str], handler):
        cont = QWidget()
        grid = QGridLayout(cont)
        grid.setSpacing(6)

        r = c = 0
        for t in items[:60]:
            btn = QPushButton(t)
            btn.setMinimumHeight(28)
            btn.clicked.connect(lambda _, x=t: handler(x))
            grid.addWidget(btn, r, c)
            c += 1
            if c >= 3:
                c = 0
                r += 1

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(cont)
        scroll.setMinimumHeight(140)
        return scroll

    # ---------------- State / Render ----------------
    def _render_info(self):
        s = self.state
        def fmt(d): return d.strftime("%d.%m.%Y") if d else "—"
        d1 = s.d_from
        d2 = s.d_to or s.d_from
        if d1 and d2 and d2 < d1:
            d1, d2 = d2, d1

        if s.mode == "PROJ":
            act = f"Projekt: {s.proj or '—'} | Stunden: {str(s.hrs).replace('.', ',') if s.hrs else '—'}"
        elif s.mode == "ABS":
            act = f"Abwesenheit: {s.abs_type or '—'}"
        else:
            act = "Tätigkeit: —"

        self.info.setText(
            f"Mitarbeiter: {s.emp or '—'}   |   Zeitraum: {fmt(d1)} bis {fmt(d2)}\n{act}"
        )

    def _render_calendar(self):
        # clear grid
        while self.cal_grid.count():
            item = self.cal_grid.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        s = self.state
        m = s.month
        self.month_label.setText(m.strftime("%B %Y"))

        dow = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
        for i, d in enumerate(dow):
            lab = QLabel(d)
            lab.setAlignment(Qt.AlignCenter)
            self.cal_grid.addWidget(lab, 0, i)

        first = m
        # monday=0.. sunday=6
        start_offset = (first.weekday())  # Monday=0
        # days in month
        nxt_month = (m.replace(day=28) + timedelta(days=4)).replace(day=1)
        last_day = (nxt_month - timedelta(days=1)).day

        # selection normalize
        d_from = s.d_from
        d_to = s.d_to or s.d_from
        if d_from and d_to and d_to < d_from:
            d_from, d_to = d_to, d_from

        day = 1
        row = 1
        col = start_offset
        while day <= last_day:
            btn = QPushButton(str(day))
            btn.setMinimumHeight(32)

            d = m.replace(day=day)
            is_weekend = d.weekday() >= 5
            if is_weekend:
                btn.setStyleSheet("opacity:0.8;")

            if d_from and d_to and d_from <= d <= d_to:
                btn.setStyleSheet("background:#32a852; color:white;")

            btn.clicked.connect(lambda _, x=day: self._click_day(x))
            self.cal_grid.addWidget(btn, row, col)

            day += 1
            col += 1
            if col >= 7:
                col = 0
                row += 1

    # ---------------- Handlers ----------------
    def _pick_emp(self, x: str):
        self.state.emp = x
        self._render_info()

    def _set_mode(self, m: str):
        self.state.mode = m
        if m == "PROJ":
            self.state.abs_type = ""
        else:
            self.state.proj = ""
            self.state.hrs = None
        self._render_info()

    def _pick_proj(self, x: str):
        self.state.mode = "PROJ"
        self.state.proj = x
        self.state.abs_type = ""
        self._render_info()

    def _pick_hours(self, h: float):
        if self.state.mode != "PROJ" or not self.state.proj:
            return
        self.state.hrs = h
        self._render_info()

    def _pick_abs(self, x: str):
        if self.state.mode != "ABS":
            return
        self.state.abs_type = x
        self._render_info()

    def _click_day(self, day: int):
        s = self.state
        clicked = s.month.replace(day=day)

        if s.d_from is None:
            s.d_from = clicked
            s.d_to = None
        elif s.d_to is None:
            s.d_to = clicked
        else:
            s.d_from = clicked
            s.d_to = None

        self._render_info()
        self._render_calendar()

    def _prev_month(self):
        m = self.state.month
        prev = (m.replace(day=1) - timedelta(days=1)).replace(day=1)
        self.state.month = prev
        self.state.d_from = None
        self.state.d_to = None
        self._render_info()
        self._render_calendar()

    def _next_month(self):
        m = self.state.month
        nxt = (m.replace(day=28) + timedelta(days=4)).replace(day=1)
        self.state.month = nxt
        self.state.d_from = None
        self.state.d_to = None
        self._render_info()
        self._render_calendar()

    def _reset(self):
        s = self.state
        s.emp = ""
        s.mode = ""
        s.proj = ""
        s.hrs = None
        s.abs_type = ""
        s.d_from = None
        s.d_to = None
        self._render_info()
        self._render_calendar()

    def _save(self):
        s = self.state

        if not s.emp:
            QMessageBox.warning(self, "Hinweis", "Bitte Mitarbeiter anklicken.")
            return
        if not s.d_from:
            QMessageBox.warning(self, "Hinweis", "Bitte Datum im Kalender anklicken.")
            return
        if s.mode not in ("PROJ", "ABS"):
            QMessageBox.warning(self, "Hinweis", "Bitte Tätigkeit wählen (Projekt oder Abwesenheit).")
            return

        d1 = s.d_from
        d2 = s.d_to or s.d_from
        if d2 < d1:
            d1, d2 = d2, d1

        if s.mode == "PROJ":
            if not s.proj:
                QMessageBox.warning(self, "Hinweis", "Bitte Projekt anklicken.")
                return
            if s.hrs is None:
                QMessageBox.warning(self, "Hinweis", f"Bitte Stunden anklicken ({H1} / {H2}).")
                return
            if abs(s.hrs - H1) > 1e-9 and abs(s.hrs - H2) > 1e-9:
                QMessageBox.warning(self, "Hinweis", f"Nur {H1} oder {H2} Stunden erlaubt.")
                return
        else:
            if not s.abs_type:
                QMessageBox.warning(self, "Hinweis", "Bitte Abwesenheitsart anklicken.")
                return

        # Schreiben
        try:
            ok, fail = self.io.write_range(
                emp=s.emp,
                mode=s.mode,
                proj=s.proj,
                hrs=float(s.hrs or 0.0),
                abs_type=s.abs_type,
                d_from=d1,
                d_to=d2,
            )
        except Exception as e:
            QMessageBox.critical(self, "Fehler", str(e))
            return

        if ok == 0:
            QMessageBox.warning(
                self,
                "Nichts eingetragen",
                "Es wurde nichts eingetragen.\nUrsache meist: Mitarbeiter/Projekt im Monatsblatt nicht gefunden.\n"
                "Prüfe: Name in Zeile 3 (Kopf) und Projekt in Zeile 4 (Subheader).",
            )
            return

        # Restlogik (3,5h)
        if s.mode == "PROJ" and abs(float(s.hrs) - H1) < 1e-9:
            if d1 == d2:
                msg = f"Du hast {H1}h eingetragen.\nSoll der Rest {H1}h auf ein anderes Projekt gebucht werden?"
            else:
                msg = (
                    f"Du hast {H1}h für den Zeitraum eingetragen: {d1.strftime('%d.%m.%Y')} bis {d2.strftime('%d.%m.%Y')}\n"
                    f"Soll der Rest {H1}h PRO TAG auf ein anderes Projekt gebucht werden?"
                )
            if QMessageBox.question(self, "Reststunden buchen?", msg, QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                dlg = RestDialog(self, self.projs, exclude=s.proj)
                if dlg.exec() == QDialog.Accepted and dlg.pick:
                    try:
                        ok2, fail2 = self.io.write_range(
                            emp=s.emp,
                            mode="PROJ",
                            proj=dlg.pick,
                            hrs=H1,
                            abs_type="",
                            d_from=d1,
                            d_to=d2,
                        )
                        QMessageBox.information(
                            self, "Rest gebucht",
                            f"Rest gebucht: {ok2} Tag(e) auf {dlg.pick}" + (f"\nFehlgeschlagen: {fail2}" if fail2 else "")
                        )
                    except Exception as e:
                        QMessageBox.critical(self, "Fehler", str(e))
                        return

                self._reset()
                return

        QMessageBox.information(self, "Gespeichert", f"Gespeichert: {ok} Tag(e)." + (f"\nFehlgeschlagen: {fail}" if fail else ""))
        self._reset()


def main():
    app = QApplication(sys.argv)
    w = App()
    w.resize(980, 620)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
