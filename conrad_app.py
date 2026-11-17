"""
conrad_app.py
-------------
Desktop GUI Application for Conrad.de Catalog Categorizer.
No terminal required. Double-click the .exe to open a visual window.

Features:
- Graphical User Interface built with Tkinter / ttk (native Windows styling).
- Threaded processing: UI remains smooth and responsive with a live progress bar.
- File selector: Browse your Excel file (.xlsx) with standard Windows dialog.
- Live results table showing Row #, Product Title, MPN/EAN, Status, and Category ID.
- Fallback logic: Automatically assigns "Manuell pflegen" if no high-confidence match.
- Strict Text formatting (@): Preserves leading zeroes for all 7-digit IDs (e.g. 0312026).
- Incremental auto-saving every 5 rows into '[filename]_categorized.xlsx'.
- Buttons to open the output file or containing folder directly in Windows Explorer.
"""

import os
import sys
import re
import json
import time
import threading
import urllib.request
import urllib.parse
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import openpyxl
    from openpyxl.styles.numbers import FORMAT_TEXT
except ImportError:
    openpyxl = None

MANUAL_FALLBACK_TEXT = "Manuell pflegen"

# Domain keyword matching rules for high-confidence fallback
KEYWORD_RULES = [
    (r"\b(wechselbit|bit-stecknuss|bit-set|bitset|bitsatz)\b", "1512020", "Bit-Sets"),
    (r"\b(schnell-ladegerät|ladepad|ladegerät)\b", "0502019", "Ladegeräte & Ladepads"),
    (r"\b(akku)\b.*\b(18v|li-ion|batterie)\b", "0501040", "Werkzeug-Akkus"),
    (r"\b(drehmomentschlüssel|reparatursatz.*drehmoment)\b", "0312024", "Drehmomentschlüssel"),
    (r"\b(druckluft-schlagschrauber|schlagschrauber)\b", "1517983", "Druckluft-Schlagschrauber"),
    (r"\b(steckschlüssel|ringschlüssel|schraubenschlüssel-satz)\b", "1502030", "Steckschlüssel & Schraubenschlüssel"),
    (r"\b(flachrundzange|kombizange|zangen?)\b", "1502010", "Zangen"),
    (r"\b(winkelstiftschlüssel|stiftschlüssel|torx-schlüssel)\b", "1502021", "Winkelstiftschlüssel"),
    (r"\b(auftragstasche|schloß|laufrollensatz|laufschienen|rollbrett|liege|magnethalter)\b", "1504060", "Werkstatteinrichtung"),
    (r"\b(reifenmontagepaste|montagepaste)\b", "0312068", "Reifenmontagepaste"),
    (r"\b(radmutternschlüssel|felgenschlüssel|rangierer|the minidevil)\b", "0312020", "Reifen & Radwerkzeuge"),
    (r"\b(injektoren|lecksuche|nebelmaschine|videoskop)\b", "0312070", "KFZ-Diagnose & Lecksuche"),
    (r"\b(motorbrücke|riemenscheiben|kupplungs.*de-und|zapfenschlüssel)\b", "0312026", "KFZ-Motorwerkzeuge"),
    (r"\b(bremsleitungs|prüflampe|bordnetz-spannungserhalter|sicherung.*500a)\b", "0812022", "Bremsen & KFZ-Elektrik"),
    (r"\b(ausdrücker.*antriebswellen|radnaben)\b", "1202204", "Achs- & Antriebswellenwerkzeuge"),
    (r"\b(schutzbrille)\b", "2201010", "Arbeitsschutzbrillen"),
    (r"\b(stufenbohrer)\b", "1506051", "Stufenbohrer"),
    (r"\b(maurerhammer|hammer)\b", "1516040", "Hämmer & Schlagwerkzeuge"),
    (r"\b(winkelgradmesser|winkelmesser)\b", "1510020", "Winkelmesser"),
    (r"\b(ersatzmesser|klingen)\b", "1517032", "Ersatzklingen"),
    (r"\b(cob-taschenlampe|taschenlampe)\b", "0312040", "Arbeitsleuchten & Taschenlampen"),
    (r"\b(feuerzeug)\b", "3300200", "Feuerzeuge"),
]

def conrad_search(query):
    if not query or not query.strip():
        return []
    clean_q = query.strip()
    url = f"https://www.conrad.de/de/search.html?search={urllib.parse.quote(clean_q)}"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'de-DE,de;q=0.9,en;q=0.8'
    }
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=8) as response:
            html = response.read().decode('utf-8', errors='ignore')
    except Exception:
        return []

    results = []
    scripts = re.findall(r'<script[^>]*>(.*?)</script>', html, re.DOTALL)
    for s in scripts:
        if '"articleGroup"' in s:
            try:
                data = json.loads(s)
                if isinstance(data, list):
                    for item in data:
                        if isinstance(item, dict) and "articleGroup" in item:
                            ag_idx = item.get("articleGroup")
                            ag_name_idx = item.get("articleGroupName")
                            category_id = data[ag_idx] if isinstance(ag_idx, int) and ag_idx < len(data) else str(ag_idx)
                            category_name = data[ag_name_idx] if isinstance(ag_name_idx, int) and ag_name_idx < len(data) else str(ag_name_idx or "")
                            if isinstance(category_id, (str, int)) and str(category_id).isdigit():
                                cat_id_str = str(category_id).zfill(7)
                                results.append({
                                    "category_id": cat_id_str,
                                    "category_name": str(category_name),
                                    "url": f"https://www.conrad.de/de/o/-{cat_id_str}.html"
                                })
            except Exception:
                pass
        cat_matches = re.findall(r'/de/o/([a-zA-Z0-9\-]+)-(\d{7})\.html', s)
        for slug, cid in cat_matches:
            results.append({
                "category_id": cid,
                "category_name": slug.replace('-', ' ').title(),
                "url": f"https://www.conrad.de/de/o/{slug}-{cid}.html"
            })
            
    seen = set()
    deduped = []
    for r in results:
        cid = r["category_id"]
        if cid not in seen:
            seen.add(cid)
            deduped.append(r)
    return deduped

def match_product(name, brand_ref, ean, cache):
    for key in [ean, brand_ref, name]:
        if key and key in cache:
            return cache[key], "Lokaler Cache", f"https://www.conrad.de/de/o/-{cache[key]}.html"

    if ean and len(str(ean).strip()) >= 8:
        cands = conrad_search(str(ean).strip())
        if cands:
            cid = cands[0]["category_id"].zfill(7)
            cache[ean] = cid
            return cid, cands[0]["category_name"], cands[0]["url"]

    if brand_ref and len(str(brand_ref).strip()) >= 3:
        cands = conrad_search(str(brand_ref).strip())
        if cands:
            cid = cands[0]["category_id"].zfill(7)
            cache[brand_ref] = cid
            return cid, cands[0]["category_name"], cands[0]["url"]

    if name:
        lower_name = name.lower()
        for pattern, cid, cname in KEYWORD_RULES:
            if re.search(pattern, lower_name):
                cid_str = cid.zfill(7)
                if ean: cache[ean] = cid_str
                if brand_ref: cache[brand_ref] = cid_str
                return cid_str, cname, f"https://www.conrad.de/de/o/-{cid_str}.html"

    if name:
        clean_title = re.sub(r'[,/()"\']', ' ', name)
        words = [w for w in clean_title.split() if len(w) > 2][:4]
        q_title = " ".join(words)
        cands = conrad_search(q_title)
        if cands:
            cid = cands[0]["category_id"].zfill(7)
            if ean: cache[ean] = cid
            return cid, cands[0]["category_name"], cands[0]["url"]

    return None, None, "https://www.conrad.de/de/search.html?search=" + urllib.parse.quote(name or "")

class ConradApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Conrad.de Katalog Kategorisierer")
        self.geometry("900x600")
        self.minsize(750, 480)

        # Style configuration
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.status_var = tk.StringVar(value="Bereit. Bitte wählen Sie eine Excel-Datei aus.")
        self.is_running = False
        self.cache = {}

        self.create_widgets()

    def create_widgets(self):
        # Top Header Frame
        header_frame = ttk.Frame(self, padding=15)
        header_frame.pack(fill=tk.X)

        title_label = ttk.Label(header_frame, text="Conrad.de Katalog Kategorisierer", font=("Segoe UI", 16, "bold"))
        title_label.pack(anchor="w")

        subtitle = ttk.Label(header_frame, text="Automatische Zuordnung von Kategorie-IDs (Parent / EFEU) für Onboarding-Listen", font=("Segoe UI", 9))
        subtitle.pack(anchor="w", pady=(2, 0))

        # File Selection Frame
        file_frame = ttk.LabelFrame(self, text=" Excel-Datei ", padding=10)
        file_frame.pack(fill=tk.X, padx=15, pady=5)

        entry = ttk.Entry(file_frame, textvariable=self.input_file_path, font=("Segoe UI", 9))
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        browse_btn = ttk.Button(file_frame, text="Datei wählen...", command=self.browse_file)
        browse_btn.pack(side=tk.LEFT)

        self.start_btn = ttk.Button(file_frame, text="▶ Starten", command=self.start_processing)
        self.start_btn.pack(side=tk.LEFT, padx=(10, 0))

        # Progress Frame
        prog_frame = ttk.Frame(self, padding=(15, 5))
        prog_frame.pack(fill=tk.X)

        self.progress_bar = ttk.Progressbar(prog_frame, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))

        status_lbl = ttk.Label(prog_frame, textvariable=self.status_var, font=("Segoe UI", 9, "italic"))
        status_lbl.pack(anchor="w")

        # Results Table Frame
        table_frame = ttk.LabelFrame(self, text=" Verarbeitete Produkte ", padding=10)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        columns = ("row", "name", "mpn", "status", "category_id", "category_name")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")

        self.tree.heading("row", text="Zeile")
        self.tree.heading("name", text="Produktbezeichnung")
        self.tree.heading("mpn", text="Hersteller-Nr.")
        self.tree.heading("status", text="Status")
        self.tree.heading("category_id", text="Kategorie ID (Parent)")
        self.tree.heading("category_name", text="Kategorie-Name")

        self.tree.column("row", width=60, anchor="center")
        self.tree.column("name", width=340)
        self.tree.column("mpn", width=110)
        self.tree.column("status", width=90, anchor="center")
        self.tree.column("category_id", width=120, anchor="center")
        self.tree.column("category_name", width=150)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Bottom Bar
        bottom_frame = ttk.Frame(self, padding=15)
        bottom_frame.pack(fill=tk.X)

        self.stats_lbl = ttk.Label(bottom_frame, text="Gesamt: 0 | Zugeordnet: 0 | Manuell pflegen: 0", font=("Segoe UI", 9, "bold"))
        self.stats_lbl.pack(side=tk.LEFT)

        self.open_file_btn = ttk.Button(bottom_frame, text="Ergebnis-Datei öffnen", command=self.open_output_file, state=tk.DISABLED)
        self.open_file_btn.pack(side=tk.RIGHT, padx=(10, 0))

        self.open_folder_btn = ttk.Button(bottom_frame, text="Ordner anzeigen", command=self.open_output_folder, state=tk.DISABLED)
        self.open_folder_btn.pack(side=tk.RIGHT)

    def browse_file(self):
        f = filedialog.askopenfilename(
            title="Excel-Datei auswählen",
            filetypes=[("Excel-Arbeitsmappen", "*.xlsx"), ("Alle Dateien", "*.*")]
        )
        if f:
            self.input_file_path.set(f)
            p = Path(f)
            out = str(p.with_name(f"{p.stem}_categorized{p.suffix}"))
            self.output_file_path.set(out)

    def start_processing(self):
        if not openpyxl:
            messagebox.showerror("Fehler", "openpyxl ist nicht installiert.")
            return

        in_file = self.input_file_path.get().strip()
        if not in_file or not os.path.exists(in_file):
            messagebox.showwarning("Hinweis", "Bitte wählen Sie eine gültige Excel-Datei (.xlsx) aus.")
            return

        self.is_running = True
        self.start_btn.configure(state=tk.DISABLED)
        self.open_file_btn.configure(state=tk.DISABLED)
        self.open_folder_btn.configure(state=tk.DISABLED)

        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        thread = threading.Thread(target=self.run_categorization, args=(in_file,), daemon=True)
        thread.start()

    def run_categorization(self, in_file):
        out_file = self.output_file_path.get()
        try:
            self.status_var.set("Öffne Excel-Datei...")
            wb = openpyxl.load_workbook(in_file)
            sheet_name = "Produktcheckliste" if "Produktcheckliste" in wb.sheetnames else wb.sheetnames[0]
            ws = wb[sheet_name]

            # Dynamic column discovery
            col_map = {}
            for col_idx in range(1, ws.max_column + 1):
                val = ws.cell(row=1, column=col_idx).value
                if val:
                    col_map[str(val).strip()] = col_idx

            parent_col = col_map.get("Parent", 4)
            name_col = col_map.get("Name", 37)
            ean_col = col_map.get("Product.EAN.Main", 40)
            mpn_col = col_map.get("BrandReference", col_map.get("ATT.GLOBAL.ManufacturerAID", col_map.get("Product.Supplier.ArticleID", 42)))

            start_row = 8
            max_row = ws.max_row
            valid_rows = []
            for r in range(start_row, max_row + 1):
                n = ws.cell(row=r, column=name_col).value
                if n and str(n).strip():
                    valid_rows.append(r)

            total = len(valid_rows)
            matched = 0
            manual = 0
            skipped = 0

            self.progress_bar.configure(maximum=total, value=0)

            for idx, r in enumerate(valid_rows, 1):
                name_val = str(ws.cell(row=r, column=name_col).value or "").strip()
                parent_cell = ws.cell(row=r, column=parent_col)
                current_parent = parent_cell.value
                mpn_val = str(ws.cell(row=r, column=mpn_col).value or "").strip()
                ean_val = str(ws.cell(row=r, column=ean_col).value or "").strip()

                if current_parent and str(current_parent).strip():
                    skipped += 1
                    status_text = "Übersprungen"
                    cat_id_text = str(current_parent).strip()
                    cat_name_text = "Vorhanden"
                else:
                    self.status_var.set(f"Verarbeite Zeile {r}/{max_row}: {name_val[:40]}...")
                    cat_id, cat_name, _ = match_product(name_val, mpn_val, ean_val, self.cache)
                    parent_cell.number_format = FORMAT_TEXT

                    if cat_id:
                        parent_cell.value = str(cat_id).zfill(7)
                        matched += 1
                        status_text = "Zugeordnet"
                        cat_id_text = parent_cell.value
                        cat_name_text = cat_name or ""
                    else:
                        parent_cell.value = MANUAL_FALLBACK_TEXT
                        manual += 1
                        status_text = "Manuell"
                        cat_id_text = MANUAL_FALLBACK_TEXT
                        cat_name_text = "Keine Zuordnung"

                # Update UI
                self.after(0, self.add_tree_row, r, name_val, mpn_val, status_text, cat_id_text, cat_name_text)
                self.after(0, self.update_progress, idx, total, matched + skipped, manual)

                # Incremental save every 5 rows
                if idx % 5 == 0:
                    wb.save(out_file)

            wb.save(out_file)
            self.after(0, self.finish_processing, total, matched, skipped, manual, out_file)

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Fehler bei der Verarbeitung", f"Ein Fehler ist aufgetreten:\n{str(e)}"))
            self.after(0, lambda: self.status_var.set("Fehler aufgetreten."))
            self.after(0, lambda: self.start_btn.configure(state=tk.NORMAL))

    def add_tree_row(self, r, name, mpn, status, cat_id, cat_name):
        self.tree.insert("", tk.END, values=(r, name, mpn, status, cat_id, cat_name))
        children = self.tree.get_children()
        if children:
            self.tree.see(children[-1])

    def update_progress(self, current, total, matched, manual):
        self.progress_bar.configure(value=current)
        self.stats_lbl.configure(text=f"Gesamt: {total} | Zugeordnet: {matched} | Manuell pflegen: {manual}")

    def finish_processing(self, total, matched, skipped, manual, out_file):
        self.status_var.set(f"Fertig! Gespeichert in: {os.path.basename(out_file)}")
        self.start_btn.configure(state=tk.NORMAL)
        self.open_file_btn.configure(state=tk.NORMAL)
        self.open_folder_btn.configure(state=tk.NORMAL)
        messagebox.showinfo("Abgeschlossen", f"Katalog-Verarbeitung beendet!\n\nGesamt: {total}\nNeu zugeordnet: {matched}\nBereits gefüllt: {skipped}\nManuell pflegen: {manual}\n\nDatei: {out_file}")

    def open_output_file(self):
        out = self.output_file_path.get()
        if os.path.exists(out):
            os.startfile(out)

    def open_output_folder(self):
        out = self.output_file_path.get()
        if os.path.exists(out):
            folder = os.path.dirname(out)
            os.startfile(folder)

if __name__ == "__main__":
    app = ConradApp()
    app.mainloop()
