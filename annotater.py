#!/usr/bin/env python3
"""
new_annotater.py

Et Tkinter-basert GUI som:
  - Leser en Excel-fil med EEG-kategorier (30 rader, 8 kolonner).
  - Lar brukeren velge "Recording X" i en Combobox.
  - Viser en sekvens med 30 kategorier (hver 10 sekunder), med nedtelling.
  - Kaller BrainVision Recorder via OLE i main thread for å unngå
    'marshalled for a different thread'-feil.
"""

import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog  # For mappe-dialog
import threading
import time
import pandas as pd

import pythoncom
import win32com.client

EXCEL_PATH = "EEG kategori rekkefølge datasett.xlsx"

SHORTKEY_MAP = {
    "REST": 1,
    "MOVE_RIGHT": 2,
    "MOVE_LEFT": 3,
    "MOVE_BOTH": 4,
    "IMAGERY_RIGHT": 5,
    "IMAGERY_LEFT": 6,
    "IMAGERY_BOTH": 7,
    "START": 8,
    "END": 9
}

CATEGORY_DURATION = 10
DELAY_BEFORE_START = 10

DEFAULT_EEG_FOLDER = r"C:\Users\vislab\Documents\Master Aksel\EEG_master\EEG_dataset"
DEFAULT_EEG_FILENAME = "OLE_Recording.eeg"


class RecordingUI:
    def __init__(self, root):
        self.root = root
        self.root.title("EEG Category Playback - OLE (Thread-Safe Edition)")
        self.root.geometry("1200x800")
        self.root.configure(bg="#FFFFFF")

        # TTK-stil
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background="#FFFFFF")
        style.configure("TLabel", font=("Helvetica", 16), background="#FFFFFF")
        style.configure("Header.TLabel", font=("Helvetica", 24, "bold"), background="#FFFFFF")
        style.configure("Big.TLabel", font=("Helvetica", 56, "bold"), background="#E0FFE0")
        style.configure("Medium.TLabel", font=("Helvetica", 22), background="#FFFFD0")
        style.configure("Small.TLabel", font=("Helvetica", 16), background="#E0E0FF")
        style.configure("TButton", font=("Helvetica", 14), padding=10)

        # (1) EGEN STIL FOR MINDRE KNAPP
        style.configure("Browse.TButton",
                        font=("Helvetica", 10),   # mindre font
                        padding=(5, 2))         # mindre padding (horisontal, vertikal)

        self.is_running = False
        self.current_index = 0
        self.remaining_time = CATEGORY_DURATION

        # OLE: Opprett connection til Recorder i main-tråd
        self.recorder = None
        self.init_recorder()

        # Les Excel
        self.df = pd.read_excel(EXCEL_PATH)
        self.recording_cols = self.df.columns[1:]  # kolonne 0 = 'Category', 1..8 = 'Recording X'

        main_frame = ttk.Frame(root, padding="20 20 20 20", style="TFrame")
        main_frame.pack(fill="both", expand=True)

        top_frame = ttk.Frame(main_frame, style="TFrame")
        top_frame.pack(fill="x", pady=10)
        self.title_label = ttk.Label(top_frame, text="EEG Category Playback (OLE)", style="Header.TLabel")
        self.title_label.pack()

        # ---------- Ramme for mappe- og filnavn ----------
        storage_frame = ttk.Frame(main_frame, style="TFrame")
        storage_frame.pack(pady=15, anchor="n")

        # Juster kolonnevekter for å få en litt mer sentrert layout
        storage_frame.grid_columnconfigure(0, weight=1)
        storage_frame.grid_columnconfigure(1, weight=1)
        storage_frame.grid_columnconfigure(2, weight=1)

        folder_label = ttk.Label(storage_frame, text="Lagringsmappe:")
        folder_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")

        self.folder_path_var = tk.StringVar(value=DEFAULT_EEG_FOLDER)
        folder_entry = ttk.Entry(storage_frame, textvariable=self.folder_path_var, width=45)
        folder_entry.grid(row=0, column=1, padx=5, pady=5, sticky="we")

        # (1) Bruker "Browse.TButton"-stil for å gjøre knappen mindre
        browse_button = ttk.Button(storage_frame, text="Browse...", style="Browse.TButton",
                                   command=self.browse_folder)
        browse_button.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        
        # (2) ELLER sett `width` direkte, f.eks. for å gjøre den smalere
        # browse_button.config(width=8)

        filename_label = ttk.Label(storage_frame, text="Filnavn:")
        filename_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")

        self.filename_var = tk.StringVar(value=DEFAULT_EEG_FILENAME)
        filename_entry = ttk.Entry(storage_frame, textvariable=self.filename_var, width=45)
        filename_entry.grid(row=1, column=1, padx=5, pady=5, sticky="we")

        # ---------- Velg Recording ----------
        rec_frame = ttk.Frame(main_frame, style="TFrame")
        rec_frame.pack(pady=10)
        ttk.Label(rec_frame, text="Velg Recording:").pack(side="left", padx=5)
        self.recording_var = tk.StringVar()
        self.recording_combobox = ttk.Combobox(
            rec_frame, textvariable=self.recording_var,
            values=list(self.recording_cols), width=20, state="readonly"
        )
        self.recording_combobox.pack(side="left", padx=5)
        self.recording_combobox.set(self.recording_cols[0])

        middle_frame = ttk.Frame(main_frame, style="TFrame")
        middle_frame.pack(fill="both", expand=True, pady=30)
        self.current_cat_label = tk.Label(
            middle_frame, text="None", font=("Helvetica", 56, "bold"),
            bg="#E0FFE0", highlightbackground="#A0D0A0", highlightthickness=2
        )
        self.current_cat_label.pack(expand=True, fill="both", padx=100, pady=20)

        self.countdown_label = tk.Label(
            main_frame, text="Countdown: 0 s",
            font=("Helvetica", 22), bg="#FFFFD0",
            highlightbackground="#D0D0A0", highlightthickness=2
        )
        self.countdown_label.pack(pady=10, fill="x", padx=200)

        self.next_cat_label = tk.Label(
            main_frame, text="Next category: None",
            font=("Helvetica", 16), bg="#E0E0FF",
            highlightbackground="#A0A0D0", highlightthickness=2
        )
        self.next_cat_label.pack(pady=10, fill="x", padx=250)

        button_frame = ttk.Frame(main_frame, style="TFrame")
        button_frame.pack(pady=20)

        self.start_button = tk.Button(
            button_frame, text="Start Recording", command=self.start_recording,
            bg="#00CC00", fg="white", font=("Helvetica", 14, "bold"),
            padx=20, pady=10
        )
        self.start_button.pack(side="left", padx=15)

        self.stop_button = tk.Button(
            button_frame, text="End Recording", command=self.stop_recording,
            bg="#FF0000", fg="white", font=("Helvetica", 14, "bold"),
            padx=20, pady=10
        )
        self.stop_button.pack(side="left", padx=15)

    def init_recorder(self):
        """Opprett COM-objekt i main-tråd, slik at OLE-kall kan gjøres her."""
        try:
            pythoncom.CoInitialize()
            self.recorder = win32com.client.Dispatch("VisionRecorder.Application")
            self.recorder.DisableThreadBlockingMode = 1
            print("Tilkoblet BrainVision Recorder via OLE.")
        except Exception as e:
            print("Kunne ikke koble til Recorder:", e)
            self.recorder = None

    def browse_folder(self):
        """Åpne en dialog for å velge mappe."""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path_var.set(folder_selected)

    def start_recording(self):
        if self.is_running:
            return
        if not self.recorder:
            print("Recorder er ikke tilgjengelig via OLE. Avbryter.")
            return

        self.is_running = True
        self.current_index = 0
        self.remaining_time = CATEGORY_DURATION

        # Hent mappe og filnavn fra GUI
        folder = self.folder_path_var.get()
        filename = self.filename_var.get()
        if not filename.lower().endswith(".eeg"):
            filename += ".eeg"

        os.makedirs(folder, exist_ok=True)
        full_eeg_path = os.path.join(folder, filename)

        # Start opptak i main-tråd
        try:
            self.recorder.Acquisition.StartRecording(full_eeg_path, "OLE testopptak")
            print(f"Startet opptak -> {full_eeg_path}")
        except Exception as e:
            print(f"Feil ved StartRecording: {e}")

        # Forsinkelse (10 sek)
        self.update_current_category("PREPARATION")
        for i in range(DELAY_BEFORE_START, 0, -1):
            self.update_next_category(f"Starting in {i} s...")
            self.update_countdown(i)
            self.root.update()
            time.sleep(1)

        chosen_col = self.recording_var.get()
        self.categories_seq = self.build_categories(chosen_col)

        # Start bakgrunnstråd for sekvens
        t = threading.Thread(target=self.run_sequence, daemon=True)
        t.start()

    def stop_recording(self):
        self.is_running = False
        self.update_current_category("Recording Stopped")
        self.update_next_category("None")
        self.update_countdown(0)

        if self.recorder:
            try:
                self.recorder.Acquisition.StopRecording()
                print("Stoppet opptak i Recorder (OLE).")
            except Exception as e:
                print(f"Feil ved StopRecording: {e}")

    def build_categories(self, col_name):
        cat_list = []
        for i in range(30):
            cat_text = str(self.df.loc[i, col_name])
            key = SHORTKEY_MAP.get(cat_text, 1)
            cat_list.append((key, cat_text))
        return cat_list

    def run_sequence(self):
        """Bakgrunnstråd: kjører kategori-sekvens (10 sek per kategori)."""
        for idx, (key_code, cat_text) in enumerate(self.categories_seq):
            if not self.is_running:
                break
            # Bestill GUI-oppdatering + marker i main-tråd
            self.root.after(0, self.schedule_marker_and_gui, cat_text, idx)

            # Teller ned i bakgrunnen
            for s in range(CATEGORY_DURATION):
                if not self.is_running:
                    break
                remaining = CATEGORY_DURATION - s
                self.root.after(0, self.update_countdown, remaining)
                time.sleep(1)

        self.is_running = False
        self.root.after(0, self.finish_sequence)

    def schedule_marker_and_gui(self, cat_text, idx):
        self.update_current_category(cat_text)

        if idx + 1 < len(self.categories_seq):
            next_cat = self.categories_seq[idx + 1][1]
        else:
            next_cat = "None"
        self.update_next_category(next_cat)

        if self.recorder:
            try:
                self.recorder.Acquisition.SetMarker(cat_text, "Stimulus")
                print(f"Sendte OLE-markør: {cat_text}")
            except Exception as e:
                print(f"Feil ved SetMarker: {e}")

    def finish_sequence(self):
        self.update_current_category("Completed")
        self.update_next_category("None")
        self.update_countdown(0)

    # GUI-hjelpefunksjoner
    def update_current_category(self, cat):
        self.current_cat_label.config(text=f"{cat}")

    def update_next_category(self, cat):
        self.next_cat_label.config(text=f"Next category: {cat}")

    def update_countdown(self, sec):
        self.countdown_label.config(text=f"Countdown: {sec} s")


def main():
    root = tk.Tk()
    app = RecordingUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
