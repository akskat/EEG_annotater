#!/usr/bin/env python3
"""
new_annotater.py


Et Tkinter-basert GUI som:
  - Leser en Excel-fil med EEG-kategorier (30 rader, 8 kolonner).
  - Lar brukeren velge "Recording X" i en Combobox.
  - Viser en sekvens med 30 kategorier (hver 10 sekunder), med nedtelling.
  - Kaller BrainVision Recorder via OLE i **main thread**, for å unngå
    'marshalled for a different thread' feil.


Vi bruker:
 - en bakgrunnstråd til venting/time.sleep
 - men *all* OLE (SetMarker, Start/StopRecording) i main-tråden via `root.after(...)`.
"""


import os
import tkinter as tk
from tkinter import ttk
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


# Sti for EEG-lagring
EEG_FOLDER = r"C:\Users\vislab\Documents\Master Aksel\EEG_master\EEG_dataset"
EEG_FILENAME = "OLE_Recording.eeg"




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


        self.is_running = False
        self.current_index = 0
        self.remaining_time = CATEGORY_DURATION


        # OLE: Opprett connection til Recorder i main-thread
        self.recorder = None
        self.init_recorder()


        # Les Excel
        self.df = pd.read_excel(EXCEL_PATH)
        self.recording_cols = self.df.columns[1:]  # kolonne 0 = 'Category', 1..8 = 'Recording X'


        # GUI-oppsett
        main_frame = ttk.Frame(root, padding="20 20 20 20", style="TFrame")
        main_frame.pack(fill="both", expand=True)


        top_frame = ttk.Frame(main_frame, style="TFrame")
        top_frame.pack(fill="x", pady=10)
        self.title_label = ttk.Label(top_frame, text="EEG Category Playback (OLE)", style="Header.TLabel")
        self.title_label.pack()


        # Velg Recording
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


        # Gjeldende kategori
        middle_frame = ttk.Frame(main_frame, style="TFrame")
        middle_frame.pack(fill="both", expand=True, pady=30)
        self.current_cat_label = tk.Label(
            middle_frame, text="None", font=("Helvetica", 56, "bold"),
            bg="#E0FFE0", highlightbackground="#A0D0A0", highlightthickness=2
        )
        self.current_cat_label.pack(expand=True, fill="both", padx=100, pady=20)


        # Nedtelling
        self.countdown_label = tk.Label(
            main_frame, text="Countdown: 0 s",
            font=("Helvetica", 22), bg="#FFFFD0",
            highlightbackground="#D0D0A0", highlightthickness=2
        )
        self.countdown_label.pack(pady=10, fill="x", padx=200)


        # Neste kategori
        self.next_cat_label = tk.Label(
            main_frame, text="Next category: None",
            font=("Helvetica", 16), bg="#E0E0FF",
            highlightbackground="#A0A0D0", highlightthickness=2
        )
        self.next_cat_label.pack(pady=10, fill="x", padx=250)


        # Knapper
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


    def start_recording(self):
        if self.is_running:
            return
        if not self.recorder:
            print("Recorder er ikke tilgjengelig via OLE. Avbryter.")
            return


        self.is_running = True
        self.current_index = 0
        self.remaining_time = CATEGORY_DURATION


        # Opprett mappe for EEG-filen
        os.makedirs(EEG_FOLDER, exist_ok=True)
        full_eeg_path = os.path.join(EEG_FOLDER, EEG_FILENAME)


        # Start opptak i main-tråd
        try:
            self.recorder.Acquisition.StartRecording(full_eeg_path, "OLE testopptak")
            print("Startet opptak i Recorder (OLE).")
        except Exception as e:
            print(f"Feil ved StartRecording: {e}")


        # 10 sek forsinkelse i main-tråd (bare for visning):
        self.update_current_category("PREPARATION")
        for i in range(DELAY_BEFORE_START, 0, -1):
            self.update_next_category(f"Starting in {i} s...")
            self.update_countdown(i)
            self.root.update()
            time.sleep(1)


        # Bygg liste av kategorier
        chosen_col = self.recording_var.get()
        self.categories_seq = self.build_categories(chosen_col)


        # Start en bakgrunnstråd for "sekvens-timing"
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
        """
        Bakgrunnstråd: Teller ned 10 sek per kategori.
        Men *OLE-kall* kan ikke gjøres direkte her,
        i stedet ber vi main-tråd gjøre SetMarker via root.after.
        """
        for idx, (key_code, cat_text) in enumerate(self.categories_seq):
            if not self.is_running:
                break


            # main-tråd: oppdater GUI og send marker
            self.root.after(0, self.schedule_marker_and_gui, cat_text, idx)


            # vent 10 sek i bakgrunn
            for s in range(CATEGORY_DURATION):
                if not self.is_running:
                    break
                # main-tråd: oppdater countdown
                remaining = CATEGORY_DURATION - s
                self.root.after(0, self.update_countdown, remaining)
                time.sleep(1)


        # Avslutt
        self.is_running = False
        self.root.after(0, self.finish_sequence)


    def schedule_marker_and_gui(self, cat_text, idx):
        """
        Kalles i main-tråd via root.after(0,...).
        Oppdaterer GUI + kaller SetMarker i main-tråden.
        """
        # Oppdater GUI
        self.update_current_category(cat_text)


        # Neste kategori
        if idx + 1 < len(self.categories_seq):
            next_cat = self.categories_seq[idx + 1][1]
        else:
            next_cat = "None"
        self.update_next_category(next_cat)


        # Kall SetMarker i main-tråden
        if self.recorder:
            try:
                self.recorder.Acquisition.SetMarker(cat_text, "Stimulus")
                print(f"Sendte OLE-markør: {cat_text}")
            except Exception as e:
                print(f"Feil ved SetMarker (main thread): {e}")


    def finish_sequence(self):
        """Når alt er ferdig, sett GUI-siste status."""
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



