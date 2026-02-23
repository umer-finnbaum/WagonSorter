"""
Simple PLC simulator for testing AutoStacker without a real PLC.

How to use:
1) Run this script alongside your app:
   python plc_simulator.py

2) In the same shell (or before launching your app), set the environment variable:
   On Windows (cmd):    set PLCSIM=1
   On Windows (PowerShell): $env:PLCSIM = "1"
   On Linux/macOS:      export PLCSIM=1

3) Start AutoStacker.py
   Use the GUI to change Wacon_INT[1], Wacon_INT[2],and Wacon_SerieName.
   The file plc_sim.json is updated automatically.
"""
import tkinter as tk
from tkinter import ttk
import json
import threading
import time
import os

SIM_FILENAME = "plc_sim.json"
UPDATE_INTERVAL = 0.25  # seconds: how often the file is refreshed when "running"

# Default simulated values
_default_state = {
    "Wacon_INT[1]": None,
    "Wacon_INT[2]": None,
    "Wacon_SerieName": ""
}


class PlcSimulatorGUI:
    def __init__(self, master):
        self.master = master
        master.title("PLC Simulator (plc_sim.json)")

        frm = ttk.Frame(master, padding=10)
        frm.grid(row=0, column=0, sticky="nsew")

        # ID1
        ttk.Label(frm, text="Wacon_INT[1]").grid(row=0, column=0, sticky="w")
        self.id1_var = tk.StringVar(value="")
        self.id1_entry = ttk.Entry(frm, textvariable=self.id1_var, width=20)
        self.id1_entry.grid(row=0, column=1, padx=6, pady=2)

        # ID2
        ttk.Label(frm, text="Wacon_INT[2]").grid(row=1, column=0, sticky="w")
        self.id2_var = tk.StringVar(value="")
        self.id2_entry = ttk.Entry(frm, textvariable=self.id2_var, width=20)
        self.id2_entry.grid(row=1, column=1, padx=6, pady=2)

        # Series name (string)
        ttk.Label(frm, text="Wacon_SerieName").grid(row=2, column=0, sticky="w")
        self.series_var = tk.StringVar(value="")
        self.series_entry = ttk.Entry(frm, textvariable=self.series_var, width=30)
        self.series_entry.grid(row=2, column=1, padx=6, pady=2)

        # Buttons: Write, Clear, Start sequence, Stop
        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=(8, 0), sticky="ew")

        self.write_btn = ttk.Button(btn_frame, text="Write Now", command=self.write_now)
        self.write_btn.pack(side="left", padx=2)

        self.clear_btn = ttk.Button(btn_frame, text="Clear IDs", command=self.clear_ids)
        self.clear_btn.pack(side="left", padx=2)

        self.seq_btn = ttk.Button(btn_frame, text="Start Demo Sequence", command=self.toggle_sequence)
        self.seq_btn.pack(side="left", padx=2)

        self.quit_btn = ttk.Button(btn_frame, text="Quit", command=self.quit)
        self.quit_btn.pack(side="right", padx=2)

        # Status label
        self.status_label = ttk.Label(frm, text="Stopped", foreground="blue")
        self.status_label.grid(row=4, column=0, columnspan=2, pady=(8,0))

        self.running = False
        self._seq_thread = None
        self._stop_event = threading.Event()

        # Ensure file exists on startup
        self.write_file(_default_state)

    def write_file(self, data):
        try:
            with open(SIM_FILENAME, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            self.status_label.config(text=f"Write error: {e}", foreground="red")

    def read_ui_state(self):
        # Convert empty inputs to None for IDs (mimics PLC empty)
        id1 = self.id1_var.get().strip()
        id2 = self.id2_var.get().strip()
        series = self.series_var.get()
        return {
            "Wacon_INT[1]": (None if id1 == "" else id1),
            "Wacon_INT[2]": (None if id2 == "" else id2),
            "Wacon_SerieName": series
        }

    def write_now(self):
        data = self.read_ui_state()
        self.write_file(data)
        self.status_label.config(text="Wrote plc_sim.json", foreground="green")

    def clear_ids(self):
        self.id1_var.set("")
        self.id2_var.set("")
        self.write_now()

    def toggle_sequence(self):
        if self.running:
            self.stop_sequence()
        else:
            self.start_sequence()

    def start_sequence(self):
        if self.running:
            return
        self._stop_event.clear()
        self.running = True
        self.status_label.config(text="Running sequence...", foreground="blue")
        self._seq_thread = threading.Thread(target=self._sequence_worker, daemon=True)
        self._seq_thread.start()

    def stop_sequence(self):
        if not self.running:
            return
        self._stop_event.set()
        if self._seq_thread:
            self._seq_thread.join(timeout=1.0)
        self.running = False
        self.status_label.config(text="Stopped", foreground="blue")

    def _sequence_worker(self):
        # A short demo sequence
        sequence = [
            {"Wacon_INT[1]": "15", "Wacon_INT[2]": "5",  "Wacon_SerieName": "SeriesDemo"},
            {"Wacon_INT[1]": -2, "Wacon_INT[2]": -2, "Wacon_SerieName": "SeriesDemo"},
            {"Wacon_INT[1]": "17", "Wacon_INT[2]": "7",  "Wacon_SerieName": "SeriesDemo"},
            {"Wacon_INT[1]": "20", "Wacon_INT[2]": -2, "Wacon_SerieName": "SeriesDemo"},
            {"Wacon_INT[1]": -2, "Wacon_INT[2]": -2, "Wacon_SerieName": ""},
        ]
        # Optionally repeat sequence until stopped
        while not self._stop_event.is_set():
            for step in sequence:
                if self._stop_event.is_set():
                    break
                # update UI fields for visibility
                self.id1_var.set("" if step["Wacon_INT[1]"] is None else str(step["Wacon_INT[1]"]))
                self.id2_var.set("" if step["Wacon_INT[2]"] is None else str(step["Wacon_INT[2]"]))
                self.series_var.set(step["Wacon_SerieName"])
                self.write_file(step)
                time.sleep(2.0)  # pause between steps
            # after one pass, wait a little before repeating
            time.sleep(1.0)
        # sequence stopped
        self.running = False
        self.status_label.config(text="Stopped", foreground="blue")

    def quit(self):
        self.stop_sequence()
        self.master.quit()


def run_gui():
    root = tk.Tk()
    app = PlcSimulatorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    print("PLC Simulator starting. This writes/updates plc_sim.json in the current directory.")
    print("Set environment variable PLCSIM=1 for AutoStacker to read the simulated values.")
    run_gui()