import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import os
import sys


class ModStatusApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Aircraft Mod Status Checker")
        self.root.geometry("520x620")

        self.aircraft_dict = {
            "A321-231": ["TC-MYA", "TC-MYB"],
            "A330-343": ["TC-MCM", "TC-MCN", "TC-MCO", "TC-MCP"],
            "A330-243": ["TC-MCU", "TC-MCZ"]
        }
        self.selected_aircraft = []
        self.entered_mods    = []      

        self.setup_ui()

    def setup_ui(self):

        # Excel file selection
        tk.Label(self.root, text="Select Database Excel File:").grid(row=0, column=0, sticky="w", padx=(10,0), pady=5)
        self.file_path = tk.StringVar()
        # Label stays in column 0
        tk.Label(self.root, text="Select Database Excel File:")\
            .grid(row=0, column=0, sticky="w", padx=10, pady=5)

        # Frame spans columns 1..3 and stretches horizontally
        file_frame = tk.Frame(self.root)
        file_frame.grid(row=0, column=1, columnspan=3, sticky="we", padx=(0,10), pady=5)
        file_frame.columnconfigure(0, weight=1)  # entry expands inside frame

        self.file_path = tk.StringVar()
        tk.Entry(file_frame, textvariable=self.file_path)\
            .grid(row=0, column=0, sticky="we")
        tk.Button(file_frame, text="Browse", command=self.browse_file)\
            .grid(row=0, column=1, sticky="w", padx=(6,0))


        tk.Label(self.root, text="↓ Selected Items").grid(row=1, column=3, sticky="sw")

        # Aircraft group checkboxes
        self.group_vars = {group: tk.BooleanVar() for group in self.aircraft_dict}
        group_frame = tk.Frame(self.root)
        group_frame.grid(row=2, column=0, pady=(0,0))
        for i, (group, var) in enumerate(self.group_vars.items()):
            tk.Checkbutton(group_frame, text=group, variable=var, command=self.update_aircraft_selection).grid(row=i, column=0, sticky="w")

        # Aircraft selection list
        tk.Label(self.root, text="Select Aircraft:").grid(row=2, column=0, sticky="nw", padx=(10,0), pady=5)
        self.aircraft_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE, width=14, height=8, exportselection=False)
        self.aircraft_listbox.grid(row=2, column=1, sticky="nw", pady=5)

        # show all aircraft by default and remember their indices
        self.all_aircraft = [ac for group in self.aircraft_dict.values() for ac in group]
        self.ac_index = {}
        for i, ac in enumerate(self.all_aircraft):
            self.aircraft_listbox.insert(tk.END, ac)
            self.ac_index[ac] = i

        # Confirm aircraft and show in right-hand box
        tk.Button(self.root,
                  text="Confirm\nSelection\n→",
                  command=self.confirm_aircraft
                 ).grid(row=2, column=2, padx=(0,15), pady=(0,50))

        # Clear all confirmed aircraft (right box)
        tk.Button(self.root,
                  text="Clear Selected",
                  command=self.clear_selected_aircrafts
                 ).grid(row=2, column=2, padx=(0,15), pady=(50,0))

        self.final_aircraft_listbox = tk.Listbox(self.root, height=8, width=14)
        self.final_aircraft_listbox.grid(row=2, column=3, pady=5, sticky="nw")

        # Mod entry
        tk.Label(self.root, text="Enter Mod Numbers\n(one per line):").grid(row=3, column=0, sticky="nw", padx=(10,0), pady=5)
        self.mod_text = scrolledtext.ScrolledText(self.root, width=10, height=6, undo=True, maxundo=-1)
        self.mod_text.grid(row=3, column=1, pady=5, sticky="nsw")

        # Keyboard shortcuts: Ctrl + A = select all
        self.mod_text.bind("<Control-a>", self._mod_select_all)
        self.mod_text.bind("<Control-A>", self._mod_select_all)

        # Build context menu and bind right-clicks
        self._build_mod_context_menu()
        self.mod_text.bind("<Button-3>", self._show_mod_context)  
        self.mod_text.bind("<Control-Button-1>", self._show_mod_context) 

        # Add mods and show in right-hand box
        tk.Button(self.root,
                  text="Add Mods\n→",
                  command=self.add_mods
                 ).grid(row=3, column=2, padx=(0,15), pady=(0,35))

        self.clear_mods_btn = tk.Button(self.root, text="Clear Mods", command=self.clear_mods)
        self.clear_mods_btn.grid(row=3, column=2, padx=(0,15), pady=(50,0))

        self.final_mod_listbox = scrolledtext.ScrolledText(self.root, height=6, width=10)
        self.final_mod_listbox.grid(row=3, column=3, pady=5, sticky="nsw")

        # Check button
        tk.Button(self.root, text="Check Mod Status", command=self.check_mod_status, bg="lightblue").grid(row=4, column=1, columnspan=3, padx=(0,18), pady=10)

        # Results display
        tk.Label(self.root, text="Results:").grid(row=5, column=0, sticky="nw", padx=(10,0))
        self.result_box = scrolledtext.ScrolledText(self.root, width=40, height=15)
        self.result_box.grid(row=5, column=1, sticky="ns", columnspan=3, padx=(0,10), pady=5)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.file_path.set(file_path)

    def update_aircraft_selection(self):
        """
        Auto-select aircraft in the left listbox based on checked groups.
        All aircraft stay visible; selection reflects current checkboxes.
        """
        # Clear all current selections
        self.aircraft_listbox.selection_clear(0, tk.END)

        # Select the union of all checked groups
        for group, var in self.group_vars.items():
            if var.get():
                for ac in self.aircraft_dict[group]:
                    idx = self.ac_index.get(ac)
                    if idx is not None:
                        self.aircraft_listbox.selection_set(idx)

    def check_mod_status(self):
        filepath = self.file_path.get()
        if not filepath or not os.path.exists(filepath):
            messagebox.showerror("Error", "Please select a valid Excel file.")
            return

        selected_indices = self.aircraft_listbox.curselection()
        all_aircraft = self.aircraft_listbox.get(0, tk.END)
        selected_aircraft = list(self.final_aircraft_listbox.get(0, tk.END))

        if not selected_aircraft:
            messagebox.showerror("Error", "Please select at least one aircraft.")
            return

        # Get only the finalized mod numbers and parse to int
        right_text = self.final_mod_listbox.get("1.0", "end-1c")
        entered_mods = [int(x) for x in right_text.splitlines() if x.strip().isdigit()]
        if not entered_mods:
            messagebox.showerror("Error", "Please enter at least one valid mod number.")
            return

        try:
            df = pd.read_excel(filepath)

            self.result_box.delete(1.0, tk.END)
            for mod in entered_mods:
                self.result_box.insert(tk.END, f"Mod Number: {mod}\n")
                for ac in selected_aircraft:
                    if ac in df.columns:
                        # Convert that column to numeric ints
                        col_vals = (
                            pd.to_numeric(df[ac], errors="coerce")
                              .dropna()
                              .astype(int)
                              .tolist()
                        )
                        status = "POST" if mod in col_vals else "PRE"
                        self.result_box.insert(tk.END, f"  {ac} : {status}\n")
                    else:
                        self.result_box.insert(tk.END, f"  {ac} : Not found in database\n")
                self.result_box.insert(tk.END, "\n")

        except Exception as e:
            messagebox.showerror("Processing Error", str(e))

    def confirm_aircraft(self):
        # Grab checked items on left, store & mirror to right
        self.selected_aircraft = [
            self.aircraft_listbox.get(i)
            for i in self.aircraft_listbox.curselection()
        ]
        self.final_aircraft_listbox.delete(0, tk.END)
        for ac in self.selected_aircraft:
            self.final_aircraft_listbox.insert(tk.END, ac)

    def add_mods(self):
        """
        Append numbers from the left box to the right box (ScrolledText),
        adding only those not already present.
        Accepts one-per-line, commas and spaces.
        """
        # tokenize left box
        raw = self.mod_text.get("1.0", "end-1c")
        tokens = []
        for line in raw.splitlines():
            tokens.extend(line.replace(",", " ").split())
        new_mods = [t for t in tokens if t.isdigit()]
        if not new_mods:
            return

        # current contents of the right box
        existing_text = self.final_mod_listbox.get("1.0", "end-1c")
        existing = set(x.strip() for x in existing_text.splitlines() if x.strip())

        # append only unique
        for m in new_mods:
            if m not in existing:
                self.final_mod_listbox.insert("end", m + "\n")
                existing.add(m)

        # clear the left box after adding
        self.mod_text.delete("1.0", "end")

    def clear_mods(self):
        # clear right text box
        self.final_mod_listbox.delete("1.0", "end")
        # clear left input
        self.mod_text.delete("1.0", "end")
        # clear cache
        self.entered_mods.clear()
        # (optional) put cursor back in input
        self.mod_text.focus_set()

    def clear_selected_aircrafts(self):
        """Clear final aircraft, left selection, and uncheck all group boxes."""
        # clear right box + cache
        self.final_aircraft_listbox.delete(0, tk.END)
        self.selected_aircraft.clear()

        # clear left listbox selection
        self.aircraft_listbox.selection_clear(0, tk.END)

        # untick all aircraft-type checkboxes
        for var in self.group_vars.values():
            var.set(False)

        # if you use an auto-selection helper, call it to sync visuals
        if hasattr(self, "update_aircraft_selection"):
            self.update_aircraft_selection()

    def _mod_select_all(self, event=None):
        # Select all text in the mod box
        self.mod_text.tag_add("sel", "1.0", "end-1c")
        self.mod_text.mark_set("insert", "end-1c")
        return "break"  # prevent default beep

    def _build_mod_context_menu(self):
        # Right-click menu for the mod box
        self._mod_ctx = tk.Menu(self.root, tearoff=0)
        self._mod_ctx.add_command(label="Undo",  command=lambda: self.mod_text.event_generate("<<Undo>>"))
        self._mod_ctx.add_command(label="Redo",  command=lambda: self.mod_text.event_generate("<<Redo>>"))
        self._mod_ctx.add_separator()
        self._mod_ctx.add_command(label="Cut",   command=lambda: self.mod_text.event_generate("<<Cut>>"))
        self._mod_ctx.add_command(label="Copy",  command=lambda: self.mod_text.event_generate("<<Copy>>"))
        self._mod_ctx.add_command(label="Paste", command=lambda: self.mod_text.event_generate("<<Paste>>"))
        self._mod_ctx.add_separator()
        self._mod_ctx.add_command(label="Select All", command=self._mod_select_all)

    def _show_mod_context(self, event):
        try:
            self._mod_ctx.tk_popup(event.x_root, event.y_root)
        finally:
            self._mod_ctx.grab_release()

if __name__ == "__main__":
    root = tk.Tk()
    app = ModStatusApp(root)
    root.mainloop()
