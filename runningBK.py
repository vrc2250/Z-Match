import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
import re
import io
import shutil
from pathlib import Path

# IMPORTANT: Ensure you ran 'pip install fpdf2'
try:
    from fpdf import FPDF
except ImportError:
    FPDF = None

class AudioRenamerApp:
    def __init__(self, root):
        self.root = root
        # Title updated as requested
        self.root.title("Z-Transmitter Scene/File Conform")
        self.root.geometry("1250x850") 

        self.df1 = None 
        self.df2 = None 
        self.matched_pairs = []
        self.filtered_df1_indices = []
        self.track_mapping = {}
        self.active_track_cols = []
        
        # Default Locations
        self.home_dir = str(Path.home())
        self.downloads_dir = str(Path.home() / "Downloads")
        self.transmitter_source_dir = None 
        
        self.sound_roll_var = tk.StringVar(value="---")
        
        self.match_colors = [
            {"bg": "#c8e6c9", "fg": "black"}, 
            {"bg": "#bbdefb", "fg": "black"}, 
            {"bg": "#fff9c4", "fg": "black"}, 
            {"bg": "#f8bbd0", "fg": "black"}, 
            {"bg": "#e1bee7", "fg": "black"}, 
            {"bg": "#ffccbc", "fg": "black"}, 
        ]

        self.create_widgets()

    def create_widgets(self):
        top_frame = tk.LabelFrame(self.root, text="Configuration & Recorder Metadata")
        top_frame.pack(pady=10, fill='x', padx=10)

        tk.Button(top_frame, text="Load Recorder CSV", command=lambda: self.load_file(1), width=18).grid(row=0, column=0, rowspan=2, padx=2, pady=10)
        tk.Button(top_frame, text="Load Transmitter CSV", command=lambda: self.load_file(2), width=18).grid(row=0, column=1, rowspan=2, padx=2)

        tk.Label(top_frame, text="Offset(s):").grid(row=0, column=2, rowspan=2, padx=2)
        self.time_offset_var = tk.DoubleVar(value=3.0)
        tk.Spinbox(top_frame, from_=0, to=15, increment=1.0, textvariable=self.time_offset_var, width=4).grid(row=0, column=3, rowspan=2)

        tk.Label(top_frame, text="Sound Roll:").grid(row=0, column=4, sticky='e', padx=5)
        self.sound_roll_display = tk.Label(top_frame, textvariable=self.sound_roll_var, font=("Arial", 11, "bold"), fg="#D32F2F")
        self.sound_roll_display.grid(row=0, column=5, sticky='w')

        tk.Label(top_frame, text="Track:").grid(row=1, column=4, sticky='e', padx=5)
        self.track_select_combo = ttk.Combobox(top_frame, width=25, state="readonly")
        self.track_select_combo.grid(row=1, column=5, sticky='w')
        self.track_select_combo.bind("<<ComboboxSelected>>", self.refresh_recorder_list)

        self.match_btn = tk.Button(top_frame, text="Match & Preview", command=self.compare_and_preview, bg="#2196F3", fg="white", width=14, state="disabled")
        self.match_btn.grid(row=0, column=6, rowspan=2, padx=10)
        
        self.rename_btn = tk.Button(top_frame, text="COMMIT RENAME", command=self.execute_rename, bg="#4CAF50", fg="white", width=14, state="disabled")
        self.rename_btn.grid(row=0, column=7, rowspan=2, padx=5)

        paned = tk.PanedWindow(self.root, orient=tk.VERTICAL)
        paned.pack(expand=True, fill='both', padx=10)

        list_frame = tk.Frame(paned)
        self.recorder_frame = tk.LabelFrame(list_frame, text="Recorder Metadata")
        self.recorder_frame.pack(side='left', expand=True, fill='both', padx=2)
        self.tree1 = self.internal_setup_tree(self.recorder_frame, ("TimeCode", "UserBits", "Scene", "Take", "All Tracks / Names"))

        self.transmitter_frame = tk.LabelFrame(list_frame, text="Transmitter Files")
        self.transmitter_frame.pack(side='left', expand=True, fill='both', padx=2)
        self.tree2 = self.internal_setup_tree(self.transmitter_frame, ("TimeCode", "UserBits", "FileName"))
        paned.add(list_frame)

        self.preview_frame = tk.LabelFrame(paned, text="Proposed Filename Changes")
        self.preview_tree = ttk.Treeview(self.preview_frame, columns=("Old", "New", "Status", "Offset"), show='headings')
        self.preview_tree.heading("Old", text="Original Filename")
        self.preview_tree.heading("New", text="New Filename")
        self.preview_tree.heading("Status", text="ID")
        self.preview_tree.heading("Offset", text="Offset")
        self.preview_tree.column("Old", width=350, stretch=True)
        self.preview_tree.column("New", width=350, stretch=True)
        self.preview_tree.column("Status", width=60, anchor='center', stretch=False)
        self.preview_tree.column("Offset", width=80, anchor='center', stretch=False)
        self.preview_tree.pack(expand=True, fill='both')
        paned.add(self.preview_frame)

    def internal_setup_tree(self, parent_frame, columns):
        tree = ttk.Treeview(parent_frame, columns=columns, show='headings')
        for col in columns:
            tree.heading(col, text=col)
            width = 200 if "Names" in col else 110
            tree.column(col, width=width)
        scrollbar = ttk.Scrollbar(parent_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side='left', expand=True, fill='both')
        scrollbar.pack(side='right', fill='y')
        for i, color in enumerate(self.match_colors):
            tree.tag_configure(f'match_{i}', background=color["bg"], foreground=color["fg"])
        return tree

    def check_ready_state(self):
        if self.df1 is not None and self.df2 is not None and self.track_select_combo.get() != "Select Track":
            self.match_btn.config(state="normal")
        else:
            self.match_btn.config(state="disabled")

    def load_file(self, list_num):
        path = filedialog.askopenfilename(
            initialdir=self.home_dir, 
            title=f"Select {'Recorder' if list_num==1 else 'Transmitter'} CSV File",
            filetypes=[("CSV files", "*.csv")]
        )
        if not path: return
        try:
            with open(path, 'r', encoding='latin1') as f:
                content = f.read()
            lines = content.splitlines()
            
            found_roll = "---"
            if list_num == 1:
                for line in lines:
                    if "FolderName" in line:
                        parts = line.split("FolderName")
                        if len(parts) > 1:
                            val = parts[1].replace("=", "").replace(":", "").strip()
                            found_roll = val.split(",")[0].strip()
                        break

            start_line_idx = next(i for i, line in enumerate(lines) if "FileID" in line)
            df = pd.read_csv(io.StringIO("\n".join(lines[start_line_idx:])), skipinitialspace=True, dtype=str)
            df.columns = [str(c).replace('"', '').strip() for c in df.columns]
            df = df.map(lambda x: str(x).replace('"', '').strip() if isinstance(x, str) and str(x).lower() != 'nan' else "")

            if list_num == 1:
                self.df1 = df
                self.sound_roll_var.set(found_roll)
                self.identify_active_tracks(df)
            else:
                self.df2 = df
                self.transmitter_source_dir = os.path.dirname(path)
                self.update_tree(self.tree2, df, ["TimeCode", "UserBits", "FileName"])
                self.transmitter_frame.config(text=f"Transmitter Files (Count: {len(df)})")
            
            self.check_ready_state()
        except Exception as e:
            messagebox.showerror("Error", f"Process Error: {e}")

    def identify_active_tracks(self, df):
        track_cols = [c for c in df.columns if (c.startswith('Name') or c.startswith('Track')) and c != "Tracks"]
        self.active_track_cols = []
        self.track_mapping = {}
        display_list = ["Select Track"]
        
        for col in track_cols:
            names_in_track = [n for n in df[col].unique() if n != ""]
            if not names_in_track: continue 
            self.active_track_cols.append(col)
            display_name = names_in_track[0] if len(names_in_track) == 1 else col
            if display_name in self.track_mapping: display_name = f"{display_name} ({col})"
            self.track_mapping[display_name] = col
            display_list.append(display_name)
            
        self.track_select_combo['values'] = display_list
        self.track_select_combo.current(0) 
        self.refresh_recorder_list()

    def refresh_recorder_list(self, event=None):
        if self.df1 is None: return
        display_selection = self.track_select_combo.get()
        self.tree1.delete(*self.tree1.get_children())
        self.filtered_df1_indices = []
        
        if display_selection == "Select Track":
            for idx, row in self.df1.iterrows():
                names = ", ".join([row.get(c, "") for c in self.active_track_cols if row.get(c, "")])
                vals = [row.get('TimeCode'), row.get('UserBits'), row.get('Scene'), row.get('Take'), names]
                self.tree1.insert("", "end", values=vals, iid=str(idx))
            self.recorder_frame.config(text=f"Recorder Metadata (Total Takes: {len(self.df1)})")
        else:
            selected_col = self.track_mapping.get(display_selection)
            count = 0
            for idx, row in self.df1.iterrows():
                char_name = row.get(selected_col, "")
                if char_name != "":
                    vals = [row.get('TimeCode'), row.get('UserBits'), row.get('Scene'), row.get('Take'), char_name]
                    self.tree1.insert("", "end", values=vals, iid=str(idx))
                    self.filtered_df1_indices.append(idx)
                    count += 1
            self.recorder_frame.config(text=f"Recorder Metadata ({display_selection} Count: {count})")
        
        self.check_ready_state()

    def update_tree(self, tree, df, cols):
        tree.delete(*tree.get_children())
        for idx, row in df.iterrows():
            tree.insert("", "end", values=[row.get(c, "") for c in cols], iid=str(idx))

    def tc_to_sec(self, tc, fps=23.98):
        try:
            parts = re.split('[:;.]', str(tc))
            h, m, s, f = map(float, parts)
            return (h * 3600) + (m * 60) + s + (f / fps)
        except: return 0

    def compare_and_preview(self):
        if self.df1 is None or self.df2 is None: return
        self.matched_pairs = []
        self.preview_tree.delete(*self.preview_tree.get_children())
        for tree in [self.tree1, self.tree2]:
            for item in tree.get_children(): tree.item(item, tags=())

        track_display = self.track_select_combo.get()
        selected_col = self.track_mapping.get(track_display)
        match_count = 0
        offset_limit = self.time_offset_var.get()
        
        first_match_r_idx = None
        first_match_t_idx = None

        for i in self.filtered_df1_indices:
            r1 = self.df1.loc[i]
            tc1 = self.tc_to_sec(r1.get('TimeCode'), float(r1.get('FrameRate', 23.98)))
            ub1 = str(r1.get('UserBits'))
            for j, r2 in self.df2.iterrows():
                tc2 = self.tc_to_sec(r2.get('TimeCode'), float(r2.get('FrameRate', 23.98)))
                if ub1 == str(r2.get('UserBits')) and abs(tc1 - tc2) <= offset_limit:
                    self.matched_pairs.append((i, j))
                    
                    if first_match_r_idx is None:
                        first_match_r_idx = self.tree1.index(str(i))
                        first_match_t_idx = self.tree2.index(str(j))

                    tag = f'match_{match_count % len(self.match_colors)}'
                    if self.tree1.exists(str(i)): self.tree1.item(str(i), tags=(tag,))
                    if self.tree2.exists(str(j)): self.tree2.item(str(j), tags=(tag,))
                    char = r1.get(selected_col)
                    new_name = f"{r1.get('Scene')}-T{r1.get('Take')}-{char}{os.path.splitext(r2.get('FileName'))[1]}"
                    self.preview_tree.insert("", "end", values=(r2.get('FileName'), new_name, f"{match_count+1}", f"{abs(tc1-tc2):.2f}s"))
                    match_count += 1
        
        # PIN TO TOP
        if first_match_r_idx is not None:
            total_r = len(self.tree1.get_children())
            if total_r > 0: self.tree1.yview_moveto(first_match_r_idx / total_r)
        
        if first_match_t_idx is not None:
            total_t = len(self.tree2.get_children())
            if total_t > 0: self.tree2.yview_moveto(first_match_t_idx / total_t)

        self.preview_frame.config(text=f"Proposed Filename Changes (Matches Found: {match_count})")
        self.rename_btn.config(state="normal" if self.matched_pairs else "disabled")
        messagebox.showinfo("Done", f"Found {len(self.matched_pairs)} matches for: {track_display}")

    def execute_rename(self):
        if not self.transmitter_source_dir or not os.path.exists(self.transmitter_source_dir):
            messagebox.showerror("Error", "Source folder not found. Please reload Transmitter CSV.")
            return

        roll = self.sound_roll_var.get()
        selected_col = self.track_mapping.get(self.track_select_combo.get())
        names_list = []
        for m1, _ in self.matched_pairs:
            n = str(self.df1.loc[m1].get(selected_col, "")).strip()
            if n and n not in names_list: names_list.append(n)
        
        char_segment = "-".join(names_list) if names_list else "Unknown"
        clean_segment = re.sub(r'[^\w\-]', '-', char_segment)
        folder_name = f"SR{roll}_{clean_segment}"

        dest_parent = filedialog.askdirectory(title=f"DESTINATION: Create Folder [{folder_name}] In:", initialdir=self.downloads_dir)
        if not dest_parent: return

        target_dir = os.path.join(dest_parent, folder_name)
        if not os.path.exists(target_dir): os.makedirs(target_dir)

        report_data = []
        count = 0
        for m1, m2 in self.matched_pairs:
            r1, r2 = self.df1.loc[m1], self.df2.loc[m2]
            char = r1.get(selected_col)
            orig_filename = str(r2.get('FileName', "")).strip()
            ext = os.path.splitext(orig_filename)[1]
            new_filename = f"{r1.get('Scene')}-T{r1.get('Take')}-{char}{ext}"
            
            src_path = os.path.join(self.transmitter_source_dir, orig_filename)
            dst_path = os.path.join(target_dir, new_filename)

            if os.path.exists(src_path):
                try:
                    shutil.copy2(src_path, dst_path)
                    count += 1
                    tc1 = self.tc_to_sec(r1.get('TimeCode'), float(r1.get('FrameRate', 23.98)))
                    tc2 = self.tc_to_sec(r2.get('TimeCode'), float(r2.get('FrameRate', 23.98)))
                    offset = f"{abs(tc1 - tc2):.2f}s"
                    report_data.append([new_filename, orig_filename, r2.get('TimeCode'), r2.get('UserBits'), offset])
                except: pass

        if report_data and FPDF:
            self.generate_pdf_report(target_dir, folder_name, report_data)
        elif not FPDF:
            messagebox.showwarning("Missing Library", "fpdf library not found. Skipping PDF generation.")

        messagebox.showinfo("Complete", f"Copied {count} files to:\n{target_dir}")
        self.rename_btn.config(state="disabled")

    def generate_pdf_report(self, target_dir, folder_name, data):
        try:
            pdf = FPDF(orientation='L', unit='mm', format='A4')
            pdf.add_page()
            pdf.set_font("helvetica", 'B', 16)
            pdf.cell(0, 10, f"Sound Report for '{folder_name}'", ln=True, align='C')
            pdf.ln(5)
            pdf.set_font("helvetica", 'B', 10)
            pdf.set_fill_color(220, 220, 220)
            pdf.cell(75, 8, "New Filename", 1, 0, 'C', 1)
            pdf.cell(75, 8, "Old Filename", 1, 0, 'C', 1)
            pdf.cell(40, 8, "Transmitter TC", 1, 0, 'C', 1)
            pdf.cell(40, 8, "Userbits", 1, 0, 'C', 1)
            pdf.cell(30, 8, "Offset", 1, 1, 'C', 1)
            pdf.set_font("helvetica", '', 9)
            for row in data:
                pdf.cell(75, 7, str(row[0]), 1)
                pdf.cell(75, 7, str(row[1]), 1)
                pdf.cell(40, 7, str(row[2]), 1, 0, 'C')
                pdf.cell(40, 7, str(row[3]), 1, 0, 'C')
                pdf.cell(30, 7, str(row[4]), 1, 1, 'C')
            report_path = os.path.join(target_dir, f"{folder_name}.pdf")
            pdf.output(report_path)
        except Exception as e:
            print(f"PDF Error: {e}")

if __name__ == "__main__":
    root = tk.Tk(); app = AudioRenamerApp(root); root.mainloop()