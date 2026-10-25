"""Desktop GUI for Viedoc EDC QC Automation.

Aligned with SOI-EDC-QC-001 v2.0: defaults to 'Default Baseline Diff / Study Amendment'
mode and retains 'Standalone Initial Setup (Exception)' for non-standard studies.
"""

from datetime import datetime
import os
from pathlib import Path
import subprocess
import sys
import threading
from typing import Optional
import customtkinter as ctk
from tkinter import filedialog, messagebox
import openpyxl

from src.config_diff import generate_configuration_report_diff
from src.delta_engine import generate_amendment_qc_form
from src.generator import find_template_path, generate_initial_qc_forms
from src.parser import generate_output_filename, parse_viedoc_report

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

MODE_AMENDMENT = "Default Baseline Diff / Study Amendment"
MODE_STANDALONE = "Standalone Initial Setup (Exception)"
MODE_CONFIG_DIFF = "Configuration Report Diff"


def get_default_templates_dir() -> Path:
    """Finds the templates directory either next to the script or in PyInstaller bundle."""
    if getattr(sys, "frozen", False):
        base_dir = Path(sys._MEIPASS)
        bundle_templates = base_dir / "templates"
        if bundle_templates.exists():
            return bundle_templates

    local_dir = Path(__file__).resolve().parent.parent / "templates"
    if local_dir.exists():
        return local_dir

    cwd_templates = Path.cwd() / "templates"
    if cwd_templates.exists():
        return cwd_templates

    argv_templates = Path(sys.argv[0]).resolve().parent / "templates"
    if argv_templates.exists():
        return argv_templates

    return local_dir


class QCAutomationApp(ctk.CTk):
    """Main Desktop Application Window."""

    def __init__(self):
        super().__init__()

        self.title("Viedoc EDC QC Automation")
        self.geometry("820x680")
        self.minsize(760, 600)

        self.templates_dir = get_default_templates_dir()
        self.last_output_dir: Optional[Path] = None

        self._build_ui()
        # Default to Baseline Diff / Study Amendment per SOI Section 2.2
        self._on_mode_change(MODE_AMENDMENT)

    def _build_ui(self):
        header_frame = ctk.CTkFrame(self, corner_radius=10)
        header_frame.pack(fill="x", padx=15, pady=(15, 10))

        title_label = ctk.CTkLabel(
            header_frame,
            text="Viedoc EDC QC Automation",
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        title_label.pack(anchor="w", padx=15, pady=(12, 2))

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Risk-based QC generation & baseline/amendment diffing (SOI-EDC-QC-001 v2.0)",
            font=ctk.CTkFont(size=13),
            text_color="gray",
        )
        subtitle_label.pack(anchor="w", padx=15, pady=(0, 12))

        mode_frame = ctk.CTkFrame(self, corner_radius=10)
        mode_frame.pack(fill="x", padx=15, pady=5)

        mode_title = ctk.CTkLabel(
            mode_frame, text="Select Operation Mode:", font=ctk.CTkFont(size=14, weight="bold")
        )
        mode_title.pack(anchor="w", padx=15, pady=(10, 5))

        self.mode_switch = ctk.CTkSegmentedButton(
            mode_frame,
            values=[MODE_AMENDMENT, MODE_STANDALONE, MODE_CONFIG_DIFF],
            command=self._on_mode_change,
            height=32,
        )
        self.mode_switch.set(MODE_AMENDMENT)
        self.mode_switch.pack(fill="x", padx=15, pady=(0, 12))

        self.input_frame = ctk.CTkFrame(self, corner_radius=10)
        self.input_frame.pack(fill="x", padx=15, pady=5)

        self.report_label = ctk.CTkLabel(
            self.input_frame,
            text="New Viedoc Configuration Report (.xlsx):",
            font=ctk.CTkFont(weight="bold"),
        )
        self.report_label.grid(row=0, column=0, sticky="w", padx=15, pady=(10, 2))

        self.report_entry = ctk.CTkEntry(
            self.input_frame, placeholder_text="Select Viedoc Configuration Report..."
        )
        self.report_entry.grid(row=1, column=0, sticky="ew", padx=(15, 5), pady=(0, 8))

        self.report_btn = ctk.CTkButton(
            self.input_frame, text="Browse...", width=90, command=self._browse_report
        )
        self.report_btn.grid(row=1, column=1, padx=(5, 15), pady=(0, 8))

        self.old_qc_label = ctk.CTkLabel(
            self.input_frame,
            text="Baseline / Previous QC Workbook (.xlsx / .xlsm):",
            font=ctk.CTkFont(weight="bold"),
        )
        self.old_qc_entry = ctk.CTkEntry(
            self.input_frame,
            placeholder_text="Select Default Study QC workbook (SharePoint) or prior study version...",
        )
        self.old_qc_btn = ctk.CTkButton(
            self.input_frame, text="Browse...", width=90, command=self._browse_old_qc
        )

        self.out_label = ctk.CTkLabel(
            self.input_frame,
            text="Output Directory:",
            font=ctk.CTkFont(weight="bold"),
        )
        default_out = str(Path(r"d:\QCproject\output"))
        self.out_entry = ctk.CTkEntry(self.input_frame, placeholder_text=default_out)
        self.out_entry.insert(0, default_out)

        self.out_btn = ctk.CTkButton(
            self.input_frame, text="Browse...", width=90, command=self._browse_out_dir
        )

        self.input_frame.columnconfigure(0, weight=1)

        action_frame = ctk.CTkFrame(self, corner_radius=10)
        action_frame.pack(fill="x", padx=15, pady=5)

        btn_box = ctk.CTkFrame(action_frame, fg_color="transparent")
        btn_box.pack(fill="x", padx=15, pady=(10, 8))

        self.run_btn = ctk.CTkButton(
            btn_box,
            text="Run Baseline / Amendment Diff & Generate QC",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=38,
            command=self._start_generation,
        )
        self.run_btn.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.open_folder_btn = ctk.CTkButton(
            btn_box,
            text="Open Output Folder",
            font=ctk.CTkFont(size=13),
            height=38,
            state="disabled",
            fg_color="gray50",
            command=self._open_output_folder,
        )
        self.open_folder_btn.pack(side="right", padx=(8, 0))

        self.progress_bar = ctk.CTkProgressBar(action_frame)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=15, pady=(0, 10))

        log_frame = ctk.CTkFrame(self, corner_radius=10)
        log_frame.pack(fill="both", expand=True, padx=15, pady=(5, 15))

        log_title = ctk.CTkLabel(
            log_frame, text="Execution Log:", font=ctk.CTkFont(size=12, weight="bold")
        )
        log_title.pack(anchor="w", padx=15, pady=(8, 2))

        self.log_textbox = ctk.CTkTextbox(
            log_frame, font=ctk.CTkFont(family="Consolas", size=11)
        )
        self.log_textbox.pack(fill="both", expand=True, padx=15, pady=(0, 10))

        self.log("Ready. Select mode and files to begin.")

    def log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", f"[{timestamp}] {message}\n")
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def _on_mode_change(self, mode: str):
        if mode == MODE_STANDALONE:
            self.report_label.configure(text="Viedoc Configuration Report (.xlsx):")
            self.old_qc_label.grid_forget()
            self.old_qc_entry.grid_forget()
            self.old_qc_btn.grid_forget()
            self.out_label.grid(row=2, column=0, sticky="w", padx=15, pady=(5, 2))
            self.out_entry.grid(row=3, column=0, sticky="ew", padx=(15, 5), pady=(0, 12))
            self.out_btn.grid(row=3, column=1, padx=(5, 15), pady=(0, 12))
            self.run_btn.configure(text="Generate Standalone Initial QC Forms (Both)")
        elif mode == MODE_CONFIG_DIFF:
            self.report_label.configure(text="New Viedoc Configuration Report (.xlsx):")
            self.old_qc_label.configure(text="Prior / Baseline Configuration Report (.xlsx):")
            self.old_qc_entry.configure(placeholder_text="Select prior Viedoc Configuration Report to compare...")
            self.old_qc_label.grid(row=2, column=0, sticky="w", padx=15, pady=(5, 2))
            self.old_qc_entry.grid(row=3, column=0, sticky="ew", padx=(15, 5), pady=(0, 8))
            self.old_qc_btn.grid(row=3, column=1, padx=(5, 15), pady=(0, 8))
            self.out_label.grid(row=4, column=0, sticky="w", padx=15, pady=(5, 2))
            self.out_entry.grid(row=5, column=0, sticky="ew", padx=(15, 5), pady=(0, 12))
            self.out_btn.grid(row=5, column=1, padx=(5, 15), pady=(0, 12))
            self.run_btn.configure(text="Run Configuration Report Diff")
        else:
            self.report_label.configure(text="New Viedoc Configuration Report (.xlsx):")
            self.old_qc_label.configure(text="Baseline / Previous QC Workbook (.xlsx / .xlsm):")
            self.old_qc_entry.configure(placeholder_text="Select Default Study QC workbook (SharePoint) or prior study version...")
            self.old_qc_label.grid(row=2, column=0, sticky="w", padx=15, pady=(5, 2))
            self.old_qc_entry.grid(row=3, column=0, sticky="ew", padx=(15, 5), pady=(0, 8))
            self.old_qc_btn.grid(row=3, column=1, padx=(5, 15), pady=(0, 8))
            self.out_label.grid(row=4, column=0, sticky="w", padx=15, pady=(5, 2))
            self.out_entry.grid(row=5, column=0, sticky="ew", padx=(15, 5), pady=(0, 12))
            self.out_btn.grid(row=5, column=1, padx=(5, 15), pady=(0, 12))
            self.run_btn.configure(text="Run Baseline / Amendment Diff & Generate QC")

    def _browse_report(self):
        filename = filedialog.askopenfilename(
            title="Select Viedoc Configuration Report",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
        )
        if filename:
            self.report_entry.delete(0, "end")
            self.report_entry.insert(0, filename)

    def _browse_old_qc(self):
        mode = self.mode_switch.get()
        if mode == MODE_CONFIG_DIFF:
            title = "Select Prior Configuration Report"
            filetypes = [("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        else:
            title = "Select Baseline or Previous QC Workbook"
            filetypes = [("Excel Files", "*.xlsx;*.xlsm"), ("All Files", "*.*")]
        filename = filedialog.askopenfilename(
            title=title,
            filetypes=filetypes,
        )
        if filename:
            self.old_qc_entry.delete(0, "end")
            self.old_qc_entry.insert(0, filename)

    def _browse_out_dir(self):
        folder = filedialog.askdirectory(title="Select Output Directory")
        if folder:
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, folder)

    def _open_output_folder(self):
        if self.last_output_dir and self.last_output_dir.exists():
            if sys.platform == "win32":
                os.startfile(str(self.last_output_dir))
            else:
                subprocess.Popen(["explorer", str(self.last_output_dir)])

    def _start_generation(self):
        report_path_str = self.report_entry.get().strip()
        out_dir_str = self.out_entry.get().strip()

        if not report_path_str:
            messagebox.showwarning("Missing Input", "Please select a Viedoc Configuration Report.")
            return

        report_path = Path(report_path_str)
        if not report_path.exists():
            messagebox.showerror("Error", f"Report file not found:\n{report_path}")
            return

        out_dir = Path(out_dir_str) if out_dir_str else report_path.parent / "output"
        mode = self.mode_switch.get()

        if mode in (MODE_AMENDMENT, MODE_CONFIG_DIFF):
            old_qc_str = self.old_qc_entry.get().strip()
            if not old_qc_str:
                msg = (
                    "Please select the prior Viedoc Configuration Report to compare against."
                    if mode == MODE_CONFIG_DIFF
                    else "Please select the Default Study baseline or previous QC workbook to compare against."
                )
                messagebox.showwarning("Missing Input", msg)
                return
            old_qc_path = Path(old_qc_str)
            if not old_qc_path.exists():
                messagebox.showerror("Error", f"File not found:\n{old_qc_path}")
                return
        else:
            old_qc_path = None

        self.run_btn.configure(state="disabled")
        self.progress_bar.set(0.1)

        threading.Thread(
            target=self._run_process_thread,
            args=(mode, report_path, old_qc_path, out_dir),
            daemon=True,
        ).start()

    def _run_process_thread(
        self,
        mode: str,
        report_path: Path,
        old_qc_path: Optional[Path],
        out_dir: Path,
    ):
        try:
            self.log(f"Starting execution in mode: {mode}")
            self.log(f"Ingesting Viedoc Configuration Report: {report_path.name}")
            parsed = parse_viedoc_report(report_path)
            self.log(
                f"Parsed study: '{parsed.metadata.study_code}', Version: '{parsed.metadata.full_version}'"
            )
            self.progress_bar.set(0.35)

            if mode == MODE_STANDALONE:
                self.log("Generating standalone initial QC workbooks...")
                val_path, design_path = generate_initial_qc_forms(
                    parsed_config=parsed,
                    templates_dir=self.templates_dir,
                    output_dir=out_dir,
                )
                self.progress_bar.set(0.9)
                self.log(f"SUCCESS: Generated Validation QC Form -> {val_path.name}")
                self.log(f"SUCCESS: Generated Design & Spec QC Form -> {design_path.name}")
                self.last_output_dir = out_dir
            elif mode == MODE_CONFIG_DIFF:
                self.log(f"Comparing New Report against Prior Report: {old_qc_path.name}")
                self.progress_bar.set(0.5)
                out_file, diffs = generate_configuration_report_diff(
                    old_report_path=old_qc_path,
                    new_report_path=report_path,
                    output_dir=out_dir,
                )
                self.progress_bar.set(0.9)
                total_mod = sum(1 for d in diffs if d["change_type"] == "MODIFIED")
                total_new = sum(1 for d in diffs if d["change_type"] == "NEW")
                total_rem = sum(1 for d in diffs if d["change_type"] == "REMOVED")
                self.log(
                    f"Diff Summary: {total_new} Added, {total_mod} Modified, {total_rem} Removed across all sheets"
                )
                self.log(f"SUCCESS: Generated Configuration Report Diff -> {out_file.name}")
                self.last_output_dir = out_dir
            else:
                self.log(f"Comparing against baseline/previous QC workbook: {old_qc_path.name}")
                wb_check = openpyxl.load_workbook(str(old_qc_path), read_only=True)
                sheets = wb_check.sheetnames
                wb_check.close()

                if "Viedoc - Data Checks" in sheets or "Viedoc - Alerts" in sheets:
                    form_type = "validation"
                    form_name = "Validation Check QC Form"
                else:
                    form_type = "design"
                    form_name = "Design and Specification QC Form"

                template = find_template_path(self.templates_dir, form_type)
                self.log(f"Detected QC Form Type: {form_name}")
                self.progress_bar.set(0.55)

                out_filename = generate_output_filename(
                    study_code=parsed.metadata.study_code,
                    qc_form_name=form_name,
                    design_version=parsed.metadata.full_version,
                )
                target_out = out_dir / out_filename

                out_file, all_deltas = generate_amendment_qc_form(
                    old_qc_path=old_qc_path,
                    new_config_report_path=report_path,
                    template_path=template,
                    output_path=target_out,
                    qc_form_type=form_type,
                )
                self.progress_bar.set(0.9)

                total_new = sum(len([d for d in deltas if d.change_type == "NEW"]) for deltas in all_deltas.values())
                total_mod = sum(len([d for d in deltas if d.change_type == "MODIFIED"]) for deltas in all_deltas.values())
                total_rem = sum(len([d for d in deltas if d.change_type == "REMOVED"]) for deltas in all_deltas.values())
                total_unch = sum(len([d for d in deltas if d.change_type == "UNCHANGED"]) for deltas in all_deltas.values())

                self.log(
                    f"Delta Summary: {total_new} Added, {total_mod} Modified, {total_rem} Removed, {total_unch} Unchanged"
                )
                self.log(f"SUCCESS: Generated Amendment QC Workbook -> {out_file.name}")
                self.last_output_dir = out_dir

            self.progress_bar.set(1.0)
            self.log("Operation completed successfully!")
            self.after(0, self._on_success)

        except Exception as e:
            self.log(f"ERROR: {str(e)}")
            self.after(0, lambda: messagebox.showerror("Execution Error", str(e)))
            self.after(0, self._on_failure)

    def _on_success(self):
        self.run_btn.configure(state="normal")
        self.open_folder_btn.configure(state="normal", fg_color=["#3B8ED0", "#1F6AA5"])
        messagebox.showinfo(
            "Success",
            "QC Documentation generation completed successfully!\nClick 'Open Output Folder' to view files.",
        )

    def _on_failure(self):
        self.run_btn.configure(state="normal")
        self.progress_bar.set(0)


def launch():
    """Launches the desktop GUI application."""
    app = QCAutomationApp()
    app.mainloop()


if __name__ == "__main__":
    launch()