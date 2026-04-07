from __future__ import annotations

import os
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from queue import Empty, Queue

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from app_metadata import APP_NAME, APP_TITLE, APP_VERSION, SUPPORT_EMAIL
from budget_automation import (
    SUPPORTED_TARGET_EXTENSIONS,
    BudgetAutomator,
    BudgetAutomationError,
    ImportSummary,
)

COMPASS_DOWNLOAD_TEXT = """How to download the Expense Sub-Program Mastersheet from Compass

1. Open Compass and go to Financial Management.
2. In the Financial Management Dashboard, select the budget year you want in Selected Financial Period.
3. Click Financial Periods.
4. On the Financial Periods page, find the same budget year.
5. Open the Budget Reports drop-down for that year and select Expense Sub-Program Mastersheet.
6. Compass will download the file. Save it, then use that downloaded file as the Expense Sub-Program file in this app.

Screenshots
- Screenshot 1 shows where to open Financial Periods from the Financial Management Dashboard.
- Screenshot 2 shows where to select Expense Sub-Program Mastersheet from the Budget Reports drop-down.
"""

HELP_TEXT = f"""Master Budget Automation Tool v{APP_VERSION} - User Instructions

What this app does
- Imports data from the Expense Sub-Program file into the Master Budget workbook.
- Saves the result as a new output workbook.
- Keeps the original template unchanged.
- Preserves original macro button bindings on Windows when Microsoft Excel is installed.
- Highlights mismatch items for review.

Before you start
1. Close the workbook in Excel before running the app.
2. Keep the original template as a separate file.
3. Choose a new file name for the output workbook.

How to use the app
1. Click Browse next to Expense Sub-Program file and select the source file.
2. Click Browse next to Master Budget template and select the original budget template.
3. Click Browse next to Output workbook and choose a new file name and location.
4. Click Generate budget workbook.
5. Wait for the progress bar to finish.
6. Open the output workbook in Excel and review the imported data.

Buttons
- Generate budget workbook: Runs the import.
- Create suggested output name: Creates a new output file name based on the template name and time.
- Open output folder: Opens the folder where the output workbook will be saved.
- Instructions: Opens this guide.
- Clear: Clears the file paths and run summary.

How mismatch handling works
- Account mismatch checking only tracks 5-digit account codes starting with 7 or 8.
- Source-only account codes and source-only sub-program codes are inserted into Master in numeric order.
- Source-only inserted items are highlighted light green.
- Rows at and below 26201 Asset Clearing Account are not changed by the import logic.

Important notes
- Do not choose the original template file as the output file.
- If the template does not contain a Compass sheet, the app still imports directly into Master.
- Keep Excel installed on Windows if you want original macro button bindings preserved.

Troubleshooting
- If the app cannot save the workbook, make sure the template and output files are not open in Excel.
- If the output file already exists and is open, close it and run again.
- If the output looks incomplete, check the Run summary for mismatch messages.

Suggestions
Please send suggestions to {SUPPORT_EMAIL}
"""


def resource_path(relative_path: str) -> Path:
    base_path = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
    return base_path / relative_path


class BudgetAutomationApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry('900x760')
        self.root.minsize(820, 660)

        self.source_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar(value='Ready')
        self.progress_text_var = tk.StringVar(value='Idle')
        self.progress_value = tk.DoubleVar(value=0)

        self.automator = BudgetAutomator()
        self.is_running = False
        self.result_queue: Queue = Queue()
        self.instructions_window: tk.Toplevel | None = None
        self._build_ui()
        self.root.after(200, self._poll_queue)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text=APP_TITLE, font=('Segoe UI', 16, 'bold')).grid(
            row=0, column=0, columnspan=3, sticky='w'
        )
        ttk.Label(
            frame,
            text=(
                'Use this tool to import an Expense Sub-Program export into the Master Budget workbook '
                'without using OFFSET/MATCH formulas.'
            ),
            wraplength=820,
            justify='left',
        ).grid(row=1, column=0, columnspan=3, sticky='w', pady=(8, 20))

        self._add_file_row(frame, 2, 'Expense Sub-Program file', self.source_var, self._browse_source)
        self._add_file_row(frame, 3, 'Master Budget template', self.template_var, self._browse_template)
        self._add_file_row(frame, 4, 'Output workbook', self.output_var, self._browse_output)

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=5, column=0, columnspan=3, sticky='w', pady=(10, 8))

        self.run_button = ttk.Button(button_frame, text='Generate budget workbook', command=self._run)
        self.run_button.pack(side='left')
        self.suggest_button = ttk.Button(button_frame, text='Create suggested output name', command=self._suggest_output)
        self.suggest_button.pack(side='left', padx=(10, 0))
        self.open_folder_button = ttk.Button(button_frame, text='Open output folder', command=self._open_output_folder)
        self.open_folder_button.pack(side='left', padx=(10, 0))
        self.instructions_button = ttk.Button(button_frame, text='Instructions', command=self._show_instructions)
        self.instructions_button.pack(side='left', padx=(10, 0))
        self.clear_button = ttk.Button(button_frame, text='Clear', command=self._clear)
        self.clear_button.pack(side='left', padx=(10, 0))

        self.banner = tk.Label(
            frame,
            text='Waiting to run.',
            anchor='w',
            justify='left',
            padx=12,
            pady=10,
            relief='groove',
            bg='#f3f3f3',
        )
        self.banner.grid(row=6, column=0, columnspan=3, sticky='ew', pady=(0, 10))

        ttk.Label(frame, textvariable=self.progress_text_var).grid(row=7, column=0, columnspan=3, sticky='w')
        self.progress_bar = ttk.Progressbar(
            frame,
            orient='horizontal',
            mode='determinate',
            variable=self.progress_value,
            maximum=100,
        )
        self.progress_bar.grid(row=8, column=0, columnspan=3, sticky='ew', pady=(6, 2))

        ttk.Label(frame, text='Run summary').grid(row=10, column=0, sticky='w')

        self.log = tk.Text(frame, height=18, wrap='word', font=('Consolas', 10))
        self.log.grid(row=11, column=0, columnspan=3, sticky='nsew', pady=(6, 12))
        self.log.tag_configure('heading', font=('Consolas', 10, 'bold'))
        self.log.tag_configure('success', foreground='#0b6e0b')
        self.log.tag_configure('warning', foreground='#9a6700')
        self.log.tag_configure('danger', foreground='#b00020')
        self.log.tag_configure('ok', foreground='#0b6e0b')
        self.log.tag_configure('extra', foreground='#2e7d32')
        self.log.tag_configure('muted', foreground='#555555')

        ttk.Label(
            frame,
            text=f'Please send suggestions to {SUPPORT_EMAIL}',
            foreground='#555555',
        ).grid(row=13, column=0, columnspan=3, sticky='w', pady=(8, 0))

        ttk.Label(frame, textvariable=self.status_var).grid(row=12, column=0, columnspan=3, sticky='w')

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(11, weight=1)

    def _set_controls_enabled(self, enabled: bool) -> None:
        state = 'normal' if enabled else 'disabled'
        for widget in [self.run_button, self.suggest_button, self.open_folder_button, self.clear_button]:
            widget.configure(state=state)
        self.instructions_button.configure(state='normal')

    def _add_file_row(self, parent, row: int, label: str, variable: tk.StringVar, browse_command) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky='w', pady=6)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky='ew', padx=(12, 12), pady=6)
        ttk.Button(parent, text='Browse', command=browse_command).grid(row=row, column=2, sticky='ew', pady=6)

    def _browse_source(self) -> None:
        path = filedialog.askopenfilename(
            title='Select Expense Sub-Program file',
            filetypes=[('Supported files', '*.csv *.xlsx *.xlsm'), ('All files', '*.*')],
        )
        if path:
            self.source_var.set(path)
            if not self.output_var.get() and self.template_var.get():
                self._suggest_output()

    def _browse_template(self) -> None:
        path = filedialog.askopenfilename(
            title='Select Master Budget workbook',
            filetypes=[('Excel files', '*.xlsm *.xlsx'), ('All files', '*.*')],
        )
        if path:
            self.template_var.set(path)
            if not self.output_var.get():
                self._suggest_output()

    def _browse_output(self) -> None:
        default_extension = self._preferred_output_suffix()
        path = filedialog.asksaveasfilename(
            title='Save output workbook as',
            defaultextension=default_extension,
            filetypes=self._output_filetypes(default_extension),
        )
        if path:
            self.output_var.set(path)

    def _preferred_output_suffix(self) -> str:
        template = self.template_var.get().strip()
        if template:
            suffix = Path(template).suffix.lower()
            if suffix in SUPPORTED_TARGET_EXTENSIONS:
                return suffix
        return '.xlsm'

    @staticmethod
    def _output_filetypes(preferred_suffix: str) -> list[tuple[str, str]]:
        if preferred_suffix == '.xlsx':
            return [('Excel Workbook', '*.xlsx'), ('Excel Macro-Enabled Workbook', '*.xlsm')]
        return [('Excel Macro-Enabled Workbook', '*.xlsm'), ('Excel Workbook', '*.xlsx')]

    def _suggest_output(self) -> None:
        template = self.template_var.get().strip()
        if not template:
            messagebox.showinfo('Select template first', 'Please choose the Master Budget workbook first.')
            return
        template_path = Path(template)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        suggested = template_path.with_name(f'{template_path.stem}_AUTO_{timestamp}{template_path.suffix}')
        self.output_var.set(str(suggested))

    def _open_output_folder(self) -> None:
        output = self.output_var.get().strip()
        if not output:
            messagebox.showinfo('No output folder yet', 'Please choose or generate an output workbook path first.')
            return
        folder = Path(output).expanduser().resolve().parent
        folder.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith('win'):
                os.startfile(str(folder))
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', str(folder)])
            else:
                subprocess.Popen(['xdg-open', str(folder)])
        except Exception as exc:
            messagebox.showerror('Could not open folder', str(exc))

    def _load_instruction_image(self, relative_path: str, max_width: int = 700) -> tk.PhotoImage | None:
        image_path = resource_path(relative_path)
        if not image_path.exists():
            return None
        try:
            image = tk.PhotoImage(file=str(image_path))
            width = image.width()
            if width > max_width:
                factor = max(1, (width + max_width - 1) // max_width)
                image = image.subsample(factor, factor)
            return image
        except Exception:
            return None

    def _show_instructions(self) -> None:
        if self.instructions_window is not None and self.instructions_window.winfo_exists():
            self.instructions_window.deiconify()
            self.instructions_window.lift()
            self.instructions_window.focus_force()
            return

        window = tk.Toplevel(self.root)
        window.title('How to use the app')
        window.geometry('780x700')
        window.minsize(660, 520)
        window.transient(self.root)
        self.instructions_window = window

        container = ttk.Frame(window, padding=16)
        container.pack(fill='both', expand=True)

        ttk.Label(container, text='How to use the app', font=('Segoe UI', 14, 'bold')).pack(anchor='w')
        ttk.Label(
            container,
            text='This guide is also included in README.md inside the app folder.',
            foreground='#555555',
        ).pack(anchor='w', pady=(4, 10))

        text_frame = ttk.Frame(container)
        text_frame.pack(fill='both', expand=True)

        scrollbar = ttk.Scrollbar(text_frame, orient='vertical')
        text_widget = tk.Text(text_frame, wrap='word', font=('Segoe UI', 10), yscrollcommand=scrollbar.set)
        scrollbar.config(command=text_widget.yview)
        scrollbar.pack(side='right', fill='y')
        text_widget.pack(side='left', fill='both', expand=True)

        image_refs: list[tk.PhotoImage] = []
        text_widget.insert('1.0', COMPASS_DOWNLOAD_TEXT)

        screenshot_1 = self._load_instruction_image('assets/compass_step1.png')
        if screenshot_1 is not None:
            image_refs.append(screenshot_1)
            text_widget.insert(tk.END, '\nScreenshot 1\n')
            text_widget.image_create(tk.END, image=screenshot_1)
            text_widget.insert(tk.END, '\n\n')

        screenshot_2 = self._load_instruction_image('assets/compass_step2.png')
        if screenshot_2 is not None:
            image_refs.append(screenshot_2)
            text_widget.insert(tk.END, 'Screenshot 2\n')
            text_widget.image_create(tk.END, image=screenshot_2)
            text_widget.insert(tk.END, '\n\n')

        text_widget.insert(tk.END, HELP_TEXT)
        text_widget.configure(state='disabled')
        window._instruction_images = image_refs

        button_row = ttk.Frame(container)
        button_row.pack(fill='x', pady=(12, 0))
        ttk.Button(button_row, text='Close', command=window.destroy).pack(side='right')

        def on_close() -> None:
            self.instructions_window = None
            window.destroy()

        window.protocol('WM_DELETE_WINDOW', on_close)

    def _clear(self) -> None:
        if self.is_running:
            return
        self.source_var.set('')
        self.template_var.set('')
        self.output_var.set('')
        self.log.delete('1.0', tk.END)
        self._set_banner('Waiting to run.', 'neutral')
        self.status_var.set('Ready')
        self.progress_text_var.set('Idle')
        self.progress_value.set(0)

    def _set_banner(self, text: str, level: str) -> None:
        colors = {
            'neutral': ('#f3f3f3', '#222222'),
            'success': ('#e8f5e9', '#0b6e0b'),
            'warning': ('#fff7e6', '#9a6700'),
            'error': ('#fdecea', '#b00020'),
        }
        bg, fg = colors.get(level, colors['neutral'])
        self.banner.configure(text=text, bg=bg, fg=fg)

    @staticmethod
    def _issues_count(summary: ImportSummary) -> int:
        return (
            len(summary.missing_master_codes)
            + len(summary.missing_source_codes)
            + len(summary.missing_subprogram_codes)
            + len(summary.source_extra_subprogram_codes)
        )

    def _append_line(self, text: str = '', tag: str | None = None) -> None:
        start = self.log.index(tk.END)
        self.log.insert(tk.END, text + '\n')
        if tag:
            end = self.log.index(tk.END)
            self.log.tag_add(tag, start, end)

    def _append_issue_block(
        self,
        title: str,
        values: list[str],
        description: str = '',
        issue_tag: str = 'danger',
        ok_tag: str = 'ok',
        issue_label: str = 'MISMATCH',
    ) -> None:
        if values:
            self._append_line(f'[{issue_label}] {title}', issue_tag)
            if description:
                self._append_line(description, 'muted')
            for value in values:
                self._append_line(f'- {value}', issue_tag)
        else:
            self._append_line(f'[OK] {title}', ok_tag)
            if description:
                self._append_line(description, 'muted')
            self._append_line('None')
        self._append_line()

    def _render_summary(self, summary: ImportSummary) -> None:
        self.log.delete('1.0', tk.END)
        issues_count = self._issues_count(summary)
        self._append_line('Run result', 'heading')
        if issues_count == 0:
            self._append_line('No mismatch or concern found. Import completed successfully.', 'success')
        else:
            self._append_line(
                f'Import completed with {issues_count} mismatch item(s). Please review highlighted items below.',
                'warning',
            )
        self._append_line()
        self._append_line(f'Output workbook: {summary.output_workbook}')
        self._append_line(f'Report file: {summary.report_file}')
        self._append_line(f'Matched rows: {summary.matched_rows}')
        self._append_line(f'Matched cells: {summary.matched_cells}')
        self._append_line()
        self._append_issue_block(
            'Master account codes missing from source',
            summary.missing_master_code_details,
            'Tracked account mismatch check only includes 5-digit account codes starting with 7 or 8.',
            issue_tag='danger',
            ok_tag='ok',
            issue_label='MISMATCH',
        )
        self._append_issue_block(
            'Source account codes not used by Master',
            summary.missing_source_code_details,
            'These tracked source account codes exist in the source file but not on Master. They are added into the Master sheet in numeric position with descriptions and light green highlighting.',
            issue_tag='extra',
            ok_tag='ok',
            issue_label='SOURCE ONLY',
        )
        self._append_issue_block(
            'Master sub-program codes missing from source',
            summary.missing_subprogram_details,
            'These sub-program codes exist on Master but were not found in the source file.',
            issue_tag='danger',
            ok_tag='ok',
            issue_label='MISMATCH',
        )
        self._append_issue_block(
            'Source sub-program codes not used by Master',
            summary.source_extra_subprogram_details,
            'These source sub-program codes do not exist on Master. They are added into the Master sheet in numeric position with descriptions and light green highlighting.',
            issue_tag='extra',
            ok_tag='ok',
            issue_label='SOURCE ONLY',
        )

    def _run(self) -> None:
        if self.is_running:
            return
        source = self.source_var.get().strip()
        template = self.template_var.get().strip()
        output = self.output_var.get().strip()
        if not source or not template or not output:
            messagebox.showwarning('Missing information', 'Please select the source file, template workbook, and output file.')
            return

        self.is_running = True
        self._set_controls_enabled(False)
        self.status_var.set('Running import...')
        self._set_banner('Running import. Please wait...', 'neutral')
        self.progress_text_var.set('Starting...')
        self.progress_value.set(0)
        self.log.delete('1.0', tk.END)
        self._append_line('Import started...')
        self._append_line('Tip: the app stays responsive while the workbook is being processed.')
        self._append_line()

        def progress_update(percent: int, message: str) -> None:
            self.result_queue.put(('progress', (percent, message)))

        def worker() -> None:
            try:
                summary = self.automator.run(source, template, output, progress_callback=progress_update)
                self.result_queue.put(('success', summary))
            except BudgetAutomationError as exc:
                self.result_queue.put(('budget_error', str(exc)))
            except Exception as exc:
                self.result_queue.put(('unexpected_error', str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self.result_queue.get_nowait()
                if kind == 'progress':
                    percent, message = payload
                    self.progress_value.set(percent)
                    self.progress_text_var.set(f'{percent}% - {message}')
                    self.status_var.set(message)
                    continue

                self.is_running = False
                self._set_controls_enabled(True)
                self.progress_value.set(100 if kind == 'success' else self.progress_value.get())
                self.progress_text_var.set('Completed.' if kind == 'success' else 'Stopped.')
        
                if kind == 'success':
                    summary: ImportSummary = payload
                    self._render_summary(summary)
                    issues_count = self._issues_count(summary)
                    if issues_count == 0:
                        self.status_var.set('Completed - no issues found')
                        self._set_banner('Completed successfully. No mismatch or concern found.', 'success')
                        messagebox.showinfo(
                            'Completed - No issues found',
                            'The budget workbook has been created successfully.\n\nNo mismatch or concern was found.\n\n'
                            f'Saved to:\n{summary.output_workbook}',
                        )
                    else:
                        self.status_var.set(f'Completed with {issues_count} mismatch item(s)')
                        self._set_banner(
                            f'Completed with {issues_count} mismatch item(s). Highlighted rows and columns need review.',
                            'warning',
                        )
                        messagebox.showwarning(
                            'Completed with mismatches',
                            'The budget workbook was created, but mismatches were found.\n\n'
                            'Highlighted rows and columns are shown in the output workbook when available. '
                            'The same issues are also listed in the Run summary and report file.\n\n'
                            f'Saved to:\n{summary.output_workbook}',
                        )
                elif kind == 'budget_error':
                    self.status_var.set('Error')
                    self._set_banner(f'Error: {payload}', 'error')
                    messagebox.showerror('Budget automation error', payload)
                else:
                    self.status_var.set('Error')
                    self._set_banner(f'Unexpected error: {payload}', 'error')
                    messagebox.showerror('Unexpected error', payload)
        except Empty:
            pass
        self.root.after(200, self._poll_queue)


if __name__ == '__main__':
    root = tk.Tk()
    style = ttk.Style()
    if 'vista' in style.theme_names():
        style.theme_use('vista')
    BudgetAutomationApp(root)
    root.mainloop()
