# -*- coding: utf-8 -*-

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Form,
    Button,
    OpenFileDialog,
    Label,
    Application,
    MessageBox,
    DialogResult,
    Control,
    ToolTip,
    DataGridView,
    DataGridViewTextBoxColumn,
    DataGridViewSelectionMode,
    CheckBox,
    DataGridViewAutoSizeColumnsMode,
    DataGridViewColumnHeadersHeightSizeMode,
    FormBorderStyle,
    FlowLayoutPanel,
    TableLayoutPanel,
    DockStyle,
    RowStyle,
    SizeType,
    FormWindowState,
    DataGridViewColumnSortMode,
    Timer,
)
from System.Drawing import Size, Color
from System.IO import Path
import System.Threading
import threading
import subprocess
import os

class MyForm(Form):
    def __init__(self):
        self.Text = "AEDT Scheduler"
        self.Size = Size(840, 550)
        self.FormBorderStyle = FormBorderStyle.FixedSingle
        self.MaximizeBox = False
        Control.CheckForIllegalCrossThreadCalls = False
        self.Load += self.on_load

        self.ansysedt_path = os.environ.get(
            "ANSYSEDT_PATH",
            r"C:\Program Files\ANSYS Inc\v251\AnsysEM\ansysedt",
        )

        # 狀態顯示
        self.status_label = Label()
        self.base_status_text = "Ready"
        self.status_label.Text = self.base_status_text
        self.status_label.AutoSize = True
        self.status_label.Dock = DockStyle.Top

        # 設定資源使用監視
        try:
            import psutil  # type: ignore
            self.psutil = psutil
        except Exception:
            self.psutil = None
        if self.psutil is not None:
            self.self_proc = self.psutil.Process(os.getpid())
            self.self_proc.cpu_percent(interval=None)
            self.update_timer = Timer()
            self.update_timer.Interval = 1000
            self.update_timer.Tick += self.update_resource_usage
            self.update_timer.Start()
            self.update_resource_usage(None, None)

        # 按鈕列使用 FlowLayoutPanel
        self.button_panel = FlowLayoutPanel()
        self.button_panel.Dock = DockStyle.Top
        self.button_panel.AutoSize = True

        # 新增檔案按鈕
        self.add_button = Button()
        self.add_button.Text = "Add File"
        self.add_button.AutoSize = True
        self.add_button.Click += self.add_file
        self.button_panel.Controls.Add(self.add_button)

        # 移除檔案按鈕
        self.remove_button = Button()
        self.remove_button.Text = "Remove"
        self.remove_button.AutoSize = True
        self.remove_button.Click += self.remove_file
        self.button_panel.Controls.Add(self.remove_button)

        # 上移按鈕
        self.up_button = Button()
        self.up_button.Text = "Up"
        self.up_button.AutoSize = True
        self.up_button.Click += self.move_up
        self.button_panel.Controls.Add(self.up_button)

        # 下移按鈕
        self.down_button = Button()
        self.down_button.Text = "Down"
        self.down_button.AutoSize = True
        self.down_button.Click += self.move_down
        self.button_panel.Controls.Add(self.down_button)

        # 非圖形化模式選項
        self.ng_checkbox = CheckBox()
        self.ng_checkbox.Text = "Non Graphical"
        self.ng_checkbox.AutoSize = True
        self.button_panel.Controls.Add(self.ng_checkbox)

        # Queue Jobs label
        self.queue_label = Label()
        self.queue_label.Text = "Queue Jobs"
        self.queue_label.AutoSize = True

        # 排程列表改為 DataGridView
        self.queue_grid = DataGridView()
        self.queue_grid.Dock = DockStyle.Fill
        self.queue_grid.ReadOnly = True
        self.queue_grid.AllowUserToAddRows = False
        self.queue_grid.ColumnHeadersHeightSizeMode = (
            DataGridViewColumnHeadersHeightSizeMode.AutoSize
        )
        self.queue_grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.queue_grid.ShowCellToolTips = False
        self.queue_grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self.queue_grid.MultiSelect = True
        self.queue_grid.SelectionChanged += self.prevent_select_running
        self.queue_grid.CellDoubleClick += self.open_queue_folder

        q_file_col = DataGridViewTextBoxColumn()
        q_file_col.HeaderText = "File"
        q_file_col.SortMode = DataGridViewColumnSortMode.NotSortable
        ng_col = DataGridViewTextBoxColumn()
        ng_col.HeaderText = "NG"
        ng_col.SortMode = DataGridViewColumnSortMode.NotSortable
        submit_col = DataGridViewTextBoxColumn()
        submit_col.HeaderText = "Submit Time"
        submit_col.SortMode = DataGridViewColumnSortMode.NotSortable
        path_col = DataGridViewTextBoxColumn()
        path_col.HeaderText = "Full Path"
        path_col.SortMode = DataGridViewColumnSortMode.NotSortable

        self.queue_grid.Columns.Add(q_file_col)
        self.queue_grid.Columns.Add(ng_col)
        self.queue_grid.Columns.Add(submit_col)
        self.queue_grid.Columns.Add(path_col)

        # Completed Jobs label
        self.finished_label = Label()
        self.finished_label.Text = "Completed Jobs"
        self.finished_label.AutoSize = True

        # 完成列表改為 DataGridView
        self.finished_grid = DataGridView()
        self.finished_grid.Dock = DockStyle.Fill
        self.finished_grid.ReadOnly = True
        self.finished_grid.AllowUserToAddRows = False
        self.finished_grid.ColumnHeadersHeightSizeMode = (
            DataGridViewColumnHeadersHeightSizeMode.AutoSize
        )
        self.finished_grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        self.finished_grid.ShowCellToolTips = False
        self.finished_grid.CellDoubleClick += self.open_finished_folder

        file_col = DataGridViewTextBoxColumn()
        file_col.HeaderText = "File"
        start_col = DataGridViewTextBoxColumn()
        start_col.HeaderText = "Start Time"
        stop_col = DataGridViewTextBoxColumn()
        stop_col.HeaderText = "Stop Time"
        dur_col = DataGridViewTextBoxColumn()
        dur_col.HeaderText = "Duration"
        path2_col = DataGridViewTextBoxColumn()
        path2_col.HeaderText = "Full Path"

        self.finished_grid.Columns.Add(file_col)
        self.finished_grid.Columns.Add(start_col)
        self.finished_grid.Columns.Add(stop_col)
        self.finished_grid.Columns.Add(dur_col)
        self.finished_grid.Columns.Add(path2_col)

        # 使用 TableLayoutPanel 組織排程和完成區域
        self.data_layout = TableLayoutPanel()
        self.data_layout.Dock = DockStyle.Fill
        self.data_layout.ColumnCount = 1
        self.data_layout.RowCount = 4
        self.data_layout.RowStyles.Add(RowStyle(SizeType.AutoSize))
        self.data_layout.RowStyles.Add(RowStyle(SizeType.Percent, 50))
        self.data_layout.RowStyles.Add(RowStyle(SizeType.AutoSize))
        self.data_layout.RowStyles.Add(RowStyle(SizeType.Percent, 50))

        self.data_layout.Controls.Add(self.queue_label, 0, 0)
        self.data_layout.Controls.Add(self.queue_grid, 0, 1)
        self.data_layout.Controls.Add(self.finished_label, 0, 2)
        self.data_layout.Controls.Add(self.finished_grid, 0, 3)

        # 將元件加入表單
        self.Controls.Add(self.data_layout)
        self.Controls.Add(self.button_panel)
        self.Controls.Add(self.status_label)

        self.queue_paths = []
        self.queue_times = []
        self.queue_ngs = []
        self.finished_paths = []
        self.tooltip = ToolTip()
        self.queue_grid.MouseMove += self.show_queue_tooltip
        self.finished_grid.MouseMove += self.show_finished_tooltip

        self.is_simulating = False
        self.current_file = None
        self.current_process = None
        self.stop_event = threading.Event()
        self.FormClosing += self.on_close

    def on_load(self, sender, event):
        self.WindowState = FormWindowState.Normal

    def update_resource_usage(self, sender, event):
        if self.psutil is None:
            self.status_label.Text = self.base_status_text
            return
        cpu = self.self_proc.cpu_percent(interval=None)
        mem = self.self_proc.memory_info().rss
        if self.current_process is not None:
            try:
                child = self.psutil.Process(self.current_process.pid)
                cpu += child.cpu_percent(interval=None)
                mem += child.memory_info().rss
            except Exception:
                pass
        info = "CPU: {:.1f}% | Mem: {:.1f} MB".format(cpu, mem / 1024.0 / 1024.0)
        self.status_label.Text = self.base_status_text + " | " + info

    def highlight_current_row(self):
        for i in range(self.queue_grid.Rows.Count):
            self.queue_grid.Rows[i].DefaultCellStyle.BackColor = Color.White
        if self.is_simulating and self.queue_grid.Rows.Count > 0:
            self.queue_grid.Rows[0].DefaultCellStyle.BackColor = Color.LightGreen
        self.prevent_select_running(None, None)

    def prevent_select_running(self, sender, event):
        if self.is_simulating and self.queue_grid.Rows.Count > 0:
            if self.queue_grid.Rows[0].Selected:
                self.queue_grid.Rows[0].Selected = False

    def add_file(self, sender, event):
        dialog = OpenFileDialog()
        dialog.Title = "選擇檔案"
        dialog.Filter = "AEDT Files (*.aedt;*.aedtz)|*.aedt;*.aedtz"
        dialog.Multiselect = True
        if dialog.ShowDialog() == DialogResult.OK:
            files_added = False
            for fname in dialog.FileNames:
                if fname.lower().endswith('.aedt') or fname.lower().endswith('.aedtz'):
                    lock_path = fname + '.lock'
                    if os.path.isfile(lock_path):
                        try:
                            os.remove(lock_path)
                        except Exception:
                            pass
                    self.queue_paths.append(fname)
                    submit_time = System.DateTime.Now
                    self.queue_times.append(submit_time)
                    self.queue_ngs.append(self.ng_checkbox.Checked)
                    self.queue_grid.Rows.Add(
                        Path.GetFileName(fname),
                        "Yes" if self.ng_checkbox.Checked else "No",
                        submit_time.ToString(),
                        fname,
                    )
                    files_added = True
                else:
                    MessageBox.Show("Only .aedt or .aedtz files are allowed.")
            if files_added and not self.is_simulating:
                self.start_simulation()
            else:
                self.highlight_current_row()

    def remove_file(self, sender, event):
        if self.queue_grid.SelectedCells.Count == 0:
            return
        indices = sorted(set([cell.RowIndex for cell in self.queue_grid.SelectedCells]), reverse=True)
        removed = False
        if self.is_simulating and 0 in indices:
            MessageBox.Show("Cannot remove file being simulated.")
            indices = [i for i in indices if i != 0]
        for idx in indices:
            self.queue_grid.Rows.RemoveAt(idx)
            del self.queue_paths[idx]
            del self.queue_times[idx]
            del self.queue_ngs[idx]
            removed = True
        if removed:
            self.highlight_current_row()

    def move_up(self, sender, event):
        index = -1
        if self.queue_grid.SelectedCells.Count > 0:
            index = self.queue_grid.SelectedCells[0].RowIndex
        if index > 0:
            if self.is_simulating and (index - 1 == 0):
                MessageBox.Show("Cannot move file being simulated.")
                return
            self.swap_queue_rows(index, index - 1)
            self.queue_paths.insert(index - 1, self.queue_paths.pop(index))
            self.queue_times.insert(index - 1, self.queue_times.pop(index))
            self.queue_ngs.insert(index - 1, self.queue_ngs.pop(index))
            self.queue_grid.ClearSelection()
            self.queue_grid.Rows[index - 1].Selected = True
            self.highlight_current_row()

    def move_down(self, sender, event):
        index = -1
        if self.queue_grid.SelectedCells.Count > 0:
            index = self.queue_grid.SelectedCells[0].RowIndex
        if index != -1 and index == self.queue_grid.Rows.Count - 1:
            MessageBox.Show("Cannot move down last item.")
            return
        if index != -1 and index < self.queue_grid.Rows.Count - 1:
            if self.is_simulating and index == 0:
                MessageBox.Show("Cannot move file being simulated.")
                return
            self.swap_queue_rows(index, index + 1)
            self.queue_paths.insert(index + 1, self.queue_paths.pop(index))
            self.queue_times.insert(index + 1, self.queue_times.pop(index))
            self.queue_ngs.insert(index + 1, self.queue_ngs.pop(index))
            self.queue_grid.ClearSelection()
            self.queue_grid.Rows[index + 1].Selected = True
            self.highlight_current_row()

    def swap_queue_rows(self, i, j):
        for col in range(self.queue_grid.ColumnCount):
            tmp = self.queue_grid.Rows[i].Cells[col].Value
            self.queue_grid.Rows[i].Cells[col].Value = self.queue_grid.Rows[j].Cells[col].Value
            self.queue_grid.Rows[j].Cells[col].Value = tmp

    def show_queue_tooltip(self, sender, event):
        hit = self.queue_grid.HitTest(event.X, event.Y)
        idx = hit.RowIndex
        if idx >= 0 and idx < len(self.queue_paths):
            text = self.queue_paths[idx]
            if self.tooltip.GetToolTip(self.queue_grid) != text:
                self.tooltip.SetToolTip(self.queue_grid, text)
        else:
            self.tooltip.SetToolTip(self.queue_grid, "")

    def show_finished_tooltip(self, sender, event):
        hit = self.finished_grid.HitTest(event.X, event.Y)
        idx = hit.RowIndex
        if idx >= 0 and idx < len(self.finished_paths):
            text = self.finished_paths[idx]
            if self.tooltip.GetToolTip(self.finished_grid) != text:
                self.tooltip.SetToolTip(self.finished_grid, text)
        else:
            self.tooltip.SetToolTip(self.finished_grid, "")

    def open_queue_folder(self, sender, event):
        if event.RowIndex >= 0 and event.RowIndex < len(self.queue_paths):
            folder = os.path.dirname(self.queue_paths[event.RowIndex])
            try:
                subprocess.Popen(["explorer", folder])
            except Exception:
                MessageBox.Show("Unable to open folder.")

    def open_finished_folder(self, sender, event):
        if event.RowIndex >= 0 and event.RowIndex < len(self.finished_paths):
            folder = os.path.dirname(self.finished_paths[event.RowIndex])
            try:
                subprocess.Popen(["explorer", folder])
            except Exception:
                MessageBox.Show("Unable to open folder.")

    def start_simulation(self, sender=None, event=None):
        if self.is_simulating:
            return
        self.is_simulating = True
        self.highlight_current_row()
        self.sim_thread = threading.Thread(target=self.run_simulation)
        self.sim_thread.IsBackground = True
        self.sim_thread.start()

    def run_simulation(self):
        while len(self.queue_paths) > 0 and not self.stop_event.is_set():
            self.highlight_current_row()
            file_path = self.queue_paths[0]
            ng_flag = self.queue_ngs[0]
            self.current_file = file_path
            self.base_status_text = "Simulating: " + file_path
            self.update_resource_usage(None, None)
            start_time = System.DateTime.Now
            if file_path.lower().endswith('.aedt'):
                cmd = [self.ansysedt_path, "-batchsolve"]
                if ng_flag:
                    cmd.append("-ng")
                cmd.append(file_path)
                self.current_process = subprocess.Popen(cmd)
                self.current_process.wait()
                self.current_process = None
                if self.stop_event.is_set():
                    break
            else:
                System.Threading.Thread.Sleep(1000)
            stop_time = System.DateTime.Now
            duration = stop_time - start_time
            duration_str = str(duration).split(".")[0]
            self.queue_grid.Rows.RemoveAt(0)
            del self.queue_paths[0]
            del self.queue_times[0]
            del self.queue_ngs[0]
            self.highlight_current_row()
            if not self.stop_event.is_set():
                self.finished_grid.Rows.Add(
                    Path.GetFileName(file_path),
                    start_time.ToString(),
                    stop_time.ToString(),
                    duration_str,
                    file_path,
                )
                self.finished_paths.append(file_path)
            self.current_file = None
        self.highlight_current_row()
        self.base_status_text = "Finished"
        self.update_resource_usage(None, None)
        self.is_simulating = False
        
    def on_close(self, sender, event):
        self.stop_event.set()
        if self.current_process is not None and self.current_process.poll() is None:
            try:
                self.current_process.terminate()
                self.current_process.wait(5)
            except Exception:
                pass
            if self.current_process.poll() is None:
                try:
                    self.current_process.kill()
                except Exception:
                    pass
        try:
            subprocess.call("taskkill /f /im ansysedt.exe", shell=True)
        except Exception:
            pass
        if hasattr(self, 'sim_thread') and self.sim_thread.is_alive():
            self.sim_thread.join(1)

form = MyForm()
form.ShowDialog()
