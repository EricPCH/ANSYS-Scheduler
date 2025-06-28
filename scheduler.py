# -*- coding: utf-8 -*-

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import (
    Form,
    Button,
    OpenFileDialog,
    Label,
    ListBox,
    Application,
    MessageBox,
    DialogResult,
)
from System.Drawing import Point, Size
import System.Threading
import subprocess

class MyForm(Form):
    def __init__(self):
        self.Text = "AEDT Scheduler"
        self.Size = Size(620, 340)

        # 狀態顯示
        self.status_label = Label()
        self.status_label.Text = "Ready"
        self.status_label.Location = Point(20, 20)
        self.status_label.Size = Size(560, 20)
        self.Controls.Add(self.status_label)

        # 排程列表
        self.queue_list = ListBox()
        self.queue_list.Location = Point(20, 50)
        self.queue_list.Size = Size(260, 200)
        self.Controls.Add(self.queue_list)

        # 完成列表
        self.finished_list = ListBox()
        self.finished_list.Location = Point(320, 50)
        self.finished_list.Size = Size(260, 200)
        self.Controls.Add(self.finished_list)

        # 新增檔案按鈕
        self.add_button = Button()
        self.add_button.Text = "Add File"
        self.add_button.Location = Point(20, 260)
        self.add_button.Click += self.add_file
        self.Controls.Add(self.add_button)

        # 移除檔案按鈕
        self.remove_button = Button()
        self.remove_button.Text = "Remove"
        self.remove_button.Location = Point(110, 260)
        self.remove_button.Click += self.remove_file
        self.Controls.Add(self.remove_button)

        # 上移按鈕
        self.up_button = Button()
        self.up_button.Text = "Up"
        self.up_button.Location = Point(200, 260)
        self.up_button.Click += self.move_up
        self.Controls.Add(self.up_button)

        # 下移按鈕
        self.down_button = Button()
        self.down_button.Text = "Down"
        self.down_button.Location = Point(260, 260)
        self.down_button.Click += self.move_down
        self.Controls.Add(self.down_button)

        # 執行按鈕
        self.start_button = Button()
        self.start_button.Text = "Start"
        self.start_button.Location = Point(340, 260)
        self.start_button.Click += self.start_simulation
        self.Controls.Add(self.start_button)

    def add_file(self, sender, event):
        dialog = OpenFileDialog()
        dialog.Title = "選擇檔案"
        dialog.Filter = "AEDT Files (*.aedt;*.aedtz)|*.aedt;*.aedtz"
        if dialog.ShowDialog() == DialogResult.OK:
            fname = dialog.FileName
            if fname.lower().endswith('.aedt') or fname.lower().endswith('.aedtz'):
                self.queue_list.Items.Add(fname)
            else:
                MessageBox.Show("Only .aedt or .aedtz files are allowed.")

    def remove_file(self, sender, event):
        index = self.queue_list.SelectedIndex
        if index != -1:
            self.queue_list.Items.RemoveAt(index)

    def move_up(self, sender, event):
        index = self.queue_list.SelectedIndex
        if index > 0:
            item = self.queue_list.Items[index]
            self.queue_list.Items.RemoveAt(index)
            self.queue_list.Items.Insert(index - 1, item)
            self.queue_list.SelectedIndex = index - 1

    def move_down(self, sender, event):
        index = self.queue_list.SelectedIndex
        if index != -1 and index < self.queue_list.Items.Count - 1:
            item = self.queue_list.Items[index]
            self.queue_list.Items.RemoveAt(index)
            self.queue_list.Items.Insert(index + 1, item)
            self.queue_list.SelectedIndex = index + 1

    def start_simulation(self, sender, event):
        while self.queue_list.Items.Count > 0:
            file_path = self.queue_list.Items[0]
            self.status_label.Text = "Simulating: " + file_path
            Application.DoEvents()
            if file_path.lower().endswith('.aedt'):
                cmd = (
                    r'"C:\\Program Files\\ANSYS Inc\\v251\\AnsysEM\\ansysedt" '
                    r'-batchsolve "{0}"'.format(file_path)
                )
                subprocess.call(cmd, shell=True)
            else:
                System.Threading.Thread.Sleep(1000)
            self.queue_list.Items.RemoveAt(0)
            self.finished_list.Items.Add(file_path)
        self.status_label.Text = "Finished"

form = MyForm()
form.ShowDialog()
