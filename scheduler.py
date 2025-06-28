# -*- coding: utf-8 -*-

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import Form, Button, OpenFileDialog, Label
from System.Drawing import Point, Size

class MyForm(Form):
    def __init__(self):
        self.Text = "IronPython WinForms Example"
        self.Size = Size(400, 200)
        
        self.label = Label()
        self.label.Text = "尚未選擇檔案"
        self.label.Location = Point(20, 20)
        self.label.Size = Size(340, 20)
        self.Controls.Add(self.label)

        self.select_button = Button()
        self.select_button.Text = "選擇檔案"
        self.select_button.Location = Point(20, 60)
        self.select_button.Click += self.select_file
        self.Controls.Add(self.select_button)

        self.apply_button = Button()
        self.apply_button.Text = "Apply"
        self.apply_button.Location = Point(120, 60)
        self.apply_button.Click += self.apply_action
        self.Controls.Add(self.apply_button)

        self.selected_file = None

    def select_file(self, sender, event):
        dialog = OpenFileDialog()
        dialog.Title = "選擇檔案"
        dialog.Filter = "All files (*.*)|*.*"
        if dialog.ShowDialog() == 1:
            self.selected_file = dialog.FileName
            self.label.Text = "已選擇檔案：" + self.selected_file

    def apply_action(self, sender, event):
        if self.selected_file:
            self.label.Text = "Apply 處理中：" + self.selected_file
            # 在這裡加入你的處理邏輯
        else:
            self.label.Text = "請先選擇檔案"

form = MyForm()
form.ShowDialog()
