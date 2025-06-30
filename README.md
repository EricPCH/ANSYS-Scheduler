# AEDT Scheduler

This project contains a small IronPython script that provides a GUI for
managing a simulation queue.  Users can add files to the queue, reorder or
remove them and items in the queue will be processed automatically in order.
The queue is displayed as a grid showing the file name and the submit time for
each item. A "Non Graphical" column indicates whether the item will be solved
with the `-ng` option. Finished items are moved to a separate grid so you can
track what has been processed.

Simulation begins as soon as a file is added to the queue, so there is no
longer a separate "Start" button.

Use the **Add File** button to select one or more `.aedt` or `.aedtz` files.
Run `run.bat` to start the application in an Ansys IronPython environment.
Use the **Settings** menu to set the full path to `ansysedt.exe`.  The value is
stored in `config.json` so you only need to configure it once.

Before adding a file you can tick the **Non Graphical** check box to run that
simulation in non-graphical mode. Each file can be configured independently.

![image](https://github.com/user-attachments/assets/cd04f12b-c9a6-4c38-91c0-c225b185cf9d)



## 使用說明

* 請將 `scheduler.py` 和 `run.bat` 置於同一目錄中使用。
* 使用前可修改 `run.bat` 中 IronPython `ipy64.exe` 的路徑。
* `ansysedt.exe` 的位置可以在程式的 **Settings** 視窗中設定，並會儲存在 `config.json`。
* 在排程列表或完成列表中，雙擊 **Full Path** 欄位可開啟該檔案所在資料夾。

