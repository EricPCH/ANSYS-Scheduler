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

Only `.aedt` or `.aedtz` files can be added to the queue. Run `run.bat` to
start the application in an Ansys IronPython environment. Edit the
`ANSYSEDT_PATH` variable inside `run.bat` if your installation lives in a
different location.

Before adding a file you can tick the **Non Graphical** check box to run that
simulation in non-graphical mode. Each file can be configured independently.

## 使用說明

* 請將 `scheduler.py` 和 `run.bat` 置於同一目錄中使用。
* 要使用程式必須修改 `run.bat` 內的兩個路徑：
  1. IronPython `ipy64.exe` 的範例路徑。
  2. `ANSYSEDT_PATH` 設定的 `ansysedt.exe` 目錄。

