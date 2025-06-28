# AEDT Scheduler

This project contains a small IronPython script that provides a GUI for
managing a simulation queue.  Users can add files to the queue, reorder or
remove them and items in the queue will be processed automatically in order.
The queue is displayed as a grid showing the file name and the submit time for
each item. Finished items are moved to a separate grid so you can track what
has been processed.

Simulation begins as soon as a file is added to the queue, so there is no
longer a separate "Start" button.

Only `.aedt` or `.aedtz` files can be added to the queue. Run `run.bat` to
start the application in an Ansys IronPython environment.
