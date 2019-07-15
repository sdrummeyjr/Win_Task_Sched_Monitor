# https://stackoverflow.com/questions/46038211/list-windows-scheduled-tasks-with-python
# https://stackoverflow.com/questions/36634214/python-check-for-completed-and-failed-task-windows-scheduler
# https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-2-0-interfaces

import os
import pywintypes
import win32com.client
import pandas as pd
from datetime import timezone
# import sqlite3

TASK_ENUM_HIDDEN = 1
TASK_STATE = {
    0: 'Unknown',
    1: 'Disabled',
    2: 'Queued',
    3: 'Ready',
    4: 'Running'
}


def walk_tasks(top, top_down=True, onerror=None, include_hidden=True, server_name=None, user=None, domain=None,
               password=None):
    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect(server_name, user, domain, password)
    if isinstance(top, bytes):
        if hasattr(os, 'fsdecode'):
            top = os.fsdecode(top)
        else:
            top = top.decode('mbcs')
    if u'/' in top:
        top = top.replace(u'/', u'\\')
    include_hidden = TASK_ENUM_HIDDEN if include_hidden else 0
    try:
        top = scheduler.GetFolder(top)
    except Exception as error:
        if onerror is not None:
            onerror(error)
        return
    for entry in _walk_tasks_internal(top, top_down, onerror, include_hidden):
        yield entry


def _walk_tasks_internal(top, topdown, onerror, flags):
    try:
        folders = list(top.GetFolders(0))
        tasks = list(top.GetTasks(flags))
    except pywintypes.com_error as error:
        if onerror is not None:
            onerror(error)
        return

    if not topdown:
        for d in folders:
            for entry in _walk_tasks_internal(d, topdown, onerror, flags):
                yield entry

    yield top, folders, tasks

    if topdown:
        for d in folders:
            for entry in _walk_tasks_internal(d, topdown, onerror, flags):
                yield entry


def task_data(win_tasks: list) -> pd.DataFrame:
    df = pd.DataFrame(win_tasks)
    df["Last_Run"] = pd.to_datetime(df["Last_Run"], unit='s')
    # print(df)
    # print(df.dtypes)
    return df


# def check_for_db():


def main():
    list_of_tasks = []
    for folder, sub_folders, tasks in walk_tasks('/'):
        for task in tasks:
            settings = task.Definition.Settings
            last_run_time = task.LastRunTime.replace(tzinfo=timezone.utc).timestamp()
            # print(last_run_time)
            list_of_tasks.append({"Path": task.Path, "Hidden": settings.Hidden, "State": TASK_STATE[task.State],
                                  "Last_Run": last_run_time, "Last_Result": task.LastTaskResult})
    data = task_data(list_of_tasks)
    writer = pd.ExcelWriter("windows_tasks.xlsx")
    data.to_excel(writer)
    writer.save()
    # print(data["Last_Run"])


if __name__ == '__main__':
    main()

# Converting POSIX timestamps from GMT to UTC so that it can be included in pd dataframe
# https://docs.python.org/3/library/datetime.html#datetime.datetime.timestamp
