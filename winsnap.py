from __future__ import annotations

import collections
from collections import defaultdict
import ctypes
from dataclasses import dataclass
from functools import reduce
import json
import os
from pathlib import Path
import sys
from uuid import uuid4

from dearpygui import core as dpg_core
from dearpygui import simple as dpg_simple
import psutil
from pywinauto import Desktop
import win32api
import win32con

user32 = ctypes.windll.user32
dwmapi = ctypes.windll.dwmapi
user32.SetProcessDPIAware()
sys.coinit_flags = 2  # STA

DWMWA_EXTENDED_FRAME_BOUNDS = 9
SPI_GETWORKAREA = 48
DWORD_DWMWA_EXTENDED_FRAME_BOUNDS = ctypes.wintypes.DWORD(DWMWA_EXTENDED_FRAME_BOUNDS)


class RectangleMixin:
    def __repr__(self):
        return (
            f"{self.__class__.__name__}(left={self.left}, top={self.top}, right={self.right}, "
            f"bottom={self.bottom})"
        )

    @property
    def width(self) -> float:
        return self.right - self.left

    @property
    def height(self) -> float:
        return self.bottom - self.top


class Rectangle(RectangleMixin):
    def __init__(self, left, right, top, bottom) -> None:
        left, right, top, bottom = float(left), float(right), float(top), float(bottom)
        if left > right:
            left, right = right, left
        if top > bottom:
            top, bottom = bottom, top
        self._left = left
        self._right = right
        self._top = top
        self._bottom = bottom

    def __getitem__(self, index) -> float:
        if index == 0:
            return self.left
        elif index == 1:
            return self.right
        elif index == 2:
            return self.top
        elif index == 3:
            return self.bottom
        else:
            raise IndexError(f"{index} is not in alllowed range of 0-3")

    @property
    def left(self) -> float:
        return self._left

    @property
    def right(self) -> float:
        return self._right

    @property
    def top(self) -> float:
        return self._top

    @property
    def bottom(self) -> float:
        return self._bottom


@dataclass(frozen=True)
class Monitor:

    area: Rectangle
    work: Rectangle
    is_primary: bool
    name: str
    id: str

    @classmethod
    def from_monitor_handle(cls, handle):
        # https://docs.microsoft.com/en-us/windows/win32/api/windef/ns-windef-rect
        info = win32api.GetMonitorInfo(handle)
        device = win32api.EnumDisplayDevices(info["Device"], 0)
        area_left, area_top, area_right, area_bottom = info["Monitor"]
        work_left, work_top, work_right, work_bottom = info["Work"]
        return cls(
            area=Rectangle(area_left, area_right, area_top, area_bottom),
            work=Rectangle(work_left, work_right, work_top, work_bottom),
            is_primary=info["Flags"] == 1,
            name=info["Device"].lstrip("\\\\.\\"),
            id=device.DeviceID,
        )

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}({self.name})"

    def __hash__(self):
        return self.id


def get_monitors() -> list[Monitor]:
    monitors = []
    for handle, *_ in win32api.EnumDisplayMonitors():
        monitors.append(Monitor.from_monitor_handle(handle))
    return monitors


def get_windows():
    windows = {}
    apps = collections.defaultdict(list)
    pids_to_windows = {}
    handles = defaultdict(list)
    for window in Desktop(backend="uia").windows():
        text = window.window_text()
        if text and text not in {"Program Manager", "Taskbar"}:
            pid = window.process_id()
            app = psutil.Process(pid).name()
            apps[app].append(pid)
            pids_to_windows[pid] = window
            handles[pid].append((window, window.handle))
    for app, pids in apps.items():
        pids = list(pids)
        if len(pids) == 1:
            windows[app] = pids_to_windows[pids[0]]
        else:
            for n, pid in enumerate(pids):
                name = f"{app} - {n + 1}"
                windows[name] = pids_to_windows[pid]

    return windows


def get_window_buffer(handle):
    rect = ctypes.wintypes.RECT()
    windowrect = ctypes.wintypes.RECT()
    handle = ctypes.wintypes.HWND(handle)
    dwmapi.DwmGetWindowAttribute(
        handle, DWORD_DWMWA_EXTENDED_FRAME_BOUNDS, ctypes.byref(rect), ctypes.sizeof(rect)
    )
    user32.GetWindowRect(handle, ctypes.byref(windowrect))
    return (
        abs(windowrect.left - rect.left),
        abs(windowrect.right - rect.right),
        abs(windowrect.top - rect.top),
        abs(windowrect.bottom - rect.bottom),
    )


def move_window(handle, x, y, width, height):
    move_flags = win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE | win32con.SWP_NOSIZE
    size_flags = win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE | win32con.SWP_NOMOVE
    left_buffer, right_buffer, top_buffer, bottom_buffer = get_window_buffer(handle)
    handle = ctypes.wintypes.HWND(int(handle))
    user32.ShowWindow(handle, win32con.SW_RESTORE)
    x, y, height, width = int(x), int(y), int(height), int(width)
    user32.SetWindowPos(handle, None, x - left_buffer, y - top_buffer, 0, 0, move_flags)
    user32.SetWindowPos(
        handle,
        None,
        0,
        0,
        width + left_buffer + right_buffer,
        height + top_buffer + bottom_buffer,
        size_flags,
    )


class UniqueWidget:
    def __init__(self) -> None:
        self._id = str(uuid4())

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}({self.id})"

    @property
    def id(self) -> str:
        return self._id

    def create_label(self, show: str, hide: str) -> str:
        if hide:
            hide = f"{hide}-{self.id}"
        else:
            hide = self.id
        return f"{show}##{hide}"


class AppTable(UniqueWidget):

    HEADER = ["Number", "Application"]
    ACTIVE_WINDOWS = {}

    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        with dpg_simple.managed_columns(f"{self._id}_head", len(self.HEADER)):
            for item in self.HEADER:
                dpg_core.add_text(item)
        self._allocated_windows = set()
        self._available_windows = []
        self._grid_mapping = defaultdict(list)
        self._nrows = 0

    @property
    def allocated_windows(self):
        return self._allocated_windows.copy()

    def clear(self):
        for row in reversed(range(1, self._nrows + 1)):
            label = f"{self._id}_{row}"
            dpg_core.delete_item(label)
        self._nrows = 0

    def selected(self, sender, data):
        selections = dpg_core.get_table_selections(sender)
        selected_items = [dpg_core.get_table_item(sender, *s) for s in selections]
        number = int(sender.split("_")[1])
        self._grid_mapping[number] = selected_items
        self._allocated_windows = set(
            reduce(lambda s1, s2: set(s1) | set(s2), self._grid_mapping.values())
        )
        self.refresh_available_windows()

    def refresh_available_windows(self, *args, **kwargs):
        available_windows = set(self.ACTIVE_WINDOWS) - self._allocated_windows
        for row in range(1, self._nrows + 1):
            dpg_core.clear_table(f"{self._id}_{row}_table")
            windows = sorted(list(available_windows | set(self._grid_mapping[row])))
            for n, name in enumerate(windows):
                dpg_core.add_row(f"{self._id}_{row}_table", [name])
                if name in self._grid_mapping[row]:
                    dpg_core.set_table_selection(f"{self._id}_{row}_table", n, 0, True)

    def set_rows(self, nrows):
        for row in range(1, nrows + 1):
            if dpg_core.does_item_exist(f"{self._id}_{row}"):
                continue
            with dpg_simple.managed_columns(
                f"{self._id}_{row}", len(self.HEADER), parent=self.parent
            ):
                dpg_core.add_input_int(
                    f"##{self._id}_{row}_number",
                    default_value=row,
                    readonly=True,
                    step=0,
                    parent=f"{self._id}_{row}",
                )
                with dpg_simple.tree_node(
                    f"##{self._id}_{row}_tree", label=f"apps{row}", parent=f"{self._id}_{row}"
                ):
                    dpg_core.add_table(
                        f"{self._id}_{row}_table",
                        [""],
                        callback=self.selected,
                        parent=f"##{self._id}_{row}_tree",
                    )
                    for name in sorted(self.ACTIVE_WINDOWS):
                        dpg_core.add_row(f"{self._id}_{row}_table", [name])
            dpg_core.add_separator(name=f"{self._id}_{row}_sep", parent=f"{self._id}_{row}")
        self._nrows = nrows


class MonitorProfile(UniqueWidget):
    def __init__(self, monitor: Monitor, parent: str) -> None:
        super().__init__()
        self.parent = parent
        self.monitor = monitor
        self._ylines = []
        self._xlines = []
        self._labels = []
        self._area_mapping = {}
        self._left_panel_id = str(uuid4())
        self._right_panel_id = str(uuid4())
        self._input_id = self.create_label("", "input")
        self._snap_id = self.create_label("Snap", "snap")
        self._plot_id = self.create_label("", "plot")
        self._tab_id = self.create_label(self.monitor.name, None)
        self._app_table = None
        self.init_ui()
        self.set_labels()
        self._app_table.set_rows(1)
        self.resize()

    def set_labels(self):
        for label in self._labels:
            dpg_core.delete_annotation(self._plot_id, label)
            dpg_core.delete_annotation(self._plot_id, f"{label}%")
        x0 = self.monitor.work.left
        xs = []
        x_percents = []
        yline_values = sorted([dpg_core.get_value(yline) for yline in self._ylines])
        for x1 in yline_values:
            xs.append(((x0 + x1) / 2))
            x_percents.append(((x1 - x0) / self.monitor.work.width) * 100)
            x0 = x1
        xs.append((self.monitor.work.right + x0) / 2)
        x_percents.append(((self.monitor.work.right - x0) / self.monitor.work.width) * 100)
        ys = []
        y_percents = []
        y0 = self.monitor.work.top
        xline_values = sorted([dpg_core.get_value(xline) for xline in self._xlines])
        for y1 in xline_values:
            ys.append(((y0 + y1) / 2))
            y_percents.append(((y1 - y0) / self.monitor.work.height) * 100)
            y0 = y1
        ys.append((self.monitor.work.bottom + y0) / 2)
        y_percents.append(((self.monitor.work.bottom - y0) / self.monitor.work.height) * 100)
        number = 1
        labels = []
        for y, y_percent in zip(reversed(sorted(ys)), reversed(y_percents)):
            for x, x_percent in zip(sorted(xs), x_percents):
                tag = str(uuid4())
                dpg_core.add_annotation(self._plot_id, str(number), x, y, 0, 0, tag=tag)
                dpg_core.add_annotation(
                    self._plot_id, f"{x_percent:.2f}x{y_percent:.2f}", x, y, 0, 2, tag=f"{tag}%"
                )
                labels.append(tag)
                number += 1
        self._labels = labels

        self._area_mapping.clear()
        y_edges = [self.monitor.work.top] + xline_values + [self.monitor.work.bottom]
        x_edges = [self.monitor.work.left] + yline_values + [self.monitor.work.right]
        row = 1
        for x_index in range(len(x_edges) - 1):
            x0, x1 = x_edges[x_index], x_edges[x_index + 1]
            for y_index in range(len(y_edges) - 1):
                y0, y1 = y_edges[y_index], y_edges[y_index + 1]
                self._area_mapping[row] = Rectangle(x0, x1, y0, y1)
                row += 1

    def line_callback(self, sender, data):
        dpg_core.set_value(sender, int(dpg_core.get_value(sender)))
        self.set_labels()

    def input_callback(self, sender, data):
        for xline in self._xlines:
            dpg_core.delete_drag_line(self._plot_id, xline)
        for yline in self._ylines:
            dpg_core.delete_drag_line(self._plot_id, yline)
        rows, cols = dpg_core.get_value(self._input_id)
        rows = 1 if rows < 1 else rows
        cols = 1 if cols < 1 else cols
        dpg_core.set_value(self._input_id, [rows, cols])
        xlines = []
        for row in range(1, rows):
            name = f"xline{row}-{self.id}"
            pos = int(self.monitor.work.top + (row / rows) * self.monitor.work.height)
            dpg_core.add_drag_line(
                self._plot_id,
                name,
                y_line=True,
                show_label=False,
                default_value=pos,
                callback=self.line_callback,
            )
            xlines.append(name)
        ylines = []
        for col in range(1, cols):
            name = f"yline{col}-{self.id}"
            pos = int(self.monitor.work.left + (col / cols) * self.monitor.work.width)
            dpg_core.add_drag_line(
                self._plot_id,
                name,
                y_line=False,
                show_label=False,
                default_value=pos,
                callback=self.line_callback,
            )
            ylines.append(name)

        self._xlines = xlines
        self._ylines = ylines

        self.set_labels()
        self._app_table.clear()
        self._app_table.set_rows(rows * cols)

    def snap(self, sender, data):
        for number, windows in self._app_table._grid_mapping.items():
            if number not in self._area_mapping:
                continue
            area = self._area_mapping[number]
            for name in windows:
                window = AppTable.ACTIVE_WINDOWS[name]
                move_window(
                    window.handle, int(area.left), int(area.top), int(area.width), int(area.height)
                )

    def resize(self):
        window_width = dpg_simple.get_item_width("Main Window")
        plot_width = window_width // 2
        table_width = window_width - plot_width
        dpg_simple.set_item_width(self._left_panel_id, plot_width)
        dpg_simple.set_item_width(self._right_panel_id, table_width)

    def init_ui(self):
        with dpg_simple.tab(self._tab_id, parent=self.parent, no_tooltip=True):
            with dpg_simple.group(self._left_panel_id, parent=self._tab_id):
                dpg_core.add_input_int2(
                    self._input_id,
                    parent=self._left_panel_id,
                    callback=self.input_callback,
                    default_value=[1, 1],
                )
                dpg_core.add_button(self._snap_id, parent=self._left_panel_id, callback=self.snap)
                dpg_core.add_plot(
                    self._plot_id,
                    parent=self._left_panel_id,
                    height=-1,
                    xaxis_lock_min=True,
                    xaxis_lock_max=True,
                    y2axis_lock_min=True,
                    y2axis_lock_max=True,
                    yaxis_no_tick_marks=True,
                    yaxis_no_tick_labels=True,
                    xaxis_no_tick_marks=True,
                    xaxis_no_tick_labels=True,
                    xaxis_no_gridlines=True,
                    yaxis_no_gridlines=True,
                )
                dpg_core.set_plot_xlimits(
                    self._plot_id, xmin=self.monitor.work.left, xmax=self.monitor.work.right
                )
                dpg_core.set_plot_ylimits(
                    self._plot_id, ymin=self.monitor.work.top, ymax=self.monitor.work.bottom
                )
            dpg_core.add_same_line(parent=self._tab_id)
            with dpg_simple.group(self._right_panel_id, parent=self._tab_id):
                self._app_table = AppTable(parent=self._right_panel_id)

    def to_dict(self) -> dict[str, list[float]]:
        return {
            "xlines": sorted(
                dpg_core.get_value(xline) / self.monitor.work.height for xline in self._xlines
            ),
            "ylines": sorted(
                dpg_core.get_value(yline) / self.monitor.work.width for yline in self._ylines
            ),
        }

    def from_dict(self, serialized_monitor_profile: dict[str, list[float]]) -> None:
        xlines = []
        for row, value in enumerate(serialized_monitor_profile["xlines"]):
            xline = int(value * self.monitor.work.height)
            name = f"xline{row}-{self.id}"
            dpg_core.add_drag_line(
                self._plot_id,
                name,
                y_line=True,
                show_label=False,
                default_value=xline,
                callback=self.line_callback,
            )
            xlines.append(name)
        ylines = []
        for col, value in enumerate(serialized_monitor_profile["ylines"]):
            yline = int(value * self.monitor.work.width)
            name = f"yline{col}-{self.id}"
            dpg_core.add_drag_line(
                self._plot_id,
                name,
                y_line=False,
                show_label=False,
                default_value=yline,
                callback=self.line_callback,
            )
            ylines.append(name)
        self._xlines = xlines
        self._ylines = ylines
        self.set_labels()
        rows, cols = (len(self._xlines) + 1), (len(self._ylines) + 1)
        dpg_core.set_value(self._input_id, [rows, cols])
        self._app_table.set_rows(rows * cols)


class Profile(UniqueWidget):
    def __init__(self, parent, monitors):
        super().__init__()
        self.parent = parent
        self._monitors = monitors
        self.monitor_profiles: dict[str, MonitorProfile] = {}
        self._tab_bar_id = self.create_label("", "tab-bar")
        self.init_ui()

    def init_ui(self):
        with dpg_simple.tab_bar(self._tab_bar_id, parent=self.parent):
            for monitor in self._monitors:
                monitor_profile = MonitorProfile(monitor, self._tab_bar_id)
                self.monitor_profiles[monitor.id] = monitor_profile

    def to_dict(self):
        serial = {}
        for monitor_id, monitor_profile in self.monitor_profiles.items():
            serial[monitor_id] = monitor_profile.to_dict()
        return serial

    def from_dict(self, serialized_profile: dict[str, dict[str, list[float]]]):
        for monitor_id, serialized_monitor_profile in serialized_profile.items():
            monitor_profile = self.monitor_profiles[monitor_id]
            monitor_profile.from_dict(serialized_monitor_profile)


AppTable.ACTIVE_WINDOWS = get_windows()


class MainWindow(UniqueWidget):

    PROFILES_PATH = Path(os.path.expandvars("%APPDATA%")) / "WinSnap" / "profiles.json"

    def __init__(self) -> None:
        super().__init__()
        self._profiles: dict[str, Profile] = {}
        self._tab_number = 1
        self._monitors = get_monitors()
        self._save_id = "Save##MainWindow-save"
        self._load_id = "Load##MainWindow-load"
        if self.PROFILES_PATH.exists():
            with self.PROFILES_PATH.open("r") as stream:
                self._saved_profiles = json.load(stream)
        else:
            if not self.PROFILES_PATH.parent.exists():
                os.mkdir(self.PROFILES_PATH.parent)
            self._saved_profiles = {}
        self.init_ui()
        dpg_core.set_resize_callback(self.resize_callback, handler="Main Window")

    def set_monitors(self):
        pass

    def get_or_create_monitor_configuration(self):
        configurations = self._saved_profiles.get("configurations", {})
        current_monitor_ids = set(monitor.id for monitor in self._monitors)
        for configuration_id, monitor_ids in configurations.items():
            if set(monitor_ids) == current_monitor_ids:
                break
        else:
            configuration_id = str(uuid4())
            configurations[configuration_id] = list(current_monitor_ids)
            self._saved_profiles["configurations"] = configurations
        return configuration_id

    def save(self, *args, **kwargs):
        serialized_profiles = []
        for profile in self._profiles.values():
            serialized_profiles.append(profile.to_dict())
        configuration_id = self.get_or_create_monitor_configuration()
        self._saved_profiles[configuration_id] = serialized_profiles
        with self.PROFILES_PATH.open("w") as stream:
            json.dump(self._saved_profiles, stream)

    def load(self, *args, **kwargs):
        configuration_id = self.get_or_create_monitor_configuration()
        serialized_profiles = self._saved_profiles.get(configuration_id)
        if not serialized_profiles:
            return

        for label in self._profiles:
            dpg_core.delete_item(label)
        self._profiles = {}
        self._tab_number = 1

        for serialized_profile in serialized_profiles:
            profile = self.add_tab()
            self.resize_callback(None, None)
            profile.from_dict(serialized_profile)
        self.resize_callback(*args, **kwargs)

    def add_tab(self, *args, **kwargs):
        label = f"{self._tab_number}##MainWindow-tab{self._tab_number}"
        with dpg_simple.tab(label, parent="##MainWindow-tabbar", closable=False, no_tooltip=True):
            profile = Profile(label, self._monitors)
            self._profiles[label] = profile
        self._tab_number += 1
        if len(self._profiles) == 2:
            for label in self._profiles:
                dpg_core.configure_item(label, closable=dpg_core.get_value(label))
                break
        return profile

    def remove_tab(self, *args, **kwargs):
        remove = set()
        for label in self._profiles:
            if dpg_core.is_item_shown(label):
                dpg_core.configure_item(label, closable=dpg_core.get_value(label))
            else:
                dpg_core.delete_item(label)
                remove.add(label)
        self._profiles = {
            label: self._profiles[label] for label in self._profiles if label not in remove
        }

    def init_ui(self):
        with dpg_simple.window("Main Window"):
            dpg_core.add_button(self._save_id, callback=self.save)
            dpg_core.add_same_line()
            dpg_core.add_button(self._load_id, callback=self.load)
            with dpg_simple.tab_bar(
                "##MainWindow-tabbar",
                parent="Main Window",
                reorderable=True,
                callback=self.remove_tab,
            ):
                self.add_tab()
                dpg_core.add_tab_button(
                    "+##MainWindow-btn",
                    parent="##MainWindow-tabbar",
                    callback=self.add_tab,
                    trailing=True,
                    no_tooltip=True,
                )

    def resize_callback(self, *args, **kwargs):
        for profile in self._profiles.values():
            for monitor_profile in profile.monitor_profiles.values():
                monitor_profile.resize()

    def refresh_windows(self):
        new_active_windows = get_windows()
        if new_active_windows != AppTable.ACTIVE_WINDOWS:
            AppTable.ACTIVE_WINDOWS = new_active_windows
            for profile in self._profiles.values():
                for monitor_profile in profile.monitor_profiles.values():
                    monitor_profile._app_table.refresh_available_windows()


main_window = MainWindow()


def mouse_click_cb(sender, data):
    main_window.refresh_windows()


def startup_cb(sender, data):
    main_window.load()


dpg_core.set_mouse_click_callback(mouse_click_cb)
dpg_core.set_start_callback(startup_cb)
dpg_core.start_dearpygui(primary_window="Main Window")
