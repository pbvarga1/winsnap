from __future__ import annotations

import argparse
import collections
from collections import defaultdict
import ctypes
from dataclasses import dataclass
from functools import reduce
import itertools
import json
import os
from pathlib import Path
import sys
from uuid import uuid4

from dearpygui import core as dpg_core
from dearpygui import simple as dpg_simple
import psutil
from pywinauto import Desktop
from pywinauto.controls.uiawrapper import UIAWrapper
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
    """Mixin class for handling rectangles"""

    def __repr__(self):
        return (
            f"{self.__class__.__name__}(left={self.left}, top={self.top}, right={self.right}, "
            f"bottom={self.bottom})"
        )

    @property
    def width(self) -> float:
        """float : Width of the rectangle"""
        return self.right - self.left

    @property
    def height(self) -> float:
        """float : Height of the rectangle"""
        return self.bottom - self.top


class Rectangle(RectangleMixin):
    """Represents a rectangle

    Parameters
    ----------
    left : float
        Coordinate of the left side of the rectangle
    right : float
        Coordinate of the right side of the rectangle
    top : float
        Coordinate of the top side of the rectangle
    bottom : float
        Coordinate of the bottom side of the rectangle
    """

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
        """float : left coordinate of the rectangle"""
        return self._left

    @property
    def right(self) -> float:
        """float : right coordinate of the rectangle"""
        return self._right

    @property
    def top(self) -> float:
        """float : top coordinate of the rectangle"""
        return self._top

    @property
    def bottom(self) -> float:
        """float : bottom coordinate of the rectangle"""
        return self._bottom


@dataclass(frozen=True)
class Monitor:
    """Information about a monitor

    Parameters
    ----------
    area : Rectangle
        Rectangle of the full monitor area
    work : Rectangle
        Rectangle of the area of the screen available to work (i.e. the part of the screen without
        the task bar)
    is_primary : bool
        Whether the monitor is the primary monitor or not
    name : str
        Name of the monitor
    id : Unique monitor ID
    """

    area: Rectangle
    work: Rectangle
    is_primary: bool
    name: str
    id: str

    @classmethod
    def from_monitor_handle(cls, handle):
        """Create a monitor from a handle for a monitor

        Parameters
        ---------
        handle : int
            Monitor handle from EnumDisplayMonitors

        Notes
        -----
        See https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getmonitorinfoa
        and
        https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumdisplaydevicesa
        for more information.
        """
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
    """Get current monitors

    Returns
    -------
    monitors : list[Monitor]
        Current monitors
    """
    monitors = []
    for handle, *_ in win32api.EnumDisplayMonitors():
        monitors.append(Monitor.from_monitor_handle(handle))
    return monitors


def get_windows() -> dict[str, UIAWrapper]:
    """Get the current running applications with visible windows

    Returns
    -------
    windows : dict[str, UIAWrapper]
        Map names of applications to a uiawrapper object of the window
    """
    windows = {}
    apps = collections.defaultdict(list)
    handles_to_windows = {}
    window: UIAWrapper

    # For each application with a moveable window, we need to map each application name to a PID.
    # Multiple handles with the same application name (i.e. two running instances of explorer) is
    # expected. We handle this by first creating a list of windows of the same application name and
    # then we enumerate the application names so they are unique. We use the handles to keep the
    # windows sorted in the same order on every call.

    for window in Desktop(backend="uia").windows():
        text = window.window_text()
        # Only use applications that have window text and are not Program Manager or Task Bar. Those
        # applications always appear but cannot be moved.
        if text and text not in {"Program Manager", "Taskbar"}:
            # Get the application name from the pid rather than the window text. The window text can
            # changes by the state of the application. For example, explorer's window text is based
            # on the current folder - however we want to get explorer.exe
            pid = window.process_id()
            app = psutil.Process(window.process_id()).name()
            apps[app].append(window.handle)
            handles_to_windows[window.handle] = window

    for app, handles in apps.items():
        handles = sorted(handles)
        if len(handles) == 1:
            windows[app] = handles_to_windows[handles[0]]
        else:
            for n, pid in enumerate(handles):
                name = f"{app} - {n + 1}"
                windows[name] = handles_to_windows[pid]

    dpg_core.log_debug(f"windows: {windows}")

    return windows


def get_window_borders(handle: int) -> tuple[float, float, float, float]:
    """Get the borders for the given window

    Most windows will have some invisible pixels around the edges of the windows. Some have 16
    pixels around each edge while others have 14 pixels around 3 of the 4 edges. These pixels
    borders will be invisible and will gaps between windows if not accounted for. See this
    stackoverflow
    https://stackoverflow.com/questions/34139450/getwindowrect-returns-a-size-including-invisible-borders
    answer for more details.
    """
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


def move_window(handle: int, x: int, y: int, width: int, height: int) -> None:
    """Move a window to a given position and size

    Unfortunately you can't just call ``user32.MoveWindow`` to move the window exactly where you
    want it. First, you have restore the window as moving windows that are maximized will not move.
    Next, get the the window borders to account for invisible pixels around the window. Last, we
    move and resize the windows.
    """
    left_border, right_border, top_border, bottom_border = get_window_borders(handle)

    handle = ctypes.wintypes.HWND(int(handle))
    user32.ShowWindow(handle, win32con.SW_RESTORE)

    x, y, height, width = int(x), int(y), int(height), int(width)
    # We can't move and resize the window at the same time. This is probably a bug in the windows
    # API but moving and then resizing seems to be a good work around.
    move_flags = win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE | win32con.SWP_NOSIZE
    user32.SetWindowPos(handle, None, x - left_border, y - top_border, 0, 0, move_flags)

    size_flags = win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE | win32con.SWP_NOMOVE
    user32.SetWindowPos(
        handle,
        None,
        0,
        0,
        width + left_border + right_border,
        height + top_border + bottom_border,
        size_flags,
    )


class UniqueContainer:
    """Parent class simplify the management of dynamic widgets

    When using dynamic widgets in dearpygui, we need to ensure that widgets are uniquely named with
    a parent. This class makes it simpler to keep track of the parent and make unique names.

    Parameters
    ----------
    parent : str
        The widget's parent, None by default
    """

    def __init__(self, parent: str = None) -> None:
        self._parent = parent
        self._id = str(uuid4())

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}({self.id})"

    @property
    def parent(self) -> str:
        """str : The widget's parent"""
        return self._parent

    @property
    def id(self) -> str:
        """str : The widget's unique id"""
        return self._id

    def create_label(self, show: str, hide: str) -> str:
        """Create a unique label for a widget

        A label is constructed with a shown part and a hidden part separated by "##". The hidden
        part will always include the unique container's id.

        Parameters
        ----------
        show : str
            The part of the label that is shown
        hide : str
            Part of the label that is hidden, this is combined with the unique container's ID
        """
        if hide:
            hide = f"{hide}-{self.id}"
        else:
            hide = self.id
        return f"{show}##{hide}"


class AppTable(UniqueContainer):
    """Table for selecting which apps go with which grid number

    Parameters
    ----------
    parent : str
        The parent to the widgets in the container
    """

    HEADER = ["Number", "Application"]
    ACTIVE_WINDOWS = {}

    def __init__(self, parent: str):
        super().__init__(parent)
        self._allocated_windows = set()
        self._available_windows = []
        self._grid_mapping = defaultdict(list)
        self._nrows = 0

    @property
    def allocated_windows(self):
        """set : windows that have already been selected"""
        return self._allocated_windows.copy()

    def init_ui(self):
        """Initialize the container's UI"""
        with dpg_simple.managed_columns(f"{self._id}_head", len(self.HEADER), parent=self.parent):
            for item in self.HEADER:
                dpg_core.add_text(item)

    def clear(self):
        """Clear the table of all rows"""
        for row in reversed(range(1, self._nrows + 1)):
            label = f"{self._id}_{row}"
            dpg_core.delete_item(label)
        self._nrows = 0

    def selected(self, sender, data):
        """Handle when a row is selected in the table"""
        # sender will be the table itself. get_table_selections will return a list of selected table
        # coordinates. We can use these table coordinates to then get names of the selected
        # applications.
        selections = dpg_core.get_table_selections(sender)
        selected_items = [dpg_core.get_table_item(sender, *s) for s in selections]

        number = int(sender.split("_")[1])
        self._grid_mapping[number] = selected_items
        self._allocated_windows = set(
            reduce(lambda s1, s2: set(s1) | set(s2), self._grid_mapping.values())
        )
        self.refresh_available_windows()

    def refresh_available_windows(self, *args, **kwargs):
        """Refresh each grid table with available tables

        After an app is selected, other grids should not be able to select it as well.
        """
        dpg_core.log_debug(f"Refreshing available windows for table {self.id}")
        active_windows = set(self.ACTIVE_WINDOWS)
        available_windows = active_windows - self._allocated_windows
        grid_mapping = defaultdict(list)
        for row in range(1, self._nrows + 1):
            dpg_core.clear_table(f"{self._id}_{row}_table")

            selected_active_windows = set(self._grid_mapping[row]) & active_windows
            grid_mapping[row] = list(selected_active_windows)

            windows = sorted(list(available_windows | selected_active_windows))
            for n, name in enumerate(windows):
                dpg_core.add_row(f"{self._id}_{row}_table", [name])

                if name in selected_active_windows:
                    dpg_core.set_table_selection(f"{self._id}_{row}_table", n, 0, True)

        self._grid_mapping = grid_mapping
        dpg_core.log_debug(f"Refreshed available windows for table {self.id}")

    def set_rows(self, nrows):
        """ "Set the rows in the table

        Each row has two columns. The first column is the grid number. The second column is a list
        of applications to snap to the grid. We use a table to enable multiselect
        """
        dpg_core.log_info(f"Refreshing rows for table {self.id}")
        for row in range(1, nrows + 1):
            name = f"{self._id}_{row}"
            # If the row already exists, we don't need to do anything else
            if dpg_core.does_item_exist(name):
                continue

            with dpg_simple.managed_columns(name, len(self.HEADER), parent=self.parent):
                # The first column is the grid number
                dpg_core.add_input_int(
                    f"##{self._id}_{row}_number",
                    default_value=row,
                    readonly=True,
                    step=0,
                    parent=name,
                )

                # The second column is the table. Wrap in a collapsing header so the screen isn't
                # too full the entire time.
                with dpg_simple.collapsing_header(f"##{self._id}_{row}_header", parent=name):
                    dpg_core.add_table(
                        f"{self._id}_{row}_table",
                        [""],  # no headers
                        callback=self.selected,
                        parent=f"##{self._id}_{row}_header",
                    )
                    # populate the table with the names of available windows
                    for window_name in sorted(self.ACTIVE_WINDOWS):
                        dpg_core.add_row(f"{self._id}_{row}_table", [window_name])

            # Separate each row with a line
            dpg_core.add_separator(name=f"{self._id}_{row}_sep", parent=name)

        self._nrows = nrows
        dpg_core.log_info(f"Refreshed rows for table {self.id}")


class MonitorProfile(UniqueContainer):
    """Each monitor will have it's own unique setup

    The profile is broken into two columns. The first column has the grid inputs, snap button, and
    plot to customize the grid. The second column has the application table to select which apps
    snap to each grid.

    Parameters
    ----------
    parent : str
        The parent to the widgets in the container
    monitor : Monitor
        The monitor this profile belongs to
    """

    def __init__(self, parent: str, monitor: Monitor) -> None:
        super().__init__(parent=parent)
        self._monitor = monitor
        self._ylines = []
        self._xlines = []
        self._labels = []
        self._rectangle_mapping = {}
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

    @property
    def monitor(self):
        """Monitor : the monitor this profile belongs to"""
        return self._monitor

    @property
    def app_table(self):
        """AppTable : The application table to select apps to snap to grids"""
        return self._app_table

    def set_labels(self):
        """Label each grid in the plot

        Grids are labeled in increasing order from left to right then top to bottom. Include below
        the label the percent amount each grid takes up width and height wise.
        """
        # First remove all the labels from the plot
        for label in self._labels:
            dpg_core.delete_annotation(self._plot_id, label)
            dpg_core.delete_annotation(self._plot_id, f"{label}%")

        x0 = self.monitor.work.left
        xs = []
        x_percents = []
        xline_values = sorted([dpg_core.get_value(xline) for xline in self._xlines])
        for x1 in itertools.chain(xline_values, [self.monitor.work.right]):
            # Get the x coordinate of the center point of the grid
            xs.append(((x0 + x1) / 2))
            # Get the percent of the width this x grid consumes
            x_percents.append(((x1 - x0) / self.monitor.work.width) * 100)
            x0 = x1

        ys = []
        y_percents = []
        y0 = self.monitor.work.top
        yline_values = sorted([dpg_core.get_value(yline) for yline in self._ylines])
        for y1 in itertools.chain(yline_values, [self.monitor.work.bottom]):
            ys.append(((y0 + y1) / 2))
            y_percents.append(((y1 - y0) / self.monitor.work.height) * 100)
            y0 = y1

        # Place each grids label in the center of each grid.
        number = 1
        labels = []
        for y, y_percent in zip(reversed(ys), y_percents):
            for x, x_percent in zip(xs, x_percents):
                tag = str(uuid4())
                dpg_core.add_annotation(self._plot_id, str(number), x, y, 0, 0, tag=tag)
                dpg_core.add_annotation(
                    self._plot_id, f"{x_percent:.2f}x{y_percent:.2f}", x, y, 0, 2, tag=f"{tag}%"
                )
                labels.append(tag)
                number += 1
        self._labels = labels

        # Map each grid number to the grids rectangle. These rectangles will be used to snap the
        # window into place
        self._rectangle_mapping.clear()
        y_edges = [self.monitor.work.top] + yline_values + [self.monitor.work.bottom]
        x_edges = [self.monitor.work.left] + xline_values + [self.monitor.work.right]
        grid = 1
        for x_index in range(len(x_edges) - 1):
            x0, x1 = x_edges[x_index], x_edges[x_index + 1]
            for y_index in range(len(y_edges) - 1):
                y0, y1 = y_edges[y_index], y_edges[y_index + 1]
                self._rectangle_mapping[grid] = Rectangle(x0, x1, y0, y1)
                grid += 1

    def line_callback(self, sender, data):
        """Whenever a line is moved, reset the labels"""
        # Since the windows can only be snapped to integer positions, ensure the grid lines are
        # always at integer values
        dpg_core.set_value(sender, int(dpg_core.get_value(sender)))
        self.set_labels()

    def input_callback(self, sender, data):
        """Callback when the input grids are changed"""
        dpg_core.log_info(
            f"Refreshing grid for monitor profile {self.monitor.name} as input changed"
        )
        # First remove each line from the plot
        for xline in self._xlines:
            dpg_core.delete_drag_line(self._plot_id, xline)
        for yline in self._ylines:
            dpg_core.delete_drag_line(self._plot_id, yline)

        # Ensure that the values are greater than or equal to 1
        rows, cols = dpg_core.get_value(self._input_id)
        rows = 1 if rows < 1 else rows
        cols = 1 if cols < 1 else cols
        dpg_core.set_value(self._input_id, [rows, cols])

        # Add horizontal lines to the plot
        ylines = []
        for row in range(1, rows):
            name = f"yline{row}-{self.id}"
            pos = int(self.monitor.work.top + (row / rows) * self.monitor.work.height)
            dpg_core.add_drag_line(
                self._plot_id,
                name,
                y_line=True,
                show_label=False,
                default_value=pos,
                callback=self.line_callback,
            )
            ylines.append(name)

        # Add vertical lines to the plot
        xlines = []
        for col in range(1, cols):
            name = f"xline{col}-{self.id}"
            pos = int(self.monitor.work.left + (col / cols) * self.monitor.work.width)
            dpg_core.add_drag_line(
                self._plot_id,
                name,
                y_line=False,
                show_label=False,
                default_value=pos,
                callback=self.line_callback,
            )
            xlines.append(name)

        self._xlines = xlines
        self._ylines = ylines

        # Reset the labels and application table to reflect the new grids
        self.set_labels()
        self._app_table.clear()
        self._app_table.set_rows(rows * cols)
        dpg_core.log_info(
            f"Refreshed grid for monitor profile {self.monitor.name} as input changed"
        )

    def snap(self, sender, data):
        """Snap each window to the selected grid"""
        dpg_core.log_info(f"Snapping windows in monitor profile {self.monitor.name}")
        for number, windows in self._app_table._grid_mapping.items():
            if number not in self._rectangle_mapping:
                continue
            rect = self._rectangle_mapping[number]
            for name in windows:
                window = AppTable.ACTIVE_WINDOWS[name]
                dpg_core.log_debug(f"Snapping window {window}")
                move_window(
                    window.handle, int(rect.left), int(rect.top), int(rect.width), int(rect.height)
                )
        dpg_core.log_info(f"Snapped windows in monitor profile {self.monitor.name}")

    def resize(self):
        """When the main window is resized, we need to resize the two columns

        The plot tends to be greedy so we have to ensure it on;y takes up half the width of the
        main window.
        """
        window_width = dpg_simple.get_item_width("Main Window")
        plot_width = window_width // 2
        table_width = window_width - plot_width
        dpg_simple.set_item_width(self._left_panel_id, plot_width)
        dpg_simple.set_item_width(self._right_panel_id, table_width)

    def init_ui(self):
        """Initialize the container's UI"""
        # the monitor profile is in a tab
        with dpg_simple.tab(self._tab_id, parent=self.parent, no_tooltip=True):
            # Create 2 groups and put them on the same line
            with dpg_simple.group(self._left_panel_id, parent=self._tab_id):
                # The grid input
                dpg_core.add_input_int2(
                    self._input_id,
                    parent=self._left_panel_id,
                    callback=self.input_callback,
                    default_value=[1, 1],
                )
                # snap button
                dpg_core.add_button(self._snap_id, parent=self._left_panel_id, callback=self.snap)
                # Customize grid with the plot
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
                # Ensure the plot's limits are the work area
                dpg_core.set_plot_xlimits(
                    self._plot_id, xmin=self.monitor.work.left, xmax=self.monitor.work.right
                )
                dpg_core.set_plot_ylimits(
                    self._plot_id, ymin=self.monitor.work.top, ymax=self.monitor.work.bottom
                )

            # Put the application table on the right
            dpg_core.add_same_line(parent=self._tab_id)
            with dpg_simple.group(self._right_panel_id, parent=self._tab_id):
                self._app_table = AppTable(parent=self._right_panel_id)

    def to_dict(self) -> dict[str, list[float]]:
        """Serialize the monitor profile

        Returns
        -------
        serialized_monitor_profile : dict[str, list[float]]
            Serialized monitor profile
        """
        dpg_core.log_info(f"Serializing monitor profile: {self.id}")
        serialized_monitor_profile = {
            "xlines": sorted(
                dpg_core.get_value(xline) / self.monitor.work.height for xline in self._xlines
            ),
            "ylines": sorted(
                dpg_core.get_value(yline) / self.monitor.work.width for yline in self._ylines
            ),
        }
        dpg_core.log_info(f"Serialized monitor profile: {self.id}")
        return serialized_monitor_profile

    def load_dict(self, serialized_monitor_profile: dict[str, list[float]]) -> None:
        """Set the monitor profile from a serialized dictionary

        Parameters
        ----------
        serialized_monitor_profile : dict[str, list[float]]
            Dictionary to load
        """
        dpg_core.log_info(f"Loading monitor profile {self.id}")
        for xline in self._xlines:
            dpg_core.delete_drag_line(self._plot_id, xline)
        for yline in self._ylines:
            dpg_core.delete_drag_line(self._plot_id, yline)

        dpg_core.log_debug("Setting loaded ylines")
        ylines = []
        for row, value in enumerate(serialized_monitor_profile["ylines"]):
            dpg_core.log_debug("Loading y lines")
            yline = int(value * self.monitor.work.height)
            name = f"yline{row}-{self.id}"
            dpg_core.add_drag_line(
                self._plot_id,
                name,
                y_line=True,
                show_label=False,
                default_value=yline,
                callback=self.line_callback,
            )
            ylines.append(name)

        dpg_core.log_debug("Setting loaded xlines")
        xlines = []
        for col, value in enumerate(serialized_monitor_profile["xlines"]):
            yline = int(value * self.monitor.work.width)
            name = f"xline{col}-{self.id}"
            dpg_core.add_drag_line(
                self._plot_id,
                name,
                y_line=False,
                show_label=False,
                default_value=yline,
                callback=self.line_callback,
            )
            xlines.append(name)

        self._xlines = xlines
        self._ylines = ylines

        self.set_labels()
        rows, cols = (len(self._xlines) + 1), (len(self._ylines) + 1)
        dpg_core.set_value(self._input_id, [rows, cols])
        self._app_table.set_rows(rows * cols)
        dpg_core.log_info(f"Loaded monitor profile {self.id}")


class Profile(UniqueContainer):
    """Setup for one or more monitor profiles

    Parameters
    ----------
    parent : str
        The parent to the widgets in the container
    monitors : list[Monitor]
        The monitors in this profile
    """

    def __init__(self, parent, monitors):
        super().__init__(parent)
        self._monitors = monitors
        # filled in init_ui
        self._monitor_profiles: dict[str, MonitorProfile] = {}
        self._tab_bar_id = self.create_label("", "tab-bar")
        self.init_ui()

    def iter_monitor_profiles(self):
        """Iterate over the monitors in this profile

        Yields
        ------
        monitor : Monitor
            monitor in this profile
        """
        yield from self._monitor_profiles.values()

    def init_ui(self):
        """Initializes this UI

        A profile is made up of tabs of monitor profiles
        """
        with dpg_simple.tab_bar(self._tab_bar_id, parent=self.parent):
            for monitor in self._monitors:
                monitor_profile = MonitorProfile(self._tab_bar_id, monitor)
                self._monitor_profiles[monitor.id] = monitor_profile

    def to_dict(self) -> dict[str, dict[str, list[float]]]:
        """Serialize profile to a dictionary

        Returns
        -------
        serialized_profile : dict[str, dict[str, list[float]]]
            The serialized profile
        """
        dpg_core.log_info(f"Serializing profile {self.id}")
        serialized_profile = {}
        for monitor_id, monitor_profile in self._monitor_profiles.items():
            serialized_profile[monitor_id] = monitor_profile.to_dict()
        dpg_core.log_info(f"Serialized profile {self.id}")
        return serialized_profile

    def load_dict(self, serialized_profile: dict[str, dict[str, list[float]]]):
        """Load a serialized profile

        Parameters
        ----------
        serialized_profile : dict[str, dict[str, list[float]]]
            The serialized profile
        """
        dpg_core.log_info(f"Loading profile {self.id}")
        for monitor_id, serialized_monitor_profile in serialized_profile.items():
            monitor_profile = self._monitor_profiles[monitor_id]
            monitor_profile.load_dict(serialized_monitor_profile)
        dpg_core.log_info(f"Loaded profile {self.id}")


AppTable.ACTIVE_WINDOWS = get_windows()


class MainWindow(UniqueContainer):
    """WinSnap's main window"""

    PROFILES_PATH = Path(os.path.expandvars("%APPDATA%")) / "WinSnap" / "profiles.json"

    def __init__(self) -> None:
        super().__init__(parent=None)
        self._profiles: dict[str, Profile] = {}
        self._tab_number = 1
        self._monitors = get_monitors()
        self._save_id = "Save##MainWindow-save"
        self._load_id = "Load##MainWindow-load"

        # Load the serialized profiles on startup
        if self.PROFILES_PATH.exists():
            with self.PROFILES_PATH.open("r") as stream:
                self._saved_profiles = json.load(stream)
        else:
            if not self.PROFILES_PATH.parent.exists():
                os.mkdir(self.PROFILES_PATH.parent)
            self._saved_profiles = {}

        self.init_ui()

    def get_or_create_monitor_set_id(self):
        """Get or create a monitor set id

        We assign each unique set of monitors their own ID so we can save different profiles to
        each set. A layout on one set of monitors does not perfectly translate to a good layout
        on another set so we give each unique set of monitors an id and it's own saved profile.

        Returns
        -------
        monitor_set_id : str
            UUID string for the monitor set
        """
        monitor_sets = self._saved_profiles.get("monitor_sets", {})
        current_monitor_ids = set(monitor.id for monitor in self._monitors)
        for monitor_set_id, monitor_ids in monitor_sets.items():
            if set(monitor_ids) == current_monitor_ids:
                dpg_core.log_debug(f"Found existing monitor set: {monitor_set_id}")
                break
        else:
            monitor_set_id = str(uuid4())
            monitor_sets[monitor_set_id] = list(current_monitor_ids)
            self._saved_profiles["monitor_sets"] = monitor_sets
            dpg_core.log_info(f"Create new monitor set: {monitor_set_id}")
        return monitor_set_id

    def save(self, *args, **kwargs):
        """Save the current layout to the profiles file for the set of monitors"""
        dpg_core.log_info("Saving configuration")
        serialized_profiles = [profile.to_dict() for profile in self._profiles.values()]
        monitor_set_id = self.get_or_create_monitor_set_id()
        self._saved_profiles[monitor_set_id] = serialized_profiles
        with self.PROFILES_PATH.open("w") as stream:
            json.dump(self._saved_profiles, stream)
        dpg_core.log_info(f"Successfully saved configuration {monitor_set_id}")

    def load(self, *args, **kwargs):
        """Load a the serialized profiles for the current set of monitors"""
        dpg_core.log_info("Loading save profile")
        monitor_set_id = self.get_or_create_monitor_set_id()
        dpg_core.log_debug(f"Monitor set ID: {monitor_set_id}")
        serialized_profiles = self._saved_profiles.get(monitor_set_id)
        if not serialized_profiles:
            dpg_core.log_debug(f"No serialized profile for monitor set ID {monitor_set_id}")
            return
        dpg_core.log_debug(f"Found serialized profile for monitor set ID {monitor_set_id}")

        for label in self._profiles:
            dpg_core.delete_item(label)
        self._profiles = {}
        self._tab_number = 1

        for serialized_profile in serialized_profiles:
            profile = self.add_tab()
            # Ensure the plot width is accurate
            self.resize_callback(None, None)
            profile.load_dict(serialized_profile)

        self.resize_callback(None, None)
        dpg_core.log_info(f"Successfully loaded saved profile {monitor_set_id}")

    def add_tab(self, *args, **kwargs):
        """Add a profile tab"""
        dpg_core.log_debug("Adding profile tab...")
        label = f"{self._tab_number}##MainWindow-tab{self._tab_number}"
        with dpg_simple.tab(label, parent="##MainWindow-tabbar", closable=False, no_tooltip=True):
            profile = Profile(label, self._monitors)
            self._profiles[label] = profile
        self._tab_number += 1

        # If we previously only had one tab then we need to make the first tab closable
        if len(self._profiles) == 2:
            label = list(self._profiles)[0]  # get the label of the first tab
            dpg_core.configure_item(label, closable=dpg_core.get_value(label))

        dpg_core.log_info(f"Profile {label} successfully added")
        return profile

    def remove_tab(self, *args, **kwargs):
        """Close a tab"""
        # Due to https://github.com/hoffstadt/DearPyGui/issues/429, when a tab is closed, we have to
        # search for the tab to remove the profile itself
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
        dpg_core.log_info(f"Profiles {remove} successfully closed")

    def init_ui(self):
        """Initialize the main window's UI

        The main window has save and load buttons and a tab bar for the profiles.
        """
        with dpg_simple.window("Main Window"):
            dpg_core.add_button(self._save_id, callback=self.save)
            dpg_core.add_same_line()
            dpg_core.add_button(self._load_id, callback=self.load)

            main_window_tab_bar_ctx = dpg_simple.tab_bar(
                "##MainWindow-tabbar",
                parent="Main Window",
                reorderable=True,
                callback=self.remove_tab,
            )
            with main_window_tab_bar_ctx:
                self.add_tab()
                dpg_core.add_tab_button(
                    "+##MainWindow-btn",
                    parent="##MainWindow-tabbar",
                    callback=self.add_tab,
                    trailing=True,
                    no_tooltip=True,
                )

    def resize_callback(self, *args, **kwargs):
        """Handle when the main window is resized"""
        for profile in self._profiles.values():
            for monitor_profile in profile.iter_monitor_profiles():
                monitor_profile.resize()

    def refresh_windows(self):
        """Refresh the currently active windows"""
        new_active_windows = get_windows()
        if new_active_windows != AppTable.ACTIVE_WINDOWS:
            dpg_core.log_debug("Active windows changed, refreshing tables")
            AppTable.ACTIVE_WINDOWS = new_active_windows
            for profile in self._profiles.values():
                for monitor_profile in profile.iter_monitor_profiles():
                    monitor_profile._app_table.refresh_available_windows()


def main():
    parser = argparse.ArgumentParser(description="Winsnap: Snap windows into a customizable grid")
    parser.add_argument("--debug", action="store_true", help="Show the log window")
    args = parser.parse_args()
    main_window = MainWindow()

    def mouse_click_cb(sender, data):
        main_window.refresh_windows()

    def startup_cb(sender, data):
        main_window.load()

    dpg_core.set_mouse_click_callback(mouse_click_cb)
    dpg_core.set_start_callback(startup_cb)
    dpg_core.set_resize_callback(main_window.resize_callback, handler="Main Window")
    scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
    dpg_core.set_global_font_scale(scale)
    w, h = dpg_core.get_main_window_size()
    dpg_core.set_main_window_size(int(w * scale), int(h * scale))
    dpg_core.set_main_window_title("WinSnap")
    if args.debug:
        dpg_core.show_logger()
    dpg_core.start_dearpygui(primary_window="Main Window")


if __name__ == "__main__":
    main()
