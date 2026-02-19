"""
pinnacle_reader.py - Extracts VINs from Pinnacle Professional via Java Access Bridge.

Connects to the running Pinnacle Professional application, locates the
Find Vehicles table via the Java Access Bridge API, and reads VIN values
from column index 3.
"""

import ctypes
import ctypes.wintypes as wintypes
import re
import sys
import time

from pywinauto import Desktop


# Valid VIN: 17 alphanumeric characters, excluding I, O, Q
VIN_PATTERN = re.compile(r'^[A-HJ-NPR-Z0-9]{17}$')

# JAB types
JOBJECT64 = ctypes.c_int64

VIN_COLUMN_INDEX = 3  # VIN is column 3 in the Find Vehicles table


class AccessibleContextInfo(ctypes.Structure):
    _fields_ = [
        ('name', ctypes.c_wchar * 1024),
        ('description', ctypes.c_wchar * 1024),
        ('role', ctypes.c_wchar * 256),
        ('role_en_US', ctypes.c_wchar * 256),
        ('states', ctypes.c_wchar * 256),
        ('states_en_US', ctypes.c_wchar * 256),
        ('indexInParent', ctypes.c_int),
        ('childrenCount', ctypes.c_int),
        ('x', ctypes.c_int),
        ('y', ctypes.c_int),
        ('width', ctypes.c_int),
        ('height', ctypes.c_int),
        ('accessibleComponent', wintypes.BOOL),
        ('accessibleAction', wintypes.BOOL),
        ('accessibleSelection', wintypes.BOOL),
        ('accessibleText', wintypes.BOOL),
        ('accessibleInterfaces', wintypes.BOOL),
    ]


class AccessibleTableInfo(ctypes.Structure):
    _fields_ = [
        ('caption', JOBJECT64),
        ('summary', JOBJECT64),
        ('rowCount', ctypes.c_int),
        ('columnCount', ctypes.c_int),
        ('accessibleContext', JOBJECT64),
        ('accessibleTable', JOBJECT64),
    ]


class AccessibleTableCellInfo(ctypes.Structure):
    _fields_ = [
        ('accessibleContext', JOBJECT64),
        ('index', ctypes.c_int),
        ('row', ctypes.c_int),
        ('column', ctypes.c_int),
        ('rowExtent', ctypes.c_int),
        ('columnExtent', ctypes.c_int),
        ('isSelected', wintypes.BOOL),
    ]


class AccessibleTextInfo(ctypes.Structure):
    _fields_ = [
        ('charCount', ctypes.c_int),
        ('caretIndex', ctypes.c_int),
        ('indexAtPoint', ctypes.c_int),
    ]


def validate_vin(vin: str) -> bool:
    """Check that a string is a valid 17-character VIN."""
    return bool(VIN_PATTERN.match(vin.upper().strip()))


class JABReader:
    """Reads data from Pinnacle Professional via Java Access Bridge."""

    def __init__(self):
        self.jab = ctypes.windll.LoadLibrary('WindowsAccessBridge-64.dll')
        self._setup_prototypes()
        self.jab.Windows_run()
        self._pump_messages(3)

    def _setup_prototypes(self):
        jab = self.jab
        jab.getAccessibleContextFromHWND.argtypes = [
            wintypes.HWND, ctypes.POINTER(ctypes.c_long), ctypes.POINTER(JOBJECT64)]
        jab.getAccessibleContextFromHWND.restype = wintypes.BOOL
        jab.getAccessibleContextInfo.argtypes = [
            ctypes.c_long, JOBJECT64, ctypes.POINTER(AccessibleContextInfo)]
        jab.getAccessibleContextInfo.restype = wintypes.BOOL
        jab.getAccessibleChildFromContext.argtypes = [
            ctypes.c_long, JOBJECT64, ctypes.c_int]
        jab.getAccessibleChildFromContext.restype = JOBJECT64
        jab.getAccessibleTableInfo.argtypes = [
            ctypes.c_long, JOBJECT64, ctypes.POINTER(AccessibleTableInfo)]
        jab.getAccessibleTableInfo.restype = wintypes.BOOL
        jab.getAccessibleTableCellInfo.argtypes = [
            ctypes.c_long, JOBJECT64, ctypes.c_int, ctypes.c_int,
            ctypes.POINTER(AccessibleTableCellInfo)]
        jab.getAccessibleTableCellInfo.restype = wintypes.BOOL
        jab.getAccessibleTextInfo.argtypes = [
            ctypes.c_long, JOBJECT64, ctypes.POINTER(AccessibleTextInfo),
            ctypes.c_int, ctypes.c_int]
        jab.getAccessibleTextInfo.restype = wintypes.BOOL
        jab.getAccessibleTextRange.argtypes = [
            ctypes.c_long, JOBJECT64, ctypes.c_int, ctypes.c_int,
            ctypes.c_wchar_p, ctypes.c_short]
        jab.getAccessibleTextRange.restype = wintypes.BOOL
        jab.getAccessibleTableColumnHeader.argtypes = [
            ctypes.c_long, JOBJECT64, ctypes.POINTER(AccessibleTableInfo)]
        jab.getAccessibleTableColumnHeader.restype = wintypes.BOOL
        jab.getAccessibleSelectionCountFromContext.argtypes = [ctypes.c_long, JOBJECT64]
        jab.getAccessibleSelectionCountFromContext.restype = ctypes.c_int
        jab.getAccessibleSelectionFromContext.argtypes = [ctypes.c_long, JOBJECT64, ctypes.c_int]
        jab.getAccessibleSelectionFromContext.restype = JOBJECT64

    def _pump_messages(self, seconds):
        """Pump Windows messages so JAB can discover Java VMs."""
        user32 = ctypes.windll.user32
        msg = wintypes.MSG()
        end = time.time() + seconds
        while time.time() < end:
            if user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1):
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
            else:
                time.sleep(0.01)

    def _find_pinnacle_hwnd(self):
        """Find the Find Vehicles window HWND."""
        desktop = Desktop(backend='uia')
        all_wins = desktop.windows()
        for w in all_wins:
            try:
                title = w.window_text()
                cls = w.element_info.class_name
                if 'SunAwtFrame' in cls and 'Find Vehicles' in title:
                    return w.handle
            except Exception:
                continue
        raise RuntimeError(
            'Pinnacle "Find Vehicles" window not found. '
            'Make sure Pinnacle Professional is running with the Find Vehicles screen open.'
        )

    def _connect_jab(self, hwnd):
        """Get JAB vmID and accessible context for a window handle."""
        vmID = ctypes.c_long()
        ac = JOBJECT64()
        result = self.jab.getAccessibleContextFromHWND(
            hwnd, ctypes.byref(vmID), ctypes.byref(ac))
        if not result:
            raise RuntimeError('Java Access Bridge could not connect to Pinnacle.')
        return vmID.value, ac.value

    def _find_table(self, vm, ctx, depth=0, max_depth=20):
        """Recursively find the first table element in the accessible tree."""
        ci = AccessibleContextInfo()
        if not self.jab.getAccessibleContextInfo(vm, ctx, ctypes.byref(ci)):
            return None
        if ci.role == 'table':
            return ctx
        if ci.childrenCount > 0 and depth < max_depth:
            for i in range(min(ci.childrenCount, 50)):
                child = self.jab.getAccessibleChildFromContext(vm, ctx, i)
                if child:
                    result = self._find_table(vm, child, depth + 1, max_depth)
                    if result:
                        return result
        return None

    def _get_cell_text(self, vm, ctx):
        """Get text content from an accessible context."""
        ti = AccessibleTextInfo()
        if self.jab.getAccessibleTextInfo(vm, ctx, ctypes.byref(ti), 0, 0):
            if ti.charCount > 0:
                buf = ctypes.create_unicode_buffer(min(ti.charCount + 1, 2048))
                if self.jab.getAccessibleTextRange(
                    vm, ctx, 0, min(ti.charCount - 1, 2046), buf, len(buf)
                ):
                    return buf.value
        return ''

    def _get_cell_value(self, vm, table_ac, row, col):
        """Read a single cell value from a JAB table."""
        cell_info = AccessibleTableCellInfo()
        if self.jab.getAccessibleTableCellInfo(
            vm, table_ac, row, col, ctypes.byref(cell_info)
        ):
            ci = AccessibleContextInfo()
            if self.jab.getAccessibleContextInfo(
                vm, cell_info.accessibleContext, ctypes.byref(ci)
            ):
                return ci.name or self._get_cell_text(vm, cell_info.accessibleContext) or ''
        return ''

    def open_selected_vehicle(self):
        """Double-click the currently selected row to open the MVR window."""
        hwnd = self._find_pinnacle_hwnd()
        vm, ac = self._connect_jab(hwnd)
        table_ctx = self._find_table(vm, ac)

        if not table_ctx:
            raise RuntimeError('Could not find the vehicles table in Pinnacle.')

        ctx = self.jab.getAccessibleSelectionFromContext(vm, table_ctx, 0)
        if not ctx:
            raise RuntimeError('No row is selected in the Find Vehicles table.')

        ci = AccessibleContextInfo()
        if not self.jab.getAccessibleContextInfo(vm, ctx, ctypes.byref(ci)):
            raise RuntimeError('Could not get context info for selected row.')

        x = ci.x + ci.width // 2
        y = ci.y + ci.height // 2
        user32 = ctypes.windll.user32
        user32.SetCursorPos(x, y)
        for _ in range(2):   # double-click
            user32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFTDOWN
            user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFTUP
            time.sleep(0.05)

    def read_all_vins(self) -> list[dict]:
        """Read all vehicle rows from the Find Vehicles table.

        Returns a list of dicts with keys: stock_num, model, year, vin,
        color, odo, location, engine, trans, doors, site.
        """
        hwnd = self._find_pinnacle_hwnd()
        vm, ac = self._connect_jab(hwnd)
        table_ctx = self._find_table(vm, ac)

        if not table_ctx:
            raise RuntimeError('Could not find the vehicles table in Pinnacle.')

        ti = AccessibleTableInfo()
        if not self.jab.getAccessibleTableInfo(vm, table_ctx, ctypes.byref(ti)):
            raise RuntimeError('Could not read table info.')

        # Column mapping (based on discovered layout)
        col_map = {
            0: 'stock_num', 1: 'model', 2: 'year', 3: 'vin',
            4: 'color', 6: 'odo', 7: 'location',
            9: 'engine', 10: 'trans', 11: 'doors', 12: 'site',
        }

        vehicles = []
        for row in range(ti.rowCount):
            vehicle = {}
            for col_idx, key in col_map.items():
                vehicle[key] = self._get_cell_value(vm, ti.accessibleTable, row, col_idx)
            if vehicle.get('vin') and validate_vin(vehicle['vin']):
                vehicles.append(vehicle)

        return vehicles

    def read_selected_vin(self) -> str:
        """Read the VIN from the currently selected row in the Find Vehicles table.

        Uses getAccessibleSelectionCountFromContext / getAccessibleSelectionFromContext
        to query the table's actual Swing selection model. Falls back to read_all_vins()
        if no valid VIN is found via the selection API.
        """
        hwnd = self._find_pinnacle_hwnd()
        vm, ac = self._connect_jab(hwnd)
        table_ctx = self._find_table(vm, ac)

        if not table_ctx:
            raise RuntimeError('Could not find the vehicles table in Pinnacle.')

        # Ask the table's selection model how many items are selected
        sel_count = self.jab.getAccessibleSelectionCountFromContext(vm, table_ctx)

        for k in range(min(sel_count, 50)):  # cap to avoid runaway on multi-select
            ctx = self.jab.getAccessibleSelectionFromContext(vm, table_ctx, k)
            if not ctx:
                continue

            ci = AccessibleContextInfo()
            if self.jab.getAccessibleContextInfo(vm, ctx, ctypes.byref(ci)):
                # Case A: selection returns cells directly — check if name is a VIN
                text = ci.name or self._get_cell_text(vm, ctx)
                if text and validate_vin(text.strip()):
                    return text.upper().strip()

                # Case B: selection returns a row context — check child at VIN_COLUMN_INDEX
                vin_child = self.jab.getAccessibleChildFromContext(vm, ctx, VIN_COLUMN_INDEX)
                if vin_child:
                    child_ci = AccessibleContextInfo()
                    if self.jab.getAccessibleContextInfo(vm, vin_child, ctypes.byref(child_ci)):
                        text = child_ci.name or self._get_cell_text(vm, vin_child)
                        if text and validate_vin(text.strip()):
                            return text.upper().strip()

        # Fallback: read_all_vins() uses a fresh table traversal
        vehicles = self.read_all_vins()
        if vehicles:
            return vehicles[0]['vin']

        raise RuntimeError('Could not find a valid VIN in the Pinnacle table.')

    def read_vin_by_row(self, row: int) -> str:
        """Read the VIN from a specific row index."""
        hwnd = self._find_pinnacle_hwnd()
        vm, ac = self._connect_jab(hwnd)
        table_ctx = self._find_table(vm, ac)

        if not table_ctx:
            raise RuntimeError('Could not find the vehicles table in Pinnacle.')

        ti = AccessibleTableInfo()
        self.jab.getAccessibleTableInfo(vm, table_ctx, ctypes.byref(ti))

        if row >= ti.rowCount:
            raise RuntimeError(f'Row {row} out of range (table has {ti.rowCount} rows).')

        vin = self._get_cell_value(vm, ti.accessibleTable, row, VIN_COLUMN_INDEX)
        if vin and validate_vin(vin.strip()):
            return vin.upper().strip()

        raise RuntimeError(f'No valid VIN at row {row}.')


# Module-level convenience functions for use by vin_automation.py

def read_vin_from_pinnacle(dump_tree: bool = False) -> str:
    """Read the selected VIN from Pinnacle Professional."""
    reader = JABReader()
    if dump_tree:
        vehicles = reader.read_all_vins()
        print(f'Found {len(vehicles)} vehicles with valid VINs:')
        for v in vehicles[:10]:
            print(f"  {v['stock_num']:>12}  {v['year']}  {v['model']:<20}  {v['vin']}")
        if len(vehicles) > 10:
            print(f'  ... and {len(vehicles) - 10} more')
        return ''
    return reader.read_selected_vin()


def read_all_vins_from_pinnacle() -> list[dict]:
    """Read all VINs and vehicle info from Pinnacle Professional."""
    reader = JABReader()
    return reader.read_all_vins()


if __name__ == '__main__':
    dump = '--dump' in sys.argv
    all_flag = '--all' in sys.argv

    try:
        if all_flag:
            reader = JABReader()
            vehicles = reader.read_all_vins()
            print(f'Found {len(vehicles)} vehicles:')
            for v in vehicles:
                print(f"  {v['stock_num']:>12}  {v['year']}  {v['model']:<20}  {v['vin']}")
        else:
            vin = read_vin_from_pinnacle(dump_tree=dump)
            if vin:
                print(f'VIN: {vin}')
    except RuntimeError as exc:
        print(f'Error: {exc}', file=sys.stderr)
        sys.exit(1)
