"""
mvr_reader.py - Reads the Parts table from the Pinnacle MVR (Motor Vehicle Record) window.

Connects to an open MVR window via Java Access Bridge, clicks the Parts tab,
and returns all un-priced parts that have a Hollander interchange number.
"""

import ctypes
import ctypes.wintypes as wintypes
import time

from pywinauto import Desktop

from pinnacle_reader import (
    AccessibleContextInfo,
    AccessibleTableInfo,
    AccessibleTableCellInfo,
    AccessibleTextInfo,
    JOBJECT64,
)


# Column indices in the Parts table (confirmed from live inspection)
_COL_CATEGORY    = 0   # Hollander interchange category name (e.g. "2nd Seat (Rear Seat)")
_COL_DESCRIPTION = 1   # Full description ("RH TOP FOLD DOWN, BLACK, WNY'S EURO PART…")
_COL_PRICE       = 4
_COL_LOCATION    = 6
_COL_GRADE       = 9
_COL_STOCK_NUM   = 11
_COL_HOLLANDER   = 14


def _extract_part_name(description: str) -> str:
    """Return the primary part name from a Pinnacle description string.

    Descriptions are comma-separated; the part name is the first segment.
    Returns an empty string if the name looks ambiguous or unresolvable.
    """
    part = description.split(',')[0].strip()
    if 'CONFIRM ID' in part.upper() or '!!' in part or not part:
        return ''   # flag for manual review
    return part


class MVRReader:
    """Reads the Parts table from a Pinnacle MVR window via Java Access Bridge."""

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

    def _find_mvr_hwnd(self, timeout=15, pre_click_titles: dict = None):
        """Poll for an MVR window that is new or has a changed title since the click.

        *pre_click_titles* is a {hwnd: title} snapshot taken before the
        double-click (from get_open_mvr_titles()).  A window qualifies if:
          - its HWND was not in the snapshot (brand-new window), OR
          - its HWND was in the snapshot but its title has changed (Pinnacle
            reused the window for the newly opened vehicle).

        If *pre_click_titles* is None, any 'Vehicle:' window is accepted
        (original behaviour — safe when no MVR is already open).
        """
        old = pre_click_titles or {}
        deadline = time.time() + timeout
        while time.time() < deadline:
            desktop = Desktop(backend='uia')
            for w in desktop.windows():
                try:
                    title = w.window_text()
                    cls = w.element_info.class_name
                    if 'SunAwtFrame' not in cls or not title.startswith('Vehicle:'):
                        continue
                    hwnd = w.handle
                    # Brand-new window
                    if hwnd not in old:
                        return hwnd
                    # Same HWND but title changed → Pinnacle loaded the new vehicle
                    if old[hwnd] != title:
                        return hwnd
                except Exception:
                    continue
            time.sleep(0.5)

        # Timeout: the same vehicle was re-opened (title unchanged) or the window
        # opened but the title update was missed.  Fall back to any open MVR window.
        desktop = Desktop(backend='uia')
        for w in desktop.windows():
            try:
                if 'SunAwtFrame' in w.element_info.class_name and w.window_text().startswith('Vehicle:'):
                    return w.handle
            except Exception:
                continue

        raise RuntimeError(
            'No MVR window ("Vehicle: …") found within '
            f'{timeout} seconds. Make sure the vehicle was opened in Pinnacle.'
        )

    def _connect_jab(self, hwnd):
        """Get JAB vmID and accessible context for a window handle."""
        vmID = ctypes.c_long()
        ac = JOBJECT64()
        result = self.jab.getAccessibleContextFromHWND(
            hwnd, ctypes.byref(vmID), ctypes.byref(ac))
        if not result:
            raise RuntimeError('Java Access Bridge could not connect to the MVR window.')
        return vmID.value, ac.value

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

    def _find_node(self, vm, ctx, role=None, name=None, require_visible=True,
                   depth=0, max_depth=25):
        """Recursively search the accessible tree for a node matching criteria."""
        ci = AccessibleContextInfo()
        if not self.jab.getAccessibleContextInfo(vm, ctx, ctypes.byref(ci)):
            return None

        matches_role = (role is None) or (ci.role == role)
        matches_name = (name is None) or (ci.name == name)
        is_visible = (not require_visible) or (ci.x > 0)

        if matches_role and matches_name and is_visible:
            return ctx

        if ci.childrenCount > 0 and depth < max_depth:
            for i in range(min(ci.childrenCount, 100)):
                child = self.jab.getAccessibleChildFromContext(vm, ctx, i)
                if child:
                    result = self._find_node(
                        vm, child, role=role, name=name,
                        require_visible=require_visible,
                        depth=depth + 1, max_depth=max_depth,
                    )
                    if result:
                        return result
        return None

    def _click_parts_tab(self, vm, ac):
        """Click the 'Parts' tab in the MVR window."""
        parts_node = self._find_node(vm, ac, role='label', name='Parts')
        if parts_node is None:
            # Also try role='page tab' which Swing sometimes reports
            parts_node = self._find_node(vm, ac, role='page tab', name='Parts')
        if parts_node is None:
            raise RuntimeError(
                'Could not find the "Parts" tab in the MVR window. '
                'The tab may have a different name or the MVR layout has changed.'
            )

        ci = AccessibleContextInfo()
        self.jab.getAccessibleContextInfo(vm, parts_node, ctypes.byref(ci))
        x = ci.x + ci.width // 2
        y = ci.y + ci.height // 2

        user32 = ctypes.windll.user32
        user32.SetCursorPos(x, y)
        user32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFTDOWN
        user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFTUP
        time.sleep(1)  # wait for Parts tab to render

    def _find_parts_table(self, vm, ac):
        """Recursively find the first visible table in the accessible tree."""
        ci = AccessibleContextInfo()
        if not self.jab.getAccessibleContextInfo(vm, ac, ctypes.byref(ci)):
            return None
        if ci.role == 'table' and ci.x > 0:
            return ac
        if ci.childrenCount > 0:
            for i in range(min(ci.childrenCount, 100)):
                child = self.jab.getAccessibleChildFromContext(vm, ac, i)
                if child:
                    result = self._find_parts_table(vm, child)
                    if result:
                        return result
        return None

    def read_unpriced_parts(self, pre_click_titles: dict = None) -> list[dict]:
        """Open the MVR window, navigate to the Parts tab, and return un-priced parts.

        A part is included if:
          - price == 0.0
          - hollander interchange number is not empty

        Returns a list of dicts with keys:
          description, part_name, hollander, price, location, grade, stock_num
        """
        hwnd = self._find_mvr_hwnd(pre_click_titles=pre_click_titles)
        vm, ac = self._connect_jab(hwnd)

        self._click_parts_tab(vm, ac)

        # Re-fetch context after tab click so the tree reflects the Parts panel
        vm, ac = self._connect_jab(hwnd)
        table_ctx = self._find_parts_table(vm, ac)

        if not table_ctx:
            raise RuntimeError(
                'Could not find the Parts table in the MVR window after clicking the Parts tab.'
            )

        ti = AccessibleTableInfo()
        if not self.jab.getAccessibleTableInfo(vm, table_ctx, ctypes.byref(ti)):
            raise RuntimeError('Could not read Parts table info from the MVR.')

        parts = []
        for row in range(ti.rowCount):
            category    = self._get_cell_value(vm, ti.accessibleTable, row, _COL_CATEGORY)
            description = self._get_cell_value(vm, ti.accessibleTable, row, _COL_DESCRIPTION)
            price_str   = self._get_cell_value(vm, ti.accessibleTable, row, _COL_PRICE)
            location    = self._get_cell_value(vm, ti.accessibleTable, row, _COL_LOCATION)
            grade       = self._get_cell_value(vm, ti.accessibleTable, row, _COL_GRADE)
            stock_num   = self._get_cell_value(vm, ti.accessibleTable, row, _COL_STOCK_NUM)
            hollander   = self._get_cell_value(vm, ti.accessibleTable, row, _COL_HOLLANDER)

            # Must have a Hollander number to be priceable
            if not hollander.strip():
                continue

            # Must be currently un-priced
            try:
                price_val = float(price_str.replace('$', '').replace(',', '').strip() or '0')
            except ValueError:
                price_val = 0.0

            if price_val != 0.0:
                continue

            part_name = _extract_part_name(description)
            # Normalize the Hollander category name (remove leading newline Pinnacle adds)
            category = category.strip()

            parts.append({
                'description': description,
                'part_name': part_name,
                'category': category,
                'hollander': hollander.strip(),
                'price': price_str,
                'location': location,
                'grade': grade,
                'stock_num': stock_num,
            })

        return parts


def get_open_mvr_titles() -> dict:
    """Return {hwnd: title} for any MVR ('Vehicle: …') windows currently open."""
    titles = {}
    desktop = Desktop(backend='uia')
    for w in desktop.windows():
        try:
            title = w.window_text()
            if 'SunAwtFrame' in w.element_info.class_name and title.startswith('Vehicle:'):
                titles[w.handle] = title
        except Exception:
            pass
    return titles


def open_mvr_and_read_parts(pre_click_titles: dict = None) -> list[dict]:
    """Connect to a newly opened (or updated) MVR window and return all un-priced parts.

    Pass *pre_click_titles* (from get_open_mvr_titles()) so the reader can
    detect whether Pinnacle opened a brand-new window OR reused an existing
    one with an updated title.
    """
    reader = MVRReader()
    return reader.read_unpriced_parts(pre_click_titles=pre_click_titles)
