"""
vinmatch_to_pinnacle.py - One-click: decode VIN on VINMatchPro → fill Pinnacle MVR.

Usage:
    py -3.12 vinmatch_to_pinnacle.py              # reads VIN from Find Vehicles (selected row)
    py -3.12 vinmatch_to_pinnacle.py --vin 1FM... # use a specific VIN
    py -3.12 vinmatch_to_pinnacle.py --dry-run    # decode and show what would be filled
"""

import argparse
import ctypes
import ctypes.wintypes as wintypes
import re
import sys
import time

from pywinauto import Desktop

from pinnacle_reader import (
    AccessibleContextInfo,
    AccessibleTextInfo,
    JOBJECT64,
    read_vin_from_pinnacle,
)
from vinmatchpro_decoder import decode as vmp_decode


# ─────────────────────────────────────────────────────────────
# Value transformers: VINMatchPro format → Pinnacle values
# ─────────────────────────────────────────────────────────────

def _extract_displacement(engine: str) -> str:
    """'2.5L 4-Cylinder DOHC' → '2.5'"""
    m = re.search(r'(\d+\.?\d*)L', engine or '', re.IGNORECASE)
    return m.group(1) if m else (engine or '')


def _extract_code(color: str) -> str:
    """'UX - INGOT SILVER METALLIC' → 'UX',  '7B - CHARCOAL BLACK' → '7B'"""
    m = re.match(r'^([A-Z0-9]{1,4})\s*[-]\s*', (color or '').strip())
    return m.group(1) if m else (color or '').strip()


def _extract_weight(weight: str) -> str:
    """'3645 lbs.' → '3645'"""
    m = re.search(r'([\d,]+)', weight or '')
    return m.group(1).replace(',', '') if m else (weight or '').strip()


_COLOR_DIRECT = {
    'black', 'blue', 'brown', 'gold', 'green', 'gray', 'maroon',
    'orange', 'pink', 'purple', 'red', 'silver', 'tan', 'white', 'yellow',
}
_COLOR_SYNONYMS = {
    'beige': 'Tan', 'ivory': 'White', 'cream': 'White', 'champagne': 'Tan',
    'charcoal': 'Gray', 'pewter': 'Gray', 'titanium': 'Gray', 'sterling': 'Gray',
    'grey': 'Gray', 'ebony': 'Black', 'midnight': 'Black', 'jet': 'Black',
    'sand': 'Tan', 'khaki': 'Tan', 'cashmere': 'Tan',
    'burgundy': 'Maroon', 'wine': 'Maroon', 'cobalt': 'Blue', 'navy': 'Blue',
    'teal': 'Blue', 'sapphire': 'Blue', 'forest': 'Green', 'olive': 'Green',
    'copper': 'Brown', 'chocolate': 'Brown', 'mocha': 'Brown',
}

def _map_color(color: str) -> str:
    if not color:
        return ''
    c = color.lower().strip()
    if c in _COLOR_DIRECT:
        return c.capitalize()
    for k, v in _COLOR_SYNONYMS.items():
        if k in c:
            return v
    return color.split()[0].capitalize()


def _map_fuel(fuel: str) -> str:
    if not fuel:
        return ''
    f = fuel.lower()
    if 'diesel' in f:
        return 'Diesel'
    if 'electric' in f and ('gas' in f or 'hybrid' in f):
        return 'Gas/Electric'
    if 'electric' in f:
        return 'Electric'
    if 'flex' in f or 'e85' in f:
        return 'Flex Fuel'
    if 'natural gas' in f or 'cng' in f:
        return 'Natural Gas'
    if 'hybrid' in f:
        return 'Gas/Electric'
    return 'Gasoline'


def _map_transmission(trans: str) -> str:
    if not trans:
        return ''
    t = trans.lower()
    if 'manual' in t or 'standard' in t or 'stick' in t:
        return 'Manual'
    return 'Automatic'


def _map_drive(drive: str) -> str:
    if not drive:
        return ''
    d = drive.lower()
    if 'front' in d or 'fwd' in d:
        return 'FWD'
    if 'rear' in d or 'rwd' in d:
        return 'RWD'
    if 'all' in d or 'awd' in d:
        return 'AWD'
    if '4wd' in d or 'four' in d or '4x4' in d:
        return '4WD'
    return drive


# Pinnacle label name → (VINMatchPro key, transformer)
# Label names must match exactly as seen in the JAB accessibility tree
_FIELD_MAP = {
    'Trim Level:':           ('trim',             None),
    'Interior Color:':       ('interior_color',   _map_color),
    'Interior Trim Code:':   ('interior_color',   _extract_code),
    'Axle Code:':            ('axle_ratio',        None),
    'External Color Code:':  ('exterior_color',   _extract_code),
    'Weight:':               ('curb_weight',       _extract_weight),
    'Engine Size:':          ('engine',            _extract_displacement),
    'Number Cylinders:':     ('engine_cylinders',  None),
    'Fuel Type:':            ('fuel_type',         _map_fuel),
    'Transmission:':         ('transmission',      _map_transmission),
    'Drive:':                ('drive_type',        _map_drive),
}


# ─────────────────────────────────────────────────────────────
# Win32 keyboard / mouse helpers
# ─────────────────────────────────────────────────────────────

def _click(x: int, y: int, delay: float = 0.15):
    user32 = ctypes.windll.user32
    user32.SetCursorPos(int(x), int(y))
    time.sleep(0.05)
    user32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFTDOWN
    user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFTUP
    time.sleep(delay)


def _ctrl_a():
    user32 = ctypes.windll.user32
    VK_CONTROL, VK_A = 0x11, 0x41
    user32.keybd_event(VK_CONTROL, 0, 0, 0)
    user32.keybd_event(VK_A, 0, 0, 0)
    user32.keybd_event(VK_A, 0, 0x0002, 0)
    user32.keybd_event(VK_CONTROL, 0, 0x0002, 0)
    time.sleep(0.1)


def _send_vk(vk: int, delay: float = 0.05):
    """Send a single virtual-key press (down + up)."""
    user32 = ctypes.windll.user32
    user32.keybd_event(vk, 0, 0, 0)
    time.sleep(0.03)
    user32.keybd_event(vk, 0, 0x0002, 0)   # KEYEVENTF_KEYUP
    time.sleep(delay)


def _open_combo():
    """Open a focused JComboBox dropdown with Alt+Down."""
    user32 = ctypes.windll.user32
    VK_MENU, VK_DOWN = 0x12, 0x28
    user32.keybd_event(VK_MENU, 0, 0, 0)
    user32.keybd_event(VK_DOWN, 0, 0, 0)
    user32.keybd_event(VK_DOWN, 0, 0x0002, 0)
    user32.keybd_event(VK_MENU, 0, 0x0002, 0)
    time.sleep(0.35)


def _type_unicode(text: str):
    """Send text as Unicode keyboard input via SendInput."""
    KEYEVENTF_UNICODE = 0x0004
    KEYEVENTF_KEYUP   = 0x0002

    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ('wVk',         wintypes.WORD),
            ('wScan',       wintypes.WORD),
            ('dwFlags',     wintypes.DWORD),
            ('time',        wintypes.DWORD),
            ('dwExtraInfo', ctypes.POINTER(wintypes.ULONG)),
        ]

    class _InputUnion(ctypes.Union):
        _fields_ = [('ki', KEYBDINPUT)]

    class INPUT(ctypes.Structure):
        _anonymous_ = ('_input',)
        _fields_    = [('type', wintypes.DWORD), ('_input', _InputUnion)]

    user32 = ctypes.windll.user32
    events = []
    for ch in text:
        for flags in (KEYEVENTF_UNICODE, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP):
            inp = INPUT(type=1)
            inp.ki.wScan   = ord(ch)
            inp.ki.dwFlags = flags
            events.append(inp)
    if events:
        arr = (INPUT * len(events))(*events)
        user32.SendInput(len(events), arr, ctypes.sizeof(INPUT))
    time.sleep(0.05)


# ─────────────────────────────────────────────────────────────
# MVR Writer
# ─────────────────────────────────────────────────────────────

class MVRWriter:
    """Fills vehicle-data fields in the open Pinnacle MVR via Java Access Bridge."""

    def __init__(self):
        self.jab  = ctypes.windll.LoadLibrary('WindowsAccessBridge-64.dll')
        self._hwnd = None
        self._setup_jab()
        self.jab.Windows_run()
        self._pump_messages(3)

    def _setup_jab(self):
        jab = self.jab
        jab.getAccessibleContextFromHWND.argtypes = [
            wintypes.HWND,
            ctypes.POINTER(ctypes.c_long),
            ctypes.POINTER(JOBJECT64),
        ]
        jab.getAccessibleContextFromHWND.restype = wintypes.BOOL

        jab.getAccessibleContextInfo.argtypes = [
            ctypes.c_long, JOBJECT64, ctypes.POINTER(AccessibleContextInfo)]
        jab.getAccessibleContextInfo.restype = wintypes.BOOL

        jab.getAccessibleChildFromContext.argtypes = [
            ctypes.c_long, JOBJECT64, ctypes.c_int]
        jab.getAccessibleChildFromContext.restype = JOBJECT64


    def _pump_messages(self, secs: float):
        user32 = ctypes.windll.user32
        msg    = wintypes.MSG()
        end    = time.time() + secs
        while time.time() < end:
            if user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1):
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
            else:
                time.sleep(0.01)

    def _find_mvr_hwnd(self):
        desktop = Desktop(backend='uia')
        for w in desktop.windows():
            try:
                if ('SunAwtFrame' in w.element_info.class_name
                        and w.window_text().startswith('Vehicle:')):
                    return w.handle
            except Exception:
                continue
        raise RuntimeError('No MVR window ("Vehicle: …") is open in Pinnacle.')

    def _connect(self, hwnd):
        vm = ctypes.c_long()
        ac = JOBJECT64()
        if not self.jab.getAccessibleContextFromHWND(
                hwnd, ctypes.byref(vm), ctypes.byref(ac)):
            raise RuntimeError('JAB could not connect to the MVR window.')
        return vm.value, ac.value

    # ── tree traversal ──────────────────────────────────────

    def _find_node(self, vm, ctx, role=None, name=None, depth=0, max_depth=30):
        ci = AccessibleContextInfo()
        if not self.jab.getAccessibleContextInfo(vm, ctx, ctypes.byref(ci)):
            return None
        if (role is None or ci.role == role) and (name is None or ci.name == name):
            return ctx
        if ci.childrenCount > 0 and depth < max_depth:
            for i in range(min(ci.childrenCount, 200)):
                child = self.jab.getAccessibleChildFromContext(vm, ctx, i)
                if child:
                    r = self._find_node(vm, child, role=role, name=name,
                                        depth=depth + 1, max_depth=max_depth)
                    if r:
                        return r
        return None

    def _collect_nodes(self, vm, ctx, roles, out, depth=0, max_depth=30):
        """Collect all visible nodes whose role is in `roles` into `out`.

        For combo boxes, options are captured IMMEDIATELY (before the JAB
        context handle can become stale) and stored in node['options'].
        """
        ci = AccessibleContextInfo()
        if not self.jab.getAccessibleContextInfo(vm, ctx, ctypes.byref(ci)):
            return
        if ci.role in roles and ci.x > 0:
            node = {
                'role': ci.role,
                'name': ci.name.strip(),
                'x': ci.x, 'y': ci.y, 'w': ci.width, 'h': ci.height,
                'ctx': ctx,
                'options': [],
            }
            # Capture combo options right now while the context is fresh.
            # Pinnacle wraps JComboBox in a panel — the direct child of the
            # 'combo box' node is the actual JComboBox with N option children.
            # We may need to descend one level to find them.
            if ci.role == 'combo box' and ci.childrenCount > 0:
                opts = self._read_combo_options(vm, ctx, ci.childrenCount)
                node['options'] = opts
            out.append(node)
        if ci.childrenCount > 0 and depth < max_depth:
            for i in range(min(ci.childrenCount, 200)):
                child = self.jab.getAccessibleChildFromContext(vm, ctx, i)
                if child:
                    self._collect_nodes(vm, child, roles, out,
                                        depth + 1, max_depth)

    # ── combo helpers ────────────────────────────────────────

    def _read_combo_options(self, vm, ctx, child_count: int) -> list:
        """Return [(index, name), …] for a combo box, looking one level deep
        if direct children have no names (Pinnacle wraps JComboBox in a panel)."""
        # First: try direct children
        direct = []
        for i in range(min(child_count, 300)):
            child = self.jab.getAccessibleChildFromContext(vm, ctx, i)
            if child:
                cci = AccessibleContextInfo()
                if self.jab.getAccessibleContextInfo(vm, child, ctypes.byref(cci)):
                    name = cci.name.strip()
                    if name:
                        direct.append((i, name))
        if direct:
            return direct

        # Second: one level deeper — find the first child that itself has
        # multiple children with names (the actual JComboBox item list)
        for i in range(min(child_count, 5)):
            child = self.jab.getAccessibleChildFromContext(vm, ctx, i)
            if not child:
                continue
            inner_ci = AccessibleContextInfo()
            if not self.jab.getAccessibleContextInfo(vm, child, ctypes.byref(inner_ci)):
                continue
            if inner_ci.childrenCount < 2:
                continue
            opts = []
            for j in range(min(inner_ci.childrenCount, 300)):
                grandchild = self.jab.getAccessibleChildFromContext(vm, child, j)
                if grandchild:
                    gcci = AccessibleContextInfo()
                    if self.jab.getAccessibleContextInfo(
                            vm, grandchild, ctypes.byref(gcci)):
                        name = gcci.name.strip()
                        if name:
                            opts.append((j, name))
            if opts:
                return opts
        return []

    def _set_combo_keyboard(self, field: dict, value: str) -> bool:
        """Select a combo option by opening the dropdown and typing to navigate.

        Swing's incremental-search (prefix matching) in an open JList lets us
        type the first characters of the target option to jump directly to it.
        No need to enumerate options; works even with lazy-loaded item lists.
        """
        cx = field['x'] + field['w'] // 2
        cy = field['y'] + field['h'] // 2

        # 1. Click to focus the combo
        ctypes.windll.user32.SetForegroundWindow(self._hwnd)
        _click(cx, cy, delay=0.25)

        # 2. Open the dropdown with Alt+Down
        _open_combo()

        # 3. Type the first 5 characters of the target value.
        #    In an open Swing JList, successive key presses within ~1 second
        #    do prefix matching (e.g. 'G','r' → selects "Gray" over "Gold").
        #    Use uppercase VK codes (A-Z: 0x41-0x5A, 0-9: 0x30-0x39).
        for ch in value[:5].upper():
            code = ord(ch)
            if (0x41 <= code <= 0x5A) or (0x30 <= code <= 0x39):
                _send_vk(code, delay=0.08)

        # 4. Confirm selection with Enter
        time.sleep(0.1)
        _send_vk(0x0D)   # VK_RETURN
        time.sleep(0.1)

        return True

    # ── button click ─────────────────────────────────────────

    def _click_button(self, vm, ac, name: str):
        btn = self._find_node(vm, ac, role='push button', name=name)
        if not btn:
            raise RuntimeError(f'Button "{name}" not found in MVR.')
        ci = AccessibleContextInfo()
        if not self.jab.getAccessibleContextInfo(vm, btn, ctypes.byref(ci)):
            raise RuntimeError(f'Cannot get info for button "{name}".')
        ctypes.windll.user32.SetForegroundWindow(self._hwnd)
        time.sleep(0.2)
        _click(ci.x + ci.width // 2, ci.y + ci.height // 2, delay=0.4)

    # ── field finder ─────────────────────────────────────────

    @staticmethod
    def _nearest_field(label_node, fields):
        """
        Given a label node dict, find the field (text or combo) that is:
          - to the RIGHT of the label
          - at approximately the same y-centre (within 15 px)
          - closest horizontal distance after the label's right edge
        """
        lx, ly, lw, lh = label_node['x'], label_node['y'], label_node['w'], label_node['h']
        label_right = lx + lw
        label_cy    = ly + lh / 2

        best, best_score = None, 9999
        for f in fields:
            if f['x'] <= lx:          # must be to the right of label
                continue
            fy_cy = f['y'] + f['h'] / 2
            y_dist = abs(fy_cy - label_cy)
            if y_dist > 15:
                continue
            x_gap = f['x'] - label_right
            if x_gap < 0:
                continue
            score = y_dist * 10 + x_gap
            if score < best_score:
                best_score = score
                best = f
        return best

    # ── main fill routine ─────────────────────────────────────

    def fill(self, vehicle_data: dict, dry_run: bool = False) -> dict:
        """
        Fill Pinnacle MVR fields from a VINMatchPro decode dict.
        Returns a dict of {pinnacle_label: value_written}.
        """
        self._hwnd = self._find_mvr_hwnd()
        vm, ac = self._connect(self._hwnd)

        # ── Enter edit mode ───────────────────────────────────────
        ctypes.windll.user32.SetForegroundWindow(self._hwnd)
        time.sleep(0.3)
        print('  Clicking Edit...')
        self._click_button(vm, ac, 'Edit')
        time.sleep(3.0)

        # Re-connect after edit mode activates
        vm, ac = self._connect(self._hwnd)

        # ── Step C: collect field positions in edit mode ──────────
        labels = []
        fields = []
        self._collect_nodes(vm, ac, {'label'},              labels)
        self._collect_nodes(vm, ac, {'text', 'combo box'},  fields)
        fields = [f for f in fields if f['x'] > 0]

        print(f'  Found {len(labels)} labels, {len(fields)} fields in edit mode.')

        results = {}

        for label_name, (vmp_key, transformer) in _FIELD_MAP.items():
            raw = vehicle_data.get(vmp_key, '')
            if not raw:
                print(f'  {label_name:<28} (no data from VINMatchPro)')
                results[label_name] = None
                continue

            value = transformer(raw) if transformer else str(raw).strip()
            if not value:
                print(f'  {label_name:<28} cannot map "{raw}"')
                results[label_name] = None
                continue

            # Find ALL labels with this name (e.g. "Drive:" appears twice)
            matching_labels = [l for l in labels if l['name'] == label_name]
            if not matching_labels:
                print(f'  {label_name:<28} label not found in UI')
                results[label_name] = None
                continue

            filled = False
            for lbl in matching_labels:
                field = self._nearest_field(lbl, fields)
                if not field:
                    continue

                cx = field['x'] + field['w'] // 2
                cy = field['y'] + field['h'] // 2
                role = field['role']

                suffix = f'[{role} @ ({field["x"]},{field["y"]})]'
                msg    = f'  {label_name:<28} "{raw}" -> "{value}"  {suffix}'

                if dry_run:
                    print(msg + '  [DRY RUN]')
                    results[label_name] = value
                    filled = True
                    continue

                if role == 'combo box':
                    self._set_combo_keyboard(field, value)
                    print(msg + '  OK')
                    results[label_name] = value
                    filled = True
                else:  # text field
                    ctypes.windll.user32.SetForegroundWindow(self._hwnd)
                    _click(cx, cy)
                    _ctrl_a()
                    _type_unicode(value)
                    print(msg + '  OK')
                    results[label_name] = value
                    filled = True

                time.sleep(0.1)

            if not filled and label_name not in results:
                results[label_name] = None

        if not dry_run:
            print('\n  Saving...')
            vm, ac = self._connect(self._hwnd)
            self._click_button(vm, ac, 'Save')
            time.sleep(1.0)
            print('  Saved.')

        return results


# ─────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='Decode VIN on VINMatchPro and fill Pinnacle MVR fields.')
    parser.add_argument('--vin',      default=None,
                        help='VIN to decode (omit to read from Pinnacle)')
    parser.add_argument('--dry-run',  action='store_true',
                        help='Show what would be filled without writing anything')
    args = parser.parse_args()

    # ── Step 1: Get VIN ──────────────────────────────────────
    if args.vin:
        vin = args.vin.upper().strip()
        print(f'Using VIN from command line: {vin}')
    else:
        print('Reading VIN from Pinnacle Find Vehicles (selected row)…')
        try:
            vin = read_vin_from_pinnacle()
        except Exception as e:
            print(f'  ERROR reading VIN from Pinnacle: {e}', file=sys.stderr)
            print('  Tip: make sure "Find Vehicles" is open and a row is selected,')
            print('       or pass --vin <VIN> to specify the VIN directly.')
            sys.exit(1)
        print(f'  VIN: {vin}')

    # ── Step 2: Decode on VINMatchPro ────────────────────────
    print(f'\nDecoding {vin} on VINMatchPro...')
    try:
        vehicle = vmp_decode(vin)
    except Exception as e:
        print(f'  ERROR decoding VIN: {e}', file=sys.stderr)
        sys.exit(1)

    print(f'  {vehicle.get("year")} {vehicle.get("make")} {vehicle.get("model")}')
    print(f'  Trim:         {vehicle.get("trim","--")}')
    print(f'  Engine:       {vehicle.get("engine","--")}')
    print(f'  Transmission: {vehicle.get("transmission","--")}')
    print(f'  Drive:        {vehicle.get("drive_type","--")}')
    print(f'  Fuel:         {vehicle.get("fuel_type","--")}')
    print(f'  Ext color:    {vehicle.get("exterior_color","--")}')
    print(f'  Int color:    {vehicle.get("interior_color","--")}')
    print(f'  Cylinders:    {vehicle.get("engine_cylinders","--")}')
    print(f'  Axle ratio:   {vehicle.get("axle_ratio","--")}')
    print(f'  Weight:       {vehicle.get("curb_weight","--")}')

    if args.dry_run:
        print('\n[DRY RUN] Fields that would be written:')
    else:
        print('\nFilling Pinnacle MVR...')

    # ── Step 3: Write to Pinnacle ────────────────────────────
    try:
        writer  = MVRWriter()
        results = writer.fill(vehicle, dry_run=args.dry_run)
    except Exception as e:
        print(f'\n  ERROR filling MVR: {e}', file=sys.stderr)
        sys.exit(1)

    # Summary
    filled  = [k for k, v in results.items() if v is not None]
    skipped = [k for k, v in results.items() if v is None]
    print(f'\nDone. {len(filled)} fields filled, {len(skipped)} skipped.')
    if skipped:
        print(f'Skipped: {", ".join(skipped)}')


if __name__ == '__main__':
    main()
