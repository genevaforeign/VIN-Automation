"""
pinnacle_field_dumper.py - Dumps all accessible/editable fields from an open
Pinnacle window via Java Access Bridge.

Run this while the Pinnacle MVR is open on whatever tab/screen contains
the factory options fields you want to write to. The output shows every
node role, name, value, and screen coordinates so we can map field targets.

Usage:
    python pinnacle_field_dumper.py
    python pinnacle_field_dumper.py --window "Find Vehicles"
    python pinnacle_field_dumper.py --roles text combobox checkbox
    python pinnacle_field_dumper.py --dump-all        # dump every node (verbose)
"""

import argparse
import ctypes
import ctypes.wintypes as wintypes
import sys
import time

from pywinauto import Desktop

from pinnacle_reader import (
    AccessibleContextInfo,
    AccessibleTableInfo,
    AccessibleTableCellInfo,
    AccessibleTextInfo,
    JOBJECT64,
)

# Roles that typically represent editable / interesting UI elements
EDITABLE_ROLES = {
    'text', 'combo box', 'combobox', 'check box', 'checkbox',
    'radio button', 'radiobutton', 'list', 'list item',
    'push button', 'button', 'label', 'panel', 'page tab',
    'page tab list', 'toggle button', 'spin box', 'spinbox',
}


class FieldDumper:
    def __init__(self):
        self.jab = ctypes.windll.LoadLibrary('WindowsAccessBridge-64.dll')
        self._setup_prototypes()
        self.jab.Windows_run()
        self._pump_messages(3)
        self._nodes = []

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
        jab.getAccessibleTextInfo.argtypes = [
            ctypes.c_long, JOBJECT64, ctypes.POINTER(AccessibleTextInfo),
            ctypes.c_int, ctypes.c_int]
        jab.getAccessibleTextInfo.restype = wintypes.BOOL
        jab.getAccessibleTextRange.argtypes = [
            ctypes.c_long, JOBJECT64, ctypes.c_int, ctypes.c_int,
            ctypes.c_wchar_p, ctypes.c_short]
        jab.getAccessibleTextRange.restype = wintypes.BOOL

    def _pump_messages(self, seconds):
        user32 = ctypes.windll.user32
        msg = wintypes.MSG()
        end = time.time() + seconds
        while time.time() < end:
            if user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1):
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
            else:
                time.sleep(0.01)

    def _get_text(self, vm, ctx):
        ti = AccessibleTextInfo()
        if self.jab.getAccessibleTextInfo(vm, ctx, ctypes.byref(ti), 0, 0):
            if ti.charCount > 0:
                buf = ctypes.create_unicode_buffer(min(ti.charCount + 1, 512))
                if self.jab.getAccessibleTextRange(
                    vm, ctx, 0, min(ti.charCount - 1, 510), buf, len(buf)
                ):
                    return buf.value
        return ''

    def _find_window(self, title_contains: str):
        """Find any open Pinnacle/Java Swing window whose title contains the given string."""
        desktop = Desktop(backend='uia')
        matches = []
        for w in desktop.windows():
            try:
                title = w.window_text()
                cls = w.element_info.class_name
                if 'SunAwtFrame' in cls and title_contains.lower() in title.lower():
                    matches.append((w.handle, title))
            except Exception:
                continue
        if not matches:
            # List all SunAwtFrame windows to help debugging
            print('\nNo matching window found. Open SunAwtFrame windows:')
            for w in Desktop(backend='uia').windows():
                try:
                    if 'SunAwtFrame' in w.element_info.class_name:
                        print(f'  [{w.handle}] "{w.window_text()}"')
                except Exception:
                    pass
            raise RuntimeError(f'No Pinnacle window with title containing "{title_contains}" found.')
        if len(matches) > 1:
            print(f'Multiple matching windows found, using first:')
            for h, t in matches:
                print(f'  [{h}] "{t}"')
        return matches[0][0]

    def _walk_tree(self, vm, ctx, depth=0, max_depth=30, roles_filter=None, dump_all=False):
        """Recursively walk the accessible tree and collect matching nodes."""
        ci = AccessibleContextInfo()
        if not self.jab.getAccessibleContextInfo(vm, ctx, ctypes.byref(ci)):
            return

        role = ci.role.strip().lower() if ci.role else ''
        name = ci.name.strip() if ci.name else ''
        desc = ci.description.strip() if ci.description else ''
        states = ci.states.strip() if ci.states else ''

        # Get text value for text nodes
        text_val = ''
        if ci.accessibleText:
            text_val = self._get_text(vm, ctx)

        include = dump_all
        if not include:
            if roles_filter:
                include = role in roles_filter
            else:
                include = role in EDITABLE_ROLES

        if include:
            self._nodes.append({
                'depth': depth,
                'role': role,
                'name': name,
                'description': desc,
                'states': states,
                'text': text_val,
                'x': ci.x,
                'y': ci.y,
                'w': ci.width,
                'h': ci.height,
            })

        if ci.childrenCount > 0 and depth < max_depth:
            for i in range(min(ci.childrenCount, 200)):
                child = self.jab.getAccessibleChildFromContext(vm, ctx, i)
                if child:
                    self._walk_tree(vm, child, depth + 1, max_depth,
                                    roles_filter=roles_filter, dump_all=dump_all)

    def dump(self, title_contains: str = 'Vehicle:', roles_filter=None, dump_all=False):
        hwnd = self._find_window(title_contains)
        vm_id = ctypes.c_long()
        ac = JOBJECT64()
        if not self.jab.getAccessibleContextFromHWND(
            hwnd, ctypes.byref(vm_id), ctypes.byref(ac)
        ):
            raise RuntimeError('JAB could not connect to the window.')

        vm = vm_id.value
        self._nodes = []
        self._walk_tree(vm, ac.value, roles_filter=roles_filter, dump_all=dump_all)
        return self._nodes


def _fmt_node(node):
    indent = '  ' * node['depth']
    parts = [f"{indent}[{node['role']}]"]
    if node['name']:
        parts.append(f"name={node['name']!r}")
    if node['text']:
        parts.append(f"text={node['text']!r}")
    if node['description']:
        parts.append(f"desc={node['description']!r}")
    if node['states']:
        parts.append(f"states={node['states']!r}")
    if node['x'] > 0:
        parts.append(f"pos=({node['x']},{node['y']}) size={node['w']}x{node['h']}")
    return ' '.join(parts)


def main():
    parser = argparse.ArgumentParser(description='Dump accessible fields from Pinnacle window.')
    parser.add_argument('--window', default='Vehicle:',
                        help='Title substring to match (default: "Vehicle:")')
    parser.add_argument('--roles', nargs='+', default=None,
                        help='Only show these roles (e.g. --roles text "combo box" label)')
    parser.add_argument('--dump-all', action='store_true',
                        help='Dump every node in the tree (very verbose)')
    parser.add_argument('--filter', default=None,
                        help='Only print nodes whose name/text contains this string (case-insensitive)')
    args = parser.parse_args()

    roles_set = None
    if args.roles:
        roles_set = {r.lower() for r in args.roles}

    print(f'Connecting to Pinnacle window matching "{args.window}"...')
    dumper = FieldDumper()
    nodes = dumper.dump(
        title_contains=args.window,
        roles_filter=roles_set,
        dump_all=args.dump_all,
    )

    if args.filter:
        f = args.filter.lower()
        nodes = [n for n in nodes if f in n['name'].lower() or f in n['text'].lower()
                 or f in n['description'].lower()]

    print(f'\nFound {len(nodes)} nodes:\n')
    for node in nodes:
        print(_fmt_node(node))

    # Summary: unique roles seen
    from collections import Counter
    role_counts = Counter(n['role'] for n in nodes)
    print('\n--- Role summary ---')
    for role, count in role_counts.most_common():
        print(f'  {role}: {count}')

    # Editable text fields specifically
    text_nodes = [n for n in nodes if n['role'] in ('text',) and n['x'] > 0]
    if text_nodes:
        print(f'\n--- Text fields with screen coordinates ({len(text_nodes)}) ---')
        for n in text_nodes:
            print(f"  name={n['name']!r}  value={n['text']!r}  pos=({n['x']},{n['y']})")

    combo_nodes = [n for n in nodes if 'combo' in n['role'] and n['x'] > 0]
    if combo_nodes:
        print(f'\n--- Combo boxes with screen coordinates ({len(combo_nodes)}) ---')
        for n in combo_nodes:
            print(f"  name={n['name']!r}  value={n['text']!r}  pos=({n['x']},{n['y']})")


if __name__ == '__main__':
    try:
        main()
    except RuntimeError as e:
        print(f'\nERROR: {e}', file=sys.stderr)
        sys.exit(1)
