import tkinter as tk
from tkinter import ttk, messagebox
from pylogix import PLC
import traceback
from typing import Any, Callable, List, Optional, Tuple

# Add aphyt for Omron CIP STRING write support
try:
    import aphyt  # noqa: F401
except Exception:
    aphyt = None

PLC_IP = "192.168.100.87"
TAG_PREFIX = "Wacon_INT"
SERIE_NAME_TAG = "Wacon_SerieName"
OMRON_STRING_LENGTH = 256  # For STRING[256]

def import_aphyt_controllers() -> List[Tuple[str, Any]]:
    """
    Try to import a few known aphyt controller classes in order.
    Returns a list of (name, class) tuples that successfully imported.
    """
    controllers: List[Tuple[str, Any]] = []
    if aphyt is None:
        return controllers  # aphyt not installed

    import_errors: List[Tuple[str, Exception]] = []

    # Try NSeries in modern and legacy paths
    try:
        from aphyt.omron.n_series import NSeries  # type: ignore
        controllers.append(("aphyt.omron.n_series.NSeries", NSeries))
    except Exception as e:
        import_errors.append(("aphyt.omron.n_series.NSeries", e))

    try:
        from aphyt.omron import NSeries as NSeriesLegacy  # type: ignore
        controllers.append(("aphyt.omron.NSeries", NSeriesLegacy))
    except Exception as e:
        import_errors.append(("aphyt.omron.NSeries", e))

    # Some older examples referenced "Controller" or "NJController"
    try:
        from aphyt.omron import Controller  # type: ignore
        controllers.append(("aphyt.omron.Controller", Controller))
    except Exception as e:
        import_errors.append(("aphyt.omron.Controller", e))

    try:
        from aphyt.omron import NJController  # type: ignore
        controllers.append(("aphyt.omron.NJController", NJController))
    except Exception as e:
        import_errors.append(("aphyt.omron.NJController", e))

    if not controllers:
        print("Failed to import any aphyt controller classes. Try: pip install --upgrade aphyt")
        print("\nImport errors:")
        for name, err in import_errors:
            print(f" - {name}: {err}")

    return controllers


def try_invoke(obj: Any, method_names: List[str]) -> Optional[Tuple[str, Callable]]:
    for name in method_names:
        fn = getattr(obj, name, None)
        if callable(fn):
            return name, fn
    return None


def maybe_connect(ctrl: Any) -> None:
    candidates = [
        "open",
        "connect",
        "create_session",
        "register_session",
        "start",
        "initialize",
        "connect_explicit",
    ]
    found = try_invoke(ctrl, candidates)
    if found:
        name, fn = found
        print(f"[CIP] Using '{name}()' to establish session...")
        try:
            fn()
            print("[CIP] Session established.")
        except Exception as e:
            print(f"[CIP] '{name}()' raised: {e}. Continuing anyway.")


def maybe_disconnect(ctrl: Any) -> None:
    for name in ["close", "disconnect", "stop"]:
        fn = getattr(ctrl, name, None)
        if callable(fn):
            try:
                fn()
                print("[CIP] Disconnected.")
                return
            except Exception:
                pass


def find_write_method(ctrl: Any) -> Tuple[str, Callable]:
    candidates = ["write_variable", "write", "write_var", "set_variable"]
    found = try_invoke(ctrl, candidates)
    if not found:
        raise RuntimeError("No write method found on controller (looked for: " + ", ".join(candidates) + ")")
    return found


def find_read_method(ctrl: Any) -> Optional[Tuple[str, Callable]]:
    candidates = ["read_variable", "read", "get_variable"]
    return try_invoke(ctrl, candidates)


# Dedicated writer for Omron STRING using aphyt
def write_omron_string_via_aphyt(ip: str, tag: str, value: str, max_bytes: int = 256) -> bool:
    """
    Write an Omron STRING tag through EtherNet/IP using aphyt.
    Attempts a direct string write; if that fails, writes .Len and .Data[i].
    Returns True on success, False otherwise.
    """
    if aphyt is None:
        messagebox.showerror("Missing Dependency", "aphyt is not installed. Install with: pip install aphyt")
        return False

    print(f"[CIP] Preparing to write STRING tag '{tag}' with value: {value!r}")

    # Enforce length
    encoded = value.encode("utf-8")
    if len(encoded) > max_bytes:
        messagebox.showwarning(
            "String Truncated",
            f"Input exceeds {max_bytes} bytes. It will be truncated."
        )
        encoded = encoded[:max_bytes]
        # Try not to cut a multi-byte sequence
        value = encoded.decode("utf-8", errors="ignore")

    controllers = import_aphyt_controllers()
    if not controllers:
        return False

    last_error = None

    for class_name, cls in controllers:
        ctrl = None
        print(f"[CIP] Trying controller class: {class_name}")
        try:
            ctrl = cls(ip)
            maybe_connect(ctrl)

            try:
                rm = find_read_method(ctrl)
                if rm:
                    rname, rfn = rm
                    curr = rfn(tag)
                    print(f"[CIP] Current value via {rname}(): {curr!r}")
            except Exception as e_read:
                print(f"[CIP] Pre-write read failed (continuing): {e_read}")

            wname, wfn = find_write_method(ctrl)

            # Direct attempt
            try:
                wfn(tag, value)
                print(f"[CIP] Direct write via {wname}() succeeded.")
                maybe_disconnect(ctrl)
                return True
            except Exception as e_direct:
                print(f"[CIP] Direct write failed, falling back. Error: {e_direct}")

            # Structured write
            data = value.encode("utf-8")
            length = len(data)
            print(f"[CIP] Structured write: .Len={length}, bytes={data}")

            wfn(f"{tag}.Len", int(length))
            for i, byte in enumerate(data):
                wfn(f"{tag}.Data[{i}]", int(byte))
            # Optional null terminator
            try:
                wfn(f"{tag}.Data[{length}]", 0)
            except Exception:
                pass

            print("[CIP] Structured write succeeded.")

            try:
                rm = find_read_method(ctrl)
                if rm:
                    rname, rfn = rm
                    rb = rfn(tag)
                    print(f"[CIP] Read-back value via {rname}(): {rb!r}")
            except Exception as e_rb:
                print(f"[CIP] Read-back failed (continuing): {e_rb}")

            maybe_disconnect(ctrl)
            return True

        except Exception as e:
            last_error = e
            print(f"[CIP] Controller '{class_name}' attempt failed: {e}")
            traceback.print_exc()
            try:
                if ctrl is not None:
                    maybe_disconnect(ctrl)
            except Exception:
                pass
            continue

    print("[CIP] All controller attempts failed.")
    if last_error:
        print("Last error:", last_error)
    return False


# Optional read via aphyt (fallback to pylogix)
def read_omron_string_via_aphyt(ip: str, tag: str) -> Optional[str]:
    if aphyt is None:
        return None

    controllers = import_aphyt_controllers()
    if not controllers:
        return None

    for class_name, cls in controllers:
        print(f"[CIP] Trying read controller: {class_name}")
        try:
            ctrl = cls(ip)
            maybe_connect(ctrl)
            rm = find_read_method(ctrl)
            if not rm:
                continue
            rname, rfn = rm
            val = rfn(tag)
            maybe_disconnect(ctrl)
            return val
        except Exception as e:
            print(f"[CIP] Read attempt with {class_name} failed: {e}")
            traceback.print_exc()
            continue
    return None


def read_plc_value(tag, is_string=False):
    try:
        if not is_string:
            with PLC() as comm:
                comm.IPAddress = PLC_IP
                result = comm.Read(tag)
                if result.Status == 'Success':
                    val = result.Value
                    return val
                else:
                    return None
        else:
            # Use aphyt for STRING instead of pylogix
            val = read_omron_string_via_aphyt(PLC_IP, tag)
            return val
    except Exception as e:
        print(f"PLC read error: {e}")
        return None


def write_plc_value(tag, value, is_string=False, tag_length=OMRON_STRING_LENGTH):
    try:
        if is_string:
            # Use the aphyt-based writer
            return write_omron_string_via_aphyt(PLC_IP, tag, value, max_bytes=tag_length)
        else:
            with PLC() as comm:
                comm.IPAddress = PLC_IP
                wr = comm.Write(tag, value)
            print("Write result:", wr)
            if hasattr(wr, 'Status'):
                print("Status:", wr.Status)
            if hasattr(wr, 'Error'):
                print("Error:", wr.Error)
            return hasattr(wr, 'Status') and wr.Status == 'Success'
    except Exception as e:
        print(f"PLC write error: {e}")
        return False


class IntEntry(ttk.Entry):
    """Entry widget that only accepts integers."""
    def __init__(self, master=None, **kwargs):
        self.var = tk.StringVar()
        super().__init__(master, textvariable=self.var, **kwargs)
        self.var.trace_add('write', self.validate)
        self.old_value = ''
        self.config(validate='key')
        self.bind('<FocusOut>', self._cleanup)

    def validate(self, *args):
        value = self.var.get()
        if value == '':
            self.old_value = value
        elif value == '-':
            self.old_value = value
        elif value.isdigit() or (value.startswith('-') and value[1:].isdigit()):
            self.old_value = value
        else:
            self.var.set(self.old_value)

    def _cleanup(self, event):
        txt = self.var.get()
        if txt == '' or txt == '-':
            self.var.set('0')


def main():
    root = tk.Tk()
    root.title("PLC Wacon_INT & Wacon_SerieName Panel")

    entries = []
    result_labels = []

    def make_row(parent, idx):
        frame = ttk.Frame(parent)
        frame.grid(row=idx, column=0, sticky='ew', pady=2)
        tag = f"{TAG_PREFIX}[{idx}]"
        label = ttk.Label(frame, text=tag, width=15)
        label.pack(side='left', padx=2)
        entry = IntEntry(frame, width=10)
        entry.pack(side='left', padx=2)
        entries.append(entry)
        write_btn = ttk.Button(frame, text="Write", width=6)
        write_btn.pack(side='left', padx=2)
        read_btn = ttk.Button(frame, text="Read", width=6)
        read_btn.pack(side='left', padx=2)
        res_label = ttk.Label(frame, text="Read: ---", width=12)
        res_label.pack(side='left', padx=2)
        result_labels.append(res_label)

        def do_write():
            try:
                val = int(entry.get())
            except ValueError:
                messagebox.showerror("Invalid Input", "Please enter a valid integer.")
                return
            success = write_plc_value(tag, val, is_string=False)
            if success:
                res_label.config(text="Write OK", foreground="green")
            else:
                res_label.config(text="Write Error", foreground="red")

        def do_read():
            read_val = read_plc_value(tag, is_string=False)
            if read_val is not None:
                res_label.config(text=f"Read: {read_val}", foreground="blue")
            else:
                res_label.config(text="Read: Error", foreground="red")

        write_btn.config(command=do_write)
        read_btn.config(command=do_read)

    for i in range(24):
        make_row(root, i)

    def make_string_row(parent, idx):
        frame = ttk.Frame(parent)
        frame.grid(row=idx, column=0, sticky='ew', pady=8)
        label = ttk.Label(frame, text=SERIE_NAME_TAG, width=15)
        label.pack(side='left', padx=2)
        entry = ttk.Entry(frame, width=30)
        entry.pack(side='left', padx=2)
        write_btn = ttk.Button(frame, text="Write", width=6)
        write_btn.pack(side='left', padx=4)
        read_btn = ttk.Button(frame, text="Read", width=6)
        read_btn.pack(side='left', padx=4)
        res_label = ttk.Label(frame, text="Read: ---", width=30)
        res_label.pack(side='left', padx=2)

        def do_write():
            val = entry.get()
            if val == '':
                messagebox.showerror("Invalid Input", "Please enter a string.")
                return
            success = write_plc_value(SERIE_NAME_TAG, val, is_string=True, tag_length=OMRON_STRING_LENGTH)
            if success:
                res_label.config(text="Write OK", foreground="green")
            else:
                res_label.config(text="Write Error", foreground="red")

        def do_read():
            read_val = read_plc_value(SERIE_NAME_TAG, is_string=True)
            if read_val is not None:
                res_label.config(text=f"Read: {read_val}", foreground="blue")
            else:
                res_label.config(text="Read: Error", foreground="red")

        write_btn.config(command=do_write)
        read_btn.config(command=do_read)

    make_string_row(root, 24)

    root.mainloop()

if __name__ == "__main__":
    main()