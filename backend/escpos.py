"""
ESC/POS Command Generator & Thermal Printer Engine for 80mm (3-inch) Receipt Printers.
Designed specifically for ATPOS AT-301 and standard ESC/POS 80mm thermal printers:
  - Strict 80mm roll width / 72mm printable width (576 dots at 203 DPI)
  - Font A: 12x24 dots = exactly 48 characters per line
  - Compact typography & clean vertical spacing to prevent clipping
  - Hardware Auto-Cutter integration after EVERY individual token slip
  - Sequential print-and-cut execution with buffer clearing to prevent paper jams & uncut strips
  - Thread-safe printer lock to prevent overlapping print jobs
"""
import base64
import os
import re
import socket
import threading
import time

try:
    import win32print
except ImportError:
    win32print = None

# Global lock to serialize physical thermal printer jobs and prevent overlapping streams
_printer_lock = threading.Lock()

# ESC/POS Command Byte Constants
ESC = b"\x1b"
GS = b"\x1d"
FS = b"\x1c"

# Control commands
INIT_PRINTER = ESC + b"@"
FONT_A = ESC + b"M\x00"        # Font A: 12x24 dots (48 cols on 80mm / 576 dots)
FONT_B = ESC + b"M\x01"        # Font B: 9x17 dots (64 cols on 80mm)
ALIGN_LEFT = ESC + b"a\x00"
ALIGN_CENTER = ESC + b"a\x01"
ALIGN_RIGHT = ESC + b"a\x02"
BOLD_ON = ESC + b"E\x01"
BOLD_OFF = ESC + b"E\x00"

# Text sizing
SIZE_NORMAL = GS + b"!\x00"
SIZE_DOUBLE_HEIGHT = GS + b"!\x01"
SIZE_DOUBLE_WIDTH = GS + b"!\x10"
SIZE_DOUBLE = GS + b"!\x11"    # Double width & double height
SIZE_LARGE = GS + b"!\x22"     # 3x width & 3x height for serial number

# Feed lines
FEED_LINES_4 = ESC + b"d\x04"
FEED_LINES_6 = ESC + b"d\x06"

# Standard Paper Cut Commands for ATPOS AT-301 and 80mm thermal receipt printers:
# Thermal printers physically have the auto-cutter knife 8-12mm above the thermal head.
# Feeding 4 lines ensures receipt text clears the blade completely before cutting.
# GS V 0 (\x1d\x56\x00) -> Function A: Full cut
# GS V 1 (\x1d\x56\x01) -> Function A: Partial cut
# GS V 66 0 (\x1d\x56\x42\x00) -> Function B: feeds to cutting position & triggers partial cut
CUT_FULL = FEED_LINES_4 + GS + b"V\x00"
CUT_PARTIAL = FEED_LINES_4 + GS + b"V\x01"
CUT_FEED_PARTIAL = FEED_LINES_4 + GS + b"V\x42\x00"

LINE_WIDTH = 48  # Strict 80mm Font A standard (576 dots / 12 dots = 48 chars)


def pad_center(text: str, width: int = LINE_WIDTH) -> str:
    """Centers text within strict 80mm line width (48 chars)."""
    text = str(text).strip()
    if len(text) >= width:
        return text[:width]
    return text.center(width)


def pad_sides(left: str, right: str, width: int = LINE_WIDTH) -> str:
    """Pads left and right text to fit exactly within width (48 chars)."""
    left = str(left).strip()
    right = str(right).strip()
    space = width - len(left) - len(right)
    if space < 1:
        return (f"{left} {right}")[:width]
    return left + (" " * space) + right


def pad_key_val(key: str, val: str, key_width: int = 14, width: int = LINE_WIDTH) -> str:
    """Formats a label: value pair strictly within width (48 chars)."""
    prefix = f"{str(key).strip():<{key_width}}: "
    rem = width - len(prefix)
    val_str = str(val).strip()
    if len(val_str) > rem:
        val_str = val_str[:rem]
    return prefix + val_str


def divider(char: str = "-", width: int = LINE_WIDTH) -> str:
    """Returns a divider line of exactly width characters."""
    return (char * width)[:width]


def build_cut_command(cut_mode: str = "full") -> bytes:
    """
    Standard reusable ESC/POS auto-cutter sequence for 80mm thermal receipt printers.
    Feeds 4 blank lines to clear the physical knife blade completely, then triggers cut.
    """
    feed = b"\n\n\n\n"
    if cut_mode == "partial":
        return feed + b"\x1d\x56\x01"  # GS V 1 (Partial cut)
    elif cut_mode == "feed_cut":
        return feed + b"\x1d\x56\x42\x00"  # GS V 66 0 (Feed and cut)
    else:
        return feed + b"\x1d\x56\x00"  # GS V 0 (Standard Full cut)


def generate_receipt_text(token: dict, copy_label: str = "CUSTOMER COPY", is_reprint: bool = False) -> str:
    """
    Generates a clean, strictly formatted 48-characters-per-line text slip
    for 80mm thermal printers. Identical clean structure for all categories:
    Royal Enfield, Silver, Saree, and Kunkuma Archana.
    """
    lines = []
    lines.append(divider("="))
    lines.append(pad_center("SVARA 2026 LUCKY DRAW"))
    lines.append(pad_center("Official Token Receipt"))
    lines.append(pad_center("Cell: +91 9848433020, 9885897093"))
    lines.append(divider("-"))

    if is_reprint or token.get("is_reprint"):
        lines.append(pad_center("*** REPRINT / DUPLICATE COPY ***"))
        reason = token.get("reprint_reason") or "Lost or torn receipt"
        lines.append(pad_key_val("Reason", reason))
        reprint_time = token.get("reprinted_at") or token.get("timestamp") or ""
        if reprint_time:
            lines.append(pad_key_val("Reprinted", reprint_time))
        lines.append(divider("-"))

    cat_name = token.get("type", token.get("token_type", "Token")).upper()
    price = token.get("price", 0)
    cat_str = f"{cat_name} - Rs. {price}" if price and int(price) > 0 else cat_name
    lines.append(pad_center(cat_str))
    lines.append(divider("-"))

    serial = str(token.get("serial", "")).strip()
    lines.append(pad_center(f">> {serial} <<"))
    lines.append(divider("-"))

    name = token.get("name", "")
    mobile = token.get("mobile", "")
    payment = token.get("payment", "")
    counter = token.get("counter_name", "Counter")
    date_val = token.get("date", token.get("created_date", ""))
    time_val = token.get("time", token.get("created_time", ""))

    if price and int(price) > 0:
        lines.append(pad_key_val("Price", f"Rs. {price}"))
    lines.append(pad_key_val("Devotee Name", name))
    lines.append(pad_key_val("Mobile No", mobile))
    pay_display = "PENDING (NOT PAID)" if payment in ("Payment Pending", "Pending") else payment
    lines.append(pad_key_val("Payment Mode", pay_display))
    if payment in ("Payment Pending", "Pending"):
        lines.append(pad_center("*** PAYMENT PENDING - PLEASE COLLECT ***"))
    lines.append(pad_key_val("Counter", counter))
    lines.append(pad_sides(f"Date: {date_val}", f"Time: {time_val}"))
    lines.append(divider("-"))

    lines.append(pad_center("Thank you for registration!"))
    lines.append(pad_center("May Goddess Durgamatha bless you and your family."))
    lines.append(divider("-"))
    lines.append(pad_center(f"[ {copy_label.upper()} ]"))
    lines.append(divider("="))
    return "\n".join(lines)


def generate_escpos_slip(token: dict, copy_label: str = "CUSTOMER COPY", is_reprint: bool = False, cut_mode: str = "full") -> bytes:
    """
    Generates raw ESC/POS binary stream for an 80mm thermal receipt printer slip.
    Strictly formatted to 48 columns (Font A, 576 dots printable area).
    Appends standard paper cut command at the end of each slip to trigger the auto-cutter.
    Follows uniform structure across Royal Enfield, Silver, Saree, and Kunkuma Archana.
    """
    out = bytearray()

    # 1. Initialize printer & select Font A (48 columns on 80mm roll)
    out.extend(INIT_PRINTER)
    out.extend(FONT_A)

    # 2. Header
    out.extend(ALIGN_CENTER)
    out.extend(BOLD_ON)
    out.extend(SIZE_DOUBLE_HEIGHT)
    out.extend(b"SVARA 2026 LUCKY DRAW\n")
    out.extend(SIZE_NORMAL)
    out.extend(BOLD_OFF)

    out.extend(b"Official Token Receipt\n")
    out.extend(b"Cell: +91 9848433020, 9885897093\n")
    out.extend(divider("-").encode("ascii", "replace") + b"\n")

    # 3. Reprint Badge (if reprint)
    if is_reprint or token.get("is_reprint"):
        out.extend(BOLD_ON)
        out.extend(b"*** REPRINT / DUPLICATE COPY ***\n")
        out.extend(BOLD_OFF)
        reason = token.get("reprint_reason") or "Lost or torn receipt"
        reprint_time = token.get("reprinted_at") or token.get("timestamp") or ""
        out.extend(ALIGN_LEFT)
        out.extend(pad_key_val("Reason", reason).encode("ascii", "replace") + b"\n")
        if reprint_time:
            out.extend(pad_key_val("Reprinted", reprint_time).encode("ascii", "replace") + b"\n")
        out.extend(ALIGN_CENTER)
        out.extend(divider("-").encode("ascii", "replace") + b"\n")

    # 4. Category & Price
    out.extend(BOLD_ON)
    cat_name = token.get("type", token.get("token_type", "Token")).upper()
    price = token.get("price", 0)
    cat_str = f"{cat_name} - Rs. {price}" if price and int(price) > 0 else cat_name
    out.extend(pad_center(cat_str).encode("ascii", "replace") + b"\n")
    out.extend(BOLD_OFF)
    out.extend(divider("-").encode("ascii", "replace") + b"\n")

    # 5. Serial Number (Prominent Double-Width & Double-Height)
    out.extend(BOLD_ON)
    out.extend(SIZE_DOUBLE)
    serial = str(token.get("serial", "")).strip()
    out.extend(f">> {serial} <<\n".encode("ascii", "replace"))
    out.extend(SIZE_NORMAL)
    out.extend(BOLD_OFF)
    out.extend(divider("-").encode("ascii", "replace") + b"\n")

    # 6. Devotee & Booking Details (Left-aligned, strict 48-char layout)
    out.extend(ALIGN_LEFT)
    name = token.get("name", "")
    mobile = token.get("mobile", "")
    payment = token.get("payment", "")
    counter = token.get("counter_name", "Counter")
    date_val = token.get("date", token.get("created_date", ""))
    time_val = token.get("time", token.get("created_time", ""))

    if price and int(price) > 0:
        out.extend(pad_key_val("Price", f"Rs. {price}").encode("ascii", "replace") + b"\n")
    out.extend(pad_key_val("Devotee Name", name).encode("ascii", "replace") + b"\n")
    out.extend(pad_key_val("Mobile No", mobile).encode("ascii", "replace") + b"\n")
    pay_display = "PENDING (NOT PAID)" if payment in ("Payment Pending", "Pending") else payment
    out.extend(pad_key_val("Payment Mode", pay_display).encode("ascii", "replace") + b"\n")
    if payment in ("Payment Pending", "Pending"):
        out.extend(BOLD_ON)
        out.extend(pad_center("*** PAYMENT PENDING - PLEASE COLLECT ***").encode("ascii", "replace") + b"\n")
        out.extend(BOLD_OFF)
    out.extend(pad_key_val("Counter", counter).encode("ascii", "replace") + b"\n")
    out.extend(pad_sides(f"Date: {date_val}", f"Time: {time_val}").encode("ascii", "replace") + b"\n")
    out.extend(divider("-").encode("ascii", "replace") + b"\n")

    # 7. Thank You Blessing Message
    out.extend(ALIGN_CENTER)
    out.extend(BOLD_ON)
    out.extend(b"Thank you for registration!\n")
    out.extend(BOLD_OFF)
    out.extend(b"May Goddess Durgamatha bless you and your family.\n")
    out.extend(divider("-").encode("ascii", "replace") + b"\n")

    # 8. Copy Label
    out.extend(BOLD_ON)
    out.extend(pad_center(f"[ {copy_label.upper()} ]").encode("ascii", "replace") + b"\n")
    out.extend(BOLD_OFF)

    # 9. Hardware Auto-Cutter Integration:
    # Append the reusable cut command to physically separate every slip
    out.extend(build_cut_command(cut_mode))

    return bytes(out)


def generate_escpos_stream_for_token(token: dict, include_office_copy: bool = True, is_reprint: bool = False, cut_mode: str = "full") -> bytes:
    """
    Generates complete raw ESC/POS stream for an individual token:
    Slip 1: Customer Copy (ends with Auto-Cutter command)
    Slip 2: Office Copy (ends with Auto-Cutter command)
    Ensures the printer cleanly cuts each slip individually.
    """
    stream = bytearray()
    # 1. Customer Copy -> Auto-Cut
    stream.extend(generate_escpos_slip(token, copy_label="CUSTOMER COPY", is_reprint=is_reprint, cut_mode=cut_mode))
    # 2. Office Copy -> Auto-Cut
    if include_office_copy:
        stream.extend(generate_escpos_slip(token, copy_label="OFFICE / COUNTER COPY", is_reprint=is_reprint, cut_mode=cut_mode))
    return bytes(stream)


def generate_escpos_streams_for_tokens(tokens_list: list, include_office_copy: bool = True, is_reprint: bool = False, cut_mode: str = "full") -> list:
    """
    Generates an itemized list of raw ESC/POS byte streams, one for each token.
    Each token stream contains its own slips and auto-cutter commands.
    """
    return [generate_escpos_stream_for_token(t, include_office_copy=include_office_copy, is_reprint=is_reprint, cut_mode=cut_mode) for t in tokens_list]


def get_installed_printers():
    """
    Enumerates local and network printers configured in Windows.
    """
    printers = []
    if win32print is not None:
        try:
            for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
                printers.append(p[2])
        except Exception:
            pass
    return printers


def send_to_thermal_printer(raw_bytes: bytes, printer_name: str = None, host: str = None, port: int = 9100):
    """
    Sends raw ESC/POS byte stream directly to a Windows thermal printer or TCP raw port 9100.
    Thread-safe to prevent overlapping print jobs.
    """
    with _printer_lock:
        # 1. TCP Network Printer (Ethernet / Wi-Fi thermal printer)
        host = host or os.environ.get("PRINTER_HOST")
        if host:
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.settimeout(5.0)
                s.connect((host, port))
                s.sendall(raw_bytes)
                s.close()
                return True, f"Printed successfully to network thermal printer at {host}:{port}"
            except Exception as e:
                return False, f"Network printer error ({host}:{port}): {e}"

        # 2. Windows Printer (USB / Spooler / Driver)
        if win32print is not None:
            target = printer_name or os.environ.get("PRINTER_NAME")
            try:
                if not target:
                    available = get_installed_printers()
                    # Auto-detect thermal / receipt printer name
                    for name in available:
                        nl = name.lower()
                        if any(k in nl for k in ["atpos", "pos", "thermal", "receipt", "at-301", "80mm", "tm-t", "tvs", "rp"]):
                            target = name
                            break
                    if not target:
                        target = win32print.GetDefaultPrinter()

                hprinter = win32print.OpenPrinter(target)
                try:
                    hjob = win32print.StartDocPrinter(hprinter, 1, ("SVARA_Receipt_80mm", None, "RAW"))
                    try:
                        win32print.StartPagePrinter(hprinter)
                        win32print.WritePrinter(hprinter, raw_bytes)
                        win32print.EndPagePrinter(hprinter)
                    finally:
                        win32print.EndDocPrinter(hprinter)
                finally:
                    win32print.ClosePrinter(hprinter)
                return True, f"Printed successfully to '{target}' with auto-cutter."
            except Exception as e:
                return False, f"Windows printer error ({target}): {e}"

        return False, "No direct printer driver available (win32print not installed or non-Windows system)."


def print_tokens_with_autocut(
    tokens_list: list,
    printer_name: str = None,
    host: str = None,
    port: int = 9100,
    include_office_copy: bool = True,
    is_reprint: bool = False,
    cut_mode: str = "full",
    inter_token_delay: float = 0.35
) -> tuple:
    """
    Unified reusable print-and-cut execution engine.
    Ensures EVERY individual token is physically separated by an auto-cut:
      Token 001 -> AUTO CUT -> Token 002 -> AUTO CUT -> Token 003 -> AUTO CUT
    Prevents printing multiple tokens as one continuous uncut strip.
    Handles printer buffer and print completion delays between tokens to prevent paper jams.
    Uses thread-safe locking to prevent overlapping print jobs.
    """
    if not tokens_list:
        return False, "No tokens provided to print."

    # Normalize if a single token dict was passed
    if isinstance(tokens_list, dict):
        tokens_list = [tokens_list]

    with _printer_lock:
        # 1. TCP Network Printer (Ethernet / Wi-Fi thermal printer)
        target_host = host or os.environ.get("PRINTER_HOST")
        if target_host:
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.settimeout(7.0)
                s.connect((target_host, port))
                try:
                    for idx, t in enumerate(tokens_list):
                        token_bytes = generate_escpos_stream_for_token(t, include_office_copy=include_office_copy, is_reprint=is_reprint, cut_mode=cut_mode)
                        s.sendall(token_bytes)
                        # Wait for printer buffer and auto-cutter blade completion before processing next token
                        if idx < len(tokens_list) - 1:
                            time.sleep(inter_token_delay)
                finally:
                    s.close()
                count = len(tokens_list)
                return True, f"Printed {count} token{'s' if count > 1 else ''} with auto-cutter to {target_host}:{port}."
            except Exception as e:
                return False, f"Network printer error ({target_host}:{port}): {e}"

        # 2. Windows Printer (USB / Spooler / Driver)
        if win32print is not None:
            target = printer_name or os.environ.get("PRINTER_NAME")
            try:
                if not target:
                    available = get_installed_printers()
                    for name in available:
                        nl = name.lower()
                        if any(k in nl for k in ["atpos", "pos", "thermal", "receipt", "at-301", "80mm", "tm-t", "tvs", "rp"]):
                            target = name
                            break
                    if not target:
                        target = win32print.GetDefaultPrinter()

                hprinter = win32print.OpenPrinter(target)
                try:
                    for idx, t in enumerate(tokens_list):
                        token_bytes = generate_escpos_stream_for_token(t, include_office_copy=include_office_copy, is_reprint=is_reprint, cut_mode=cut_mode)
                        doc_name = f"SVARA_Token_{t.get('serial', idx+1)}"
                        # Individual document job ensures driver flushes buffer and executes cut
                        hjob = win32print.StartDocPrinter(hprinter, 1, (doc_name, None, "RAW"))
                        try:
                            win32print.StartPagePrinter(hprinter)
                            win32print.WritePrinter(hprinter, token_bytes)
                            win32print.EndPagePrinter(hprinter)
                        finally:
                            win32print.EndDocPrinter(hprinter)

                        # Wait for printer buffer & physical knife completion before sending next token
                        if idx < len(tokens_list) - 1:
                            time.sleep(inter_token_delay)
                finally:
                    win32print.ClosePrinter(hprinter)

                count = len(tokens_list)
                return True, f"Printed {count} token{'s' if count > 1 else ''} with individual auto-cutter to '{target}'."
            except Exception as e:
                return False, f"Windows printer error ({target}): {e}"

        return False, "No direct printer driver available (win32print not installed or non-Windows system)."
