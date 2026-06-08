"""
LED Meeting Badge - Shows your next Outlook meeting on a USB LED name badge.
Fetches next meeting from Outlook via COM, pushes scrolling text to the
LED badge (LS32 / BMP Badge) via USB HID or Bluetooth Low Energy.

Requirements:
    pip install pyusb libusb pywin32    # for USB mode
    pip install bleak                    # for Bluetooth mode

Usage:
    python led_meeting_badge.py                  # USB one-shot
    python led_meeting_badge.py --loop           # USB refresh every 60s
    python led_meeting_badge.py --ble            # Bluetooth one-shot
    python led_meeting_badge.py --ble --loop     # Bluetooth refresh every 60s
    python led_meeting_badge.py --ble --scan     # scan for BLE badges
"""

import sys
import time
import asyncio
import argparse
from datetime import datetime, timedelta
from array import array


# ── Outlook COM: fetch next meeting ──────────────────────────────────────────

def fetch_next_meeting():
    """Return dict with subject, start, end, minutes_until for the next meeting, or None."""
    import pythoncom
    pythoncom.CoInitialize()
    try:
        import win32com.client
        ol = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        cal = ol.GetDefaultFolder(9)
        items = cal.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        now = datetime.now()
        end = now + timedelta(days=1)
        restriction = (
            "[Start] >= '" + now.strftime('%m/%d/%Y %H:%M %p') + "' AND "
            "[Start] <= '" + end.strftime('%m/%d/%Y %H:%M %p') + "'"
        )
        restricted = items.Restrict(restriction)

        for m in restricted:
            try:
                subj = getattr(m, "Subject", "")
                start = m.Start
                end_t = m.End
                s_dt = datetime(start.year, start.month, start.day, start.hour, start.minute)
                e_dt = datetime(end_t.year, end_t.month, end_t.day, end_t.hour, end_t.minute)
                mins = int((s_dt - now).total_seconds() / 60)
                if mins < -5:
                    continue  # skip meetings that ended > 5 min ago
                return {
                    "subject": subj,
                    "start": s_dt,
                    "end": e_dt,
                    "minutes_until": max(mins, 0),
                }
            except Exception:
                continue
        return None
    except Exception as e:
        print(f"Outlook error: {e}")
        return None
    finally:
        pythoncom.CoUninitialize()


def format_meeting_text(meeting):
    """Format meeting info into a short scrolling string for the LED badge."""
    if not meeting:
        return "NO MEETINGS"

    subj = meeting["subject"]
    start = meeting["start"].strftime("%H:%M")
    mins = meeting["minutes_until"]

    if mins == 0:
        urgency = "NOW"
    elif mins <= 5:
        urgency = f"IN {mins}MIN"
    elif mins <= 60:
        urgency = f"IN {mins}MIN"
    else:
        hours = mins // 60
        remaining = mins % 60
        if remaining > 0:
            urgency = f"IN {hours}H{remaining:02d}M"
        else:
            urgency = f"IN {hours}H"

    # Keep it short for the 44x11 LED display
    if len(subj) > 40:
        subj = subj[:37] + "..."

    return f"{urgency} {start} {subj}"


# ── LED Badge Protocol (LS32 / BMP Badge) ───────────────────────────────────
# Based on led-name-badge-ls32 by jnweiger (GPLv2+)
# USB HID: vendor=0x0416, product=0x5020

# 11-row font data for A-Z, a-z, 0-9, punctuation
FONT_11 = (
    # A-Z
    0x00,0x38,0x6c,0xc6,0xc6,0xfe,0xc6,0xc6,0xc6,0xc6,0x00,
    0x00,0xfc,0x66,0x66,0x66,0x7c,0x66,0x66,0x66,0xfc,0x00,
    0x00,0x7c,0xc6,0xc6,0xc0,0xc0,0xc0,0xc6,0xc6,0x7c,0x00,
    0x00,0xfc,0x66,0x66,0x66,0x66,0x66,0x66,0x66,0xfc,0x00,
    0x00,0xfe,0x66,0x62,0x68,0x78,0x68,0x62,0x66,0xfe,0x00,
    0x00,0xfe,0x66,0x62,0x68,0x78,0x68,0x60,0x60,0xf0,0x00,
    0x00,0x7c,0xc6,0xc6,0xc0,0xc0,0xce,0xc6,0xc6,0x7e,0x00,
    0x00,0xc6,0xc6,0xc6,0xc6,0xfe,0xc6,0xc6,0xc6,0xc6,0x00,
    0x00,0x3c,0x18,0x18,0x18,0x18,0x18,0x18,0x18,0x3c,0x00,
    0x00,0x1e,0x0c,0x0c,0x0c,0x0c,0x0c,0xcc,0xcc,0x78,0x00,
    0x00,0xe6,0x66,0x6c,0x6c,0x78,0x6c,0x6c,0x66,0xe6,0x00,
    0x00,0xf0,0x60,0x60,0x60,0x60,0x60,0x62,0x66,0xfe,0x00,
    0x00,0x82,0xc6,0xee,0xfe,0xd6,0xc6,0xc6,0xc6,0xc6,0x00,
    0x00,0x86,0xc6,0xe6,0xf6,0xde,0xce,0xc6,0xc6,0xc6,0x00,
    0x00,0x7c,0xc6,0xc6,0xc6,0xc6,0xc6,0xc6,0xc6,0x7c,0x00,
    0x00,0xfc,0x66,0x66,0x66,0x7c,0x60,0x60,0x60,0xf0,0x00,
    0x00,0x7c,0xc6,0xc6,0xc6,0xc6,0xc6,0xd6,0xde,0x7c,0x06,
    0x00,0xfc,0x66,0x66,0x66,0x7c,0x6c,0x66,0x66,0xe6,0x00,
    0x00,0x7c,0xc6,0xc6,0x60,0x38,0x0c,0xc6,0xc6,0x7c,0x00,
    0x00,0x7e,0x7e,0x5a,0x18,0x18,0x18,0x18,0x18,0x3c,0x00,
    0x00,0xc6,0xc6,0xc6,0xc6,0xc6,0xc6,0xc6,0xc6,0x7c,0x00,
    0x00,0xc6,0xc6,0xc6,0xc6,0xc6,0xc6,0x6c,0x38,0x10,0x00,
    0x00,0xc6,0xc6,0xc6,0xc6,0xd6,0xfe,0xee,0xc6,0x82,0x00,
    0x00,0xc6,0xc6,0x6c,0x7c,0x38,0x7c,0x6c,0xc6,0xc6,0x00,
    0x00,0x66,0x66,0x66,0x66,0x3c,0x18,0x18,0x18,0x3c,0x00,
    0x00,0xfe,0xc6,0x86,0x0c,0x18,0x30,0x62,0xc6,0xfe,0x00,
    # a-z
    0x00,0x00,0x00,0x00,0x78,0x0c,0x7c,0xcc,0xcc,0x76,0x00,
    0x00,0xe0,0x60,0x60,0x7c,0x66,0x66,0x66,0x66,0x7c,0x00,
    0x00,0x00,0x00,0x00,0x7c,0xc6,0xc0,0xc0,0xc6,0x7c,0x00,
    0x00,0x1c,0x0c,0x0c,0x7c,0xcc,0xcc,0xcc,0xcc,0x76,0x00,
    0x00,0x00,0x00,0x00,0x7c,0xc6,0xfe,0xc0,0xc6,0x7c,0x00,
    0x00,0x1c,0x36,0x30,0x78,0x30,0x30,0x30,0x30,0x78,0x00,
    0x00,0x00,0x00,0x00,0x76,0xcc,0xcc,0x7c,0x0c,0xcc,0x78,
    0x00,0xe0,0x60,0x60,0x6c,0x76,0x66,0x66,0x66,0xe6,0x00,
    0x00,0x18,0x18,0x00,0x38,0x18,0x18,0x18,0x18,0x3c,0x00,
    0x0c,0x0c,0x00,0x1c,0x0c,0x0c,0x0c,0x0c,0xcc,0xcc,0x78,
    0x00,0xe0,0x60,0x60,0x66,0x6c,0x78,0x78,0x6c,0xe6,0x00,
    0x00,0x38,0x18,0x18,0x18,0x18,0x18,0x18,0x18,0x3c,0x00,
    0x00,0x00,0x00,0x00,0xec,0xfe,0xd6,0xd6,0xd6,0xc6,0x00,
    0x00,0x00,0x00,0x00,0xdc,0x66,0x66,0x66,0x66,0x66,0x00,
    0x00,0x00,0x00,0x00,0x7c,0xc6,0xc6,0xc6,0xc6,0x7c,0x00,
    0x00,0x00,0x00,0x00,0xdc,0x66,0x66,0x7c,0x60,0x60,0xf0,
    0x00,0x00,0x00,0x00,0x7c,0xcc,0xcc,0x7c,0x0c,0x0c,0x1e,
    0x00,0x00,0x00,0x00,0xde,0x76,0x60,0x60,0x60,0xf0,0x00,
    0x00,0x00,0x00,0x00,0x7c,0xc6,0x70,0x1c,0xc6,0x7c,0x00,
    0x00,0x10,0x30,0x30,0xfc,0x30,0x30,0x30,0x34,0x18,0x00,
    0x00,0x00,0x00,0x00,0xcc,0xcc,0xcc,0xcc,0xcc,0x76,0x00,
    0x00,0x00,0x00,0x00,0xc6,0xc6,0xc6,0x6c,0x38,0x10,0x00,
    0x00,0x00,0x00,0x00,0xc6,0xd6,0xd6,0xd6,0xfe,0x6c,0x00,
    0x00,0x00,0x00,0x00,0xc6,0x6c,0x38,0x38,0x6c,0xc6,0x00,
    0x00,0x00,0x00,0x00,0xc6,0xc6,0xc6,0x7e,0x06,0x0c,0xf8,
    0x00,0x00,0x00,0x00,0xfe,0x8c,0x18,0x30,0x62,0xfe,0x00,
    # 0987654321
    0x00,0x7c,0xc6,0xce,0xde,0xf6,0xe6,0xc6,0xc6,0x7c,0x00,
    0x00,0x7c,0xc6,0xc6,0xc6,0x7e,0x06,0x06,0xc6,0x7c,0x00,
    0x00,0x7c,0xc6,0xc6,0xc6,0x7c,0xc6,0xc6,0xc6,0x7c,0x00,
    0x00,0xfe,0xc6,0x06,0x0c,0x18,0x30,0x30,0x30,0x30,0x00,
    0x00,0x7c,0xc6,0xc0,0xc0,0xfc,0xc6,0xc6,0xc6,0x7c,0x00,
    0x00,0xfe,0xc0,0xc0,0xfc,0x06,0x06,0x06,0xc6,0x7c,0x00,
    0x00,0x0c,0x1c,0x3c,0x6c,0xcc,0xfe,0x0c,0x0c,0x1e,0x00,
    0x00,0x7c,0xc6,0x06,0x06,0x3c,0x06,0x06,0xc6,0x7c,0x00,
    0x00,0x7c,0xc6,0x06,0x0c,0x18,0x30,0x60,0xc6,0xfe,0x00,
    0x00,0x18,0x38,0x78,0x18,0x18,0x18,0x18,0x18,0x7e,0x00,
    # ^ !"$%&/()=?` space
    0x38,0x6c,0xc6,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
    0x00,0x00,0x00,0x00,0x00,0x40,0x3c,0x00,0x00,0x00,0x00,
    0x00,0x18,0x3c,0x3c,0x3c,0x18,0x18,0x00,0x18,0x18,0x00,
    0x66,0x66,0x22,0x22,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
    0x10,0x7c,0xd6,0xd6,0x70,0x1c,0xd6,0xd6,0x7c,0x10,0x10,
    0x00,0x60,0x92,0x96,0x6c,0x10,0x6c,0xd2,0x92,0x0c,0x00,
    0x00,0x38,0x6c,0x6c,0x38,0x76,0xdc,0xcc,0xcc,0x76,0x00,
    0x00,0x00,0x02,0x06,0x0c,0x18,0x30,0x60,0xc0,0x80,0x00,
    0x00,0x0c,0x18,0x30,0x30,0x30,0x30,0x30,0x18,0x0c,0x00,
    0x00,0x30,0x18,0x0c,0x0c,0x0c,0x0c,0x0c,0x18,0x30,0x00,
    0x00,0x00,0x00,0x7e,0x00,0x00,0x7e,0x00,0x00,0x00,0x00,
    0x00,0x7c,0xc6,0xc6,0x0c,0x18,0x18,0x00,0x18,0x18,0x00,
    0x18,0x18,0x10,0x08,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
    0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,
    # .:- (dot, colon, dash)
    0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x30,0x30,0x00,
    0x00,0x00,0x00,0x18,0x18,0x00,0x00,0x18,0x18,0x00,0x00,
    0x00,0x00,0x00,0x00,0x00,0xfe,0x00,0x00,0x00,0x00,0x00,
)

CHARMAP = (
    'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    'abcdefghijklmnopqrstuvwxyz'
    '0987654321'
    '^ !"$%&/()=?` '
    '.:- '  # extra punctuation we added (reusing from the full font)
)

# Build lookup: char -> offset in FONT_11
CHAR_OFFSETS = {}
# We'll handle the charmap carefully
_FULL_CHARMAP = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0987654321^ !"$%&/()=?` .:-'
for _i, _ch in enumerate(_FULL_CHARMAP):
    CHAR_OFFSETS[_ch] = 11 * _i


def text_to_bitmap(text):
    """Convert text string to LED badge bitmap (11 bytes per character column)."""
    buf = array('B')
    cols = 0
    for ch in text:
        if ch in CHAR_OFFSETS:
            off = CHAR_OFFSETS[ch]
            buf.extend(FONT_11[off:off + 11])
            cols += 1
        elif ch.upper() in CHAR_OFFSETS:
            off = CHAR_OFFSETS[ch.upper()]
            buf.extend(FONT_11[off:off + 11])
            cols += 1
        else:
            # Unknown char -> space
            off = CHAR_OFFSETS.get(' ', 0)
            buf.extend(FONT_11[off:off + 11])
            cols += 1
    return buf, cols


# ── Protocol header ──────────────────────────────────────────────────────────

HEADER_TEMPLATE = (
    0x77, 0x61, 0x6e, 0x67, 0x00, 0x00, 0x00, 0x00,
    0x40, 0x40, 0x40, 0x40, 0x40, 0x40, 0x40, 0x40,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
)


def make_header(length, speed=4, mode=0, blink=0, ants=0, brightness=100):
    """Build the 64-byte protocol header for one message.
    speed: 1-8, mode: 0=scroll-left, brightness: 25/50/75/100
    """
    h = list(HEADER_TEMPLATE)

    # Brightness
    if brightness <= 25:
        h[5] = 0x40
    elif brightness <= 50:
        h[5] = 0x20
    elif brightness <= 75:
        h[5] = 0x10

    # Blink / ants for message 0
    h[6] = blink & 1
    h[7] = ants & 1

    # Speed (0-7 internal) + mode for message 0
    h[8] = 16 * max(0, min(7, speed - 1)) + max(0, min(8, mode))

    # Length of message 0 (in byte-columns)
    h[16] = length // 256
    h[17] = length % 256

    # Timestamp
    now = datetime.now()
    h[38] = now.year % 100
    h[39] = now.month
    h[40] = now.day
    h[41] = now.hour
    h[42] = now.minute
    h[43] = now.second

    return h


# ── USB Write ────────────────────────────────────────────────────────────────

def _get_usb_backend():
    """Get the libusb backend, trying the bundled 'libusb' package first."""
    try:
        import usb.backend.libusb1
        import libusb as libusb_pkg
        # The 'libusb' pip package bundles the DLL
        dll_path = libusb_pkg.dll._name if hasattr(libusb_pkg, 'dll') and libusb_pkg.dll else None
        if dll_path:
            import ctypes.util
            be = usb.backend.libusb1.get_backend(find_library=lambda x: dll_path)
            if be:
                return be
    except Exception:
        pass
    # Fallback: let pyusb find it on its own (needs libusb-win32 driver)
    return None


def write_to_badge(buf):
    """Write buffer to LED badge via USB (pyusb / libusb)."""
    # Pad to 64-byte blocks
    remainder = len(buf) % 64
    if remainder:
        buf.extend((0,) * (64 - remainder))

    if len(buf) > 8192:
        print("Data too large for badge! Max 8192 bytes.")
        return False

    try:
        import usb.core
        import usb.util
    except ImportError:
        print("pyusb not installed. Run: pip install pyusb")
        return False

    backend = _get_usb_backend()
    dev = usb.core.find(idVendor=0x0416, idProduct=0x5020, backend=backend)
    if dev is None:
        print("LED badge not found! Check USB connection.")
        print("  - Is the badge plugged in?")
        print("  - Is the libusb-win32 driver installed? (Use Zadig)")
        return False

    try:
        if dev.is_kernel_driver_active(0):
            dev.detach_kernel_driver(0)
    except (NotImplementedError, Exception):
        pass

    try:
        dev.set_configuration()
    except Exception as e:
        print(f"USB config error: {e}")
        print("Try running as Administrator.")
        return False

    cfg = dev.get_active_configuration()[(0, 0)]
    ep = usb.util.find_descriptor(
        cfg,
        custom_match=lambda e: usb.util.endpoint_direction(e.bEndpointAddress) == usb.util.ENDPOINT_OUT
    )

    if ep is None:
        print("Could not find USB endpoint!")
        return False

    print(f"Writing {len(buf)} bytes to LED badge...")
    for i in range(len(buf) // 64):
        time.sleep(0.1)
        ep.write(buf[i * 64: i * 64 + 64])

    print("Done!")
    return True


# ── Bluetooth Low Energy Write ───────────────────────────────────────────────
# Badge advertises as "LSLED" over BLE
# Characteristic UUID: 0000fee1-0000-1000-8000-00805f9b34fb
# Same protocol header, but sent in 16-byte chunks instead of 64-byte

LSLED_BLE_NAME = "LSLED"
LSLED_CHAR_UUID = "0000fee1-0000-1000-8000-00805f9b34fb"
BLE_CHUNK_SIZE = 16


async def _ble_scan():
    """Scan for BLE LED badges and print results."""
    from bleak import BleakScanner
    print("Scanning for BLE LED badges (10 seconds)...")
    devices = await BleakScanner.discover(timeout=10.0, return_adv=True)
    found = []
    for d, adv in devices.values():
        name = d.name or adv.local_name or ""
        rssi = adv.rssi if adv else ""
        if "LSLED" in name.upper() or "LED" in name.upper() or "BADGE" in name.upper():
            found.append(d)
            print(f"  >> MATCH: {name} [{d.address}] RSSI={rssi}")
    if not found:
        print("No LED badges found. Nearby BLE devices:")
        sorted_devs = sorted(devices.values(), key=lambda x: x[1].rssi if x[1] else -999, reverse=True)
        for d, adv in sorted_devs[:15]:
            name = d.name or adv.local_name or "(unnamed)"
            rssi = adv.rssi if adv else "?"
            print(f"  {name} [{d.address}] RSSI={rssi}")
    return found


def ble_scan():
    """Synchronous wrapper for BLE scan."""
    asyncio.run(_ble_scan())


async def _write_to_badge_ble(buf, device_address=None):
    """Write buffer to LED badge via Bluetooth Low Energy."""
    from bleak import BleakClient, BleakScanner

    # Pad to 16-byte blocks for BLE
    remainder = len(buf) % BLE_CHUNK_SIZE
    if remainder:
        buf.extend((0,) * (BLE_CHUNK_SIZE - remainder))

    device = None

    # Try direct address first, fall back to scan
    if device_address:
        print(f"Connecting to BLE badge at {device_address}...")
        try:
            async with BleakClient(device_address, timeout=10.0) as client:
                print(f"Connected! Writing {len(buf)} bytes in {len(buf) // BLE_CHUNK_SIZE} chunks...")
                for i in range(len(buf) // BLE_CHUNK_SIZE):
                    chunk = bytes(buf[i * BLE_CHUNK_SIZE: (i + 1) * BLE_CHUNK_SIZE])
                    await client.write_gatt_char(LSLED_CHAR_UUID, chunk, response=True)
                print("Done!")
                return True
        except Exception as e:
            print(f"Direct connect failed ({e}), scanning...")

    # Scan fallback
    print("Scanning for LSLED badge via BLE...")
    device = await BleakScanner.find_device_by_name(LSLED_BLE_NAME, timeout=10.0)
    if device is None:
        print("BLE badge not found! Make sure:")
        print("  - Badge is powered on (not connected to USB)")
        print("  - Bluetooth is enabled on your PC")
        print("  - Badge is within range")
        print("  Use --ble --scan to see nearby devices")
        return False

    print(f"Connecting to {device}...")
    async with BleakClient(device, timeout=10.0) as client:
        print(f"Connected! Writing {len(buf)} bytes in {len(buf) // BLE_CHUNK_SIZE} chunks...")
        for i in range(len(buf) // BLE_CHUNK_SIZE):
            chunk = bytes(buf[i * BLE_CHUNK_SIZE: (i + 1) * BLE_CHUNK_SIZE])
            await client.write_gatt_char(LSLED_CHAR_UUID, chunk, response=True)
        print("Done!")
    return True


def write_to_badge_ble(buf, device_address=None):
    """Synchronous wrapper for BLE write."""
    return asyncio.run(_write_to_badge_ble(buf, device_address))


# ── Main ─────────────────────────────────────────────────────────────────────

def push_meeting_to_badge(speed=4, mode=0, brightness=100, use_ble=False, ble_address=None):
    """Fetch next meeting and push to LED badge. Returns True on success."""
    meeting = fetch_next_meeting()
    text = format_meeting_text(meeting)
    transport = "BLE" if use_ble else "USB"
    print(f"[{datetime.now().strftime('%H:%M:%S')}] ({transport}) Badge text: {text}")

    bitmap, cols = text_to_bitmap(text)
    header = make_header(cols, speed=speed, mode=mode, brightness=brightness)

    buf = array('B')
    buf.extend(header)
    buf.extend(bitmap)

    if use_ble:
        return write_to_badge_ble(buf, ble_address)
    else:
        return write_to_badge(buf)


def main():
    parser = argparse.ArgumentParser(description="Show next Outlook meeting on USB/BLE LED badge")
    parser.add_argument("--loop", nargs="?", const=60, type=int, metavar="SECONDS",
                        help="Refresh every N seconds (default 60)")
    parser.add_argument("--speed", type=int, default=4, choices=range(1, 9),
                        help="Scroll speed 1-8 (default 4)")
    parser.add_argument("--mode", type=int, default=0,
                        help="Display mode: 0=scroll-left, 1=scroll-right, 4=static, 5=animation")
    parser.add_argument("--brightness", type=int, default=100, choices=[25, 50, 75, 100],
                        help="Brightness percent (default 100)")
    parser.add_argument("--text", type=str, default=None,
                        help="Override: send custom text instead of meeting")
    parser.add_argument("--ble", action="store_true",
                        help="Use Bluetooth Low Energy instead of USB")
    parser.add_argument("--ble-address", type=str, default=None, metavar="ADDR",
                        help="BLE device address (skip scan, e.g. AA:BB:CC:DD:EE:FF)")
    parser.add_argument("--scan", action="store_true",
                        help="Scan for BLE LED badges and exit")
    args = parser.parse_args()

    if args.scan:
        ble_scan()
        return

    write_fn = write_to_badge_ble if args.ble else write_to_badge

    if args.text:
        text = args.text.upper()
        print(f"Sending custom text via {'BLE' if args.ble else 'USB'}: {text}")
        bitmap, cols = text_to_bitmap(text)
        header = make_header(cols, speed=args.speed, mode=args.mode, brightness=args.brightness)
        buf = array('B')
        buf.extend(header)
        buf.extend(bitmap)
        if args.ble:
            write_to_badge_ble(buf, args.ble_address)
        else:
            write_to_badge(buf)
        return

    if args.loop:
        transport = "BLE" if args.ble else "USB"
        print(f"LED Meeting Badge ({transport}) - refreshing every {args.loop}s (Ctrl+C to stop)")
        while True:
            try:
                push_meeting_to_badge(
                    speed=args.speed, mode=args.mode, brightness=args.brightness,
                    use_ble=args.ble, ble_address=args.ble_address
                )
            except Exception as e:
                print(f"Error: {e}")
            time.sleep(args.loop)
    else:
        push_meeting_to_badge(
            speed=args.speed, mode=args.mode, brightness=args.brightness,
            use_ble=args.ble, ble_address=args.ble_address
        )


if __name__ == "__main__":
    main()
