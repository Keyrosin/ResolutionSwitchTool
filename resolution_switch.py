"""Toggle the primary display between two resolutions on Windows."""
from __future__ import annotations

import argparse
import ctypes
import sys
from ctypes import Structure, byref, sizeof
from ctypes.wintypes import DWORD, WCHAR


if sys.platform != "win32":
    raise SystemExit("This tool can only run on Windows.")

user32 = ctypes.windll.user32


WORD = ctypes.c_ushort
SHORT = ctypes.c_short


class DEVMODE(Structure):
    _fields_ = [
        ("dmDeviceName", WCHAR * 32),
        ("dmSpecVersion", WORD),
        ("dmDriverVersion", WORD),
        ("dmSize", WORD),
        ("dmDriverExtra", WORD),
        ("dmFields", DWORD),
        ("dmOrientation", SHORT),
        ("dmPaperSize", SHORT),
        ("dmPaperLength", SHORT),
        ("dmPaperWidth", SHORT),
        ("dmScale", SHORT),
        ("dmCopies", SHORT),
        ("dmDefaultSource", SHORT),
        ("dmPrintQuality", SHORT),
        ("dmColor", SHORT),
        ("dmDuplex", SHORT),
        ("dmYResolution", SHORT),
        ("dmTTOption", SHORT),
        ("dmCollate", SHORT),
        ("dmFormName", WCHAR * 32),
        ("dmLogPixels", WORD),
        ("dmBitsPerPel", DWORD),
        ("dmPelsWidth", DWORD),
        ("dmPelsHeight", DWORD),
        ("dmDisplayFlags", DWORD),
        ("dmDisplayFrequency", DWORD),
        ("dmICMMethod", DWORD),
        ("dmICMIntent", DWORD),
        ("dmMediaType", DWORD),
        ("dmDitherType", DWORD),
        ("dmReserved1", DWORD),
        ("dmReserved2", DWORD),
        ("dmPanningWidth", DWORD),
        ("dmPanningHeight", DWORD),
    ]


DM_BITSPERPEL = 0x00040000
DM_PELSWIDTH = 0x00080000
DM_PELSHEIGHT = 0x00100000
DM_DISPLAYFREQUENCY = 0x00400000


ENUM_CURRENT_SETTINGS = DWORD(-1).value
CDS_UPDATEREGISTRY = 0x00000001
CDS_FULLSCREEN = 0x00000004
DISP_CHANGE_SUCCESSFUL = 0


DISPLAY_MODES = (
    {"width": 2560, "height": 1440},
    {"width": 1680, "height": 1050},
)


def enum_display_settings(mode_number: int) -> DEVMODE | None:
    devmode = DEVMODE()
    devmode.dmSize = sizeof(DEVMODE)
    if not user32.EnumDisplaySettingsW(None, mode_number, byref(devmode)):
        return None
    return devmode


def get_current_settings() -> DEVMODE:
    devmode = enum_display_settings(ENUM_CURRENT_SETTINGS)
    if devmode is None:
        raise RuntimeError("Unable to read current display settings")
    return devmode


def find_max_refresh_rate(width: int, height: int) -> int | None:
    max_rate: int | None = None
    mode_index = 0
    while True:
        devmode = enum_display_settings(mode_index)
        if devmode is None:
            break
        if devmode.dmPelsWidth == width and devmode.dmPelsHeight == height:
            rate = int(devmode.dmDisplayFrequency)
            if rate:
                max_rate = rate if max_rate is None or rate > max_rate else max_rate
        mode_index += 1
    return max_rate


def build_devmode(width: int, height: int, refresh_rate: int | None = None) -> DEVMODE:
    base = get_current_settings()
    target = DEVMODE()
    ctypes.memmove(byref(target), byref(base), sizeof(DEVMODE))
    target.dmFields = DM_BITSPERPEL | DM_PELSWIDTH | DM_PELSHEIGHT
    target.dmBitsPerPel = base.dmBitsPerPel
    target.dmPelsWidth = width
    target.dmPelsHeight = height

    if refresh_rate is None:
        refresh_rate = find_max_refresh_rate(width, height)
    if refresh_rate:
        target.dmFields |= DM_DISPLAYFREQUENCY
        target.dmDisplayFrequency = refresh_rate
    return target


def apply_settings(devmode: DEVMODE) -> None:
    result = user32.ChangeDisplaySettingsExW(
        None,
        byref(devmode),
        None,
        CDS_UPDATEREGISTRY | CDS_FULLSCREEN,
        None,
    )
    if result != DISP_CHANGE_SUCCESSFUL:
        raise RuntimeError(f"Display change failed with code {result}")


def determine_target_mode(current: DEVMODE, force: tuple[int, int] | None) -> tuple[int, int]:
    if force is not None:
        return force

    current_size = (int(current.dmPelsWidth), int(current.dmPelsHeight))
    if current_size == (DISPLAY_MODES[0]["width"], DISPLAY_MODES[0]["height"]):
        return DISPLAY_MODES[1]["width"], DISPLAY_MODES[1]["height"]
    if current_size == (DISPLAY_MODES[1]["width"], DISPLAY_MODES[1]["height"]):
        return DISPLAY_MODES[0]["width"], DISPLAY_MODES[0]["height"]
    # Default to the first resolution if the current size doesn't match either.
    return DISPLAY_MODES[0]["width"], DISPLAY_MODES[0]["height"]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--set",
        metavar="WIDTHxHEIGHT",
        help="Set a specific resolution instead of toggling between the presets.",
    )
    parser.add_argument(
        "--refresh",
        type=int,
        default=None,
        help="Force a refresh rate (Hz). Defaults to the maximum available for the chosen resolution.",
    )
    return parser.parse_args()


def parse_resolution(value: str) -> tuple[int, int]:
    try:
        width_str, height_str = value.lower().split("x", 1)
        return int(width_str), int(height_str)
    except Exception as exc:  # noqa: BLE001
        raise argparse.ArgumentTypeError(f"Invalid resolution format: {value!r}") from exc


def main() -> None:
    args = parse_args()
    forced_resolution: tuple[int, int] | None = None
    if args.set:
        forced_resolution = parse_resolution(args.set)

    current = get_current_settings()
    target_width, target_height = determine_target_mode(current, forced_resolution)

    devmode = build_devmode(target_width, target_height, args.refresh)
    apply_settings(devmode)
    print(
        f"Switched to {target_width}x{target_height} at {devmode.dmDisplayFrequency or 'default'} Hz",
    )


if __name__ == "__main__":
    main()
