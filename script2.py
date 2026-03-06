import json  # Read/write JSON reports.
import math  # Math helpers (LCD diagonal calculation).
import platform  # Detect OS/platform details.
import re  # Regex parsing for text extraction.
import shutil  # Check command availability (e.g., sudo).
import socket  # Get hostname.
import subprocess  # Run system/PowerShell commands.
from datetime import datetime  # Timestamp report generation.
from pathlib import Path  # Safe file path/file reading helpers.


def decode_cmd_output(raw):
    """Decode subprocess bytes safely across common Windows/Linux encodings."""
    if raw is None:
        return ""
    if isinstance(raw, str):
        return raw
    for enc in ("utf-8", "cp1252", "cp850", "latin-1"):
        try:
            return raw.decode(enc)
        except Exception:
            continue
    return raw.decode("utf-8", errors="replace")


def run_cmd(cmd, timeout=30, shell=False, use_sudo=False):
    """Run a system command and return (success: bool, output: str)."""
    if isinstance(cmd, str) and not shell:
        cmd = cmd.split()

    final_cmd = cmd
    if use_sudo and platform.system() == "Linux" and shutil.which("sudo") and not shell:
        final_cmd = ["sudo", "-n"] + cmd

    try:
        result = subprocess.run(
            final_cmd,
            capture_output=True,
            text=False,
            timeout=timeout,
            check=False,
            shell=shell,
        )
        stdout = decode_cmd_output(result.stdout).strip()
        stderr = decode_cmd_output(result.stderr).strip()
        if result.returncode == 0:
            return True, stdout
        err = stderr or stdout or "Unknown command error"
        return False, err
    except Exception as exc:
        return False, str(exc)


def run_ps(command, timeout=30):
    """Run a PowerShell command and return (success: bool, output: str)."""
    return run_cmd(["powershell", "-NoProfile", "-Command", command], timeout=timeout)


def run_ps_json(command):
    """Run PowerShell and parse JSON output; return dict/list or None."""
    ok, out = run_ps(command)
    if not ok or not out:
        return None
    try:
        return json.loads(out)
    except Exception:
        return None


def read_file(path):
    """Read a text file and return stripped content, or empty string on failure."""
    try:
        return Path(path).read_text(encoding="utf-8", errors="ignore").strip()
    except Exception:
        return ""


def parse_key_value_lines(text, sep=":"):
    """Parse 'key: value' lines into a dictionary."""
    data = {}
    for line in text.splitlines():
        if sep in line:
            k, v = line.split(sep, 1)
            data[k.strip()] = v.strip()
    return data


def bytes_to_gib(value):
    """Convert bytes to GiB string (binary units)."""
    try:
        return f"{float(value) / (1024 ** 3):.2f} GiB"
    except Exception:
        return "Unknown"


def bytes_to_installed_gb(value):
    """Convert bytes to rounded installed GB value (binary base, human label)."""
    try:
        gb = float(value) / (1024 ** 3)
        return f"{int(round(gb))}GB"
    except Exception:
        return "Unknown"


def bytes_to_marketed_gb(value):
    """Convert bytes to marketed GB value (decimal base)."""
    try:
        gb = float(value) / 1_000_000_000
        return f"{int(round(gb))}GB"
    except Exception:
        return "Unknown"


def memory_type_label(code):
    """Map SMBIOS memory type code to readable DDR generation label."""
    mapping = {
        20: "DDR",
        21: "DDR2",
        24: "DDR3",
        26: "DDR4",
        34: "DDR5",
    }
    return mapping.get(int(code), "Unknown")


def uniq_keep_order(items):
    """Return unique non-empty strings while preserving initial order."""
    seen = set()
    out = []
    for item in items:
        key = item.strip().lower()
        if key and key not in seen:
            seen.add(key)
            out.append(item.strip())
    return out


def infer_cpu_base_ghz_from_name(model_name):
    """Extract base GHz from CPU model string like '@ 2.40GHz'."""
    if not model_name:
        return None
    m = re.search(r"@\s*([0-9]+(?:\.[0-9]+)?)\s*GHz", model_name, flags=re.IGNORECASE)
    if m:
        return f"{float(m.group(1)):.2f} GHz"
    return None


def get_os_info():
    """Collect OS metadata (name, version, build/release) for Windows/Linux."""
    system = platform.system()
    info = {
        "system": system,
        "release": platform.release(),
        "version": platform.version(),
        "machine": platform.machine(),
        "pretty_name": "",
        "kernel": "",
        "display_version": "",
    }

    if system == "Linux":
        os_release = read_file("/etc/os-release")
        if os_release:
            for line in os_release.splitlines():
                if line.startswith("PRETTY_NAME="):
                    info["pretty_name"] = line.split("=", 1)[1].strip().strip('"')
                    break
        ok, out = run_cmd(["uname", "-sr"])
        if ok:
            info["kernel"] = out

    elif system == "Windows":
        os_data = run_ps_json(
            "Get-CimInstance Win32_OperatingSystem | "
            "Select-Object Caption,OSArchitecture,Version,BuildNumber | ConvertTo-Json -Compress"
        )
        if isinstance(os_data, dict):
            caption = os_data.get("Caption", "Windows")
            arch = os_data.get("OSArchitecture", "")
            info["pretty_name"] = f"{caption} {arch}".strip()
            info["version"] = os_data.get("Version", info["version"])
            info["kernel"] = f"Build {os_data.get('BuildNumber', '')}".strip()

        ver_data = run_ps_json(
            "Get-ItemProperty 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion' | "
            "Select-Object DisplayVersion,ReleaseId | ConvertTo-Json -Compress"
        )
        if isinstance(ver_data, dict):
            info["display_version"] = ver_data.get("DisplayVersion") or ver_data.get("ReleaseId") or ""

    return info


def get_model_info():
    """Collect machine vendor/model/serial information."""
    if platform.system() == "Windows":
        cs = run_ps_json(
            "Get-CimInstance Win32_ComputerSystem | "
            "Select-Object Manufacturer,Model | ConvertTo-Json -Compress"
        )
        bios = run_ps_json(
            "Get-CimInstance Win32_BIOS | Select-Object SerialNumber | ConvertTo-Json -Compress"
        )
        return {
            "vendor": (cs or {}).get("Manufacturer", "Unknown"),
            "model": (cs or {}).get("Model", "Unknown"),
            "serial": (bios or {}).get("SerialNumber", "Unknown"),
        }

    model = read_file("/sys/devices/virtual/dmi/id/product_name")
    vendor = read_file("/sys/devices/virtual/dmi/id/sys_vendor")
    serial = read_file("/sys/devices/virtual/dmi/id/product_serial")
    return {"vendor": vendor or "Unknown", "model": model or "Unknown", "serial": serial or "Unknown"}


def get_cpu_info():
    """Collect CPU model, base clock, and core/thread counts."""
    if platform.system() == "Windows":
        cpu = run_ps_json(
            "Get-CimInstance Win32_Processor | "
            "Select-Object Name,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed | "
            "ConvertTo-Json -Compress"
        )
        if isinstance(cpu, list):
            cpu = cpu[0] if cpu else {}
        cpu = cpu or {}

        model_name = cpu.get("Name", "Unknown")
        base_from_name = infer_cpu_base_ghz_from_name(model_name)
        base_from_clock = f"{float(cpu.get('MaxClockSpeed', 0)) / 1000:.2f} GHz" if cpu.get("MaxClockSpeed") else None

        return {
            "model_name": model_name,
            "base_clock_ghz": base_from_name or base_from_clock or "Unknown",
            "cores_physical": str(cpu.get("NumberOfCores", "Unknown")),
            "cores_logical": str(cpu.get("NumberOfLogicalProcessors", "Unknown")),
        }

    ok, out = run_cmd(["lscpu"])
    if not ok:
        return {"model_name": "Unknown", "base_clock_ghz": "Unknown", "cores_physical": "Unknown", "cores_logical": "Unknown"}
    data = parse_key_value_lines(out)
    model_name = data.get("Model name", "Unknown")
    base_from_name = infer_cpu_base_ghz_from_name(model_name)
    base_from_clock = f"{float(data.get('CPU max MHz', '0')) / 1000:.2f} GHz" if data.get("CPU max MHz") else None
    return {
        "model_name": model_name,
        "base_clock_ghz": base_from_name or base_from_clock or "Unknown",
        "cores_physical": data.get("Core(s) per socket", "Unknown"),
        "cores_logical": data.get("CPU(s)", "Unknown"),
    }


def get_ram_info():
    """Collect installed RAM size/type and module details when available."""
    if platform.system() == "Windows":
        modules = run_ps_json(
            "Get-CimInstance Win32_PhysicalMemory | "
            "Select-Object Capacity,SMBIOSMemoryType,ConfiguredClockSpeed,Manufacturer,PartNumber | "
            "ConvertTo-Json -Compress"
        )
        if isinstance(modules, dict):
            modules = [modules]

        if isinstance(modules, list) and modules:
            total_bytes = 0
            mem_types = []
            parsed_modules = []
            for m in modules:
                cap = int(m.get("Capacity") or 0)
                total_bytes += cap
                mt = memory_type_label(m.get("SMBIOSMemoryType") or 0)
                if mt != "Unknown":
                    mem_types.append(mt)
                parsed_modules.append(
                    {
                        "capacity_bytes": cap,
                        "capacity_installed": bytes_to_installed_gb(cap),
                        "type": mt,
                        "speed_mhz": m.get("ConfiguredClockSpeed"),
                        "manufacturer": m.get("Manufacturer"),
                        "part_number": (m.get("PartNumber") or "").strip(),
                    }
                )

            mem_type = mem_types[0] if mem_types else "Unknown"
            return {
                "installed_bytes": total_bytes,
                "installed_human": bytes_to_installed_gb(total_bytes),
                "installed_gib": bytes_to_gib(total_bytes),
                "type": mem_type,
                "modules": parsed_modules,
            }

        ok, out = run_ps("(Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory")
        raw = out.strip() if ok else "0"
        return {
            "installed_bytes": int(raw or 0),
            "installed_human": bytes_to_installed_gb(raw),
            "installed_gib": bytes_to_gib(raw),
            "type": "Unknown",
            "modules": [],
        }

    mem_total_kb = ""
    meminfo = read_file("/proc/meminfo")
    for line in meminfo.splitlines():
        if line.startswith("MemTotal:"):
            mem_total_kb = line.split(":", 1)[1].strip()
            break
    return {
        "installed_bytes": 0,
        "installed_human": mem_total_kb or "Unknown",
        "installed_gib": "Unknown",
        "type": "Unknown",
        "modules": [],
    }


def get_disks_info():
    """Collect physical disk inventory (model, serial, size, transport)."""
    if platform.system() == "Windows":
        disks = run_ps_json(
            "Get-CimInstance Win32_DiskDrive | "
            "Select-Object DeviceID,Model,SerialNumber,Size,InterfaceType,MediaType | ConvertTo-Json -Compress"
        )
        if isinstance(disks, dict):
            disks = [disks]
        if not isinstance(disks, list):
            return []

        parsed = []
        for d in disks:
            size = d.get("Size")
            model = (d.get("Model") or "Unknown").strip()
            bus_guess = "NVMe" if "nvme" in model.lower() else (d.get("InterfaceType") or "")
            parsed.append(
                {
                    "name": d.get("DeviceID"),
                    "model": model,
                    "serial": (d.get("SerialNumber") or "").strip(),
                    "size_bytes": int(size) if str(size).isdigit() else 0,
                    "size_gb": bytes_to_marketed_gb(size),
                    "transport": bus_guess,
                    "media_type": d.get("MediaType"),
                }
            )

        parsed.sort(key=lambda x: x.get("size_bytes", 0), reverse=True)
        return parsed

    ok, out = run_cmd(["lsblk", "-J", "-o", "NAME,SIZE,MODEL,SERIAL,TYPE,TRAN"])
    if not ok:
        return []
    try:
        data = json.loads(out)
    except Exception:
        return []
    disks = []
    for dev in data.get("blockdevices", []):
        if dev.get("type") == "disk":
            disks.append(
                {
                    "name": dev.get("name"),
                    "model": dev.get("model"),
                    "serial": dev.get("serial"),
                    "size_gb": dev.get("size"),
                    "transport": dev.get("tran"),
                }
            )
    return disks


def infer_optical_type(name_text, media_text=""):
    """Infer optical drive type (DVD-RW/DVD-ROM/etc.) from device strings."""
    text = f"{name_text} {media_text}".lower()
    if "dvd-rw" in text or "dvdrw" in text or "dvd+rw" in text:
        return "DVD-RW"
    if "dvd-rom" in text or "dvdrom" in text:
        return "DVD-ROM"
    if "bd-re" in text or "blu-ray" in text or "bluray" in text:
        return "Blu-ray"
    if "cd-rw" in text or "cdrw" in text:
        return "CD-RW"
    if "cd-rom" in text or "cdrom" in text:
        return "CD-ROM"
    if "dvd" in text:
        return "DVD"
    if "cd" in text:
        return "CD"
    return "Optical drive"


def get_mechanika_info():
    """Detect optical drive presence and return detailed drive type list."""
    if platform.system() == "Windows":
        drives = run_ps_json(
            "Get-CimInstance Win32_CDROMDrive | "
            "Select-Object Name,MediaType,Drive | ConvertTo-Json -Compress"
        )
        if isinstance(drives, dict):
            drives = [drives]
        if not isinstance(drives, list) or not drives:
            return {"present": False, "items": []}

        items = []
        for d in drives:
            name = (d.get("Name") or "").strip()
            media = (d.get("MediaType") or "").strip()
            dtype = infer_optical_type(name, media)
            details = f"{dtype}"
            if name:
                details += f" ({name})"
            items.append(details)
        return {"present": True, "items": uniq_keep_order(items)}

    # Linux: type "rom" from lsblk indicates optical drive.
    ok, out = run_cmd(["lsblk", "-J", "-o", "NAME,TYPE,MODEL"])
    if not ok:
        return {"present": False, "items": []}
    try:
        data = json.loads(out)
    except Exception:
        return {"present": False, "items": []}

    items = []
    for dev in data.get("blockdevices", []):
        if dev.get("type") == "rom":
            model = (dev.get("model") or "").strip()
            dtype = infer_optical_type(model)
            if model:
                items.append(f"{dtype} ({model})")
            else:
                items.append(dtype)

    return {"present": bool(items), "items": uniq_keep_order(items)}


def get_gpu_info():
    """Collect GPU adapter names."""
    if platform.system() == "Windows":
        gpus = run_ps_json(
            "Get-CimInstance Win32_VideoController | Select-Object Name | ConvertTo-Json -Compress"
        )
        if isinstance(gpus, dict):
            gpus = [gpus]
        names = []
        if isinstance(gpus, list):
            names = [g.get("Name", "Unknown") for g in gpus]
        names = uniq_keep_order(names)
        return names or ["NE"]

    ok, out = run_cmd(["lspci"])
    if not ok:
        return ["NE"]
    items = []
    for line in out.splitlines():
        low = line.lower()
        if "vga" in low or "3d controller" in low or "display controller" in low:
            items.append(line.strip())
    return items or ["NE"]


def get_network_info():
    """Collect Wi-Fi and Ethernet adapter names."""
    if platform.system() == "Windows":
        nics = run_ps_json(
            "Get-CimInstance Win32_NetworkAdapter | "
            "Where-Object { $_.PhysicalAdapter -eq $true } | "
            "Select-Object Name,AdapterType | ConvertTo-Json -Compress"
        )
        if isinstance(nics, dict):
            nics = [nics]

        wifi = []
        ethernet = []
        if isinstance(nics, list):
            for nic in nics:
                name = (nic.get("Name") or "").strip()
                if not name:
                    continue
                n = name.lower()
                adapter_type = (nic.get("AdapterType") or "").lower()
                if "bluetooth" in n:
                    continue
                if "wireless" in adapter_type or "wi-fi" in n or "wifi" in n or "802.11" in n:
                    wifi.append(name)
                else:
                    ethernet.append(name)
        return {"wifi": uniq_keep_order(wifi), "ethernet": uniq_keep_order(ethernet)}

    return {"wifi": [], "ethernet": []}


def get_battery_info():
    """Collect battery percentage and full/design capacity when available."""
    if platform.system() == "Windows":
        batt = run_ps_json(
            "Get-CimInstance Win32_Battery | "
            "Select-Object Name,EstimatedChargeRemaining,BatteryStatus | ConvertTo-Json -Compress"
        )
        if isinstance(batt, dict):
            batt = [batt]
        if not isinstance(batt, list) or not batt:
            return {"present": False, "percentage": "NE", "full_wh": None, "design_wh": None}

        pct_values = [b.get("EstimatedChargeRemaining") for b in batt if b.get("EstimatedChargeRemaining") is not None]
        pct = int(round(sum(pct_values) / len(pct_values))) if pct_values else None

        full_caps = run_ps_json(
            "Get-CimInstance -Namespace root\\wmi -Class BatteryFullChargedCapacity | "
            "Select-Object FullChargedCapacity | ConvertTo-Json -Compress"
        )

        design_caps = run_ps_json(
            "Get-CimInstance -Namespace root\\wmi -Class BatteryStaticData | "
            "Select-Object DesignedCapacity | ConvertTo-Json -Compress"
        )
        if design_caps is None:
            design_caps = run_ps_json(
                "Get-WmiObject -Namespace root\\wmi -Class BatteryStaticData | "
                "Select-Object DesignedCapacity | ConvertTo-Json -Compress"
            )

        if isinstance(full_caps, dict):
            full_caps = [full_caps]
        if isinstance(design_caps, dict):
            design_caps = [design_caps]

        full_mwh = sum(int(x.get("FullChargedCapacity") or 0) for x in (full_caps or []))
        design_mwh = sum(int(x.get("DesignedCapacity") or 0) for x in (design_caps or []))

        full_wh = round(full_mwh / 1000, 2) if full_mwh else None
        design_wh = round(design_mwh / 1000, 2) if design_mwh else None

        return {
            "present": True,
            "percentage": f"{pct}%" if pct is not None else "NE",
            "full_wh": full_wh,
            "design_wh": design_wh,
        }

    return {"present": False, "percentage": "NE", "full_wh": None, "design_wh": None}


def _decode_wmi_ushort_string(values):
    """Decode WMI ushort array into a string (used by monitor WMI classes)."""
    if not isinstance(values, list):
        return ""
    chars = []
    for v in values:
        try:
            iv = int(v)
        except Exception:
            continue
        if iv == 0:
            break
        chars.append(chr(iv))
    return "".join(chars).strip()


def get_lcd_info():
    """Collect LCD size/resolution and append 'IPS' only if detected."""
    if platform.system() != "Windows":
        return "NE"

    res = run_ps_json(
        "Get-CimInstance Win32_VideoController | "
        "Where-Object { $_.CurrentHorizontalResolution -gt 0 -and $_.CurrentVerticalResolution -gt 0 } | "
        "Select-Object -First 1 CurrentHorizontalResolution,CurrentVerticalResolution | ConvertTo-Json -Compress"
    )
    width = (res or {}).get("CurrentHorizontalResolution")
    height = (res or {}).get("CurrentVerticalResolution")

    mon = run_ps_json(
        "Get-CimInstance -Namespace root\\wmi -Class WmiMonitorBasicDisplayParams | "
        "Where-Object { $_.MaxHorizontalImageSize -gt 0 -and $_.MaxVerticalImageSize -gt 0 } | "
        "Select-Object -First 1 MaxHorizontalImageSize,MaxVerticalImageSize | ConvertTo-Json -Compress"
    )

    inch_text = ""
    if isinstance(mon, dict):
        w_cm = float(mon.get("MaxHorizontalImageSize") or 0)
        h_cm = float(mon.get("MaxVerticalImageSize") or 0)
        if w_cm > 0 and h_cm > 0:
            inches = math.sqrt(w_cm * w_cm + h_cm * h_cm) / 2.54
            inch_text = f"{inches:.1f} "

    # IPS is not always exposed in a standard way. We check monitor model strings;
    # add "IPS" only when we find a reliable hint.
    ips_detected = False

    wmi_id = run_ps_json(
        "Get-CimInstance -Namespace root\\wmi -Class WmiMonitorID | "
        "Select-Object UserFriendlyName,ManufacturerName,ProductCodeID | ConvertTo-Json -Compress"
    )
    if isinstance(wmi_id, dict):
        wmi_id = [wmi_id]
    if isinstance(wmi_id, list):
        for m in wmi_id:
            combined = " ".join(
                [
                    _decode_wmi_ushort_string(m.get("UserFriendlyName", [])),
                    _decode_wmi_ushort_string(m.get("ManufacturerName", [])),
                    _decode_wmi_ushort_string(m.get("ProductCodeID", [])),
                ]
            ).lower()
            if "ips" in combined:
                ips_detected = True
                break

    if not ips_detected:
        mon_pnp = run_ps_json(
            "Get-CimInstance Win32_PnPEntity | "
            "Where-Object { $_.PNPClass -eq 'Monitor' } | "
            "Select-Object Name | ConvertTo-Json -Compress"
        )
        if isinstance(mon_pnp, dict):
            mon_pnp = [mon_pnp]
        if isinstance(mon_pnp, list):
            for item in mon_pnp:
                name = (item.get("Name") or "").lower()
                if "ips" in name:
                    ips_detected = True
                    break

    if width and height:
        suffix = " IPS" if ips_detected else ""
        return f"{inch_text}{width}x{height}{suffix}".strip()
    return "NE"


def pnp_exists_by_regex(regex):
    """Return True if a PnP device name matches the provided regex."""
    cmd = (
        "Get-CimInstance Win32_PnPEntity | "
        f"Where-Object {{ $_.Name -match '{regex}' }} | "
        "Select-Object -First 1 Name | ConvertTo-Json -Compress"
    )
    return run_ps_json(cmd) is not None


def pnp_exists_by_class(pnp_class):
    """Return True if a PnP device exists in the given PNPClass."""
    cmd = (
        "Get-CimInstance Win32_PnPEntity | "
        f"Where-Object {{ $_.PNPClass -eq '{pnp_class}' }} | "
        "Select-Object -First 1 Name | ConvertTo-Json -Compress"
    )
    return run_ps_json(cmd) is not None


def get_windows_biometric_info():
    """Infer fingerprint status and face hints from biometric/camera devices."""
    if platform.system() != "Windows":
        return {"fp": "NE", "face_hint": False, "biometric_devices": []}

    bio = run_ps_json(
        "Get-CimInstance Win32_PnPEntity | "
        "Where-Object { $_.PNPClass -eq 'Biometric' } | "
        "Select-Object Name,PNPClass | ConvertTo-Json -Compress"
    )
    if isinstance(bio, dict):
        bio = [bio]
    if not isinstance(bio, list):
        bio = []

    bio_names = [str(x.get("Name") or "").strip() for x in bio if str(x.get("Name") or "").strip()]
    low_names = [n.lower() for n in bio_names]

    fp_keywords = [
        "finger",
        "fingerprint",
        "synaptics wbdi",
        "validity sensor",
        "goodix fingerprint",
        "elan wbf",
    ]
    face_keywords = [
        "hello face",
        "facial",
        "face",
        "iris",
    ]

    has_fp_strong = any(any(k in n for k in fp_keywords) for n in low_names)
    has_face_bio_hint = any(any(k in n for k in face_keywords) for n in low_names)

    # Extra face hint from camera/IR device names (common for Windows Hello face).
    cam = run_ps_json(
        "Get-CimInstance Win32_PnPEntity | "
        "Where-Object { $_.PNPClass -eq 'Camera' -or $_.Name -match 'IR|Infrared|Hello|RealSense|Depth' } | "
        "Select-Object Name | ConvertTo-Json -Compress"
    )
    if isinstance(cam, dict):
        cam = [cam]
    if not isinstance(cam, list):
        cam = []
    cam_names = [str(x.get("Name") or "").lower() for x in cam]
    has_face_cam_hint = any(
        ("ir" in n and "camera" in n)
        or "infrared" in n
        or "hello" in n
        or "realsense" in n
        or "depth" in n
        for n in cam_names
    )

    if has_fp_strong:
        fp_status = "Ano"
    elif bio_names:
        # Biometric class exists but no clear fingerprint signature: manual check.
        fp_status = "Check"
    else:
        fp_status = "NE"

    return {
        "fp": fp_status,
        "face_hint": bool(has_face_bio_hint or has_face_cam_hint),
        "biometric_devices": bio_names,
    }


def get_windows_extra_flags():
    """Build Windows flags for BT/FP/WC/B-KBD/FACE summary display."""
    if platform.system() != "Windows":
        return {"bt": "NE", "fp": "NE", "wc": "NE", "b_kbd": "NE", "face": "NE"}

    has_bt = pnp_exists_by_regex("Bluetooth")
    has_wc = pnp_exists_by_class("Camera") or pnp_exists_by_regex("Webcam|Camera")
    bio = get_windows_biometric_info()

    # Backlit keyboard is not reliably exposed via standard WMI on all vendors.
    # We mark as "Check" unless we find explicit keywords.
    has_backlit_hint = pnp_exists_by_regex("Backlit|Backlight|Illuminat|RGB Keyboard")
    b_kbd = "Ano" if has_backlit_hint else "Check"

    face_status = "Ano" if bio.get("face_hint") else "Check"

    return {
        "bt": "Ano" if has_bt else "NE",
        "fp": bio["fp"],
        "wc": "Ano" if has_wc else "NE",
        "b_kbd": b_kbd,
        "face": face_status,
    }


def build_summary(inventory):
    """Build final one-line summary fields used by the text report output."""
    model = inventory["system"]
    cpu = inventory["cpu"]
    ram = inventory["ram"]
    disks = inventory["disks"]
    gpus = inventory["gpu"]
    net = inventory["network"]
    bat = inventory["battery"]
    mech = inventory.get("mechanika", {"present": False, "items": []})
    osi = inventory["os"]
    flags = inventory["flags"]

    gpu1 = gpus[0] if gpus else "NE"
    gpu2 = gpus[1] if len(gpus) > 1 else "NE"

    disk_line = "NE"
    if disks:
        d = disks[0]
        model_text = d.get("model", "Unknown")
        # Some vendors prepend "NVMe" in the model name; remove duplicate prefix.
        model_text = re.sub(r"^\s*nvme\s+", "", str(model_text), flags=re.IGNORECASE).strip()
        size_text = d.get("size_gb", "Unknown")
        transport = (d.get("transport") or "").strip()
        suffix = ""
        if transport:
            if transport.lower() == "nvme" and "nvme" not in model_text.lower():
                suffix = " NVMe"
            elif transport.upper() not in {"SCSI", "UNKNOWN"}:
                suffix = f" {transport}"
        disk_line = f"{model_text} {size_text}{suffix}".strip()

    wifi_line = net["wifi"][0] if net.get("wifi") else "NE"
    if wifi_line != "NE":
        wifi_low = wifi_line.lower()
        has_ax = (
            " ax" in wifi_low
            or "wi-fi 6" in wifi_low
            or "wifi 6" in wifi_low
            or "wi-fi 6e" in wifi_low
            or "wifi 6e" in wifi_low
            or "wi-fi 7" in wifi_low
            or "wifi 7" in wifi_low
        )
        if has_ax and "[wifi ax]" not in wifi_low:
            wifi_line = f"{wifi_line} [Wifi AX]"

    bat_line = "NE"
    if bat.get("present"):
        if bat.get("full_wh") is not None and bat.get("design_wh") is not None and bat.get("design_wh") != 0:
            health = (bat["full_wh"] / bat["design_wh"]) * 100
            bat_line = f"{health:.2f}% [{bat['full_wh']:.2f}/{bat['design_wh']:.2f} Wh]"
        else:
            bat_line = f"{bat.get('percentage', 'NE')}"

    os_line = osi.get("pretty_name", "Unknown")
    if osi.get("display_version"):
        os_line += f" ({osi['display_version']})"

    mechanika_line = "NE"
    if mech.get("present"):
        mechanika_line = " | ".join(mech.get("items", [])) or "NE"

    cpu_model = cpu.get("model_name", "Unknown")
    cpu_base = cpu.get("base_clock_ghz", "Unknown")
    cpu_line = cpu_model
    if cpu_base != "Unknown":
        # Avoid duplicate frequency when it is already present in the model string.
        if cpu_base.lower().replace(" ", "") not in cpu_model.lower().replace(" ", ""):
            cpu_line = f"{cpu_model}  {cpu_base}"

    return {
        "model_line": f"{model.get('vendor', 'Unknown')} {model.get('model', 'Unknown')} S/N: {model.get('serial', 'Unknown')}",
        "cpu_line": cpu_line,
        "ram_line": f"{ram.get('installed_human', 'Unknown')} {ram.get('type', '').strip()}".strip(),
        "disk_line": disk_line,
        "mechanika_line": mechanika_line,
        "gpu1": gpu1,
        "gpu2": gpu2,
        "lcd_line": inventory.get("lcd", "NE"),
        "wifi_line": wifi_line,
        "bat_line": bat_line,
        "os_line": os_line,
        "flags_line": (
            ("BT:Ano " if flags.get("bt") == "Ano" else "")
            + ("FP:Ano " if flags.get("fp") == "Ano" else "")
            + ("FP:Check " if flags.get("fp") == "Check" else "")
            + ("WC:Ano " if flags.get("wc") == "Ano" else "")
            + (f"FACE:{flags.get('face')} " if flags.get("face") in {"Ano", "Check"} else "")
            + f"B-KBD:{flags['b_kbd']}"
        ),
    }


def collect_inventory():
    """Collect all hardware/OS sections into one inventory dictionary."""
    now = datetime.now()
    inventory = {
        "collected_at": now.isoformat(timespec="seconds"),
        "hostname": socket.gethostname(),
        "os": get_os_info(),
        "system": get_model_info(),
        "cpu": get_cpu_info(),
        "ram": get_ram_info(),
        "disks": get_disks_info(),
        "mechanika": get_mechanika_info(),
        "gpu": get_gpu_info(),
        "network": get_network_info(),
        "battery": get_battery_info(),
        "lcd": get_lcd_info(),
        "flags": get_windows_extra_flags(),
    }
    inventory["summary"] = build_summary(inventory)
    return inventory


def to_human_report(data):
    """Render inventory dictionary into the final human-readable text report."""
    s = data["summary"]
    header_sep = "=" * 143
    header = [
        header_sep,
        "SYSTEM MATERIAL REPORT",
        f"Generated the : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        header_sep,
        "",
    ]
    lines = [
        f"Model: {s['model_line']}",
        f"CPU: {s['cpu_line']}",
        f"RAM: {s['ram_line']}",
        f"Disk: {s['disk_line']}",
        f"Mechanika: {s['mechanika_line']}",
        f"GPU: {s['gpu1']}",
        f"GPU2: {s['gpu2']}",
        f"LCD: {s['lcd_line']}",
        f"Wifi: {s['wifi_line']}",
        f"BAT: {s['bat_line']}",
        f"OS: {s['os_line']}",
        s["flags_line"],
    ]
    return "\n".join(header + lines)


def main():
    """Entry point: collect inventory, write JSON/TXT files, print report."""
    inventory = collect_inventory()

    with open("machine_report.json", "w", encoding="utf-8") as f:
        json.dump(inventory, f, indent=2, ensure_ascii=False)

    text_report = to_human_report(inventory)
    with open("machine_report.txt", "w", encoding="utf-8") as f:
        f.write(text_report)

    print("Generated: machine_report.json")
    print("Generated: machine_report.txt")
    print()
    print(text_report)


if __name__ == "__main__":
    main()
