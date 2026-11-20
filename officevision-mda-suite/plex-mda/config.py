"""Configuration and device discovery/persistence for Plex MDA music player."""

from concurrent.futures import ThreadPoolExecutor
import json
import os
import re
import socket
from typing import List, Optional, Tuple

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("PMDA_DATA_DIR") or APP_DIR

DEFAULT_DB_PATHS = [
    "/var/lib/plexmediaserver/Library/Application Support/Plex Media Server/Plug-in Support/Databases/com.plexapp.plugins.library.db",
    os.path.expanduser(
        "~/Library/Application Support/Plex Media Server/Plug-in Support/Databases/com.plexapp.plugins.library.db"
    ),
    os.path.expanduser(
        "~/.local/share/plexmediaserver/Library/Application Support/Plex Media Server/Plug-in Support/Databases/com.plexapp.plugins.library.db"
    ),
]
DB_PATH = DEFAULT_DB_PATHS[0]

SETTINGS_FILE = os.path.join(DATA_DIR, "settings.json")
LEGACY_CONFIG_FILE = os.path.expanduser("~/.config/plex-mda/config.json")
DEVICES_FILE = os.path.join(DATA_DIR, "devices.json")

# Discovered playback devices are saved to devices.json via `pmda scan` or pressing [r] in [d] menu
DEFAULT_DEVICES: List[dict] = []

DEFAULT_STREAM_PORT = 8888


def get_db_path() -> str:
    """Resolve the Plex library SQLite database path from env, settings, or standard locations."""
    env_path = os.environ.get("PLEX_DB_PATH", "").strip()
    if env_path:
        return os.path.expanduser(env_path)
    settings = load_settings()
    cfg_path = str(settings.get("plex_db_path", "")).strip()
    if cfg_path:
        return os.path.expanduser(cfg_path)
    for candidate in DEFAULT_DB_PATHS:
        if os.path.exists(candidate):
            return candidate
    return DEFAULT_DB_PATHS[0]


def get_stream_port() -> int:
    """Resolve the local HTTP streaming server port."""
    env_port = os.environ.get("PMDA_STREAM_PORT", "").strip()
    if env_port.isdigit():
        return int(env_port)
    settings = load_settings()
    try:
        return int(settings.get("stream_port", DEFAULT_STREAM_PORT))
    except (TypeError, ValueError):
        return DEFAULT_STREAM_PORT


def get_local_ip() -> str:
    """Detect local IP address on LAN."""
    for target in (("10.255.255.255", 1), ("8.8.8.8", 80), ("192.168.1.1", 80)):
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(target)
            ip = s.getsockname()[0]
            if ip and not ip.startswith("127."):
                return ip
        except Exception:
            pass
        finally:
            s.close()
    try:
        return socket.gethostbyname(socket.gethostname())
    except Exception:
        return "127.0.0.1"


def load_settings() -> dict:
    """Load user preferences."""
    defaults = {
        "selected_device_id": "",
        "plex_db_path": "",
        "stream_port": DEFAULT_STREAM_PORT,
        "border_style": "ascii",  # "ascii" or "cp437"
        "volume_step": 5,
        "seek_step": 10,
    }
    for path in (SETTINGS_FILE, LEGACY_CONFIG_FILE):
        if os.path.exists(path):
            try:
                with open(path, "r") as f:
                    data = json.load(f)
                    defaults.update(data)
                break
            except Exception:
                pass
    return defaults


def save_settings(settings: dict):
    """Save user preferences."""
    for path in (SETTINGS_FILE, LEGACY_CONFIG_FILE):
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w") as f:
                json.dump(settings, f, indent=2)
        except Exception:
            pass


def load_devices() -> List[dict]:
    """Load the cached list of last discovered playback devices."""
    if os.path.exists(DEVICES_FILE):
        try:
            with open(DEVICES_FILE, "r") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception:
            pass
    return [dict(d) for d in DEFAULT_DEVICES]


def save_devices(devices: List[dict]):
    """Persist the discovered playback devices list to disk."""
    try:
        os.makedirs(os.path.dirname(DEVICES_FILE), exist_ok=True)
        with open(DEVICES_FILE, "w") as f:
            json.dump(devices, f, indent=2)
    except Exception:
        pass


def _slugify(name: str, existing_ids: set) -> str:
    base = re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_") or "device"
    slug = base
    counter = 2
    while slug in existing_ids:
        slug = f"{base}_{counter}"
        counter += 1
    return slug


def _check_tcp_port(ip: str, port: int, timeout: float = 0.35) -> bool:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(timeout)
    try:
        s.connect((ip, port))
        return True
    except Exception:
        return False
    finally:
        s.close()


def _get_arp_ips(subnet_prefix: str) -> List[str]:
    """Read known neighbor IPs on the subnet from /proc/net/arp."""
    arp_ips: List[str] = []
    try:
        with open("/proc/net/arp", "r") as f:
            for line in f.readlines()[1:]:
                parts = line.split()
                if len(parts) >= 4 and parts[0].startswith(subnet_prefix + ".") and parts[2] != "0x0":
                    arp_ips.append(parts[0])
    except Exception:
        pass
    return arp_ips


def _probe_subnet_hosts(subnet_prefix: str) -> Tuple[List[str], List[str]]:
    """Fast parallel TCP check for Cast (8008/8443/8009) and HEOS (1255) on the /24 subnet."""
    eureka_hosts: List[str] = []
    heos_hosts: List[str] = []

    def check_ip(ip: str, timeout: float = 0.5):
        # Check Cast Eureka ports (8008 / 8443) and Cast TLS port (8009)
        if _check_tcp_port(ip, 8009, timeout):
            if _check_tcp_port(ip, 8008, timeout) or _check_tcp_port(ip, 8443, timeout):
                if ip not in eureka_hosts:
                    eureka_hosts.append(ip)
        # Check HEOS port 1255
        if _check_tcp_port(ip, 1255, 0.35):
            if ip not in heos_hosts:
                heos_hosts.append(ip)

    # 1. Check known ARP neighbors first with a generous timeout (wakes sleeping Wi-Fi devices without ARP flood)
    arp_ips = _get_arp_ips(subnet_prefix)
    if arp_ips:
        with ThreadPoolExecutor(max_workers=16) as pool:
            list(pool.map(lambda ip: check_ip(ip, 0.8), arp_ips))

    # 2. Sweep remaining subnet IPs with moderate concurrency to avoid kernel ARP queue overflow
    arp_set = set(arp_ips)
    remaining_ips = [f"{subnet_prefix}.{i}" for i in range(1, 255) if f"{subnet_prefix}.{i}" not in arp_set]
    with ThreadPoolExecutor(max_workers=24) as pool:
        list(pool.map(lambda ip: check_ip(ip, 0.25), remaining_ips))

    return eureka_hosts, heos_hosts


def _probe_heos_host(ip: str, port: int = 1255) -> List[dict]:
    """Query a HEOS device on port 1255 for its players list."""
    found = []
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.settimeout(2.0)
    try:
        s.connect((ip, port))
        s.sendall(b"heos://player/get_players\r\n")
        data = s.recv(4096)
        if data:
            for line in data.decode("utf-8", errors="replace").split("\r\n"):
                if line.strip():
                    try:
                        resp = json.loads(line)
                        for p in resp.get("payload", []):
                            found.append(
                                {
                                    "name": p.get("name") or f"HEOS ({ip})",
                                    "model": p.get("model") or "HEOS Receiver",
                                    "host": p.get("ip") or ip,
                                    "port": port,
                                    "pid": p.get("pid"),
                                }
                            )
                    except Exception:
                        pass
    except Exception:
        pass
    finally:
        s.close()
    return found


def discover_devices(
    existing_devices: Optional[List[dict]] = None, timeout: float = 4.5
) -> Tuple[List[dict], List[str]]:
    """
    Scan the local network for Google Cast (including Chromecast Audio) and HEOS devices.
    Updates existing devices in place, appends any newly discovered devices, preserves
    offline known devices so powered-off TVs/receivers aren't lost, and saves to devices.json.

    Returns (updated_devices, newly_added_names).
    """
    import pychromecast

    current = [dict(d) for d in (existing_devices if existing_devices is not None else load_devices())]
    existing_ids = {d["id"] for d in current if "id" in d}
    new_names: List[str] = []

    local_ip = get_local_ip()
    parts = local_ip.split(".")
    subnet_prefix = ".".join(parts[:3]) if len(parts) == 4 else "192.168.1"

    # Fast subnet probe (~0.4s) to find all online Cast Eureka hosts & HEOS hosts even if mDNS is filtered
    eureka_hosts, heos_hosts = _probe_subnet_hosts(subnet_prefix)

    # Discover all Cast devices via mDNS + parallel direct Eureka info check (no 30s multizone hang)
    cast_infos = []
    browser = None
    zc = None
    try:
        import zeroconf

        zc = zeroconf.Zeroconf(interfaces=[local_ip])
    except Exception:
        zc = None
    try:
        cast_infos, browser = pychromecast.discover_chromecasts(
            timeout=min(3.5, timeout),
            zeroconf_instance=zc,
        )
    except Exception:
        pass
    finally:
        if browser is not None:
            try:
                pychromecast.stop_discovery(browser)
            except Exception:
                pass
        elif zc is not None:
            try:
                zc.close()
            except Exception:
                pass

    # Supplement with direct Eureka device_info queries for any eureka_hosts not already in cast_infos
    found_hosts = {ci.host for ci in cast_infos}
    missing_eureka = [ip for ip in eureka_hosts if ip not in found_hosts]
    if missing_eureka:
        def _query_eureka(ip: str):
            try:
                ds = pychromecast.dial.get_device_info(ip, timeout=1.5)
                if ds and ds.friendly_name:
                    return pychromecast.CastInfo(
                        services={pychromecast.HostServiceInfo(ip, 8009)},
                        uuid=ds.uuid,
                        model_name=ds.model_name,
                        friendly_name=ds.friendly_name,
                        host=ip,
                        port=8009,
                        cast_type=ds.cast_type,
                        manufacturer=ds.manufacturer,
                    )
            except Exception:
                pass
            return None

        with ThreadPoolExecutor(max_workers=16) as pool:
            for extra_ci in pool.map(_query_eureka, missing_eureka):
                if extra_ci is not None:
                    cast_infos.append(extra_ci)

    # Process discovered Cast devices
    for ci in cast_infos:
        friendly_name = (ci.friendly_name or ci.model_name or ci.host or "").strip()
        if not friendly_name:
            continue
        model_name = (ci.model_name or "Chromecast").strip()
        uuid_str = str(ci.uuid) if ci.uuid else ""
        host = ci.host
        port = ci.port or 8009
        cast_type = ci.cast_type or (
            "group"
            if port != 8009
            else ("audio" if "audio" in model_name.lower() else "cast")
        )

        # Match against existing devices by UUID, or (host + port), or case-insensitive name
        match = None
        if uuid_str:
            match = next((d for d in current if d.get("uuid") == uuid_str), None)
        if not match and port == 8009:
            match = next(
                (
                    d
                    for d in current
                    if d.get("type") == "cast"
                    and d.get("host") == host
                    and (d.get("port", 8009) == port)
                ),
                None,
            )
        if not match:
            match = next(
                (
                    d
                    for d in current
                    if d.get("type") == "cast"
                    and d.get("name", "").lower() == friendly_name.lower()
                ),
                None,
            )

        if match:
            match["name"] = friendly_name
            match["host"] = host
            match["port"] = port
            match["cast_type"] = cast_type
            if uuid_str:
                match["uuid"] = uuid_str
            if model_name:
                match["model"] = model_name
            if not match.get("alias"):
                match["alias"] = model_name
        else:
            dev_id = _slugify(friendly_name, existing_ids)
            existing_ids.add(dev_id)
            new_dev = {
                "id": dev_id,
                "name": friendly_name,
                "alias": model_name,
                "model": model_name,
                "type": "cast",
                "cast_type": cast_type,
                "host": host,
                "port": port,
                "uuid": uuid_str,
            }
            current.append(new_dev)
            new_names.append(friendly_name)

    # Process discovered HEOS devices
    for h_ip in heos_hosts:
        players = _probe_heos_host(h_ip, 1255)
        for p in players:
            p_name = p["name"]
            p_host = p["host"]
            p_pid = p["pid"]
            match = next(
                (
                    d
                    for d in current
                    if d.get("type") == "heos"
                    and (
                        d.get("pid") == p_pid
                        or d.get("host") == p_host
                        or d.get("name", "").lower() == p_name.lower()
                    )
                ),
                None,
            )
            if match:
                match["name"] = p_name
                match["host"] = p_host
                match["port"] = 1255
                if p_pid is not None:
                    match["pid"] = p_pid
            else:
                dev_id = _slugify(p_name, existing_ids)
                existing_ids.add(dev_id)
                current.append(
                    {
                        "id": dev_id,
                        "name": p_name,
                        "alias": p["model"],
                        "type": "heos",
                        "host": p_host,
                        "port": 1255,
                        "pid": p_pid,
                    }
                )
                new_names.append(p_name)

    save_devices(current)
    return current, new_names
