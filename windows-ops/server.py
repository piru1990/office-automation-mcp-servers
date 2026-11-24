"""MCP: diagnóstico de Windows (solo lectura)."""
import platform, socket, subprocess
import psutil
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

mcp = FastMCP("windows-ops")
RO = ToolAnnotations(readOnlyHint=True)


@mcp.tool(annotations=RO)
def system_info() -> dict:
    """CPU, RAM, uptime y SO."""
    import time
    vm = psutil.virtual_memory()
    return {
        "host": socket.gethostname(), "os": platform.platform(),
        "cpu_percent": psutil.cpu_percent(interval=0.5), "cpu_count": psutil.cpu_count(),
        "ram_total_gb": round(vm.total / 2**30, 1), "ram_used_percent": vm.percent,
        "uptime_hours": round((time.time() - psutil.boot_time()) / 3600, 1),
    }


@mcp.tool(annotations=RO)
def disk_usage() -> list[dict]:
    """Uso de cada disco."""
    out = []
    for p in psutil.disk_partitions():
        try:
            u = psutil.disk_usage(p.mountpoint)
            out.append({"mount": p.mountpoint, "fs": p.fstype, "total_gb": round(u.total / 2**30, 1),
                        "free_gb": round(u.free / 2**30, 1), "used_percent": u.percent})
        except OSError:
            pass
    return out


@mcp.tool(annotations=RO)
def top_processes(n: int = 15, by: str = "memory") -> list[dict]:
    """Top N procesos por 'memory' o 'cpu'."""
    procs = []
    for p in psutil.process_iter(["pid", "name", "memory_info", "cpu_percent"]):
        i = p.info
        procs.append({"pid": i["pid"], "name": i["name"],
                      "mem_mb": round(i["memory_info"].rss / 2**20, 1) if i["memory_info"] else 0,
                      "cpu": i["cpu_percent"]})
    procs.sort(key=lambda x: x["cpu" if by == "cpu" else "mem_mb"], reverse=True)
    return procs[:n]


@mcp.tool(annotations=RO)
def list_services(state: str = "all", name_filter: str = "") -> list[dict]:
    """Servicios de Windows. state: all|running|stopped."""
    out = []
    for s in psutil.win_service_iter():
        d = s.as_dict()
        if state != "all" and d["status"] != state:
            continue
        if name_filter.lower() in (d["name"] + d["display_name"]).lower():
            out.append({"name": d["name"], "display": d["display_name"], "status": d["status"], "start": d["start_type"]})
    return out[:300]


@mcp.tool(annotations=RO)
def ping(host: str, count: int = 4) -> str:
    """Ping a un host."""
    if not all(c.isalnum() or c in ".-_" for c in host):
        raise ValueError("host inválido")
    r = subprocess.run(["ping", "-n", str(min(count, 10)), host], capture_output=True, text=True, timeout=60)
    return r.stdout or r.stderr


@mcp.tool(annotations=RO)
def recent_events(log: str = "System", level: int = 2, max_events: int = 20) -> str:
    """Eventos recientes del visor de eventos. level: 1=crítico,2=error,3=advertencia."""
    if log not in ("System", "Application", "Security"):
        raise ValueError("log inválido")
    ps = (f"Get-WinEvent -FilterHashtable @{{LogName='{log}';Level={int(level)}}} -MaxEvents {int(max_events)} "
          "| Select TimeCreated,ProviderName,Id,@{n='Msg';e={$_.Message.Substring(0,[Math]::Min(200,$_.Message.Length))}} | Format-List | Out-String")
    r = subprocess.run(["pwsh", "-NoProfile", "-Command", ps], capture_output=True, text=True, timeout=60)
    return r.stdout or r.stderr


def _host(host: str) -> str:
    if not host or not all(c.isalnum() or c in ".-_:" for c in host):
        raise ValueError("host inválido")
    return host


@mcp.tool(annotations=RO)
def nslookup(host: str, server: str = "") -> str:
    """Resolución DNS de un nombre o IP (opcionalmente contra un servidor DNS concreto)."""
    cmd = ["nslookup", _host(host)] + ([_host(server)] if server else [])
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    return (r.stdout + r.stderr).strip()


@mcp.tool(annotations=RO)
def test_port(host: str, port: int, timeout: float = 4.0) -> dict:
    """Prueba si un puerto TCP responde (p. ej. 3389 RDP, 445 SMB, 1433 SQL, 30015 SAP HANA)."""
    import time
    _host(host)
    if not 1 <= int(port) <= 65535:
        raise ValueError("puerto inválido")
    t0 = time.time()
    try:
        ip = socket.gethostbyname(host)
    except OSError as e:
        return {"host": host, "port": port, "open": False, "error": f"DNS: {e}"}
    try:
        with socket.create_connection((ip, int(port)), timeout=min(timeout, 15)):
            return {"host": host, "ip": ip, "port": port, "open": True, "ms": round((time.time() - t0) * 1000)}
    except OSError as e:
        return {"host": host, "ip": ip, "port": port, "open": False, "error": str(e)}


@mcp.tool(annotations=RO)
def tracert(host: str, max_hops: int = 15) -> str:
    """Traza la ruta a un host (sin resolver nombres, para que sea rápido)."""
    r = subprocess.run(["tracert", "-d", "-w", "800", "-h", str(max(1, min(max_hops, 30))), _host(host)],
                       capture_output=True, text=True, timeout=120)
    return r.stdout or r.stderr


@mcp.tool(annotations=RO)
def ip_config() -> str:
    """Adaptadores de red con IP, puerta de enlace y DNS (ipconfig /all resumido)."""
    out = []
    for nombre, dirs in psutil.net_if_addrs().items():
        st = psutil.net_if_stats().get(nombre)
        if not st or not st.isup:
            continue
        ips = [f"{d.address}/{d.netmask}" for d in dirs if d.family == socket.AF_INET]
        if ips:
            out.append(f"{nombre}: {', '.join(ips)} ({st.speed} Mbps)")
    r = subprocess.run(["pwsh", "-NoProfile", "-Command",
                        "Get-NetIPConfiguration | ForEach-Object { '{0}: gw {1} | dns {2}' -f $_.InterfaceAlias, "
                        "($_.IPv4DefaultGateway.NextHop -join ','), ($_.DNSServer.ServerAddresses -join ',') }"],
                       capture_output=True, text=True, timeout=30, encoding="utf-8", errors="replace")
    return "\n".join(out) + "\n" + (r.stdout or r.stderr)


if __name__ == "__main__":
    mcp.run()
