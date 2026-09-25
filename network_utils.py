import errno
import socket

try:
    import socks  # PySocks, shipped with AstrBot core
except ImportError:  # Keep the plugin importable even without PySocks.
    socks = None


# Plugin proxy type name -> PySocks constant name.
_SOCKS_PROXY_TYPES = {"socks5": "SOCKS5", "socks4": "SOCKS4", "http": "HTTP"}


def parse_socks_proxy(settings: dict | None) -> dict | None:
    """Convert the plugin's `proxy_settings` config into a proxy descriptor.

    Args:
        settings: The `proxy_settings` dict from plugin config.

    Returns:
        A descriptor with type/host/port/username/password when the proxy is
        enabled, or None when it is disabled.

    Raises:
        ValueError: The proxy is enabled but host, port or type is invalid.
    """
    if not isinstance(settings, dict) or not settings.get("enable"):
        return None

    host = str(settings.get("host") or "").strip()
    if not host:
        raise ValueError("已启用代理，但未填写代理服务器地址。")
    if "://" in host:
        # PySocks expects a bare host; the scheme comes from the type dropdown.
        raise ValueError(
            f"代理服务器地址不要包含协议前缀（当前填写: {host}）。"
            "只填写 IP 或域名即可，协议类型请通过上方的“代理类型”选择。"
        )

    try:
        port = int(settings.get("port") or 0)
    except (TypeError, ValueError):
        port = 0
    if not 0 < port < 65536:
        raise ValueError("已启用代理，但代理端口配置错误。")

    proxy_type = str(settings.get("type") or "socks5").strip().lower()
    if proxy_type not in _SOCKS_PROXY_TYPES:
        raise ValueError(
            f"不支持的代理类型: {proxy_type}，可选值: {', '.join(_SOCKS_PROXY_TYPES)}"
        )

    return {
        "type": proxy_type,
        "host": host,
        "port": port,
        "username": str(settings.get("username") or "").strip() or None,
        "password": str(settings.get("password") or "").strip() or None,
    }


def open_tcp_socket(
    host: str,
    port: int,
    timeout: float | None = None,
    source_address: tuple[str, int] | None = None,
    proxy: dict | None = None,
) -> socket.socket:
    """Open a TCP socket with IPv4-first fallback for mixed-network hosts.

    Args:
        host: Target hostname or IP address.
        port: Target port.
        timeout: Connect timeout in seconds, None keeps the system default.
        source_address: Optional local source address binding (direct only).
        proxy: Proxy descriptor from parse_socks_proxy(), None means direct.

    Returns:
        A connected socket; the caller is responsible for closing it.

    Raises:
        OSError: Every candidate address failed to connect.
        RuntimeError: A proxy is requested but PySocks is unavailable.
    """
    if proxy:
        # With a proxy the domain is resolved by the proxy server, so no local
        # getaddrinfo() is needed. This also bypasses the all-IPv6 resolved
        # addresses that have no usable route.
        if socks is None:
            raise RuntimeError(
                "未安装 PySocks，无法使用代理，请先执行 pip install pysocks。"
            )
        proxy_socket = socks.socksocket()
        proxy_socket.set_proxy(
            getattr(socks, _SOCKS_PROXY_TYPES[proxy["type"]]),
            proxy["host"],
            proxy["port"],
            username=proxy.get("username"),
            password=proxy.get("password"),
        )
        if timeout is not None:
            proxy_socket.settimeout(timeout)
        try:
            proxy_socket.connect((host, port))
        except Exception:
            proxy_socket.close()
            raise
        return proxy_socket

    addrinfos = socket.getaddrinfo(host, port, type=socket.SOCK_STREAM)
    addrinfos.sort(key=lambda item: 0 if item[0] == socket.AF_INET else 1)

    errors: list[OSError] = []
    for family, socktype, proto, _, sockaddr in addrinfos:
        sock: socket.socket | None = None
        try:
            sock = socket.socket(family, socktype, proto)
            if timeout is not None:
                sock.settimeout(timeout)
            if source_address:
                sock.bind(source_address)
            sock.connect(sockaddr)
            return sock
        except OSError as exc:
            errors.append(exc)
            if sock is not None:
                sock.close()

    if not errors:
        raise OSError(f"Unable to resolve {host}:{port}")

    if any(exc.errno == errno.ENETUNREACH for exc in errors):
        raise OSError(
            errno.ENETUNREACH,
            (
                f"Unable to connect to {host}:{port}. The current environment may "
                "lack a usable route for one or more resolved addresses."
            ),
        ) from errors[-1]

    raise errors[-1]
