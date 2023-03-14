import ssl

from socks import create_connection
from socks import PROXY_TYPE_SOCKS4
from socks import PROXY_TYPE_SOCKS5
from socks import PROXY_TYPE_HTTP

from imaplib import IMAP4
from imaplib import IMAP4_PORT
from imaplib import IMAP4_SSL_PORT


class SocksIMAP4(IMAP4):
    """
    IMAP service trough SOCKS proxy. PySocks module required.
    """

    PROXY_TYPES = {
        "socks4": PROXY_TYPE_SOCKS4,
        "socks5": PROXY_TYPE_SOCKS5,
        "http": PROXY_TYPE_HTTP,
    }

    def __init__(
        self,
        host,
        port=IMAP4_PORT,
        proxy_addr=None,
        proxy_port=None,
        rdns=True,
        username=None,
        password=None,
        proxy_type="socks5",
        timeout=None,
    ):
        self.proxy_addr = proxy_addr
        self.proxy_port = proxy_port
        self.rdns = rdns
        self.username = username
        self.password = password
        self.proxy_type = SocksIMAP4.PROXY_TYPES[proxy_type.lower()]

        IMAP4.__init__(self, host, port, timeout)

    def _create_socket(self):
        return create_connection(
            (self.host, self.port),
            proxy_type=self.proxy_type,
            proxy_addr=self.proxy_addr,
            proxy_port=self.proxy_port,
            proxy_rdns=self.rdns,
            proxy_username=self.username,
            proxy_password=self.password,
        )


class SocksIMAP4SSL(SocksIMAP4):
    def __init__(
        self,
        host="",
        port=IMAP4_SSL_PORT,
        keyfile=None,
        certfile=None,
        ssl_context=None,
        proxy_addr=None,
        proxy_port=None,
        rdns=True,
        username=None,
        password=None,
        proxy_type="socks5",
        timeout=None,
    ):
        if ssl_context is not None and keyfile is not None:
            raise ValueError(
                "ssl_context and keyfile arguments are mutually " "exclusive"
            )
        if ssl_context is not None and certfile is not None:
            raise ValueError(
                "ssl_context and certfile arguments are mutually " "exclusive"
            )

        self.keyfile = keyfile
        self.certfile = certfile
        if ssl_context is None:
            ssl_context = ssl._create_stdlib_context(certfile=certfile, keyfile=keyfile)
        self.ssl_context = ssl_context

        SocksIMAP4.__init__(
            self,
            host,
            port,
            proxy_addr=proxy_addr,
            proxy_port=proxy_port,
            rdns=rdns,
            username=username,
            password=password,
            proxy_type=proxy_type,
            timeout=timeout,
        )

    def _create_socket(self):
        sock = SocksIMAP4._create_socket(self)
        server_hostname = self.host if ssl.HAS_SNI else None
        return self.ssl_context.wrap_socket(sock, server_hostname=server_hostname)

    # Looks like newer versions of Python changed open() method in IMAP4 lib.
    # Adding timeout as additional parameter should resolve issues mentioned in comments.
    # Check https://github.com/python/cpython/blob/main/Lib/imaplib.py#L202 for more details.
    def open(self, host="", port=IMAP4_PORT, timeout=None):
        SocksIMAP4.open(self, host, port, timeout)
