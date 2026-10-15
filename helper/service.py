"""Punto de entrada del helper instalado (ver installer/). En desarrollo se usa app.py.

Sin argumentos: registra el complemento en Word para el usuario actual y sirve la API
y el panel por HTTPS en https://localhost:8766.
  --generate-cert  crea la CA local y el certificado de localhost (lo corre el instalador)
  --unregister     quita el complemento de Word del usuario actual (lo corre el desinstalador)
"""

import argparse
import datetime
import ipaddress
import os
import socket
import sys
import winreg
from pathlib import Path

PORT = 8766
BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
MANIFEST = BASE_DIR / "web" / "manifest.xml"
ADDIN_ID = "361ff4c7-d62a-48e9-9dc1-cb225cf6a1c9"  # <Id> del manifest
# Misma clave que usa `npm start` (office-addin-dev-settings) para cargar el complemento.
WEF_DEVELOPER_KEY = r"Software\Microsoft\Office\16.0\WEF\Developer"
CERT_DIR = Path(os.environ.get("PROGRAMDATA", r"C:\ProgramData")) / "TTS Word"
LOG_FILE = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "TTS Word" / "helper.log"
CA_NAME = "TTS Word - CA local"  # el instalador la busca por este nombre para quitarla


def generate_cert() -> None:
    """Crea una CA propia de esta máquina y con ella firma el certificado de localhost.

    La clave privada de la CA no se guarda: nadie más puede emitir certificados con ella.
    """
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import ec
    from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID

    now = datetime.datetime.now(datetime.timezone.utc)
    not_after = now + datetime.timedelta(days=3650)

    ca_key = ec.generate_private_key(ec.SECP256R1())
    ca_name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, CA_NAME)])
    ca_cert = (
        x509.CertificateBuilder()
        .subject_name(ca_name)
        .issuer_name(ca_name)
        .public_key(ca_key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - datetime.timedelta(days=1))
        .not_valid_after(not_after)
        .add_extension(x509.BasicConstraints(ca=True, path_length=0), critical=True)
        .add_extension(
            x509.KeyUsage(False, False, False, False, False, True, True, False, False), critical=True
        )
        .add_extension(x509.SubjectKeyIdentifier.from_public_key(ca_key.public_key()), critical=False)
        .sign(ca_key, hashes.SHA256())
    )

    key = ec.generate_private_key(ec.SECP256R1())
    cert = (
        x509.CertificateBuilder()
        .subject_name(x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "localhost")]))
        .issuer_name(ca_name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - datetime.timedelta(days=1))
        .not_valid_after(not_after)
        .add_extension(
            x509.SubjectAlternativeName(
                [x509.DNSName("localhost"), x509.IPAddress(ipaddress.ip_address("127.0.0.1"))]
            ),
            critical=False,
        )
        .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
        .add_extension(
            x509.AuthorityKeyIdentifier.from_issuer_public_key(ca_key.public_key()), critical=False
        )
        .sign(ca_key, hashes.SHA256())
    )

    CERT_DIR.mkdir(parents=True, exist_ok=True)
    (CERT_DIR / "ca.cer").write_bytes(ca_cert.public_bytes(serialization.Encoding.DER))
    (CERT_DIR / "cert.pem").write_bytes(cert.public_bytes(serialization.Encoding.PEM))
    (CERT_DIR / "key.pem").write_bytes(
        key.private_bytes(
            serialization.Encoding.PEM,
            serialization.PrivateFormat.PKCS8,
            serialization.NoEncryption(),
        )
    )


def register() -> None:
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, WEF_DEVELOPER_KEY) as key:
        winreg.SetValueEx(key, ADDIN_ID, 0, winreg.REG_SZ, str(MANIFEST))


def unregister() -> None:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, WEF_DEVELOPER_KEY, 0, winreg.KEY_SET_VALUE) as key:
            winreg.DeleteValue(key, ADDIN_ID)
    except FileNotFoundError:
        pass


def already_running() -> bool:
    with socket.socket() as sock:
        sock.settimeout(0.5)
        return sock.connect_ex(("127.0.0.1", PORT)) == 0


def serve() -> None:
    if sys.stdout is None:  # ejecutable sin consola: uvicorn necesita dónde escribir
        LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        sys.stdout = sys.stderr = open(LOG_FILE, "w", encoding="utf-8", buffering=1)
    register()  # por usuario: así también queda registrado para otras cuentas de la PC
    if already_running():
        return

    import uvicorn

    from app import app

    uvicorn.run(
        app,
        host="127.0.0.1",
        port=PORT,
        ssl_certfile=CERT_DIR / "cert.pem",
        ssl_keyfile=CERT_DIR / "key.pem",
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--generate-cert", action="store_true")
    parser.add_argument("--unregister", action="store_true")
    args = parser.parse_args()
    if args.generate_cert:
        generate_cert()
    elif args.unregister:
        unregister()
    else:
        serve()


if __name__ == "__main__":
    main()
