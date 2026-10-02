"""
Generate a self-signed TLS certificate for the FIRS dashboard.
Run once: python generate_ssl_cert.py
The certificate is valid for 365 days and covers 10.40.24.41 and localhost.
"""
import os, datetime, ipaddress
from cryptography import x509
from cryptography.x509.oid import NameOID
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa

os.makedirs('ssl', exist_ok=True)
KEY_FILE  = 'ssl/server.key'
CERT_FILE = 'ssl/server.crt'

key = rsa.generate_private_key(public_exponent=65537, key_size=2048)

with open(KEY_FILE, 'wb') as f:
    f.write(key.private_bytes(
        serialization.Encoding.PEM,
        serialization.PrivateFormat.TraditionalOpenSSL,
        serialization.NoEncryption()
    ))

name = x509.Name([
    x509.NameAttribute(NameOID.COUNTRY_NAME,             'NG'),
    x509.NameAttribute(NameOID.STATE_OR_PROVINCE_NAME,   'Lagos'),
    x509.NameAttribute(NameOID.ORGANIZATION_NAME,        'Keystone Bank Limited'),
    x509.NameAttribute(NameOID.ORGANIZATIONAL_UNIT_NAME, 'FIRS e-Invoice'),
    x509.NameAttribute(NameOID.COMMON_NAME,              '10.40.24.41'),
])

cert = (
    x509.CertificateBuilder()
    .subject_name(name)
    .issuer_name(name)
    .public_key(key.public_key())
    .serial_number(x509.random_serial_number())
    .not_valid_before(datetime.datetime.utcnow())
    .not_valid_after(datetime.datetime.utcnow() + datetime.timedelta(days=365))
    .add_extension(x509.SubjectAlternativeName([
        x509.DNSName('localhost'),
        x509.DNSName('10.40.24.41'),
        x509.IPAddress(ipaddress.IPv4Address('10.40.24.41')),
        x509.IPAddress(ipaddress.IPv4Address('127.0.0.1')),
    ]), critical=False)
    .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
    .sign(key, hashes.SHA256())
)

with open(CERT_FILE, 'wb') as f:
    f.write(cert.public_bytes(serialization.Encoding.PEM))

print(f"Certificate : {CERT_FILE}")
print(f"Private key : {KEY_FILE}")
print(f"Expires     : {(datetime.datetime.utcnow() + datetime.timedelta(days=365)).strftime('%Y-%m-%d')}")
print(f"Covers      : 10.40.24.41, 127.0.0.1, localhost")
