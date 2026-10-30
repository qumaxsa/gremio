#!/usr/bin/env python3
"""
Cifrado de la LISTA GREMIO para publicarla en un sitio estatico (GitHub Pages) con una clave por tecnico.

Esquema (compatible con WebCrypto, ver plantilla_gremio_privado.html):
  - Cada tecnico tiene su propio par de claves ECDH P-256.
      keys.json guarda por tecnico: clave PUBLICA + clave PRIVADA cifrada con su contrasena
      (AES-GCM, clave derivada por PBKDF2-SHA256 600.000 iteraciones). Nada de esto es secreto.
  - Cada semana los datos se cifran con una clave AES-256 nueva (K), y K se envuelve para cada
    tecnico con ECDH efimero contra su clave publica. Publicar NO requiere ninguna contrasena.
  - Dar de baja a un tecnico = sacarlo de keys.json. Desde la siguiente publicacion ya no puede abrir la lista.
  - Usuarios identificados solo por hash SHA-256 de "qumax-gremio:" + usuario en minusculas.

Comandos:
  python3 -I gremio_crypto.py adduser <keys.json>            (stdin: usuario, contrasena)
  python3 -I gremio_crypto.py deluser <keys.json> <usuario>
  python3 -I gremio_crypto.py list    <keys.json>
  python3 -I gremio_crypto.py cifrar  <keys.json> <web_data.json> <carpeta_sitio> [lista.xlsx]
  python3 -I gremio_crypto.py probar  <carpeta_sitio>        (stdin: usuario, contrasena; descifra para verificar)
Las contrasenas se leen de stdin y no se escriben en ningun archivo.
"""
import base64, gzip, hashlib, json, os, sys
from pathlib import Path

from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.hkdf import HKDF
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

ITER = 600_000
INFO = b"qumax-gremio-v1"
X962, UNC = serialization.Encoding.X962, serialization.PublicFormat.UncompressedPoint
b64 = lambda b: base64.b64encode(b).decode()
ub64 = lambda s: base64.b64decode(s)


def user_id(usuario):
    return hashlib.sha256(("qumax-gremio:" + usuario.strip().lower()).encode()).hexdigest()


def _pwkey(password, salt):
    return PBKDF2HMAC(hashes.SHA256(), 32, salt, ITER).derive(password.encode())


def _kek(shared):
    return HKDF(hashes.SHA256(), 32, bytes(32), INFO).derive(shared)


def _load(path):
    p = Path(path)
    return json.loads(p.read_text()) if p.exists() else {"v": 2, "iter": ITER, "users": {}}


def cmd_adduser(path, user, pw):
    if len(pw) < 8:
        raise SystemExit("La clave debe tener al menos 8 caracteres")
    keys = _load(path)
    priv = ec.generate_private_key(ec.SECP256R1())
    der = priv.private_bytes(serialization.Encoding.DER, serialization.PrivateFormat.PKCS8,
                             serialization.NoEncryption())
    salt, iv = os.urandom(16), os.urandom(12)
    keys["users"][user_id(user)] = {
        "pub": b64(priv.public_key().public_bytes(X962, UNC)),
        "salt": b64(salt), "iv": b64(iv),
        "wk": b64(AESGCM(_pwkey(pw, salt)).encrypt(iv, der, None)),
    }
    Path(path).write_text(json.dumps(keys, indent=1))
    print(f"OK usuario agregado. Total usuarios: {len(keys['users'])}")


def cmd_deluser(path, user):
    keys = _load(path)
    if keys["users"].pop(user_id(user), None) is None:
        raise SystemExit("No existe ese usuario")
    Path(path).write_text(json.dumps(keys, indent=1))
    print(f"OK usuario eliminado. Total usuarios: {len(keys['users'])}")


def cmd_cifrar(path, web_data, outdir, xlsx=None):
    keys = _load(path)
    if not keys["users"]:
        raise SystemExit("keys.json no tiene usuarios")
    eph = ec.generate_private_key(ec.SECP256R1())
    k = AESGCM.generate_key(256)
    wraps = {}
    for uid, u in keys["users"].items():
        pub = ec.EllipticCurvePublicKey.from_encoded_point(ec.SECP256R1(), ub64(u["pub"]))
        ivk = os.urandom(12)
        wraps[uid] = {"ivk": b64(ivk), "wk": b64(AESGCM(_kek(eph.exchange(ec.ECDH(), pub))).encrypt(ivk, k, None))}
    raw = Path(web_data).read_bytes()
    ivd = os.urandom(12)
    out = Path(outdir); out.mkdir(parents=True, exist_ok=True)
    (out / "data.bin").write_bytes(AESGCM(k).encrypt(ivd, gzip.compress(raw, 9), None))
    meta = {"v": 2, "fecha": json.loads(raw)["fecha"], "eph": b64(eph.public_key().public_bytes(X962, UNC)),
            "ivd": b64(ivd), "wraps": wraps}
    if xlsx:
        ivx = os.urandom(12)
        (out / "lista.bin").write_bytes(AESGCM(k).encrypt(ivx, Path(xlsx).read_bytes(), None))
        meta["ivx"], meta["xname"] = b64(ivx), Path(xlsx).name
    (out / "data.json").write_text(json.dumps(meta))
    print(f"OK cifrado para {len(wraps)} usuarios | data.bin {(out / 'data.bin').stat().st_size/1e6:.2f} MB | fecha {meta['fecha']}")


def cmd_probar(site, user, pw):
    """Descifra como lo haria el navegador, para verificar una publicacion."""
    keys = json.loads((Path(site) / "keys.json").read_text())
    meta = json.loads((Path(site) / "data.json").read_text())
    uid = user_id(user)
    u = keys["users"][uid]
    der = AESGCM(_pwkey(pw, ub64(u["salt"]))).decrypt(ub64(u["iv"]), ub64(u["wk"]), None)
    priv = serialization.load_der_private_key(der, None)
    eph = ec.EllipticCurvePublicKey.from_encoded_point(ec.SECP256R1(), ub64(meta["eph"]))
    w = meta["wraps"][uid]
    k = AESGCM(_kek(priv.exchange(ec.ECDH(), eph))).decrypt(ub64(w["ivk"]), ub64(w["wk"]), None)
    data = json.loads(gzip.decompress(AESGCM(k).decrypt(ub64(meta["ivd"]), (Path(site) / "data.bin").read_bytes(), None)))
    print(f"OK descifrado: {len(data['items'])} artículos, fecha {data['fecha']}")


if __name__ == "__main__":
    cmd, args = sys.argv[1], sys.argv[2:]
    if cmd in ("adduser", "probar"):
        user, pw = sys.stdin.read().splitlines()[:2]
        (cmd_adduser if cmd == "adduser" else cmd_probar)(args[0], user, pw)
    elif cmd == "deluser":
        cmd_deluser(args[0], args[1])
    elif cmd == "list":
        print(f"{len(_load(args[0])['users'])} usuarios")
    elif cmd == "cifrar":
        cmd_cifrar(*args)
    else:
        raise SystemExit(__doc__)
