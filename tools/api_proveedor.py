#!/usr/bin/env python3
"""
Baja catálogo, precios, stock y tránsito desde la API del proveedor y arma el paquete
JSON que usa generar_lista_gremio.py (mismo formato que la extracción manual anterior).

Uso:  python3 -I api_proveedor.py <salida.json> [carpeta_cache_fotos]

Variables de entorno (secretos de GitHub, nunca en el código ni en documentos):
  PROV_API_URL   URL base de la API (termina en /api)
  PROV_API_KEY   clave x-api-key
Marca: solo Hikvision.
Nunca se usa el campo "price" (costo de Qumax): solo "price_no_list" (precio gremio).
"""
import base64, io, json, os, sys, time, urllib.request
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path

from PIL import Image

MARCA = "Hikvision"
BASE = os.environ["PROV_API_URL"].rstrip("/") + "/drop-shipping-api/"
KEY = os.environ["PROV_API_KEY"]
UA = {"User-Agent": "qumax-gremio/1.0"}


def get(path, page=1, limit=500, tries=4):
    url = f"{BASE}{path}?page={page}&limit={limit}"
    for i in range(tries):
        try:
            req = urllib.request.Request(url, headers={"x-api-key": KEY, "Accept": "application/json", **UA})
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            if e.code in (401, 403):
                raise SystemExit(f"API rechazó la clave o la IP ({e.code}) en {path}")
            err = e
        except Exception as e:  # red
            err = e
        time.sleep(3 * (i + 1))
    raise SystemExit(f"API sin respuesta en {path}: {err}")


def todo(path):
    out, page = [], 1
    while True:
        j = get(path, page)
        out += j.get("data", [])
        if page >= int(j.get("total_pages") or 1):
            return out
        page += 1


def foto(url, cache):
    """Miniatura 80x60 JPG en base64 (embebida: el dominio del proveedor no aparece en la lista)."""
    if not url:
        return None
    name = url.rsplit("/", 1)[-1]
    f = cache / name
    if not f.exists():
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30) as r:
                im = Image.open(io.BytesIO(r.read())).convert("RGB")
            im.thumbnail((80, 60))
            canvas = Image.new("RGB", (80, 60), "white")
            canvas.paste(im, ((80 - im.width) // 2, (60 - im.height) // 2))
            buf = io.BytesIO(); canvas.save(buf, "JPEG", quality=80)
            f.write_bytes(buf.getvalue())
        except Exception:
            return None
    return base64.b64encode(f.read_bytes()).decode()


def foto_grande(url, cache):
    """Foto grande (máx. 480x360, JPG) para ampliar en la lista. Se re-aloja en el sitio de Qumax.
    Devuelve el nombre del archivo en la caché o None."""
    if not url:
        return None
    name = "g480_" + url.rsplit("/", 1)[-1].rsplit(".", 1)[0] + ".jpg"
    f = cache / name
    if not f.exists():
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30) as r:
                im = Image.open(io.BytesIO(r.read())).convert("RGB")
            im.thumbnail((480, 360))
            buf = io.BytesIO(); im.save(buf, "JPEG", quality=78, optimize=True)
            f.write_bytes(buf.getvalue())
        except Exception:
            return None
    return name


def main():
    out = Path(sys.argv[1])
    cache = Path(sys.argv[2] if len(sys.argv) > 2 else "fotos_cache"); cache.mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    brands = todo("brands")
    bid = next(b["brand_id"] for b in brands if str(b.get("brand_desc", "")).strip().lower() == MARCA.lower())
    cats = {c["cat_id"]: c["cat_desc"] for c in todo("categories")}
    subs = {c["subcat_id"]: c["subcat_desc"] for c in todo("subcategories")}
    prods = [p for p in todo("products") if p.get("brand_id") == bid]
    ids = {p["item_id"] for p in prods}
    precios = {x["item_id"]: x for x in todo("prices") if x["item_id"] in ids}
    stock = {x["item_id"]: x for x in todo("stock") if x["item_id"] in ids}
    transito = [x for x in todo("in-transit-stock") if x.get("item_id") in ids]

    # diagnóstico temporal: qué campos de texto trae cada producto y qué tan largos son (sin contenido)
    diag = {}
    for p in prods:
        for k, v in p.items():
            if isinstance(v, str) and v.strip():
                d = diag.setdefault(k, {"con_dato": 0, "max": 0, "suma": 0})
                d["con_dato"] += 1; d["max"] = max(d["max"], len(v)); d["suma"] += len(v)
    diag = {k: {"con_dato": d["con_dato"], "max": d["max"], "prom": round(d["suma"] / d["con_dato"])} for k, d in diag.items()}
    Path("diag-campos.json").write_text(json.dumps({"productos": len(prods), "campos": diag,
        "todos_los_campos": sorted({k for p in prods for k in p})}, ensure_ascii=False, indent=1), encoding="utf-8")

    items = []
    for p in prods:
        pr = precios.get(p["item_id"])
        if not pr or "U$S" not in str(pr.get("currency_symbol", "")) and "USD" not in str(pr.get("currency_symbol", "")):
            continue
        items.append({
            "Id": str(p["item_id"]),
            "Marca": MARCA,
            "Código EAN": str(p.get("item_code") or "").strip(),
            "Artículo": p.get("item_desc") or "",
            "Detalle Categoría": p.get("description") or "",
            "Categoría": cats.get(p.get("cat_id"), ""),
            "Sub Categoría": subs.get(p.get("subcat_id"), ""),
            "Stock (*)": str((stock.get(p["item_id"]) or {}).get("stock_quantity", 0)),
            "Precio Gremio USD": str(pr["price_no_list"]),   # precio gremio (nunca "price")
            "IVA": str(pr.get("iva", 21)),
            "_foto": p.get("image_url_100") or "",
            "_foto_g": p.get("image_url_400") or p.get("image_url_250") or "",
        })

    with ThreadPoolExecutor(8) as ex:
        fotos = list(ex.map(lambda x: foto(x["_foto"], cache), items))
        grandes = list(ex.map(lambda x: foto_grande(x["_foto_g"], cache), items))
    imgs, imgs_g = {}, {}
    for x, b, g in zip(items, fotos, grandes):
        if b:
            imgs[x["Código EAN"]] = b
        if g:
            imgs_g[x["Código EAN"]] = g
        x.pop("_foto"); x.pop("_foto_g")

    # CONTROL DE PRECIOS: se publica solo el precio gremio (price_no_list, sin IVA; la lista le suma el IVA).
    # Si algún precio gremio es menor que el costo o falta, algo cambió en la API: no se publica nada.
    malos = [x for x in prods if x["item_id"] in precios
             and not (float(precios[x["item_id"]].get("price_no_list") or 0) >= float(precios[x["item_id"]].get("price") or 0) > 0)]
    if malos:
        raise SystemExit(f"CONTROL DE PRECIOS: {len(malos)} artículos con precio gremio menor al costo o sin precio. No se publica.")
    if len(items) < 800:
        raise SystemExit(f"Muy pocos artículos ({len(items)}): algo cambió en la API, no se publica")
    pkg = {"fecha": datetime.now(timezone.utc).isoformat(), "items": items,
           "transito": {"products": transito}, "imgs": imgs,
           "imgs_g": imgs_g, "cache": str(cache)}
    out.write_text(json.dumps(pkg, ensure_ascii=False), encoding="utf-8")
    print(f"OK API {len(items)} artículos | fotos {len(imgs)} | fotos grandes {len(imgs_g)} | tránsito {len(transito)} | {time.time() - t0:.0f}s")


if __name__ == "__main__":
    main()
