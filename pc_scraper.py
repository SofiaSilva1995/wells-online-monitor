# -*- coding: utf-8 -*-
"""
Scraper Perfumes & Companhia (perfumesecompanhia.pt) — marcas Pierre Fabre.

Módulo ADITIVO: não toca no fluxo Wells. Devolve linhas com a MESMA forma das
linhas Wells, mais o campo `loja` = "P&C", para alimentarem o mesmo dashboard.

A P&C é Salesforce Commerce Cloud (SFCC) e a listagem de cada marca vem
COMPLETAMENTE renderizada no HTML do servidor — não precisa de browser.
Usamos `requests` (Accept-Encoding gzip,deflate para evitar o bug do brotli).

Estrutura real do tile (confirmada no HTML 2026-05-30):
  <div class="product pc-product-tile-content" data-pid="438077">
    <div class="product-tile js-product-tile"
         data-gtm-event='{"event":"impressionView","ecommerce":{"select_item":[
            {"name":"Ultra%20S%E9rum...","id":"81060M","price":"25.95",
             "brand":"AVENE","variant":"438077","Original_price":"39.95",
             "category":"SL50"}]}}'>
      <a href="/pt/avene-ultra-serum-preenchimento/438077.html"> ...
      ... <link/botão Cart-AddProduct>  (presente = EM STOCK)
      ... "Notifiquem-me!"              (presente = RUTURA)

Detecção de rutura (confirmada pela utilizadora):
  - EM STOCK  = tile tem o ícone/botão preto "adicionar ao carrinho" (Cart-AddProduct).
  - RUTURA    = NÃO tem esse ícone (P&C mostra "Notifiquem-me!" no lugar).

A P&C NÃO vende Oral Care nem Dexeryl → essas marcas não existem aqui (não é erro).
Identificação por NOME (vem do JSON gtm). SKU também vem (campo `id`).
"""

import json
import re
import urllib.parse
from datetime import datetime, timezone
from typing import Dict, List, Optional

try:
    import requests
except Exception:  # pragma: no cover
    requests = None

# ---------------------------------------------------------------------------
# Config das marcas P&C (5 marcas — SEM Oral Care e SEM Dexeryl)
# `label` TEM de coincidir com o label Wells para emparelhar por marca.
# ---------------------------------------------------------------------------
PC_BRANDS: List[Dict] = [
    {"key": "pc_avene",         "label": "Avène",         "url": "https://www.perfumesecompanhia.pt/pt/marcas/avene/",        "cgid": "AVENE"},
    {"key": "pc_a_derma",       "label": "A-Derma",       "url": "https://www.perfumesecompanhia.pt/pt/marcas/a-derma/",      "cgid": "A-DERMA"},
    {"key": "pc_ducray",        "label": "Ducray",        "url": "https://www.perfumesecompanhia.pt/pt/marcas/ducray/",       "cgid": "DUCRAY"},
    {"key": "pc_klorane",       "label": "Klorane",       "url": "https://www.perfumesecompanhia.pt/pt/marcas/klorane/",      "cgid": "KLORANE"},
    {"key": "pc_rene_furterer", "label": "René Furterer", "url": "https://www.perfumesecompanhia.pt/pt/marcas/rene-furterer/","cgid": "RENE FURTERER"},
]

_BASE = "https://www.perfumesecompanhia.pt"
_AJAX_BASE = f"{_BASE}/on/demandware.store/Sites-PC-Site/pt_PT/Search-UpdateGrid"
_PAGE_SIZE = 24  # P&C sempre devolve 24 por página (ignoram sz != 24)

_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate",   # evitar 'br' (bug de decode)
    "Accept-Language": "pt-PT,pt;q=0.9",
}

_HEADERS_AJAX = {**_HEADERS, "X-Requested-With": "XMLHttpRequest"}

# Início de cada tile de produto na listagem
_TILE_START_RE = re.compile(
    r'class="product pc-product-tile-content"\s+data-pid="(\d+)"')
_GTM_RE = re.compile(r'data-gtm-event="([^"]+)"')
_HREF_RE = re.compile(r'href="(/pt/[^"]+\.html)"')


def _log(msg: str) -> None:
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S+00:00")
    print(f"[{ts}] [P&C] {msg}", flush=True)


def _fetch(url: str, timeout: int = 30) -> Optional[str]:
    """GET com encoding seguro. Devolve HTML ou None."""
    if requests is None:
        _log("ERRO: módulo 'requests' não disponível")
        return None
    try:
        r = requests.get(url, headers=_HEADERS, timeout=timeout)
        if r.status_code != 200:
            _log(f"HTTP {r.status_code} em {url}")
            return None
        return r.text
    except Exception as e:
        _log(f"ERRO fetch {url}: {e}")
        return None


def _decode_gtm_name(name_raw: str) -> str:
    """O nome no JSON gtm vem percent-encoded em latin-1 (ex.: 'S%E9rum')."""
    if not name_raw:
        return ""
    try:
        return urllib.parse.unquote(name_raw, encoding="latin-1")
    except Exception:
        try:
            return urllib.parse.unquote(name_raw)
        except Exception:
            return name_raw


def _split_tiles(html: str) -> List[str]:
    """Parte o HTML em blocos, um por tile de produto."""
    starts = [m.start() for m in _TILE_START_RE.finditer(html)]
    if not starts:
        return []
    starts.append(len(html))
    return [html[starts[i]:starts[i + 1]] for i in range(len(starts) - 1)]


def _parse_tile(blk: str, data_str: str, label: str) -> Optional[Dict]:
    pid_m = re.search(r'data-pid="(\d+)"', blk)
    pid = pid_m.group(1) if pid_m else ""

    name = ""
    sku = ""
    price = orig = None
    g = _GTM_RE.search(blk)
    if g:
        try:
            j = json.loads(g.group(1).replace("&quot;", '"'))
            it = j["ecommerce"]["select_item"][0]
            name = _decode_gtm_name(it.get("name", ""))
            sku = it.get("id", "") or ""
            price = it.get("price")
            orig = it.get("Original_price")
        except Exception:
            pass

    href_m = _HREF_RE.search(blk)
    url = (_BASE + href_m.group(1)) if href_m else ""

    if not name:
        # fallback: alt da imagem ou título do link
        alt = re.search(r'<img[^>]*\salt="([^"]+)"', blk)
        if alt:
            name = alt.group(1).strip()
    if not name or not url:
        return None

    # ---- Stock ----
    # "data-add-cart-url" aparece no <a> de adicionar ao carrinho — ausente em OOS.
    # NÃO usar "add-to-cart" (string presente no nome da class wrapper do tile).
    has_cart = "data-add-cart-url" in blk
    has_notify = bool(re.search(r"notifiqu|esgotad|indispon|fora de stock", blk, re.I))
    is_oos = has_notify or (not has_cart)

    # ---- Desconto ----
    desconto = "Sem desconto"
    try:
        if price is not None and orig is not None:
            p, o = float(price), float(orig)
            if o > p > 0:
                desconto = f"{round((1 - p / o) * 100)}%"
    except Exception:
        pass

    return {
        "data":          data_str,
        "loja":          "P&C",
        "marca":         label,
        "titulo":        name,
        "url":           url,
        "nome_variante": "",
        "is_oos":        is_oos,
        "ref_produto":   sku or pid,
        "desconto":      desconto,
    }


def _fetch_ajax(url: str, referer: str, timeout: int = 30) -> Optional[str]:
    """GET AJAX com header X-Requested-With."""
    if requests is None:
        return None
    try:
        r = requests.get(url, headers={**_HEADERS_AJAX, "Referer": referer}, timeout=timeout)
        if r.status_code != 200:
            _log(f"HTTP {r.status_code} em {url}")
            return None
        return r.text
    except Exception as e:
        _log(f"ERRO fetch AJAX {url}: {e}")
        return None


def scrape_pc_brand(label: str, url: str, cgid: str, max_pages: int = 30) -> List[Dict]:
    """Devolve linhas (loja='P&C') para uma marca P&C.

    P&C usa SFCC com 'Carregar mais' via AJAX (Search-UpdateGrid).
    O parâmetro start/sz na URL principal é ignorado pelo servidor — tem de
    usar-se o endpoint AJAX para obter produtos além dos primeiros 24.
    """
    data_str = datetime.now(timezone.utc).strftime("%d/%m/%Y")
    _log(f"{label}: a ler {url} (cgid={cgid})")

    rows: List[Dict] = []
    seen_url: set = set()

    # Página 0: HTML completo da página da marca
    html0 = _fetch(url)
    if html0:
        for blk in _split_tiles(html0):
            row = _parse_tile(blk, data_str, label)
            if row and row["url"] not in seen_url:
                seen_url.add(row["url"])
                rows.append(row)

    # Páginas seguintes via endpoint AJAX
    import urllib.parse as _up
    cgid_enc = _up.quote(cgid)
    for page in range(1, max_pages):
        start = page * _PAGE_SIZE
        ajax_url = f"{_AJAX_BASE}?cgid={cgid_enc}&start={start}&sz={_PAGE_SIZE}"
        html = _fetch_ajax(ajax_url, referer=url)
        if not html:
            break
        tiles = _split_tiles(html)
        if not tiles:
            break
        new_count = 0
        for blk in tiles:
            row = _parse_tile(blk, data_str, label)
            if row and row["url"] not in seen_url:
                seen_url.add(row["url"])
                rows.append(row)
                new_count += 1
        if new_count == 0:
            break

    oos = sum(1 for r in rows if r["is_oos"])
    _log(f"{label}: {len(rows)} produtos | {oos} OOS | {len(rows) - oos} em stock")
    return rows


def run_pc_all() -> List[Dict]:
    """Percorre as 5 marcas P&C e devolve todas as linhas juntas.
    Self-contained (requests) — NÃO precisa de browser/Playwright."""
    all_rows: List[Dict] = []
    for cfg in PC_BRANDS:
        try:
            all_rows.extend(scrape_pc_brand(cfg["label"], cfg["url"], cfg["cgid"]))
        except Exception as e:
            _log(f"{cfg['label']}: FALHA geral: {e}")
    return all_rows


# ---------------------------------------------------------------------------
# Teste isolado:  python3 pc_scraper.py            (todas as marcas)
#                 python3 pc_scraper.py avene       (uma marca)
#                 python3 pc_scraper.py --file pc_avene.html  (parse de HTML local)
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import sys

    if len(sys.argv) > 2 and sys.argv[1] == "--file":
        data_str = datetime.now(timezone.utc).strftime("%d/%m/%Y")
        html = open(sys.argv[2], encoding="utf-8").read()
        out = []
        for blk in _split_tiles(html):
            r = _parse_tile(blk, data_str, "Avène")
            if r:
                out.append(r)
    else:
        arg = sys.argv[1].lower() if len(sys.argv) > 1 else None
        targets = PC_BRANDS
        if arg:
            targets = [b for b in PC_BRANDS if arg in b["key"] or arg in b["label"].lower()]
        out = []
        for cfg in targets:
            out.extend(scrape_pc_brand(cfg["label"], cfg["url"], cfg["cgid"]))

    print(f"\n===== TOTAL: {len(out)} produtos =====")
    by_brand = {}
    for r in out:
        by_brand.setdefault(r["marca"], [0, 0])
        by_brand[r["marca"]][0] += 1
        if r["is_oos"]:
            by_brand[r["marca"]][1] += 1
    for m, (tot, oos) in by_brand.items():
        print(f"  {m:16s} {tot:4d} produtos | {oos:3d} OOS")
    print("\n--- amostra (6) ---")
    for r in out[:6]:
        print(json.dumps(r, ensure_ascii=False))
    print("\n--- ruturas ---")
    for r in [x for x in out if x["is_oos"]][:15]:
        print(f"  {r['marca']:14s} | {r['titulo'][:50]} | {r['ref_produto']}")
