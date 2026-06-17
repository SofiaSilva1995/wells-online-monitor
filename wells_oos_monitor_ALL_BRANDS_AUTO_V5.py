# -*- coding: utf-8 -*-
"""
Wells - Monitor OOS (ALL BRANDS) - AUTO V4
Master unificado: Avene (155), Ducray (68), Klorane (85), Rene Furterer (72), A-Derma (60)

NOVIDADES V4 (vs V3):
- expand_variant_urls(): ao encontrar um produto com "+X Tamanho(s)" na listagem,
  entra na pagina do produto principal e recolhe os URLs de TODAS as variantes.
  Garante que todos os produtos sao detetados (ex: 151/151 Avene).
- Coluna "ref_produto": referencia interna (REF: XXXXXXX)
- Coluna "variantes_oos": variantes de tamanho indisponiveis separadas por " | "
- Email: so envia quando ha NOVOS OOS

Requisitos:
- config_email.py na mesma pasta
- playwright instalado
- openpyxl instalado
"""

import csv
import json
import math
import os
import re
import mimetypes
import smtplib
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from email.message import EmailMessage
from datetime import datetime, timezone
from typing import Optional, Set, Tuple, Dict, List
from urllib.parse import urlsplit, urlunsplit

from playwright.sync_api import sync_playwright

# Numero de workers paralelos para verificacao OOS
OOS_WORKERS = 5
_log_lock = threading.Lock()

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

import config_email

PER_PAGE = 24

OOS_KEYWORDS = (
    "indisponivel",
    "indisponível",
    "esgotado",
    "sem stock",
    "temporariamente indisponivel",
    "temporariamente indisponível",
)

BRANDS: List[Dict] = [
    {
        "key": "avene",
        "label": "Avène",
        "base_url": "https://wells.pt/avene.html",
        "max_pages_fallback": 60,
        "brand_filter": ["avene", "avène"],
        "exclude_urls": {
            "https://wells.pt/ultra-facial-meltdown-recovery-cream-446915.html",
            "https://wells.pt/ella-ella-flora-azura-eau-de-parfum-446720.html",
            "https://wells.pt/break-fix-liquid-nail-patch-8684662.html",  # OPI - intruso na pagina Avene
        }
    },
    {"key": "ducray",       "label": "Ducray",       "base_url": "https://wells.pt/marcas/ducray",       "max_pages_fallback": 10, "exclude_urls": set(), "brand_filter": ["ducray"]},
    {"key": "klorane",      "label": "Klorane",      "base_url": "https://wells.pt/marcas/k/klorane",      "max_pages_fallback": 12, "exclude_urls": set(), "brand_filter": ["klorane"]},
    {"key": "rene_furterer","label": "René Furterer","base_url": "https://wells.pt/marcas/rene-furterer","max_pages_fallback": 12, "exclude_urls": set(), "brand_filter": ["rene furterer", "rene-furterer", "rene furtere"]},
    {"key": "a_derma",      "label": "A-Derma",      "base_url": "https://wells.pt/marcas/a-derma",      "max_pages_fallback": 10, "exclude_urls": set(), "brand_filter": ["a-derma", "aderma"]},
    {
        "key": "oral_care",
        "label": "Oral Care",
        "multi_brand": True,
        "sub_brands": [
            {"name": "Eludril",    "url": "https://wells.pt/resultados-pesquisa-wells?q=ELUDRIL&prefn1=brand&prefv1=Eludril"},
            {"name": "Elgydium",   "url": "https://wells.pt/resultados-pesquisa-wells?q=ELGYDIUM&prefn1=brand&prefv1=Elgydium"},
            {"name": "Arthrodont", "url": "https://wells.pt/resultados-pesquisa-wells?q=ARTHRODONT&prefn1=brand&prefv1=Arthrodont"},
            {"name": "Elugel",     "url": "https://wells.pt/gel-redutor-placa-bacteriana-3017892.html", "single_product": True},
            {"name": "Parodium",   "url": "https://wells.pt/gel-gengival-3065953.html", "single_product": True},
        ],
        "max_pages_fallback": 10,
        "exclude_urls": set()
    },
    {
        "key": "dexeryl",
        "label": "Dexeryl",
        "base_url": "https://wells.pt/resultados-pesquisa-wells?q=DEXERYL&prefn1=brand&prefv1=Dexeryl",
        "max_pages_fallback": 10,
        "exclude_urls": set()
    },
]


# ---------------------------------------------------------------------------
# UTILS
# ---------------------------------------------------------------------------

def utc_now_ts() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S+00:00")

def log(msg: str) -> None:
    print(f"[{utc_now_ts()}] {msg}", flush=True)

def normalize_url(url: str) -> str:
    parts = list(urlsplit(url))
    parts[3] = ""
    parts[4] = ""
    return urlunsplit(parts)

def ensure_dirs(state_dir: str, logs_dir: str) -> None:
    os.makedirs(state_dir, exist_ok=True)
    os.makedirs(logs_dir, exist_ok=True)

def try_accept_cookies(page) -> None:
    """Remove banner de cookies via JS (mais fiavel que click em headless)."""
    KILL_JS = """
    () => {
        ['CybotCookiebotDialog','CybotCookiebotDialogBodyUnderlay'].forEach(function(id){
            var el = document.getElementById(id);
            if(el) el.remove();
        });
        document.querySelectorAll('[id*="Cookiebot"],[class*="cookiebot"],[id*="onetrust"],[class*="onetrust"]')
            .forEach(function(el){ el.remove(); });
        document.body.style.overflow = '';
        document.documentElement.style.overflow = '';
    }
    """
    try:
        page.evaluate(KILL_JS)
        page.wait_for_timeout(150)
        return
    except Exception:
        pass
    # Fallback: clique normal
    for sel in [
        "#CybotCookiebotDialogBodyLevelButtonLevelOptinAllowAll",
        "#onetrust-accept-btn-handler",
        "button:has-text('Aceitar tudo')",
        "button:has-text('Aceitar')",
        "button:has-text('Accept')",
    ]:
        try:
            loc = page.locator(sel).first
            if loc.is_visible(timeout=500):
                loc.click(timeout=800)
                page.wait_for_timeout(200)
                return
        except Exception:
            pass

def detect_total_results(page) -> Optional[int]:
    pats = [
        r"\b(\d{2,4})\s+Resultados\b",
        r"\b(\d{2,4})\s+resultados\b",
    ]
    def scan(t: str) -> Optional[int]:
        if not t:
            return None
        for p in pats:
            m = re.search(p, t, flags=re.IGNORECASE)
            if m:
                try:
                    v = int(m.group(1))
                    if 10 <= v <= 5000:
                        return v
                except Exception:
                    pass
        return None
    try:
        v = scan(page.inner_text("body"))
        if v:
            return v
    except Exception:
        pass
    try:
        v = scan(page.content())
        if v:
            return v
    except Exception:
        pass
    return None


# ---------------------------------------------------------------------------
# FASE 1 - RECOLHA DE URLs DA LISTAGEM
# ---------------------------------------------------------------------------

def extract_urls_from_dom(page, brand_label: str, exclude_urls: Set[str], expected_min: int = 0, brand_filter: List[str] = None) -> Set[str]:
    """
    Recolhe URLs de produto da pagina de listagem atual.
    Se brand_filter for fornecido (lista de strings), só aceita produtos cujo
    card contenha pelo menos uma dessas strings (case-insensitive, sem acentos).
    Usado como fallback quando a deteção de intrusos por pág1∩pág2 falha
    (ex: Wells repete resultados entre páginas num catálogo pequeno).
    """

    if expected_min > 0:
        try:
            page.wait_for_function(
                r"""(minCount) => {
                    const urls = Array.from(document.querySelectorAll("a[href$='.html']"))
                      .map(a => a.href)
                      .filter(u => /-\d+\.html(\?|#|$)/i.test(u));
                    return urls.length >= minCount;
                }""",
                expected_min,
                timeout=8000
            )
        except Exception:
            try:
                page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                page.wait_for_timeout(400)
            except Exception:
                pass

    EXCLUDE_PATTERNS = [
        r"/marcas/", r"/categoria/",
        r"/avene\.html$", r"/ducray\.html$", r"/klorane\.html$",
        r"/rene-furterer\.html$", r"/a-derma\.html$",
    ]

    js = r"""
    () => {
      const norm = (u) => {
        try {
          const url = new URL(u, window.location.origin);
          url.search = ''; url.hash = '';
          return url.toString();
        } catch(e) { return u || ""; }
      };
      const isProductUrl = (u) => /-\d+\.html(\?|#|$)/i.test(u);
      const anchors = Array.from(document.querySelectorAll("a[href$='.html']"));
      const out = [];
      const seen = new Set();
      for (const a of anchors) {
        const url = norm(a.href);
        if (!url || !url.includes("wells.pt/")) continue;
        if (!isProductUrl(url)) continue;
        if (seen.has(url)) continue;
        seen.add(url);
        out.push(url);
      }
      return out;
    }
    """
    try:
        hrefs = page.evaluate(js) or []
    except Exception:
        hrefs = []

    # Se brand_filter activo, recolhe também o texto de cada card para validar a marca
    card_brand_map: dict = {}  # url -> card_text (só preenchido se brand_filter activo)
    if brand_filter:
        js_cards = r"""
        () => {
          const norm = (u) => {
            try { const url = new URL(u, window.location.origin); url.search=''; url.hash=''; return url.toString(); }
            catch(e) { return u || ""; }
          };
          const isProductUrl = (u) => /-\d+\.html(\?|#|$)/i.test(u);
          const result = {};
          const anchors = Array.from(document.querySelectorAll("a[href$='.html']"));
          for (const a of anchors) {
            const url = norm(a.href);
            if (!url || !isProductUrl(url)) continue;
            // Sobe até ao card (li, article ou div com class product)
            let el = a;
            for (let i = 0; i < 6; i++) {
              if (!el.parentElement) break;
              el = el.parentElement;
              const tag = el.tagName.toLowerCase();
              const cls = (el.className || '').toLowerCase();
              if (tag === 'li' || tag === 'article' ||
                  cls.includes('product') || cls.includes('w-product')) break;
            }
            const cardText = (el.innerText || el.textContent || '').toLowerCase();
            if (!result[url]) result[url] = cardText;
          }
          return result;
        }
        """
        try:
            card_brand_map = page.evaluate(js_cards) or {}
        except Exception:
            card_brand_map = {}

    import unicodedata as _ucd
    def _strip_accents(s: str) -> str:
        return "".join(c for c in _ucd.normalize("NFD", s) if _ucd.category(c) != "Mn")

    brand_filter_norm = [_strip_accents(b.lower()) for b in brand_filter] if brand_filter else []

    out: Set[str] = set()
    for u in hrefs:
        if not u:
            continue
        u = normalize_url(str(u))
        if "wells.pt/" not in u:
            continue
        if not re.search(r"-\d+\.html$", u, flags=re.IGNORECASE):
            continue
        if u in exclude_urls:
            continue
        skip = False
        for pat in EXCLUDE_PATTERNS:
            if re.search(pat, u, re.IGNORECASE):
                skip = True
                break
        if skip:
            continue
        # Filtro de marca: se brand_filter activo, o card tem de mencionar a marca
        if brand_filter_norm:
            card_text = _strip_accents(card_brand_map.get(u, "").lower())
            if not any(bf in card_text for bf in brand_filter_norm):
                log(f"  [brand_filter] Rejeitado (sem '{brand_filter}'): {u}")
                continue
        out.add(u)

    return out

def get_cards_with_variants(page) -> List[Dict]:
    """
    Devolve lista de dicts com:
      - url: URL principal do produto
      - has_variants: True se o card tem badge '+X Tamanho(s)'
    Usado para saber quais produtos precisam de expansao de variantes.
    """
    js = r"""
    () => {
      const norm = (u) => {
        try { const url = new URL(u, window.location.origin); url.search=''; url.hash=''; return url.toString(); }
        catch(e) { return u || ""; }
      };
      const isProductUrl = (u) => /-\d+\.html(\?|#|$)/i.test(u);

      const results = [];
      const seen = new Set();

      // Cada card de produto na listagem
      const cards = Array.from(document.querySelectorAll("li,article,[class*='product-item'],[class*='w-product']"));

      for (const card of cards) {
        const anchor = card.querySelector("a[href$='.html']");
        if (!anchor) continue;

        const url = norm(anchor.href);
        if (!url || !isProductUrl(url)) continue;
        if (seen.has(url)) continue;
        seen.add(url);

        // Deteta badge "+X Tamanho(s)" ou "+X Size(s)"
        const cardText = (card.innerText || card.textContent || "").toLowerCase();
        const hasVariants = /\+\s*\d+\s*tamanho|variant|size/i.test(cardText);

        results.push({ url, has_variants: hasVariants });
      }
      return results;
    }
    """
    try:
        return page.evaluate(js) or []
    except Exception:
        return []


def expand_variant_urls(page, main_url: str, exclude_urls: Set[str]) -> Set[str]:
    """
    Entra na pagina do produto principal e clica em cada botao de variante
    (js-w-variation-tile) para capturar o URL final de cada tamanho.

    Estrutura real do Wells:
      <a class="js-w-variation-tile w-variation-tile ..."
         href="https://wells.pt/on/demandware.store/.../Product-Variation?dwvar_PID_capacity=200ml&pid=PID"
         data-attr-value="200 ml">
        <span>200 ml</span>
      </a>

    O href nao e um URL de produto direto, por isso navegamos para cada href
    e capturamos o URL final resolvido pelo browser.
    """
    found: Set[str] = {main_url}

    try:
        page.goto(main_url, wait_until="domcontentloaded")
        page.wait_for_timeout(700)
        try_accept_cookies(page)
    except Exception:
        return found

    # Recolhe label + href de cada botao de variante
    js_get_tiles = """
    () => {
      const tiles = Array.from(document.querySelectorAll(
        "a.js-w-variation-tile, a[class*='variation-tile'], a[class*='w-variation-tile']"
      ));
      return tiles.map(a => ({
        label: (a.getAttribute('data-attr-value') || a.innerText || '').trim(),
        href:  a.href || ''
      })).filter(t => t.href);
    }
    """
    try:
        tiles = page.evaluate(js_get_tiles) or []
    except Exception:
        tiles = []

    if len(tiles) <= 1:
        return found

    for tile in tiles:
        label = tile.get("label", "")
        href  = tile.get("href", "")
        if not href:
            continue

        try:
            # Navega para o href da variante — o Wells redireciona para o URL de produto correto
            page.goto(href, wait_until="domcontentloaded")
            page.wait_for_timeout(600)

            # URL final apos redirecionamento = URL de produto desta variante
            final = normalize_url(page.url)

            if (final
                    and "wells.pt/" in final
                    and re.search(r"-\d+\.html$", final)
                    and final not in exclude_urls):
                found.add(final)
                log(f"    Variante '{label}' -> {final}")

        except Exception:
            pass

    return found


# ---------------------------------------------------------------------------
# FASE 2 - VERIFICACAO OOS POR PRODUTO
# ---------------------------------------------------------------------------

def extract_ref(page, previous_ref: str = "") -> str:
    """
    Extrai a referencia interna do produto (ex: REF: 6552595).
    Se previous_ref for fornecido, aguarda ate o DOM mostrar uma REF diferente
    (util apos click AJAX numa variante).
    """
    js = """
    () => {
        const el = document.querySelector(
            'p.w-product-id, span.w-product-id, div.w-product-id, [class*="w-product-id"]'
        );
        if (!el) return '';
        return el.innerText || el.textContent || '';
    }
    """
    attempts = 5 if previous_ref else 1
    for _ in range(attempts):
        try:
            txt = (page.evaluate(js) or "").strip()
            if txt:
                m = re.search(r"REF[:\s]+(\d+)", txt, re.IGNORECASE)
                ref = m.group(1) if m else re.sub(r"[^\d]", "", txt)
                if ref and ref != previous_ref:
                    return ref
                if ref and not previous_ref:
                    return ref
        except Exception:
            pass
        if previous_ref and attempts > 1:
            try: page.wait_for_timeout(200)
            except Exception: pass
    try:
        html = page.content()
        m = re.search(r"REF[:\s]+(\d{5,10})", html, re.IGNORECASE)
        if m:
            return m.group(1)
    except Exception:
        pass
    return ""
def extract_discount(page) -> str:
    """
    Extrai o desconto do produto (ex: '25%').
    Retorna '25%', '30%', etc. ou 'Sem desconto' se nao houver.
    """
    try:
        result = page.evaluate("""
        () => {
            // Método 1: lê directamente span.w-discount (estrutura confirmada por inspect)
            // <div class="w-badge-price"><span class="w-discount">35</span><span class="w-unit">%</span>...
            const disc = document.querySelector('.w-badge-price .w-discount, .w-badge-price [class*="discount"]');
            if (disc) {
                const val = (disc.innerText || disc.textContent || '').trim();
                if (/^\\d+$/.test(val)) return val + '%';
            }
            // Método 2: lê o badge-price inteiro e extrai número antes de %
            const selectors = [
                '.js-w-circle-badge .w-badge-price',
                '.w-circle-product-badge .w-badge-price',
                '.w-circle-badge .w-badge-price',
                '.js-w-badge-product-image .w-badge-price',
                '[class*="circle-badge"] .w-badge-price'
            ];
            for (const sel of selectors) {
                const badge = document.querySelector(sel);
                if (!badge) continue;
                const full = (badge.innerText || badge.textContent || '').replace(/\\s+/g,' ').trim();
                const m = full.match(/(\\d+)\\s*%/);
                if (m) return m[1] + '%';
            }
            // Método 3: fallback PVPR — restrito à zona do produto
            const zone = document.querySelector('#product-content, .product-detail, .pdp-main, [class*="product-detail"]') || document.body;
            const txt = (zone.innerText || zone.textContent || '');
            const mp = txt.match(/(\\d+[,.]\\d+)\\s*€?\\s*(\\d+[,.]\\d+)\\s*PVPR/i);
            if (mp) {
                const preco = parseFloat(mp[1].replace(',','.'));
                const pvpr  = parseFloat(mp[2].replace(',','.'));
                if (pvpr > preco && pvpr > 0) return Math.round((pvpr-preco)/pvpr*100) + '%';
            }
            return '';
        }
        """)
        if result:
            log(f"      [DESCONTO] -> {result}")
            return result
        log(f"      [DESCONTO] sem desconto")
        return "Sem desconto"
    except Exception as e:
        log(f"      [DESCONTO] exception: {e}")
        return "Sem desconto"

def extract_titulo(page, brand_label=""):
    import unicodedata, re as _re
    def _norm(s):
        return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").lower()

    BRAND_VARIANTS = [
        "Avène", "Avene",
        "Ducray",
        "Klorane",
        "René Furterer", "Rene Furterer",
        "A-Derma", "Aderma",
        "Dexeryl",
    ]
    if brand_label:
        BRAND_VARIANTS = [brand_label] + BRAND_VARIANTS

    for sel in ["h1.w-product-name","h1[class*='product-name']","h1[class*='product-title']","h1"]:
        try:
            loc = page.locator(sel).first
            if loc.is_visible(timeout=500):
                txt = loc.inner_text(timeout=500).strip()
                if txt:
                    # Remove nome da marca do inicio (mesmo que colado sem espaco)
                    txt_norm = _norm(txt)
                    for bv in BRAND_VARIANTS:
                        bv_norm = _norm(bv)
                        if txt_norm.startswith(bv_norm):
                            # Usa comprimento do texto original (com acentos)
                            txt = txt[len(bv):].strip()
                            break
                    return txt
        except Exception:
            pass
    return ""

def extract_variants(page) -> List[Dict]:
    """
    Extrai todas as variantes de tamanho da pagina de produto.
    Tenta detectar OOS directamente do DOM sem clicar ou navegar.
    Devolve lista de dicts: [{label, href, selected, is_oos}]
    """
    js = """
    () => {
      const tiles = Array.from(document.querySelectorAll(
        "a.js-w-variation-tile, a[class*='w-variation-tile']"
      ));
      return tiles.map(a => {
        const cls = (a.className || '').toLowerCase();
        const txt = (a.innerText || '').toLowerCase();
        // Wells marca variantes OOS com classes como 'unselectable', 'disabled', 'out-of-stock'
        const isOos = cls.includes('unselectable') ||
                      cls.includes('disabled') ||
                      cls.includes('out-of-stock') ||
                      cls.includes('oos') ||
                      a.hasAttribute('disabled') ||
                      txt.includes('esgotado') ||
                      txt.includes('indispon');
        return {
          label:    (a.getAttribute('data-attr-value') || a.innerText || '').trim(),
          href:     a.href || '',
          selected: a.classList.contains('select-capacity') || a.classList.contains('selected'),
          is_oos:   isOos ? 1 : -1  // -1 = desconhecido (nao conseguiu determinar do DOM)
        };
      }).filter(t => t.label && t.href);
    }
    """
    try:
        return page.evaluate(js) or []
    except Exception:
        return []


def check_product(page, url: str, brand_label: str) -> List[Dict]:
    """
    Verifica um produto e devolve uma lista de dicts - um por variante.
    Se o produto nao tem variantes, devolve uma lista com 1 elemento.

    Estrutura de cada dict:
        data, marca, titulo, url, nome_variante, is_oos, ref_produto
    """
    final_url = url
    try:
        resp = page.goto(url, wait_until="domcontentloaded")
        if resp is not None:
            try:
                final_url = normalize_url(resp.url)
            except Exception:
                pass
    except Exception:
        pass

    try_accept_cookies(page)
    page.wait_for_timeout(600)

    # Aguarda elemento REF estar no DOM antes de extrair (até 3s)
    try:
        page.wait_for_selector(
            'p.w-product-id, span.w-product-id, div.w-product-id, [class*="w-product-id"]',
            timeout=3000
        )
    except Exception:
        pass  # Continua mesmo que não encontre — alguns produtos podem não ter REF

    data_hoje = datetime.now().strftime("%d/%m/%Y")
    titulo    = extract_titulo(page, brand_label)
    ref_base  = extract_ref(page)
    variants  = extract_variants(page)

    rows: List[Dict] = []

    if len(variants) <= 1:
        # Sem variantes - 1 linha
        # Verifica OOS restringindo à zona do produto (evita falsos positivos de produtos relacionados)
        try:
            is_oos = page.evaluate("""() => {
                // 1. Botão "Adicionar ao carrinho" desativado ou ausente → OOS
                var btn = document.querySelector(
                    'button.js-add-to-cart, button[data-action="add-to-cart"], ' +
                    'button.add-to-cart, button[name="add-to-cart"]'
                );
                if (btn) {
                    if (btn.disabled || btn.classList.contains('disabled') ||
                        btn.classList.contains('unselectable')) return 1;
                    return 0;
                }
                // 2. Fallback: keywords apenas na zona do produto
                var zone = document.querySelector(
                    '#product-content, .product-detail, .product-info, ' +
                    '.pdp-main, [class*="product-detail"], [class*="pdp"]'
                );
                var txt = (zone ? zone.innerText : '').toLowerCase();
                var kws = ['indisponivel','indisponível','esgotado','sem stock',
                           'temporariamente indisponivel','temporariamente indisponível'];
                return kws.some(function(k){ return txt.includes(k); }) ? 1 : 0;
            }""")
        except Exception:
            is_oos = 0
        # Aguarda badge de desconto carregar (importante no GitHub Actions)
        for _d in range(8):
            page.wait_for_timeout(200)
            _b = page.evaluate("""() => {
                var b = document.querySelector('.js-w-circle-badge .w-badge-price, .w-circle-product-badge .w-badge-price, .w-circle-badge .w-badge-price');
                return b ? (b.innerText || b.textContent || '').trim() : '';
            }""")
            if _b: break
        # extract_discount já tem 3 métodos: w-discount span, badge-price, fallback PVPR
        desconto = extract_discount(page)
        rows.append({
            "data":         data_hoje,
            "marca":        brand_label,
            "titulo":       titulo,
            "url":          final_url,
            "nome_variante": variants[0]["label"] if variants else "",
            "is_oos":       is_oos,
            "ref_produto":  ref_base,
            "desconto":     desconto,
        })
    else:
        # Com variantes - lê OOS directamente do DOM sem clicar ou navegar
        # O Wells marca tiles OOS com classes CSS (unselectable, disabled, etc.)
        last_ref = ref_base  # REF actual no DOM antes de cada click
        for tile in variants:
            tile_label = tile.get("label", "")
            if not tile_label:
                continue

            dom_oos = tile.get("is_oos", -1)  # -1=desconhecido, 0=OK, 1=OOS
            ref_v   = ref_base  # fallback

            # Clica na variante e aguarda AJAX actualizar REF
            try:
                tile_loc = page.locator(
                    f"a.js-w-variation-tile[data-attr-value='{tile_label}'], "
                    f"a[class*='w-variation-tile'][data-attr-value='{tile_label}']"
                ).first
                # Remove banner cookies via JS antes do click (headless-safe)
                KILL_JS = """() => {
                    ['CybotCookiebotDialog','CybotCookiebotDialogBodyUnderlay'].forEach(function(id){
                        var el=document.getElementById(id); if(el) el.remove();
                    });
                    document.querySelectorAll('[id*="Cookiebot"],[class*="cookiebot"],[id*="onetrust"],[class*="onetrust"]')
                        .forEach(function(el){ el.remove(); });
                    document.body.style.overflow='';
                }"""
                try: page.evaluate(KILL_JS)
                except Exception: pass
                # Captura badge ANTES do clique para detectar mudança
                _GET_BADGE_JS = """() => {
                    // Método 1: lê directamente span.w-discount (estrutura confirmada por inspect)
                    // <div class="w-badge-price"><span class="w-discount">35</span><span class="w-unit">%</span>...
                    var disc = document.querySelector('.w-badge-price .w-discount, .w-badge-price [class*="discount"]');
                    if (disc) {
                        var val = (disc.innerText || disc.textContent || '').trim();
                        if (/^\\d+$/.test(val)) return val + '%';
                    }
                    // Método 2: lê o badge-price inteiro e extrai número antes de %
                    var sels = [
                        '.js-w-circle-badge .w-badge-price',
                        '.w-circle-product-badge .w-badge-price',
                        '.w-circle-badge .w-badge-price',
                        '.js-w-badge-product-image .w-badge-price',
                        '[class*="circle-badge"] .w-badge-price'
                    ];
                    for (var s of sels) {
                        var b = document.querySelector(s);
                        if (!b) continue;
                        var t = (b.innerText || b.textContent || '').replace(/\\s+/g,' ').trim();
                        var m = t.match(/(\\d+)\\s*%/);
                        if (m) return m[1]+'%';
                    }
                    // Método 3: fallback PVPR — restrito à zona do produto
                    var zone = document.querySelector('#product-content, .product-detail, .pdp-main, [class*="product-detail"]') || document.body;
                    var txt = (zone.innerText || zone.textContent || '');
                    var mp = txt.match(/(\\d+[,.]\\d+)\\s*€?\\s*(\\d+[,.]\\d+)\\s*PVPR/i);
                    if (mp) {
                        var p = parseFloat(mp[1].replace(',','.'));
                        var v = parseFloat(mp[2].replace(',','.'));
                        if (v > p && v > 0) return Math.round((v-p)/v*100)+'%';
                    }
                    return '';
                }"""
                try:
                    _badge_before = page.evaluate(_GET_BADGE_JS)
                except Exception:
                    _badge_before = ''
                # Clica na variante e aguarda AJAX
                # Sempre clica — se já estava seleccionada, REF não muda mas extract_discount
                # ainda é chamado e lê o desconto do estado actual da página
                tile_loc.click(timeout=2000, force=True)
                # Aguarda AJAX - tenta até 6x com 150ms até REF mudar
                ref_v = ref_base
                for _w in range(6):
                    page.wait_for_timeout(150)
                    candidate = extract_ref(page, previous_ref=last_ref)
                    if candidate and candidate != last_ref:
                        ref_v = candidate
                        break
                else:
                    ref_v = extract_ref(page) or ref_base
                    log(f"      [VARIANTE] '{tile_label}' REF não mudou após click")
                last_ref = ref_v
                # Aguarda badge estabilizar após AJAX (até 3s)
                page.wait_for_timeout(400)
                _badge = ''
                for _d in range(15):
                    page.wait_for_timeout(200)
                    try:
                        _badge_now = page.evaluate(_GET_BADGE_JS)
                    except Exception:
                        _badge_now = ''
                    if _badge_now:
                        _badge = _badge_now
                        break
                # Extrai desconto apos AJAX (pode variar por variante)
                # extract_discount já tem 3 métodos incluindo fallback PVPR — não usar body.innerText
                # pois apanharia desconto de outras variantes na mesma página
                desconto_v = extract_discount(page)
                # Se DOM nao foi conclusivo, verifica botão add-to-cart (mais fiável que body.innerText)
                # Não usar body.innerText — apanha OOS de outras variantes/produtos relacionados
                if dom_oos == -1:
                    try:
                        dom_oos = page.evaluate("""() => {
                            // 1. Botão add-to-cart desativado → OOS
                            var btn = document.querySelector(
                                'button.js-add-to-cart, button[data-action="add-to-cart"], ' +
                                'button.add-to-cart, button[name="add-to-cart"]'
                            );
                            if (btn) {
                                if (btn.disabled || btn.classList.contains('disabled') ||
                                    btn.classList.contains('unselectable')) return 1;
                                return 0;
                            }
                            // 2. Fallback: apenas zona do produto activo (não body inteiro)
                            var zone = document.querySelector(
                                '#product-content, .product-detail, .product-info, ' +
                                '.pdp-main, [class*="product-detail"], [class*="pdp"]'
                            );
                            if (!zone) return 0;
                            var txt = zone.innerText.toLowerCase();
                            var kws = ['indisponivel','indisponível','esgotado','sem stock',
                                       'temporariamente indisponivel','temporariamente indisponível',
                                       'avisar quando disponivel','avisar-me'];
                            return kws.some(function(k){ return txt.includes(k); }) ? 1 : 0;
                        }""")
                    except Exception:
                        dom_oos = 0
            except Exception as _ex:
                log(f"      [VARIANTE] ERRO ao processar '{tile_label}': {_ex}")
                if dom_oos == -1:
                    dom_oos = 0
                # Tenta extrair desconto mesmo após erro no click/AJAX
                try:
                    desconto_v = extract_discount(page)
                    log(f"      [DESCONTO] fallback após erro -> {desconto_v}")
                except Exception:
                    desconto_v = "Sem desconto"

            rows.append({
                "data":          data_hoje,
                "marca":         brand_label,
                "titulo":        titulo,
                "url":           final_url,
                "nome_variante": tile_label,
                "is_oos":        dom_oos,
                "ref_produto":   ref_v,
                "desconto":      desconto_v,
            })


    return rows



# ---------------------------------------------------------------------------
# PERSISTENCIA
# ---------------------------------------------------------------------------

def load_known(path: str) -> Set[str]:
    if not os.path.exists(path):
        return set()
    try:
        data = json.load(open(path, "r", encoding="utf-8"))
        if isinstance(data, dict) and isinstance(data.get("known_oos"), list):
            return set(data["known_oos"])
        if isinstance(data, list):
            return set(data)
    except Exception:
        pass
    return set()

def save_known(path: str, known: Set[str]) -> None:
    payload = {"updated_utc": utc_now_ts(), "known_oos": sorted(known)}
    json.dump(payload, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

def append_history(history_csv: str, run_id: str, urls_oos: Set[str]) -> None:
    exists = os.path.exists(history_csv)
    with open(history_csv, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        if not exists:
            w.writerow(["run_id", "timestamp_utc", "url"])
        ts = utc_now_ts()
        for u in sorted(urls_oos):
            w.writerow([run_id, ts, u])


# ---------------------------------------------------------------------------
# EXCEL
# ---------------------------------------------------------------------------

def write_xlsx(xlsx_path: str, headers: List[str], rows: List[List]) -> None:
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.styles import PatternFill

    RED_FILL   = PatternFill("solid", fgColor="FFCCCC")   # vermelho claro - OOS
    GREEN_FILL = PatternFill("solid", fgColor="CCFFCC")   # verde claro - disponivel
    GREY_FILL  = PatternFill("solid", fgColor="F2F2F2")   # cinza - linha alternada de produto
    HEADER_FILL= PatternFill("solid", fgColor="1F4E79")   # azul escuro - cabecalho
    HEADER_FONT= Font(color="FFFFFF", bold=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Verificacao OOS"

    # Cabecalho
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 22

    # Identifica coluna is_oos, nome_variante e url
    try:
        oos_col       = headers.index("is_oos") + 1
        variante_col  = headers.index("nome_variante") + 1
        url_col       = headers.index("url") + 1
    except ValueError:
        oos_col = variante_col = url_col = None

    # Agrupa linhas por URL para colorir alternadamente por produto
    prev_url = None
    use_grey = False

    for row_idx, r in enumerate(rows, 2):
        ws.append(r)

        # Alterna cor de fundo por produto (agrupa variantes do mesmo produto)
        cur_url = r[url_col - 1] if url_col else ""
        if cur_url != prev_url:
            use_grey = not use_grey
            prev_url = cur_url

        is_oos = r[oos_col - 1] if oos_col else 0

        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.alignment = Alignment(vertical="center", wrap_text=(c == url_col))

            if is_oos:
                cell.fill = RED_FILL
            elif use_grey:
                cell.fill = GREY_FILL

        # Coluna is_oos: mostra "OOS" ou "OK" em vez de 1/0
        if oos_col:
            cell_oos = ws.cell(row=row_idx, column=oos_col)
            cell_oos.value = "Out of Stock" if is_oos else "In Stock"
            cell_oos.font = Font(bold=True, color="CC0000" if is_oos else "006600")
            cell_oos.alignment = Alignment(horizontal="center", vertical="center")

    # Largura das colunas
    col_widths = {
        "data":          12,
        "marca":         14,
        "titulo":        40,
        "url":           55,
        "nome_variante": 18,
        "is_oos":        10,
        "ref_produto":   14,
    }
    for c, h in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(c)].width = col_widths.get(h, 15)

    # Tabela
    last_row = len(rows) + 1
    last_col = get_column_letter(len(headers))
    if last_row > 1:
        table = Table(displayName="Verificacao", ref=f"A1:{last_col}{last_row}")
        style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False, showLastColumn=False,
            showRowStripes=False, showColumnStripes=False
        )
        table.tableStyleInfo = style
        ws.add_table(table)

    wb.save(xlsx_path)


# ---------------------------------------------------------------------------
# EMAIL
# ---------------------------------------------------------------------------

def send_email(subject: str, body: str, attachments: List[str]) -> None:
    host = getattr(config_email, "SMTP_HOST", "smtp.gmail.com")
    port = int(getattr(config_email, "SMTP_PORT", 587))
    user = getattr(config_email, "SMTP_USER", None)
    pw = getattr(config_email, "SMTP_PASS", None)
    to_addr  = getattr(config_email, "EMAIL_TO",  None)
    bcc_addr = getattr(config_email, "EMAIL_BCC", None)
    if not user or not pw or not to_addr:
        raise RuntimeError("config_email.py precisa de SMTP_USER, SMTP_PASS e EMAIL_TO")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = to_addr
    if bcc_addr:
        msg["Bcc"] = bcc_addr
    msg.set_content(body)

    for path in attachments:
        if not path or not os.path.exists(path):
            continue
        ctype, enc = mimetypes.guess_type(path)
        if ctype is None or enc is not None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        with open(path, "rb") as f:
            msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(path))

    with smtplib.SMTP(host, port) as server:
        server.starttls()
        server.login(user, pw)
        # Passa destinatários explicitamente para que o BCC seja entregue
        # send_message sozinho ignora o header Bcc
        all_recipients = [to_addr]
        if bcc_addr:
            all_recipients += [a.strip() for a in bcc_addr.split(",") if a.strip()]
        server.send_message(msg, to_addrs=all_recipients)


# ---------------------------------------------------------------------------
# RUN POR MARCA
# ---------------------------------------------------------------------------

def _oos_worker(args: Tuple) -> List[Dict]:
    """
    Worker top-level para verificacao OOS em paralelo.
    Usa check_product que devolve 1 linha por variante.
    Cada chamada lanca o seu proprio sync_playwright isolado.
    """
    widx, wu, label, total_urls = args
    try:
        from playwright.sync_api import sync_playwright as _spw
        with _spw() as _pw:
            _br = _pw.chromium.launch(headless=True)
            _ctx = _br.new_context(viewport={"width": 1365, "height": 900})
            _pg = _ctx.new_page()
            _pg.set_default_timeout(20000)   # 20s max por operacao
            _pg.set_default_navigation_timeout(25000)  # 25s max por navegacao
            try:
                product_rows = check_product(_pg, wu, label)
                oos_count = sum(r["is_oos"] for r in product_rows)
                status = "OOS" if oos_count > 0 else "OK "
                variants_info = f" ({len(product_rows)} variantes)" if len(product_rows) > 1 else ""
                log(f"{label}: [{widx}/{total_urls}] {status}: {wu}{variants_info}")
                return product_rows
            finally:
                try: _ctx.close()
                except Exception: pass
                try: _br.close()
                except Exception: pass
    except Exception as e:
        log(f"{label}: [{widx}/{total_urls}] ERRO: {wu} -> {e}")
    return []


def run_brand(cfg: Dict, page, historico: list = None) -> Dict:
    key = cfg["key"]
    label = cfg["label"]
    base_url = cfg.get("base_url", "")
    exclude_urls = set(cfg.get("exclude_urls") or set())
    max_pages_fallback = int(cfg.get("max_pages_fallback", 20))
    multi_brand = cfg.get("multi_brand", False)
    brand_filter: List[str] = cfg.get("brand_filter") or []

    state_dir = f"state_{key}"
    logs_dir  = f"logs_{key}"
    ensure_dirs(state_dir, logs_dir)

    known_json  = os.path.join(state_dir, "oos_known.json")
    all_urls_json = os.path.join(state_dir, "all_urls.json")  # Para detetar removidos
    history_csv = os.path.join(state_dir, "oos_history.csv")

    run_id   = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%SZ")
    _date_fmt = datetime.now(timezone.utc).strftime("%d_%m_%Y")
    # Sanitiza label para nome de ficheiro (remove acentos e caracteres especiais)
    import unicodedata as _ud
    _label_safe = "".join(
        c if c.isalnum() or c in "-_" else "_"
        for c in _ud.normalize("NFD", label)
        if _ud.category(c) != "Mn"
    ).strip("_")
    run_csv  = os.path.join(logs_dir, f"verificacao_{run_id}.csv")
    run_xlsx = os.path.join(logs_dir, f"{_label_safe}_{_date_fmt}.xlsx")

    log(f"--- {label} | Start ---")
    known_oos = load_known(known_json)
    
    # Carrega URLs do run anterior (para detetar removidos)
    prev_all_urls = set()
    if os.path.exists(all_urls_json):
        try:
            with open(all_urls_json, "r", encoding="utf-8") as f:
                prev_all_urls = set(json.load(f))
        except:
            prev_all_urls = set()
    
    # Carrega ESTADO COMPLETO do run anterior (para detetar mudanças OOS ↔ In Stock)
    state_json = os.path.join(state_dir, "products_state.json")
    prev_state = {}  # {url: is_oos}
    if os.path.exists(state_json):
        try:
            with open(state_json, "r", encoding="utf-8") as f:
                prev_state = json.load(f)
        except:
            prev_state = {}

    # Se não há state JSON (ex: GitHub Actions sem persistência),
    # reconstrói o estado anterior a partir do último run do histórico CSV.
    # Isto evita tratar todos os OOS actuais como "novos" a cada run.
    if not prev_state and historico:
        # Agrupa por URL, ficando com o registo mais recente
        # (histórico está em ordem cronológica — o último valor de cada URL é o mais recente)
        last_seen = {}  # {url: is_oos}
        for row in historico:
            url = row.get("url", "")
            if url:
                last_seen[url] = bool(int(row.get("is_oos", 0)))
        prev_state = last_seen
        log(f"{label}: prev_state reconstruído do histórico CSV ({len(prev_state)} URLs)")

    first_run = (len(known_oos) == 0 and len(prev_state) == 0)

    # Inicializa variáveis que serão usadas em ambos os modos (multi-brand ou normal)
    all_urls: Set[str] = set()
    variant_candidates: Set[str] = set()
    total = None

    # ---- MULTI-BRAND SUPPORT (Oral Care) ----
    if multi_brand:
        log(f"{label}: Multi-brand category with {len(cfg.get('sub_brands', []))} sub-brands")
        all_urls_collected = set()
        urls_by_subbrand = {}  # Rastreia quais URLs aparecem em cada sub-marca
        
        for sub_brand in cfg.get("sub_brands", []):
            sub_name = sub_brand.get("name")
            sub_url = sub_brand.get("url")
            is_single = sub_brand.get("single_product", False)
            
            if is_single:
                # Produto individual - adiciona diretamente
                all_urls_collected.add(normalize_url(sub_url))
                log(f"{label}/{sub_name}: Single product added")
            else:
                # Pesquisa normal - processa como marca regular
                try:
                    page.goto(sub_url, timeout=15000, wait_until="domcontentloaded")
                    try_accept_cookies(page)
                    page.wait_for_timeout(600)
                    
                    # Primeira página
                    sub_urls = extract_urls_from_dom(page, label, exclude_urls, expected_min=5, brand_filter=brand_filter)
                    
                    # Rastreia URLs desta sub-marca
                    urls_by_subbrand[sub_name] = set(sub_urls)
                    
                    all_urls_collected.update(sub_urls)
                    log(f"{label}/{sub_name}: {len(sub_urls)} URLs found on page 1")
                    
                    # Verifica se tem mais páginas
                    total = detect_total_results(page)
                    if total and total > PER_PAGE:
                        max_pages = min(math.ceil(total / PER_PAGE), 15)
                        for pg_num in range(2, max_pages + 1):
                            try:
                                if "?" in sub_url:
                                    page_url = f"{sub_url}&start={(pg_num-1)*PER_PAGE}&sz={PER_PAGE}"
                                else:
                                    page_url = f"{sub_url}?start={(pg_num-1)*PER_PAGE}&sz={PER_PAGE}"
                                
                                page.goto(page_url, timeout=12000, wait_until="domcontentloaded")
                                try_accept_cookies(page)
                                page.wait_for_timeout(400)
                                
                                page_urls = extract_urls_from_dom(page, label, exclude_urls, brand_filter=brand_filter)
                                if page_urls:
                                    urls_by_subbrand[sub_name].update(page_urls)
                                    all_urls_collected.update(page_urls)
                                    log(f"{label}/{sub_name} Page {pg_num}: +{len(page_urls)} URLs")
                                else:
                                    log(f"{label}/{sub_name} Page {pg_num}: no new URLs, stopping")
                                    break
                            except Exception as e:
                                log(f"{label}/{sub_name} Page {pg_num}: error - {e}")
                                break
                except Exception as e:
                    log(f"{label}/{sub_name}: Error - {e}")
        
        # DETEÇÃO DE INTRUSOS PINNED (produtos que aparecem em múltiplas sub-marcas)
        # Se um produto aparece em Eludril E Elgydium E Arthrodont, é intruso!
        if len(urls_by_subbrand) >= 2:
            # Conta quantas sub-marcas cada URL aparece
            url_counts = {}
            for sub_name, urls in urls_by_subbrand.items():
                for url in urls:
                    url_counts[url] = url_counts.get(url, 0) + 1
            
            # URLs que aparecem em 2+ sub-marcas são intrusos
            pinned_intruders = set(url for url, count in url_counts.items() if count >= 2)
            
            if pinned_intruders:
                log(f"{label}: Detetados {len(pinned_intruders)} produtos intrusos (aparecem em múltiplas sub-marcas)")
                for intruso in list(pinned_intruders)[:5]:  # Mostra até 5
                    log(f"{label}:   Intruso removido: {intruso}")
                all_urls_collected -= pinned_intruders
                exclude_urls.update(pinned_intruders)
        
        # FILTRO HÍBRIDO para Oral Care:
        # O Wells filtra por marca (&prefn1=brand&prefv1=X), MAS pode retornar produtos relacionados
        # Estratégia: CONFIAR no Wells, mas REJEITAR marcas conhecidas não-PF
        
        excluded_brands = [
            # Maquilhagem
            'lancome', 'nyx', 'it-cosmetics', 'cosmetics-do-it', 'do-it-all',
            'loreal', 'maybelline', 'bourjois', 'rimmel',
            'clinique', 'estee-lauder', 'mac', 'benefit',
            # Perfumes/Luxo
            'peachn-roses', 'idole', 'jelly-job', 'makeup',
            'armani', 'giorgio-armani', 'acqua-di-gio',
            'hugo-boss', 'boss-bottled', 'bottled-beyond',
            # L'Oréal Grupo
            'revitalift', 'loreal-paris', 'paris-revitalift',
            'garnier', 'glass-skin',
            # Dermocosméticos (não PF)
            'nivea', 'dove', 'neutrogena',
            'vichy', 'la-roche', 'bioderma', 'cerave',
        ]
        
        filtered_urls = set()
        for url in all_urls_collected:
            url_lower = url.lower()
            
            # Rejeita apenas marcas conhecidas não-PF
            if any(excluded in url_lower for excluded in excluded_brands):
                log(f"{label}: Produto rejeitado (marca não-PF): {url}")
                continue
            
            # Aceita todo o resto (confia no filtro &prefn1=brand do Wells)
            filtered_urls.add(url)
        
        rejected_count = len(all_urls_collected) - len(filtered_urls)
        if rejected_count > 0:
            log(f"{label}: {rejected_count} produtos rejeitados (marcas não Pierre Fabre)")
        
        all_urls_collected = filtered_urls
        
        # Pula a paginação normal e vai direto para expansão de variantes
        all_urls.update(all_urls_collected)
        # variant_candidates já foi inicializado (multi-brand não usa por agora)
        log(f"{label}: Multi-brand collected {len(all_urls)} URLs total")
    else:
        # ---- Paginação normal para marcas únicas ----
        # Usa ordenacao fixa para garantir consistencia entre paginas
        base_url_sorted = (base_url + ("&" if "?" in base_url else "?") + "srule=product-name-ascending")
        page.goto(base_url_sorted, wait_until="domcontentloaded")
        try_accept_cookies(page)
        try:
            page.wait_for_load_state("networkidle", timeout=5000)
        except Exception:
            pass
        page.wait_for_timeout(500)

        total = detect_total_results(page)
        if total:
            pages     = math.ceil(total / PER_PAGE)
            page_range = range(1, pages + 1)
            log(f"{label}: Total detetado: {total} resultados -> {pages} paginas")
        else:
            page_range = range(1, max_pages_fallback + 1)
            log(f"{label}: AVISO: nao consegui ler o total. Vou paginar ate nao haver produtos novos.")

        # all_urls e variant_candidates já foram inicializados acima
        empty_in_row = 0
        page1_urls: Set[str] = set()  # URLs da pagina 1 para detetar intrusos pinned

        for pnum in page_range:
            # srule fixo para garantir ordenacao consistente entre paginas (evita duplicados).
            # O Wells usa paginação por offset (start=N&sz=24), NÃO page=N.
            # Usar page=N faz o Wells devolver sempre a página 1, o que dispara
            # o detetor de intrusos e apaga todos os produtos da marca.
            if pnum == 1:
                list_url = base_url + ("&" if "?" in base_url else "?") + "srule=product-name-ascending"
            else:
                start = (pnum - 1) * PER_PAGE
                list_url = base_url + ("&" if "?" in base_url else "?") + f"srule=product-name-ascending&start={start}&sz={PER_PAGE}"
            log(f"{label}: Abrindo pagina {pnum}: {list_url}")
            page.goto(list_url, wait_until="domcontentloaded")
            try_accept_cookies(page)

            # Calcula quantos produtos esperamos nesta pagina
            is_last_page = total and (pnum == math.ceil(total / PER_PAGE))
            if is_last_page:
                expected_min = total - (PER_PAGE * (pnum - 1))
            elif total:
                expected_min = PER_PAGE  # 24
            else:
                expected_min = 6

            # Espera activa com retries e scroll para forcar lazy-load
            for attempt in range(3):
                try:
                    page.wait_for_function(
                        """(minCount) => {
                            const urls = Array.from(document.querySelectorAll("a[href]"))
                              .map(a => a.href)
                              .filter(u => /-\\d{5,}\\.html(\\?|#|$)/i.test(u));
                            return new Set(urls).size >= minCount;
                        }""",
                        expected_min,
                        timeout=8000
                    )
                    break
                except Exception:
                    try:
                        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                        page.wait_for_timeout(600)
                        page.evaluate("window.scrollTo(0, 0)")
                        page.wait_for_timeout(500)
                    except Exception:
                        pass

            before = len(all_urls)

            # Recolhe URLs base desta pagina
            page_urls = extract_urls_from_dom(page, label, exclude_urls, expected_min=expected_min, brand_filter=brand_filter)

            if pnum == 1:
                # Guarda URLs da pagina 1 - ainda nao adicionamos a all_urls
                # Esperamos pela pagina 2 para identificar intrusos pinned primeiro
                page1_urls = set(page_urls)
            elif pnum == 2 and page1_urls:
                # Intrusos = URLs que aparecem na pag1 E na pag2
                pinned_intruders = page1_urls & page_urls
                # GUARDA: só remove como intrusos se a sobreposição for pequena (≤ 3 produtos).
                # Se a interseção for grande (> 30% da pág1), o Wells está a repetir resultados
                # entre páginas (ex: catálogo pequeno, bug de paginação) — nesse caso não há
                # verdadeiros intrusos e NÃO removemos nada para evitar perder toda a marca.
                MAX_INTRUDERS = 3
                overlap_ratio = len(pinned_intruders) / len(page1_urls) if page1_urls else 0
                if pinned_intruders and len(pinned_intruders) <= MAX_INTRUDERS and overlap_ratio <= 0.30:
                    for intruder in pinned_intruders:
                        log(f"{label}: Intruso pinned detetado e removido: {intruder}")
                        page1_urls.discard(intruder)
                        exclude_urls.add(intruder)
                elif pinned_intruders:
                    log(f"{label}: AVISO — sobreposição pág1∩pág2 = {len(pinned_intruders)} URLs "
                        f"({overlap_ratio:.0%}) — possível repetição de resultados pelo Wells. "
                        f"Intrusos NÃO removidos para não perder produtos da marca.")
                # Agora adiciona pagina 1 limpa + pagina 2 (sem intrusos confirmados)
                all_urls.update(page1_urls)
                all_urls.update(page_urls - exclude_urls)
            else:
                all_urls.update(page_urls - exclude_urls)

            # Deteta cards com variantes de tamanho (+X Tamanho(s))
            # NOTA: na pág1, os URLs ainda estão em page1_urls (buffer), não em all_urls.
            # Por isso verificamos ambos para não perder candidatos a variante da pág1.
            cards = get_cards_with_variants(page)
            for card in cards:
                u = normalize_url(card.get("url", ""))
                if u and card.get("has_variants") and (u in all_urls or u in page1_urls):
                    variant_candidates.add(u)

            added = len(all_urls) - before
            if pnum == 1:
                log(f"{label}:  +{len(page1_urls)} em buffer | candidatos a expansao: {len(variant_candidates)}")
            else:
                log(f"{label}:  +{added} produtos (total={len(all_urls)}) | candidatos a expansao: {len(variant_candidates)}")

            effective_added = len(page1_urls) if pnum == 1 else added
            if effective_added == 0:
                empty_in_row += 1
            else:
                empty_in_row = 0
            if not total and empty_in_row >= 2:
                log(f"{label}: Paragem automatica (2 paginas seguidas sem novos produtos).")
                break

    # ---- Expandir URLs de variantes ----
    # Para cada produto com "+X Tamanho(s)", entra na pagina e recolhe todos os URLs de variante
    if variant_candidates:
        log(f"{label}: A expandir {len(variant_candidates)} produto(s) com multiplos tamanhos...")
        expanded_count = 0
        for main_url in sorted(variant_candidates):
            before_expand = len(all_urls)
            new_variant_urls = expand_variant_urls(page, main_url, exclude_urls)
            all_urls.update(new_variant_urls)
            added_variants = len(all_urls) - before_expand
            if added_variants > 0:
                expanded_count += added_variants
                log(f"{label}:   {main_url} -> +{added_variants} variantes")
        log(f"{label}: Expansao concluida. +{expanded_count} URLs adicionados.")

    urls_sorted = sorted(all_urls)
    log(f"{label}: Total de produtos a verificar: {len(urls_sorted)}" +
        (f" | total site: {total}" if total else ""))

    # ---- Verificar OOS produto a produto (PARALELO) ----
    # Cada produto devolve N linhas (1 por variante de tamanho)
    current_oos: Set[str] = set()
    all_rows: List[Dict] = []

    HEADERS = ["data", "marca", "titulo", "url", "nome_variante", "is_oos", "ref_produto", "desconto"]

    total_urls = len(urls_sorted)
    log(f"{label}: A verificar {total_urls} produtos com {OOS_WORKERS} workers paralelos...")

    # Resultados indexados por posicao original (para manter ordem)
    results_map: Dict[int, List[Dict]] = {}

    with ThreadPoolExecutor(max_workers=OOS_WORKERS) as executor:
        futures = {executor.submit(_oos_worker, (i, u, label, total_urls)): i
                   for i, u in enumerate(urls_sorted, 1)}
        for future in as_completed(futures):
            idx = futures[future]
            try:
                product_rows = future.result(timeout=60)  # max 60s por produto
                if product_rows:
                    results_map[idx] = product_rows
            except TimeoutError:
                url_stuck = [u for i,u in enumerate(urls_sorted,1) if i==idx]
                log(f"{label}: TIMEOUT (>60s) produto [{idx}] {url_stuck[0] if url_stuck else ''} — a saltar")
            except Exception as e:
                log(f"{label}: Erro em worker [{idx}]: {e}")

    # Ordena por index original e aplana lista de listas
    for idx in sorted(results_map.keys()):
        for row_dict in results_map[idx]:
            all_rows.append(row_dict)
            if row_dict["is_oos"]:
                current_oos.add(row_dict["url"])

    # Escreve CSV
    with open(run_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=HEADERS, delimiter=";")
        w.writeheader()
        for row_dict in all_rows:
            w.writerow(row_dict)

    # Escreve XLSX com formatacao
    rows_list = [[r[h] for h in HEADERS] for r in all_rows]
    write_xlsx(run_xlsx, HEADERS, rows_list)

    # ═══════════════════════════════════════════════════════════════════════
    # NOVA LÓGICA: Detecta mudanças de estado (IN STOCK ↔ OOS)
    # ═══════════════════════════════════════════════════════════════════════
    
    # Guarda all_urls atuais para próximo run
    current_all_urls = set(row.get("url") for row in all_rows if row.get("url"))
    with open(all_urls_json, "w", encoding="utf-8") as f:
        json.dump(sorted(current_all_urls), f, ensure_ascii=False, indent=2)
    
    # Cria estado atual: {url: is_oos}
    current_state = {}
    for row in all_rows:
        url = row.get("url")
        if url:
            current_state[url] = bool(row.get("is_oos"))
    
    # Calcula mudanças de estado
    newly_oos = set()       # IN STOCK → OOS (Novos OOS)
    newly_in_stock = set()  # OOS → IN STOCK (Recuperados)

    for url in current_all_urls:
        was_oos = prev_state.get(url, None)
        is_oos_now = current_state.get(url, False)

        if was_oos == False and is_oos_now:
            # Estava In Stock, agora está OOS → novo OOS (reset oos_desde)
            newly_oos.add(url)
        elif was_oos == True and not is_oos_now:
            # Estava OOS, agora está In Stock → recuperado
            newly_in_stock.add(url)
        elif was_oos is None and is_oos_now:
            # URL não estava no state JSON.
            # Consulta o histórico CSV para saber o último estado conhecido:
            #   - Nunca visto          → novo OOS (reset oos_desde)
            #   - Último estado OOS    → OOS contínuo (mantém oos_desde)
            #   - Último estado In Stock → voltou a ficar OOS (reset oos_desde)
            last_hist_is_oos = None
            for row in (historico or []):
                if row.get("url") == url:
                    last_hist_is_oos = bool(int(row.get("is_oos", 0)))
            # se nunca visto OU último estado era In Stock → é novo OOS → reset
            if last_hist_is_oos is None or last_hist_is_oos == False:
                newly_oos.add(url)
    
    # Guarda estado atual para próximo run
    with open(state_json, "w", encoding="utf-8") as f:
        json.dump(current_state, f, ensure_ascii=False, indent=2)
    
    # Usa novos conjuntos (substituem old logic)
    new_oos = newly_oos
    recovered = newly_in_stock
    
    # Calcula removidos (estavam no catálogo, desapareceram)
    removed = prev_all_urls - current_all_urls if prev_all_urls else set()
    
    # Criar rows para produtos REMOVIDOS (para aparecerem na tabela)
    removed_rows = []
    if removed and os.path.exists(history_csv):
        # Lê último CSV para obter info dos produtos removidos
        try:
            import csv as _csv_mod
            with open(history_csv, "r", encoding="utf-8") as csvf:
                reader = _csv_mod.DictReader(csvf, delimiter=";")
                last_run_rows = list(reader)
            
            # Para cada URL removido, cria row com dados do último run
            for url in removed:
                # Procura row no último CSV
                matching_rows = [r for r in last_run_rows if r.get("url") == url]
                if matching_rows:
                    # Usa primeira variante encontrada
                    old_row = matching_rows[0]
                    removed_rows.append({
                        "data": run_id.split("_")[0][6:8] + "/" + run_id.split("_")[0][4:6] + "/" + run_id.split("_")[0][:4],
                        "marca": old_row.get("marca", label),
                        "titulo": old_row.get("titulo", "Produto removido"),
                        "ref_produto": old_row.get("ref_produto", ""),
                        "is_oos": "REMOVIDO",  # Flag especial
                        "desconto": "—",
                        "nome_variante": old_row.get("nome_variante", ""),
                        "url": url,
                        "oos_desde": ""
                    })
                else:
                    # Sem info, cria row básica
                    removed_rows.append({
                        "data": run_id.split("_")[0][6:8] + "/" + run_id.split("_")[0][4:6] + "/" + run_id.split("_")[0][:4],
                        "marca": label,
                        "titulo": "Produto removido do catálogo",
                        "ref_produto": url.split("/")[-1].replace(".html", ""),
                        "is_oos": "REMOVIDO",
                        "desconto": "—",
                        "nome_variante": "—",
                        "url": url,
                        "oos_desde": ""
                    })
        except Exception as e:
            log(f"{label}: AVISO - Erro ao ler histórico para removidos: {e}")
    
    save_known(known_json, current_oos)
    append_history(history_csv, run_id, current_oos)

    current_count_variants = sum(1 for row in all_rows if row["is_oos"])
    new_oos_count_variants = sum(1 for row in all_rows if row["is_oos"] and row.get("url") in new_oos)
    recovered_count_variants = sum(1 for row in all_rows if not row["is_oos"] and row.get("url") in recovered)
    removed_count = len(removed)  # Removidos são por URL (não têm variantes visíveis)
    
    log(f"{label}: OOS atuais: {len(current_oos)} produtos / {current_count_variants} variantes | NOVOS OOS: {len(new_oos)} produtos / {new_oos_count_variants} variantes | RECUPERADOS: {len(recovered)} / {recovered_count_variants} | REMOVIDOS: {removed_count}")
    log(f"{label}: XLSX: {run_xlsx}")
    log(f"--- {label} | End ---")

    return {
        "label":                  label,
        "first_run":              first_run,
        "run_xlsx":               run_xlsx,
        "current_oos":            sorted(current_oos),
        "new_oos":                sorted(new_oos),
        "recovered":              sorted(recovered),
        "removed":                sorted(removed),
        "removed_rows":           removed_rows,  # Rows para tabela
        "products_found":         len(urls_sorted),
        "total_detected":         total,
        "current_count":          len(current_oos),            # por produto/URL (para known_oos)
        "current_count_variants": current_count_variants,      # por variante (igual ao dashboard)
        "new_count":              len(new_oos),                 # por produto/URL
        "new_count_variants":     new_oos_count_variants,      # por variante (igual ao dashboard)
        "recovered_count":        len(recovered),
        "recovered_count_variants": recovered_count_variants,
        "removed_count":          removed_count,
        "all_rows":               all_rows,
        "headers":                HEADERS,
    }


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------


def write_consolidated_xlsx(xlsx_path: str, results: List[Dict]) -> None:
    """Cria Excel consolidado com uma sheet por marca + sheet de resumo."""
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.styles import PatternFill

    RED_FILL    = PatternFill("solid", fgColor="FFCCCC")
    GREEN_FILL  = PatternFill("solid", fgColor="CCFFCC")
    GREY_FILL   = PatternFill("solid", fgColor="F2F2F2")
    HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
    HEADER_FONT = Font(color="FFFFFF", bold=True)
    OOS_FONT_RED   = Font(bold=True, color="CC0000")
    OOS_FONT_GREEN = Font(bold=True, color="006600")

    wb = Workbook()

    # ---- Sheet de Resumo ----
    ws_resumo = wb.active
    ws_resumo.title = "Resumo"
    resumo_headers = ["Marca", "Produtos", "Out of Stock", "Novos Out of Stock", "Sem Desconto"]
    ws_resumo.append(resumo_headers)
    for c in range(1, 6):
        cell = ws_resumo.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for r in results:
        # Sem desconto: conta por variante/linha (igual ao dashboard)
        sem_desc = sum(
            1 for row in r.get("all_rows", [])
            if (row.get("desconto", "") or "").strip() in ("", "Sem desconto")
        )
        ws_resumo.append([
            r["label"],
            r["products_found"],
            r.get("current_count_variants", r["current_count"]),  # por variante, igual ao dashboard
            r.get("new_count_variants", r["new_count"]),
            sem_desc,
        ])
    for c, w in zip(range(1, 6), [20, 14, 16, 14, 16]):
        ws_resumo.column_dimensions[get_column_letter(c)].width = w

    # ---- Sheet por marca ----
    col_widths = {
        "data": 12, "marca": 14, "titulo": 40,
        "url": 55, "nome_variante": 18, "is_oos": 16, "ref_produto": 14, "desconto": 14,
    }

    for r in results:
        headers  = r.get("headers", [])
        all_rows = r.get("all_rows", [])
        if not headers or not all_rows:
            continue

        # Nome da sheet limitado a 31 chars (limite Excel)
        sheet_name = r["label"][:31]
        ws = wb.create_sheet(title=sheet_name)

        # Cabecalho
        ws.append(headers)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=c)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 22

        try:
            oos_col      = headers.index("is_oos") + 1
            url_col      = headers.index("url") + 1
        except ValueError:
            oos_col = url_col = None

        prev_url = None
        use_grey = False

        for row_dict in all_rows:
            row_vals = [row_dict.get(h, "") for h in headers]
            ws.append(row_vals)
            row_idx = ws.max_row

            cur_url  = row_dict.get("url", "")
            if cur_url != prev_url:
                use_grey = not use_grey
                prev_url = cur_url

            is_oos = row_dict.get("is_oos", 0)

            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=c)
                cell.alignment = Alignment(vertical="center")
                if is_oos:
                    cell.fill = RED_FILL
                elif use_grey:
                    cell.fill = GREY_FILL

            if oos_col:
                cell_oos = ws.cell(row=row_idx, column=oos_col)
                cell_oos.value = "Out of Stock" if is_oos else "In Stock"
                cell_oos.font = OOS_FONT_RED if is_oos else OOS_FONT_GREEN
                cell_oos.alignment = Alignment(horizontal="center", vertical="center")

        for c, h in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(c)].width = col_widths.get(h, 15)

        # Tabela
        last_row = len(all_rows) + 1
        last_col = get_column_letter(len(headers))
        if last_row > 1:
            tbl = Table(displayName=f"Marca_{r['label'].replace(' ','_').replace('-','_')[:20]}",
                        ref=f"A1:{last_col}{last_row}")
            style = TableStyleInfo(name="TableStyleMedium9",
                showFirstColumn=False, showLastColumn=False,
                showRowStripes=False, showColumnStripes=False)
            tbl.tableStyleInfo = style
            ws.add_table(tbl)

    # ---- Sheet de Variantes ----
    all_variant_rows = []
    for r in results:
        for row in r.get("all_rows", []):
            if row.get("nome_variante") and row["nome_variante"].strip():
                all_variant_rows.append(row)

    if all_variant_rows:
        ws_var = wb.create_sheet(title="Variantes")
        var_headers = ["marca", "titulo", "nome_variante", "is_oos", "ref_produto", "desconto", "url"]
        var_labels  = ["Marca", "Produto", "Variante", "Estado", "REF", "Desconto", "URL"]
        ws_var.append(var_labels)
        for c in range(1, len(var_labels) + 1):
            cell = ws_var.cell(row=1, column=c)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center")

        prev_titulo = None
        use_grey_v  = False
        for row_dict in sorted(all_variant_rows, key=lambda x: (x.get("marca",""), x.get("titulo",""), x.get("nome_variante",""))):
            is_oos = row_dict.get("is_oos", 0)
            titulo = row_dict.get("titulo", "")
            if titulo != prev_titulo:
                use_grey_v = not use_grey_v
                prev_titulo = titulo

            ws_var.append([
                row_dict.get("marca", ""),
                titulo,
                row_dict.get("nome_variante", ""),
                "Out of Stock" if is_oos else "In Stock",
                row_dict.get("ref_produto", ""),
                row_dict.get("desconto", ""),
                row_dict.get("url", ""),
            ])
            row_idx = ws_var.max_row
            for c in range(1, 8):
                cell = ws_var.cell(row=row_idx, column=c)
                cell.alignment = Alignment(vertical="center")
                if is_oos:
                    cell.fill = RED_FILL
                elif use_grey_v:
                    cell.fill = GREY_FILL
            # Estado cell font
            estado_cell = ws_var.cell(row=row_idx, column=4)
            estado_cell.font  = OOS_FONT_RED if is_oos else OOS_FONT_GREEN
            estado_cell.alignment = Alignment(horizontal="center", vertical="center")

        for c, w in zip(range(1, 8), [16, 42, 20, 16, 14, 12, 55]):
            ws_var.column_dimensions[get_column_letter(c)].width = w
        ws_var.freeze_panes = "A2"

    wb.save(xlsx_path)

# Funcao a integrar no bot - gera dashboard HTML com o mesmo estilo do PF_Top10_Dashboard

DASHBOARD_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="pt">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Wells &amp; P&amp;C Online Daily Stocks Monitor</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/flatpickr/dist/flatpickr.min.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/flatpickr"></script>
<script src="https://cdn.jsdelivr.net/npm/flatpickr/dist/l10n/pt.js"></script>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --blue:#1565C0;--blue-lt:#1E88E5;--blue-pale:#E3F2FD;--blue-mid:#90CAF9;
  --text:#0f1923;--muted:#6b7a8d;--line:#e4e9f0;--card:#fff;--bg:#f4f6fb;
  --shadow:0 2px 14px rgba(15,25,50,.07);--shadow-h:0 8px 28px rgba(15,25,50,.14);
  --radius:14px;
  --red:#c62828;--red-bg:#ffebee;--red-mid:#ef9a9a;
  --green:#2e7d32;--green-bg:#e8f5e9;--green-mid:#a5d6a7;
  --orange:#e65100;
  --c-avene:#FF8874;
  --c-ducray:#007AB0;
  --c-klorane:#008000;
  --c-rene:#111111;
  --c-aderma:#99b445;
  --c-oral:#00878e;
  --c-dexeryl:#6B0B16;
}
html{font-family:"DM Sans",sans-serif;background:var(--bg);color:var(--text);font-size:14px;line-height:1.55}
a{color:var(--blue);text-decoration:none}a:hover{text-decoration:underline}

/* HEADER */
.header{background:linear-gradient(118deg,#0D47A1 0%,#1565C0 55%,#1976D2 100%);color:#fff;padding:16px 36px;
  display:flex;align-items:center;gap:16px;border-bottom:3px solid rgba(255,255,255,.10);
  position:sticky;top:0;z-index:200;box-shadow:0 4px 24px rgba(13,71,161,.38);}
.pf-logo-wrap{width:52px;height:52px;border-radius:50%;flex-shrink:0;overflow:hidden;
  border:2.5px solid rgba(255,255,255,.6);box-shadow:0 2px 12px rgba(0,0,0,.3);}
.pf-logo-wrap img{width:52px;height:52px;object-fit:cover;display:block}
.header-text h1{font-size:20px;font-weight:800;letter-spacing:-.4px;line-height:1.2}
.header-text .sub{font-size:12px;opacity:.7;margin-top:3px}
.header-right{margin-left:auto;text-align:right;font-size:11px;opacity:.75;line-height:1.9;flex-shrink:0}
.header-right strong{font-size:13px;opacity:1;font-family:"DM Mono",monospace}

/* LAYOUT */
.wrap{max-width:1460px;margin:0 auto;padding:26px 28px 64px}
.sec{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);
  margin:28px 0 12px;display:flex;align-items:center;gap:8px}
.sec::after{content:"";flex:1;height:1px;background:var(--line)}

/* KPI */
.kpi-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:14px}
.kpi{background:var(--card);border-radius:var(--radius);border:1px solid var(--line);
  padding:18px 20px 14px;box-shadow:var(--shadow);transition:box-shadow .2s,transform .2s;position:relative;overflow:hidden;cursor:pointer}
.kpi:hover{box-shadow:var(--shadow-h);transform:translateY(-2px)}
.kpi::before{content:"";position:absolute;top:0;left:0;right:0;height:4px}
.kpi.k-total::before{background:var(--blue)}
.kpi.k-oos::before{background:var(--red)}
.kpi.k-ok::before{background:var(--green)}
.kpi.k-new::before{background:var(--orange)}
.kpi.k-nodesc::before{background:#6a1b9a}
/* Barras coloridas para Recuperados (verde) e Removidos (vermelho) */
.kpi.k-new.k-new-down::before{background:var(--green)}
.kpi.k-new.k-new-down .kpi-val{color:var(--green)}
.kpi.k-removed::before{background:#b71c1c}
.kpi.k-removed .kpi-val{color:#b71c1c}
.kpi.k-nodesc .kpi-val{color:#6a1b9a}
.kpi-lbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--muted)}
.kpi-val{font-size:34px;font-weight:800;margin-top:6px;line-height:1;font-variant-numeric:tabular-nums}
.kpi.k-total .kpi-val{color:var(--blue)}
.kpi.k-oos   .kpi-val{color:var(--red)}
.kpi.k-ok    .kpi-val{color:var(--green)}
.kpi.k-new   .kpi-val{color:var(--orange)}
.kpi-sub{font-size:11px;color:var(--muted);margin-top:6px}
.kpi-click{cursor:pointer}
.kpi-click:hover .kpi-val{text-decoration:underline}
.kpi-click.selected{outline:2.5px solid currentColor;outline-offset:2px}

/* BRAND CARDS */
.brand-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:12px}
.brand-card{background:var(--card);border-radius:var(--radius);border:1px solid var(--line);
  padding:16px 18px;box-shadow:var(--shadow);cursor:pointer;transition:all .15s}
.brand-card:hover{box-shadow:var(--shadow-h);transform:translateY(-1px)}
.brand-card.active{background:var(--blue-pale);border-color:var(--blue)}
.bc-name{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;margin-bottom:6px}
.bc-oos{font-size:28px;font-weight:800;line-height:1;margin-bottom:4px}
.bc-oos.zero{color:var(--green)}.bc-oos.some{color:var(--red)}
.bc-stats{font-size:11px;color:var(--muted);margin-bottom:8px}
.prog{height:5px;background:#e8edf3;border-radius:3px;overflow:hidden}
.prog-fill{height:100%;border-radius:3px;transition:width .5s cubic-bezier(.4,0,.2,1)}
.prog-fill.good{background:var(--green)}.prog-fill.warn{background:var(--orange)}.prog-fill.bad{background:var(--red)}

/* FILTERS */
.filters{background:var(--card);border-radius:var(--radius);border:1px solid var(--line);
  padding:18px 22px;box-shadow:var(--shadow);display:flex;flex-wrap:wrap;gap:14px;align-items:flex-end}
.fg{display:flex;flex-direction:column;gap:4px;min-width:130px}
.fg label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--muted)}
.fg select,.fg input[type=text]{padding:7px 11px;border-radius:8px;height:34px;border:1.5px solid var(--line);
  font-size:13px;font-family:inherit;color:var(--text);background:#fff;outline:none;transition:border-color .15s}
.fg select:focus,.fg input:focus{border-color:var(--blue)}
.fg-date-pair{display:flex;flex-direction:column;gap:4px}
.fg-date-pair>label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--muted)}
.date-row{display:flex;align-items:center;gap:8px}
.date-row input{width:130px;padding:7px 10px;height:34px;border-radius:8px;border:1.5px solid var(--line);
  font-size:13px;font-family:inherit;color:var(--text);background:#fff;outline:none;transition:border-color .15s;cursor:pointer}
.date-row input:focus{border-color:var(--blue)}
.date-sep{font-size:13px;color:var(--muted)}
.chk{display:flex;align-items:center;gap:7px;padding-bottom:4px;
  font-size:13px;font-weight:600;color:var(--red);cursor:pointer;align-self:flex-end}
.chk input{transform:scale(1.2);accent-color:var(--red)}
.btn-reset{padding:7px 18px;height:34px;border-radius:8px;background:var(--blue-pale);color:var(--blue);
  border:1.5px solid var(--blue-mid);font-size:13px;font-weight:700;font-family:inherit;
  cursor:pointer;transition:background .15s;align-self:flex-end}
.btn-reset:hover{background:var(--blue-mid)}
.btn-export{padding:7px 18px;height:34px;border-radius:8px;background:var(--blue);color:#fff;
  border:none;font-size:13px;font-weight:700;font-family:inherit;
  cursor:pointer;transition:background .15s;align-self:flex-end}
.btn-export:hover{background:var(--blue-lt)}

/* Botões modernos (layout exato da imagem) */
.btn-reset-modern{padding:9px 18px;height:38px;border-radius:6px;
  background:white;color:#0ea5e9;border:1.5px solid #0ea5e9;
  font-size:14px;font-weight:600;font-family:inherit;cursor:pointer;
  transition:all .2s}
.btn-reset-modern:hover{background:#f0f9ff}
.btn-export-modern{padding:9px 18px;height:38px;border-radius:6px;
  background:#0ea5e9;color:white;border:1.5px solid #0ea5e9;
  font-size:14px;font-weight:600;font-family:inherit;cursor:pointer;
  transition:all .2s}
.btn-export-modern:hover{background:#0284c7;border-color:#0284c7}

/* PERIOD BUTTONS & DROPDOWNS - Mesmo estilo da imagem */
.period-btn{padding:9px 18px;height:38px;border:1.5px solid #0ea5e9;background:white;
  border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;transition:all .2s;
  color:#0ea5e9;font-family:inherit}
.period-btn:hover{background:#f0f9ff}
.period-btn.active{background:#0ea5e9;color:white;border-color:#0ea5e9}
.period-dropdown{padding:9px 16px 9px 18px;height:38px;border:1.5px solid #0ea5e9;background:white;
  border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;transition:all .2s;
  color:#0ea5e9;font-family:inherit;appearance:none;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%230ea5e9' d='M6 8L2 4h8z'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 10px center;padding-right:32px}
.period-dropdown:hover{background:#f0f9ff}
.period-dropdown:focus{outline:none;border-color:#0ea5e9;box-shadow:0 0 0 3px rgba(14,165,233,0.1)}
#btnChartReset{padding:9px 18px;height:38px;border:1.5px solid #0ea5e9;background:white;
  border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;transition:all .2s;
  color:#0ea5e9;font-family:inherit}
#btnChartReset:hover{background:#f0f9ff}

/* TABLE */
.table-card{background:var(--card);border-radius:var(--radius);border:1px solid var(--line);box-shadow:var(--shadow);overflow:hidden}
.tbl-head{padding:14px 20px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--line);flex-wrap:wrap;gap:10px}
.tbl-head h3{font-size:14px;font-weight:700}
.tbl-count{font-size:12px;color:var(--muted);font-weight:500;font-family:"DM Mono",monospace}
.tbl-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
thead th{background:var(--blue-pale);color:var(--blue);font-weight:700;font-size:10px;
  text-transform:uppercase;letter-spacing:.5px;padding:0;
  border-bottom:2px solid var(--blue-mid);white-space:nowrap;vertical-align:top}
.th-inner{display:flex;flex-direction:column;gap:5px;padding:10px 12px 8px}
.th-lbl{display:flex;align-items:center;gap:5px;white-space:nowrap;cursor:pointer;user-select:none;
  padding:2px 5px;border-radius:5px;transition:background .12s}
.th-lbl:hover{background:rgba(21,101,192,.1)}
.th-lbl.no-sort{cursor:default}.th-lbl.no-sort:hover{background:none}
.sort-icon{font-size:10px;opacity:.35;transition:opacity .15s}
.th-lbl.active .sort-icon{opacity:1;color:var(--blue-lt)}
.col-filter{width:100%;min-width:70px;padding:4px 7px;height:26px;border-radius:5px;
  border:1.5px solid var(--blue-mid);font-size:11px;color:var(--text);
  background:rgba(255,255,255,.9);font-family:inherit;outline:none;cursor:pointer;transition:border-color .15s}
.col-filter:focus{border-color:var(--blue);background:#fff}
.col-filter-txt{width:100%;min-width:70px;padding:4px 7px;height:26px;border-radius:5px;
  border:1.5px solid var(--blue-mid);font-size:11px;color:var(--text);
  background:rgba(255,255,255,.9);font-family:"DM Mono",monospace;outline:none;transition:border-color .15s}
.col-filter-txt:focus{border-color:var(--blue);background:#fff}
tbody tr{border-bottom:1px solid var(--line)}
tbody tr:last-child{border-bottom:none}
tbody tr:hover{background:var(--blue-pale)}
tbody tr.oos-row{background:var(--red-bg)}
tbody tr.oos-row:hover{background:#ffd7d7}
tbody tr.oos-critical{background:#ffcdd2;border-left:3px solid #b71c1c}
tbody tr.oos-critical:hover{background:#ef9a9a}
tbody tr.oos-critical td{font-weight:600}
tbody td{padding:9px 12px;vertical-align:middle}
.badge-oos{display:inline-flex;align-items:center;gap:4px;background:var(--red);color:#fff;
  padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700}
.badge-ok{display:inline-flex;align-items:center;background:var(--green-bg);color:var(--green);
  padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700;border:1px solid var(--green-mid)}
.badge-removed{display:inline-flex;align-items:center;gap:4px;background:#b71c1c;color:#fff;
  padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700}
tbody tr.removed-row{background:#ffebee}
tbody tr.removed-row td{color:#b71c1c;font-weight:600}
td.ttl-col{max-width:250px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
td.url-col{max-width:100px}

/* PAGINATION */
.pagination{display:flex;align-items:center;justify-content:center;gap:8px;padding:14px;border-top:1px solid var(--line)}
.btn-pg{padding:5px 14px;border-radius:8px;border:1.5px solid var(--line);background:#fff;
  color:var(--text);font-size:13px;font-weight:600;font-family:inherit;cursor:pointer;transition:all .15s}
.btn-pg:hover{border-color:var(--blue);color:var(--blue)}
.btn-pg:disabled{opacity:.35;cursor:default}
.pg-info{font-size:13px;color:var(--muted);font-family:"DM Mono",monospace}
.empty{text-align:center;padding:48px;color:var(--muted);font-size:14px}
.flatpickr-day.selected,.flatpickr-day.selected:hover{background:var(--blue);border-color:var(--blue)}
.chart-card{background:var(--card);border-radius:var(--radius);border:1px solid var(--line);padding:20px 24px;box-shadow:var(--shadow)}
.chart-card h3{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);margin-bottom:16px}
.chart-wrap{position:relative;height:260px}

/* ── COLLAPSIBLE SECTIONS ── */
.sec{cursor:pointer;user-select:none}
.sec:hover{color:var(--blue-lt)}
.sec .sec-chev{font-size:11px;opacity:.55;transition:transform .25s;margin-left:2px;flex-shrink:0}
.sec .sec-chev.open{transform:rotate(180deg)}
.sec-body{display:none}
.sec-body.open{display:block}

/* ── DISCOUNT RADAR ── */
.radar-toggle{display:flex;align-items:center;justify-content:space-between;cursor:pointer;user-select:none;padding:2px 0}
.radar-toggle:hover .radar-chevron{color:var(--blue)}
.radar-chevron{font-size:16px;color:var(--muted);transition:transform .25s;line-height:1}
.radar-chevron.open{transform:rotate(180deg)}
.radar-body{overflow:hidden;transition:max-height .35s cubic-bezier(.4,0,.2,1),opacity .25s;max-height:0;opacity:0}
.radar-body.open{max-height:700px;opacity:1}
.radar-inner{padding-top:20px;display:grid;grid-template-columns:220px 1fr;gap:28px;align-items:center}
.radar-donut-wrap{position:relative;height:200px;width:200px}
.radar-donut-center{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);text-align:center;pointer-events:none}
.radar-donut-center .rdc-val{font-size:26px;font-weight:800;color:var(--text);line-height:1}
.radar-donut-center .rdc-lbl{font-size:10px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:.05em;margin-top:2px}
.radar-right{display:flex;flex-direction:column;gap:6px}
.radar-band-row{display:grid;grid-template-columns:90px 1fr 44px 44px;align-items:center;gap:10px;padding:7px 8px;border-bottom:1px solid var(--line);border-radius:6px}
.radar-band-row:last-child{border-bottom:none}
.radar-band-label{font-size:12px;font-weight:700;display:flex;align-items:center;gap:6px}
.radar-band-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.radar-track{height:6px;background:#edf0f5;border-radius:3px;overflow:hidden}
.radar-fill{height:100%;border-radius:3px;transition:width .6s cubic-bezier(.4,0,.2,1)}
.radar-count{font-size:12px;font-weight:700;text-align:right;font-family:"DM Mono",monospace}
.radar-pct{font-size:11px;color:var(--muted);text-align:right;font-family:"DM Mono",monospace;display:flex;align-items:center;justify-content:flex-end;gap:4px}
.radar-brand-mini{display:flex;gap:3px;align-items:center;flex-wrap:wrap;margin-top:3px}
.radar-brand-pip{display:flex;align-items:center;gap:3px;font-size:10px;color:var(--muted);font-weight:600;padding:1px 5px;border-radius:8px;background:#f0f2f6}
.radar-sem-desc{display:flex;align-items:center;gap:8px;padding:8px 12px;background:#fff3e0;border-radius:10px;border:1px solid #ffe0b2;margin-top:4px;transition:background .15s,box-shadow .15s}
.radar-sem-desc:hover{background:#ffe0b2;box-shadow:0 2px 8px rgba(230,81,0,.15)}
.radar-sem-desc-icon{font-size:16px}
.radar-sem-desc-text{font-size:12px;font-weight:600;color:#e65100}
.radar-sem-desc-brands{font-size:11px;color:#bf360c;margin-top:1px}
.radar-band-clickable{transition:background .15s,box-shadow .15s;cursor:pointer}
.radar-band-clickable:hover{background:#edf3fb;box-shadow:inset 3px 0 0 var(--blue)}
.radar-band-active{background:var(--blue-pale) !important;box-shadow:inset 3px 0 0 var(--blue) !important}
.radar-band-active .radar-band-label{color:var(--blue)}
.radar-filter-icon{font-size:10px;opacity:.3;transition:opacity .15s}
.radar-band-clickable:hover .radar-filter-icon,.radar-band-active .radar-filter-icon{opacity:1;color:var(--blue)}

/* ── STORE TABS ── */
.store-tabs{display:flex;gap:8px;margin-bottom:20px}
.store-tab{padding:7px 20px;border-radius:20px;border:2px solid var(--blue);background:#fff;color:var(--blue);font-weight:700;font-size:13px;cursor:pointer;transition:all .15s}
.store-tab.active{background:var(--blue);color:#fff}
.store-tab:hover:not(.active){background:var(--blue-pale)}
@media(max-width:768px){
  .header{padding:12px 16px}
  .header-right{display:none}
  .wrap{padding:12px 14px 48px}
  .kpi-grid{grid-template-columns:repeat(3,1fr);gap:10px}
  .brand-grid{grid-template-columns:repeat(3,1fr);gap:8px}
  .kpi-val{font-size:26px}
  .bc-oos{font-size:22px}
  .radar-inner{grid-template-columns:1fr;gap:16px}
  .radar-band-row{grid-template-columns:75px 1fr 38px 38px;font-size:11px}
  .store-tabs{flex-wrap:wrap;gap:6px}
  .store-tab{padding:6px 14px;font-size:12px}
  .chart-card{padding:14px 12px}
  .fg{min-width:100px}
}
@media(max-width:480px){
  .kpi-grid{grid-template-columns:repeat(2,1fr);gap:8px}
  .brand-grid{grid-template-columns:repeat(2,1fr);gap:8px}
  .kpi-val{font-size:22px}
  .bc-oos{font-size:20px}
  .radar-inner{grid-template-columns:1fr}
  .tbl-head{flex-direction:column;align-items:flex-start}
  .filter-bar{gap:8px}
  .fg{min-width:calc(50% - 4px)}
}
</style>
</head>
<body>

<div class="header">
  <div class="pf-logo-wrap">
    <img src="data:image/jpeg;base64,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" alt="PF" id="pfLogoImg">
  </div>
  <div class="header-text">
    <h1>Wells &amp; P&amp;C Online Daily Stocks Monitor</h1>
    <div class="sub">Wells.pt &nbsp;·&nbsp; Perfumes&amp;Companhia &nbsp;·&nbsp; Avène · Ducray · Klorane · René Furterer · A-Derma · Dexeryl · Oral Care</div></div>
  </div>
  <div class="header-right">
    Run <strong id="hRunId">{{RUN_ID}}</strong><br>
    <span id="hDate">{{RUN_DATE}}</span>
  </div>
</div>

<div class="wrap">

  <div class="store-tabs">
    <button class="store-tab active" id="tabAll">Todos</button>
    <button class="store-tab" id="tabWells">Wells.pt</button>
    <button class="store-tab" id="tabPC">P&amp;C</button>
  </div>

  <div class="sec" id="secResumo">Resumo <span class="sec-chev open" id="chevResumo">▼</span></div>
  <div class="sec-body open" id="bodyResumo"><div class="kpi-grid">
    <div class="kpi k-total kpi-click" id="kpi-total" title="Ver todos os produtos"><div class="kpi-lbl">Total Verificados</div><div class="kpi-val" id="kTotal">—</div><div class="kpi-sub">produtos + variantes</div></div>
    <div class="kpi k-oos kpi-click" id="kpi-oos" title="Ver só Out of Stock">
      <div class="kpi-lbl">Out of Stock</div>
      <div class="kpi-val" id="kOos">—</div>
      <div class="kpi-sub" id="kOosPct">% do catálogo</div>
      <div id="kOosArrow" style="font-size:13px;font-weight:700;margin-top:3px;height:18px"></div>
    </div>
    <div class="kpi k-ok kpi-click" id="kpi-ok" title="Ver só In Stock"><div class="kpi-lbl">In Stock</div><div class="kpi-val" id="kOk">—</div><div class="kpi-sub" id="kOkPct">% in stock</div></div>
    <div class="kpi k-new kpi-click" id="kpi-new" title="Ver novos Out of Stock">
      <div class="kpi-lbl">Novos Out of Stock</div>
      <div class="kpi-val" id="kNew">—</div>
      <div class="kpi-sub" id="kNewPct" style="display:flex;align-items:center;gap:6px">— do catálogo</div>
      <div id="kNewArrow" style="font-size:13px;font-weight:700;margin-top:3px;height:18px"></div>
    </div>
    <div class="kpi k-new k-new-down kpi-click" id="kpi-gone" title="Ver recuperados">
      <div class="kpi-lbl">Recuperados</div>
      <div class="kpi-val" id="kGone">—</div>
      <div class="kpi-sub" id="kGonePct">vs dia anterior</div>
    </div>
    <div class="kpi k-removed kpi-click" id="kpi-removed" title="Ver removidos do catálogo">
      <div class="kpi-lbl">Removidos</div>
      <div class="kpi-val" id="kRemoved">—</div>
      <div class="kpi-sub" id="kRemovedPct">desapareceram hoje</div>
    </div>
    <div class="kpi k-nodesc kpi-click" id="kpi-nodesc" title="Ver produtos sem desconto">
      <div class="kpi-lbl">Sem Desconto</div>
      <div class="kpi-val" id="kNoDesc">—</div>
      <div class="kpi-sub" id="kNoDescPct">— % do catálogo</div>
    </div>
  </div></div>

  <div class="sec" id="secMarca">Por Marca <span class="sec-chev open" id="chevMarca">▼</span></div>
  <div class="sec-body open" id="bodyMarca"><div class="brand-grid" id="brandGrid"></div></div>

  <div class="sec" id="secFiltros">Filtros <span class="sec-chev open" id="chevFiltros">▼</span></div>
  <div class="sec-body open" id="bodyFiltros"><div class="filters">
    <div class="fg">
      <label>Marca</label>
      <select id="fMarca"><option value="">Todas</option></select>
    </div>
    <div class="fg">
      <label>Estado</label>
      <select id="fEstado">
        <option value="">Todos</option>
        <option value="1">Out of Stock</option>
        <option value="0">In Stock</option>
      </select>
    </div>
    <div class="fg-date-pair">
      <label>Data</label>
      <div class="date-row">
        <input type="text" id="fDateFrom" placeholder="De…">
        <span class="date-sep">→</span>
        <input type="text" id="fDateTo" placeholder="Até…">
      </div>
    </div>
    <div class="fg" style="min-width:200px">
      <label>Pesquisar</label>
      <input type="text" id="fSearch" placeholder="Nome, REF, variante…">
    </div>
    <label class="chk"><input type="checkbox" id="fSoOos"> Só Out of Stock</label>
    <button class="btn-reset" id="btnReset">↺ Limpar</button>
  </div></div>

  <div class="sec" id="secOos">Evolução Out of Stock <span class="sec-chev open" id="chevOos">▼</span></div>
  <div class="sec-body open" id="bodyOos"><div class="chart-card">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px">
      <h3 style="margin-bottom:0">Out of Stock por Dia e Marca</h3>
      <div style="display:flex;align-items:center;gap:8px">
        <div style="display:flex;gap:6px;align-items:center">
          <button class="period-btn" data-period="all">Tudo</button>
          <button class="period-btn" data-period="ytd">YTD</button>
          <select class="period-dropdown" id="periodYear" data-period="year">
            <option value="">Ano</option>
            <option value="2026">2026</option>
            <option value="2025">2025</option>
            <option value="2024">2024</option>
          </select>
          <select class="period-dropdown" id="periodMonth" data-period="month">
            <option value="">Mês</option>
            <option value="01">Janeiro</option>
            <option value="02">Fevereiro</option>
            <option value="03">Março</option>
            <option value="04">Abril</option>
            <option value="05">Maio</option>
            <option value="06">Junho</option>
            <option value="07">Julho</option>
            <option value="08">Agosto</option>
            <option value="09">Setembro</option>
            <option value="10">Outubro</option>
            <option value="11">Novembro</option>
            <option value="12">Dezembro</option>
          </select>
        </div>
        <button id="btnChartReset">↺ Limpar filtros</button>
      </div>
    </div>
    <div class="chart-wrap"><canvas id="oosChart"></canvas></div>
  </div></div>

  <div class="sec" id="secRadar">Radar de Descontos <span class="sec-chev open" id="chevRadar">▼</span></div>
  <div class="sec-body open" id="bodyRadar"><div class="chart-card" style="padding:20px 24px">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:0">
      <h3 style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);margin:0">Distribuição de Descontos por Marca</h3>
      <div style="display:flex;align-items:center;gap:10px">
        <span id="discTotal" style="font-size:12px;color:var(--muted);font-family:'DM Mono',monospace"></span>
        <button id="btnRadarReset" style="display:none;padding:4px 14px;height:28px;border-radius:8px;background:var(--blue-pale);color:var(--blue);border:1.5px solid var(--blue-mid);font-size:12px;font-weight:700;font-family:inherit;cursor:pointer">↺ Limpar filtro</button>
      </div>
    </div>
    <div class="radar-inner">
      <div><div class="radar-donut-wrap"><canvas id="discChart" style="cursor:pointer"></canvas>
        <div class="radar-donut-center"><div class="rdc-val" id="rdcTotal">—</div><div class="rdc-lbl">produtos</div></div>
      </div></div>
      <div class="radar-right" id="radarBands"></div>
    </div>
    <div id="radarSemDesc" style="margin-top:12px"></div>
  </div></div>

  <div class="sec" id="secProdutos">Produtos <span class="sec-chev open" id="chevProdutos">▼</span></div>
  <div class="sec-body open" id="bodyProdutos"><div class="table-card">
    <div class="tbl-head">
      <h3 style="margin:0;font-size:15px;font-weight:600;color:#1e293b">Listagem de Produtos</h3>
      <div style="display:flex;align-items:center;gap:10px">
        <span style="font-size:13px;color:#64748b;font-weight:500" id="rowCount">— resultados</span>
        <button class="btn-reset-modern" id="btnResetTable">↺ Limpar filtros</button>
        <button class="btn-export-modern" id="btnExport">↓ Exportar CSV</button>
      </div>
    </div>
    <div class="tbl-wrap">
      <table>
        <thead id="tHead">
          <tr>
            <th><div class="th-inner">
              <div class="th-lbl" data-col="data">Data <span class="sort-icon" id="si-data">↕</span></div>
            </div></th>
            <th><div class="th-inner">
              <div class="th-lbl" data-col="marca">Marca <span class="sort-icon" id="si-marca">↕</span></div>
              <select class="col-filter" id="cf-marca"><option value="">Todas</option></select>
            </div></th>
            <th><div class="th-inner">
              <div class="th-lbl" data-col="titulo">Produto <span class="sort-icon" id="si-titulo">↕</span></div>
              <input class="col-filter-txt" id="cf-titulo" placeholder="filtrar…">
            </div></th>
            <th><div class="th-inner">
              <div class="th-lbl" data-col="nome_variante">Variante <span class="sort-icon" id="si-nome_variante">↕</span></div>
              <input class="col-filter-txt" id="cf-variante" placeholder="filtrar…">
            </div></th>
            <th><div class="th-inner">
              <div class="th-lbl" data-col="is_oos">Estado <span class="sort-icon" id="si-is_oos">↕</span></div>
              <select class="col-filter" id="cf-estado">
                <option value="">Todos</option>
                <option value="1">Out of Stock</option>
                <option value="0">In Stock</option>
              </select>
            </div></th>
            <th><div class="th-inner">
              <div class="th-lbl" data-col="ref_produto">REF <span class="sort-icon" id="si-ref_produto">↕</span></div>
              <input class="col-filter-txt" id="cf-ref" placeholder="filtrar…">
            </div></th>
            <th><div class="th-inner">
              <div class="th-lbl" data-col="oos_desde">Out of Stock há <span class="sort-icon" id="si-oos_desde">↕</span></div>
              <select class="col-filter" id="cf-oos-dias">
                <option value="">Todos</option>
                <option value="0">Hoje</option>
                <option value="3">≤ 3 dias</option>
                <option value="7">4 a 7 dias</option>
                <option value="30">8 a 30 dias</option>
                <option value="999">&gt; 30 dias</option>
              </select>
            </div></th>
            <th><div class="th-inner">
              <div class="th-lbl" data-col="desconto">Desconto <span class="sort-icon" id="si-desconto">↕</span></div>
              <select class="col-filter" id="cf-desconto"><option value="">Todos</option></select>
            </div></th>
            <th><div class="th-inner">
              <div class="th-lbl no-sort">Link</div>
            </div></th>
          </tr>
        </thead>
        <tbody id="tBody"></tbody>
      </table>
      <div class="empty" id="emptyState" style="display:none">Nenhum produto encontrado.</div>
    </div>
    <div class="pagination">
      <button class="btn-pg" id="btnPrev">← Anterior</button>
      <span class="pg-info" id="pageInfo">Página 1 / 1</span>
      <button class="btn-pg" id="btnNext">Próxima →</button>
    </div>
  </div></div>

</div>

<script>
(function() {

/* ── BRAND COLOURS ── */
var BRAND_COLORS = {
  "Avène":         "#FF8874",
  "Ducray":        "#007AB0",
  "Klorane":       "#008000",
  "René Furterer": "#111111",
  "A-Derma":       "#99b445",
  "Dexeryl":       "#6B0B16",
  "Oral Care":     "#00878e"
};

/* ── DATA ── */
var NEW_OOS = {{NEW_OOS_COUNT}};
var GONE_OOS = {{RECOVERED_COUNT}};
var REMOVED = {{REMOVED_COUNT}};
var DATA = {{DATA_JSON}};  // Run atual (hoje)
var HIST = {{HIST_JSON}};  // Histórico para gráficos

/* ── STATE ── */
var sortState = {col:"oos_desde", dir:"desc"};
var page = 1;
var PAGE_SIZE = 50;
var _activeStore = "all";  /* "all" | "wells" | "pc" */

/* ── HELPERS ── */
function $(id) { return document.getElementById(id); }
function esc(s) {
  return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}
function unique(arr) { return [...new Set(arr)].filter(Boolean).sort(); }
function parseDate(d) {
  if (!d) return "";
  var p = d.split("/");
  return p.length === 3 ? p[2]+p[1]+p[0] : "";
}

/* ── DIAS OOS ── */
function diasOos(oos_desde) {
  if (!oos_desde) return -1;
  var p = oos_desde.split("/");
  if (p.length !== 3) return -1;
  var d = new Date(parseInt(p[2]), parseInt(p[1])-1, parseInt(p[0]));
  var today = new Date(); today.setHours(0,0,0,0);
  var diff = Math.round((today - d) / (1000*60*60*24));
  return diff >= 0 ? diff : -1;
}
function diasBadge(oos_desde, is_oos) {
  if (!is_oos) return "<span style='color:#aaa;font-size:11px'>—</span>";
  var d = diasOos(oos_desde);
  if (d < 0) return "<span style='color:#aaa;font-size:11px'>—</span>";
  var color = d <= 3 ? "#e65100" : d <= 14 ? "#c62828" : "#7b0000";
  var bg    = d <= 3 ? "#fff3e0" : d <= 14 ? "#ffebee" : "#fce4ec";
  return "<span style='background:" + bg + ";color:" + color + ";padding:2px 9px;border-radius:20px;font-size:11px;font-weight:700;font-family:DM Mono,monospace'>" + d + "d</span>";
}

/* ── DESCONTO BADGE ── */
function descontoBadge(desconto) {
  if (!desconto || desconto === "Sem desconto") {
    return "<span style='color:#aaa;font-size:11px'>—</span>";
  }
  return "<span style='background:#fff8e1;color:#e65100;padding:2px 9px;border-radius:20px;font-size:11px;font-weight:700;font-family:DM Mono,monospace'>" + esc(desconto) + "</span>";
}

/* ── FILTERING ── */
function globalFiltered() {
  var fM   = $("fMarca").value;
  var fE   = $("fEstado").value;
  var fS   = $("fSearch").value.toLowerCase();
  var soO  = $("fSoOos").checked;
  var dFrom = window._fpFrom && window._fpFrom.selectedDates[0] ? window._fpFrom.selectedDates[0] : null;
  var dTo   = window._fpTo   && window._fpTo.selectedDates[0]   ? window._fpTo.selectedDates[0]   : null;

  return DATA.filter(function(r) {
    if (_activeStore === "wells" && r.loja === "P&C") return false;
    if (_activeStore === "pc"    && r.loja !== "P&C") return false;
    if (fM && r.marca !== fM) return false;
    if (fE !== "" && String(r.is_oos) !== fE) return false;
    if (soO && !r.is_oos) return false;
    if (fS) {
      var hay = (r.marca+r.titulo+r.nome_variante+r.ref_produto+r.data).toLowerCase();
      if (hay.indexOf(fS) === -1) return false;
    }
    if (dFrom || dTo) {
      var p = r.data.split("/");
      if (p.length === 3) {
        var rd = new Date(parseInt(p[2]), parseInt(p[1])-1, parseInt(p[0]));
        if (dFrom && rd < dFrom) return false;
        if (dTo   && rd > dTo)   return false;
      }
    }
    return true;
  });
}

function tableFiltered(base) {
  var cfM = $("cf-marca").value;
  var cfT = $("cf-titulo").value.toLowerCase();
  var cfV = $("cf-variante").value.toLowerCase();
  var cfE = $("cf-estado").value;
  var cfR = $("cf-ref").value.toLowerCase();
  // var cfD = $("cf-data").value.toLowerCase(); // Dropdown data é informativo apenas
  var cfDesc = $("cf-desconto").value;
  var cfDias = $("cf-oos-dias").value;

  return base.filter(function(r) {
    if (cfDias !== "") {
      var d = diasOos(r.oos_desde);
      if (cfDias === "0"   && d !== 0) return false;
      if (cfDias === "3"   && !(d >= 1 && d <= 3)) return false;
      if (cfDias === "7"   && !(d >= 4 && d <= 7)) return false;
      if (cfDias === "30"  && !(d >= 8 && d <= 30)) return false;
      if (cfDias === "999" && !(d > 30)) return false;
    }
    if (cfM && r.marca !== cfM) return false;
    if (cfT && r.titulo.toLowerCase().indexOf(cfT) === -1) return false;
    if (cfV && (r.nome_variante||"").toLowerCase().indexOf(cfV) === -1) return false;
    if (cfE !== "" && String(r.is_oos) !== cfE) return false;
    if (cfR && (r.ref_produto||"").toLowerCase().indexOf(cfR) === -1) return false;
    // if (cfD && r.data !== cfD) return false; // Dropdown data desabilitado
    if (cfDesc && (r.desconto||"") !== cfDesc) return false;
    if(window._radarBandFilter&&!window._radarBandFilter.test(r.desconto||""))return false;
    return true;
  });
}

function applySort(rows) {
  var col = sortState.col, dir = sortState.dir === "asc" ? 1 : -1;
  return rows.slice().sort(function(a, b) {
    var av, bv;
    if (col === "data")        { av = parseDate(a.data);   bv = parseDate(b.data); }
    else if (col === "oos_desde") { av = diasOos(a.oos_desde); bv = diasOos(b.oos_desde); }
    else if (col === "is_oos")    { av = a.is_oos; bv = b.is_oos; }
    else { av = (a[col]||"").toString().toLowerCase(); bv = (b[col]||"").toString().toLowerCase(); }
    return av < bv ? -dir : av > bv ? dir : 0;
  });
}

/* ── SORT HEADERS ── */
document.querySelectorAll(".th-lbl[data-col]").forEach(function(el) {
  el.addEventListener("click", function() {
    var col = this.getAttribute("data-col");
    if (sortState.col === col) {
      sortState.dir = sortState.dir === "asc" ? "desc" : "asc";
    } else {
      sortState.col = col; sortState.dir = "asc";
    }
    ["data","marca","titulo","nome_variante","is_oos","ref_produto","oos_desde","desconto"].forEach(function(c) {
      var si = $("si-"+c); if (!si) return;
      si.textContent = c === sortState.col ? (sortState.dir === "asc" ? "↑" : "↓") : "↕";
      si.parentElement.classList.toggle("active", c === sortState.col);
    });
    page = 1; renderTable();
  });
});

/* ── DROPDOWNS ── */
function populateGlobal() {
  var marcas = unique(DATA.map(function(r) { return r.marca; }));
  [$("fMarca"), $("cf-marca")].forEach(function(sel) {
    sel.innerHTML = "<option value=''>Todas</option>";
    marcas.forEach(function(m) {
      var opt = document.createElement("option");
      opt.value = m; opt.textContent = m;
      sel.appendChild(opt);
    });
  });
}

function populateColFilters(base) {
  var marcas = unique(base.map(function(r) { return r.marca; }));
  var cf = $("cf-marca"), cur = cf.value;
  cf.innerHTML = "<option value=''>Todas</option>";
  marcas.forEach(function(m) {
    var opt = document.createElement("option");
    opt.value = m; opt.textContent = m;
    if (m === cur) opt.selected = true;
    cf.appendChild(opt);
  });
  // Desconto dropdown
  var descontos = unique(base.map(function(r) { return r.desconto || "Sem desconto"; }));
  var cfD = $("cf-desconto"), curD = cfD.value;
  cfD.innerHTML = "<option value=''>Todos</option>";
  descontos.forEach(function(d) {
    var opt = document.createElement("option");
    opt.value = d; opt.textContent = d;
    if (d === curD) opt.selected = true;
    cfD.appendChild(opt);
  });
}

/* ── KPIs ── */
function renderKPIs(rows) {
  var total = rows.length;
  var oos   = rows.filter(function(r) { return r.is_oos; }).length;
  var ok    = total - oos;
  // Sem desconto: conta por linha/variante (cada variante conta individualmente)
  var noDesc = DATA.filter(function(r) { return !r.desconto || r.desconto === "Sem desconto"; }).length;

  $("kTotal").textContent = total.toLocaleString("pt-PT");
  $("kOos").textContent   = oos.toLocaleString("pt-PT");
  $("kOk").textContent    = ok.toLocaleString("pt-PT");
  $("kNew").textContent   = NEW_OOS;
  
  // Recuperados
  var kGoneEl = $("kGone");
  if (kGoneEl) kGoneEl.textContent = GONE_OOS > 0 ? "-" + GONE_OOS : "0";
  var kGonePct = $("kGonePct");
  if (kGonePct) kGonePct.innerHTML = GONE_OOS > 0
    ? "<span style='color:#27ae60;font-weight:700'>▼ -" + GONE_OOS + " vs dia anterior</span>"
    : "<span style='color:#6b7a8d'>→ nenhum recuperado</span>";
  
  // Removidos
  var kRemovedEl = $("kRemoved");
  if (kRemovedEl) kRemovedEl.textContent = REMOVED > 0 ? REMOVED : "0";
  var kRemovedPct = $("kRemovedPct");
  if (kRemovedPct) kRemovedPct.innerHTML = REMOVED > 0
    ? "<span style='color:#e74c3c;font-weight:700'>▲ " + REMOVED + " removidos hoje</span>"
    : "<span style='color:#6b7a8d'>→ nenhum removido</span>";
  
  $("kNoDesc").textContent = noDesc.toLocaleString("pt-PT");

  // Percentagens
  var oosPct   = total ? (oos/total*100)     : 0;
  var okPct    = total ? (ok/total*100)      : 0;
  var newPct   = total ? (NEW_OOS/total*100) : 0;
  var ndPct    = total ? (noDesc/DATA.length*100) : 0;

  $("kOosPct").textContent  = oosPct.toFixed(1)  + "% do catálogo";
  $("kOkPct").textContent   = okPct.toFixed(1)   + "% in stock";
  $("kNoDescPct").textContent = ndPct.toFixed(1) + "% do catálogo";

  // Trend arrows for OOS
  var oosT = $("kOosTrend");
  if (oosT) {
    var prev = parseFloat(oosT.getAttribute("data-prev") || oosPct);
    if (oosT.getAttribute("data-prev") === null) { oosT.setAttribute("data-prev", oosPct); }
    var diff = oosPct - prev;
    if (Math.abs(diff) > 0.05) {
      oosT.innerHTML = diff > 0
        ? "<span style='color:#c0392b'>▲ +" + diff.toFixed(1) + "%</span>"
        : "<span style='color:#27ae60'>▼ "  + diff.toFixed(1) + "%</span>";
    }
  }

  // % do catalogo + trend arrow for Novos OOS
  var kNewPct = $("kNewPct");
  if (kNewPct) kNewPct.textContent = newPct.toFixed(1) + "% do catálogo";
  var kNewTrend = $("kNewTrend");
  if (kNewTrend) {
    if (NEW_OOS === 0) {
      kNewTrend.innerHTML = "<span style='color:#27ae60'>✓ nenhum novo</span>";
    } else {
      kNewTrend.innerHTML = "<span style='color:#c0392b'>▲ " + NEW_OOS + " novos</span>";
    }
  }

  // Trend arrow for OOS vs HIST (previous day)
  var prevDayOosPct = (function() {
    if (!HIST || !HIST.length) return null;
    var dates = HIST.map(function(r){ return r.data; });
    var allDates = dates.filter(function(v,i,a){ return a.indexOf(v)===i; }).sort(function(a,b){ return parseDate(a)>parseDate(b)?1:-1; });
    var prevDate = allDates.length >= 2 ? allDates[allDates.length-2] : null;
    if (!prevDate) return null;
    var prevRows = HIST.filter(function(r){ return r.data===prevDate; });
    if (!prevRows.length) return null;
    return prevRows.filter(function(r){ return r.is_oos; }).length / prevRows.length * 100;
  })();

  var oosT = $("kOosTrend");
  if (oosT) {
    if (prevDayOosPct !== null) {
      var diff = oosPct - prevDayOosPct;
      if (Math.abs(diff) > 0.05) {
        oosT.innerHTML = diff > 0
          ? "<span style='color:#c0392b'>▲ +" + diff.toFixed(1) + "% vs ontem</span>"
          : "<span style='color:#27ae60'>▼ " + Math.abs(diff).toFixed(1) + "% vs ontem</span>";
      } else {
        oosT.innerHTML = "<span style='color:#6b7a8d'>= sem alteração</span>";
      }
    }
  }
}

/* ── BRAND CARDS ── */
function renderBrands() {
  var BRAND_ORDER = ['Avène','Ducray','Klorane','René Furterer','A-Derma','Dexeryl','Oral Care'];
  var storeData = DATA.filter(function(r) {
    if (_activeStore === "wells") return !r.loja;
    if (_activeStore === "pc")    return r.loja === "P&C";
    return true;
  });
  var marcas = unique(storeData.map(function(r) { return r.marca; })).sort(function(a,b){
    var ai=BRAND_ORDER.indexOf(a),bi=BRAND_ORDER.indexOf(b);
    if(ai===-1)ai=999;if(bi===-1)bi=999;return ai-bi;
  });
  var act    = $("fMarca").value;
  var grid   = $("brandGrid");
  grid.innerHTML = "";

  marcas.forEach(function(m) {
    var rows = storeData.filter(function(r) { return r.marca === m; });
    var oos  = rows.filter(function(r) { return r.is_oos; }).length;
    var pct  = rows.length ? (rows.length - oos) / rows.length * 100 : 100;
    var cls  = pct >= 95 ? "good" : pct >= 80 ? "warn" : "bad";
    var color = BRAND_COLORS[m] || "#666";

    var card = document.createElement("div");
    card.className = "brand-card" + (act === m ? " active" : "");
    card.innerHTML =
      "<div class='bc-name' style='color:" + color + "'>" + esc(m) + "</div>" +
      "<div class='bc-oos' style='color:" + color + "'>" + oos + "</div>" +
      "<div class='bc-stats'>" + oos + " Out of Stock · " + rows.length + " total</div>" +
      (function(){
        try {
          if (!HIST || !HIST.length) return "<div style='font-size:11px;color:#888;margin-top:2px'>→ sem histórico anterior</div>";
          var allDates = [...new Set(HIST.map(function(r){return r.data;}))].sort(function(a,b){return parseDate(a)>parseDate(b)?1:-1;});
          if (allDates.length < 2) return "<div style='font-size:11px;color:#888;margin-top:2px'>→ primeiro registo</div>";
          var prevDate = allDates[allDates.length-2];
          var prevRows = HIST.filter(function(r){return r.data===prevDate && r.marca===m;});
          if (!prevRows.length) return "<div style='font-size:11px;color:#888;margin-top:2px'>→ sem dados do dia anterior</div>";
          var prevDayOos = prevRows.filter(function(r){return r.is_oos;}).length;
          var diff = oos - prevDayOos;
          if (diff===0) return "<div style='font-size:11px;color:#888;margin-top:2px'>→ igual ao dia anterior</div>";
          return diff>0
            ? "<div style='font-size:11px;font-weight:700;color:#c62828;margin-top:2px'>▲ +"+diff+" vs dia anterior</div>"
            : "<div style='font-size:11px;font-weight:700;color:#2e7d32;margin-top:2px'>▼ "+Math.abs(diff)+" vs dia anterior</div>";
        } catch(e) { return "<div style='font-size:11px;color:#888;margin-top:2px'>→ —</div>"; }
      })() +
      "<div class='prog'><div class='prog-fill' style='width:" + pct + "%;background:" + color + "'></div></div>";

    (function(brand) {
      card.addEventListener("click", function() {
        $("fMarca").value = ($("fMarca").value === brand) ? "" : brand;
        page = 1; render();
        // Garante que a secção Produtos está aberta antes de fazer scroll
        var body = document.getElementById("bodyProdutos");
        var chev = document.getElementById("chevProdutos");
        if (body && !body.classList.contains("open")) {
          body.classList.add("open");
          if (chev) chev.classList.add("open");
        }
        setTimeout(scrollToTable, 150);
      });
    })(m);
    grid.appendChild(card);
  });
}

/* ── TABLE ── */
function renderTable() {
  var base  = globalFiltered();
  populateColFilters(base);
  // populateDateFilter(base); // Removido - dropdown de data não existe mais
  var rows  = applySort(tableFiltered(base));
  var total = rows.length;
  $("rowCount").textContent = total.toLocaleString("pt-PT") + " resultados";

  var tp = Math.max(1, Math.ceil(total / PAGE_SIZE));
  if (page > tp) page = tp;
  var slice = rows.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);

  var tbody = $("tBody");
  var empty = $("emptyState");

  if (slice.length === 0) {
    tbody.innerHTML = "";
    empty.style.display = "block";
  } else {
    empty.style.display = "none";
    var color;
    tbody.innerHTML = slice.map(function(r) {
      color = BRAND_COLORS[r.marca] || "#333";
      var badge;
      var rowCls;
      
      if (r.is_oos === "REMOVIDO") {
        badge = "<span class='badge-removed'>🗑 Removido</span>";
        rowCls = "removed-row";
      } else if (r.is_oos) {
        badge = "<span class='badge-oos'>● Out of Stock</span>";
        rowCls = diasOos(r.oos_desde) > 30 ? "oos-critical" : "oos-row";
      } else {
        badge = "<span class='badge-ok'>● In Stock</span>";
        rowCls = "";
      }
      
      return "<tr class='" + rowCls + "'>" +
        "<td style='font-family:\"DM Mono\",monospace;font-size:12px;white-space:nowrap'>" + esc(r.data) + "</td>" +
        "<td><strong style='color:" + color + "'>" + esc(r.marca) + "</strong></td>" +
        "<td class='ttl-col' title='" + esc(r.titulo) + "'>" + esc(r.titulo) + "</td>" +
        "<td style='font-family:\"DM Mono\",monospace;font-size:12px'>" + esc(r.nome_variante||"—") + "</td>" +
        "<td>" + badge + "</td>" +
        "<td style='font-family:\"DM Mono\",monospace;font-size:12px;color:#6b7a8d'>" + esc(r.ref_produto) + "</td>" +
        "<td style='text-align:center'>" + diasBadge(r.oos_desde, r.is_oos) + "</td>" +
        "<td style='text-align:center'>" + descontoBadge(r.desconto) + "</td>" +
        "<td class='url-col'><a href='" + esc(r.url) + "' target='_blank' rel='noopener'>Abrir ↗</a></td>" +
        "</tr>";
    }).join("");
  }

  $("pageInfo").textContent = "Página " + page + " / " + tp;
  $("btnPrev").disabled = page <= 1;
  $("btnNext").disabled = page >= tp;
}

/* ── FULL RENDER ── */
function render() {
  page = 1;
  var gf   = globalFiltered();
  var full = tableFiltered(gf);
  renderKPIs(full);
  renderBrands();
  renderTable();
  renderChart();
}

/* ── PAGINATION ── */
$("btnPrev").addEventListener("click", function() {
  if (page > 1) { page--; renderTable(); scrollToTable(); }
});
$("btnNext").addEventListener("click", function() {
  var tp = Math.max(1, Math.ceil(tableFiltered(globalFiltered()).length / PAGE_SIZE));
  if (page < tp) { page++; renderTable(); scrollToTable(); }
});
function scrollToTable() {
  var el = document.querySelector(".table-card");
  if (el) el.scrollIntoView({behavior:"smooth", block:"start"});
}

/* ── KPI CLICKS ── */
/* Event listeners movidos para baixo (após definição de variáveis) */

$("kpi-nodesc").addEventListener("click", function() {
  _newOosFilter = false; _semDescontoFilter = false;
  $("fMarca").value=""; $("fEstado").value=""; $("fSearch").value=""; $("fSoOos").checked=false;
  if (window._fpFrom) window._fpFrom.clear();
  if (window._fpTo) window._fpTo.clear();
  $("cf-marca").value=""; $("cf-titulo").value=""; $("cf-variante").value="";
  $("cf-estado").value=""; $("cf-ref").value=""; $("cf-desconto").value=""; $("cf-oos-dias").value="";
  _semDescontoFilter = true;
  page = 1; renderTable();
  setTimeout(scrollToTable, 150);
});

/* ── FILTER EVENTS ── */
["fMarca","fEstado"].forEach(function(id) {
  $(id).addEventListener("change", function() { page=1; render(); });
});
$("fSearch").addEventListener("input", function() { page=1; render(); });
$("fSoOos").addEventListener("change", function() { page=1; render(); });
["cf-marca","cf-estado","cf-desconto","cf-oos-dias"].forEach(function(id) {
  $(id).addEventListener("change", function() { if(_resetting) return; page=1; renderTable(); });
});
["cf-titulo","cf-variante","cf-ref"].forEach(function(id) {
  $(id).addEventListener("input", function() { page=1; renderTable(); });
});
// Dropdown "Data" é apenas informativo - mostra datas disponíveis no histórico
// mas não filtra a tabela (que sempre mostra run de hoje)
// $("cf-data").addEventListener("change", function() { page=1; renderTable(); });

/* ── RESET ── */
function resetFilters() {
  _newOosFilter = false; _semDescontoFilter = false; _recoveredFilter = false; _removedFilter = false;
  if(window._clearRadarFilter)window._clearRadarFilter();
  _resetting = true;
  $("fMarca").value=""; $("fEstado").value=""; $("fSearch").value=""; $("fSoOos").checked=false;
  if (window._fpFrom) window._fpFrom.clear();
  if (window._fpTo)   window._fpTo.clear();
  $("cf-marca").value=""; $("cf-titulo").value=""; $("cf-variante").value="";
  $("cf-estado").value=""; $("cf-ref").value=""; $("cf-desconto").value=""; $("cf-oos-dias").value="";
  _newOosFilter = false; _semDescontoFilter = false; _recoveredFilter = false; _removedFilter = false;
  page = 1;
  _resetting = false;
  render();
}

$("btnReset").addEventListener("click", function() {
  _newOosFilter = false; _semDescontoFilter = false; _recoveredFilter = false; _removedFilter = false;
  resetFilters();
});
$("btnResetTable").addEventListener("click", function() {
  _newOosFilter = false; _semDescontoFilter = false; _recoveredFilter = false; _removedFilter = false;
  resetFilters();
});

/* ── EXPORT CSV ── */
$("btnExport").addEventListener("click", function() {
  var rows = applySort(tableFiltered(globalFiltered()));
  var cols = ["data","marca","titulo","nome_variante","is_oos","oos_desde","desconto","ref_produto","url"];
  var hdrs = ["Data","Marca","Produto","Variante","Estado","OOS Desde","Desconto","REF","URL"];
  var lines = ["sep=;", hdrs.join(";")];
  rows.forEach(function(r) {
    var vals = cols.map(function(c) {
      var v = c === "is_oos" ? (r.is_oos ? "Out of Stock" : "In Stock") : (r[c]||"").toString();
      return '"' + v.replace(/"/g, '""') + '"';
    });
    lines.push(vals.join(";"));
  });
  var content = lines.join("\n");
  var buf = new ArrayBuffer(2 + content.length * 2);
  var view = new DataView(buf);
  view.setUint8(0, 0xFF); view.setUint8(1, 0xFE);
  for (var i = 0; i < content.length; i++) view.setUint16(2 + i*2, content.charCodeAt(i), true);
  var blob = new Blob([buf], {type:"text/csv;charset=utf-16le"});
  var url  = URL.createObjectURL(blob);
  var a = document.createElement("a");
  a.href = url; var today = new Date();
  var dd = String(today.getDate()).padStart(2,"0");
  var mm = String(today.getMonth()+1).padStart(2,"0");
  var yyyy = today.getFullYear();
  a.download = "wells_stock_watch_" + dd + "-" + mm + "-" + yyyy + ".csv";
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(function() { URL.revokeObjectURL(url); }, 1000);
});

/* ── FLATPICKR ── */
var fpBase = {locale:"pt", dateFormat:"d/m/Y", allowInput:false, disableMobile:true};
window._fpFrom = flatpickr("#fDateFrom", Object.assign({}, fpBase, {
  onChange: function(sel) {
    if (window._fpTo && sel[0]) window._fpTo.set("minDate", sel[0]);
    render();
  }
}));
window._fpTo = flatpickr("#fDateTo", Object.assign({}, fpBase, {
  onChange: function(sel) {
    if (window._fpFrom && sel[0]) window._fpFrom.set("maxDate", sel[0]);
    render();
  }
}));



/* ── POPULATE DATE FILTER (col) ── */
/* ── KPI CLICK ── */
var NEW_OOS_REFS = {{NEW_OOS_REFS_JSON}};
var RECOVERED_REFS = {{RECOVERED_REFS_JSON}};
var REMOVED_REFS = {{REMOVED_REFS_JSON}};

$("kpi-total").addEventListener("click", function() {
  resetFilters();
  scrollToTable();
});

$("kpi-oos").addEventListener("click", function() {
  $("fEstado").value = "1"; $("fSoOos").checked = false;
  $("cf-estado").value = "1";
  page = 1; render(); scrollToTable();
});

$("kpi-ok").addEventListener("click", function() {
  $("fEstado").value = "0"; $("fSoOos").checked = false;
  $("cf-estado").value = "0";
  page = 1; render(); scrollToTable();
});

$("kpi-new").addEventListener("click", function() {
  resetFilters();
  _newOosFilter = true;
  page = 1; renderTable();
  setTimeout(scrollToTable, 150);
});

$("kpi-gone").addEventListener("click", function() {
  resetFilters();
  _recoveredFilter = true;
  page = 1; renderTable();
  setTimeout(scrollToTable, 150);
});

$("kpi-removed").addEventListener("click", function() {
  resetFilters();
  _removedFilter = true;
  page = 1; renderTable();
  setTimeout(scrollToTable, 150);
});

var _newOosFilter = false;
var _semDescontoFilter = false;
var _recoveredFilter = false;
var _removedFilter = false;
var _resetting = false;

var _origTableFiltered = tableFiltered;
tableFiltered = function(base) {
  var rows;
  if (_newOosFilter) {
    // Mostra produtos que ficaram OOS HOJE (oos_desde = data do run)
    var today = new Date();
    var dd = String(today.getDate()).padStart(2, '0');
    var mm = String(today.getMonth() + 1).padStart(2, '0');
    var yyyy = today.getFullYear();
    var todayStr = dd + "/" + mm + "/" + yyyy;
    
    rows = DATA.filter(function(r) {
      return r.is_oos && r.oos_desde === todayStr;
    });
  } else if (_recoveredFilter) {
    // Mostra produtos recuperados (que voltaram a stock - agora IN STOCK)
    rows = DATA.filter(function(r) {
      return RECOVERED_REFS.indexOf(r.url) !== -1 && !r.is_oos;
    });
  } else if (_removedFilter) {
    // Mostra produtos removidos do catálogo
    rows = DATA.filter(function(r) {
      return r.is_oos === "REMOVIDO";
    });
  } else if (_semDescontoFilter) {
    // Ignora filtros globais — mostra todos sem desconto
    rows = DATA.filter(function(r) {
      return !r.desconto || r.desconto === "Sem desconto";
    });
  } else {
    rows = _origTableFiltered(base);
  }
  return rows;
};

// resetNewOosFilter called by resetFilters()
function resetNewOosFilter() {
  _newOosFilter = false;
  _semDescontoFilter = false;
  _recoveredFilter = false;
  _removedFilter = false;
}

/* ── CHART ── */
var _chart = null;
var _activePeriod = "month"; // Período ativo: all, ytd, year, month

function filterDataByPeriod(rows, period) {
  if (period === "all") return rows;
  
  var now = new Date();
  var currentYear = now.getFullYear();
  var currentMonth = now.getMonth() + 1;
  
  // Se dropdown foi usado, usa valores seleccionados
  var targetYear = _selectedYear || currentYear;
  var targetMonth = _selectedMonth || currentMonth;
  
  return rows.filter(function(r) {
    var parts = r.data.split("/");
    if (parts.length !== 3) return false;
    var rowDay = parseInt(parts[0], 10);
    var rowMonth = parseInt(parts[1], 10);
    var rowYear = parseInt(parts[2], 10);
    
    if (period === "month") {
      return rowYear === targetYear && rowMonth === targetMonth;
    }
    if (period === "year") {
      return rowYear === targetYear;
    }
    if (period === "ytd") {
      // Year to Date: desde 1 janeiro até hoje do ano ACTUAL
      if (rowYear < currentYear) return false;
      if (rowYear > currentYear) return false;
      return true; // Mesmo ano = YTD
    }
    return true;
  });
}

function renderChart() {
  var BRAND_COLORS_CHART = {
    "Avène":              "#FF8874",
    "Ducray":             "#007AB0",
    "Klorane":            "#008000",
    "René Furterer":      "#111111",
    "A-Derma":            "#99b445",
    "Dexeryl":            "#6B0B16",
    "Oral Care":          "#00878e",
    "Avène (P&C)":        "#FF5533",
    "Ducray (P&C)":       "#0099CC",
    "Klorane (P&C)":      "#33AA33",
    "René Furterer (P&C)":"#555555",
    "A-Derma (P&C)":      "#BDD455"
  };

  // HIST is global

  // Combina historico + run actual para o grafico
  var allRows = (HIST && HIST.length > 0) ? HIST : DATA;
  // Garante que a run actual esta incluida
  var dataKeys = new Set(DATA.map(function(r){ return r.data; }));
  var histKeys = new Set(allRows.map(function(r){ return r.data; }));
  dataKeys.forEach(function(d){ if(!histKeys.has(d)) allRows = allRows.concat(DATA.filter(function(r){ return r.data===d; })); });

  // Aplica filtro de período
  allRows = filterDataByPeriod(allRows, _activePeriod);

  // Filtro por loja — sincronizado com separadores Wells / P&C / Todos
  if (_activeStore === "wells") {
    allRows = allRows.filter(function(r){ return r.loja !== "P&C"; });
  } else if (_activeStore === "pc") {
    allRows = allRows.filter(function(r){ return r.loja === "P&C"; });
  } else {
    // "all": distinguir marcas com sufixo (P&C) para não sobrepor linhas
    allRows = allRows.map(function(r){
      return r.loja === "P&C" ? Object.assign({}, r, {marca: r.marca + " (P&C)"}) : r;
    });
  }

  // Marcas presentes nos dados filtrados (na ordem do mapa de cores)
  var marcas = Object.keys(BRAND_COLORS_CHART).filter(function(m){
    return allRows.some(function(r){ return r.marca === m; });
  });
  var allDates = [...new Set(allRows.map(function(r){ return r.data; }))].sort(function(a,b){
    return parseDate(a) > parseDate(b) ? 1 : -1;
  });

  var brandData = {};
  marcas.forEach(function(m) {
    brandData[m] = allDates.map(function(d) {
      return allRows.filter(function(r){ return r.marca===m && r.data===d && r.is_oos; }).length;
    });
  });

  // Formata labels: mm/aaaa só quando "all" com dados de múltiplos meses; senão dd/mm
  var _uniqueMonths = new Set(allDates.map(function(d){ var p=d.split("/"); return p.length===3?p[1]+"/"+p[2]:d; }));
  var _multiMonth = (_activePeriod === "all") && _uniqueMonths.size > 1;
  var days = allDates.map(function(d) {
    var p = d.split("/");
    if (p.length !== 3) return d;
    return _multiMonth ? p[1] + "/" + p[2] : p[0] + "/" + p[1];
  });

  var datasets = Object.keys(brandData).map(function(m) {
    return {
      label: m,
      data: brandData[m],
      borderColor: BRAND_COLORS_CHART[m] || "#666",
      backgroundColor: (BRAND_COLORS_CHART[m] || "#666") + "22",
      borderWidth: 2.5,
      pointRadius: 4,
      pointHoverRadius: 6,
      tension: 0.3,
      fill: false
    };
  });

  var ctx = document.getElementById("oosChart").getContext("2d");
  if (_chart) _chart.destroy();
  _chart = new Chart(ctx, {
    type: "line",
    data: { labels: days, datasets: datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: {
          position: "top",
          labels: {
            font: { family: "DM Sans", size: 12, weight: "600" },
            usePointStyle: true,
            pointStyleWidth: 10,
            padding: 16
          }
        },
        tooltip: {
          backgroundColor: "#fff",
          borderColor: "#e4e9f0",
          borderWidth: 1,
          titleColor: "#0f1923",
          bodyColor: "#6b7a8d",
          titleFont: { family: "DM Sans", weight: "700" },
          bodyFont: { family: "DM Sans" },
          padding: 12,
          boxPadding: 4,
          callbacks: {
            afterBody: function(items) {
              if (!items.length) return "";
              var dayLabel = items[0].label;
              var total = allRows.filter(function(r){ return r.data === dayLabel; }).length;
              var totalOos = items.reduce(function(s,i){ return s + i.parsed.y; }, 0);
              if (!total) return "";
              return ["\n" + totalOos + " OOS de " + total + " (" + (totalOos/total*100).toFixed(1) + "%)"];
            }
          }
        }
      },
      scales: {
        x: {
          grid: { color: "#f0f4f8" },
          ticks: { font: { family: "DM Mono", size: 11 }, color: "#6b7a8d" }
        },
        y: {
          beginAtZero: true,
          ticks: {
            stepSize: 1,
            font: { family: "DM Mono", size: 11 },
            color: "#6b7a8d"
          },
          grid: { color: "#f0f4f8" }
        }
      }
    }
  });
}

/* ── CHART RESET ── */
document.getElementById("btnChartReset").addEventListener("click", function() {
  if (!_chart) return;
  _activePeriod = "all";
  _selectedYear = null;
  _selectedMonth = null;
  
  // Reset botões
  document.querySelectorAll(".period-btn").forEach(function(b) {
    b.classList.remove("active");
  });
  document.querySelector(".period-btn[data-period='all']").classList.add("active");
  
  // Reset dropdowns
  document.getElementById("periodYear").value = "";
  document.getElementById("periodMonth").value = "";
  
  // Mostra todas as marcas
  var count = _chart.data.datasets.length;
  for (var i = 0; i < count; i++) {
    _chart.show(i);
  }
  
  renderChart();
});

/* ── PERIOD FILTERS ── */
var _selectedYear  = new Date().getFullYear();
var _selectedMonth = new Date().getMonth() + 1;
document.getElementById("periodYear").value  = _selectedYear.toString();
document.getElementById("periodMonth").value = ("0" + _selectedMonth).slice(-2);

document.querySelectorAll(".period-btn").forEach(function(btn) {
  btn.addEventListener("click", function() {
    var period = this.getAttribute("data-period");
    _activePeriod = period;
    _selectedYear = null;
    _selectedMonth = null;
    
    // Reset dropdowns
    document.getElementById("periodYear").value = "";
    document.getElementById("periodMonth").value = "";
    
    // Atualiza classes dos botões
    document.querySelectorAll(".period-btn").forEach(function(b) {
      b.classList.remove("active");
    });
    this.classList.add("active");
    
    // Re-renderiza gráfico com período filtrado
    renderChart();
  });
});

// Dropdown ANO
document.getElementById("periodYear").addEventListener("change", function() {
  var year = this.value;
  if (!year) return;
  
  _activePeriod = "year";
  _selectedYear = parseInt(year);
  _selectedMonth = null;
  
  // Reset botões e mês
  document.querySelectorAll(".period-btn").forEach(function(b) {
    b.classList.remove("active");
  });
  document.getElementById("periodMonth").value = "";
  
  renderChart();
});

// Dropdown MÊS
document.getElementById("periodMonth").addEventListener("change", function() {
  var monthVal = this.value;
  if (!monthVal) return;
  
  // Mês agora é só o número (01-12), combina com ano selecionado
  var yearSelect = document.getElementById("periodYear");
  var yearVal = yearSelect.value;
  
  // Se não há ano selecionado, usa ano atual
  if (!yearVal) {
    var currentYear = new Date().getFullYear();
    yearVal = currentYear.toString();
    yearSelect.value = yearVal;
  }
  
  _activePeriod = "month";
  _selectedYear = parseInt(yearVal);
  _selectedMonth = parseInt(monthVal);
  
  // Reset botões
  document.querySelectorAll(".period-btn").forEach(function(b) {
    b.classList.remove("active");
  });
  
  renderChart();
});

/* ── BOOT ── */
populateGlobal();
render();
renderChart();


/* ── COLLAPSIBLE SECTIONS ── */
(function(){
  [['secResumo','chevResumo','bodyResumo'],
   ['secMarca','chevMarca','bodyMarca'],
   ['secFiltros','chevFiltros','bodyFiltros'],
   ['secOos','chevOos','bodyOos'],
   ['secRadar','chevRadar','bodyRadar'],
   ['secProdutos','chevProdutos','bodyProdutos'],
  ].forEach(function(s){
    var sec=document.getElementById(s[0]),chev=document.getElementById(s[1]),body=document.getElementById(s[2]);
    if(!sec||!body) return;
    sec.addEventListener('click',function(){
      var open=body.classList.toggle('open');
      if(chev) chev.classList.toggle('open',open);
    });
  });
})();

/* ── STORE TAB EVENTS ── */
["tabAll","tabWells","tabPC"].forEach(function(id) {
  var el = $(id);
  if (!el) return;
  el.addEventListener("click", function() {
    _activeStore = id === "tabAll" ? "all" : id === "tabWells" ? "wells" : "pc";
    document.querySelectorAll(".store-tab").forEach(function(b) { b.classList.remove("active"); });
    el.classList.add("active");
    $("fMarca").value = "";
    page = 1;
    render();
    window.renderRadar();
  });
});

/* ── DISCOUNT RADAR ── */
(function(){
  var BANDS=[
    {key:'<=20',label:'≤ 20%',color:'#D0E8FF',test:function(d){return d&&d!=='Sem desconto'&&parseInt(d)<=20;}},
    {key:'25',  label:'25%',  color:'#79BAEF',test:function(d){return d==='25%';}},
    {key:'2830',label:'28–30%',color:'#2E7FD9',test:function(d){var n=parseInt(d);return !isNaN(n)&&n>=28&&n<=30;}},
    {key:'35p', label:'35%+', color:'#1249A0',test:function(d){var n=parseInt(d);return !isNaN(n)&&n>=35&&n<50;}},
    {key:'50',  label:'50%',  color:'#071D47',test:function(d){return d==='50%';}},
  ];
  var BRANDS=['Avène','Ducray','Klorane','René Furterer','A-Derma','Dexeryl','Oral Care'];
  var _discChart=null;

  /* Helpers definidos uma vez */
  window._radarBandFilter=null;
  function showRadarBtn(show){var b=document.getElementById('btnRadarReset');if(b)b.style.display=show?'':'none';}
  function clearAllFilters(){
    _newOosFilter=false;_semDescontoFilter=false;window._radarBandFilter=null;
    $("fMarca").value="";$("fEstado").value="";$("fSearch").value="";$("fSoOos").checked=false;
    if(window._fpFrom)window._fpFrom.clear();if(window._fpTo)window._fpTo.clear();
    $("cf-marca").value="";$("cf-titulo").value="";$("cf-variante").value="";
    $("cf-estado").value="";$("cf-ref").value="";
    $("cf-desconto").value="";$("cf-oos-dias").value="";
  }
  function applyRadarFilter(band){
    clearAllFilters();
    window._radarBandFilter={test:band.test};
    document.querySelectorAll('.radar-band-clickable').forEach(function(r){r.classList.remove('radar-band-active');});
    var m=document.querySelector('.radar-band-clickable[data-band-key="'+band.key+'"]');
    if(m)m.classList.add('radar-band-active');
    showRadarBtn(true); page=1; renderTable(); setTimeout(scrollToTable,150);
  }
  window._clearRadarFilter=function(){
    window._radarBandFilter=null;_semDescontoFilter=false;
    document.querySelectorAll('.radar-band-clickable').forEach(function(r){r.classList.remove('radar-band-active');});
    showRadarBtn(false);
  };
  document.getElementById('btnRadarReset').addEventListener('click',function(){
    window._clearRadarFilter(); page=1; renderTable();
  });

  window.renderRadar=function(){
    /* Filtra DATA pelo separador activo */
    var rows=DATA.filter(function(r){
      if(_activeStore==="wells") return r.loja!=="P&C";
      if(_activeStore==="pc") return r.loja==="P&C";
      return true;
    });

    var bandTotals={},bandByBrand={},semDesc={},totalRows=0;
    BANDS.forEach(function(b){bandTotals[b.key]=0;bandByBrand[b.key]={};});
    rows.forEach(function(r){
      totalRows++;
      var d=r.desconto||'';
      if(!d||d==='Sem desconto'){semDesc[r.marca]=(semDesc[r.marca]||0)+1;return;}
      BANDS.forEach(function(b){if(b.test(d)){bandTotals[b.key]++;bandByBrand[b.key][r.marca]=(bandByBrand[b.key][r.marca]||0)+1;}});
    });

    document.getElementById('rdcTotal').textContent=totalRows;
    document.getElementById('discTotal').textContent=totalRows+' produtos analisados';

    var maxBand=Math.max.apply(null,BANDS.map(function(b){return bandTotals[b.key]||0;}));

    /* Donut — destroi instância anterior */
    if(_discChart){_discChart.destroy();_discChart=null;}
    var ctx=document.getElementById('discChart').getContext('2d');
    _discChart=new Chart(ctx,{
      type:'doughnut',
      data:{labels:BANDS.map(function(b){return b.label;}),datasets:[{
        data:BANDS.map(function(b){return bandTotals[b.key];}),
        backgroundColor:BANDS.map(function(b){return b.color;}),
        borderWidth:3,borderColor:'#fff',hoverOffset:10
      }]},
      options:{responsive:true,maintainAspectRatio:false,cutout:'68%',
        plugins:{legend:{display:false},tooltip:{callbacks:{label:function(c){
          var pct=((c.parsed/totalRows)*100).toFixed(1);
          return ' '+c.parsed+' produtos ('+pct+'%) — clica para filtrar';
        }}}},
        onClick:function(evt,elements){
          if(!elements||!elements.length)return;
          applyRadarFilter(BANDS[elements[0].index]);
        }
      }
    });

    /* Band rows — limpa e repopula */
    var bandsEl=document.getElementById('radarBands');
    bandsEl.innerHTML='';
    BANDS.forEach(function(b){
      var n=bandTotals[b.key],pct=((n/totalRows)*100).toFixed(1);
      var fillW=maxBand>0?Math.round((n/maxBand)*100):0;
      var pips=BRANDS.filter(function(br){return bandByBrand[b.key][br]>0;})
        .map(function(br){return '<span class="radar-brand-pip"><span style="width:6px;height:6px;border-radius:50%;background:'+(BRAND_COLORS[br]||'#999')+';display:inline-block"></span>'+br.split(' ')[0]+'&nbsp;'+bandByBrand[b.key][br]+'</span>';}).join('');
      var row=document.createElement('div');
      row.className='radar-band-row radar-band-clickable';
      row.setAttribute('data-band-key',b.key);
      row.innerHTML='<div class="radar-band-label"><div class="radar-band-dot" style="background:'+b.color+'"></div>'+b.label+'</div>'+
        '<div><div class="radar-track"><div class="radar-fill" style="width:'+fillW+'%;background:'+b.color+'"></div></div>'+(pips?'<div class="radar-brand-mini">'+pips+'</div>':'')+
        '</div><div class="radar-count" style="color:'+b.color+'">'+n+'</div>'+
        '<div class="radar-pct">'+pct+'% <span class="radar-filter-icon">↗</span></div>';
      (function(band,rowEl){rowEl.addEventListener('click',function(){applyRadarFilter(band);});})(b,row);
      bandsEl.appendChild(row);
    });

    /* Sem desconto callout — limpa e repopula */
    var semDescTotal=Object.values(semDesc).reduce(function(a,v){return a+v;},0);
    var semDescEl=document.getElementById('radarSemDesc');
    semDescEl.innerHTML='';
    if(semDescTotal>0){
      var brandList=Object.entries(semDesc).sort(function(a,b){return b[1]-a[1];})
        .map(function(e){return e[0].split(' ')[0]+' ('+e[1]+')';}).join(' · ');
      semDescEl.innerHTML='<div class="radar-sem-desc" style="cursor:pointer" title="Clica para filtrar produtos sem desconto">'+
        '<span class="radar-sem-desc-icon">⚠️</span>'+
        '<div><div class="radar-sem-desc-text">'+semDescTotal+' produtos sem desconto registado</div>'+
        '<div class="radar-sem-desc-brands">'+brandList+'</div></div>'+
        '<span style="margin-left:auto;font-size:11px;color:#e65100;font-weight:700">Ver na tabela ↗</span></div>';
      semDescEl.firstChild.addEventListener('click',function(){
        clearAllFilters(); _semDescontoFilter=true;
        showRadarBtn(true); page=1; renderTable(); setTimeout(scrollToTable,150);
      });
    }
  };

  /* Render inicial */
  window.renderRadar();
})();

})();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# HISTORICO CSV
# ---------------------------------------------------------------------------

HISTORICO_CSV = os.path.join("logs", "historico.csv")
HISTORICO_COLS = ["run_id", "data", "loja", "marca", "titulo", "url", "nome_variante",
                  "is_oos", "oos_desde", "ref_produto", "desconto"]

_LOJAS_SET = {"Wells", "P&C"}

def load_historico() -> list:
    """Carrega todo o historico de runs anteriores."""
    if not os.path.exists(HISTORICO_CSV):
        return []
    rows = []
    try:
        # Detecta se o cabeçalho CSV tem coluna "loja" (formato novo, 11 cols)
        # ou não (formato antigo, 10 cols sem loja).
        # Bug: quando "loja" foi adicionado a HISTORICO_COLS mas o CSV já existia,
        # o cabeçalho ficou antigo e as colunas das rows novas ficaram shifted:
        #   col2 = loja value → lido como "marca"
        #   col3 = marca value → lido como "titulo"  etc.
        with open(HISTORICO_CSV, "r", encoding="utf-8") as f:
            header_line = f.readline().strip()
        has_loja_col = "loja" in header_line.split(";")

        with open(HISTORICO_CSV, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f, delimiter=";")
            for row in reader:
                if not has_loja_col:
                    if row.get("marca") in _LOJAS_SET:
                        # Row com 11 cols mas lida com cabeçalho de 10 cols:
                        # repor o shift — cada campo está 1 posição à direita do esperado.
                        row["loja"]          = row.get("marca", "")
                        row["marca"]         = row.get("titulo", "")
                        row["titulo"]        = row.get("url", "")
                        row["url"]           = row.get("nome_variante", "")
                        row["nome_variante"] = row.get("is_oos", "")
                        real_is_oos          = row.get("oos_desde", "0")
                        row["oos_desde"]     = row.get("ref_produto", "")
                        row["ref_produto"]   = row.get("desconto", "")
                        # campo extra (desconto real) fica no restkey None do DictReader
                        extra = row.get(None, [])
                        row["desconto"]      = extra[0] if isinstance(extra, list) and extra else (extra or "")
                        row["is_oos"]        = real_is_oos
                    else:
                        # Row antiga (10 cols, só Wells): colunas correctas, apenas falta loja
                        row.setdefault("loja", "Wells")
                v = row.get("is_oos", "0")
                row["is_oos"] = 1 if str(v).strip().lower() in ("1", "true") else 0
                rows.append(row)
    except Exception as e:
        log(f"Erro ao carregar historico: {e}")
    return rows

def save_historico(run_id: str, all_rows: list, oos_desde_map: dict) -> None:
    """
    Acrescenta os dados da run actual ao historico CSV.
    oos_desde_map: {ref_produto -> data_oos_desde} calculado a partir do historico.
    """
    os.makedirs("logs", exist_ok=True)
    exists = os.path.exists(HISTORICO_CSV)
    try:
        with open(HISTORICO_CSV, "a", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=HISTORICO_COLS, delimiter=";",
                                    extrasaction="ignore")
            if not exists:
                writer.writeheader()
            for row in all_rows:
                ref = row.get("ref_produto", "")
                # oos_desde: data em que ficou OOS pela primeira vez
                oos_desde = ""
                if row.get("is_oos"):
                    oos_desde = oos_desde_map.get(ref, row.get("data", ""))
                writer.writerow({
                    "run_id":        run_id,
                    "data":          row.get("data", ""),
                    "loja":          row.get("loja", "Wells"),
                    "marca":         row.get("marca", ""),
                    "titulo":        row.get("titulo", ""),
                    "url":           row.get("url", ""),
                    "nome_variante": row.get("nome_variante", ""),
                    "is_oos":        row.get("is_oos", 0),
                    "oos_desde":     oos_desde,
                    "ref_produto":   ref,
                    "desconto":      row.get("desconto", "Sem desconto"),
                })
    except Exception as e:
        log(f"Erro ao guardar historico: {e}")

def build_oos_desde_map(historico: list, current_rows: list, results_list: list) -> dict:
    """
    Calcula para cada REF a data em que ficou OOS.
    CORRIGIDO: Reinicia contagem quando produto volta a ficar OOS após recuperação.
    Retorna dict {ref_produto -> data_oos_desde}.
    """
    oos_desde = {}
    
    # Coleta URLs que mudaram para OOS HOJE (newly_oos de cada marca)
    newly_oos_urls = set()
    for r in results_list:
        newly_oos_urls.update(r.get("new_oos", []))
    
    # Para produtos que mudaram para OOS HOJE, usa data de HOJE
    for row in current_rows:
        if row.get("is_oos"):
            ref = row.get("ref_produto", "")
            url = row.get("url", "")
            
            if url in newly_oos_urls:
                # Produto mudou para OOS HOJE (In Stock → OOS)
                oos_desde[ref] = row.get("data", "")
            elif ref and ref not in oos_desde:
                # Produto continua OOS, usa data do histórico
                # Procura no histórico
                hist_date = None
                for hist_row in historico:
                    if hist_row.get("ref_produto") == ref and int(hist_row.get("is_oos", 0)):
                        hist_date = hist_row.get("oos_desde") or hist_row.get("data", "")
                        break
                oos_desde[ref] = hist_date or row.get("data", "")
    
    return oos_desde


# ---------------------------------------------------------------------------
# GERAR DASHBOARD
# ---------------------------------------------------------------------------

def generate_dashboard(run_id: str, all_results: list, output_path: str,
                        historico: list = None) -> None:
    """
    Gera o dashboard HTML com historico completo.
    historico: lista de rows de runs anteriores (do CSV).
    """
    import json as _json
    from datetime import datetime as _dt

    # Agrega rows da run actual
    current_rows = []
    new_oos_count = 0
    for r in all_results:
        current_rows.extend(r.get("all_rows", []))
        new_oos_count += r.get("new_count", 0)
    
    # Adiciona rows de produtos REMOVIDOS (para aparecerem na tabela ao clicar KPI)
    for r in all_results:
        current_rows.extend(r.get("removed_rows", []))

    # Calcula oos_desde (com informação de newly_oos)
    hist = historico or []
    oos_desde_map = build_oos_desde_map(hist, current_rows, all_results)

    # Enriquece rows com oos_desde
    for row in current_rows:
        ref = row.get("ref_produto", "")
        row["oos_desde"] = oos_desde_map.get(ref, "") if row.get("is_oos") else ""

    # Historico de runs anteriores para o grafico (agrupado por data+marca)
    # Inclui runs anteriores + run actual
    # Deduplicação: por data+marca+run_id, mantém só a ultima run de cada DIA
    # Evita que runs de teste do mesmo dia distorçam o historico
    _seen_run_dates = {}
    for row in hist:
        day = row.get("data", "")
        run = row.get("run_id", "")
        key = (day, row.get("marca", ""))
        # guarda o run_id mais recente por (dia, marca)
        if key not in _seen_run_dates or run > _seen_run_dates[key]:
            _seen_run_dates[key] = run

    # Data de hoje (para excluir do hist e usar só current_rows para hoje)
    _today_str = _dt.now().strftime("%d/%m/%Y")

    hist_rows_for_chart = []
    for row in hist:
        day = row.get("data", "")
        run = row.get("run_id", "")
        key = (day, row.get("marca", ""))
        # Exclui o dia de hoje do histórico — será adicionado via current_rows
        if day == _today_str:
            continue
        # só inclui rows da ultima run desse dia
        if _seen_run_dates.get(key) == run:
            hist_rows_for_chart.append({
                "data":  day,
                "loja":  row.get("loja", "Wells"),
                "marca": row.get("marca", ""),
                "is_oos": int(row.get("is_oos", 0)),
            })
    # Adiciona run actual (hoje) — única fonte para hoje, sem duplicação
    for row in current_rows:
        hist_rows_for_chart.append({
            "data":  row.get("data", ""),
            "loja":  row.get("loja", "Wells"),
            "marca": row.get("marca", ""),
            "is_oos": 1 if row.get("is_oos") else 0,
        })

    run_date = _dt.now().strftime("%d/%m/%Y %H:%M")

    # Normaliza is_oos para int 0/1 em current_rows antes do JSON
    # (pc_scraper devolve bool Python True/False; Wells devolve int 0/1;
    #  o JS usa String(r.is_oos)==="1" nos filtros, que falha com "true")
    for row in current_rows:
        row["is_oos"] = 1 if row.get("is_oos") else 0

    # NOVOS OOS: apenas produtos com oos_desde == data de hoje (0 dias de OOS)
    _today_str = _dt.now().strftime("%d/%m/%Y")
    new_oos_refs = []
    for row in current_rows:
        if row.get("is_oos") and row.get("oos_desde") == _today_str:
            ref = row.get("ref_produto", "")
            if ref and ref not in new_oos_refs:
                new_oos_refs.append(ref)

    # Garantia: NEW_OOS_COUNT deve ser sempre igual a len(new_oos_refs)
    new_oos_count = len(new_oos_refs)

    html = DASHBOARD_TEMPLATE
    html = html.replace("{{DATA_JSON}}",         _json.dumps(current_rows,       ensure_ascii=False))
    # Calcula RECOVERED e REMOVED do consolidado
    recovered_count = sum(r.get("recovered_count_variants", 0) for r in all_results)
    removed_count = sum(r.get("removed_count", 0) for r in all_results)
    
    # Coleta URLs/REFs de recuperados e removidos
    recovered_refs = []
    removed_refs = []
    for r in all_results:
        recovered_refs.extend(r.get("recovered", []))
        removed_refs.extend(r.get("removed", []))
    
    html = html.replace("{{NEW_OOS_COUNT}}",     str(new_oos_count))
    html = html.replace("{{RECOVERED_COUNT}}",   str(recovered_count))
    html = html.replace("{{REMOVED_COUNT}}",     str(removed_count))
    html = html.replace("{{NEW_OOS_REFS_JSON}}", _json.dumps(new_oos_refs,        ensure_ascii=False))
    html = html.replace("{{RECOVERED_REFS_JSON}}", _json.dumps(recovered_refs,    ensure_ascii=False))
    html = html.replace("{{REMOVED_REFS_JSON}}", _json.dumps(removed_refs,        ensure_ascii=False))
    html = html.replace("{{RUN_ID}}",            run_id)
    html = html.replace("{{RUN_DATE}}",          run_date)
    html = html.replace("{{HIST_JSON}}",         _json.dumps(hist_rows_for_chart, ensure_ascii=False))

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)


PC_HEADERS = ["data", "loja", "marca", "titulo", "url", "nome_variante", "is_oos", "ref_produto", "desconto"]


def build_pc_result(pc_rows: list, historico: list) -> dict:
    """Converte rows brutas do pc_scraper no mesmo formato que run_brand() devolve."""
    import unicodedata as _ud
    label = "P&C"
    state_dir = "state_pc_all"
    logs_dir = "logs_pc_all"
    ensure_dirs(state_dir, logs_dir)

    known_json = os.path.join(state_dir, "oos_known.json")
    run_id = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%SZ")
    _date_fmt = datetime.now(timezone.utc).strftime("%d_%m_%Y")
    run_xlsx = os.path.join(logs_dir, f"PC_{_date_fmt}.xlsx")

    known_oos: set = load_known(known_json)

    # Reconstruir estado anterior do histórico se não há state local
    prev_state: dict = {}
    if not prev_state and historico:
        for row in historico:
            url = row.get("url", "")
            loja = row.get("loja", "")
            if url and loja == "P&C":
                prev_state[url] = bool(int(row.get("is_oos", 0)))

    first_run = (len(known_oos) == 0 and len(prev_state) == 0)

    current_oos: set = {r["url"] for r in pc_rows if r["is_oos"]}
    all_urls: set = {r["url"] for r in pc_rows}

    new_oos = current_oos - known_oos if not first_run else set()
    recovered = (known_oos - current_oos) if not first_run else set()
    removed_rows: list = []

    save_known(known_json, current_oos)

    # Marca rows com newly_oos para o dashboard
    for row in pc_rows:
        row["newly_oos"] = row["url"] in new_oos

    # XLSX
    rows_list = [[r[h] for h in PC_HEADERS] for r in pc_rows]
    write_xlsx(run_xlsx, PC_HEADERS, rows_list)

    current_count_variants = len(current_oos)
    new_oos_count_variants = len(new_oos)
    recovered_count_variants = len(recovered)

    log(f"P&C: {len(pc_rows)} produtos | {current_count_variants} OOS | {new_oos_count_variants} NOVOS | {recovered_count_variants} RECUPERADOS")

    return {
        "label":                    label,
        "first_run":                first_run,
        "run_xlsx":                 run_xlsx,
        "current_oos":              sorted(current_oos),
        "new_oos":                  sorted(new_oos),
        "recovered":                sorted(recovered),
        "removed":                  [],
        "removed_rows":             removed_rows,
        "products_found":           len(all_urls),
        "total_detected":           len(pc_rows),
        "current_count":            len(current_oos),
        "current_count_variants":   current_count_variants,
        "new_count":                len(new_oos),
        "new_count_variants":       new_oos_count_variants,
        "recovered_count":          len(recovered),
        "recovered_count_variants": recovered_count_variants,
        "removed_count":            0,
        "all_rows":                 pc_rows,
        "headers":                  PC_HEADERS,
    }


def main() -> None:
    master_run_id = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%SZ")
    log(f"START | Monitor OOS Wells (ALL BRANDS) | run={master_run_id}")

    # Carrega historico anterior
    historico = load_historico()
    log(f"Historico carregado: {len(historico)} linhas de runs anteriores.")

    results: List[Dict] = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1365, "height": 900})
        page    = context.new_page()
        page.set_default_navigation_timeout(60000)  # 60s para lidar com wells.pt lento

        for cfg in BRANDS:
            try:
                results.append(run_brand(cfg, page, historico))
            except Exception as e:
                log(f"ERRO em {cfg.get('label', cfg.get('key'))}: {e}")
                log(f"A tentar novamente {cfg.get('label', cfg.get('key'))}...")
                try:
                    results.append(run_brand(cfg, page, historico))
                except Exception as e2:
                    log(f"ERRO (2ª tentativa) em {cfg.get('label', cfg.get('key'))}: {e2}")

        browser.close()

    # --- P&C (Perfumes & Companhia) — não precisa Playwright ---
    try:
        from pc_scraper import run_pc_all
        pc_rows = run_pc_all()
        if pc_rows:
            pc_result = build_pc_result(pc_rows, historico)
            results.append(pc_result)
        else:
            log("P&C: nenhuma row obtida (scraper falhou ou site inacessível)")
    except Exception as e:
        log(f"ERRO P&C: {e}")

    send        = False
    attachments: List[str] = []
    body_lines: List[str]  = ["Resumo da execucao (Wells & P&C Online):\n"]

    # Agrega todas as rows da run actual
    all_current_rows = []
    for r in results:
        all_current_rows.extend(r.get("all_rows", []))

    for r in results:
        body_lines.append(
            f"- {r['label']}: produtos={r['products_found']}" +
            (f" | total_detetado={r['total_detected']}" if r['total_detected'] else "")
        )
        body_lines.append(f"  OOS atuais: {r.get('current_count_variants', r['current_count'])} variantes | NOVOS OOS: {r.get('new_count_variants', r['new_count'])} variantes\n")

        if r["first_run"] and r.get("current_count_variants", r["current_count"]) > 0:
            send = True
            attachments.append(r["run_xlsx"])
            body_lines.append(f"{r['label']} - Primeira execucao (OOS atuais):")
            body_lines.extend([f"  {u}" for u in r["current_oos"]])
            body_lines.append("")
        elif (not r["first_run"]) and r["new_count"] > 0:
            send = True
            attachments.append(r["run_xlsx"])
            body_lines.append(f"{r['label']} - Novos OOS desde a ultima execucao:")
            body_lines.extend([f"  {u}" for u in r["new_oos"]])
            body_lines.append("")

    # Calcula oos_desde e guarda historico
    oos_desde_map = build_oos_desde_map(historico, all_current_rows, results)
    save_historico(master_run_id, all_current_rows, oos_desde_map)
    log(f"Historico actualizado: {HISTORICO_CSV}")

    # Cria Excel consolidado (sempre)
    _master_date_fmt = datetime.now(timezone.utc).strftime("%d_%m_%Y")
    consolidated_xlsx      = os.path.join("logs_consolidado", f"Consolidado_{_master_date_fmt}.xlsx")
    consolidated_xlsx_repo = os.path.join("logs", f"Consolidado_{_master_date_fmt}.xlsx")
    os.makedirs("logs_consolidado", exist_ok=True)
    os.makedirs("logs", exist_ok=True)
    try:
        write_consolidated_xlsx(consolidated_xlsx, results)
        import shutil as _shutil2
        _shutil2.copy2(consolidated_xlsx, consolidated_xlsx_repo)
        log(f"Excel consolidado: {consolidated_xlsx}")
        log(f"Excel consolidado copiado para repo: {consolidated_xlsx_repo}")
    except Exception as e:
        log(f"Erro ao criar Excel consolidado: {e}")
        consolidated_xlsx = None

    # Gera dashboard HTML com historico completo
    # Guarda em logs_consolidado (local) e em logs/ (para push ao GitHub)
    os.makedirs("logs", exist_ok=True)
    dashboard_html       = os.path.join("logs_consolidado", f"Dashboard_Update_{_master_date_fmt}.html")
    dashboard_html_repo        = os.path.join("logs", "dashboard_latest.html")
    dashboard_html_repo_dated  = os.path.join("logs", f"Dashboard_Update_{_master_date_fmt}.html")
    try:
        generate_dashboard(master_run_id, results, dashboard_html, historico=historico)
        # Copia para logs/ — ficheiro fixo (latest) e ficheiro datado
        import shutil as _shutil
        os.makedirs("logs", exist_ok=True)
        _shutil.copy2(dashboard_html, dashboard_html_repo)
        _shutil.copy2(dashboard_html, dashboard_html_repo_dated)
        log(f"Dashboard HTML: {dashboard_html}")
        log(f"Dashboard copiado para repo: {dashboard_html_repo} e {dashboard_html_repo_dated}")
    except Exception as e:
        log(f"Erro ao criar dashboard: {e}")
        dashboard_html = None

    if send:
        import shutil, tempfile
        subject = f"Wells & P&C Online - Daily Monitor Update — {datetime.now(timezone.utc).strftime('%d/%m/%Y')}"
        body    = "\n".join(body_lines).strip() + "\n\nRelatorios Excel e Dashboard em anexo."
        uniq = []
        tmp_copies = []
        _tmp_dir = tempfile.mkdtemp()
        # 1. Excel consolidado — nome correcto (ex: Consolidado_26_02_2026.xlsx)
        if consolidated_xlsx and os.path.exists(consolidated_xlsx):
            try:
                dest = os.path.join(_tmp_dir, os.path.basename(consolidated_xlsx))
                shutil.copy2(consolidated_xlsx, dest)
                uniq.append(dest); tmp_copies.append(dest)
            except Exception:
                uniq.append(consolidated_xlsx)
        # 2. Dashboard HTML — nome correcto (ex: Dashboard_Update_26_02_2026.html)
        if dashboard_html and os.path.exists(dashboard_html):
            try:
                dest = os.path.join(_tmp_dir, os.path.basename(dashboard_html))
                shutil.copy2(dashboard_html, dest)
                uniq.append(dest); tmp_copies.append(dest)
            except Exception:
                uniq.append(dashboard_html)
        # 3. Excel por marca — agrupa num ZIP (ex: Todas_as_marcas_individuais_26_02_2026.zip)
        seen = set()
        _marca_files = []
        for a in attachments:
            if a and a not in seen and os.path.exists(a):
                seen.add(a)
                _marca_files.append(a)
        if _marca_files:
            import zipfile as _zf
            _zip_name = f"todas_as_marcas_individuais_{_master_date_fmt}.zip"
            _zip_path = os.path.join(_tmp_dir, _zip_name)
            try:
                with _zf.ZipFile(_zip_path, "w", _zf.ZIP_DEFLATED) as _zzip:
                    for a in _marca_files:
                        _zzip.write(a, os.path.basename(a))
                uniq.append(_zip_path); tmp_copies.append(_zip_path)
            except Exception as _ze:
                log(f"Erro ao criar ZIP das marcas: {_ze}")
                for a in _marca_files:
                    uniq.append(a)
        send_email(subject, body, uniq)
        log(f"Email enviado com {len(uniq)} anexo(s): Excel + Dashboard.")
        try: shutil.rmtree(_tmp_dir, ignore_errors=True)
        except Exception: pass
    else:
        log("Sem novos OOS (em todas as marcas) - nenhum email enviado.")

    log("END | Monitor OOS Wells (ALL BRANDS)")


if __name__ == "__main__":
    main()
