#!/usr/bin/env python3
"""
Wireframe Analyzer
==================
Takes a wireframe (HTML file OR image: png/jpg) and produces an Excel spec
sheet with, for every detected visual/section on the page:
  - x, y, width, height (px)
  - background color (hex)
  - font family, font size, font color (hex)  [HTML: exact. Image: estimated]
Also flags overflow risk (visuals that exceed the canvas / would cause
scrolling or squeezing in Power BI).

USAGE:
    python3 wireframe_analyzer.py <input.html|input.png> [--canvas 1280x720] [--out report.xlsx]

Requirements (already satisfied in most environments):
    pip install playwright beautifulsoup4 opencv-python pillow pytesseract openpyxl
    playwright install chromium
    (Linux) apt install tesseract-ocr
"""

import sys
import os
import json
import argparse
import statistics
from collections import Counter

# ----------------------------------------------------------------------------
# Shared helpers
# ----------------------------------------------------------------------------

def rgb_to_hex(r, g, b):
    return "#{:02X}{:02X}{:02X}".format(int(r), int(g), int(b))


def px_to_pt(px, dpi=96):
    """Convert a CSS/image pixel measurement to points (used by PowerPoint/Power BI)."""
    return round(px * 72.0 / dpi, 1)


def hex_to_rgb(h):
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def brightness_of(hex_color):
    try:
        r, g, b = hex_to_rgb(hex_color)
        return (r * 299 + g * 587 + b * 114) / 1000
    except Exception:
        return 255


def blend_hex(hex_a, hex_b, t):
    """Blend two hex colors; t=0 -> hex_a, t=1 -> hex_b."""
    ra, ga, ba = hex_to_rgb(hex_a)
    rb, gb, bb = hex_to_rgb(hex_b)
    r = ra + (rb - ra) * t
    g = ga + (gb - ga) * t
    b = ba + (bb - ba) * t
    return rgb_to_hex(r, g, b)


# ----------------------------------------------------------------------------
# HTML analysis (exact — reads real computed CSS via a headless browser)
# ----------------------------------------------------------------------------

JS_EXTRACT = r"""
() => {
  function parseColor(c) {
    if (!c) return null;
    const m = c.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/);
    if (!m) return null;
    const alpha = m[4] === undefined ? 1 : parseFloat(m[4]);
    return { r: +m[1], g: +m[2], b: +m[3], a: alpha };
  }

  const results = [];
  const all = document.querySelectorAll('body, body *');
  const idxMap = new Map();
  all.forEach((el, i) => idxMap.set(el, i));
  all.forEach((el, idx) => {
    const rect = el.getBoundingClientRect();
    if (rect.width < 20 || rect.height < 6) return; // skip tiny/invisible elements
    const tagLower = el.tagName.toLowerCase();
    const isSvgRoot = tagLower === 'svg';
    // SVG internals are drawing primitives of ONE chart, not layout -- the
    // <svg> root is captured as a single typed visual instead
    if (!isSvgRoot && el.namespaceURI === 'http://www.w3.org/2000/svg') return;
    if (isSvgRoot) {
      const shapes = { polygon: 0, rect: 0, circle: 0, ellipse: 0, path: 0, line: 0 };
      el.querySelectorAll('polygon,rect,circle,ellipse,path,line').forEach(s => {
        const t = s.tagName.toLowerCase();
        if (shapes[t] !== undefined) shapes[t]++;
      });
      const dataRects = Array.from(el.querySelectorAll('rect'))
        .map(r => r.getBoundingClientRect()).filter(r => r.width >= 6 && r.height >= 6);
      let svgType = 'custom', svgCount = 0;
      if (shapes.polygon >= 2 && shapes.polygon >= dataRects.length) {
        svgType = 'funnel'; svgCount = shapes.polygon;
      } else if (dataRects.length >= 3) {
        const ws = dataRects.map(r => r.width).sort((a, b) => a - b);
        const hs = dataRects.map(r => r.height).sort((a, b) => a - b);
        const wVar = (ws[ws.length - 1] - ws[0]) / Math.max(1, ws[ws.length - 1]);
        const hVar = (hs[hs.length - 1] - hs[0]) / Math.max(1, hs[hs.length - 1]);
        svgType = hVar > wVar ? 'column' : 'bar'; svgCount = dataRects.length;
      } else if (shapes.circle + shapes.ellipse >= 1 && shapes.path <= 2) {
        svgType = 'pie'; svgCount = shapes.circle + shapes.ellipse;
      } else if (shapes.path >= 1) {
        svgType = 'line'; svgCount = shapes.path;
      }
      let fnt = null;
      const svgText = el.querySelector('text');
      if (svgText) fnt = window.getComputedStyle(svgText);
      results.push({
        tag: 'svg', id: el.id || null, cls: null,
        x: Math.round(rect.left + window.scrollX), y: Math.round(rect.top + window.scrollY),
        w: Math.round(rect.width), h: Math.round(rect.height),
        hasBg: false, hasBorder: false, bg: null, repeatKey: null,
        text: '', svgType: svgType, svgCount: svgCount,
        fontFamily: fnt ? fnt.fontFamily : '', fontSize: fnt ? parseFloat(fnt.fontSize) : 0,
        fontWeight: fnt ? fnt.fontWeight : '', color: fnt ? parseColor(fnt.fill || fnt.color) : null,
        borderRadius: '0px',
      });
      return;
    }
    const cs = window.getComputedStyle(el);
    const bg = parseColor(cs.backgroundColor);
    const hasBg = bg && bg.a > 0.05;
    if (rect.height < 12 && !hasBg) return; // 6-11px band: colored only (chart bars)
    const borderSides = [cs.borderTopWidth, cs.borderRightWidth, cs.borderBottomWidth, cs.borderLeftWidth]
      .map(v => parseFloat(v)).filter(v => v > 0).length;
    const hasBorder = borderSides >= 3; // a real box, not just a single divider line

    // does this element have direct (non-whitespace) text content of its own?
    let ownText = '';
    for (const node of el.childNodes) {
      if (node.nodeType === 3) ownText += node.textContent;
    }
    ownText = ownText.trim();

    if (!hasBg && !hasBorder && !ownText) return;

    const color = parseColor(cs.color);

    // repeatKey: marks this element as belonging to a run of >=3
    // structurally-identical siblings (same parent + same first class, or
    // tr/td/th/li tags) -- either the element itself is one of the run
    // ('self'), or a thin row-like ANCESTOR of it is (typical HTML tables:
    // the row divs are transparent, only the badges/cells inside them get
    // captured). Everything sharing a runKey is the DATA of one visual --
    // table rows, chart bars, list entries -- and gets collapsed into a
    // single typed visual later instead of documented item by item.
    const tagName = el.tagName.toLowerCase();
    function runKey(node) {
      const par = node.parentElement;
      if (!par) return null;
      if (node.className && typeof node.className === 'string') {
        const cls = node.className.trim().split(/\s+/)[0];
        if (cls) {
          let same = 0;
          for (const sib of par.children) {
            if (sib.classList && sib.classList.contains(cls)) same++;
          }
          if (same >= 3) return (idxMap.has(par) ? idxMap.get(par) : -1) + ':' + cls;
        }
      }
      // classless markup (inline-styled mockups): >=3 THIN same-tag sibling
      // divs are a row-run too. Thin-only, so stacks of big layout wrappers
      // or cards never merge.
      const nr = node.getBoundingClientRect();
      const classless = !((node.className || '').toString().trim());
      if (classless && nr.height >= 8 && nr.height <= 45) {
        let same = 0;
        for (const sib of par.children) {
          if (sib.tagName === node.tagName && !((sib.className || '').toString().trim())) {
            const sr = sib.getBoundingClientRect();
            if (sr.height >= 8 && sr.height <= 45) same++;
          }
        }
        if (same >= 3) return (idxMap.has(par) ? idxMap.get(par) : -1) + ':thinrun';
      }
      // metric/KPI pairs: a wrapper whose ONLY children are 2 same-class,
      // similar-size, small siblings (label+value metric pair, comparison
      // pair). Small-only and exactly-2-only, so two-column page layouts
      // never merge.
      if (node.className && typeof node.className === 'string') {
        const cls2 = node.className.trim().split(/\s+/)[0];
        if (cls2 && par.children.length === 2) {
          const sibs = Array.from(par.children);
          if (sibs.every(s => s.classList && s.classList.contains(cls2))) {
            const rA = sibs[0].getBoundingClientRect(), rB = sibs[1].getBoundingClientRect();
            const pageArea = Math.max(1, document.body.offsetWidth * document.body.offsetHeight);
            const small = r => (r.width * r.height) / pageArea <= 0.01 && r.height <= 90;
            const similar = Math.abs(rA.width - rB.width) <= 0.15 * Math.max(rA.width, rB.width) &&
                            Math.abs(rA.height - rB.height) <= 0.15 * Math.max(rA.height, rB.height);
            if (small(rA) && small(rB) && similar) {
              return (idxMap.has(par) ? idxMap.get(par) : -1) + ':@pair';
            }
          }
        }
      }
      // classless equal-width sibling groups laid out side by side or as a
      // grid (>=2 distinct x positions) = the clusters/tiles of ONE chart
      // (quarter columns of a trend, stat tiles of a distribution...).
      // Small-only (each <=3% of the page) so page sections never merge.
      if (classless && nr.width >= 40) {
        const pageArea = Math.max(1, document.body.offsetWidth * document.body.offsetHeight);
        if (nr.width <= 0.5 * document.body.offsetWidth && (nr.width * nr.height) / pageArea <= 0.03) {
          const matches = [];
          for (const sib of par.children) {
            if (sib.tagName !== node.tagName) continue;
            if (((sib.className || '').toString().trim())) continue;
            const sr = sib.getBoundingClientRect();
            if (sr.width >= 40 && Math.abs(sr.width - nr.width) <= 0.25 * nr.width &&
                (sr.width * sr.height) / pageArea <= 0.03) {
              matches.push(Math.round(sr.left));
            }
          }
          const distinctX = new Set(matches.map(x => Math.round(x / 24))).size;
          if (matches.length >= 3 && distinctX >= 2) {
            return (idxMap.has(par) ? idxMap.get(par) : -1) + ':@grp';
          }
        }
      }
      return null;
    }
    let repeatKey = null;
    // real HTML tables/lists: EVERY row/cell of one <table> (or item of one
    // <ul>/<ol>) belongs to the SAME visual -- key by the table itself, not
    // by each row, or the group fragments into one-row slivers
    if (['td', 'th', 'tr'].includes(tagName)) {
      const tbl = el.closest('table');
      if (tbl && tbl.querySelectorAll('tr').length >= 3) {
        repeatKey = (idxMap.has(tbl) ? idxMap.get(tbl) : -1) + ':table';
      }
    } else if (tagName === 'li') {
      const lst = el.closest('ul,ol');
      if (lst && lst.querySelectorAll('li').length >= 3) {
        repeatKey = (idxMap.has(lst) ? idxMap.get(lst) : -1) + ':list';
      }
    }
    if (!repeatKey) repeatKey = runKey(el);
    if (!repeatKey) {
      // inherit from a thin, row-like repeated ancestor (h<=45 or w<=60) --
      // big repeated cards (dashboard tiles) deliberately DON'T propagate,
      // so their contents stay individually documented
      let anc = el.parentElement;
      for (let d = 0; d < 4 && anc && anc !== document.body; d++, anc = anc.parentElement) {
        const at = anc.tagName.toLowerCase();
        if (['td', 'th', 'tr'].includes(at)) {
          const tbl = anc.closest('table');
          if (tbl && tbl.querySelectorAll('tr').length >= 3) {
            repeatKey = (idxMap.has(tbl) ? idxMap.get(tbl) : -1) + ':table';
            break;
          }
        }
        const ar = anc.getBoundingClientRect();
        if (ar.height <= 45 || ar.width <= 60) {
          const k = runKey(anc);
          if (k) { repeatKey = k; break; }
        } else {
          const k = runKey(anc);
          if (k && k.endsWith(':@grp')) { repeatKey = k; break; }
        }
      }
    }

    results.push({
      tag: tagName,
      id: el.id || null,
      cls: el.className && typeof el.className === 'string' ? el.className : null,
      x: Math.round(rect.left + window.scrollX),
      y: Math.round(rect.top + window.scrollY),
      w: Math.round(rect.width),
      h: Math.round(rect.height),
      hasBg, hasBorder,
      bg: hasBg ? bg : null,
      repeatKey: repeatKey,
      text: ownText.slice(0, 40),
      fontFamily: cs.fontFamily,
      fontSize: parseFloat(cs.fontSize),
      fontWeight: cs.fontWeight,
      color: color,
      borderRadius: cs.borderTopLeftRadius,
    });
  });

  return {
    page: { w: document.documentElement.scrollWidth, h: document.documentElement.scrollHeight },
    elements: results
  };
}
"""


JS_DETECT_PAGE_GROUP = r"""
() => {
  const all = document.querySelectorAll('body *');
  const classGroups = {};
  all.forEach(el => {
    if (!el.className || typeof el.className !== 'string') return;
    el.className.trim().split(/\s+/).forEach(cls => {
      if (!cls) return;
      if (!classGroups[cls]) classGroups[cls] = [];
      classGroups[cls].push(el);
    });
  });

  let best = null;
  for (const [cls, els] of Object.entries(classGroups)) {
    if (els.length < 2 || els.length > 40) continue;
    const infos = els.map(el => ({ el, display: window.getComputedStyle(el).display }));
    const visibleCount = infos.filter(i => i.display !== 'none').length;
    const hiddenCount = infos.length - visibleCount;
    // the classic toggled-page/tab pattern: several share a class, most are
    // hidden, one (or a couple) currently shown
    if (hiddenCount < 1 || visibleCount < 1 || visibleCount > 2) continue;
    if (!best || els.length > best.els.length) {
      best = { cls, els: infos.map(i => i.el) };
    }
  }
  if (!best) return null;

  function findLabel(el) {
    const id = el.id || '';
    const suffix = id.replace(/^(page|tab|view|section)[-_]?/i, '');
    const candidates = document.querySelectorAll('[onclick], [data-target], [data-page], [href]');
    for (const c of candidates) {
      const blob = (c.getAttribute('onclick') || '') + ' ' + (c.getAttribute('data-target') || '') +
                   ' ' + (c.getAttribute('data-page') || '') + ' ' + (c.getAttribute('href') || '');
      if ((suffix && blob.includes(suffix)) || (id && blob.includes(id))) {
        const t = c.innerText && c.innerText.trim();
        if (t) return t;
      }
    }
    return null;
  }

  return {
    groupClass: best.cls,
    pages: best.els.map((el, i) => ({
      index: i,
      id: el.id || null,
      label: findLabel(el) || el.id || `Page ${i + 1}`,
    })),
  };
}
"""

JS_SHOW_ONLY = r"""
(args) => {
  const { groupClass, targetId, targetIndex } = args;
  const els = Array.from(document.getElementsByClassName(groupClass));
  els.forEach((el, i) => {
    const isTarget = targetId ? (el.id === targetId) : (i === targetIndex);
    el.style.setProperty('display', isTarget ? 'block' : 'none', 'important');
  });
}
"""

JS_CLICK_NAV = r"""
(args) => {
  const { groupClass, targetId, targetIndex } = args;
  const els = Array.from(document.getElementsByClassName(groupClass));
  const target = targetId ? els.find(e => e.id === targetId) : els[targetIndex];
  if (!target) return false;
  const id = target.id || '';
  const suffix = id.replace(/^(page|tab|view|section)[-_]?/i, '');
  const candidates = document.querySelectorAll('[onclick], [data-target], [data-page], [href]');
  for (const c of candidates) {
    const blob = (c.getAttribute('onclick') || '') + ' ' + (c.getAttribute('data-target') || '') +
                 ' ' + (c.getAttribute('data-page') || '') + ' ' + (c.getAttribute('href') || '');
    if ((id && blob.includes(id)) || (suffix && suffix !== id && blob.includes(suffix))) {
      c.click();
      return true;
    }
  }
  return false;
}
"""

JS_EXPAND_TOGGLES = r"""
(args) => {
  const { groupClass, targetId, targetIndex } = args;
  const els = Array.from(document.getElementsByClassName(groupClass));
  const container = targetId ? els.find(e => e.id === targetId) : els[targetIndex];
  if (!container) return 0;

  function visCount() {
    let n = 0;
    container.querySelectorAll('*').forEach(el => {
      const r = el.getBoundingClientRect();
      if (r.width >= 20 && r.height >= 12) n++;
    });
    return n;
  }

  // small, button-sized clickables INSIDE this page only (never the shared
  // nav/tab bar, which lives outside the page container)
  const clickables = Array.from(
    container.querySelectorAll('[onclick], [data-toggle], [data-target], [aria-expanded], button, summary')
  ).filter(el => {
    const r = el.getBoundingClientRect();
    return r.width >= 10 && r.height >= 8 && r.width <= 420 && r.height <= 60;
  });

  let expanded = 0;
  for (const el of clickables.slice(0, 15)) {
    const before = visCount();
    const beforeH = document.body.offsetHeight;
    try { el.click(); } catch (e) { continue; }
    const after = visCount();
    const afterH = document.body.offsetHeight;
    if (after > before || afterH > beforeH) {
      expanded++;            // revealed new content -- keep it open
    } else if (after < before || afterH < beforeH) {
      try { el.click(); } catch (e) {}  // it collapsed/hid something -- undo
    }
  }
  return expanded;
}
"""

JS_IS_VISIBLE = r"""
(args) => {
  const { groupClass, targetId, targetIndex } = args;
  const els = Array.from(document.getElementsByClassName(groupClass));
  const target = targetId ? els.find(e => e.id === targetId) : els[targetIndex];
  if (!target) return false;
  const cs = window.getComputedStyle(target);
  if (cs.display === 'none' || cs.visibility === 'hidden') return false;
  // and every sibling page must actually be hidden, or we'd extract a mix
  return els.every(el => {
    if (el === target) return true;
    const s = window.getComputedStyle(el);
    return s.display === 'none' || s.visibility === 'hidden';
  });
}
"""


def _norm_text(t):
    return " ".join((t or "").split())


def _rows_match(a, b, tol=3):
    """Same element on two pages, allowing a few px of layout jitter.
    Text and background must still match exactly (normalized whitespace)."""
    return (abs(a["x"] - b["x"]) <= tol and abs(a["y"] - b["y"]) <= tol and
            abs(a["w"] - b["w"]) <= tol + 1 and abs(a["h"] - b["h"]) <= tol + 1 and
            a.get("bg_hex", "") == b.get("bg_hex", "") and
            _norm_text(a.get("sample_text")) == _norm_text(b.get("sample_text")))


def extract_common_elements(pages_data, canvas_w):
    """Two-tier de-duplication of elements that repeat across pages:

    Tier 1 (Common Elements section): rows that appear -- within a small
    position tolerance, same bg + same text -- on EVERY page (the shared
    header/nav/filter-bar "shell"). Removed from all pages, documented
    once, with a cropped screenshot of the shell region.

    Tier 2 (per-page notes): rows with real text that appear identically
    on SOME (>= 2) pages but not all (per-page Export buttons, By-Factory
    toggles, etc). They stay fully documented on the first page they occur
    on; later pages drop them and instead get a one-line 'Also on this
    page ...' note pointing back to where they're documented. This is what
    stops the same filter/button being redocumented on page after page."""
    if len(pages_data) < 2:
        return pages_data, None

    # ---- Tier 1: on every page (tolerance-based, not pixel-exact) ----
    common_rows = []
    removed_ids = set()  # id() of row objects to drop, per page
    for r in pages_data[0]["rows"]:
        matches = [r]
        for pd in pages_data[1:]:
            m = next((o for o in pd["rows"] if _rows_match(r, o)), None)
            if m is None:
                matches = None
                break
            matches.append(m)
        if matches:
            common_rows.append(r)
            for m in matches:
                removed_ids.add(id(m))

    for pd in pages_data:
        pd["rows"] = [r for r in pd["rows"] if id(r) not in removed_ids]

    # ---- Tier 2: identical on >= 2 pages (but not all) ----------------
    # Only rows with real text are deduped here: an empty container at the
    # same spot on two pages can legitimately hold different content, so
    # removing it would orphan that page's visuals.
    for i, pd_src in enumerate(pages_data):
        src_label = pd_src["label"] or f"Page {i + 1}"
        for r in pd_src["rows"]:
            if not _norm_text(r.get("sample_text")):
                continue
            for j in range(i + 1, len(pages_data)):
                pd_dst = pages_data[j]
                dups = [o for o in pd_dst["rows"] if _rows_match(r, o)]
                if not dups:
                    continue
                dup_ids = {id(o) for o in dups}
                pd_dst["rows"] = [o for o in pd_dst["rows"] if id(o) not in dup_ids]
                notes = pd_dst.setdefault("_dedup_map", {})
                notes.setdefault(src_label, []).append(_norm_text(r.get("sample_text"))[:30])

    for pd in pages_data:
        dmap = pd.pop("_dedup_map", None)
        if dmap:
            pd["dedup_notes"] = [
                f"Also on this page, identical to '{src}' (see there for full details): " + ", ".join(titles)
                for src, titles in dmap.items()
            ]

    if not common_rows:
        return pages_data, None

    shell_bottom = max(r["y"] + r["h"] for r in common_rows)
    shell_canvas = {"w": canvas_w, "h": shell_bottom, "detected_w": canvas_w, "detected_h": shell_bottom}

    shell_screenshot = None
    first_screenshot = pages_data[0].get("screenshot_path")
    if first_screenshot and os.path.exists(first_screenshot):
        try:
            from PIL import Image
            with Image.open(first_screenshot) as im:
                crop = im.crop((0, 0, canvas_w, min(shell_bottom + 4, im.height)))
                shell_dir = os.path.dirname(first_screenshot)
                shell_screenshot = os.path.join(shell_dir, "_common_shell.png")
                crop.save(shell_screenshot)
        except Exception:
            shell_screenshot = None

    return pages_data, {
        "label": "Common Elements  (appear identically on every page)",
        "canvas": shell_canvas,
        "rows": common_rows,
        "screenshot_path": shell_screenshot,
    }


def detect_and_extract_pages(path, canvas_w, canvas_h, work_dir=None):
    """If this HTML has multiple tab/page-style sections (several sibling
    elements sharing a class, toggled via display:none/block -- the
    standard pattern for hand-built dashboard mockups with a tab bar),
    extract each one separately. Returns a list of
    {label, canvas, rows, screenshot_path} dicts, or None if no such
    multi-page pattern was detected (caller should fall back to normal
    single-page analysis)."""
    from playwright.sync_api import sync_playwright

    abs_path = os.path.abspath(path)
    url = f"file://{abs_path}"

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": canvas_w, "height": canvas_h})
        page.on("dialog", lambda d: d.dismiss())
        page.goto(url)
        page.wait_for_timeout(300)

        group = page.evaluate(JS_DETECT_PAGE_GROUP)
        if not group or len(group["pages"]) < 2:
            browser.close()
            return None

        results = []
        for p_info in group["pages"]:
            nav_args = {
                "groupClass": group["groupClass"],
                "targetId": p_info["id"],
                "targetIndex": p_info["index"],
            }
            # Activate the page the way a real user would: click its nav
            # button, so the wireframe's own JS updates BOTH the page
            # visibility AND the tab bar's active styling (underline,
            # highlight). Direct display-toggling alone leaves the tab bar
            # stuck on whichever tab was active when the file was saved.
            clicked = page.evaluate(JS_CLICK_NAV, nav_args)
            if clicked:
                page.wait_for_timeout(200)
            if not page.evaluate(JS_IS_VISIBLE, nav_args):
                # no clickable nav found, or clicking didn't switch pages --
                # fall back to forcing visibility directly
                page.evaluate(JS_SHOW_ONLY, nav_args)
                page.wait_for_timeout(150)

            # Click expander-style controls inside this page ('[+] By
            # Factory' etc.) so tables/sections that only appear on click
            # are open -- and therefore extracted + screenshotted -- too.
            n_expanded = page.evaluate(JS_EXPAND_TOGGLES, nav_args)
            if n_expanded:
                page.wait_for_timeout(200)
                print(f"  {p_info['label']}: expanded {n_expanded} collapsible section(s) before analysis")
            if not page.evaluate(JS_IS_VISIBLE, nav_args):
                # an expander click accidentally switched/hid the page --
                # bring it back before extracting
                if not page.evaluate(JS_CLICK_NAV, nav_args):
                    page.evaluate(JS_SHOW_ONLY, nav_args)
                page.wait_for_timeout(150)

            # this page's real content may be taller than the initial probe
            # viewport (each tab can have very different content length) --
            # resize to fit so nothing gets clipped in the screenshot or
            # miscounted as "exceeds canvas"
            content_h = page.evaluate("() => document.body.offsetHeight")
            page_canvas_h = max(canvas_h, content_h + 10)
            if page_canvas_h != canvas_h:
                page.set_viewport_size({"width": canvas_w, "height": page_canvas_h})
                page.wait_for_timeout(100)

            data = page.evaluate(JS_EXTRACT)

            screenshot_path = None
            if work_dir:
                os.makedirs(work_dir, exist_ok=True)
                safe_label = "".join(c if c.isalnum() else "_" for c in p_info["label"])[:40]
                screenshot_path = os.path.join(work_dir, f"page_{p_info['index']+1}_{safe_label}.png")
                page.screenshot(path=screenshot_path)

            kept, texts = _elements_to_rows(data["elements"], canvas_w, page_canvas_h)
            canvas, rows = _build_rows_from_containers(kept, texts, canvas_w, page_canvas_h,
                                                         data["page"]["w"], data["page"]["h"])

            results.append({
                "label": p_info["label"],
                "canvas": canvas,
                "rows": rows,
                "screenshot_path": screenshot_path,
            })

            if page_canvas_h != canvas_h:
                page.set_viewport_size({"width": canvas_w, "height": canvas_h})

        browser.close()
    return results



_SVG_TYPE_MAP = {
    "funnel": ("Funnel", "Funnel"),
    "bar": ("Bar chart (horizontal)", "Clustered bar chart"),
    "column": ("Column chart (vertical)", "Clustered column chart"),
    "pie": ("Pie / Donut chart", "Pie chart or Donut chart"),
    "line": ("Line / Area chart", "Line chart or Area chart"),
    "custom": ("Custom chart (SVG)", "Closest native chart visual, or a custom visual"),
}

def collapse_repeated_rows(rows, canvas_w, canvas_h):
    """Collapse runs of repeated DOM siblings (same repeat_key) into ONE
    synthetic visual, classified as the Power BI native visual it maps to.
    A bar chart's 8 bars / a table's 12 rows are DATA, not layout -- the
    spec only needs 'a Bar chart sits here, this big, this font', not one
    entry per bar/row.

    Classification (geometry of the repeated items):
      - vertically stacked, similar heights, similar widths  -> Table rows
        (narrow stacks with text read as a List / menu instead)
      - vertically stacked, similar heights, VARYING widths, left-aligned
                                                    -> Horizontal Bar chart
      - side by side, similar widths, varying heights, bottom-aligned
                                                    -> Column chart
      - <tr>/<td>/<th>/<li> tags                    -> Table / List
    Groups that don't match a pattern (e.g. a grid of big dashboard cards
    that happen to share a class) are left fully itemized."""
    from collections import OrderedDict

    groups = OrderedDict()
    singles = []
    for r in rows:
        k = r.get("repeat_key")
        if k:
            groups.setdefault(k, []).append(r)
        else:
            singles.append(r)

    out = list(singles)
    deferred_lists = []
    next_id = max((r["id"] for r in rows), default=0)

    def _y_levels(members):
        cs = sorted(m["y"] + m["h"] / 2 for m in members)
        levels, last = 0, None
        for c in cs:
            if last is None or c - last > 8:
                levels += 1
                last = c
        return levels

    def _wordy(m):
        """True when the item carries WORD text (>=3 letters, more letters
        than digits). Chart bars are textless or numeric ('38%'); colored
        word-pills are status badges in a table -- their widths vary with
        word length, which fakes the bar-chart geometry."""
        t = _norm_text(m.get("sample_text"))
        letters = sum(c.isalpha() for c in t)
        digits = sum(c.isdigit() for c in t)
        # narrow-only: a NARROW colored item sized to its word = a badge
        # pill. A WIDE colored item is a real bar even when it carries a
        # category label -- its width is data-driven, not text-driven.
        return letters >= 3 and letters > digits and m["w"] < 140

    def _bar_subcluster(members):
        """>=3 colored, thin, left-aligned, textless/numeric items whose
        widths vary = bars."""
        from collections import defaultdict
        byx = defaultdict(list)
        for m in members:
            if 6 <= m["h"] <= 30 and m.get("bg_hex") and not _wordy(m):
                byx[round(m["x"] / 8)].append(m)
        union_w = (max(m["x"] + m["w"] for m in members) -
                   min(m["x"] for m in members)) if members else 1
        for grp in byx.values():
            if len(grp) < 3:
                continue
            hs = sorted(g["h"] for g in grp)
            ws = sorted(g["w"] for g in grp)
            # bars must DOMINATE the visual's width -- a bar chart's bars
            # span most of the plot; narrow colored cells (numbers, accents,
            # mini progress fills) inside a wide table never do
            if ws[len(ws) // 2] < 0.25 * union_w:
                continue
            if (hs[-1] - hs[0]) <= max(6, 0.35 * hs[len(hs) // 2]) and \
               (ws[-1] - ws[0]) / max(1.0, float(ws[-1])) > 0.18 and _y_levels(grp) >= 3:
                return grp
        return None

    def _column_subcluster(members):
        """>=3 thin, bottom-aligned items whose heights vary = columns."""
        from collections import defaultdict
        byb = defaultdict(list)
        for m in members:
            if m["w"] <= 50 and m.get("bg_hex") and not _wordy(m):
                byb[round((m["y"] + m["h"]) / 8)].append(m)
        for grp in byb.values():
            if len(grp) < 3:
                continue
            ws = sorted(g["w"] for g in grp)
            hs = sorted(g["h"] for g in grp)
            if (ws[-1] - ws[0]) <= 8 and (hs[-1] - hs[0]) > 0.25 * hs[-1]:
                return grp
        return None

    for key, members in groups.items():
        min_members = 2 if key.endswith(":@pair") else 3
        if len(members) < min_members:
            out.extend(members)
            continue
        if key.endswith(":@pair"):
            vtype = ("KPI / metric group", "Multi-row card")
            # fall through to the union/synthetic build below
        else:
            vtype = None

        def _text_column_count(members):
            """Distinct x-aligned columns holding >=3 WORD-text items. A data
            TABLE has 3+ such columns; a bar chart has at most 1-2 (category
            labels, maybe a status column) -- its VALUE columns are numeric
            and don't count, since numeric data labels are normal on charts.
            This is the structural tiebreaker when a table contains bar-like
            colored cells or a chart contains label text."""
            from collections import defaultdict
            cols = defaultdict(int)
            for m in members:
                t = _norm_text(m.get("sample_text"))
                letters = sum(c.isalpha() for c in t)
                digits = sum(c.isdigit() for c in t)
                if letters >= 3 and letters > digits:
                    cols[round(m["x"] / 12)] += 1
            return sum(1 for n in cols.values() if n >= 3)

        if vtype is None and _bar_subcluster(members) and _text_column_count(members) <= 2:
            vtype = ("Bar chart (horizontal)", "Clustered bar chart")
        elif vtype is None and _column_subcluster(members):
            vtype = ("Column chart (vertical)", "Clustered column chart")
        elif vtype is None and key.endswith(":@grp"):
            # a run of equal-width classless sibling groups laid out side by
            # side / in a grid = the clusters of ONE chart or stat-tile group
            vtype = ("Chart / stat group", "Clustered bar or column chart (or Multi-row card)")
        elif vtype is None:
            # table / list: repetition down the page, and no member so tall
            # that this could be a grid of real dashboard cards
            max_h = max(m["h"] for m in members)
            if max_h <= 60 and _y_levels(members) >= 3:
                ux0 = min(m["x"] for m in members)
                ux1 = max(m["x"] + m["w"] for m in members)
                uy0 = min(m["y"] for m in members)
                uy1 = max(m["y"] + m["h"] for m in members)
                if (ux1 - ux0) >= 0.3 * canvas_w:
                    vtype = ("Table", "Table / Matrix")
                elif ux1 < 0.3 * canvas_w and (uy1 - uy0) > 0.25 * canvas_h:
                    vtype = None  # tall narrow left-edge list = likely a nav menu; keep itemized
                else:
                    vtype = ("List / menu", "Table (or Buttons + bookmarks for navigation)")

        if not vtype:
            out.extend(members)
            continue

        ux0 = min(m["x"] for m in members)
        uy0 = min(m["y"] for m in members)
        ux1 = max(m["x"] + m["w"] for m in members)
        uy1 = max(m["y"] + m["h"] for m in members)
        font_src = next((m for m in members if m.get("font_family")), members[0])
        bg_src = next((m for m in members if m.get("bg_hex")), None)
        next_id += 1
        synth = {
            "id": next_id,
            "tag": "chart",
            "label": vtype[0][:30],
            "x": ux0, "y": uy0, "w": ux1 - ux0, "h": uy1 - uy0,
            "x_pt": px_to_pt(ux0), "y_pt": px_to_pt(uy0),
            "w_pt": px_to_pt(ux1 - ux0), "h_pt": px_to_pt(uy1 - uy0),
            "bg_hex": bg_src["bg_hex"] if bg_src else "",
            "font_family": font_src.get("font_family", ""),
            "font_size_px": font_src.get("font_size_px", ""),
            "font_size_pt": font_src.get("font_size_pt", ""),
            "font_color_hex": font_src.get("font_color_hex", ""),
            "sample_text": vtype[0],
            "repeat_key": None,
            "visual_type": vtype[1],
            "item_count": len(members),
            "note": next((m["note"] for m in members if m.get("note")), ""),
        }
        if vtype[0] == "List / menu":
            deferred_lists.append(synth)
        else:
            out.append(synth)

    # A 'List / menu' has no native Power BI visual. When one sits INSIDE a
    # documented container (a dashboard card), that block's own snapshot,
    # position and font already cover it -- emitting it as a separate visual
    # is redundant, so it's dropped entirely. Standalone lists (no enclosing
    # element) are kept, so nothing is silently lost.
    for synth in deferred_lists:
        area = max(1, synth["w"] * synth["h"])
        contained = False
        for o in out:  # all kept rows, incl. members of unclassified runs (cards)
            if o["w"] * o["h"] <= area:
                continue
            ix = max(0, min(synth["x"] + synth["w"], o["x"] + o["w"]) - max(synth["x"], o["x"]))
            iy = max(0, min(synth["y"] + synth["h"], o["y"] + o["h"]) - max(synth["y"], o["y"]))
            if ix * iy >= 0.7 * area:
                contained = True
                break
        if not contained:
            out.append(synth)

    out.sort(key=lambda r: (r["y"], r["x"]))
    for i, r in enumerate(out, 1):
        r["id"] = i
    return out


def _elements_to_rows(elements, canvas_w, canvas_h):
    """Shared post-processing: turn raw JS_EXTRACT elements into the
    deduplicated container list used everywhere else (containers +
    inner-text font matching). Factored out of analyze_html so both the
    single-page and multi-page paths use identical logic."""
    containers = [e for e in elements if (e["hasBg"] or e["hasBorder"] or e.get("svgType")) and e["tag"] != "body"]
    texts = [e for e in elements if e["text"]]

    def is_nested_dup(a, b):
        inside = (b["x"] >= a["x"] - 2 and b["y"] >= a["y"] - 2 and
                  b["x"] + b["w"] <= a["x"] + a["w"] + 2 and
                  b["y"] + b["h"] <= a["y"] + a["h"] + 2)
        area_a = a["w"] * a["h"]
        area_b = b["w"] * b["h"]
        return inside and area_b > 0.85 * area_a

    containers.sort(key=lambda e: e["w"] * e["h"], reverse=True)
    kept = []
    for c in containers:
        # an <svg> filling most of its card is the card's CHART, not a
        # duplicate box -- always keep it
        if not c.get("svgType") and any(is_nested_dup(k, c) for k in kept):
            continue
        kept.append(c)

    kept.sort(key=lambda e: (e["y"], e["x"]))
    return kept, texts


def _build_rows_from_containers(kept, texts, canvas_w, canvas_h, detected_w, detected_h):
    rows = []
    for i, c in enumerate(kept, 1):
        bg_hex = rgb_to_hex(c["bg"]["r"], c["bg"]["g"], c["bg"]["b"]) if c["bg"] else ""
        inner_texts = [t for t in texts if
                        t["x"] >= c["x"] - 2 and t["y"] >= c["y"] - 2 and
                        t["x"] + t["w"] <= c["x"] + c["w"] + 2 and
                        t["y"] + t["h"] <= c["y"] + c["h"] + 2]
        # title = first inner text containing alphanumerics; a leading icon
        # (emoji, arrow, symbol-only text) makes a useless title like '🎯'
        title_src = next((t for t in inner_texts if any(ch.isalnum() for ch in t["text"])),
                          inner_texts[0] if inner_texts else None)
        font_family = title_src["fontFamily"].split(",")[0].strip().strip('"') if title_src else ""
        font_size_px = title_src["fontSize"] if title_src else ""
        font_color_hex = (rgb_to_hex(title_src["color"]["r"], title_src["color"]["g"], title_src["color"]["b"])
                           if title_src and title_src["color"] else "")
        sample_text = title_src["text"] if title_src else ""

        overflow_note = ""
        if c["x"] + c["w"] > canvas_w or c["y"] + c["h"] > canvas_h:
            overflow_note = "EXCEEDS CANVAS — will clip/scroll"

        svg_label, svg_native = _SVG_TYPE_MAP.get(c.get("svgType"), (None, None)) if c.get("svgType") else (None, None)
        rows.append({
            "id": i,
            "tag": c["tag"],
            "label": (svg_label or c["id"] or c["cls"] or sample_text or "")[:30],
            "x": c["x"], "y": c["y"], "w": c["w"], "h": c["h"],
            "x_pt": px_to_pt(c["x"]), "y_pt": px_to_pt(c["y"]),
            "w_pt": px_to_pt(c["w"]), "h_pt": px_to_pt(c["h"]),
            "bg_hex": bg_hex,
            "font_family": font_family,
            "font_size_px": font_size_px,
            "font_size_pt": px_to_pt(font_size_px) if font_size_px else "",
            "font_color_hex": font_color_hex,
            "sample_text": svg_label or sample_text,
            "repeat_key": c.get("repeatKey"),
            "note": overflow_note,
        })
        if svg_label:
            rows[-1]["visual_type"] = svg_native
            if c.get("svgCount"):
                rows[-1]["item_count"] = c["svgCount"]

    kept_ids = {id(c) for c in kept}
    for t in texts:
        if not t.get("repeatKey") or id(t) in kept_ids:
            continue
        rows.append({
            "id": len(rows) + 1, "tag": t["tag"],
            "label": (t.get("id") or t.get("cls") or t["text"] or "")[:30],
            "x": t["x"], "y": t["y"], "w": t["w"], "h": t["h"],
            "x_pt": px_to_pt(t["x"]), "y_pt": px_to_pt(t["y"]),
            "w_pt": px_to_pt(t["w"]), "h_pt": px_to_pt(t["h"]),
            "bg_hex": rgb_to_hex(t["bg"]["r"], t["bg"]["g"], t["bg"]["b"]) if t.get("bg") else "",
            "font_family": (t.get("fontFamily") or "").split(",")[0].strip().strip('"'),
            "font_size_px": t.get("fontSize", ""),
            "font_size_pt": px_to_pt(t["fontSize"]) if t.get("fontSize") else "",
            "font_color_hex": (rgb_to_hex(t["color"]["r"], t["color"]["g"], t["color"]["b"])
                                if t.get("color") else ""),
            "sample_text": t["text"],
            "repeat_key": t["repeatKey"],
            "_from_text": True,
            "note": "",
        })

    rows = collapse_repeated_rows(rows, canvas_w, canvas_h)
    rows = [r for r in rows if not r.get("_from_text")]
    for i, r in enumerate(rows, 1):
        r["id"] = i
    canvas = {"w": canvas_w, "h": canvas_h, "detected_w": detected_w, "detected_h": detected_h}
    return canvas, rows


def detect_html_canvas_size(path, probe_width=1280, probe_height=900):
    """Render once at a generous probe viewport and read back the page's
    real content size (body.offsetWidth/offsetHeight). Works for both
    fixed-size dashboards (canvas == the probe size, unchanged) and tall,
    naturally-scrolling report pages (canvas height grows to fit the real
    content)."""
    from playwright.sync_api import sync_playwright

    abs_path = os.path.abspath(path)
    url = f"file://{abs_path}"
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": probe_width, "height": probe_height})
        page.goto(url)
        page.wait_for_timeout(300)
        dims = page.evaluate("() => ({w: document.body.offsetWidth, h: document.body.offsetHeight})")
        browser.close()
    return dims["w"], dims["h"]


def analyze_html(path, canvas_w, canvas_h, screenshot_path=None):
    from playwright.sync_api import sync_playwright

    abs_path = os.path.abspath(path)
    url = f"file://{abs_path}"

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": canvas_w, "height": canvas_h})
        page.goto(url)
        page.wait_for_timeout(300)  # let fonts/layout settle
        data = page.evaluate(JS_EXTRACT)
        if screenshot_path:
            page.screenshot(path=screenshot_path)
        browser.close()

    kept, texts = _elements_to_rows(data["elements"], canvas_w, canvas_h)
    return _build_rows_from_containers(kept, texts, canvas_w, canvas_h, data["page"]["w"], data["page"]["h"])


# ----------------------------------------------------------------------------
# Image analysis (estimated — pixel/contour/OCR based)
# ----------------------------------------------------------------------------

def dominant_color(img_bgr, x, y, w, h, exclude_dark=True):
    """Return the most common color in a region, ignoring near-black text pixels."""
    region = img_bgr[y:y + h, x:x + w]
    if region.size == 0:
        return (255, 255, 255)
    pixels = region.reshape(-1, 3)
    if exclude_dark:
        brightness = pixels.sum(axis=1)
        mask = brightness > 150  # drop very dark (likely text) pixels
        if mask.sum() > 20:
            pixels = pixels[mask]
    # quantize to reduce noise, then find mode
    quant = (pixels // 8 * 8)
    tuples = [tuple(p) for p in quant[::max(1, len(quant) // 2000)]]  # sample for speed
    common = Counter(tuples).most_common(1)[0][0]
    b, g, r = common
    return (r, g, b)


def cluster_ocr_into_text_blocks(img):
    """Group raw OCR words into readable text blocks (labels/values), since
    tesseract's own line/block grouping falls apart on dashboard-style grids.
    Groups words that sit on the same visual row (similar y-center, similar
    height) and are close together horizontally."""
    import pytesseract
    import numpy as np

    data = pytesseract.image_to_data(img, config="--psm 11", output_type=pytesseract.Output.DICT)
    words = []
    for i in range(len(data["text"])):
        txt = data["text"][i].strip()
        if not txt:
            continue
        words.append({"text": txt, "x": data["left"][i], "y": data["top"][i],
                       "w": data["width"][i], "h": data["height"][i]})
    words.sort(key=lambda w: (w["y"], w["x"]))

    used = [False] * len(words)
    rows = []
    for i, w in enumerate(words):
        if used[i]:
            continue
        row = [w]
        used[i] = True
        anchor_cy = w["y"] + w["h"] / 2
        anchor_h = w["h"]
        for j in range(i + 1, len(words)):
            if used[j]:
                continue
            w2 = words[j]
            cy2 = w2["y"] + w2["h"] / 2
            thresh = max(4, 0.5 * min(anchor_h, w2["h"]))
            if abs(anchor_cy - cy2) < thresh and abs(anchor_h - w2["h"]) < 8:
                row.append(w2)
                used[j] = True
        rows.append(row)

    blocks = []
    gap_limit = 20
    for row in rows:
        row.sort(key=lambda w: w["x"])
        cur = [row[0]]
        for w in row[1:]:
            prev = cur[-1]
            gap = w["x"] - (prev["x"] + prev["w"])
            if gap > gap_limit:
                blocks.append(cur)
                cur = [w]
            else:
                cur.append(w)
        blocks.append(cur)

    text_items = []
    for b in blocks:
        text = " ".join(x["text"] for x in b)
        xs = [x["x"] for x in b]
        ys = [x["y"] for x in b]
        xe = [x["x"] + x["w"] for x in b]
        ye = [x["y"] + x["h"] for x in b]
        x, y = min(xs), min(ys)
        w, h = max(xe) - x, max(ye) - y
        heights = [it["h"] for it in b]
        font_size_px = round(statistics.median(heights))

        region = img[y:y + h, x:x + w].reshape(-1, 3)
        font_color_hex = ""
        if len(region):
            brightness = region.sum(axis=1)
            dark_pixels = region[brightness < 250]
            if len(dark_pixels):
                bm, gm, rm = np.median(dark_pixels, axis=0)
                font_color_hex = rgb_to_hex(rm, gm, bm)

        text_items.append({"text": text, "x": x, "y": y, "w": w, "h": h,
                            "font_size_px": font_size_px, "font_color_hex": font_color_hex})
    return text_items


def _determine_bg_color(quant, w_img, h_img):
    """Background color: prefer the corners IF at least 3 of the 4 agree
    (handles low-contrast "card on slightly different page background"
    designs, where the card's own color covers more total area than the
    true page margin and would otherwise be mistaken for the background).
    Falls back to the globally most-frequent color when corners disagree
    (handles full-bleed headers/sidebars that touch multiple corners)."""
    from collections import Counter
    import numpy as np
    corner_pts = [(0, 0), (w_img - 1, 0), (0, h_img - 1), (w_img - 1, h_img - 1)]
    corner_colors = [tuple(quant[y, x]) for (x, y) in corner_pts]
    corner_counter = Counter(corner_colors)
    most_common_corner, corner_agreement = corner_counter.most_common(1)[0]
    if corner_agreement >= 3:
        return most_common_corner
    flat = quant.reshape(-1, 3)
    colors, counts = np.unique(flat, axis=0, return_counts=True)
    return tuple(colors[np.argmax(counts)])


def _color_region_boxes(img, quant, bg_color, total, min_area_frac, min_fill):
    import cv2
    import numpy as np
    boxes = []
    flat = quant.reshape(-1, 3)
    colors, counts = np.unique(flat, axis=0, return_counts=True)
    for color, count in zip(colors, counts):
        if tuple(color) == bg_color or count < min_area_frac * total:
            continue
        mask = cv2.inRange(quant, color, color)
        num_labels, labels = cv2.connectedComponents(mask)
        for label in range(1, num_labels):
            ys, xs = np.where(labels == label)
            if len(xs) < min_area_frac * total:
                continue
            x, y, w, h = int(xs.min()), int(ys.min()), int(xs.max() - xs.min() + 1), int(ys.max() - ys.min() + 1)
            if len(xs) / (w * h) < min_fill:
                continue
            boxes.append((x, y, w, h))
    return boxes


def detect_panels(img, w_img, h_img):
    """Detect visual containers: solid color-fill blocks (Pass A -- catches
    sidebars, nav-item highlights, header bars, icons, chart segments) plus
    outlined-but-unfilled cards (Pass B -- catches white cards with a thin
    border on a white/light canvas), plus a supplementary low-contrast pass
    for cards that differ from the page background by only a few RGB
    levels (common in minimalist "white card on off-white page" designs)."""
    import cv2
    import numpy as np

    total = w_img * h_img
    boxes = []

    # Pass A (primary): coarse quantization, as tuned against dense/high-
    # contrast dashboards -- keep this exactly as before.
    quant_main = (img // 16 * 16).astype(np.uint8)
    bg_main = _determine_bg_color(quant_main, w_img, h_img)
    boxes += _color_region_boxes(img, quant_main, bg_main, total, min_area_frac=0.0003, min_fill=0.4)

    # Pass A2 (supplementary): finer quantization, but only large regions --
    # catches subtle low-contrast cards without flooding small dashboards
    # with noise (a big region is very unlikely to be an artifact).
    quant_fine = (img // 6 * 6).astype(np.uint8)
    bg_fine = _determine_bg_color(quant_fine, w_img, h_img)
    boxes += _color_region_boxes(img, quant_fine, bg_fine, total, min_area_frac=0.012, min_fill=0.5)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 40, 120)
    kernel = np.ones((3, 3), np.uint8)
    dilated = cv2.dilate(edges, kernel, iterations=1)
    contours, _ = cv2.findContours(dilated, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        area = w * h
        if area < 0.0006 * total or area > 0.95 * total:
            continue
        boxes.append((x, y, w, h))

    # drop hairline dividers / anti-aliasing slivers -- not real "visuals"
    filtered = []
    for (x, y, w, h) in boxes:
        if min(w, h) < 8:
            continue
        if max(w, h) / max(1, min(w, h)) > 20:
            continue
        filtered.append((x, y, w, h))
    boxes = filtered

    # de-duplicate near-identical/heavily-overlapping boxes, keep the larger one
    # (RETR_TREE produces multiple nested contours a couple px apart around the
    # same real edge -- collapse those, but still keep genuinely distinct
    # nested sub-elements like an individual card inside a larger section)
    boxes.sort(key=lambda b: b[2] * b[3], reverse=True)
    kept = []
    for b in boxes:
        x, y, w, h = b
        dup = False
        for kx, ky, kw, kh in kept:
            inter_x = max(0, min(x + w, kx + kw) - max(x, kx))
            inter_y = max(0, min(y + h, ky + kh) - max(y, ky))
            inter = inter_x * inter_y
            union = w * h + kw * kh - inter
            if union and inter / union > 0.6:
                dup = True
                break
        if not dup:
            kept.append(b)
    return kept


def analyze_image(path, canvas_w_hint, canvas_h_hint):
    import cv2
    import numpy as np

    img = cv2.imread(path)
    if img is None:
        raise ValueError(f"Could not read image: {path}")
    h_img, w_img = img.shape[:2]

    panel_boxes = detect_panels(img, w_img, h_img)
    text_items = cluster_ocr_into_text_blocks(img)

    panel_boxes.sort(key=lambda b: (b[1], b[0]))
    panel_rows = []
    for i, (x, y, w, h) in enumerate(panel_boxes, 1):
        r, g, b = dominant_color(img, x, y, w, h)
        bg_hex = rgb_to_hex(r, g, b)

        inner_texts = [t for t in text_items if
                        t["x"] >= x - 3 and t["y"] >= y - 3 and
                        t["x"] + t["w"] <= x + w + 3 and t["y"] + t["h"] <= y + h + 3]

        font_size_px = ""
        font_color_hex = ""
        sample_text = ""
        if inner_texts:
            font_size_px = statistics.median(t["font_size_px"] for t in inner_texts)
            sample_text = inner_texts[0]["text"]
            font_color_hex = inner_texts[0]["font_color_hex"]

        overflow_note = "EXCEEDS CANVAS" if (x + w > w_img or y + h > h_img) else ""

        panel_rows.append({
            "id": i, "tag": "panel", "label": (sample_text or f"panel_{i}")[:30],
            "x": x, "y": y, "w": w, "h": h,
            "x_pt": px_to_pt(x), "y_pt": px_to_pt(y), "w_pt": px_to_pt(w), "h_pt": px_to_pt(h),
            "bg_hex": bg_hex,
            "font_family": "NOT DETECTABLE FROM IMAGE — verify manually",
            "font_size_px": font_size_px, "font_size_pt": px_to_pt(font_size_px) if font_size_px else "",
            "font_color_hex": font_color_hex, "sample_text": sample_text, "note": overflow_note,
        })

    text_rows = []
    for i, t in enumerate(text_items, 1):
        overflow_note = "EXCEEDS CANVAS" if (t["x"] + t["w"] > w_img or t["y"] + t["h"] > h_img) else ""
        text_rows.append({
            "id": i, "tag": "text", "label": t["text"][:30],
            "x": t["x"], "y": t["y"], "w": t["w"], "h": t["h"],
            "x_pt": px_to_pt(t["x"]), "y_pt": px_to_pt(t["y"]), "w_pt": px_to_pt(t["w"]), "h_pt": px_to_pt(t["h"]),
            "bg_hex": "",
            "font_family": "NOT DETECTABLE FROM IMAGE — verify manually",
            "font_size_px": t["font_size_px"], "font_size_pt": px_to_pt(t["font_size_px"]),
            "font_color_hex": t["font_color_hex"], "sample_text": t["text"], "note": overflow_note,
        })

    canvas = {"w": w_img, "h": h_img, "detected_w": w_img, "detected_h": h_img}
    return canvas, panel_rows, text_rows


# ----------------------------------------------------------------------------
# Section / widget detection
# ----------------------------------------------------------------------------
# Groups the flat list of detected elements into a professional document
# structure: Navigation, Header, and N content Sections (left-to-right),
# each containing individual "widgets" (a card/chart/nav-item/etc.) with a
# title, bounding box, and font/color info -- instead of one flat list.

def project_bands(elems, axis, lo, hi, gap_threshold):
    """Project elements onto an axis (x or y) and return contiguous covered
    ranges, merging across gaps smaller than gap_threshold."""
    n = hi - lo
    if n <= 0:
        return []
    coverage = [False] * n
    size_key = "w" if axis == "x" else "h"
    for e in elems:
        a0 = e[axis] - lo
        a1 = a0 + e[size_key]
        a0c, a1c = max(0, a0), min(n, a1)
        for i in range(a0c, a1c):
            coverage[i] = True
    bands = []
    i = 0
    while i < n:
        if not coverage[i]:
            i += 1
            continue
        j = i
        while j < n:
            if coverage[j]:
                j += 1
            else:
                k = j
                while k < n and not coverage[k]:
                    k += 1
                if k - j < gap_threshold and k < n:
                    j = k
                else:
                    break
        bands.append((i + lo, j + lo))
        i = j
    return bands


def union_bbox(elems):
    x0 = min(e["x"] for e in elems)
    y0 = min(e["y"] for e in elems)
    x1 = max(e["x"] + e["w"] for e in elems)
    y1 = max(e["y"] + e["h"] for e in elems)
    return x0, y0, x1 - x0, y1 - y0


def pick_title_style(elems):
    """Choose title text + font/color for a widget: prefer the smallest,
    topmost *legible* text-bearing element (typically a chart/card title
    label, not a 1-2 char OCR noise glyph)."""
    texts = [e for e in elems if e.get("sample_text")]
    if not texts:
        return "", "", "", "", ""
    legible = [t for t in texts if len(t["sample_text"].strip()) >= 3]
    candidates = legible or texts
    candidates.sort(key=lambda e: (e["y"], e["h"]))
    best = candidates[0]
    return (best["sample_text"], best.get("font_family", ""), best.get("font_size_px", ""),
            best.get("font_size_pt", ""), best.get("font_color_hex", ""))


def find_enclosing_panel(all_panels, bbox, margin=3, min_ratio=1.05):
    """Find the tightest (smallest-area) panel that fully contains bbox --
    i.e. the actual background card a piece of content sits on, searched
    across ALL detected panels (not just ones incidentally in the same
    local grouping), since a card is often much larger than the text/chart
    grouped inside it."""
    if not all_panels:
        return None
    x, y, w, h = bbox
    best = None
    for p in all_panels:
        if (p["x"] - margin <= x and p["y"] - margin <= y and
                p["x"] + p["w"] + margin >= x + w and p["y"] + p["h"] + margin >= y + h):
            area = p["w"] * p["h"]
            if area <= min_ratio * (w * h):  # not meaningfully bigger than the content -- not a real "card"
                continue
            if best is None or area < best["w"] * best["h"]:
                best = p
    return best


def _finalize_widget(group_elems, all_panels):
    x, y, w, h = union_bbox(group_elems)
    title, font_family, font_size_px, font_size_pt, font_color_hex = pick_title_style(group_elems)
    enclosing = find_enclosing_panel(all_panels, (x, y, w, h))
    bg_hex = enclosing.get("bg_hex", "") if enclosing else ""
    widget = {
        "title": title or "(untitled visual)",
        "x": x, "y": y, "w": w, "h": h,
        "x_pt": px_to_pt(x), "y_pt": px_to_pt(y), "w_pt": px_to_pt(w), "h_pt": px_to_pt(h),
        "font_family": font_family or "NOT DETECTABLE FROM IMAGE — verify manually",
        "font_size_px": font_size_px, "font_size_pt": font_size_pt,
        "font_color_hex": font_color_hex, "bg_hex": bg_hex,
    }
    if enclosing:
        bx, by, bw, bh = enclosing["x"], enclosing["y"], enclosing["w"], enclosing["h"]
        widget.update({
            "bg_x": bx, "bg_y": by, "bg_w": bw, "bg_h": bh,
            "bg_x_pt": px_to_pt(bx), "bg_y_pt": px_to_pt(by),
            "bg_w_pt": px_to_pt(bw), "bg_h_pt": px_to_pt(bh),
        })
    return widget


def _old_style_row_widgets(elems, lo, hi, gap_threshold, title_max_size, all_panels):
    """The original row-banding heuristic (title-band starts a new widget,
    everything else merges into it) -- used as a fallback for content that
    has no individually-detected background card (so per-element panel
    grouping can't place it), e.g. a chart's title + its plotted body."""
    bands = project_bands(elems, "y", lo, hi, gap_threshold)
    current = []
    widgets_ranges = []
    for b in bands:
        b_elems = [e for e in elems if b[0] <= e["y"] + e["h"] / 2 < b[1]]
        heights = [e["h"] for e in b_elems] or [999]
        is_title_band = min(heights) <= title_max_size and max(heights) <= title_max_size + 4
        if is_title_band and current:
            widgets_ranges.append(current)
            current = [b]
        else:
            current.append(b)
    if current:
        widgets_ranges.append(current)
    widgets_ranges = [(g[0][0], g[-1][1]) for g in widgets_ranges]

    widgets = []
    for (a0, a1) in widgets_ranges:
        in_range = [e for e in elems if e["y"] < a1 and e["y"] + e["h"] > a0]
        if in_range:
            widgets.append(_finalize_widget(in_range, all_panels))
    return widgets


def group_into_widgets(elems, axis, lo, hi, gap_threshold=8, title_max_size=15, merge_titles=True, all_panels=None):
    """Split a region into widgets. For vertical stacks (axis='y') with
    merge_titles=True: a short band (<=title_max_size px tall) starts a new
    widget and everything up to the next short band is merged into it
    (title + its chart/content) -- for uneven chart+title layouts.
    With merge_titles=False (e.g. a simple nav list, or axis='x' header
    strip): every band is its own widget, no merging.

    all_panels, if given, is the FULL global panel list (not just elems in
    this call) -- used to find the actual background card each widget sits
    on, which is frequently larger than the widget's own content.

    Note: grouping purely by "each element's own tightest enclosing panel"
    was tried and reverted -- it correctly separates a grid of same-sized
    cards (e.g. KPI cards) on some dashboards, but badly over-fragments
    others where individual numbers/badges sit in their own small pill-
    shaped background within a card (each pill then wrongly becomes its
    own "widget"). Row-banding by title-height is the more robust default
    across different dashboard styles; a tightly-packed side-by-side grid
    of same-sized cards may end up under-split rather than over-fragmented
    elsewhere -- a deliberate trade-off."""
    bands = project_bands(elems, axis, lo, hi, gap_threshold)
    widgets_ranges = []

    if axis == "x" or not merge_titles:
        widgets_ranges = bands
    else:
        current = []
        for b in bands:
            # use element vertical CENTER (not any overlap) to avoid a
            # neighboring band's element bleeding across the boundary
            b_elems = [e for e in elems if b[0] <= e["y"] + e["h"] / 2 < b[1]]
            heights = [e["h"] for e in b_elems] or [999]
            is_title_band = min(heights) <= title_max_size and max(heights) <= title_max_size + 4
            # a large gap after we're already a couple of bands deep into
            # accumulated content is a hard break -- without this, a non-
            # title band absorbs EVERY later band indefinitely (even an
            # unrelated chart two sections down). The very first band right
            # after a title gets more slack, since a chart body often
            # starts with extra spacing below its own title.
            hard_gap = (current and len(current) >= 2 and
                        (b[0] - current[-1][1]) > max(3 * gap_threshold, 35))
            if (is_title_band or hard_gap) and current:
                widgets_ranges.append(current)
                current = [b]
            else:
                current.append(b)
        if current:
            widgets_ranges.append(current)
        widgets_ranges = [(g[0][0], g[-1][1]) for g in widgets_ranges]

    widgets = []
    for (a0, a1) in widgets_ranges:
        if axis == "y":
            in_range = [e for e in elems if e["y"] < a1 and e["y"] + e["h"] > a0]
        else:
            in_range = [e for e in elems if e["x"] < a1 and e["x"] + e["w"] > a0]
        if in_range:
            widgets.append(_finalize_widget(in_range, all_panels))
    return widgets


def _assign_uids(sections):
    uid = 0
    for sec in sections:
        for w in sec["widgets"]:
            w["uid"] = f"w{uid}"
            uid += 1
    return sections


_FILTER_KEYWORDS = ("year", "region", "city", "category", "filter", "date", "status", "type",
                    "country", "state", "product", "brand", "month", "quarter", "all", "select")


def _classify_sidebar_block(block_texts, last_label):
    """Given the text found in one sidebar sub-block, guess whether it's the
    nav menu or a filter panel. Keyword hits (Year, Region, Category, ...,
    or the word 'All' which is the classic dropdown placeholder) mark a
    Filters block. With no keyword hit (e.g. OCR mangled the word), continue
    whatever section came right before it -- an ambiguous item inside the
    filters area is almost always another filter, not a new nav item."""
    joined = " ".join(block_texts).lower()
    hits = sum(1 for k in _FILTER_KEYWORDS if k in joined)
    if hits >= 1:
        return "Filters"
    return last_label or "Navigation"


def strip_banner_noise(elems, W):
    """Remove long, thin banner/watermark-style text that spans a large
    fraction of the canvas width but isn't real dashboard content (e.g. a
    stock-template promo banner like "Build your next report..."). Left
    in place, a single such element can wrongly bridge the gap between two
    otherwise-separate cards during column/section detection."""
    out = []
    for e in elems:
        aspect = e["w"] / max(1, e["h"])
        if e.get("tag") == "text" and e["w"] > 0.2 * W and aspect > 8:
            continue
        out.append(e)
    return out


def strip_uncontained_margin_text(elems, panel_rows, W, H):
    """Remove stray text sitting entirely above the topmost real card (e.g.
    a stock-template promo banner like "Build your next report..."). OCR
    often fragments such banners into several small pieces at slightly
    different y-positions, so a shape/width heuristic alone can miss some
    fragments -- this instead uses "is it above any real content" which
    catches all of them regardless of fragmentation."""
    big_panels = [p for p in panel_rows if p["w"] * p["h"] > 0.01 * W * H]
    if not big_panels:
        return elems
    top_of_content = min(p["y"] for p in big_panels)
    return [e for e in elems if not (e.get("tag") == "text" and e["y"] + e["h"] <= top_of_content + 2)]


def strip_big_cards(elems, W, H, area_frac=0.03):
    """Remove large card-like panels from a list of elements before using
    them for content banding/title-extraction. A big background card
    shouldn't itself be treated as 'the content' when grouping widgets --
    it's looked up separately via find_enclosing_panel for the Background
    Block info instead. Small colored panels (icons, mini bar-chart
    segments, KPI fills) stay, since those genuinely are visual content."""
    threshold = area_frac * W * H
    return [e for e in elems if not (e.get("tag") == "panel" and e["w"] * e["h"] > threshold)]


def _drop_redundant_untitled(widgets):
    """Drop an untitled widget only if it overlaps an already-titled widget
    significantly (a true duplicate/artifact, e.g. a dropdown's border box
    detected separately from its label) -- keep untitled widgets that don't
    overlap anything else, since those are legitimate content (icons,
    dividers, etc.) that OCR simply couldn't read a label for."""
    titled = [w for w in widgets if w["title"] != "(untitled visual)"]
    kept = list(titled)
    for w in widgets:
        if w["title"] != "(untitled visual)":
            continue
        overlaps = False
        for t in titled:
            ix = max(0, min(w["x"] + w["w"], t["x"] + t["w"]) - max(w["x"], t["x"]))
            iy = max(0, min(w["y"] + w["h"], t["y"] + t["h"]) - max(w["y"], t["y"]))
            inter = ix * iy
            if inter > 0.3 * min(w["w"] * w["h"], t["w"] * t["h"]):
                overlaps = True
                break
        if not overlaps:
            kept.append(w)
    return kept


def find_sidebar_region(all_elems, W, H):
    """Find the sidebar's horizontal extent even when it's split into
    multiple disjoint cards (e.g. a nav-menu card and a separate filters
    card) with a gap between them, and even when the nav items are just
    icons/text with no strong colored container box around them."""
    candidates = [e for e in all_elems if e["x"] < max(20, 0.06 * W)]
    if not candidates:
        return None
    max_right = max(e["x"] + e["w"] for e in candidates)
    if max_right > 0.4 * W:
        return None  # too wide to plausibly be a sidebar
    top = min(e["y"] for e in candidates)
    bottom = max(e["y"] + e["h"] for e in candidates)
    if bottom - top < 0.5 * H:
        return None

    # Grid guard: if the left-edge elements are just the first column of a
    # grid of equal cards (KPI strips, department card grids...), they will
    # have same-row, same-size "twins" further right on the canvas. A real
    # sidebar's nav/filter items don't. Without this check the whole left
    # column of content gets misfiled into a fake "Navigation" section and
    # the rest of the page's cards get scattered into the wrong sections.
    right_elems = [e for e in all_elems if e["x"] >= max_right - 4]
    if right_elems:
        twin_count = 0
        for c in candidates:
            for r in right_elems:
                if (abs(r["y"] - c["y"]) <= 6 and
                        abs(r["w"] - c["w"]) <= 8 and abs(r["h"] - c["h"]) <= 8):
                    twin_count += 1
                    break
        if twin_count >= max(2, 0.4 * len(candidates)):
            return None  # it's a card grid, not a sidebar

    return max_right, top, bottom


def build_sidebar_sections(all_elems, nav_w, top, bottom, W, H, all_panels=None):
    """Split the sidebar region into labeled sections (Navigation / Filters
    / ...) using a wide gap threshold to separate visually distinct cards,
    then extract individual item-widgets within each with a tight gap
    threshold."""
    sidebar_elems = [e for e in all_elems if e["x"] < nav_w and e["y"] < bottom
                      and e["y"] + e["h"] > top and e["h"] < 0.85 * (bottom - top)]
    sidebar_elems = strip_big_cards(sidebar_elems, W, H)
    if not sidebar_elems:
        return []

    coarse_bands = project_bands(sidebar_elems, "y", top, bottom, gap_threshold=35)
    sections = []
    last_label = None
    for (a0, a1) in coarse_bands:
        block_elems = [e for e in sidebar_elems if a0 <= e["y"] + e["h"] / 2 < a1]
        if not block_elems:
            continue
        block_texts = [e["sample_text"] for e in block_elems if e.get("sample_text")]
        label = _classify_sidebar_block(block_texts, last_label)
        last_label = label
        widgets = group_into_widgets(block_elems, axis="y", lo=a0, hi=a1, gap_threshold=6,
                                      merge_titles=False, all_panels=all_panels)
        widgets = _drop_redundant_untitled(widgets)
        if not widgets:
            continue
        # merge into an existing section of the same name if one's already open
        existing = next((s for s in sections if s["name"] == label), None)
        if existing:
            existing["widgets"].extend(widgets)
        else:
            sections.append({"name": label, "widgets": widgets})
    return sections


def build_sidebar_sections_prebuilt(all_elems, nav_w, top, bottom):
    """Same idea as build_sidebar_sections but for already-correct HTML
    widgets (already_widgets=True) -- no fragment merging needed, each
    element is already one complete widget."""
    sidebar_elems = [e for e in all_elems if e["x"] < nav_w and e["y"] < bottom and e["y"] + e["h"] > top]
    if not sidebar_elems:
        return []
    coarse_bands = project_bands(sidebar_elems, "y", top, bottom, gap_threshold=35)
    sections = []
    last_label = None
    for (a0, a1) in coarse_bands:
        block_elems = sorted([e for e in sidebar_elems if a0 <= e["y"] + e["h"] / 2 < a1], key=lambda e: e["y"])
        if not block_elems:
            continue
        block_texts = [e.get("sample_text") or e.get("label") or "" for e in block_elems]
        label = _classify_sidebar_block(block_texts, last_label)
        last_label = label
        existing = next((s for s in sections if s["name"] == label), None)
        if existing:
            existing["widgets"].extend(block_elems)
        else:
            sections.append({"name": label, "widgets": list(block_elems)})
    return sections


def build_sections(canvas, panel_rows, text_rows=None, already_widgets=False):
    """Top-level structure builder. Returns an ordered list of
    {"name": ..., "widgets": [...]} sections: Navigation / Filters (if a
    left-edge sidebar region is found -- even split across multiple cards),
    Header (if a distinct top band is found), then content Sections
    left-to-right.

    If already_widgets=True (HTML mode), each item in panel_rows is treated
    as a complete, already-correct widget (exact computed style/box) and is
    simply bucketed into Navigation/Filters/Header/Section N -- no
    fragment-merging is applied, since there's nothing fragmented to merge."""
    W, H = canvas["w"], canvas["h"]

    def is_decorative(e):
        # a small element with no text of its own (e.g. a progress-bar
        # track/fill div) isn't a meaningful standalone visual -- its title
        # would otherwise just fall back to a raw CSS class name
        return not e.get("sample_text") and e["h"] < 20

    def to_widget(e):
        w = {
            "title": e.get("sample_text") or e.get("label") or "(untitled visual)",
            "x": e["x"], "y": e["y"], "w": e["w"], "h": e["h"],
            "x_pt": e.get("x_pt", px_to_pt(e["x"])), "y_pt": e.get("y_pt", px_to_pt(e["y"])),
            "w_pt": e.get("w_pt", px_to_pt(e["w"])), "h_pt": e.get("h_pt", px_to_pt(e["h"])),
            "font_family": e.get("font_family") or "",
            "font_size_px": e.get("font_size_px", ""), "font_size_pt": e.get("font_size_pt", ""),
            "font_color_hex": e.get("font_color_hex", ""), "bg_hex": e.get("bg_hex", ""),
        }
        if e.get("visual_type"):
            w["visual_type"] = e["visual_type"]
            w["item_count"] = e.get("item_count")
        # in HTML mode the container element IS the background block -- its
        # own box is already exact, so report it as the bg block too
        if e.get("bg_hex"):
            w.update({
                "bg_x": e["x"], "bg_y": e["y"], "bg_w": e["w"], "bg_h": e["h"],
                "bg_x_pt": w["x_pt"], "bg_y_pt": w["y_pt"], "bg_w_pt": w["w_pt"], "bg_h_pt": w["h_pt"],
            })
        return w

    if already_widgets:
        all_elems = strip_uncontained_margin_text(strip_banner_noise(panel_rows, W), panel_rows, W, H)
        sections = []
        sidebar = find_sidebar_region(all_elems, W, H)
        nav_w = 0
        if sidebar:
            nav_w, top, bottom = sidebar
            sections.extend(build_sidebar_sections_prebuilt(all_elems, nav_w, top, bottom))
            sections = [{"name": s["name"], "widgets": [to_widget(e) for e in s["widgets"] if not is_decorative(e)]} for s in sections]

        rest = [e for e in all_elems if e["x"] + e["w"] > nav_w]
        y_bands = project_bands(rest, "y", 0, H, gap_threshold=10)
        header_bottom = y_bands[0][1] if (y_bands and y_bands[0][0] < 25 and len(y_bands) >= 2) else 0

        if header_bottom > 0:
            header_items = sorted([e for e in rest if e["y"] < header_bottom], key=lambda e: e["x"])
            if header_items:
                sections.append({"name": "Header", "widgets": [to_widget(e) for e in header_items if not is_decorative(e)]})

        body = [e for e in rest if e["y"] >= header_bottom]
        x_bands = project_bands(body, "x", nav_w, W, gap_threshold=8)
        for idx, (x0, x1) in enumerate(x_bands, 1):
            col_items = sorted([e for e in body if e["x"] >= x0 - 3 and e["x"] + e["w"] <= x1 + 3], key=lambda e: e["y"])
            if col_items:
                sections.append({"name": f"Section {idx}", "widgets": [to_widget(e) for e in col_items if not is_decorative(e)]})
        return _assign_uids(sections)

    all_elems = strip_uncontained_margin_text(strip_banner_noise(panel_rows + (text_rows or []), W), panel_rows, W, H)
    sections = []

    # --- Sidebar: Navigation / Filters, even if split across multiple cards
    sidebar = find_sidebar_region(all_elems, W, H)
    nav_w = 0
    if sidebar:
        nav_w, top, bottom = sidebar
        sections.extend(build_sidebar_sections(all_elems, nav_w, top, bottom, W, H, all_panels=panel_rows))

    # --- Header vs. Body: find the first big horizontal gap below the top
    rest_elems = [e for e in all_elems if e["x"] + e["w"] > nav_w]
    y_bands = project_bands(rest_elems, "y", 0, H, gap_threshold=10)
    header_bottom = 0
    body_top = 0
    if len(y_bands) >= 2 and y_bands[0][0] < 25:
        header_bottom = y_bands[0][1]
        body_top = y_bands[1][0]
    elif y_bands:
        header_bottom = min(60, y_bands[0][1])
        body_top = header_bottom

    if header_bottom > 0:
        header_elems = strip_big_cards([e for e in rest_elems if e["y"] < header_bottom], W, H)
        header_widgets = group_into_widgets(header_elems, axis="x", lo=nav_w, hi=W, gap_threshold=15, all_panels=panel_rows)
        if header_widgets:
            sections.append({"name": "Header", "widgets": header_widgets})

    # --- Content: column-cluster the remaining body area, then row-band each column
    body_elems = [e for e in rest_elems if e["y"] >= body_top]
    x_bands = project_bands(body_elems, "x", nav_w, W, gap_threshold=8)
    for idx, (x0, x1) in enumerate(x_bands, 1):
        col_elems = strip_big_cards([e for e in body_elems if e["x"] >= x0 - 3 and e["x"] + e["w"] <= x1 + 3], W, H)
        if not col_elems:
            continue
        widgets = group_into_widgets(col_elems, axis="y", lo=body_top, hi=H, gap_threshold=15, all_panels=panel_rows)
        if widgets:
            sections.append({"name": f"Section {idx}", "widgets": widgets})

    return _assign_uids(sections)

# ----------------------------------------------------------------------------
# Screenshot cropping + plain-English descriptions
# ----------------------------------------------------------------------------

def crop_widget_images(screenshot_path, sections, out_dir, canvas, pad=6):
    """Crop each widget's region out of the full wireframe screenshot (using
    ORIGINAL, pre-auto-fit coordinates) and save as a small PNG. Returns a
    dict of {uid: image_path}. Silently returns {} if no screenshot / PIL
    is unavailable / a crop fails -- images are a nice-to-have, not a
    hard requirement."""
    try:
        from PIL import Image
    except Exception:
        return {}
    if not screenshot_path or not os.path.exists(screenshot_path):
        return {}

    try:
        img = Image.open(screenshot_path)
    except Exception:
        return {}

    os.makedirs(out_dir, exist_ok=True)
    W, H = canvas["w"], canvas["h"]
    image_map = {}
    for sec in sections:
        for w in sec["widgets"]:
            uid = w.get("uid")
            if not uid:
                continue
            x0 = max(0, w["x"] - pad)
            y0 = max(0, w["y"] - pad)
            x1 = min(W, w["x"] + w["w"] + pad)
            y1 = min(H, w["y"] + w["h"] + pad)
            if x1 <= x0 or y1 <= y0:
                continue
            try:
                crop = img.crop((x0, y0, x1, y1))
                if crop.width < 4 or crop.height < 4:
                    continue
                crop_path = os.path.join(out_dir, f"{uid}.png")
                crop.save(crop_path)
                image_map[uid] = crop_path
            except Exception:
                continue
    return image_map


_DESCRIPTION_RULES = [
    (("dashboard", "menu", "logo"), "Likely the app logo / brand mark for the dashboard."),
    (("severity", "priority"), "A breakdown widget listing severity/priority levels and their counts, typically shown with small horizontal progress bars."),
    (("satisfaction", "score"), "A breakdown widget showing a satisfaction or scoring distribution across categories."),
    (("category",), "A breakdown widget listing categories and their associated counts or values."),
    (("owner", "group"), "A breakdown widget grouping records by owner or team, with counts per group."),
    (("heatmap",), "A heatmap grid, typically showing intensity/volume across two dimensions (e.g. day vs. week)."),
    (("week", "day of week", "days open"), "A small bar or column chart showing a distribution across days/time buckets."),
    (("month", "monthly"), "A bar chart showing a metric trending across months."),
    (("status", "type"), "Likely a donut/pie chart summarizing a proportion or split between categories (e.g. ticket type or status)."),
    (("vs prev", "trend"), "A trend indicator comparing the current period to a previous one."),
]


def describe_visual(section_name, widget):
    """Best-effort, heuristic plain-English guess at what a visual is, based
    on its title text, its section, and its rough shape. Always hedged
    ('likely'/'appears to be') since this is inferred, not read from the
    source."""
    title = (widget.get("title") or "").strip().lower()

    if widget.get("visual_type"):
        n = widget.get("item_count")
        items = f" with {n} data items detected" if n else ""
        return (f"A {widget['title']}{items} -- build with Power BI's native "
                f"'{widget['visual_type']}' visual. The individual bars/rows/segments are "
                f"data-driven and intentionally not documented one by one.")

    if section_name == "Navigation":
        return f"A navigation menu item — clicking it likely switches the dashboard to the '{widget['title']}' view."
    if section_name == "Filters":
        return f"A filter control — likely a dropdown ('{widget['title']}') that narrows the whole report by that field."
    if section_name == "Header":
        return "A header element — likely the page title, branding, or user/account info shown at the top of the dashboard."

    for keywords, desc in _DESCRIPTION_RULES:
        if any(k in title for k in keywords):
            return desc

    # KPI/metric heuristic: short title, mostly digits/punctuation
    stripped = title.replace(",", "").replace(".", "").replace("%", "").strip()
    if stripped and sum(c.isdigit() for c in stripped) / max(1, len(stripped)) > 0.4 and widget["w"] < 400:
        return "Likely a KPI/metric card — a headline number with a short label, often paired with a period-over-period comparison."

    aspect = widget["w"] / max(1, widget["h"])
    if 0.8 <= aspect <= 1.25 and widget["w"] < 250:
        return "Likely a donut/pie chart or a compact icon-style widget, based on its roughly square shape."
    if aspect > 2.5:
        return "A wide, short element — likely a single-row widget such as a label, filter bar, or divider."

    return f"A dashboard widget titled '{widget['title']}', part of the '{section_name}' section."


def extract_theme(sections):
    """Derive a color theme from the wireframe itself: the accent (brand)
    color -- typically the nav/header's dark color -- and a light tint of it
    for banners/shading. Falls back to a neutral navy if nothing usable is
    found (e.g. an all-white wireframe)."""
    DEFAULT_ACCENT = "1F3864"

    def area(w):
        return w["w"] * w["h"]

    # prefer the Navigation or Header section's dominant background --
    # that's almost always the brand color on a real dashboard
    for pref in ("Navigation", "Header"):
        for sec in sections:
            if sec["name"] == pref:
                candidates = [w for w in sec["widgets"] if w.get("bg_hex")]
                if candidates:
                    biggest = max(candidates, key=area)
                    if brightness_of(biggest["bg_hex"]) < 200:  # skip near-white
                        return biggest["bg_hex"], blend_hex(biggest["bg_hex"], "FFFFFF", 0.85)

    # otherwise, the largest reasonably dark background anywhere
    all_bg = [w for sec in sections for w in sec["widgets"] if w.get("bg_hex")]
    dark_candidates = [w for w in all_bg if brightness_of(w["bg_hex"]) < 180]
    if dark_candidates:
        biggest = max(dark_candidates, key=area)
        return biggest["bg_hex"], blend_hex(biggest["bg_hex"], "FFFFFF", 0.85)

    # light theme with no dark fills at all (accents only on text/borders):
    # fall back to the most-used saturated (non-gray) font color as the brand accent
    font_colors = Counter()
    for sec in sections:
        for w in sec["widgets"]:
            fc = w.get("font_color_hex")
            if fc:
                try:
                    r, g, b = hex_to_rgb(fc)
                except Exception:
                    continue
                if max(r, g, b) - min(r, g, b) > 70 and brightness_of(fc) < 200:  # skip grays/near-white
                    font_colors[fc.upper()] += 1
    if font_colors:
        best = font_colors.most_common(1)[0][0]
        return best, blend_hex(best, "FFFFFF", 0.85)

    return DEFAULT_ACCENT, blend_hex(DEFAULT_ACCENT, "FFFFFF", 0.85)


def extract_palette(sections, max_colors=10):
    """Collect the distinct colors actually used across the wireframe
    (background + font colors), each with a usage count, most-used first --
    a quick reference sheet for building a matching Power BI theme."""
    from collections import Counter
    counts = Counter()
    for sec in sections:
        for w in sec["widgets"]:
            if w.get("bg_hex"):
                counts[w["bg_hex"].upper()] += 1
            if w.get("font_color_hex"):
                counts[w["font_color_hex"].upper()] += 1

    # de-duplicate near-identical colors (quantize to nearest 16 per channel)
    merged = {}
    for hexval, cnt in counts.items():
        try:
            r, g, b = hex_to_rgb(hexval)
        except Exception:
            continue
        key = (r // 20, g // 20, b // 20)
        if key not in merged or cnt > merged[key][1]:
            merged[key] = (hexval, cnt)
        else:
            merged[key] = (merged[key][0], merged[key][1] + cnt)

    ranked = sorted(merged.values(), key=lambda t: -t[1])
    return ranked[:max_colors]


def split_grid_widgets(sections, all_elems, all_panels):
    """If a widget's background block contains several genuinely distinct,
    similarly-sized sub-panels (e.g. 6 individual KPI card boxes inside one
    outer container), split it into that many sub-visuals -- all still
    pointing at the SAME parent block, so they render as N visuals nested
    under one Block heading. Deliberately conservative: requires at least
    2 sub-panels, each meaningfully smaller than the outer block, and all
    within 2.5x of each other in area (a real grid is fairly uniform) --
    a single ambiguous or lopsided case is left untouched."""
    for sec in sections:
        new_widgets = []
        for w in sec["widgets"]:
            if w.get("bg_x") is None:
                new_widgets.append(w)
                continue
            region = [e for e in all_elems if e.get("tag") == "text" and
                      e["x"] >= w["x"] - 2 and e["x"] + e["w"] <= w["x"] + w["w"] + 2 and
                      e["y"] >= w["y"] - 2 and e["y"] + e["h"] <= w["y"] + w["h"] + 2]
            if len(region) < 2:
                new_widgets.append(w)
                continue

            groups = {}
            for e in region:
                enc = find_enclosing_panel(all_panels, (e["x"], e["y"], e["w"], e["h"]), min_ratio=1.3)
                if enc:
                    key = (enc["x"], enc["y"], enc["w"], enc["h"])
                    groups.setdefault(key, []).append(e)

            outer_area = w["bg_w"] * w["bg_h"]
            valid = {k: v for k, v in groups.items() if k[2] * k[3] < 0.5 * outer_area}
            if len(valid) < 2:
                new_widgets.append(w)
                continue
            areas = [k[2] * k[3] for k in valid]
            if max(areas) / min(areas) > 2.5:
                new_widgets.append(w)
                continue

            for key, elems_in in valid.items():
                sub = _finalize_widget(elems_in, all_panels)
                sub["bg_x"], sub["bg_y"], sub["bg_w"], sub["bg_h"] = w["bg_x"], w["bg_y"], w["bg_w"], w["bg_h"]
                sub["bg_x_pt"], sub["bg_y_pt"] = w["bg_x_pt"], w["bg_y_pt"]
                sub["bg_w_pt"], sub["bg_h_pt"] = w["bg_w_pt"], w["bg_h_pt"]
                sub["bg_hex"] = w["bg_hex"]
                new_widgets.append(sub)
        new_widgets.sort(key=lambda w: (w["y"], w["x"]))
        sec["widgets"] = new_widgets
    return sections


def _merge_widgets_into_one(members):
    x0 = min(m["x"] for m in members)
    y0 = min(m["y"] for m in members)
    x1 = max(m["x"] + m["w"] for m in members)
    y1 = max(m["y"] + m["h"] for m in members)
    titled = sorted(members, key=lambda m: -len(m.get("title", "")))
    merged = dict(titled[0])
    merged["x"], merged["y"] = x0, y0
    merged["w"], merged["h"] = x1 - x0, y1 - y0
    merged["x_pt"], merged["y_pt"] = px_to_pt(x0), px_to_pt(y0)
    merged["w_pt"], merged["h_pt"] = px_to_pt(merged["w"]), px_to_pt(merged["h"])
    return merged


def merge_same_block_fragments(sections):
    """Within content sections (Header, Section N -- NOT Navigation/Filters,
    where each item is deliberately kept separate), if 2+ widgets already
    resolve to the exact same background block, they're almost always
    fragments of ONE visual (a chart title + a stray axis-tick label, or a
    table's header + body after merge_overlapping_blocks unified their
    background) rather than meaningfully separate items -- collapse them
    into one."""
    for sec in sections:
        if sec["name"] in ("Navigation", "Filters"):
            continue
        widgets = sec["widgets"]
        groups = {}
        order = []
        for w in widgets:
            key = (w["bg_x"], w["bg_y"], w["bg_w"], w["bg_h"]) if w.get("bg_x") is not None else None
            if key is None:
                order.append((None, w))
                continue
            if key not in groups:
                groups[key] = []
                order.append((key, None))
            groups[key].append(w)

        new_widgets = []
        seen_keys = set()
        for key, standalone in order:
            if key is None:
                new_widgets.append(standalone)
                continue
            if key in seen_keys:
                continue
            seen_keys.add(key)
            members = groups[key]
            new_widgets.append(members[0] if len(members) == 1 else _merge_widgets_into_one(members))
        sec["widgets"] = new_widgets
    return sections


def resolve_html_nesting(sections):
    """If one widget's box fully contains another, meaningfully smaller
    widget (e.g. an outer card containing two distinct sub-cards, each with
    their own background/border), the smaller one becomes a Visual nested
    under the bigger one's Block -- not merged away, not left as an
    unrelated sibling.

    The bigger container's OWN visual entry (e.g. a table with no separate
    box of its own) is then trimmed to just the region NOT covered by its
    children, so it represents only its own remaining content rather than
    overlapping them. If nothing substantial is left outside the children
    (e.g. the "card" is really just a title sitting above 4 metric boxes
    that already cover everything meaningful), the parent's own visual
    entry is dropped entirely -- its box still lives on as the shared
    Block background for its children, just not as a redundant sibling
    Visual."""
    MIN_REMAINING = 50  # px -- below this, the "leftover" is just a title sliver, not real content

    for sec in sections:
        widgets = sec["widgets"]
        parent_children = {}
        for child in widgets:
            child_area = child["w"] * child["h"]
            if child_area <= 0:
                continue
            best_parent = None
            for cand in widgets:
                if cand is child:
                    continue
                cand_area = cand["w"] * cand["h"]
                if cand_area <= child_area * 1.5:  # not meaningfully bigger -- not a container
                    continue
                contains = (cand["x"] - 2 <= child["x"] and cand["y"] - 2 <= child["y"] and
                            cand["x"] + cand["w"] + 2 >= child["x"] + child["w"] and
                            cand["y"] + cand["h"] + 2 >= child["y"] + child["h"])
                if contains and (best_parent is None or cand_area < best_parent["w"] * best_parent["h"]):
                    best_parent = cand
            if best_parent is not None:
                child["bg_x"], child["bg_y"] = best_parent["x"], best_parent["y"]
                child["bg_w"], child["bg_h"] = best_parent["w"], best_parent["h"]
                child["bg_x_pt"], child["bg_y_pt"] = px_to_pt(best_parent["x"]), px_to_pt(best_parent["y"])
                child["bg_w_pt"], child["bg_h_pt"] = px_to_pt(best_parent["w"]), px_to_pt(best_parent["h"])
                child["bg_hex"] = best_parent.get("bg_hex", "")
                parent_children.setdefault(id(best_parent), []).append(child)

        margin = 8
        for parent in widgets:
            children = parent_children.get(id(parent))
            if not children:
                continue
            children_top = min(c["y"] for c in children)
            children_bottom = max(c["y"] + c["h"] for c in children)
            space_above = children_top - parent["y"]

            if space_above >= MIN_REMAINING:
                parent["h"] = space_above - margin
                parent["h_pt"] = px_to_pt(parent["h"])
            else:
                parent["_drop"] = True
                earliest_child = min(children, key=lambda c: (c["y"], c["x"]))
                earliest_child["_block_title_override"] = parent["title"]

        sec["widgets"] = [w for w in widgets if not w.get("_drop")]
    return sections


def merge_overlapping_blocks(sections):
    """Merge two distinct background blocks within the same section into
    one if they overlap heavily in x and are adjacent/overlapping in y --
    the classic case of a table's header row and body rows resolving to
    two slightly different panels instead of being recognized as one
    table card."""
    for sec in sections:
        widgets = sec["widgets"]
        keys = []
        for w in widgets:
            if w.get("bg_x") is not None:
                key = (w["bg_x"], w["bg_y"], w["bg_w"], w["bg_h"])
                if key not in keys:
                    keys.append(key)
        if len(keys) < 2:
            continue

        parent = {k: k for k in keys}

        def find(k):
            while parent[k] != k:
                k = parent[k]
            return k

        for i in range(len(keys)):
            for j in range(i + 1, len(keys)):
                r1, r2 = find(keys[i]), find(keys[j])
                if r1 == r2:
                    continue
                x1, y1, w1, h1 = r1
                x2, y2, w2, h2 = r2
                ix = max(0, min(x1 + w1, x2 + w2) - max(x1, x2))
                x_overlap_frac = ix / min(w1, w2) if min(w1, w2) else 0
                if y1 + h1 <= y2:
                    y_gap = y2 - (y1 + h1)
                elif y2 + h2 <= y1:
                    y_gap = y1 - (y2 + h2)
                else:
                    y_gap = -1  # already overlapping
                if x_overlap_frac > 0.6 and y_gap < 10:
                    parent[r2] = r1

        groups = {}
        for k in keys:
            groups.setdefault(find(k), []).append(k)

        final_bbox = {}
        for rep, members in groups.items():
            if len(members) == 1:
                final_bbox[members[0]] = members[0]
            else:
                xs0 = min(m[0] for m in members)
                ys0 = min(m[1] for m in members)
                xs1 = max(m[0] + m[2] for m in members)
                ys1 = max(m[1] + m[3] for m in members)
                merged_box = (xs0, ys0, xs1 - xs0, ys1 - ys0)
                for m in members:
                    final_bbox[m] = merged_box

        for w in widgets:
            if w.get("bg_x") is not None:
                key = (w["bg_x"], w["bg_y"], w["bg_w"], w["bg_h"])
                nx, ny, nw, nh = final_bbox[key]
                w["bg_x"], w["bg_y"], w["bg_w"], w["bg_h"] = nx, ny, nw, nh
                w["bg_x_pt"], w["bg_y_pt"] = px_to_pt(nx), px_to_pt(ny)
                w["bg_w_pt"], w["bg_h_pt"] = px_to_pt(nw), px_to_pt(nh)
    return sections


def dedupe_cross_section_blocks(sections):
    """If the same physical background card shows up in two different
    sections (a stray OCR fragment that got column-banded into the wrong
    section but still matched a real card via find_enclosing_panel), keep
    it only in the section it first appeared in."""
    box_to_section = {}
    for sec in sections:
        keep = []
        for w in sec["widgets"]:
            if w.get("bg_x") is not None:
                key = (w["bg_x"], w["bg_y"], w["bg_w"], w["bg_h"])
                owner = box_to_section.get(key)
                if owner is None:
                    box_to_section[key] = sec["name"]
                elif owner != sec["name"]:
                    continue  # belongs to an earlier section -- drop the stray duplicate here
            keep.append(w)
        sec["widgets"] = keep
    return [s for s in sections if s["widgets"]]


def assign_block_groups(sections):
    """Give every widget a block_uid: widgets within the same section that
    share the same background block (same bg_x/y/w/h) get the same
    block_uid, so the document can render 'Block N -> its visuals' instead
    of a flat list. Widgets with no detected background block get
    block_uid=None (rendered as standalone, no Block wrapper)."""
    counter = 0
    for sec in sections:
        seen = {}
        for w in sec["widgets"]:
            if w.get("bg_x") is not None:
                key = (w["bg_x"], w["bg_y"], w["bg_w"], w["bg_h"])
                if key not in seen:
                    counter += 1
                    seen[key] = f"blk{counter}"
                w["block_uid"] = seen[key]
            else:
                w["block_uid"] = None
    return sections


def crop_block_images(screenshot_path, sections, out_dir, canvas, pad=4):
    """Crop each distinct background BLOCK (not each individual visual) out
    of the full wireframe screenshot, using ORIGINAL pre-auto-fit
    coordinates. Returns {block_uid: image_path}."""
    try:
        from PIL import Image
    except Exception:
        return {}
    if not screenshot_path or not os.path.exists(screenshot_path):
        return {}
    try:
        img = Image.open(screenshot_path)
    except Exception:
        return {}

    os.makedirs(out_dir, exist_ok=True)
    W, H = canvas["w"], canvas["h"]
    image_map = {}
    seen_uids = set()
    for sec in sections:
        for w in sec["widgets"]:
            uid = w.get("block_uid")
            if not uid or uid in seen_uids:
                continue
            seen_uids.add(uid)
            x0 = max(0, w["bg_x"] - pad)
            y0 = max(0, w["bg_y"] - pad)
            x1 = min(W, w["bg_x"] + w["bg_w"] + pad)
            y1 = min(H, w["bg_y"] + w["bg_h"] + pad)
            if x1 <= x0 or y1 <= y0:
                continue
            try:
                crop = img.crop((x0, y0, x1, y1))
                if crop.width < 4 or crop.height < 4:
                    continue
                crop_path = os.path.join(out_dir, f"{uid}.png")
                crop.save(crop_path)
                image_map[uid] = crop_path
            except Exception:
                continue
    return image_map


def _shade_cell(cell, hex_color):
    """Set a table cell's background fill color (python-docx has no public
    API for this, so we drop into the underlying XML)."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    if not hex_color:
        return
    hex_color = hex_color.lstrip("#")
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _set_cell_valign_center(cell):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    tcPr = cell._tc.get_or_add_tcPr()
    valign = OxmlElement("w:vAlign")
    valign.set(qn("w:val"), "center")
    tcPr.append(valign)


def _add_fitted_picture(run, image_path, max_width_in, max_height_in):
    """Add a picture sized to fit within BOTH max_width_in and max_height_in
    (whichever constraint is tighter for this image's actual aspect ratio)
    -- prevents a very tall/narrow crop (e.g. a full sidebar) from being
    forced to a fixed width and ballooning to many inches tall, which would
    spill across several mostly-blank pages."""
    from docx.shared import Inches
    try:
        from PIL import Image
        with Image.open(image_path) as im:
            px_w, px_h = im.size
    except Exception:
        run.add_picture(image_path, width=Inches(max_width_in))
        return
    if px_w <= 0 or px_h <= 0:
        run.add_picture(image_path, width=Inches(max_width_in))
        return
    aspect = px_h / px_w
    width_in = max_width_in
    height_in = width_in * aspect
    if height_in > max_height_in:
        height_in = max_height_in
        width_in = height_in / aspect
    run.add_picture(image_path, width=Inches(width_in), height=Inches(height_in))


def _rgb_color(hexval):
    from docx.shared import RGBColor
    return RGBColor.from_string(hexval.lstrip("#"))


def _readable_text_color(hex_bg):
    """Pick black or white text so it stays readable against a swatch fill."""
    try:
        return "000000" if brightness_of(hex_bg) > 140 else "FFFFFF"
    except Exception:
        return "000000"


def _add_bottom_border(paragraph, color="1F3864", size=12):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def _add_page_number_footer(doc):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = 1  # center
    run = p.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


def _add_bookmark(paragraph, name):
    """Mark this paragraph so a PAGEREF field elsewhere can resolve to
    whatever physical page it lands on once Word/LibreOffice lays out the
    document."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(abs(hash(name)) % 100000))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(abs(hash(name)) % 100000))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def _add_pageref_field(paragraph, bookmark_name):
    """Insert a PAGEREF field run -- resolves to the real page number of
    the matching bookmark when the document is opened/converted (both Word
    and LibreOffice update these fields automatically on open)."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f'PAGEREF {bookmark_name} \\h'
    fld_separate = OxmlElement("w:fldChar")
    fld_separate.set(qn("w:fldCharType"), "separate")
    placeholder = OxmlElement("w:t")
    placeholder.text = "?"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_separate)
    run._r.append(placeholder)
    run._r.append(fld_end)
    return run


def generate_design_doc(pages, out_path, source_path, mode):
    """pages: list of dicts, each with:
         label (str or None -- None means "single page, no Page heading"),
         canvas, sections, overflow_count, scale, hero_image_path,
         image_map, block_image_map
    A single-entry list with label=None renders identically to the
    original single-page-only version of this function (no behavior
    change for existing callers). 2+ entries (or a single entry with a
    real label) render each under its own "N. Page Label" heading, with
    every deeper heading level and numbering shifted to make room."""
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE

    multi_page = len(pages) > 1
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)
    _add_page_number_footer(doc)

    all_sections = [s for pg in pages for s in pg["sections"]]
    content_pages = [pg for pg in pages if not pg.get("is_shell")] or pages
    first_canvas = content_pages[0]["canvas"]
    first_hero = next((pg["hero_image_path"] for pg in content_pages if pg.get("hero_image_path")), None)
    first_hero_label = content_pages[0].get("label")
    total_overflow = sum(pg.get("overflow_count", 0) for pg in pages)
    pages_with_overflow = [pg for pg in pages if pg.get("overflow_count", 0)]

    ACCENT, ACCENT_LIGHT = extract_theme(all_sections)
    accent_text_color = _readable_text_color(ACCENT)
    palette = extract_palette(all_sections)

    # tighten margins so the cover block reads as closer to full-bleed
    sec0 = doc.sections[0]
    sec0.top_margin = Inches(0.4)
    sec0.bottom_margin = Inches(0.4)
    sec0.left_margin = Inches(0.6)
    sec0.right_margin = Inches(0.6)
    usable_width = sec0.page_width - sec0.left_margin - sec0.right_margin
    usable_height = sec0.page_height - sec0.top_margin - sec0.bottom_margin - Inches(0.6)

    doc.core_properties.author = "Ajith Khan (IN1661), Senior Consultant"
    doc.core_properties.title = "Wireframe Design Specification"

    # ---------------- Cover page: full-bleed colored block, sized to fill the page ----------------
    cover_table = doc.add_table(rows=1, cols=1)
    cover_table.autofit = False
    cover_cell = cover_table.rows[0].cells[0]
    cover_cell.width = usable_width
    cover_table.rows[0].height = usable_height
    cover_table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    _shade_cell(cover_cell, ACCENT)

    # clear the default empty paragraph's spacing so content starts near the top
    cover_cell.paragraphs[0].paragraph_format.space_after = Pt(0)

    _cover_state = {"first_used": False}

    def cover_para(text=None, size=11, bold=False, italic=False, color=None, align=WD_ALIGN_PARAGRAPH.CENTER,
                   space_before=0, space_after=6):
        if not _cover_state["first_used"]:
            p = cover_cell.paragraphs[0]
            _cover_state["first_used"] = True
        else:
            p = cover_cell.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        if text:
            run = p.add_run(text)
            run.font.size = Pt(size)
            run.bold = bold
            run.italic = italic
            run.font.color.rgb = _rgb_color(color or accent_text_color)
        return p

    cover_para(space_before=14, space_after=2)
    cover_para("WIREFRAME DESIGN SPECIFICATION", size=30, bold=True, space_after=6)
    subtitle = f"{len(pages)} pages - " if multi_page else ""
    cover_para(f"{subtitle}Layout, color & typography reference for Power BI rebuild", size=13, italic=True, space_after=10)
    # author credit at the TOP of the cover -- guaranteed to be on page 1
    # even if the lower part of the cover block ever reflows
    cover_para("Prepared by  Ajith Khan  \u00b7  IN1661  \u00b7  Senior Consultant", size=13, bold=True, space_after=20)

    if first_hero and os.path.exists(first_hero):
        hero_p = cover_para(space_after=4)
        hero_run = hero_p.add_run()
        try:
            _add_fitted_picture(hero_run, first_hero, 5.6, 3.5)
        except Exception:
            pass
        cap = f"'{first_hero_label}' page as analyzed" if first_hero_label else "Full wireframe as analyzed"
        cover_para(cap, size=8.5, italic=True, space_after=20)
    else:
        cover_para(space_after=20)

    # quick-stats strip
    total_sections = len(all_sections)
    total_widgets = sum(len(s["widgets"]) for s in all_sections)
    ratio = first_canvas["w"] / first_canvas["h"] if first_canvas["h"] else 0
    is_16_9 = abs(ratio - (16 / 9)) < 0.03
    if multi_page:
        stats_labels = ["Pages", "Sections", "Visuals Documented", "Canvas Size"]
        stats_values = [str(len(pages)), str(total_sections), str(total_widgets), f"{first_canvas['w']} x {first_canvas['h']} px"]
    else:
        stats_labels = ["Sections", "Visuals Documented", "Canvas Size", "Aspect Ratio"]
        stats_values = [str(total_sections), str(total_widgets), f"{first_canvas['w']} x {first_canvas['h']} px",
                         f"{ratio:.3f}  ({'16:9' if is_16_9 else 'custom'})"]
    stats_table = cover_cell.add_table(rows=2, cols=4)
    stats_table.autofit = True
    for i in range(4):
        c1 = stats_table.rows[0].cells[i]
        c1.text = stats_labels[i]
        c1.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c1.paragraphs[0].runs[0].font.size = Pt(9)
        c1.paragraphs[0].runs[0].font.color.rgb = _rgb_color(accent_text_color)
        c1.paragraphs[0].runs[0].italic = True
        c2 = stats_table.rows[1].cells[i]
        c2.text = stats_values[i]
        c2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c2.paragraphs[0].runs[0].font.size = Pt(13)
        c2.paragraphs[0].runs[0].bold = True
        c2.paragraphs[0].runs[0].font.color.rgb = _rgb_color(accent_text_color)

    cover_para(space_before=20, space_after=2)
    cover_para(f"Source file: {os.path.basename(source_path)}   |   Mode: {mode}   |   "
               f"Generated: {__import__('datetime').date.today().isoformat()}", size=9, space_after=2)
    cover_para("Generated by  Ajith Khan - IN1661,  Senior Consultant", size=9, bold=True, space_after=10)

    if total_overflow:
        where = f" (on {len(pages_with_overflow)} of {len(pages)} pages)" if multi_page else ""
        cover_para(f"[!] {total_overflow} visual(s) exceeded the canvas in raw detection{where}. "
                   f"An auto-fit scale has been applied per-page so everything fits with no scrolling/clipping.",
                   size=9, italic=True, space_after=10)

    if mode.startswith("Image"):
        cover_para("Generated from an image: positions/colors are estimated via computer vision and OCR, "
                   "font family can't be detected from a flat image, visual descriptions below are best-effort "
                   "guesses, and a few tightly packed rows may be grouped together. Cross-check against "
                   "the source before finalizing.", size=8.5, italic=True, space_after=4)

    # ---------------- Contents + Color Palette (fills the second page properly) ----------------
    contents_heading = doc.add_heading("Contents", level=1)
    contents_heading.paragraph_format.page_break_before = True
    for run in contents_heading.runs:
        run.font.color.rgb = _rgb_color(ACCENT)
    _add_bottom_border(contents_heading, color=ACCENT, size=16)

    if multi_page:
        index_note = doc.add_paragraph()
        index_note_run = index_note.add_run("Page Index")
        index_note_run.bold = True
        index_note_run.font.size = Pt(13)
        index_note_run.font.color.rgb = _rgb_color(ACCENT)

        index_table = doc.add_table(rows=0, cols=2)
        index_table.style = "Light Grid Accent 1"
        index_table.columns[0].width = Inches(5.0)
        index_table.columns[1].width = Inches(1.2)
        for p_idx, pg in enumerate(pages, 1):
            row = index_table.add_row()
            row.cells[0].text = f"{p_idx}. {pg['label'] or f'Page {p_idx}'}"
            row.cells[0].paragraphs[0].runs[0].bold = True
            row.cells[0].paragraphs[0].runs[0].font.size = Pt(11)
            row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            pageref_run = _add_pageref_field(row.cells[1].paragraphs[0], f"page_{p_idx}")
            pageref_run.font.bold = True
            pageref_run.font.size = Pt(11)
            _set_cell_valign_center(row.cells[0])
            _set_cell_valign_center(row.cells[1])
        doc.add_paragraph()

    toc_table = doc.add_table(rows=1, cols=3)
    toc_table.style = "Light Grid Accent 1"
    hdr = toc_table.rows[0].cells
    toc_headers = ["Page / Section", "Visuals", "Dominant Color"] if multi_page else ["Section", "Visuals", "Dominant Color"]
    for i, label in enumerate(toc_headers):
        hdr[i].text = label
        hdr[i].paragraphs[0].runs[0].bold = True
        hdr[i].paragraphs[0].runs[0].font.size = Pt(11)
        _shade_cell(hdr[i], ACCENT_LIGHT)
    toc_table.columns[0].width = Inches(3.4)
    toc_table.columns[1].width = Inches(1.4)
    toc_table.columns[2].width = Inches(2.0)
    toc_table.rows[0].height = Inches(0.4)
    toc_table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    def add_toc_row(name, widget_count, dom_color, bold=True, indent=False):
        row = toc_table.add_row()
        row.height = Inches(0.4 if indent else 0.45)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row.cells[0].text = ("      " if indent else "") + name
        row.cells[0].paragraphs[0].runs[0].bold = bold
        row.cells[0].paragraphs[0].runs[0].font.size = Pt(10.5 if indent else 12)
        if not indent:
            _shade_cell(row.cells[0], ACCENT_LIGHT)
        row.cells[1].text = str(widget_count)
        row.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row.cells[1].paragraphs[0].runs[0].font.size = Pt(10.5 if indent else 12)
        if dom_color:
            row.cells[2].text = dom_color
            _shade_cell(row.cells[2], dom_color)
            row.cells[2].paragraphs[0].runs[0].font.color.rgb = _rgb_color(_readable_text_color(dom_color))
            row.cells[2].paragraphs[0].runs[0].bold = True
            row.cells[2].paragraphs[0].runs[0].font.size = Pt(10.5 if indent else 11)
        _set_cell_valign_center(row.cells[0])
        _set_cell_valign_center(row.cells[1])
        _set_cell_valign_center(row.cells[2])

    def dominant_color_of(sec):
        sec_colors = [w["bg_hex"] for w in sec["widgets"] if w.get("bg_hex")]
        return Counter(sec_colors).most_common(1)[0][0] if sec_colors else ""

    if multi_page:
        for pg in pages:
            pg_widget_total = sum(len(s["widgets"]) for s in pg["sections"])
            add_toc_row(pg["label"] or "(untitled page)", pg_widget_total, "", bold=True, indent=False)
            for sec in pg["sections"]:
                add_toc_row(sec["name"], len(sec["widgets"]), dominant_color_of(sec), bold=False, indent=True)
    else:
        for sec in all_sections:
            add_toc_row(sec["name"], len(sec["widgets"]), dominant_color_of(sec), bold=True, indent=False)

    doc.add_paragraph()
    palette_heading = doc.add_heading("Detected Color Palette", level=2)
    for run in palette_heading.runs:
        run.font.color.rgb = _rgb_color(ACCENT)
    palette_note = doc.add_paragraph()
    palette_note_run = palette_note.add_run(
        "The most-used colors found across the wireframe - use these directly when building a matching "
        "custom theme in Power BI (View -> Themes -> Customize current theme)."
    )
    palette_note_run.italic = True
    palette_note_run.font.size = Pt(10)
    palette_note_run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

    if palette:
        n = len(palette)
        pal_table = doc.add_table(rows=2, cols=n)
        for i, (hexval, cnt) in enumerate(palette):
            swatch_cell = pal_table.rows[0].cells[i]
            swatch_cell.text = " "
            _shade_cell(swatch_cell, hexval)
            label_cell = pal_table.rows[1].cells[i]
            label_cell.text = hexval
            label_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            label_cell.paragraphs[0].runs[0].font.size = Pt(9)
            label_cell.paragraphs[0].runs[0].bold = True
        pal_table.rows[0].height = Inches(0.9)
        pal_table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        pal_table.rows[1].height = Inches(0.3)
        pal_table.rows[1].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    # --- Power BI setup guide: practical, uses the extracted theme, fills remaining space ---
    doc.add_paragraph()
    setup_heading = doc.add_heading("Power BI Setup Guide", level=2)
    for run in setup_heading.runs:
        run.font.color.rgb = _rgb_color(ACCENT)
    setup_note = doc.add_paragraph()
    setup_note_run = setup_note.add_run(
        "Recommended starting settings so the report canvas matches this wireframe before you place visuals."
        + (" Sizes shown are for the first page -- check each page's own Canvas Size if they differ." if multi_page else "")
    )
    setup_note_run.italic = True
    setup_note_run.font.size = Pt(10)
    setup_note_run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

    setup_table = doc.add_table(rows=0, cols=2)
    setup_table.style = "Light Grid Accent 1"
    setup_table.columns[0].width = Inches(2.4)
    setup_table.columns[1].width = Inches(4.4)

    def setup_row(prop, value, swatch=None):
        row = setup_table.add_row()
        row.height = Inches(0.4)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row.cells[0].text = prop
        row.cells[0].paragraphs[0].runs[0].bold = True
        row.cells[0].paragraphs[0].runs[0].font.size = Pt(10.5)
        _shade_cell(row.cells[0], ACCENT_LIGHT)
        row.cells[1].text = value
        row.cells[1].paragraphs[0].runs[0].font.size = Pt(10.5)
        if swatch:
            _shade_cell(row.cells[1], swatch)
            row.cells[1].paragraphs[0].runs[0].font.color.rgb = _rgb_color(_readable_text_color(swatch))
            row.cells[1].paragraphs[0].runs[0].bold = True
        _set_cell_valign_center(row.cells[0])
        _set_cell_valign_center(row.cells[1])

    setup_row("Page Size (Custom)", f"{first_canvas['w']/96:.2f}\" x {first_canvas['h']/96:.2f}\"   ({first_canvas['w']} x {first_canvas['h']} px)")
    setup_row("Page Background", ACCENT_LIGHT if brightness_of(ACCENT_LIGHT) > 200 else "#FFFFFF")
    setup_row("Primary / Accent Color", f"#{ACCENT.lstrip('#')}", swatch=f"#{ACCENT.lstrip('#')}")
    setup_row("Suggested Font", "Segoe UI (Power BI default) - verify against source" if mode.startswith("Image") else "See each visual's Font Family below")
    setup_row("Grid / Snap", "Enable snap-to-grid at a fine interval (e.g. 4-8 px) for pixel-accurate placement")

    doc.add_page_break()

    # ---------------- Section-by-section detail, per page ----------------
    def render_sections(sections, level_offset, numbering_prefix):
        for s_idx, sec in enumerate(sections, 1):
            sec_number = f"{numbering_prefix}{s_idx}"
            sec_heading = doc.add_heading(f"{sec_number}. {sec['name']}", level=1 + level_offset)
            for run in sec_heading.runs:
                run.font.color.rgb = _rgb_color(ACCENT)
            _add_bottom_border(sec_heading, color=ACCENT, size=16)

            count_p = doc.add_paragraph()
            count_run = count_p.add_run(f"{len(sec['widgets'])} visual(s) detected in this section.")
            count_run.italic = True
            count_run.font.size = Pt(9.5)
            count_run.font.color.rgb = RGBColor(0x70, 0x70, 0x70)

            def make_table():
                t = doc.add_table(rows=0, cols=2)
                t.style = "Light Grid Accent 1"
                t.alignment = WD_TABLE_ALIGNMENT.LEFT
                t.columns[0].width = Inches(1.9)
                t.columns[1].width = Inches(4.3)
                return t

            def add_row(table, prop, value, color_swatch=None):
                row = table.add_row()
                row.cells[0].text = prop
                row.cells[0].paragraphs[0].runs[0].bold = True
                _shade_cell(row.cells[0], ACCENT_LIGHT)
                if color_swatch:
                    row.cells[1].text = value
                    _shade_cell(row.cells[1], color_swatch)
                    run = row.cells[1].paragraphs[0].runs[0]
                    run.font.color.rgb = _rgb_color(_readable_text_color(color_swatch))
                    run.bold = True
                else:
                    row.cells[1].text = value

            def render_visual(numbering, w, indent=False):
                heading_level = (3 if indent else 2) + level_offset
                prefix = "->  " if indent else ""
                visual_heading = doc.add_heading(f"{numbering}  {prefix}{w['title']}", level=heading_level)
                for run in visual_heading.runs:
                    run.font.color.rgb = RGBColor(0x2E, 0x50, 0x90)

                img_path = image_map.get(w.get("uid"))
                if img_path and os.path.exists(img_path):
                    img_p = doc.add_paragraph()
                    img_run = img_p.add_run()
                    try:
                        _add_fitted_picture(img_run, img_path, 2.3 if indent else 2.6, 3.0)
                    except Exception:
                        pass

                desc_p = doc.add_paragraph()
                desc_run = desc_p.add_run(describe_visual(sec["name"], w))
                desc_run.italic = True
                desc_run.font.size = Pt(9.5)
                desc_run.font.color.rgb = RGBColor(0x45, 0x45, 0x45)

                table = make_table()
                if w.get("visual_type"):
                    add_row(table, "Power BI Visual", w["visual_type"])
                add_row(table, "Position (X, Y)", f"{w['x']} px, {w['y']} px   ({w['x_pt']} pt, {w['y_pt']} pt)")
                add_row(table, "Dimensions (W x H)", f"{w['w']} x {w['h']} px   ({w['w_pt']} x {w['h_pt']} pt)")
                if not w.get("block_uid") and w.get("bg_hex"):
                    add_row(table, "Background Color", w["bg_hex"], color_swatch=w["bg_hex"])
                if w.get("font_family"):
                    add_row(table, "Font Family", w["font_family"])
                if w.get("font_size_px") not in ("", None):
                    add_row(table, "Font Size", f"{w['font_size_px']} px  ({w.get('font_size_pt', '')} pt)")
                if w.get("font_color_hex"):
                    add_row(table, "Font Color", w["font_color_hex"], color_swatch=w["font_color_hex"])
                doc.add_paragraph()

            groups = []  # list of (block_uid_or_None, [widgets])
            group_index_by_buid = {}
            for w in sec["widgets"]:
                buid = w.get("block_uid")
                if buid:
                    if buid in group_index_by_buid:
                        groups[group_index_by_buid[buid]][1].append(w)
                    else:
                        group_index_by_buid[buid] = len(groups)
                        groups.append((buid, [w]))
                else:
                    groups.append((None, [w]))

            block_num = 0
            for buid, group_widgets in groups:
                if buid:
                    block_num += 1
                    rep = group_widgets[0]
                    override = next((w.get("_block_title_override") for w in group_widgets if w.get("_block_title_override")), None)
                    block_title = override or rep["title"]
                    block_heading = doc.add_heading(f"{sec_number}.{block_num}  BLOCK: {block_title}", level=2 + level_offset)
                    for run in block_heading.runs:
                        run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
                        run.bold = True

                    block_img = block_image_map.get(buid)
                    if block_img and os.path.exists(block_img):
                        bimg_p = doc.add_paragraph()
                        bimg_run = bimg_p.add_run()
                        try:
                            _add_fitted_picture(bimg_run, block_img, 3.4, 4.5)
                        except Exception:
                            pass

                    block_table = make_table()
                    add_row(block_table, "Position (X, Y)", f"{rep['bg_x']} px, {rep['bg_y']} px   ({rep['bg_x_pt']} pt, {rep['bg_y_pt']} pt)")
                    add_row(block_table, "Dimensions (W x H)", f"{rep['bg_w']} x {rep['bg_h']} px   ({rep['bg_w_pt']} x {rep['bg_h_pt']} pt)")
                    if rep.get("bg_hex"):
                        add_row(block_table, "Background Color", rep["bg_hex"], color_swatch=rep["bg_hex"])
                    doc.add_paragraph()

                    for v_idx, w in enumerate(group_widgets, 1):
                        render_visual(f"{sec_number}.{block_num}.{v_idx}", w, indent=True)
                else:
                    for w in group_widgets:
                        block_num += 1
                        render_visual(f"{sec_number}.{block_num}", w, indent=False)

    for p_idx, pg in enumerate(pages, 1):
        image_map = pg.get("image_map") or {}
        block_image_map = pg.get("block_image_map") or {}

        if multi_page:
            page_heading = doc.add_heading(f"{p_idx}. {pg['label'] or f'Page {p_idx}'}", level=1)
            page_heading.paragraph_format.page_break_before = (p_idx > 1)
            page_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in page_heading.runs:
                run.font.color.rgb = _rgb_color(ACCENT)
                run.bold = True
                run.font.size = Pt(26)
            _add_bottom_border(page_heading, color=ACCENT, size=24)
            _add_bookmark(page_heading, f"page_{p_idx}")

            page_widget_total = sum(len(s["widgets"]) for s in pg["sections"])
            pg_count_p = doc.add_paragraph()
            pg_count_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pg_count_run = pg_count_p.add_run(
                f"Canvas: {pg['canvas']['w']} x {pg['canvas']['h']} px   |   "
                f"{len(pg['sections'])} section(s), {page_widget_total} visual(s) total."
            )
            pg_count_run.italic = True
            pg_count_run.font.size = Pt(10)
            pg_count_run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

            # full-page snapshot right up front, so anyone scrolling through
            # the document can immediately see which page they're looking at
            pg_hero = pg.get("hero_image_path")
            if pg_hero and os.path.exists(pg_hero):
                snap_p = doc.add_paragraph()
                snap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                snap_run = snap_p.add_run()
                try:
                    _add_fitted_picture(snap_run, pg_hero, 6.2, 4.3)
                except Exception:
                    pass
                snap_cap = doc.add_paragraph()
                snap_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                snap_cap_run = snap_cap.add_run(f"'{pg['label'] or f'Page {p_idx}'}' -- full page snapshot")
                snap_cap_run.italic = True
                snap_cap_run.font.size = Pt(8.5)
                snap_cap_run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
                doc.add_paragraph()

            for note in pg.get("dedup_notes") or []:
                note_p = doc.add_paragraph()
                note_run = note_p.add_run("[shared] " + note)
                note_run.italic = True
                note_run.font.size = Pt(9)
                note_run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

            render_sections(pg["sections"], level_offset=1, numbering_prefix=f"{p_idx}.")
        else:
            render_sections(pg["sections"], level_offset=0, numbering_prefix="")

    doc.save(out_path)
    return out_path


# ----------------------------------------------------------------------------
# Auto-fit: rescale all visuals proportionally so nothing exceeds the canvas
# (canvas size and background stay exactly as detected — only the visuals'
# x/y/w/h shrink, keeping their relative layout, so nothing scrolls/clips)
# ----------------------------------------------------------------------------

def fit_to_canvas(canvas, rows):
    if not rows:
        return 1.0, rows

    max_x = max(r["x"] + r["w"] for r in rows)
    max_y = max(r["y"] + r["h"] for r in rows)

    # small margin so nothing sits flush against the edge after scaling
    margin = 0.98
    scale = min(
        (canvas["w"] * margin) / max_x if max_x > 0 else 1.0,
        (canvas["h"] * margin) / max_y if max_y > 0 else 1.0,
        1.0,  # never upscale — only shrink to fit
    )

    fitted = []
    for r in rows:
        fr = dict(r)
        fr["x"] = round(r["x"] * scale)
        fr["y"] = round(r["y"] * scale)
        fr["w"] = round(r["w"] * scale)
        fr["h"] = round(r["h"] * scale)
        fr["x_pt"] = px_to_pt(fr["x"])
        fr["y_pt"] = px_to_pt(fr["y"])
        fr["w_pt"] = px_to_pt(fr["w"])
        fr["h_pt"] = px_to_pt(fr["h"])
        if r.get("font_size_px"):
            fr["font_size_px"] = round(r["font_size_px"] * scale, 1)
            fr["font_size_pt"] = px_to_pt(fr["font_size_px"])
        fr["note"] = ""  # no longer overflowing after scaling
        fitted.append(fr)

    return scale, fitted


# ----------------------------------------------------------------------------
# Report writer
# ----------------------------------------------------------------------------

def write_report(canvas, rows, out_path, source_path, mode, text_rows=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()

    scale, fitted_rows = fit_to_canvas(canvas, rows)

    # --- Canvas Info sheet ---
    ws1 = wb.active
    ws1.title = "Canvas Info"
    ws1.append(["Source file", os.path.basename(source_path)])
    ws1.append(["Analysis mode", mode])
    ws1.append(["Canvas width (px)", canvas["w"]])
    ws1.append(["Canvas height (px)", canvas["h"]])
    ratio = canvas["w"] / canvas["h"] if canvas["h"] else 0
    ws1.append(["Aspect ratio", round(ratio, 3)])
    is_16_9 = abs(ratio - (16 / 9)) < 0.03
    ws1.append(["Matches 16:9?", "Yes" if is_16_9 else f"No (16:9 = {16/9:.3f})"])
    ws1.append(["Power BI custom page size (in)",
                f"{canvas['w']/96:.2f} x {canvas['h']/96:.2f} (at 96 DPI)"])
    ws1.append(["Background color", "See Visuals sheet / source — canvas bg is unchanged by auto-fit"])
    ws1.append(["Panels/Cards detected", len(rows)])
    if text_rows is not None:
        ws1.append(["Text labels/values detected", len(text_rows)])
    overflow_rows = [r for r in rows if r["note"]]
    ws1.append(["Panels exceeding canvas bounds", len(overflow_rows)])
    ws1.append(["Auto-fit scale factor applied", f"{scale:.3f}" + ("  (no scaling needed)" if scale >= 0.999 else "  (see 'Auto-Fit Layout' tab)")])
    for col in "AB":
        ws1.column_dimensions[col].width = 40
    for r in ws1["A1:A10"]:
        for c in r:
            c.font = Font(bold=True)

    def build_visuals_sheet(title, data_rows, note_text=None):
        ws = wb.create_sheet(title)
        if note_text:
            ws.append([note_text])
            ws.cell(row=1, column=1).font = Font(italic=True, color="806000")
            ws.append([])
        headers = ["ID", "Tag/Type", "Label", "X (px)", "Y (px)", "W (px)", "H (px)",
                   "X (pt)", "Y (pt)", "W (pt)", "H (pt)", "Background Hex",
                   "Font Family", "Font Size (px)", "Font Size (pt)", "Font Color Hex",
                   "Sample Text", "Note"]
        ws.append(headers)
        header_row = ws.max_row
        for c in ws[header_row]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="4472C4")
            c.alignment = Alignment(horizontal="center")

        for r in sorted(data_rows, key=lambda r: (r["y"], r["x"])):
            row_vals = [r["id"], r["tag"], r["label"], r["x"], r["y"], r["w"], r["h"],
                        r["x_pt"], r["y_pt"], r["w_pt"], r["h_pt"], r["bg_hex"],
                        r["font_family"], r["font_size_px"], r["font_size_pt"],
                        r["font_color_hex"], r["sample_text"], r["note"]]
            ws.append(row_vals)
            excel_row = ws.max_row
            if r["bg_hex"]:
                try:
                    ws.cell(row=excel_row, column=12).fill = PatternFill("solid", fgColor=r["bg_hex"].lstrip("#"))
                except Exception:
                    pass
            if r["note"]:
                for col_idx in range(1, len(headers) + 1):
                    ws.cell(row=excel_row, column=col_idx).fill = PatternFill("solid", fgColor="FFC7CE")

        for i, h in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(i)].width = max(12, len(h) + 2)
        return ws

    panels_title = "Panels & Cards" if text_rows is not None else "Visuals (As Detected)"
    build_visuals_sheet(panels_title, rows)

    if text_rows:
        build_visuals_sheet("Text Elements (Labels & Values)", text_rows,
                             note_text=("Every OCR-detected text block — nav items, KPI labels/values, chart "
                                        "titles, filters, headings — with estimated font size/color. "
                                        "Cross-check 'Sample Text' against the source; small dashboard fonts "
                                        "can be misread by OCR."))

    # --- Auto-Fit Layout sheet (always present; scale=1 means no change was needed) ---
    if scale < 0.999:
        note = (f"Scaled to {scale:.1%} of original size so every panel fits within the "
                f"{canvas['w']}x{canvas['h']}px canvas — no scrolling/clipping. "
                f"Canvas size and background are unchanged; only panel positions/sizes were shrunk proportionally.")
    else:
        note = "Everything already fit within the canvas — no scaling was needed."
    build_visuals_sheet("Auto-Fit Layout (No Scroll)", fitted_rows, note_text=note)
    wb.save(out_path)
    return out_path


def scale_sections(sections, scale):
    if scale >= 0.999:
        return sections
    scaled = []
    for sec in sections:
        new_widgets = []
        for w in sec["widgets"]:
            nw = dict(w)
            nw["x"] = round(w["x"] * scale)
            nw["y"] = round(w["y"] * scale)
            nw["w"] = round(w["w"] * scale)
            nw["h"] = round(w["h"] * scale)
            nw["x_pt"] = px_to_pt(nw["x"])
            nw["y_pt"] = px_to_pt(nw["y"])
            nw["w_pt"] = px_to_pt(nw["w"])
            nw["h_pt"] = px_to_pt(nw["h"])
            if w.get("font_size_px") not in ("", None):
                nw["font_size_px"] = round(w["font_size_px"] * scale, 1)
                nw["font_size_pt"] = px_to_pt(nw["font_size_px"])
            if w.get("bg_x") is not None:
                nw["bg_x"] = round(w["bg_x"] * scale)
                nw["bg_y"] = round(w["bg_y"] * scale)
                nw["bg_w"] = round(w["bg_w"] * scale)
                nw["bg_h"] = round(w["bg_h"] * scale)
                nw["bg_x_pt"] = px_to_pt(nw["bg_x"])
                nw["bg_y_pt"] = px_to_pt(nw["bg_y"])
                nw["bg_w_pt"] = px_to_pt(nw["bg_w"])
                nw["bg_h_pt"] = px_to_pt(nw["bg_h"])
            new_widgets.append(nw)
        scaled.append({"name": sec["name"], "widgets": new_widgets})
    return scaled


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------

def _process_page(canvas, rows, text_rows, is_html, screenshot_path, work_dir, no_images, label=None, is_shell=False,
                  crop_subdir=None):
    """Run the full sections/blocks/visuals pipeline for one page's worth
    of extracted rows, returning a dict ready to hand to generate_design_doc.

    crop_subdir: unique per-page folder name for cropped snapshots. Visual
    uids restart at w0/blk1 on every page, so without a per-page folder each
    page's crops OVERWRITE the previous page's files -- and since images are
    embedded only after all pages are processed, every page's document would
    end up showing the LAST-processed page's snapshots (the shuffled-images
    bug)."""
    scale, _ = fit_to_canvas(canvas, rows)
    overflow_count = len([r for r in rows if r["note"]])

    original_sections = build_sections(canvas, rows, text_rows, already_widgets=is_html)
    if is_html:
        original_sections = resolve_html_nesting(original_sections)
    else:
        original_sections = merge_overlapping_blocks(original_sections)
        original_sections = merge_same_block_fragments(original_sections)
    original_sections = dedupe_cross_section_blocks(original_sections)
    assign_block_groups(original_sections)

    image_map = {}
    block_image_map = {}
    if not no_images and screenshot_path:
        crop_dir = os.path.join(work_dir, crop_subdir) if crop_subdir else work_dir
        image_map = crop_widget_images(screenshot_path, original_sections, crop_dir, canvas)
        block_image_map = crop_block_images(screenshot_path, original_sections, crop_dir, canvas)

    sections = scale_sections(original_sections, scale)

    return {
        "label": label,
        "is_shell": is_shell,
        "canvas": canvas,
        "sections": sections,
        "overflow_count": overflow_count,
        "scale": scale,
        "hero_image_path": screenshot_path,
        "image_map": image_map,
        "block_image_map": block_image_map,
    }


def main():
    ap = argparse.ArgumentParser(description="Analyze a wireframe (HTML or image) and produce a professional design specification document.")
    ap.add_argument("input", help="Path to .html/.htm file or .png/.jpg/.jpeg image")
    ap.add_argument("--canvas", default=None, help="Canvas size for HTML rendering, e.g. 1920x1080. "
                    "If omitted, the real page size is auto-detected (recommended -- works for both "
                    "fixed-size dashboards and tall, naturally-scrolling report pages).")
    ap.add_argument("--out", default=None, help="Exact output .docx path. Overrides --outdir/auto-naming if given.")
    ap.add_argument("--outdir", default=None, help="Folder to write the output into, using the auto-generated "
                    "name <input>_design_spec.docx. Created if it doesn't exist. Ignored if --out is given.")
    ap.add_argument("--xlsx", action="store_true", help="Also generate the old-style Excel spec sheet alongside the design document (single-page documents only)")
    ap.add_argument("--no-images", action="store_true", help="Skip embedding cropped screenshots of each visual (faster, smaller file)")
    ap.add_argument("--no-multipage", action="store_true", help="Disable automatic multi-page/tab detection for HTML -- always treat the file as a single page")
    args = ap.parse_args()

    if not os.path.isfile(args.input):
        print(f"Input file not found: {args.input}")
        sys.exit(1)

    ext = os.path.splitext(args.input)[1].lower()
    is_html = ext in (".html", ".htm")

    if args.canvas:
        cw, ch = (int(v) for v in args.canvas.lower().split("x"))
    elif is_html:
        cw, ch = detect_html_canvas_size(args.input)
        print(f"Auto-detected canvas size: {cw} x {ch} px")
    else:
        cw, ch = 1280, 720  # unused for image mode (real image dimensions are read directly)

    if args.out:
        out_path = args.out
    else:
        base_name = os.path.splitext(os.path.basename(args.input))[0] + "_design_spec.docx"
        out_dir = args.outdir or os.path.dirname(os.path.abspath(args.input))
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, base_name)

    work_dir = os.path.splitext(out_path)[0] + "_assets"

    text_rows = None
    pages_data = None  # set if multi-page HTML detected
    screenshot_path = None

    if is_html:
        mode = "HTML (exact)"
        if not args.no_multipage:
            probe_work_dir = work_dir if not args.no_images else None
            pages_data = detect_and_extract_pages(args.input, cw, ch, work_dir=probe_work_dir)

        if pages_data:
            print(f"Detected {len(pages_data)} pages (tabs) in this HTML file:")
            for pd in pages_data:
                print(f"  - {pd['label']}")
            pages_data, common_shell = extract_common_elements(pages_data, cw)
            if common_shell:
                print(f"  Found {len(common_shell['rows'])} elements common to every page "
                      f"(header/nav/filters) -- documented once instead of per page.")
        else:
            common_shell = None
            if not args.no_images:
                os.makedirs(work_dir, exist_ok=True)
                screenshot_path = os.path.join(work_dir, "_full_page.png")
            canvas, rows = analyze_html(args.input, cw, ch, screenshot_path=screenshot_path)
    elif ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp"):
        mode = "Image (estimated)"
        canvas, rows, text_rows = analyze_image(args.input, cw, ch)
        screenshot_path = args.input
        common_shell = None
    else:
        print(f"Unsupported file type: {ext}. Use .html/.htm or .png/.jpg/.jpeg")
        sys.exit(1)

    if pages_data:
        pages = []
        for p_i, pd in enumerate(pages_data, 1):
            safe = "".join(c if c.isalnum() else "_" for c in (pd["label"] or ""))[:30]
            pg = _process_page(pd["canvas"], pd["rows"], None, True, pd["screenshot_path"],
                                work_dir, args.no_images, label=pd["label"],
                                crop_subdir=f"page_{p_i}_{safe}")
            if pd.get("dedup_notes"):
                pg["dedup_notes"] = pd["dedup_notes"]
            pages.append(pg)
        if common_shell and common_shell["rows"]:
            shell_page = _process_page(common_shell["canvas"], common_shell["rows"], None, True,
                                        common_shell["screenshot_path"], work_dir, args.no_images,
                                        label=common_shell["label"], is_shell=True,
                                        crop_subdir="common_elements")
            pages = [shell_page] + pages
    else:
        pages = [_process_page(canvas, rows, text_rows, is_html, screenshot_path,
                                work_dir, args.no_images, label=None)]

    generate_design_doc(pages, out_path, args.input, mode)

    total_sections = sum(len(pg["sections"]) for pg in pages)
    total_widgets = sum(len(s["widgets"]) for pg in pages for s in pg["sections"])
    total_overflow = sum(pg["overflow_count"] for pg in pages)
    print(f"\nMode: {mode}")
    if len(pages) > 1:
        print(f"Pages documented: {len(pages)}")
    print(f"Sections detected: {total_sections}")
    print(f"Total visuals documented: {total_widgets}")
    if total_overflow:
        print(f"[!] {total_overflow} visual(s) exceeded the canvas in raw detection across the document -- auto-fit scale applied per page.")
    else:
        print("Everything already fits within its canvas -- no scaling needed.")
    print(f"Design document written to: {out_path}")

    if args.xlsx:
        if pages_data:
            print("Note: --xlsx is only supported for single-page documents; skipping Excel export for this multi-page file.")
        else:
            xlsx_path = os.path.splitext(out_path)[0] + ".xlsx"
            write_report(canvas, rows, xlsx_path, args.input, mode, text_rows=text_rows)
            print(f"Excel spec sheet also written to: {xlsx_path}")


if __name__ == "__main__":
    main()
