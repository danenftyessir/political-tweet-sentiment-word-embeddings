# -*- coding: utf-8 -*-
"""Screenshot sel notebook Praktikum 2 (final) dari HTML nbconvert classic."""
import pathlib
from playwright.sync_api import sync_playwright

HTML = pathlib.Path(r"C:\Users\HYPERS~1\AppData\Local\Temp\nb_final.html").as_uri()
OUT = pathlib.Path(__file__).parent / "img"
OUT.mkdir(exist_ok=True)

# nama file -> (teks unik penemu sel, "full"=kode+output | "input"=kode saja)
targets = {
    "cell_preprocess.png":     ("def bersihkan_teks", "input"),   # kode preprocessing
    "cell_embed_compare.png":  ("embedding_matrices = {", "full"),# kode + output perbandingan embedding
    "cell_lstm.png":           ("def bangun_lstm", "input"),      # kode arsitektur LSTM
    "cell_infer.png":          ("peta_exact", "input"),           # kode inferensi final
}
css = """
#notebook-container{width:1340px!important;max-width:none!important;box-shadow:none!important;
  margin:0!important;padding:14px!important;border:none!important;}
body,div#notebook{background:#ffffff!important;}
div.output_subarea{max-width:none!important;}
div.prompt{min-width:82px!important;}
"""
with sync_playwright() as p:
    browser = p.chromium.launch(channel="chrome")
    ctx = browser.new_context(viewport={"width": 1420, "height": 1400}, device_scale_factor=2)
    page = ctx.new_page()
    page.goto(HTML)
    page.add_style_tag(content=css)
    page.wait_for_timeout(500)
    for fn, (txt, mode) in targets.items():
        cell = page.locator("div.code_cell", has_text=txt).first
        target = cell if mode == "full" else cell.locator("div.input")
        target.scroll_into_view_if_needed()
        page.wait_for_timeout(150)
        target.screenshot(path=str(OUT / fn))
        box = target.bounding_box()
        print("saved %-22s %dx%d" % (fn, box["width"], box["height"]))
    browser.close()
print("DONE")
