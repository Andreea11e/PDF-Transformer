# FruitAI 8.0 pentru web

Versiunea JavaScript a aplicației FruitAI (`main.py`). Recunoaște fructele cu camera sau dintr-o fotografie și afișează valorile nutriționale (USDA, per porție).

## Ce conține

| Fișier | Rol |
|---|---|
| `index.html` | Tot site-ul: interfața, stilurile și codul JavaScript |
| `start-mac-linux.sh` | Pornește un server local și deschide site-ul (macOS / Linux) |
| `start-windows.bat` | Același lucru pe Windows |

## Cum îl pornești

Browserele permit camera doar pe `https://` sau pe `http://localhost`, deci **nu deschide `index.html` cu dublu-click**. Pornește-l printr-un server local:

- **macOS / Linux:** în Terminal, în acest dosar: `sh start-mac-linux.sh`
- **Windows:** dublu-click pe `start-windows.bat`
- **Manual:** `python3 -m http.server 8000`, apoi deschide http://localhost:8000

Ai nevoie doar de Python 3 instalat, fără alte pachete.

## Cum funcționează

- **Detecție:** OWL-ViT (`Xenova/owlvit-base-patch32`) găsește fructul în cadru.
- **Clasificare:** CLIP (`Xenova/clip-vit-base-patch32`) alege dintre 61 de fructe și 58 de etichete care nu sunt fructe.
- Modelele rulează **în browser** prin transformers.js (WebGPU dacă există, altfel WASM). Prima dată se descarcă aproximativ 240 MB, după care rămân în cache. Imaginile nu sunt trimise nicăieri.
- Pragurile sunt cele din `main.py`: detector 0,10, fruct 0,30, cadru întreg 0,60, o analiză la 0,8 s, iar un rezultat rămâne afișat 4 s.
- **Nutriție:** valorile USDA FoodData Central (SR Legacy) sunt incluse direct în cod, așa că nu e nevoie de cheie API. Sunt orientative. Pentru moșmon nu există date USDA.

## Publicare online

Poți urca `index.html` pe orice hosting static cu HTTPS: GitHub Pages, Netlify, Vercel sau Cloudflare Pages. Nu are nevoie de server propriu.

## Fișierele vechi

`yolov8n.pt` și `yolov8s-worldv2.pt` nu sunt necesare: `main.py` nu le folosea, iar formatul PyTorch nu rulează în browser. Nici `fruitai_cache.json` nu mai e necesar, pentru că valorile nutriționale sunt acum în cod.
