# FruitLab AR — site + aplicație Python

Site-ul face același lucru ca `main.py`, dar direct în browser (fără server de AI):

1. **OWL-ViT** caută *unde* este fructul în cadru și îl compară cu „persoană / față / mână”.
2. **CLIP** decide *ce* fruct este (60+ fructe) sau spune că **nu e fruct** (cană, telefon, legumă…).
3. Dacă detectorul nu găsește nimic, CLIP verifică tot cadrul (prag 60%).
4. Afișează numele în **română + engleză**, scorul și **nutriția / 100 g** (USDA).

Aceleași praguri ca în Python: detector 0.10, fruct pe decupaj 0.30, tot cadrul 0.60, analiză la 0.8 s, rezultat valabil 4 s.

## Cum pornești site-ul

Camera NU merge dacă deschizi `index.html` cu dublu-click. Pornește un server local în folderul acesta:

```bash
python -m http.server 8000
```

Apoi deschide **http://localhost:8000** (Chrome sau Edge recomandat).
Pe telefon sau online trebuie HTTPS (ex. GitHub Pages, Netlify).

La prima scanare se descarcă modelele (~300 MB, o singură dată; apoi rămân în cache-ul browserului).

## Fișiere

| Fișier | Rol |
|---|---|
| `index.html`, `scan.html`, `ghid.html`, `joaca.html`, `fructopedia.html`, `pasaport.html` | paginile site-ului |
| `script.js` | scannerul AI (echivalentul lui `main.py`) |
| `site.js` | meniu, quiz, fructopedia, pașaport |
| `styles.css` | stilurile |
| `fruitai_cache.json` | valori nutriționale salvate (folosite și de site, și de Python) |
| `main.py` | aplicația desktop originală (OpenCV + PyTorch) |

## Nutriție (USDA)

Se folosește `DEMO_KEY`, care are o limită de cereri pe oră. Pentru mai multe, ia o cheie gratuită de la
https://api.data.gov/signup/ și pune-o în `script.js` (`USDA_KEY`) sau în variabila `FDC_API_KEY` pentru Python.

## Python

```bash
pip install torch transformers opencv-python pillow requests numpy
python main.py
```
