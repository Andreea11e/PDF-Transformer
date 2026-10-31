import json
import sys
import threading
import time
import os
from pathlib import Path

import cv2
import numpy as np
import requests
import torch
from PIL import Image, ImageDraw, ImageFont
from transformers import (
    CLIPModel,
    CLIPProcessor,
    OwlViTForObjectDetection,
    OwlViTProcessor,
)

BASE_DIR = Path(__file__).resolve().parent
CACHE_FILE = BASE_DIR / "fruitai_cache.json"

CLIP_NAME = "openai/clip-vit-base-patch32"
DETECTOR_NAME = "google/owlvit-base-patch32"
CAMERA_INDEX = 0
CHECK_INTERVAL = 0.8
RESULT_TTL = 4.0            # secunde cat ramane valabil un rezultat
DETECT_THRESHOLD = 0.10     # prag detector pentru "fruct"
FRUIT_MIN_SCORE = 0.30      # prag CLIP pentru a accepta un fruct
FULLFRAME_MIN_SCORE = 0.60  # prag cand fructul acopera tot cadrul

NUTRITION_URL = "https://api.nal.usda.gov/fdc/v1/foods/search"
USDA_KEY = os.getenv("FDC_API_KEY", "DEMO_KEY")

FRUITS = {
    "apple": "Măr", "banana": "Banană", "orange": "Portocală",
    "pear": "Pară", "peach": "Piersică", "mango": "Mango",
    "pineapple": "Ananas", "strawberry": "Căpșună",
    "grape": "Struguri", "watermelon": "Pepene verde",
    "lemon": "Lămâie", "lime": "Lime", "kiwi": "Kiwi",
    "plum": "Prună", "cherry": "Cireașă",
    "coconut": "Nucă de cocos", "papaya": "Papaya",
    "pomegranate": "Rodie", "avocado": "Avocado",
    "apricot": "Caisă", "blueberry": "Afină",
    "raspberry": "Zmeură", "blackberry": "Mură",
    "grapefruit": "Grepfrut", "fig": "Smochină", "guava": "Guava",
    "dragon fruit": "Fructul dragonului",
    "passion fruit": "Fructul pasiunii", "lychee": "Litchi",
    "nectarine": "Nectarină", "cranberry": "Merișor",
    "tangerine": "Mandarină", "clementine": "Clementină",
    "persimmon": "Curmal japonez", "date": "Curmală",
    "durian": "Durian", "jackfruit": "Jackfruit",
    "starfruit": "Carambolă", "longan": "Longan",
    "rambutan": "Rambutan", "quince": "Gutui", "mulberry": "Dudă",
    "gooseberry": "Agrișă", "pomelo": "Pomelo",
    "kumquat": "Kumquat", "tamarind": "Tamarind",
    "cantaloupe": "Pepene galben", "honeydew": "Pepene honeydew",
    "plantain": "Banană plantain", "soursop": "Guanabana",
    "custard apple": "Anona", "elderberry": "Soc",
    "mangosteen": "Mangostan", "blackcurrant": "Coacăză neagră",
    "redcurrant": "Coacăză roșie", "greengage": "Reine Claude",
    "blood orange": "Portocală roșie", "prickly pear": "Smochin de India",
    "cherimoya": "Cherimoya", "medlar": "Moșmon", "olive": "Măslină",
}

# Etichete "negative": ajuta modelul sa spuna ca NU e fruct.
NON_FRUITS = {
    "person": "o persoană", "human face": "o față",
    "human hand": "o mână", "man": "un bărbat", "woman": "o femeie",
    "child": "un copil", "head and shoulders of a person": "o persoană",
    "cat": "o pisică", "dog": "un câine",
    "mobile phone": "un telefon", "laptop computer": "un laptop",
    "computer keyboard": "o tastatură", "computer monitor": "un monitor",
    "bottle": "o sticlă", "cup": "o cană", "mug": "o cană",
    "glass": "un pahar", "plate": "o farfurie", "bowl": "un bol",
    "book": "o carte", "pen": "un pix", "remote control": "o telecomandă",
    "chair": "un scaun", "table": "o masă", "sofa": "o canapea",
    "bed": "un pat", "door": "o ușă", "window": "o fereastră",
    "wall": "un perete", "ceiling": "un tavan", "room": "o cameră",
    "curtain": "o perdea", "lamp": "o lampă", "poster": "un poster",
    "backpack": "un rucsac", "clothes": "haine", "shirt": "o cămașă",
    "toy": "o jucărie", "ball": "o minge", "plant": "o plantă",
    "flower": "o floare", "tree": "un copac",
    "tomato": "o roșie (legumă)", "cucumber": "un castravete (legumă)",
    "potato": "un cartof (legumă)", "carrot": "un morcov (legumă)",
    "onion": "o ceapă (legumă)", "bell pepper": "un ardei (legumă)",
    "cabbage": "o varză (legumă)", "bread": "pâine",
    "cheese": "brânză", "egg": "un ou", "pizza": "o pizza",
    "blurry image": "o imagine neclară", "dark image": "o imagine întunecată",
    "empty background": "fundal gol",
}

NUTRIENTS = {
    "calories": ("Calorii", "kcal"), "protein": ("Proteine", "g"),
    "carbs": ("Carbohidrați", "g"), "fat": ("Grăsimi", "g"),
    "fiber": ("Fibre", "g"),
}
# ID-uri USDA -> cheie (datele "Foundation" pot da caloriile ca 2047/2048)
NUTRIENT_IDS = {1008: "calories", 2047: "calories", 2048: "calories",
                1003: "protein", 1005: "carbs", 1004: "fat", 1079: "fiber"}

USDA_NAMES = {
    "apple": "apples raw with skin",
    "orange": "oranges raw all commercial varieties",
    "blood orange": "oranges raw all commercial varieties",
    "lemon": "lemons raw without peel",
    "grape": "grapes red or green raw", "strawberry": "strawberries raw",
    "blueberry": "blueberries raw", "raspberry": "raspberries raw",
    "blackberry": "blackberries raw", "cherry": "sweet cherries raw",
    "kiwi": "kiwifruit raw", "coconut": "coconut meat raw",
    "dragon fruit": "dragon fruit raw",
    "passion fruit": "passion fruit raw",
    "watermelon": "watermelon raw", "cantaloupe": "cantaloupe raw",
    "honeydew": "honeydew melon raw",
}

lock = threading.RLock()
cache = {}


def normalize_entry(entry):
    """Acceptă și formatul vechi (chei "1008") și cel nou ("calories")."""
    if not isinstance(entry, dict) or not isinstance(entry.get("values"), dict):
        return None
    values = {}
    for key, value in entry["values"].items():
        name = NUTRIENT_IDS.get(int(key)) if str(key).isdigit() else key
        if name in NUTRIENTS and value is not None and name not in values:
            values[name] = float(value)
    if "calories" not in values:
        return None  # intrare incompletă -> se cere din nou de la USDA
    return {"food": entry.get("food", "USDA"), "values": values,
            "basis": "100 g", "source": "USDA FoodData Central"}


def load_cache():
    try:
        data = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            clean = {k: normalize_entry(v) for k, v in data.items()}
            return {k: v for k, v in clean.items() if v}
    except (OSError, json.JSONDecodeError):
        pass
    return {}


def save_cache():
    try:
        temp = CACHE_FILE.with_suffix(".tmp")
        temp.write_text(json.dumps(cache, ensure_ascii=False, indent=2),
                        encoding="utf-8")
        temp.replace(CACHE_FILE)
    except OSError as exc:
        print("Eroare la salvarea cache-ului:", exc)


def nutrition_for(fruit):
    with lock:
        if fruit in cache:
            return cache[fruit]

    query = USDA_NAMES.get(fruit, fruit + " raw")
    response = requests.get(
        NUTRITION_URL,
        params={"api_key": USDA_KEY, "query": query, "pageSize": 20,
                "dataType": ["Foundation", "SR Legacy"]},
        timeout=20,
    )
    response.raise_for_status()
    foods = response.json().get("foods", [])
    if not foods:
        raise ValueError("USDA nu a găsit alimentul.")

    def food_score(food):
        desc = food.get("description", "").lower()
        score = 0
        if "raw" in desc:
            score += 5
        bad = ("juice", "dried", "canned", "frozen", "syrup", "sweetened",
               "cooked", "leaves", "peel", "seed", "nectar", "baby", "pie",
               "jam", "candied")
        if any(word in desc for word in bad):
            score -= 6
        if fruit.replace(" ", "") in desc.replace(" ", ""):
            score += 3
        if food.get("dataType") == "SR Legacy":
            score += 1  # SR Legacy are valori complete
        return score

    food = max(foods, key=food_score)
    values = {}
    for item in food.get("foodNutrients", []):
        name = NUTRIENT_IDS.get(item.get("nutrientId"))
        value = item.get("value")
        if name and value is not None and (
                name not in values or item.get("nutrientId") == 1008):
            values[name] = value

    result = {"food": food.get("description", query),
              "values": values, "basis": "100 g",
              "source": "USDA FoodData Central"}
    with lock:
        cache[fruit] = result
        save_cache()
    return result


def get_font(size):
    paths = [
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for path in paths:
        try:
            if Path(path).exists():
                return ImageFont.truetype(path, size)
        except OSError:
            pass
    return ImageFont.load_default()


FONT_TITLE = get_font(24)
FONT_TEXT = get_font(18)
FONT_SMALL = get_font(15)


class FruitAI:
    def __init__(self):
        self.device = (
            "mps" if torch.backends.mps.is_available()
            else "cuda" if torch.cuda.is_available()
            else "cpu"
        )
        # Detectorul ruleaza pe CPU/CUDA (MPS poate da erori la OWL-ViT).
        self.det_device = "cuda" if torch.cuda.is_available() else "cpu"

        self.result = None   # dict cu rezultatul curent
        self.nutrition = None
        self.status = "Se încarcă modelele AI..."
        self.busy = False
        self.last_check = 0.0
        self.lock = threading.RLock()

        print("Se încarcă CLIP...")
        self.processor = CLIPProcessor.from_pretrained(CLIP_NAME)
        self.model = CLIPModel.from_pretrained(CLIP_NAME)
        self.model.to(self.device).eval()

        print("Se încarcă detectorul OWL-ViT (prima dată se descarcă)...")
        self.det_processor = OwlViTProcessor.from_pretrained(DETECTOR_NAME)
        self.detector = OwlViTForObjectDetection.from_pretrained(
            DETECTOR_NAME)
        self.detector.to(self.det_device).eval()

        self.fruit_labels = list(FRUITS.keys())
        self.labels = self.fruit_labels + list(NON_FRUITS.keys())

        text_inputs = self.processor(
            text=[f"a photo of a {label}" for label in self.labels],
            padding=True, truncation=True, return_tensors="pt",
        ).to(self.device)
        with torch.inference_mode():
            feats = self.model.get_text_features(**text_inputs)
            self.text_features = torch.nn.functional.normalize(feats, dim=-1)

        # Interogari pentru detector: prima = fruct, restul = "negative".
        self.det_queries = [
            "a photo of a fruit",
            "a photo of a person",
            "a photo of a human face",
            "a photo of a hand",
        ]
        self.status = "Modele pregătite"
        print("Pregătit. CLIP pe:", self.device, "| detector pe:",
              self.det_device)

    # ---------- detector: gaseste fructul in cadru ----------
    def find_fruit_box(self, frame_bgr):
        h, w = frame_bgr.shape[:2]
        scale = 640 / max(h, w)
        small = cv2.resize(frame_bgr, (int(w * scale), int(h * scale)))
        image = Image.fromarray(cv2.cvtColor(small, cv2.COLOR_BGR2RGB))

        inputs = self.det_processor(
            text=[self.det_queries], images=image, return_tensors="pt"
        ).to(self.det_device)

        with torch.inference_mode():
            outputs = self.detector(**inputs)

        probs = outputs.logits[0].sigmoid()      # (patch-uri, interogari)
        boxes = outputs.pred_boxes[0]            # cx, cy, w, h normalizate

        fruit = probs[:, 0]
        others = probs[:, 1:].max(dim=1).values
        valid = fruit > others                    # exclude persoane/maini
        scores = torch.where(valid, fruit, torch.zeros_like(fruit))

        idx = int(scores.argmax().item())
        best = float(scores[idx].item())
        if best < DETECT_THRESHOLD:
            return None, best

        cx, cy, bw, bh = [float(v) for v in boxes[idx]]
        x1 = int(max(0, (cx - bw / 2)) * w)
        y1 = int(max(0, (cy - bh / 2)) * h)
        x2 = int(min(1, (cx + bw / 2)) * w)
        y2 = int(min(1, (cy + bh / 2)) * h)

        # Margine mica in jurul fructului.
        mx = int((x2 - x1) * 0.12)
        my = int((y2 - y1) * 0.12)
        x1, y1 = max(0, x1 - mx), max(0, y1 - my)
        x2, y2 = min(w, x2 + mx), min(h, y2 + my)

        if x2 - x1 < 40 or y2 - y1 < 40:
            return None, best
        return (x1, y1, x2, y2), best

    # ---------- CLIP: ce este in imagine ----------
    def classify(self, crop_bgr):
        image = Image.fromarray(cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2RGB))
        inputs = self.processor(images=image, return_tensors="pt").to(
            self.device)
        with torch.inference_mode():
            feats = self.model.get_image_features(**inputs)
            feats = torch.nn.functional.normalize(feats, dim=-1)
            logits = self.model.logit_scale.exp() * (
                feats @ self.text_features.T)
            probs = logits.softmax(dim=-1)[0]
            index = int(probs.argmax().item())
            score = float(probs[index].item())
        return self.labels[index], score

    def worker(self, frame):
        try:
            box, det_score = self.find_fruit_box(frame)
            now = time.monotonic()

            if box is not None:
                x1, y1, x2, y2 = box
                label, score = self.classify(frame[y1:y2, x1:x2])

                if label in FRUITS and score >= FRUIT_MIN_SCORE:
                    self.set_fruit(label, score, box, now)
                elif label in NON_FRUITS:
                    self.set_result({
                        "kind": "notfruit", "label": label,
                        "score": score, "box": box, "time": now,
                    }, f"Nu este un fruct: {NON_FRUITS[label]}")
                else:
                    self.set_result({"kind": "none", "time": now},
                                    "Nu sunt sigur ce este")
                return

            # Niciun fruct gasit: verificam tot cadrul.
            label, score = self.classify(frame)
            h, w = frame.shape[:2]
            if label in FRUITS and score >= FULLFRAME_MIN_SCORE:
                self.set_fruit(label, score, (0, 0, w, h), now)
            elif label in NON_FRUITS and score >= 0.25:
                self.set_result({
                    "kind": "notfruit", "label": label,
                    "score": score, "box": None, "time": now,
                }, f"Nu văd niciun fruct – văd {NON_FRUITS[label]}")
            else:
                self.set_result({"kind": "none", "time": now},
                                "Nu văd niciun fruct")

        except Exception as exc:
            print("EROARE:", repr(exc))
            with self.lock:
                self.status = "Eroare AI: " + str(exc)[:100]
        finally:
            with self.lock:
                self.busy = False

    def set_result(self, result, status):
        with self.lock:
            self.result = result
            self.nutrition = None
            self.status = status

    def set_fruit(self, fruit, score, box, now):
        with self.lock:
            same = (self.result and self.result.get("kind") == "fruit"
                    and self.result.get("label") == fruit)
            self.result = {"kind": "fruit", "label": fruit,
                           "score": score, "box": box, "time": now}
            if not same:
                self.nutrition = None
            self.status = f"Identificat: {FRUITS[fruit]} ({score:.0%})"

        if same and self.nutrition:
            return
        try:
            data = nutrition_for(fruit)
            with self.lock:
                if self.result and self.result.get("label") == fruit:
                    self.nutrition = data
        except Exception as exc:
            print("Eroare USDA:", repr(exc))
            with self.lock:
                self.status = "Fruct identificat; nutriție indisponibilă"

    def process(self, frame):
        now = time.monotonic()
        with self.lock:
            if self.busy or now - self.last_check < CHECK_INTERVAL:
                return
            self.busy = True
            self.last_check = now
        threading.Thread(target=self.worker, args=(frame.copy(),),
                         daemon=True).start()


def draw_interface(frame, ai):
    rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    image = Image.fromarray(rgb)
    draw = ImageDraw.Draw(image)

    with ai.lock:
        result = ai.result
        nutrition = ai.nutrition
        status = ai.status
        busy = ai.busy

    if result and time.monotonic() - result["time"] > RESULT_TTL:
        result = None

    width, height = image.size
    kind = result["kind"] if result else None

    # Dreptunghiul pe fructul gasit automat.
    if result and result.get("box") and kind in ("fruit", "notfruit"):
        x1, y1, x2, y2 = result["box"]
        color = (70, 230, 100) if kind == "fruit" else (240, 80, 80)
        draw.rectangle((x1, y1, x2, y2), outline=color, width=4)
        tag = (FRUITS[result["label"]] if kind == "fruit"
               else NON_FRUITS[result["label"]])
        draw.rectangle((x1, max(0, y1 - 28), x1 + 12 + 10 * len(tag),
                        max(0, y1 - 28) + 28), fill=color)
        draw.text((x1 + 6, max(0, y1 - 26)), tag,
                  font=FONT_SMALL, fill=(0, 0, 0))

    panel_bottom = min(height - 10, 345 if kind == "fruit" else 140)
    draw.rounded_rectangle(
        (10, 10, min(width - 10, 440), panel_bottom), radius=14,
        fill=(25, 36, 30), outline=(100, 225, 120), width=2)

    x, y = 25, 20
    draw.text((x, y), "FruitAI 8.0", font=FONT_TITLE, fill=(140, 255, 155))
    y += 40

    if kind == "fruit":
        label = result["label"]
        draw.text((x, y), "Română: " + FRUITS[label],
                  font=FONT_TEXT, fill=(255, 255, 255))
        y += 30
        draw.text((x, y), "English: " + label.title(),
                  font=FONT_TEXT, fill=(220, 240, 220))
        y += 30
        draw.text((x, y), f"Scor de clasificare: {result['score']:.0%}",
                  font=FONT_SMALL, fill=(220, 240, 220))
        y += 27
        draw.text((x, y), "Nutriție estimată / 100 g",
                  font=FONT_TEXT, fill=(140, 255, 155))
        y += 29
        if nutrition:
            for key, (nlabel, unit) in NUTRIENTS.items():
                value = nutrition["values"].get(key)
                text = (f"{nlabel}: {float(value):g} {unit}"
                        if value is not None
                        else f"{nlabel}: indisponibil")
                draw.text((x, y), text, font=FONT_SMALL,
                          fill=(255, 255, 255))
                y += 22
                if y > panel_bottom - 20:
                    break
            if y < panel_bottom - 18:
                draw.text((x, y), "USDA: " + nutrition["food"][:38],
                          font=FONT_SMALL, fill=(190, 210, 190))
        else:
            draw.text((x, y), "Se caută valorile nutriționale...",
                      font=FONT_SMALL, fill=(255, 220, 100))
    elif kind == "notfruit":
        draw.text((x, y), "Acesta nu este un fruct",
                  font=FONT_TEXT, fill=(255, 130, 130))
        y += 30
        draw.text((x, y), "Văd: " + NON_FRUITS[result["label"]],
                  font=FONT_SMALL, fill=(255, 255, 255))
    else:
        draw.text((x, y), ("Caut un fruct..." if busy else status[:60]),
                  font=FONT_TEXT, fill=(255, 255, 255))
        y += 32
        draw.text((x, y), "Arată un fruct oriunde în cadru.",
                  font=FONT_SMALL, fill=(220, 240, 220))

    draw.text((15, height - 28),
              "Q = ieșire | Detector și clasificare locale, "
              "internet doar pentru USDA",
              font=FONT_SMALL, fill=(255, 255, 255))

    return cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)


def open_camera():
    backends = []
    if sys.platform == "darwin":
        backends.append(cv2.CAP_AVFOUNDATION)
    backends.append(cv2.CAP_ANY)

    for backend in backends:
        for index in (CAMERA_INDEX, 1, 2):
            cap = cv2.VideoCapture(index, backend)
            if not cap.isOpened():
                cap.release()
                continue
            cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
            cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
            for _ in range(30):  # încălzire: primele cadre pot fi negre
                ok, frame = cap.read()
                if ok and frame is not None and frame.mean() > 3:
                    return cap
                time.sleep(0.05)
            return cap
    return None


def main():
    cache.update(load_cache())

    try:
        ai = FruitAI()
    except Exception as exc:
        print("Nu s-au putut încărca modelele:", repr(exc))
        print("Verifică internetul, spațiul pe disc și instalarea "
              "transformers/torch.")
        return

    camera = open_camera()
    if camera is None:
        print("Camera nu se deschide. Verifică permisiunile "
              "(System Settings → Privacy & Security → Camera).")
        return

    print("FruitAI rulează. Apasă Q pentru a ieși.")

    try:
        while True:
            ok, frame = camera.read()
            if not ok or frame is None:
                time.sleep(0.02)
                continue

            ai.process(frame)
            cv2.imshow("FruitAI 8.0", draw_interface(frame, ai))

            if cv2.waitKey(1) & 0xFF == ord("q"):
                break
    finally:
        camera.release()
        cv2.destroyAllWindows()
        print("FruitAI oprit.")


if __name__ == "__main__":
    main()