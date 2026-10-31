// FruitLab AR — scanner în browser care face exact ce face main.py:
//   1) OWL-ViT caută "unde este fructul" (comparat patch cu patch cu persoană/față/mână)
//   2) CLIP decide "ce fruct este" (embedding-urile textului calculate O SINGURĂ DATĂ)
//   3) dacă detectorul nu găsește nimic, CLIP verifică tot cadrul (prag 60%)
//   4) valorile nutriționale vin din fruitai_cache.json / localStorage / USDA FoodData Central
// Totul rulează local în browser; internetul e folosit doar pentru modele (prima dată) și USDA.

const TRANSFORMERS_URL = 'https://cdn.jsdelivr.net/npm/@huggingface/transformers@3.8.1';
const CLIP_NAME = 'Xenova/clip-vit-base-patch32';
const DETECTOR_NAME = 'Xenova/owlvit-base-patch32';

// Aceleași praguri ca în main.py
const CHECK_INTERVAL = 800;        // ms între analize
const RESULT_TTL = 4000;           // cât rămâne valabil un rezultat
const DETECT_THRESHOLD = 0.10;     // prag detector pentru "fruct"
const FRUIT_MIN_SCORE = 0.30;      // prag CLIP pe decupaj
const FULLFRAME_MIN_SCORE = 0.60;  // prag CLIP pe tot cadrul
const OTHER_MIN_SCORE = 0.25;      // prag "nu e fruct" pe tot cadrul
const CLIP_LOGIT_SCALE = 100;      // exp(logit_scale) din CLIP ViT-B/32

const NUTRITION_URL = 'https://api.nal.usda.gov/fdc/v1/foods/search';
const USDA_KEY = 'DEMO_KEY';       // pune aici cheia ta gratuită de la api.data.gov dacă vrei
const NUTRITION_CACHE_KEY = 'fruitlab-usda-cache-v2';
const PASSPORT_KEY = 'fruitlab-explorer-passport-v1';
const LAST_FRUIT_KEY = 'fruitlab-last-fruit-v1';

const FRUITS = {
  apple:'Măr', banana:'Banană', orange:'Portocală', pear:'Pară', peach:'Piersică', mango:'Mango', pineapple:'Ananas', strawberry:'Căpșună', grape:'Struguri', watermelon:'Pepene verde', lemon:'Lămâie', lime:'Lime', kiwi:'Kiwi', plum:'Prună', cherry:'Cireașă', coconut:'Nucă de cocos', papaya:'Papaya', pomegranate:'Rodie', avocado:'Avocado', apricot:'Caisă', blueberry:'Afină', raspberry:'Zmeură', blackberry:'Mură', grapefruit:'Grepfrut', fig:'Smochină', guava:'Guava', 'dragon fruit':'Fructul dragonului', 'passion fruit':'Fructul pasiunii', lychee:'Litchi', nectarine:'Nectarină', cranberry:'Merișor', tangerine:'Mandarină', clementine:'Clementină', persimmon:'Curmal japonez', date:'Curmală', durian:'Durian', jackfruit:'Jackfruit', starfruit:'Carambolă', longan:'Longan', rambutan:'Rambutan', quince:'Gutui', mulberry:'Dudă', gooseberry:'Agrișă', pomelo:'Pomelo', kumquat:'Kumquat', tamarind:'Tamarind', cantaloupe:'Pepene galben', honeydew:'Pepene honeydew', plantain:'Banană plantain', soursop:'Guanabana', 'custard apple':'Anona', elderberry:'Soc', mangosteen:'Mangostan', blackcurrant:'Coacăză neagră', redcurrant:'Coacăză roșie', greengage:'Reine Claude', 'blood orange':'Portocală roșie', 'prickly pear':'Smochin de India', cherimoya:'Cherimoya', medlar:'Moșmon', olive:'Măslină'
};
// Etichete "negative": ajută modelul să spună că NU e fruct (identic cu main.py)
const NON_FRUITS = {
  person:'o persoană','human face':'o față','human hand':'o mână',man:'un bărbat',woman:'o femeie',child:'un copil','head and shoulders of a person':'o persoană',cat:'o pisică',dog:'un câine','mobile phone':'un telefon','laptop computer':'un laptop','computer keyboard':'o tastatură','computer monitor':'un monitor',bottle:'o sticlă',cup:'o cană',mug:'o cană',glass:'un pahar',plate:'o farfurie',bowl:'un bol',book:'o carte',pen:'un pix','remote control':'o telecomandă',chair:'un scaun',table:'o masă',sofa:'o canapea',bed:'un pat',door:'o ușă',window:'o fereastră',wall:'un perete',ceiling:'un tavan',room:'o cameră',curtain:'o perdea',lamp:'o lampă',poster:'un poster',backpack:'un rucsac',clothes:'haine',shirt:'o cămașă',toy:'o jucărie',ball:'o minge',plant:'o plantă',flower:'o floare',tree:'un copac',tomato:'o roșie (legumă)',cucumber:'un castravete (legumă)',potato:'un cartof (legumă)',carrot:'un morcov (legumă)',onion:'o ceapă (legumă)','bell pepper':'un ardei (legumă)',cabbage:'o varză (legumă)',bread:'pâine',cheese:'brânză',egg:'un ou',pizza:'o pizza','blurry image':'o imagine neclară','dark image':'o imagine întunecată','empty background':'fundal gol'
};
const FRUIT_IDS = Object.keys(FRUITS);
const LABELS = [...FRUIT_IDS, ...Object.keys(NON_FRUITS)];
// Prima interogare = fruct, restul = "negative" (identic cu main.py)
const DET_QUERIES = ['a photo of a fruit', 'a photo of a person', 'a photo of a human face', 'a photo of a hand'];

const ENGLISH = { grape:'Grapes', kiwi:'Kiwi fruit', starfruit:'Star fruit', honeydew:'Honeydew melon' };
const EMOJI = {apple:'🍎',banana:'🍌',orange:'🍊',strawberry:'🍓',grape:'🍇',watermelon:'🍉',lemon:'🍋',pineapple:'🍍',peach:'🍑',pear:'🍐',kiwi:'🥝',mango:'🥭',coconut:'🥥',avocado:'🥑',cherry:'🍒',blueberry:'🫐',olive:'🫒',tangerine:'🍊',clementine:'🍊','blood orange':'🍊',lime:'🍋',nectarine:'🍑',apricot:'🍑',cantaloupe:'🍈',honeydew:'🍈',plantain:'🍌'};
const FACTS = { apple:'Mărul face parte din familia trandafirilor.', banana:'Botanic, banana este o bacă.', strawberry:'Semințele de la exterior sunt fructele propriu-zise.', pineapple:'Ananasul este alcătuit din multe fructe mici unite.', watermelon:'Pepenele verde este înrudit cu castravetele.', avocado:'Avocado este un fruct, chiar dacă îl folosim adesea în mâncăruri sărate.', kiwi:'Miezul verde ascunde semințe mici, comestibile.', orange:'Portocala este un fruct citric, înrudit cu lămâia.', lemon:'Coaja lămâii conține uleiuri parfumate.', grape:'Strugurii cresc în ciorchini pe viță-de-vie.', blueberry:'Afinele își iau culoarea de la pigmenți numiți antociani.', peach:'Piersica are un sâmbure mare și coaja catifelată.' };
const FACTS_EN = { banana:'Botanically, a banana is a berry.', strawberry:'The tiny seeds on the outside are the actual fruits.', apple:'Apples belong to the rose family.', pineapple:'A pineapple is made of many small fruits fused together.' };

const NUTRIENTS = [['calories','Calorii','kcal'],['protein','Proteine','g'],['carbs','Carbohidrați','g'],['fat','Grăsimi','g'],['fiber','Fibre','g']];
// ID-uri USDA → cheie. Pentru calorii, datele "Foundation" folosesc uneori 2047/2048 în loc de 1008.
const NUTRIENT_IDS = { 1008:'calories', 2047:'calories', 2048:'calories', 1003:'protein', 1005:'carbs', 1004:'fat', 1079:'fiber' };
const USDA_NAMES = { apple:'apples raw with skin', orange:'oranges raw all commercial varieties', 'blood orange':'oranges raw all commercial varieties', grape:'grapes red or green raw', strawberry:'strawberries raw', blueberry:'blueberries raw', raspberry:'raspberries raw', blackberry:'blackberries raw', cherry:'cherries sweet raw', kiwi:'kiwifruit green raw', coconut:'coconut meat raw', lemon:'lemons raw without peel', lime:'limes raw', 'dragon fruit':'pitaya raw', 'passion fruit':'passion-fruit purple raw', watermelon:'watermelon raw', cantaloupe:'melons cantaloupe raw', honeydew:'melons honeydew raw', peach:'peaches yellow raw', plum:'plums raw', pear:'pears raw', mango:'mangos raw', pineapple:'pineapple raw all varieties', avocado:'avocados raw all commercial varieties', tangerine:'tangerines mandarin oranges raw', clementine:'clementines raw', date:'dates medjool', fig:'figs raw' };

// ---------- elemente din pagină ----------
const $ = s => document.querySelector(s);
const video = $('#camera-video'), preview = $('#preview-image'), placeholder = $('#camera-placeholder');
const upload = $('#image-upload'), cameraButton = $('#camera-button'), scanButton = $('#capture-button');
const cameraWindow = $('#camera-window'), cameraHint = $('#camera-hint');
const resultContent = $('#result-content'), resultFacts = $('#result-facts');
const arLabel = $('#ar-label'), arBox = $('#ar-box'), arRomanian = $('#ar-romanian'), arEnglish = $('#ar-english');
const fruitActions = $('#fruit-actions'), arAdventure = $('#ar-adventure');
const challengeId = new URLSearchParams(location.search).get('challenge');

let tf, models, modelPromise;
let cameraStream, scanning = false, scanBusy = false, scanTimer, previewUrl;
let lastFruitAt = 0, lastResultKey = '', currentFruitId = null;

const english = id => ENGLISH[id] || id.replace(/\b\w/g, c => c.toUpperCase());
const fruitEmoji = id => EMOJI[id] || '🍒';
const readJSON = (key, fallback) => { try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch { return fallback; } };
const writeJSON = (key, value) => { try { localStorage.setItem(key, JSON.stringify(value)); } catch {} };
const escapeHTML = s => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c]);

// ---------- pașaport + aventura AR (neschimbat) ----------
function getPassport() { const v = readJSON(PASSPORT_KEY, []); return Array.isArray(v) ? v : []; }
function collectFruit(id = currentFruitId) {
  if (!id || !FRUITS[id]) return;
  const found = getPassport();
  if (!found.includes(id)) { found.push(id); writeJSON(PASSPORT_KEY, found); }
  document.querySelectorAll('.passport-count').forEach(n => n.textContent = found.length);
  const b = $('#collect-fruit'); if (b) { b.textContent = '✅ Adăugat în jurnal'; b.classList.add('collected'); }
}
if (challengeId && FRUITS[challengeId]) {
  const banner = $('#scan-challenge');
  if (banner) { banner.hidden = false; banner.querySelector('strong').textContent = `Misiune: găsește ${FRUITS[challengeId]} (${english(challengeId)})!`; }
}
function activateAR() {
  if (!currentFruitId) return; const id = currentFruitId;
  $('#ar-world-fruit').textContent = fruitEmoji(id);
  $('#ar-adventure-name').textContent = `${FRUITS[id]} · ${english(id)}`;
  $('#ar-adventure-fact').textContent = FACTS[id] || `${FRUITS[id]} este o descoperire minunată. Privește forma și culoarea lui chiar în camera ta!`;
  $('#collect-fruit').textContent = getPassport().includes(id) ? '✅ Adăugat în jurnal' : 'Adaugă în jurnal';
  arAdventure.hidden = false; arLabel.hidden = true; cameraWindow.classList.add('in-adventure');
}
$('#ar-adventure-button')?.addEventListener('click', activateAR);
$('#ar-exit')?.addEventListener('click', () => { arAdventure.hidden = true; cameraWindow.classList.remove('in-adventure'); if (currentFruitId) arLabel.hidden = false; });
$('#collect-fruit')?.addEventListener('click', () => collectFruit());
const goQuiz = () => { if (currentFruitId) location.href = `joaca.html?fruit=${encodeURIComponent(currentFruitId)}`; };
$('#fruit-quiz-button')?.addEventListener('click', goQuiz);
$('#ar-quiz-shortcut')?.addEventListener('click', goQuiz);

// ---------- încărcarea modelelor (echivalent FruitAI.__init__) ----------
async function loadModels() {
  if (modelPromise) return modelPromise;
  modelPromise = (async () => {
    cameraHint.textContent = 'Se pregătește motorul AI în browser…';
    tf = tf || await import(TRANSFORMERS_URL);
    tf.env.allowLocalModels = false;
    tf.env.useBrowserCache = true;

    const progress_callback = info => {
      if (info?.status === 'progress' && info.total) {
        const file = String(info.file || '').split('/').pop();
        cameraHint.textContent = `Se descarcă ${info.name || ''} ${file}: ${Math.round(info.progress || 0)}% (doar prima dată)`;
      }
    };
    const build = async device => {
      const opts = { progress_callback, dtype: device === 'webgpu' ? 'fp16' : 'q8', ...(device ? { device } : {}) };
      // CLIP — text și imagine separat, ca să calculăm textul o singură dată
      const clipTokenizer = await tf.AutoTokenizer.from_pretrained(CLIP_NAME);
      const clipProcessor = await tf.AutoProcessor.from_pretrained(CLIP_NAME);
      const clipText = await tf.CLIPTextModelWithProjection.from_pretrained(CLIP_NAME, opts);
      const clipVision = await tf.CLIPVisionModelWithProjection.from_pretrained(CLIP_NAME, opts);
      cameraHint.textContent = 'CLIP pregătit · se încarcă detectorul OWL‑ViT (prima dată se descarcă)…';
      // OWL-ViT cuantizat (~150 MB) pe WASM — stabil pe orice browser
      const detTokenizer = await tf.AutoTokenizer.from_pretrained(DETECTOR_NAME);
      const detProcessor = await tf.AutoProcessor.from_pretrained(DETECTOR_NAME);
      const detector = await tf.OwlViTForObjectDetection.from_pretrained(DETECTOR_NAME, { progress_callback, dtype: 'q8' });

      cameraHint.textContent = 'Se pregătesc etichetele CLIP…';
      const textInputs = clipTokenizer(LABELS.map(l => `a photo of a ${l}`), { padding: true, truncation: true });
      const { text_embeds } = await clipText(textInputs);
      const textFeatures = normalizeRows(text_embeds.data, text_embeds.dims[0], text_embeds.dims[1]);
      const detTextInputs = detTokenizer(DET_QUERIES, { padding: true, truncation: true });
      return { clipProcessor, clipVision, textFeatures, dim: text_embeds.dims[1], detProcessor, detector, detTextInputs };
    };
    if ('gpu' in navigator) {
      try { models = await build('webgpu'); }
      catch (e) { console.warn('WebGPU indisponibil, folosesc WASM:', e); models = await build(); }
    } else models = await build();
  })().catch(e => { modelPromise = undefined; throw e; });
  return modelPromise;
}
function normalizeRows(data, rows, dim) {
  const out = new Float32Array(rows * dim);
  for (let r = 0; r < rows; r++) {
    let s = 0; for (let k = 0; k < dim; k++) s += data[r * dim + k] ** 2;
    s = Math.sqrt(s) || 1; for (let k = 0; k < dim; k++) out[r * dim + k] = data[r * dim + k] / s;
  }
  return out;
}

// ---------- imagini ----------
function sourceSize(src) {
  return [src.videoWidth || src.naturalWidth || src.width, src.videoHeight || src.naturalHeight || src.height];
}
function makeCanvas(source, maxSide = 640) {         // la fel ca în main.py: latura mare 640 px
  const [sw, sh] = sourceSize(source), scale = Math.min(1, maxSide / Math.max(sw, sh));
  const c = document.createElement('canvas');
  c.width = Math.max(1, Math.round(sw * scale)); c.height = Math.max(1, Math.round(sh * scale));
  c.getContext('2d', { willReadFrequently: true }).drawImage(source, 0, 0, c.width, c.height);
  return c;
}
function cropCanvas(canvas, { x1, y1, x2, y2 }) {
  const w = x2 - x1, h = y2 - y1, out = document.createElement('canvas');
  out.width = w; out.height = h;
  out.getContext('2d', { willReadFrequently: true }).drawImage(canvas, x1, y1, w, h, 0, 0, w, h);
  return out;
}
const toRaw = canvas => tf.RawImage.fromCanvas(canvas);

// ---------- detector: găsește fructul în cadru (find_fruit_box) ----------
async function findFruitBox(canvas) {
  const { detProcessor, detector, detTextInputs } = models;
  const { pixel_values } = await detProcessor(toRaw(canvas));
  const out = await detector({ ...detTextInputs, pixel_values });
  const [, P, Q] = out.logits.dims, logits = out.logits.data, boxes = out.pred_boxes.data;
  let best = 0, idx = -1;
  for (let p = 0; p < P; p++) {
    const fruit = sigmoid(logits[p * Q]);
    let others = 0; for (let q = 1; q < Q; q++) others = Math.max(others, sigmoid(logits[p * Q + q]));
    const score = fruit > others ? fruit : 0;        // exclude persoane / mâini / fețe
    if (score > best) { best = score; idx = p; }
  }
  if (idx < 0 || best < DETECT_THRESHOLD) return { box: null, score: best };
  const W = canvas.width, H = canvas.height;
  const [cx, cy, bw, bh] = boxes.slice(idx * 4, idx * 4 + 4);
  let x1 = Math.max(0, cx - bw / 2) * W, y1 = Math.max(0, cy - bh / 2) * H;
  let x2 = Math.min(1, cx + bw / 2) * W, y2 = Math.min(1, cy + bh / 2) * H;
  const mx = (x2 - x1) * 0.12, my = (y2 - y1) * 0.12;   // margine mică în jurul fructului
  x1 = Math.round(Math.max(0, x1 - mx)); y1 = Math.round(Math.max(0, y1 - my));
  x2 = Math.round(Math.min(W, x2 + mx)); y2 = Math.round(Math.min(H, y2 + my));
  if (x2 - x1 < 40 || y2 - y1 < 40) return { box: null, score: best };
  return { box: { x1, y1, x2, y2 }, score: best };
}
const sigmoid = v => 1 / (1 + Math.exp(-v));

// ---------- CLIP: ce este în imagine (classify) ----------
async function classify(canvas) {
  const { clipProcessor, clipVision, textFeatures, dim } = models;
  const { image_embeds } = await clipVision(await clipProcessor(toRaw(canvas)));
  const img = normalizeRows(image_embeds.data, 1, dim);
  const logits = new Float32Array(LABELS.length);
  let max = -Infinity;
  for (let i = 0; i < LABELS.length; i++) {
    let s = 0; for (let k = 0; k < dim; k++) s += img[k] * textFeatures[i * dim + k];
    logits[i] = CLIP_LOGIT_SCALE * s; if (logits[i] > max) max = logits[i];
  }
  let sum = 0, bestI = 0;
  for (let i = 0; i < logits.length; i++) { logits[i] = Math.exp(logits[i] - max); sum += logits[i]; if (logits[i] > logits[bestI]) bestI = i; }
  return { label: LABELS[bestI], score: logits[bestI] / sum };
}

// ---------- logica principală (worker din main.py) ----------
async function analyzeSource(source) {
  const canvas = makeCanvas(source);
  const { box } = await findFruitBox(canvas);
  if (box) {
    const { label, score } = await classify(cropCanvas(canvas, box));
    if (FRUITS[label] && score >= FRUIT_MIN_SCORE) return { kind: 'fruit', id: label, score, box, localized: true, canvas };
    if (NON_FRUITS[label]) return { kind: 'notfruit', label, score, box, canvas, text: `Nu este un fruct: ${NON_FRUITS[label]}` };
    return { kind: 'none', canvas, text: 'Nu sunt sigur ce este' };
  }
  // Niciun fruct găsit: verificăm tot cadrul
  const { label, score } = await classify(canvas);
  const W = canvas.width, H = canvas.height;
  if (FRUITS[label] && score >= FULLFRAME_MIN_SCORE) return { kind: 'fruit', id: label, score, box: { x1: 0, y1: 0, x2: W, y2: H }, localized: false, canvas };
  if (NON_FRUITS[label] && score >= OTHER_MIN_SCORE) return { kind: 'notfruit', label, score, box: null, canvas, text: `Nu văd niciun fruct – văd ${NON_FRUITS[label]}` };
  return { kind: 'none', canvas, text: 'Nu văd niciun fruct' };
}

// ---------- afișare (draw_interface) ----------
function showRecognition(result, source) {
  if (result.kind === 'fruit') {
    lastFruitAt = Date.now(); cameraWindow.classList.add('has-result');
    placeAR(scaleBox(result.box, result.canvas, source), 'fruit', FRUITS[result.id], english(result.id), result.localized);
    renderFruit(result.id, result.score, result.localized);
    return;
  }
  // un fruct găsit recent rămâne pe ecran RESULT_TTL ms (ca în Python)
  if (Date.now() - lastFruitAt <= RESULT_TTL) return;
  currentFruitId = null; fruitActions.hidden = true; cameraWindow.classList.remove('has-result');
  if (result.kind === 'notfruit') {
    if (result.box) placeAR(scaleBox(result.box, result.canvas, source), 'notfruit', NON_FRUITS[result.label], english(result.label), true);
    else clearAR();
    setResult('🚫', 'Acesta nu este un fruct', `Văd ${NON_FRUITS[result.label]} (${Math.round(result.score * 100)}%). Arată un fruct camerei.`, `notfruit-${result.label}`);
  } else {
    clearAR();
    setResult('✳', scanning ? 'Caut un fruct…' : result.text, 'Arată un fruct oriunde în cadru. O lumină bună ajută mult.', 'searching');
  }
  cameraHint.textContent = result.text;
}
function scaleBox(b, canvas, source) {
  const [nw, nh] = sourceSize(source), sx = nw / canvas.width, sy = nh / canvas.height;
  return { x: b.x1 * sx, y: b.y1 * sy, width: (b.x2 - b.x1) * sx, height: (b.y2 - b.y1) * sy };
}
function setResult(emoji, title, description, key = '') {
  if (key && key === lastResultKey) return; lastResultKey = key;
  resultContent.innerHTML = `<span class="result-emoji" aria-hidden="true">${emoji}</span><h3>${escapeHTML(title)}</h3><p>${escapeHTML(description)}</p>`;
  resultFacts.hidden = true; resultFacts.replaceChildren();
}
function renderFruit(id, score, localized) {
  currentFruitId = id; try { localStorage.setItem(LAST_FRUIT_KEY, id); } catch {}
  fruitActions.hidden = false;
  const ro = FRUITS[id], en = english(id), percent = Math.round(score * 100), key = `fruit-${id}`;
  if (key === lastResultKey) {
    const c = resultContent.querySelector('.confidence strong'), bar = resultContent.querySelector('.confidence b');
    if (c) c.textContent = `${percent}%`; if (bar) bar.style.width = `${percent}%`;
    return;
  }
  if (!arAdventure.hidden) { arAdventure.hidden = true; cameraWindow.classList.remove('in-adventure'); }
  lastResultKey = key;
  resultContent.innerHTML = `<span class="result-emoji" aria-hidden="true">${fruitEmoji(id)}</span><p class="result-overline">TE-AM GĂSIT · FRUIT FOUND</p><h3>${ro}</h3><p class="english-name"><span>ENGLISH</span><strong>${en}</strong></p><div class="confidence"><span>Scor de clasificare CLIP</span><strong>${percent}%</strong><i><b style="width:${percent}%"></b></i></div>`;
  resultFacts.innerHTML = `<div class="fact-item"><span>🌿</span><p><strong>Știai că? · Română</strong>${FACTS[id] || `${ro} este un fruct pe care îl poți explora și compara cu alte specii.`}</p></div><div class="fact-item"><span>🌎</span><p><strong>Did you know? · English</strong>${FACTS_EN[id] || `${en} is a fruit you can explore and compare with other species.`}</p></div><div class="nutrition-card"><div class="nutrition-heading"><span>🧪</span><div><strong>Nutriție estimată / 100 g</strong><small>Se caută valorile nutriționale…</small></div></div></div><div class="live-pulse"><span></span> OWL‑VIT + CLIP · ${localized ? 'OBIECT LOCALIZAT' : 'RECUNOAȘTERE PE TOT CADRUL'}</div>`;
  resultFacts.hidden = false;
  cameraHint.textContent = `Identificat: ${ro} (${percent}%)${scanning ? ' · camera continuă să caute automat.' : ''}`;
  const msg = $('#scan-challenge-message');
  if (msg && challengeId && FRUITS[challengeId]) { msg.textContent = id === challengeId ? 'Misiune reușită! Ai găsit fructul cerut! ⭐' : `Ai găsit ${ro}. Mai caută ${FRUITS[challengeId]} pentru misiune!`; msg.hidden = false; }
  nutritionFor(id)
    .then(data => { if (lastResultKey === key) renderNutrition(data); })
    .catch(err => { console.warn('Eroare USDA:', err); if (lastResultKey === key) renderNutrition(null); });
}
function renderNutrition(data) {
  const node = resultFacts.querySelector('.nutrition-card'); if (!node) return;
  if (!data) { node.querySelector('small').textContent = 'Fruct identificat; nutriție indisponibilă momentan'; return; }
  const rows = NUTRIENTS.map(([k, label, unit]) => { const v = data.values[k]; return `<li><span>${label}</span><strong>${v == null ? 'indisponibil' : `${Number(Number(v).toFixed(2))} ${unit}`}</strong></li>`; }).join('');
  node.innerHTML = `<div class="nutrition-heading"><span>🧪</span><div><strong>Nutriție estimată / 100 g</strong><small>USDA · ${escapeHTML(data.food)}</small></div></div><ul>${rows}</ul>`;
}
function clearAR() { arLabel.hidden = true; arBox.hidden = true; arBox.classList.remove('is-notfruit'); arLabel.classList.remove('is-notfruit'); }
function placeAR(box, kind, title, subtitle, localized) {
  const [sw, sh] = sourceSize(video.hidden ? preview : video);
  const rect = cameraWindow.getBoundingClientRect(), isVideo = !video.hidden;
  // video = object-fit: cover, poză = object-fit: contain
  const fit = isVideo ? Math.max(rect.width / sw, rect.height / sh) : Math.min(rect.width / sw, rect.height / sh);
  const ox = (rect.width - sw * fit) / 2, oy = (rect.height - sh * fit) / 2;
  const left = ox + box.x * fit, top = oy + box.y * fit, w = box.width * fit, h = box.height * fit;
  Object.assign(arBox.style, { left: `${left}px`, top: `${top}px`, width: `${w}px`, height: `${h}px` });
  arBox.hidden = false;
  const notFruit = kind === 'notfruit';
  arBox.classList.toggle('is-notfruit', notFruit); arLabel.classList.toggle('is-notfruit', notFruit);
  const world = $('#ar-world'); world.style.left = `${left + w / 2}px`; world.style.top = `${top + h / 2}px`;
  arLabel.querySelector('.ar-label-kicker').textContent = notFruit ? 'ACESTA NU ESTE UN FRUCT' : 'FRUCT IDENTIFICAT';
  arRomanian.textContent = title; arEnglish.textContent = subtitle;
  arLabel.style.left = `${Math.min(rect.width - 105, Math.max(105, left + w / 2))}px`;
  arLabel.style.top = `${Math.max(95, top)}px`;
  arLabel.classList.toggle('ar-label-wide', !localized);
  arLabel.hidden = !arAdventure.hidden;
}

// ---------- nutriție (nutrition_for) ----------
let seedCache;
async function loadSeedCache() {
  if (seedCache) return seedCache;
  try { const r = await fetch('fruitai_cache.json', { cache: 'no-cache' }); seedCache = r.ok ? await r.json() : {}; }
  catch { seedCache = {}; }
  return seedCache;
}
function normalizeEntry(entry) {
  if (!entry || typeof entry !== 'object' || !entry.values) return null;
  const values = {};
  for (const [k, v] of Object.entries(entry.values)) {
    const key = NUTRIENT_IDS[k] || k;
    if (NUTRIENTS.some(([n]) => n === key) && v != null && values[key] == null) values[key] = Number(v);
  }
  if (values.calories == null) return null;            // intrare incompletă → o cerem din nou
  return { food: entry.food || 'USDA', values, basis: '100 g' };
}
function foodScore(food, fruit) {
  const d = (food.description || '').toLowerCase();
  let s = 0;
  if (d.includes('raw')) s += 5;
  if (/juice|dried|canned|frozen|syrup|sweetened|cooked|leaves|peel|seed|nectar|baby|pie|jam|candied/.test(d)) s -= 6;
  if (d.replaceAll(' ', '').includes(fruit.replaceAll(' ', ''))) s += 3;
  if (food.dataType === 'SR Legacy') s += 1;             // SR Legacy are valori complete (calorii, fibre)
  return s;
}
async function nutritionFor(fruit) {
  const cache = readJSON(NUTRITION_CACHE_KEY, {});
  const local = normalizeEntry(cache[fruit]); if (local) return local;
  const seed = normalizeEntry((await loadSeedCache())[fruit]); if (seed) return seed;

  const query = USDA_NAMES[fruit] || `${fruit} raw`;
  const url = `${NUTRITION_URL}?api_key=${USDA_KEY}&query=${encodeURIComponent(query)}&pageSize=20&dataType=Foundation,SR%20Legacy`;
  const response = await fetch(url);
  if (!response.ok) throw new Error(`USDA ${response.status}`);
  const foods = (await response.json()).foods || [];
  if (!foods.length) throw new Error('USDA nu a găsit alimentul.');
  const food = foods.reduce((a, b) => foodScore(b, fruit) > foodScore(a, fruit) ? b : a);
  const values = {};
  for (const n of food.foodNutrients || []) {
    const key = NUTRIENT_IDS[n.nutrientId];
    if (key && n.value != null && (values[key] == null || n.nutrientId === 1008)) values[key] = n.value;
  }
  const result = { food: food.description || query, values, basis: '100 g', source: 'USDA FoodData Central' };
  cache[fruit] = result; writeJSON(NUTRITION_CACHE_KEY, cache);
  return result;
}

// ---------- cameră ----------
function stopCamera(clear = true) {
  scanning = false; scanBusy = false; clearTimeout(scanTimer);
  cameraStream?.getTracks().forEach(t => t.stop()); cameraStream = undefined;
  video.srcObject = null; video.hidden = true; scanButton.hidden = true; scanButton.disabled = false;
  cameraButton.innerHTML = '<span aria-hidden="true">◎</span> Deschide camera';
  cameraWindow.classList.remove('is-scanning', 'has-result'); clearAR();
  arAdventure.hidden = true; cameraWindow.classList.remove('in-adventure');
  if (clear) { preview.hidden = true; placeholder.hidden = false; }
}
cameraButton.addEventListener('click', async () => {
  if (cameraStream) { stopCamera(); cameraHint.textContent = 'Scanarea s-a oprit. Imaginea este procesată numai în browser.'; return; }
  if (!window.isSecureContext || !navigator.mediaDevices?.getUserMedia) {
    cameraHint.textContent = 'Camera merge doar prin http://localhost sau HTTPS, nu deschizând fișierul direct. Vezi README.';
    return;
  }
  cameraButton.disabled = true; lastResultKey = ''; lastFruitAt = 0; cameraHint.textContent = 'Se pornește camera…';
  try {
    cameraStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' }, width: { ideal: 1280 }, height: { ideal: 720 } }, audio: false });
    video.srcObject = cameraStream; video.hidden = false; preview.hidden = true; placeholder.hidden = true; clearAR();
    await new Promise((resolve, reject) => {
      if (video.readyState >= 2) return resolve();
      const t = setTimeout(() => reject(new Error('Camera nu a trimis imagine. Verifică permisiunile și încearcă din nou.')), 12000);
      video.addEventListener('loadeddata', () => { clearTimeout(t); resolve(); }, { once: true });
    });
    await video.play();
    cameraWindow.classList.add('is-scanning'); scanButton.hidden = false; scanButton.disabled = true;
    scanButton.textContent = 'Încarc AI…'; cameraButton.innerHTML = '<span>…</span> Încarc modelele AI';
    await loadModels();
    if (!cameraStream) return;                           // utilizatorul a oprit camera între timp
    scanning = true; cameraButton.disabled = false; cameraButton.innerHTML = '<span>×</span> Oprește camera';
    scanButton.textContent = 'Scanează din nou'; scanButton.disabled = false;
    cameraHint.textContent = 'Modele pregătite · arată un fruct oriunde în cadru.';
    setResult('✳', 'Caut un fruct…', 'Arată un fruct oriunde în cadru.', 'searching');
    scanLoop();
  } catch (error) {
    stopCamera(true); cameraButton.disabled = false;
    cameraHint.textContent = error.name === 'NotAllowedError' ? 'Accesul la cameră a fost refuzat. Permite accesul pentru acest site și încearcă din nou.'
      : error.name === 'NotFoundError' ? 'Nu am găsit o cameră conectată. Poți încărca o fotografie.'
      : `Nu am putut porni: ${error.message || error}`;
    console.error(error);
  }
});
scanButton.addEventListener('click', () => {
  if (cameraStream && !scanning) { scanning = true; scanButton.textContent = 'Scanează din nou'; cameraHint.textContent = 'Analiza automată a fost reluată.'; }
  if (scanning) { clearTimeout(scanTimer); scanLoop(); }
});
async function scanLoop() {
  if (!scanning || scanBusy) return;
  if (video.readyState < 2) { scanTimer = setTimeout(scanLoop, CHECK_INTERVAL); return; }
  scanBusy = true;
  const started = performance.now();
  try { const result = await analyzeSource(video); if (scanning) showRecognition(result, video); }
  catch (error) {
    scanning = false; scanButton.textContent = 'Reîncearcă analiza'; scanButton.disabled = false;
    cameraHint.textContent = `Eroare AI: ${String(error.message || error).slice(0, 100)}`;
    setResult('AI', 'Analiza s-a oprit', 'Apasă „Reîncearcă analiza”. Camera rămâne pornită.', 'inference-error');
    console.error(error);
  } finally {
    scanBusy = false;
    if (scanning) scanTimer = setTimeout(scanLoop, Math.max(0, CHECK_INTERVAL - (performance.now() - started)));
  }
}

// ---------- poză încărcată ----------
upload.addEventListener('change', async () => {
  const file = upload.files?.[0]; if (!file) return;
  stopCamera(false); lastResultKey = ''; lastFruitAt = 0; clearAR();
  if (previewUrl) URL.revokeObjectURL(previewUrl);
  previewUrl = URL.createObjectURL(file); preview.src = previewUrl; preview.hidden = false; placeholder.hidden = true; upload.value = '';
  try {
    await preview.decode();
    setResult('⏳', 'Analizez poza…', 'OWL‑ViT caută fructul, apoi CLIP îl identifică.', 'loading');
    await loadModels();
    cameraHint.textContent = 'OWL‑ViT caută fructul, apoi CLIP îl identifică…';
    const result = await analyzeSource(preview);
    lastResultKey = '';
    showRecognition(result, preview);
    if (result.kind === 'fruit') cameraHint.textContent = `Identificat: ${FRUITS[result.id]} (${Math.round(result.score * 100)}%). Încarcă altă poză sau pornește camera.`;
  } catch (error) {
    setResult('🌱', 'Analiza nu a pornit', error.message || String(error), 'error');
    cameraHint.textContent = error.message || String(error); console.error(error);
  }
});

cameraHint.textContent = 'Deschide camera sau încarcă o poză: OWL‑ViT caută fructul, iar CLIP îl identifică în română și engleză.';
