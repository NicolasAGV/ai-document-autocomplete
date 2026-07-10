import subprocess, sys
for pkg in ['opencv-python-headless', 'Pillow', 'numpy', 'pandas', 'anthropic']:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', pkg, '-q'])

import cv2
import numpy as np
import pandas as pd
import anthropic
import base64, json, os, re, time
from pathlib import Path
from PIL import Image
import sys

sys.path.insert(0, str(Path(__file__).parent.parent / '00'))
import main_path as _mp  # type: ignore

input_folder = Path(_mp.main_path)

# ── PATHS ──────────────────────────────────────────────────────────────────────
# Algun_path = Path(Master_path.main_path) / 'output'
IMAGE_FOLDER  = input_folder / 'output' / 'fotos_renamed'
OUTPUT_JSON   = input_folder / 'output' / 'json' / "extracted_api_ocr_marking.json"

# ── Maximum allowed estimated cost (USD) — script aborts if exceeded ───────────
COST_THRESHOLD = 0.03

# ── Image processing settings ──────────────────────────────────────────────────
MAX_LONG_SIDE  = 600
JPEG_QUALITY   = 82

# ── API settings ───────────────────────────────────────────────────────────────
MODEL               = 'claude-sonnet-4-5'
MAX_TOKENS          = 64
DELAY_BETWEEN_CALLS = 1.0   # seconds between requests (increase to 12 on free tier)

# ── Pricing (per million tokens) ───────────────────────────────────────────────
INPUT_COST_PER_MTK  = 3.00
OUTPUT_COST_PER_MTK = 15.00

SUPPORTED = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'}

# ──────────────────────────────────────────────────────────────────────────────
# 1. Validate paths
# ──────────────────────────────────────────────────────────────────────────────
folder_in = Path(IMAGE_FOLDER)
if not folder_in.exists():
    raise FileNotFoundError(f'Input folder not found: {IMAGE_FOLDER}')

# ──────────────────────────────────────────────────────────────────────────────
# 2. Filter images
# ──────────────────────────────────────────────────────────────────────────────
all_files = sorted(f for f in folder_in.iterdir() if f.is_file() and f.suffix.lower() in SUPPORTED)

image_files = sorted(
    f for f in all_files
    if 'marcado'        in f.stem.lower()
    and 'producto'      not in f.stem.lower()
    and 'certificadora' not in f.stem.lower()
)

print(f'📂 Folder : {folder_in}')
print(f'   {len(all_files)} total images, {len(image_files)} pass filter (marcado / no producto / no certificadora)')
for f in all_files:
    stem = f.stem.lower()
    passes = ('marcado' in stem) and ('producto' not in stem) and ('certificadora' not in stem)
    print(f'   {"✅" if passes else "❌"}  {f.name}')

# ──────────────────────────────────────────────────────────────────────────────
# 3. Preprocess (resize + JPEG encode) — saved back into IMAGE_FOLDER
# ──────────────────────────────────────────────────────────────────────────────
def resize_for_api(img_bgr, max_long_side=MAX_LONG_SIDE):
    h, w = img_bgr.shape[:2]
    scale = max_long_side / max(h, w)
    if abs(scale - 1.0) < 0.01:
        return img_bgr
    interp = cv2.INTER_CUBIC if scale > 1 else cv2.INTER_AREA
    return cv2.resize(img_bgr, (int(w * scale), int(h * scale)), interpolation=interp)

def preprocess_image(src_path, dst_path, max_long_side=MAX_LONG_SIDE, quality=JPEG_QUALITY):
    img = cv2.imread(str(src_path))
    if img is None:
        raise ValueError(f'Cannot read: {src_path}')
    orig_bytes = src_path.stat().st_size
    resized    = resize_for_api(img, max_long_side)
    dst_path   = dst_path.with_suffix('.jpg')
    cv2.imwrite(str(dst_path), resized, [cv2.IMWRITE_JPEG_QUALITY, quality])
    final_bytes = dst_path.stat().st_size
    h_orig, w_orig = img.shape[:2]
    h_fin,  w_fin  = resized.shape[:2]
    return {
        'orig_size' : f'{w_orig}x{h_orig}px ({orig_bytes/1024:.0f} KB)',
        'final_size': f'{w_fin}x{h_fin}px ({final_bytes/1024:.0f} KB)',
        'reduction' : f'{(1 - final_bytes/orig_bytes)*100:.0f}%',
        'dst_path'  : dst_path,
    }

print(f'\nPreprocessing {len(image_files)} image(s)...')
for src in image_files:
    try:
        info = preprocess_image(src, folder_in / src.name)
        print(f'  ✅ {src.name:<45} {info["orig_size"]} → {info["final_size"]}  saved {info["reduction"]}')
    except Exception as e:
        print(f'  ❌ {src.name}: {e}')

# ──────────────────────────────────────────────────────────────────────────────
# 4. Prompts (inline)
# ──────────────────────────────────────────────────────────────────────────────
_BASE = (
    'Extract all text, numbers, symbols, and characters visible on this component exactly as they appear. '
    'The marking may include standard characters as well as less common ones such as "µ" (micro), '
    '"~" (AC/alternating current), "²" (squared), "º" (degree), "Ω" (ohm), "±" (tolerance). '
    'Capture every character precisely — do not confuse 0 (zero) with O (letter), 1 (one) with I (letter), '
    'or µ (micro) with u. If a character is uncertain, write it followed by a question mark in brackets, '
    'e.g. "B[?]". The text may be rotated, embossed, or printed in very small size — read it regardless. '
    'Return only the raw extracted text, nothing else.'
)

_CAP_EXTRA = (
    ' IMPORTANT — post-processing rules: '
    '(1) Voltage and capacitance are two separate values: if you read them joined without a space '
    '(e.g. "400V15µF" or "400v15µf"), always insert a space between them (e.g. "400V 15µF"). '
    '(2) Replace any "." used as a decimal separator between digits with "," '
    '(e.g. "4.7µF" → "4,7µF", "63.0V" → "63,0V"). Return only the raw extracted text, nothing else.'
)

_FUSIBLE_PROMPT = (
    'Extract all text, numbers, and symbols visible on this component exactly as they appear. '
    'Only output characters that can be represented as a single Unicode character — letters, digits, '
    'and symbols such as "µ" (micro), "~" (AC), "²" (squared), "º" (degree), "Ω" (ohm), "±" (tolerance), '
    '"©", "®", "™", "⚠". Do NOT describe or name any graphical logo or certification mark '
    '(such as CE logo, UL logo, Kitemark, CCC, VDE, NF, N-mark, or similar): if it has no single Unicode '
    'character, skip it entirely — no words, no brackets, no descriptions. '
    'Capture every typeable character precisely — do not confuse 0 (zero) with O (letter), '
    '1 (one) with I (letter), or µ (micro) with u. If a character is uncertain, write it followed by a '
    'question mark in brackets, e.g. "B[?]". The text may be rotated, embossed, or printed in very small '
    'size — read it regardless. Return only the raw extracted text, nothing else.'
)

_DEFAULT_PROMPT = _BASE
_COMPONENT_PROMPTS: dict = {
    'Capacitor_Electrolitico': _BASE + _CAP_EXTRA,
    'Capacitor_Y':             _BASE + _CAP_EXTRA,
    'Capacitor_X':             _BASE + _CAP_EXTRA,
    'Fusible':                 _FUSIBLE_PROMPT,
    'Varistor':                _BASE,
    'Resistencia':             _BASE,
    'Bobina':                  _BASE,
    'Transformador':           _BASE,
    'Optoacoplador':           _BASE,
    'Interruptor':             _BASE,
    'Parlante':                _BASE,
    'Cordon_de_alimentacion':  _BASE,
    'Ficha_de_alimentacion':   _BASE,
    'PCB':                     _BASE,
    'Bornera':                 _BASE,
}

def _get_prompt(stem: str) -> str:
    """Return the component-specific prompt for this filename stem, or the default."""
    stem_lower = stem.lower()
    match = max(
        (k for k in _COMPONENT_PROMPTS if stem_lower.startswith(k.lower())),
        key=len,
        default=None,
    )
    return _COMPONENT_PROMPTS[match] if match else _DEFAULT_PROMPT

print(f'\n📄 Prompts loaded inline  ({len(_COMPONENT_PROMPTS)} component(s) + default)')

# ──────────────────────────────────────────────────────────────────────────────
# 5. Estimate cost — abort if over threshold
# ──────────────────────────────────────────────────────────────────────────────
def estimate_image_tokens(path):
    img = Image.open(path)
    w, h = img.size
    return max(100, int(w * h / 750))

preprocessed_files = [
    folder_in / (f.stem + '.jpg')
    for f in image_files
    if (folder_in / (f.stem + '.jpg')).exists()
]

OUTPUT_TOKENS_EST   = 30
total_input_tokens  = sum(
    estimate_image_tokens(f) + int(len(_get_prompt(f.stem)) / 4)
    for f in preprocessed_files
)
total_output_tokens = len(preprocessed_files) * OUTPUT_TOKENS_EST
estimated_cost      = (total_input_tokens * INPUT_COST_PER_MTK +
                       total_output_tokens * OUTPUT_COST_PER_MTK) / 1_000_000

print(f'\n💰 Estimated cost : ${estimated_cost:.4f}')
print(f'   Threshold      : ${COST_THRESHOLD:.4f}')
if estimated_cost > COST_THRESHOLD:
    raise SystemExit(
        f'\n❌ Estimated cost ${estimated_cost:.4f} exceeds threshold ${COST_THRESHOLD:.4f}.\n'
        '   Increase COST_THRESHOLD or reduce the number of images.'
    )
print('✅ Cost within threshold — proceeding.')

# ──────────────────────────────────────────────────────────────────────────────
# 6. Connect to Anthropic API
# ──────────────────────────────────────────────────────────────────────────────
api_key = os.environ.get('ANTHROPIC_API_KEY', '').strip()
if not api_key:
    # Fallback: read the key from the local key file if the env var isn't set.
    # Project root = two levels up from this script (…/Main/01 Extractors/ -> project root).
    _project_root = Path(__file__).resolve().parent.parent.parent
    _key_file = _project_root / '.XLSX .DOCX patrones' / 'Super secret API key maybe anthropic.txt'
    if _key_file.is_file():
        api_key = _key_file.read_text(encoding='utf-8').strip()
if not api_key:
    raise EnvironmentError(
        'ANTHROPIC_API_KEY not found.\n'
        'Set it with:  setx ANTHROPIC_API_KEY "sk-ant-your-key-here"\n'
        'Then restart your terminal.\n'
        f'(Or place the key in: {_key_file})'
    )
client = anthropic.Anthropic(api_key=api_key)
try:
    client.models.list()
    print(f'✅ API connected  (key: ...{api_key[-6:]})')
except anthropic.AuthenticationError:
    raise ValueError('❌ Invalid API key. Check it at console.anthropic.com')
except Exception as e:
    raise RuntimeError(f'❌ Connection error: {e}')

# ──────────────────────────────────────────────────────────────────────────────
# 7. Batch recognition
# ──────────────────────────────────────────────────────────────────────────────
def _postprocess_capacitor(text: str) -> str:
    # Insert space between voltage and capacitance when joined: "400v15µF" → "400v 15µF"
    text = re.sub(r'(\d+\s*[vV])(\d)', r'\1 \2', text)
    # Replace decimal dot between digits with comma: "4.7µF" → "4,7µF"
    text = re.sub(r'(\d)\.(\d)', r'\1,\2', text)
    return text

_CAPACITOR_PREFIXES = ('capacitor',)

def recognize_image(image_path, prompt, retries=1):
    with open(image_path, 'rb') as f:
        image_data = base64.standard_b64encode(f.read()).decode()
    for attempt in range(retries + 1):
        try:
            response = client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                messages=[{
                    'role': 'user',
                    'content': [
                        {
                            'type'  : 'image',
                            'source': {'type': 'base64', 'media_type': 'image/jpeg', 'data': image_data},
                        },
                        {'type': 'text', 'text': prompt},
                    ]
                }]
            )
            result = next(
                (b.text.strip() for b in response.content if b.type == 'text'),
                ''
            )
            usage  = response.usage
            return result, usage.input_tokens, usage.output_tokens, 'ok'
        except anthropic.RateLimitError:
            if attempt < retries:
                print('     ⚠️  Rate limit — waiting 60 s...')
                time.sleep(60)
            else:
                return '', 0, 0, 'rate_limit'
        except Exception as e:
            if attempt < retries:
                time.sleep(3)
            else:
                return '', 0, 0, str(e)

if not preprocessed_files:
    raise FileNotFoundError('No preprocessed images found — check IMAGE_FOLDER.')

print(f'\n🚀 Sending {len(preprocessed_files)} image(s) to {MODEL}...\n')
rows = []
total_in, total_out = 0, 0

for i, img_path in enumerate(preprocessed_files, 1):
    result, in_tok, out_tok, status = recognize_image(img_path, _get_prompt(img_path.stem))
    if img_path.stem.lower().startswith(_CAPACITOR_PREFIXES):
        result = _postprocess_capacitor(result)
    total_in  += in_tok
    total_out += out_tok
    rows.append({
        'file_name'    : img_path.name,
        'recognized'   : result,
        'input_tokens' : in_tok,
        'output_tokens': out_tok,
        'status'       : status,
    })
    print(f'  [{i}/{len(preprocessed_files)}] {img_path.name:<45} {result[:50]}  ({status})')
    if i < len(preprocessed_files):
        time.sleep(DELAY_BETWEEN_CALLS)

df = pd.DataFrame(rows)
actual_cost = (total_in * INPUT_COST_PER_MTK + total_out * OUTPUT_COST_PER_MTK) / 1_000_000
print(f'\n✅ Done — {len(df)} image(s) processed  |  actual cost: ${actual_cost:.5f}')

# ──────────────────────────────────────────────────────────────────────────────
# 8. Export results to JSON
# ──────────────────────────────────────────────────────────────────────────────
results_dict = {
    row['file_name']: row['recognized']
    for _, row in df.iterrows()
    if row['status'] == 'ok'
}

json_path = Path(OUTPUT_JSON)
json_path.parent.mkdir(parents=True, exist_ok=True)
json_path.write_text(json.dumps(results_dict, ensure_ascii=False, indent=2), encoding='utf-8')

print(f'\n✅ Results saved to: {json_path}')
print(json.dumps(results_dict, ensure_ascii=False, indent=2))
