"""
esrgan_helper.py  —  AI image upscaling (local Real-ESRGAN or Replicate API)
Importable: enhance_pil(img, scale=4) -> PIL Image
CLI:        python esrgan_helper.py <input_path> <output_path> [scale]

Backends (set UPSCALE_BACKEND below):
  "local"     — local Real-ESRGAN Python package (free, needs GPU for speed)
  "replicate" — prunaai/p-image-upscale via Replicate API (fast, ~$0.01/image)
  "auto"      — try local first, fall back to replicate if not installed

Setup:
  Local:     pip install realesrgan basicsr torch torchvision
  Replicate: pip install replicate
             set env var REPLICATE_API_TOKEN=your_token
"""

import os
import sys
import io
import numpy as np
from PIL import Image

# ── Config ────────────────────────────────────────────────────────────────────
UPSCALE_BACKEND = "auto"   # "local" | "replicate" | "auto"

# Local Real-ESRGAN settings
LOCAL_MODEL_DIR  = r'C:\Varsany\realesrgan'
LOCAL_MODEL_FILE = 'RealESRGAN_x4plus.pth'
LOCAL_MODEL_URL  = 'https://github.com/xinntao/Real-ESRGAN/releases/download/v0.1.0/RealESRGAN_x4plus.pth'
LOCAL_TILE_SIZE  = 200   # tune for your GPU VRAM; 128=safe/slow, 200=balanced, 400=OOM on 2GB

# Replicate settings
REPLICATE_MODEL  = "prunaai/p-image-upscale:7135ff723ecea89c0f67afcd51e4904904586e351093465bdc7beed45941b3e0"
REPLICATE_TOKEN  = os.environ.get("REPLICATE_API_TOKEN", "")
REPLICATE_TARGET = 8              # megapixels (max 8)
REPLICATE_FORMAT = "png"          # png keeps full quality (no compression)
# ─────────────────────────────────────────────────────────────────────────────

_local_upsampler = None   # singleton — loaded once per process


# ── Local backend ─────────────────────────────────────────────────────────────

def _get_local_upsampler():
    global _local_upsampler
    if _local_upsampler is not None:
        return _local_upsampler

    try:
        import torch
        from realesrgan import RealESRGANer
        from basicsr.archs.rrdbnet_arch import RRDBNet
    except ImportError as e:
        raise ImportError(
            f"realesrgan not installed ({e}). "
            "Run: pip install realesrgan basicsr torch torchvision"
        ) from e

    os.makedirs(LOCAL_MODEL_DIR, exist_ok=True)
    model_path = os.path.join(LOCAL_MODEL_DIR, LOCAL_MODEL_FILE)
    if not os.path.isfile(model_path):
        model_path = LOCAL_MODEL_URL
        print("[ESRGAN] Model not found locally — will download from GitHub", flush=True)

    gpu_id = 0 if torch.cuda.is_available() else None
    # fp16 (half precision) can hang on low-VRAM cards (e.g. MX550 2GB) — use fp32
    half   = False

    model = RRDBNet(
        num_in_ch=3, num_out_ch=3,
        num_feat=64, num_block=23, num_grow_ch=32,
        scale=4
    )
    _local_upsampler = RealESRGANer(
        scale=4,
        model_path=model_path,
        model=model,
        tile=LOCAL_TILE_SIZE,
        tile_pad=10,
        pre_pad=0,
        half=half,
        gpu_id=gpu_id,
    )
    device = "GPU" if gpu_id is not None else "CPU"
    print(f"[ESRGAN] Local model loaded ({device})", flush=True)
    return _local_upsampler


def _enhance_local(img_pil: Image.Image, scale: int = 4) -> Image.Image:
    has_alpha = img_pil.mode == 'RGBA'
    if has_alpha:
        r, g, b, a = img_pil.split()
        rgb = Image.merge('RGB', (r, g, b))
        a_up = a.resize((max(1, a.width * scale), max(1, a.height * scale)), Image.LANCZOS)
    else:
        rgb = img_pil.convert('RGB')
        a_up = None

    img_bgr = np.array(rgb)[:, :, ::-1]
    out_bgr, _ = _get_local_upsampler().enhance(img_bgr, outscale=scale)
    out_pil = Image.fromarray(out_bgr[:, :, ::-1], 'RGB')

    if a_up is not None:
        if a_up.size != out_pil.size:
            a_up = a_up.resize(out_pil.size, Image.LANCZOS)
        out_pil = out_pil.convert('RGBA')
        out_pil.putalpha(a_up)

    return out_pil


# ── Replicate backend ─────────────────────────────────────────────────────────

def _enhance_replicate(img_pil: Image.Image) -> Image.Image:
    try:
        import replicate
    except ImportError as e:
        raise ImportError(
            "replicate not installed. Run: pip install replicate"
        ) from e

    if not REPLICATE_TOKEN:
        raise ValueError(
            "REPLICATE_API_TOKEN not set. "
            "Add it to your .env file or set the environment variable."
        )

    # Convert PIL image → PNG bytes → base64 data URI
    buf = io.BytesIO()
    img_pil.convert('RGB').save(buf, format='PNG')
    buf.seek(0)

    print("[ESRGAN] Sending to Replicate (prunaai/p-image-upscale)...", flush=True)

    output = replicate.run(
        REPLICATE_MODEL,
        input={
            "image":            buf,
            "upscale_mode":     "target",
            "target":           REPLICATE_TARGET,
            "enhance_details":  True,
            "enhance_realism":  True,
            "output_format":    REPLICATE_FORMAT,
            "output_quality":   100,
        }
    )

    # output is a FileOutput object — read its bytes
    img_bytes = output.read() if hasattr(output, 'read') else output
    out_pil = Image.open(io.BytesIO(img_bytes)).convert('RGBA')

    # Re-attach original alpha if input had transparency
    if img_pil.mode == 'RGBA':
        a = img_pil.split()[3]
        a_up = a.resize(out_pil.size, Image.LANCZOS)
        out_pil.putalpha(a_up)

    print(f"[ESRGAN] Replicate done: {img_pil.width}x{img_pil.height} -> {out_pil.width}x{out_pil.height}", flush=True)
    return out_pil


# ── Public API ────────────────────────────────────────────────────────────────

def enhance_pil(img_pil: Image.Image, scale: int = 4) -> Image.Image:
    """
    Upscale a PIL Image using the configured backend.
    scale param is used by local backend only (Replicate targets 8MP).
    Raises ImportError / ValueError if backend is not configured.
    """
    backend = UPSCALE_BACKEND

    if backend == "replicate":
        return _enhance_replicate(img_pil)

    if backend == "local":
        return _enhance_local(img_pil, scale)

    # "auto" — try local first, fall back to replicate
    try:
        return _enhance_local(img_pil, scale)
    except ImportError:
        print("[ESRGAN] Local not available, trying Replicate...", flush=True)
        return _enhance_replicate(img_pil)


def enhance_file(input_path: str, output_path: str, scale: int = 4) -> str:
    img = Image.open(input_path)
    out = enhance_pil(img, scale=scale)
    out.save(output_path, format='PNG')
    print(f"[ESRGAN] Saved {out.width}x{out.height} -> {output_path}", flush=True)
    return output_path


# ── CLI ───────────────────────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 3:
        print("Usage: python esrgan_helper.py <input> <output> [scale=4]", file=sys.stderr)
        sys.exit(1)
    inp   = sys.argv[1]
    out   = sys.argv[2]
    scale = int(sys.argv[3]) if len(sys.argv) > 3 else 4
    try:
        enhance_file(inp, out, scale)
    except (ImportError, ValueError) as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    main()
