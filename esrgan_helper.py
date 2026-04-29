"""
esrgan_helper.py
Called as: py -3.10 esrgan_helper.py <input_path> <output_path> <scale>
Upscales input image with Real-ESRGAN and saves result to output_path.
"""
import sys
import numpy as np
from PIL import Image
from basicsr.archs.rrdbnet_arch import RRDBNet
from realesrgan import RealESRGANer

def main():
    if len(sys.argv) < 4:
        print("Usage: esrgan_helper.py <input> <output> <scale>", file=sys.stderr)
        sys.exit(1)

    inp, out, scale = sys.argv[1], sys.argv[2], float(sys.argv[3])
    model_path = r'C:\Varsany\realesrgan\RealESRGAN_x4plus.pth'

    img = Image.open(inp).convert('RGBA')
    alpha = img.split()[3]
    rgb   = img.convert('RGB')
    bgr   = np.array(rgb)[:, :, ::-1]

    model = RRDBNet(num_in_ch=3, num_out_ch=3, num_feat=64,
                    num_block=23, num_grow_ch=32, scale=4)
    upsampler = RealESRGANer(
        scale=4, model_path=model_path, model=model,
        tile=128, tile_pad=10, pre_pad=0, half=True)

    out_bgr, _ = upsampler.enhance(bgr, outscale=max(2.0, scale))
    out_rgb    = out_bgr[:, :, ::-1]
    result     = Image.fromarray(out_rgb)

    alpha_up = alpha.resize(result.size, Image.LANCZOS)
    result   = result.convert('RGBA')
    result.putalpha(alpha_up)
    result.save(out, format='PNG')
    print(f"OK {result.width}x{result.height}")

if __name__ == '__main__':
    main()
