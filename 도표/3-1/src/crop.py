import subprocess, sys
from PIL import Image
pdf = sys.argv[1]; out = sys.argv[2]
subprocess.run(["pdftoppm", "-png", "-r", "400", "-f", "2", "-l", "2", pdf, "render/hi"], check=True)
import glob
im = Image.open(sorted(glob.glob("render/hi*.png"))[-1])
dpc = 400 / 2.54
box = (int((1.6 - 0.3) * dpc), int((0.8) * dpc), int((1.6 + 9.25 + 0.3) * dpc), int((2.9 + 5.1 + 1.6) * dpc))
im.crop(box).save(out)
