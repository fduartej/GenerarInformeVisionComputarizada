from PIL import Image
import os

img_path = os.path.abspath(r"evidencias\1003187\Visita_1\TimePhoto_20250505_091029.jpg")
img_saneada = os.path.abspath(r"evidencias\1003187\Visita_1\TimePhoto_20250505_091029_saneada.jpg")

try:
    with Image.open(img_path) as img:
        img = img.convert("RGB")
        img.save(img_saneada, "JPEG")
    print(f"Imagen saneada guardada en: {img_saneada}")
except Exception as e:
    print(f"Error al sanear la imagen: {e}")