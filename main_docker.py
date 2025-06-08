import os
from dotenv import load_dotenv
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from PIL import Image
import requests

def sanear_imagen(origen):
    destino = origen.replace(".jpg", "_saneada.jpg").replace(".jpeg", "_saneada.jpeg").replace(".png", "_saneada.png")
    try:
        img = corregir_orientacion(origen)
        img = img.convert("RGB")
        img = img.rotate(-90, expand=True)  # fuerza giro antihorario
        img.save(destino, "JPEG")
        return destino
    except Exception as e:
        print(f"⚠️ No se pudo sanear la imagen {origen}: {e}")
        return origen

def recortar_contador_gas(imagen_path, predicciones):
    # Cargar imagen y obtener dimensiones
    imagen = Image.open(imagen_path)
    width, height = imagen.size

    # Filtrar solo detecciones con etiqueta 'contador_gas' y probabilidad ≥ 0.08
    detecciones_validas = [
        p for p in predicciones
        if p["tagName"] == "contador_gas" and p["probability"] >= 0.08
    ]

    if not detecciones_validas:
        print(f"⚠️ No se encontró contador_gas confiable en {imagen_path}")
        return None

    # Tomar la detección con mayor probabilidad
    mejor = max(detecciones_validas, key=lambda x: x["probability"])
    box = mejor["boundingBox"]

    # Calcular coordenadas reales
    x1 = int(box["left"] * width)
    y1 = int(box["top"] * height)
    x2 = int((box["left"] + box["width"]) * width)
    y2 = int((box["top"] + box["height"]) * height)

    # Recortar y devolver
    return imagen.crop((x1, y1, x2, y2))

def corregir_orientacion(path):
    try:
        img = Image.open(path)
        for orientation in ExifTags.TAGS.keys():
            if ExifTags.TAGS[orientation] == 'Orientation':
                break
        exif = img._getexif()
        if exif is not None:
            orient = exif.get(orientation, None)
            if orient == 3:
                img = img.rotate(180, expand=True)
            elif orient == 6:
                img = img.rotate(270, expand=True)
            elif orient == 8:
                img = img.rotate(90, expand=True)
        return img
    except Exception as e:
        print(f"⚠️ No se pudo corregir la orientación de {path}: {e}")
        return Image.open(path)

def es_imagen_valida(path):
    try:
        with Image.open(path) as img:
            img.verify()
        with Image.open(path) as img:
            img.load()
            img.convert("RGB")
        return True
    except Exception as e:
        print(f"⚠️ Imagen corrupta detectada: {path} - {e}")
        return False

def imagen_inline(doc, info):
    if info and "path" in info:
        path = info["path"]
        print(f"Intentando insertar imagen: {path}")
        if os.path.isfile(path) and es_imagen_valida(path):
            try:
                path_saneada = sanear_imagen(path)
                return InlineImage(doc, path_saneada, width=Inches(3))
            except Exception as e:
                print(f"⚠️ Error insertando imagen: {path} - {e}")
        else:
            print(f"⚠️ Imagen inválida o corrupta (no insertada): {path}")
    return ""

def clasificar_desde_docker(imagen_path):
    try:
        with open(imagen_path, "rb") as f:
            res = requests.post(
                "http://127.0.0.1:5001/image",
                data=f,
                headers={"Content-Type": "application/octet-stream"}
            )
        if res.status_code != 200:
            print(f"❌ Error en predicción Docker para {imagen_path}: {res.status_code}")
            return []
        return res.json().get("predictions", [])
    except Exception as e:
        print(f"❌ Excepción al clasificar imagen {imagen_path}: {e}")
        return []


# Cargar variables de entorno
load_dotenv()
excel_file = os.getenv("EXCEL_FILE")
template_file = os.getenv("TEMPLATE_FILE")
evidencia_dir = os.getenv("EVIDENCIA_DIR")
output_dir = os.getenv("OUTPUT_DIR")

os.makedirs(output_dir, exist_ok=True)

# Leer Excel de visitas
df = pd.read_excel(excel_file, sheet_name="completo", dtype=str)

# Procesar cada fila (una visita)
for _, row in df.iterrows():
    cpno = str(row["CPNO"])
    cuentaContrato = str(row["CUENTA CONTRATO"])
    visita = "1"
    carpeta = os.path.join(evidencia_dir, cuentaContrato, f"Visita_{visita}")
    print(f"Procesando visita {visita} para cuenta {cuentaContrato} (CPNO: {cpno})...carpeta: {carpeta}")
    if not os.path.isdir(carpeta):
        print(f"❌ Carpeta no encontrada: {carpeta}")
        continue

    tags_interes = {
        "medidor": None,
        "sin_medidor": None,
        "bolsa_plastica": None
    }

    for archivo in os.listdir(carpeta):
        if not archivo.lower().endswith((".jpg", ".jpeg", ".png")):
            continue
        ruta_img = os.path.join(carpeta, archivo)
        if not es_imagen_valida(ruta_img):
            print(f"⚠️ Imagen inválida o corrupta (omitida): {ruta_img}")
            continue

        predicciones = clasificar_desde_docker(ruta_img)
        for pred in predicciones:
            tag = pred["tagName"]
            prob = pred["probability"]
            if tag in tags_interes and prob > 0.8:
                actual = tags_interes[tag]
                if actual is None or prob > actual["prob"]:
                    tags_interes[tag] = {"path": ruta_img, "prob": prob}

    # Cargar plantilla Word
    doc = DocxTemplate(template_file)

    contexto = {
        "CUENTA_CONTRATO": row["CUENTA CONTRATO"],
        "DIRECCION": row["DIRECCIÓN"],
        "FECHA": str(row["FECHA"]),
        "RAZON_SOCIAL": row["RAZÓN SOCIAL"],
        "CPNO": row["CPNO"],
        "FOTO_MEDIDOR": imagen_inline(doc, tags_interes["medidor"]),
        "FOTO_SINMEDIDOR": imagen_inline(doc, tags_interes["sin_medidor"]),
        "FOTO_BOLSA": imagen_inline(doc, tags_interes["bolsa_plastica"]),
    }
    print(contexto)
    output_path = os.path.join(output_dir, f"informe_{cpno}_{cuentaContrato}_visita_{visita}.docx")
    doc.render(contexto)
    doc.save(output_path)
    print(f"✅ Informe generado: {output_path}")
