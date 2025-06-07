import os
from dotenv import load_dotenv
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from azure.cognitiveservices.vision.customvision.prediction import CustomVisionPredictionClient
from msrest.authentication import ApiKeyCredentials
from PIL import Image

def sanear_imagen(origen):
    destino = origen.replace(".jpg", "_saneada.jpg").replace(".jpeg", "_saneada.jpeg").replace(".png", "_saneada.png")
    try:
        with Image.open(origen) as img:
            img = img.convert("RGB")
            img.save(destino, "JPEG")
        return destino
    except Exception as e:
        print(f"⚠️ No se pudo sanear la imagen {origen}: {e}")
        return origen  # Devuelve la original si falla

def es_imagen_valida(path):
    try:
        with Image.open(path) as img:
            img.verify()
        with Image.open(path) as img:
            img.load()
            img.convert("RGB")  # Fuerza la carga completa de datos
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


# Cargar variables de entorno
load_dotenv()
prediction_key = os.getenv("CUSTOM_VISION_KEY")
endpoint = os.getenv("CUSTOM_VISION_ENDPOINT")
project_id = os.getenv("CUSTOM_VISION_PROJECT_ID")
model_id = os.getenv("CUSTOM_VISION_MODEL_ID")
iteration = os.getenv("CUSTOM_VISION_ITERATION")
excel_file = os.getenv("EXCEL_FILE")
template_file = os.getenv("TEMPLATE_FILE")
evidencia_dir = os.getenv("EVIDENCIA_DIR")
output_dir = os.getenv("OUTPUT_DIR")

# Asegurar carpeta de salida
os.makedirs(output_dir, exist_ok=True)

# Inicializar cliente de Custom Vision
credentials = ApiKeyCredentials(in_headers={"Prediction-key": prediction_key})
predictor = CustomVisionPredictionClient(endpoint, credentials)

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

    # Clasificar imágenes relevantes
    tags_interes = {
        "medidor": None,
        "sin_medidor": None,
        "bolsa": None
    }

    for archivo in os.listdir(carpeta):
        if not archivo.lower().endswith((".jpg", ".jpeg", ".png")):
            continue
        ruta_img = os.path.join(carpeta, archivo)
        if not es_imagen_valida(ruta_img):
            print(f"⚠️ Imagen inválida o corrupta (omitida): {ruta_img}")
            continue
        with open(ruta_img, "rb") as f:
            result = predictor.classify_image(project_id, iteration, f.read())
            for pred in result.predictions:
                tag = pred.tag_name
                if tag in tags_interes and pred.probability > 0.8:
                    actual = tags_interes[tag]
                    if actual is None or pred.probability > actual["prob"]:
                        tags_interes[tag] = {"path": ruta_img, "prob": pred.probability}

    # Cargar plantilla Word
    doc = DocxTemplate(template_file)

    # Construir contexto con datos + imágenes
    contexto = {
        "CUENTA_CONTRATO": row["CUENTA CONTRATO"],
        "DIRECCION": row["DIRECCIÓN"],
        "FECHA": str(row["FECHA"]),
        "RAZON_SOCIAL": row["RAZÓN SOCIAL"],
        "CPNO": row["CPNO"],
        "FOTO_MEDIDOR": imagen_inline(doc, tags_interes["medidor"]),
        "FOTO_SINMEDIDOR": imagen_inline(doc, tags_interes["sin_medidor"]),
        "FOTO_BOLSA": imagen_inline(doc, tags_interes["bolsa"])
    }
    print(contexto)
    # Renderizar y guardar
    output_path = os.path.join(output_dir, f"informe_{cpno}_{cuentaContrato}_visita_{visita}.docx")
    doc.render(contexto)
    doc.save(output_path)
    print(f"✅ Informe generado: {output_path}")
