import os
from dotenv import load_dotenv
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from azure.cognitiveservices.vision.customvision.prediction import CustomVisionPredictionClient
from msrest.authentication import ApiKeyCredentials

# Cargar variables de entorno
load_dotenv()
prediction_key = os.getenv("CUSTOM_VISION_KEY")
endpoint = os.getenv("CUSTOM_VISION_ENDPOINT")
project_id = os.getenv("CUSTOM_VISION_PROJECT_ID")
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
df = pd.read_excel(excel_file)

# Procesar cada fila (una visita)
for _, row in df.iterrows():
    cliente = str(row["cliente_id"])
    visita = str(row["visita"])
    carpeta = os.path.join(evidencia_dir, cliente, f"visita_{visita}")
    if not os.path.isdir(carpeta):
        print(f"❌ Carpeta no encontrada: {carpeta}")
        continue

    # Clasificar imágenes relevantes
    tags_interes = {
        "medidor_antes": None,
        "bypass": None,
        "medidor_cortado": None
    }

    for archivo in os.listdir(carpeta):
        if not archivo.lower().endswith((".jpg", ".jpeg", ".png")):
            continue
        ruta_img = os.path.join(carpeta, archivo)
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
        "cliente": row["cliente_id"],
        "direccion": row["direccion"],
        "fecha": str(row["fecha"]),
        "tecnico": row["tecnico"],
        "observacion": row["observacion"],
        "FOTO_MEDIDOR_ANTES": InlineImage(doc, tags_interes["medidor_antes"]["path"], width=Inches(3)) if tags_interes["medidor_antes"] else "",
        "FOTO_BYPASS": InlineImage(doc, tags_interes["bypass"]["path"], width=Inches(3)) if tags_interes["bypass"] else "",
        "FOTO_MEDIDOR_CORTADO": InlineImage(doc, tags_interes["medidor_cortado"]["path"], width=Inches(3)) if tags_interes["medidor_cortado"] else "",
    }

    # Renderizar y guardar
    output_path = os.path.join(output_dir, f"informe_{cliente}_visita_{visita}.docx")
    doc.render(contexto)
    doc.save(output_path)
    print(f"✅ Informe generado: {output_path}")
