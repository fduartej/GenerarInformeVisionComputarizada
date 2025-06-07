import os
from dotenv import load_dotenv
from azure.cognitiveservices.vision.customvision.prediction import CustomVisionPredictionClient
from msrest.authentication import ApiKeyCredentials

# Cargar variables de entorno
load_dotenv()
prediction_key = os.getenv("CUSTOM_VISION_KEY")
endpoint = os.getenv("CUSTOM_VISION_ENDPOINT")
project_id = os.getenv("CUSTOM_VISION_PROJECT_ID")
iteration = os.getenv("CUSTOM_VISION_ITERATION")

# Ruta de la imagen a probar
#ruta_img = r"evidencias\1003187\Visita_1\TimePhoto_20250425_121523.jpg" 
#ruta_img = r"evidencias\1003187\Visita_1\TimePhoto_20250425_104241.jpg"
ruta_img = r"evidencias\1003187\Visita_1\TimePhoto_20250505_091029.jpg"

# Inicializar cliente
credentials = ApiKeyCredentials(in_headers={"Prediction-key": prediction_key})
predictor = CustomVisionPredictionClient(endpoint, credentials)

with open(ruta_img, "rb") as image_data:
    results = predictor.classify_image(project_id, iteration, image_data.read())

print(f"Resultados para {ruta_img}:")
for prediction in results.predictions:
    print(f"Etiqueta: {prediction.tag_name}, Probabilidad: {prediction.probability:.2%}")