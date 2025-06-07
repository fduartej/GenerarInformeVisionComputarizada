from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
import os

doc = DocxTemplate("INFORME_PLANTILLA.docx")
img_path = os.path.abspath(r"evidencias\1003187\Visita_1\TimePhoto_20250425_104241_saneada.jpg")
contexto = {"FOTO_MEDIDOR": InlineImage(doc, img_path, width=Inches(3))}
doc.render(contexto)
doc.save("test_output.docx")
print("Listo")