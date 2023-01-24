from cgitb import text
import datetime
from io import BytesIO
import io
import math
import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from datetime import date
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt
import requests



# Funcion para generar archivo pptx
def generate_pptx(prs):
    print("-------")
    binary_output = BytesIO()
    prs.save(binary_output)
    return binary_output.getvalue()


st.set_page_config(layout="wide")
st.title("Reporte")

name = st.text_input(label="Nombre del proyecto", key='name')



r = requests.get(
        "https://github.com/Neuronsinc/DataMiningReport/blob/main/template.pptx?raw=true"
    )
prs = Presentation(io.BytesIO(r.content))



st.download_button(
    label="Descargar pptx",
    data=generate_pptx(prs),
    file_name="efectividad_" + date.today().strftime("%d_%m_%Y") + ".pptx",
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
<<<<<<< HEAD
)
=======
)
>>>>>>> 247ede044da201c1475b75f9f1c2d50e995a94b7
