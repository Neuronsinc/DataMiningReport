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


def generate_pptx(pptx):
    print('------')


st.set_page_config(layout="wide")
st.title("Reporte")

name = st.text_input(label="Nombre del proyecto", key='name')



prs = Presentation('template.pptx')


            
st.download_button(
    label="Descargar pptx",
    data=generate_pptx(prs),
    file_name="efectividad_" + date.today().strftime("%d_%m_%Y") + ".pptx",
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
)

