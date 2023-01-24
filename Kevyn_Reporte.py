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
from pptx.chart.data import CategoryChartData


# Funcion para generar archivo pptx
def generate_pptx(prs):
    print("-------")
    for shape in prs.slides[6].shapes:
        print(shape.shape_type)
        if shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX and shape.text == 'title':
            shape.text = ""
            frame2 = shape.text_frame.paragraphs[0]
            frame2.alignment = PP_ALIGN.CENTER
            run2 = frame2.add_run()
            run2.text = name
     

            
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

slide = prs.slides.add_slide(prs.slide_layouts[10])
table_placeholder = slide.shapes[0]
shape = table_placeholder.insert_table(rows=3, cols=4)
table = shape.table
cell = table.cell(0, 0)
cell.text = 'Unladen Swallow'

prs.save('test.pptx')


st.download_button(
    label="Descargar pptx",
    data=generate_pptx(prs),
    file_name="efectividad_" + date.today().strftime("%d_%m_%Y") + ".pptx",
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
)
