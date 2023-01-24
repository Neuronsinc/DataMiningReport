from cgitb import text
import datetime
from io import BytesIO
import io
import math
import csv
import pandas
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
        if shape.shape_type == MSO_SHAPE_TYPE.CHART:
            chart_title = shape.chart.chart_title.text_frame.text
            if chart_title == 'Género':
                genero_categories = []
                genero_series = ()
                if generos is not None:
                    generos_df = pandas.read_csv(generos, header=None)
                    genero_categories = generos_df.iloc[0].values[1:]
                    genero_series = generos_df.iloc[1].values[1:]
                    genero_sum = sum(map(int, genero_series))
                    genero_series = [int(x) / genero_sum for x in genero_series]                         
                    chart_data = CategoryChartData()
                    chart_data.categories = genero_categories
                    chart_data.add_series('Género', genero_series)       
                    shape.chart.replace_data(chart_data)
            if chart_title == 'Etapa familiar':
                etapa_categories = []
                etapa_series = ()
                if etapas is not None:
                    etapas_df = pandas.read_csv(etapas, header=None)
                    etapa_categories = etapas_df.iloc[0].values[1:]
                    etapa_series = etapas_df.iloc[1].values[1:]
                    etapa_sum = sum(map(int, etapa_series))
                    etapa_series = [int(x) / etapa_sum for x in etapa_series]                         
                    chart_data = CategoryChartData()
                    chart_data.categories = etapa_categories
                    chart_data.add_series('Etapa Familiar', etapa_series)       
                    shape.chart.replace_data(chart_data)
            if chart_title == 'Edad':
                edad_categories = []
                edad_series = ()
                if edades is not None:
                    df = pandas.read_csv(edades, header=None)
                    edad_categories = df.iloc[:, 0].values[1:]
                    edad_series = df.iloc[:, 1].values[1:]
                    edad_sum = sum(map(int, edad_series))
                    edad_series = [int(x) / edad_sum for x in edad_series]                         
                    chart_data = CategoryChartData()
                    chart_data.categories = edad_categories
                    chart_data.add_series('Edad', edad_series)       
                    shape.chart.replace_data(chart_data)
            if chart_title == 'Ocupaciones':
                ocupacion_categories = []
                ocupacion_series = ()
                if ocupaciones is not None:
                    ocupaciones_df = pandas.read_csv(ocupaciones, header=None)
                    ocupacion_categories = ocupaciones_df.iloc[0].values[1:]
                    ocupacion_series = ocupaciones_df.iloc[1].values[1:]
                    ocupacion_sum = sum(map(int, ocupacion_series))
                    ocupacion_series = [int(x) / ocupacion_sum for x in ocupacion_series]                         
                    chart_data = CategoryChartData()
                    chart_data.categories = ocupacion_categories
                    chart_data.add_series('Ocupaciones', ocupacion_series)       
                    shape.chart.replace_data(chart_data)
            if chart_title == 'Intereses':
                interes_categories = []
                interes_series = ()
                if intereses is not None:
                    intereses_df = pandas.read_csv(intereses, header=None)
                    interes_categories = intereses_df.iloc[0].values[1:]
                    interes_series = intereses_df.iloc[1].values[1:]
                    interes_sum = sum(map(int, interes_series))
                    interes_series = [int(x) / interes_sum for x in interes_series]                         
                    chart_data = CategoryChartData()
                    chart_data.categories = interes_categories
                    chart_data.add_series('Intereses', interes_series)       
                    shape.chart.replace_data(chart_data)
    binary_output = BytesIO()
    prs.save(binary_output)
    return binary_output.getvalue()


st.set_page_config(layout="wide")
st.title("Reporte")

name = st.text_input(label="Nombre del proyecto", key='name')

edades = st.file_uploader('CSV Edades')
ocupaciones = st.file_uploader('CSV Ocupaciones')
etapas = st.file_uploader('CSV Etapas')
intereses = st.file_uploader('CSV Intereses')
generos = st.file_uploader('CSV Generos')


r = requests.get(
        "https://github.com/Neuronsinc/DataMiningReport/blob/main/template.pptx?raw=true"
    )
prs = Presentation(io.BytesIO(r.content))



st.download_button(
    label="Descargar pptx",
    data=generate_pptx(prs),
    file_name="efectividad_" + date.today().strftime("%d_%m_%Y") + ".pptx",
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
)
