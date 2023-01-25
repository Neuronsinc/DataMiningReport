from cgitb import text
import datetime
from io import BytesIO
import io
from pptx.util import Inches
import math
import csv
import pandas as pd
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
import dataframe_image as dfi
import numpy as np
import df2img
import plotly.graph_objects as go

def iter_cells(table):
    for row in table.rows:
        for cell in row.cells:
            yield cell

def color_negative_red(value):
    color = 'green'
    return 'color: %s' % color


def highlight(s):
    if s.duration < 3:
        return ['background-color: yellow'] * len(s)
    else:
        return ['background-color: red'] * len(s)



# Funcion para generar archivo pptx
def generate_pptx(prs):
    print("-------")
    for shape in prs.slides[6].shapes:
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
                    generos_df = pd.read_csv(generos, header=None)
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
                    etapas_df = pd.read_csv(etapas, header=None)
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
                    df = pd.read_csv(edades, header=None)
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
                    ocupaciones_df = pd.read_csv(ocupaciones, header=None)
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
                    intereses_df = pd.read_csv(intereses, header=None)
                    interes_categories = intereses_df.iloc[0].values[1:]
                    interes_series = intereses_df.iloc[1].values[1:]
                    interes_sum = sum(map(int, interes_series))
                    interes_series = [int(x) / interes_sum for x in interes_series]                         
                    chart_data = CategoryChartData()
                    chart_data.categories = interes_categories
                    chart_data.add_series('Intereses', interes_series)       
                    shape.chart.replace_data(chart_data)
    for shape in prs.slides[8].shapes:
        print(shape.shape_type)

    if keyword_file is not None:
        keyword_df = pd.read_csv(keyword_file)
        col_formats = {'Competition': '.2%'}
        plot_df = keyword_df.reset_index()
        font_colours_df = pd.DataFrame(
            'black',  # Set default font colour
            index=plot_df.index, columns=plot_df.columns
        )
                
        colors = ['rgb(239, 243, 255)', 'rgb(189, 215, 231)', 'rgb(107, 174, 214)']
        data = {'Year' : [2010, 2011, 2012], 'Color' : colors}
        df = pd.DataFrame(data)

        fig = go.Figure(data=[go.Table(
                                header=dict(
                                values=["Color", "<b>YEAR</b>"],
                                line_color='white', fill_color='white',
                                align='center', font=dict(color='black', size=12)
                            ),
                            cells=dict(
                                values=[df.Color, df.Year],
                                line_color=[df.Color]
                                , fill_color=[df.Color]
                                , align='center'
                                , font=dict(color='black', size=11)
                            ))
                        ])
        #fig.show()  
        print("figura")
        fig.write_image('plot1.png')
       


        row, col = keyword_df.shape

        shapes = prs.slides[8].shapes
        left = Inches(1.0)
        top = Inches(1.0)
        width = Inches(1.0)
        height = Inches(1.0)


        # table = shapes.add_table(row, col, left, top, width, height).table
        promedio = 0.0

        # for i in range(col):
        #     table.columns[i].width = Inches(1.0)
        #     table.cell(0, i).text = encabezados[i]

        colores = []
        colores = ['rgb(255, 255, 255)' for i in range(row-1)]
        

        for i in range(col):            
            for j in range(row):
                if j == 0 or j == 1 or j == 2 :
                        colores[j] = 'rgb(252, 248, 3)'
                if j + 1 < row :                    
                    # table.cell(j + 1, i).text = str(keyword_df.iloc[j + 1, i]) 
                    if i == 1 :
                        promedio += float(keyword_df.iloc[j + 1, i])
                    
                    # if i > 0 and i < 4 :
                    #     colores.append ('rgb(235, 222, 52)')
                    # else :
                    #     colores.append ('rgb(255, 255, 255)')

        print('colores')
        print(colores)
        print('promedio')
        print(promedio)

        shapes.add_picture('plot1.png', Inches(1), Inches(1))
#row col
        # set column widths
        # table.columns[0].width = Inches(1.0)
        # table.columns[1].width = Inches(1.0)
        # table.columns[2].width = Inches(1.0)
        # table.columns[3].width = Inches(1.0)
        # table.columns[4].width = Inches(1.0)


        # write column headings
        # table.cell(0, 0).text = 'Keyword'
        # table.cell(0, 1).text = 'Búsqueda Promedio'
        # table.cell(0, 2).text = 'Impresiones'
        # table.cell(0, 3).text = 'Competencia'
        # table.cell(0, 4).text = 'CPC Promedio'

        # write body cells
        # table.cell(1, 0).text = 'Baz'
        # table.cell(1, 1).text = 'Qux'
        # for cell in iter_cells(table):
        #     for paragraph in cell.text_frame.paragraphs:
        #         for run in paragraph.runs:
        #             run.font.size = Pt(1)
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

keyword_file = st.file_uploader('Google Ads')


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
