import pandas as pd
from plotly.offline import plot
import numpy as np
import plotly.offline as py
import plotly.graph_objs as go
import plotly.express as px
import matplotlib.pyplot as plt
import calendar
from datetime import datetime
import colorlover as cl
import plotly.offline as offline

def mensal_bar(mes,tipo,ano,link,df): #ex: 1,Presencial
	df['DATA'] = pd.to_datetime(df['DATA'], dayfirst=True)
	df         = df.loc[df['DATA'].dt.month == mes]
	df         = df.loc[df['DATA'].dt.year == ano]
	if tipo   != 'todos':
		df     = df[df['tipo']== tipo]
	if len(df)==0:
		return 'nan'
	else:
		df['Hora'] = df['HORAS'].apply(lambda x: x.hour + x.minute / 60 + x.second / 3600)
		df_sum     = df.groupby(['GRUPO','SUBCATEGORIA']).agg({'Hora': 'sum'}).reset_index()
		gray_palette = cl.scales['9']['seq']['Greys']
		fig = px.bar(df_sum, x='Hora', y='GRUPO', color='SUBCATEGORIA', orientation='h',color_discrete_sequence=gray_palette)
		fig.update_layout(title='Horas total trabalhadas por categoria e subcategoria')
		fig.update_layout( xaxis_title='Horas',yaxis_title='Categoria',legend_title='Subcategoria')
		if link == 1:
			fig=offline.plot(fig,output_type='div')
		return fig
	
def diario_bar (dia,tipo,df): #ex: 10/01/2022,Remoto
	df['DATA'] = pd.to_datetime(df['DATA'], dayfirst=True)
	dia        = datetime.strptime(dia, '%d/%m/%Y')
	df         = df[df['DATA']== dia]
	if tipo   != 'todos':
		df     = df[df['tipo']== tipo]
	if len(df)==0:
		return 'nan'
	else:
		df['Hora'] = df['HORAS'].apply(lambda x: x.hour + x.minute / 60 + x.second / 3600)
		df_sum     = df.groupby(['GRUPO','SUBCATEGORIA']).agg({'Hora': 'sum'}).reset_index()
		fig = px.bar(df_sum, x='Hora', y='GRUPO', color='SUBCATEGORIA', orientation='h')
		fig.update_layout( xaxis_title='Horas',yaxis_title='Categoria',legend_title='Subcategoria')
		fig.update_layout(title='Horas total trabalhadas no dia por categoria e subcategoria')
		return fig
	
	
def mensal_line(mes,tipo,ano,link,df): #ex: 1,Remoto
	year       = ano
	month      = mes
	start_date = pd.Timestamp(year, month, 1)
	end_date   = start_date + pd.offsets.MonthEnd(0)
	df['DATA'] = pd.to_datetime(df['DATA'], dayfirst=True)
	df         = df.loc[df['DATA'].dt.month == mes]
	df         = df.loc[df['DATA'].dt.year == ano]
	if tipo   != 'todos':
		df     = df[df['tipo']== tipo]
	if len(df)==0:
		return 'nan'
	else:
		df['Hora'] = df['HORAS'].apply(lambda x: x.hour + x.minute / 60 + x.second / 3600)
		df_sum     = df.groupby(['GRUPO','DATA']).agg({'Hora': 'sum'}).reset_index()
		color_map = {group: px.colors.qualitative.Plotly[i % len(px.colors.qualitative.Plotly)] for i, group in enumerate(df_sum['GRUPO'].unique())}
		colors = [color_map[group] for group in df_sum['GRUPO']]
		fig = px.bar(df_sum, x='DATA', y='Hora', color='GRUPO', orientation='v')
		#fig = go.Figure(data=[go.Scatter(x=df_sum[df_sum['GRUPO']==group]['DATA'], y=df_sum[df_sum['GRUPO']==group]['Hora'], mode='markers', marker=dict(color=color_map[group]), name=group) for group in df_sum['GRUPO'].unique()])
		fig.update_layout(xaxis_title='Data', yaxis_title='Horas', legend_title='Grupo')
		fig.update_layout(xaxis_range=[start_date,end_date])
		fig.update_layout(title='Horas trabalhadas por dia')
		fig.update_layout(xaxis_tickmode='linear')
		fig.update_layout(xaxis_tickangle=-90)
		if link ==1:
			fig=offline.plot(fig,output_type='div')
		return fig

def mensal_todos(mes,ano,link,df): #ex: 1
	df['DATA'] = pd.to_datetime(df['DATA'], dayfirst=True)
	df         = df.loc[df['DATA'].dt.month == mes]
	df         = df.loc[df['DATA'].dt.year == ano]
	if len(df)==0:
		return 'nan'
	else:
		df['Hora'] = df['HORAS'].apply(lambda x: x.hour + x.minute / 60 + x.second / 3600)
		df_sum     = df.groupby(['tipo']).agg({'Hora': 'sum'}).reset_index()
		fig        = px.bar(df_sum, x='tipo', y='Hora', color='tipo', orientation='v')
		fig.update_layout(yaxis_title='Horas',xaxis_title=' ',legend_title='Tipo')
		fig.update_layout(title='Horas total trabalhadas por tipo')
		if link ==1:
			fig=offline.plot(fig,output_type='div')
		return fig
        
