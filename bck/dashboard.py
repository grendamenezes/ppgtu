# Run this app with `python app.py` and
# visit http://127.0.0.1:8050/ in your web browser.
import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output, State
import pandas as pd
import datetime
import locale
import graficos
import zipfile
import zip_gera
import matplotlib.pyplot as plt

df         = pd.read_excel('atividade.xlsx')
df['DATA'] = pd.to_datetime(df['DATA'], dayfirst=True)


def download_zip(mes,ano,nome):
	locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
	month_name = datetime.date(2000, mes, 1).strftime('%B')
	month_name = month_name.capitalize()	
	zip_path = 'relatorio_'+month_name+'.zip'
	with zipfile.ZipFile(zip_path, 'w') as zip_file:
		zip_file.write(zip_gera.preenche_modelo(mes,ano,nome))
		zip_file.writestr('horas_'+month_name+'_Presencial.html', graficos.mensal_bar(mes,'Presencial',ano,1))
		zip_file.writestr('horas_'+month_name+'_Remoto.html', graficos.mensal_bar(mes,'Remoto',ano,1))
		zip_file.writestr('horas_'+month_name+'_todos.html', graficos.mensal_bar(mes,'todos',ano,1))
		zip_file.writestr('horas_por_dia_'+month_name+'_Presencial.html', graficos.mensal_line(mes,'Presencial',ano,1))
		zip_file.writestr('horas_por_dia_'+month_name+'_Remoto.html', graficos.mensal_line(mes,'Remoto',ano,1))
		zip_file.writestr('horas_por_dia_'+month_name+'_todos.html', graficos.mensal_line(mes,'todos',ano,1))
		zip_file.writestr('horas_por_tipo_'+month_name+'.html', graficos.mensal_todos(mes,ano,1))
	return dcc.send_file(zip_path, filename=zip_path)



# Initialize the app
app = dash.Dash(__name__)

# Define the layout
app.layout = html.Div([
    html.H1('Relatório Atividades PPGTU'),
    html.Div(id='tipo-container', children=[
    dcc.RadioItems(
        id='freq-tipo',
        options=[
            {'label': 'Presencial', 'value': 'Presencial'},
            {'label': 'Remoto', 'value': 'Remoto'},
            {'label': 'Todos', 'value': 'todos'}
        ], value=None
    )]),
    html.Div(id='rela-container', children=[
    dcc.RadioItems(
        id='freq-radio',
        options=[
            {'label': 'Mensal', 'value': 'mensal'},
            {'label': 'Diário', 'value': 'diario'}
        ], value=None
    )]),
    
    html.Div(id='mensal-container', children=[
        html.Label('Ano:'),
        html.Div([
        dcc.Input(id='year-input', type='number', placeholder='Ano')]),
        html.Div([
        html.Label('Mês:'),
        dcc.Dropdown(
            id='month-dropdown',
            options=[
                {'label': 'Janeiro'  , 'value': '01'},
                {'label': 'Fevereiro', 'value': '02'},
                {'label': 'Março'    , 'value': '03'},
                {'label': 'Abril'    , 'value': '04'},
                {'label': 'Maio'     , 'value': '05'},
                {'label': 'Junho'    , 'value': '06'},
                {'label': 'Julho'    , 'value': '07'},
                {'label': 'Agosto'   , 'value': '08'},
                {'label': 'Setembro' , 'value': '09'},
                {'label': 'Outubro'  , 'value': '10'},
                {'label': 'Novembro' , 'value': '11'},
                {'label': 'Dezembro' , 'value': '12'}
            ],
            placeholder='Mês'
        )]),
        html.Button('Enter', id='submit-btn')
    ], style={'display': 'none'}),
    
    html.Div(id='diario-container', children=[
        dcc.Input(id='date-input', type='text', placeholder='DD/MM/YYYY'),
        html.Button('Enter', id='submit-btn-2')
    ], style={'display': 'none'}),
    html.Div(id="mensal-graphs1", children=[
        dcc.Graph(id="graph-1-mes"),
        dcc.Graph(id="graph-2")
    ], style={'display': 'none'}),
    html.Div(id="mensal-graphs2", children=[
        dcc.Graph(id="graph-1-1-mes"),
        dcc.Graph(id="graph-2-2"),
        dcc.Graph(id="graph-3")
    ], style={'display': 'none'}),
    html.Div(id="diario-graphs", children=[
        dcc.Graph(id="graph-1-dia")
    ], style={'display': 'none'}),
    html.Div(id="gerar",children=[html.H2('Gerar relatório'), html.Label('Aluno(a):'),
                                  dcc.Input(id='name-input', type='text', placeholder='Digite seu nome'),
                                  html.Br(),
                                  html.Button('Download', id='download-link'),
                                  dcc.Download(id='download')],style={'display': 'block'})
])

'''
html.Div(id="gerar",children=[html.Label('Aluno(a):'),
				dcc.Input(id='name-input', type='text', 
				placeholder='Digite seu nome'),html.Button('Gerar relatório', id='relato')],style={'display': 'none'})
				'''

# Define the callbacks
@app.callback(
    [Output('mensal-container', 'style'), Output('diario-container', 'style'),Output("mensal-graphs1", "style"),Output("mensal-graphs2", "style"),
    Output("diario-graphs", "style"),Output("gerar", "style")],
    [Input('freq-radio', 'value'),Input('freq-tipo', 'value')]
)
def show_hide_divs(frequency,tipo):
	mensal_style         = {'display': 'block'} if frequency == "mensal" else {'display': 'none'}
	diario_style         = {'display': 'block'} if frequency == "diario" else {'display': 'none'}
	mensal_graphs1_style = {'display': 'block'} if frequency == "mensal" and tipo != 'todos' else {'display': 'none'}
	mensal_graphs2_style = {'display': 'block'} if frequency == "mensal" and tipo == 'todos' else {'display': 'none'}
	diario_graphs_style  = {'display': 'block'} if frequency == "diario" else {'display': 'none'}
	gerar_style          = {'display': 'block'} if frequency == "mensal" and tipo == 'todos' else {'display': 'none'}
	return mensal_style, diario_style, mensal_graphs1_style, mensal_graphs2_style, diario_graphs_style,gerar_style


@app.callback(
    [Output('graph-1-mes', 'figure'), Output('graph-2', 'figure')],
    [Input('submit-btn', 'n_clicks'),Input('freq-tipo', 'value')],
    [State('year-input', 'value'), State('month-dropdown', 'value')]
)
def update_graphs_1(n_clicks,tipo, year, month):
	if not year or not month:
		return {}, {}
	fig1=graficos.mensal_bar(int(month),tipo,year,0)
	fig2=graficos.mensal_line(int(month),tipo,year,0)
	if fig1 == 'nan' or fig2 == 'nan':
		return {},{}
	else:
		return fig1, fig2
				
@app.callback(
    [Output('graph-1-1-mes', 'figure'), Output('graph-2-2', 'figure'), Output('graph-3', 'figure'),Output('download', 'data')],
    [Input('submit-btn', 'n_clicks'),Input('freq-tipo', 'value'),Input('download-link', 'n_clicks'),Input('name-input', 'value')],
    [State('year-input', 'value'), State('month-dropdown', 'value')]
)
def update_graphs_2(n_clicks,tipo,n_clicks2,nome, year, month):
	if not year or not month:
		return {}, {},{},None
	fig1=graficos.mensal_bar(int(month),tipo,year,0)
	fig2=graficos.mensal_line(int(month),tipo,year,0)
	fig3=graficos.mensal_todos(int(month),year,0)
	if fig1 == 'nan' or fig2 == 'nan' or fig3 == 'nan':
		return {},{},{},None
	elif n_clicks2 is not None:
		return fig1, fig2,fig3,download_zip(int(month),year,str(nome))
	else:
		return fig1, fig2,fig3, None


@app.callback(
    Output('graph-1-dia', 'figure'),
    [Input('submit-btn-2', 'n_clicks'),Input('freq-tipo', 'value')],
    [State('date-input', 'value')]
)
def update_graphs_3(n_clicks,tipo, date):
	if not date:
		return {}
	fig = graficos.diario_bar (date,tipo)
	if fig == 'nan':
		return {}
	else:
		return fig

if __name__ == '__main__':
    app.run_server(debug=True)

