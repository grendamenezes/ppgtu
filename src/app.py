import dash
import dash_html_components as html
import dash_table
import dash_core_components as dcc
from dash.dependencies import Input, Output, State
import pandas as pd

# Criando um DataFrame com dados de exemplo
df = pd.DataFrame({
    'Nome': ['Alice', 'Bob', 'Charlie', 'David'],
    'Idade': [25, 30, 35, 40],
    'Cidade': ['Rio de Janeiro', 'São Paulo', 'Belo Horizonte', 'Curitiba']
})

# Criando o aplicativo Dash
app = dash.Dash(__name__)
server=app.server

# Definindo o layout do dashboard
app.layout = html.Div(children=[
    html.H1(children='Dashboard'),
    dash_table.DataTable(
        id='tabela',
        columns=[{'name': col, 'id': col} for col in df.columns],
        data=df.to_dict('records')
    ),
    html.Button('Baixar tabela', id='botao-download'),
    dcc.Download(id='download')
])

# Definindo o callback para gerar o arquivo Excel para download
@app.callback(Output('download', 'data'),
              Input('botao-download', 'n_clicks'),
              State('tabela', 'data'))
def download_table(n_clicks, data):
    if n_clicks is None:
        return None
    df = pd.DataFrame.from_dict(data)
    return dcc.send_data_frame(df.to_excel, "tabela.xlsx", sheet_name='Sheet1', index=False)

if __name__ == '__main__':
    app.run_server(debug=True)
