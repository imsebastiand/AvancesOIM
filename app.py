import dash
from dash import dcc
from dash import html
import pandas as pd
import plotly.express as px
from dash import dash_table

from PIL import Image
import requests
from io import BytesIO

dfproyectosnombre =  pd.read_excel('OIMV33.xlsx')
dfproyectosnombre['COD_proy']= dfproyectosnombre['COD_proy'].fillna("Sin Código")
project_list = dfproyectosnombre['COD_proy'].unique()

dfproyectosnombre = pd.read_excel('OIMV33.xlsx')
dfproyectosnombre['COD_proy']= dfproyectosnombre['COD_proy'].fillna("Sin Código")
listaProyectos =dfproyectosnombre['COD_proy'].unique()

external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

url = "https://joseluisrosado.com/OIMPeru.jpeg"
response = requests.get(url)
image = Image.open(BytesIO(response.content))
new_width = 382
new_height = 139
resized_image = image.resize((new_width, new_height))

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

app.layout = html.Div([
    html.Div([
        html.Div([
            html.H2("Avance de Proyectos",
                    style={'color': 'darkblue',
                           'padding': '12px 12px 6px 12px', 'margin': '0px'}),
            html.P("Enero 2021 - Diciembre 2022",
                   style={'color': 'darkblue',
                          'padding': '12px 12px 6px 12px', 'margin': '0px'}
                   ),
        ], className='eight columns',
        ),
        html.Div([
            html.Img(src=resized_image)
        ], className='four columns',
        ),
    ], className='twelve columns' ,),
    dcc.Dropdown(
        id='project-dropdown',
        options=[{'label': project, 'value': project} for project in listaProyectos],
        value=None
    ),
    html.Br(),
    html.Div(id='table-container' ,)
])


@app.callback(
    dash.dependencies.Output('table-container', 'children'),
    [dash.dependencies.Input('project-dropdown', 'value')]
)
def update_table(selected_project):
    if selected_project is None:
        return dash_table.DataTable()

    # Load the Excel
    try:
        project_df = pd.read_excel(selected_project + '.xlsx')
        project_df.dropna(how='all', inplace=True)

    except FileNotFoundError:
        return html.Div(f'El archivo {selected_project} no está disponible.')

    project_df['Avance(%)'] = project_df['Avance(%)'].round(2)

    # Define the conditional formatting rules
    formatting_rules = [
        {
            'if': {
                'filter_query': '{Avance(%)} < 34',
                'column_id': 'Avance(%)'
            },
            'backgroundColor': 'red',
            'color': 'white'
        },
        {
            'if': {
                'filter_query': '{Avance(%)} >= 34 && {Avance(%)} < 67',
                'column_id': 'Avance(%)'
            },
            'backgroundColor': 'yellow'
        },
        {
            'if': {
                'filter_query': '{Avance(%)} >= 67 && {Avance(%)} < 100',
                'column_id': 'Avance(%)'
            },
            'backgroundColor': 'lightgreen'
        },
        {
            'if': {
                'filter_query': '{Avance(%)} >= 100',
                'column_id': 'Avance(%)'
            },
            'backgroundColor': 'darkgreen',
            'color': 'white'
        }
    ]

    table = dash_table.DataTable(
        id='project-table',
        columns=[{'name': col, 'id': col} for col in project_df.columns],
        data=project_df.to_dict('records'),
        style_cell={
            'textAlign': 'center',
            'whiteSpace': 'normal',
            'height': 'auto',
        },
        style_data=True,
        style_data_conditional=formatting_rules,
        filter_action="native",
        sort_action="native",
        sort_mode="multi",
        page_action="native",
        page_current=0,
        page_size=10
    )

    return table

if __name__ == '__main__':
    app.run_server(debug=False)
