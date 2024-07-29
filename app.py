from dash import Dash, html, dcc, dash_table
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.express as px
from dash.dependencies import Input, Output

data = pd.read_excel('./data/eje.xlsx', engine='openpyxl')
dataa = pd.read_excel('./data/com.xlsx', engine='openpyxl')

app = Dash(external_stylesheets=[dbc.themes.CERULEAN])
server = app.server

# Barra principal
navbar = dbc.NavbarSimple(
    brand='Analisis De Datos - Upres Cauca',
    brand_style={'marginLeft': 10, 'fontFamily': 'Helvetica'},
    children=[
        html.A('Data Source', href='#', target='_blank')
    ],
    color='primary',
    fluid=True
)

# Menú
menu_drop = html.Div([
    dcc.Dropdown(data['DEP GASTO'].unique(), placeholder='Seleccione Categoria', id='eleccion')
])

# Tarjetas
tarjetas = html.Div([
    dbc.Row([
        dbc.Col(dbc.Card([
            html.H4('Valor '),
            html.H5(id='vlr'),
            html.H6('PESOS')
        ], body=True, style={'textAlign': 'center', 'height': '100%'})),
        dbc.Col(dbc.Card([
            html.H4('Cantidad '),
            html.H5(id='ctr'),
            html.H6('REGISTROS')
        ], body=True, style={'textAlign': 'center', 'height': '100%'}))
    ])
])

# Tabla de categorías
table = dash_table.DataTable(
    id='tbl-categoria',
    columns=[
        {'name': 'Numero Documento Soporte', 'id': 'Numero Documento Soporte'},
        {'name': 'Nombre Razon Social', 'id': 'Nombre Razon Social'},
        {'name': 'Valor Actual', 'id': 'Valor Actual'},
        {'name': 'Saldo por Utilizar', 'id': 'Saldo por Utilizar'}
    ],
    page_size=10,
    style_table={'height': '300px', 'overflowY': 'auto'},
    style_cell={'textAlign': 'center', 'padding': '5px', 'fontSize': '12px'},
    style_header={'backgroundColor': 'rgb(230, 230, 230)', 'fontWeight': 'bold'},
    style_data_conditional=[
        {
            'if': {
                'filter_query': '{Saldo por Utilizar} = 0'
            },
            'backgroundColor': 'green',
            'color': 'white'
        },
        {
            'if': {
                'filter_query': '{Saldo por Utilizar} > 0'
            },
            'backgroundColor': 'yellow',
            'color': 'black'
        }
    ]
)

# Gráfico
grafico = dcc.Graph(id='grapih', style={'height': '400px'})

@app.callback(
    Output('vlr', 'children'),
    Output('ctr', 'children'),
    Output('tbl-categoria', 'data'),
    Output('grapih', 'figure'),
    Input('eleccion', 'value')
)
def Actualizar(opcion):
    if opcion is None:
        valor_A = dataa["Valor Actual"].sum()
        cantidadctr_A = dataa["Valor Actual"].count()
        tabl = dataa.groupby('Numero Documento Soporte').agg({
            'Nombre Razon Social': 'first',
            'Valor Actual': 'sum',
            'Saldo por Utilizar': 'sum'
        }).reset_index().to_dict('records')
        
        total_apropiacion = data["APR, VIGENTE"].sum()
        ejecutado = dataa["Valor Actual"].sum()
        sin_ejecutar = total_apropiacion - ejecutado
        
        categorias = pd.DataFrame({
            'Categoria': ['Apropiacion Total', 'Ejecutado', 'Sin Ejecutar'],
            'Valor': [total_apropiacion, ejecutado, sin_ejecutar]
        })
        
        fig = px.bar(categorias, x='Categoria', y='Valor', title=f'Apropiacion vs Ejecutado - Sin Ejecutar: {sin_ejecutar:,.2f}')

        return f"{valor_A:,.2f}", f"{cantidadctr_A:,}", tabl, fig

    else:
        valor_A = dataa[dataa["Dependencia"] == opcion]
        
        if valor_A.empty:
            # Si no hay datos para la opción seleccionada, devolver valores predeterminados
            valorctr_AAA = 0
            cantidadctr_AAA = 0
            tabl = []
            total_apropiacion = data[data["DEP GASTO"] == opcion]["APR, VIGENTE"].sum()
            ejecutado = 0
            sin_ejecutar = total_apropiacion - ejecutado
            
            categorias = pd.DataFrame({
                'Categoria': ['Apropiacion Total', 'Ejecutado', 'Sin Ejecutar'],
                'Valor': [total_apropiacion, ejecutado, sin_ejecutar]
            })
            
            fig = px.bar(categorias, x='Categoria', y='Valor', title=f'Apropiacion vs Ejecutado para {opcion} - Sin Ejecutar: {sin_ejecutar:,.2f}')
        else:
            valorctr_AA = valor_A[["Dependencia", "Valor Actual"]].groupby("Dependencia").sum()
            valorctr_AAA = valorctr_AA["Valor Actual"].sum()
            cantidadctr_AA = valor_A[["Dependencia", "Valor Actual"]].groupby("Dependencia").count()
            cantidadctr_AAA = cantidadctr_AA["Valor Actual"].sum()
            tabl = valor_A.groupby('Numero Documento Soporte').agg({
                'Nombre Razon Social': 'first',
                'Valor Actual': 'sum',
                'Saldo por Utilizar': 'sum'
            }).reset_index().to_dict('records')

            total_apropiacion = data[data["DEP GASTO"] == opcion]["APR, VIGENTE"].sum()
            ejecutado = valor_A["Valor Actual"].sum()
            sin_ejecutar = total_apropiacion - ejecutado
            
            categorias = pd.DataFrame({
                'Categoria': ['Apropiacion Total', 'Ejecutado', 'Sin Ejecutar'],
                'Valor': [total_apropiacion, ejecutado, sin_ejecutar]
            })

            fig = px.bar(categorias, x='Categoria', y='Valor', title=f'Apropiacion vs Ejecutado para {opcion} - Sin Ejecutar: {sin_ejecutar:,.2f}')

        return f"{valorctr_AAA:,.2f}", f"{cantidadctr_AAA:,}", tabl, fig

app.layout = html.Div([
    dbc.Row(navbar),
    dbc.Row([
        dbc.Col([
            dbc.Row(menu_drop, style={'marginBottom': '10px'}),
            dbc.Row(tarjetas, style={'marginBottom': '10px'}),
            dbc.Row(table, style={'height': '300px', 'overflowY': 'auto'})
        ], width=6),
        dbc.Col([
            dbc.Row(grafico),
            dbc.Row()
        ], width=6)
    ]),
    html.Footer([
        dbc.Row([
            dbc.Col(html.P(' Todos los derechos reservados - Diego A. Hoyos Dorado', className='text-center'), width={'size': 6, 'offset': 3})
        ], style={'backgroundColor': '#f8f9fa', 'padding': '10px'})
    ], style={'position': 'fixed', 'bottom': 0, 'width': '100%', 'backgroundColor': '#f8f9fa', 'textAlign': 'center'})
])

if __name__ == '__main__':
    app.run_server(debug=True)
