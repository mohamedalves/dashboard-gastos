import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output, State
import plotly.express as px
import pandas as pd
import base64
import io

# Função para carregar os dados
def load_data():
    excel_file = pd.ExcelFile("gastos.xlsx")
    abas = {
        'Mohamed': pd.read_excel(excel_file, sheet_name='Mohamed'),
        'Evelyn': pd.read_excel(excel_file, sheet_name='Evelyn')
    }

    for pessoa, df in abas.items():
        if pessoa == 'Evelyn':
            df.columns = df.columns.str.strip()
            rename_dict = {
                'MES': 'Mês',
                'CATEGORIA': 'Categoria',
                'DESPESAS': 'Despesas',
                'VALOR (RS)': 'Valor (R$)',
                'TOTAL CATEGORIA': 'Total Categoria',
                'TOTAL MENSAL': 'Total Mensal'
            }
            df.rename(columns=rename_dict, inplace=True)
        
        df['Valor (R$)'] = pd.to_numeric(df['Valor (R$)'], errors='coerce')
        df = df.dropna(subset=['Valor (R$)'])
        
        gastos_por_mes = df.groupby('Mês')['Valor (R$)'].sum().reset_index()
        meses_validos = gastos_por_mes[gastos_por_mes['Valor (R$)'] > 0]['Mês'].tolist()
        df = df[df['Mês'].isin(meses_validos)]
        
        df['Pessoa'] = pessoa
        abas[pessoa] = df

    df_combinado = pd.concat([abas['Mohamed'], abas['Evelyn']], ignore_index=True)
    return abas, df_combinado

# Carregar os dados inicialmente
abas, df_combinado = load_data()

# Definir paletas personalizadas
paleta_mohamed = ['#ff4040', '#e63939', '#cc3333', '#b32d2d', '#992626', '#8b0000']  # Vermelhos com contraste
paleta_evelyn = ['#32cd32', '#2eb82e', '#29a329', '#248f24', '#1f7a1f', '#006400']  # Verdes com contraste

# Inicializar o app Dash
app = dash.Dash(__name__)

# Layout do dashboard
app.layout = html.Div(
    style={'backgroundColor': '#f9f9f9', 'padding': '20px'},
    children=[
        html.H1("Dashboard de Gastos 2025 - Comparativo", style={'textAlign': 'center', 'color': '#333'}),
        
        html.Button("Atualizar Dados", id="btn_atualizar", style={'marginBottom': '20px'}),
        
        html.Label("Selecione a Pessoa:"),
        dcc.Dropdown(
            id='pessoa-dropdown',
            options=[
                {'label': 'Mohamed', 'value': 'Mohamed'},
                {'label': 'Evelyn', 'value': 'Evelyn'},
                {'label': 'Comparativo', 'value': 'Comparativo'}
            ],
            value='Mohamed',
            multi=False,
            style={'width': '50%', 'marginBottom': '20px'}
        ),
        
        html.Label("Selecione o Mês:"),
        dcc.Dropdown(
            id='mes-dropdown',
            multi=False,
            style={'width': '50%', 'marginBottom': '20px'}
        ),
        
        html.Label("Selecione a Categoria:"),
        dcc.Dropdown(
            id='categoria-dropdown',
            multi=False,
            style={'width': '50%', 'marginBottom': '20px'}
        ),
        
        dcc.Graph(id='grafico-pizza', config={'displayModeBar': True, 'toImageButtonOptions': {'format': 'png', 'filename': 'grafico_pizza'}}),
        dcc.Graph(id='grafico-barras', config={'displayModeBar': True, 'toImageButtonOptions': {'format': 'png', 'filename': 'grafico_barras'}}),
        dcc.Graph(id='grafico-linha', config={'displayModeBar': True, 'toImageButtonOptions': {'format': 'png', 'filename': 'grafico_linha'}}),
        dcc.Graph(id='grafico-total-meses', config={'displayModeBar': True, 'toImageButtonOptions': {'format': 'png', 'filename': 'grafico_total_meses'}}),
        
        html.Label("Selecione a Despesa para Comparativo (Opcional):"),
        dcc.Dropdown(
            id='despesa-dropdown',
            multi=False,
            style={'width': '50%', 'marginBottom': '20px'}
        ),
        
        dcc.Graph(id='grafico-comparativo', config={'displayModeBar': True, 'toImageButtonOptions': {'format': 'png', 'filename': 'grafico_comparativo'}}),
        
        html.H3("Detalhamento dos Gastos", style={'color': '#333'}),
        dash_table.DataTable(
            id='tabela-dados',
            columns=[{'name': col, 'id': col} for col in ['Pessoa', 'Mês', 'Categoria', 'Despesas', 'Valor (R$)']],
            style_table={'overflowX': 'auto'},
            style_cell={'textAlign': 'left', 'padding': '5px'},
            style_header={'backgroundColor': '#ddd', 'fontWeight': 'bold'}
        ),
        
        html.Div([
            html.Button("Exportar Tabela como CSV", id="btn_csv"),
            dcc.Download(id="download-dataframe-csv")
        ], style={'marginTop': '20px'}),
        
        # Armazenar os dados em um componente oculto
        dcc.Store(id='data-store', data={'abas': {k: v.to_dict() for k, v in abas.items()}, 'df_combinado': df_combinado.to_dict()})
    ]
)

# Callback para atualizar os dados ao clicar no botão
@app.callback(
    Output('data-store', 'data'),
    Input('btn_atualizar', 'n_clicks'),
    prevent_initial_call=False
)
def atualizar_dados(n_clicks):
    global abas, df_combinado
    abas, df_combinado = load_data()
    return {'abas': {k: v.to_dict() for k, v in abas.items()}, 'df_combinado': df_combinado.to_dict()}

# Callback para atualizar os dropdowns
@app.callback(
    [Output('mes-dropdown', 'options'),
     Output('mes-dropdown', 'value'),
     Output('categoria-dropdown', 'options'),
     Output('categoria-dropdown', 'value'),
     Output('despesa-dropdown', 'options'),
     Output('despesa-dropdown', 'value')],
    [Input('pessoa-dropdown', 'value'),
     Input('data-store', 'data')]
)
def update_dropdowns(selected_pessoa, data_store):
    # Restaurar os DataFrames a partir do data_store
    abas_local = {k: pd.DataFrame(v) for k, v in data_store['abas'].items()}
    df_combinado_local = pd.DataFrame(data_store['df_combinado'])
    
    if selected_pessoa == 'Comparativo':
        df_pessoa = df_combinado_local
    else:
        df_pessoa = abas_local[selected_pessoa]
    
    meses = [{'label': mes, 'value': mes} for mes in df_pessoa['Mês'].unique()]
    categorias = [{'label': cat, 'value': cat} for cat in df_pessoa['Categoria'].unique()] + [{'label': 'Todas', 'value': 'Todas'}]
    despesas = [{'label': desp, 'value': desp} for desp in df_pessoa['Despesas'].unique()]
    
    return meses, meses[0]['value'] if meses else None, categorias, 'Todas', despesas, None

# Callback para atualizar os gráficos e a tabela
@app.callback(
    [Output('grafico-pizza', 'figure'),
     Output('grafico-barras', 'figure'),
     Output('grafico-linha', 'figure'),
     Output('grafico-total-meses', 'figure'),
     Output('grafico-comparativo', 'figure'),
     Output('tabela-dados', 'data')],
    [Input('pessoa-dropdown', 'value'),
     Input('mes-dropdown', 'value'),
     Input('categoria-dropdown', 'value'),
     Input('despesa-dropdown', 'value'),
     Input('data-store', 'data')]
)
def update_dashboard(selected_pessoa, selected_month, selected_category, selected_despesa, data_store):
    # Restaurar os DataFrames a partir do data_store
    abas_local = {k: pd.DataFrame(v) for k, v in data_store['abas'].items()}
    df_combinado_local = pd.DataFrame(data_store['df_combinado'])
    
    if selected_pessoa == 'Comparativo':
        df_pessoa = df_combinado_local
        cores = {'Mohamed': '#d62728', 'Evelyn': '#2ca02c'}  # Vermelho e verde vibrantes
    else:
        df_pessoa = abas_local[selected_pessoa]
        cores = paleta_mohamed if selected_pessoa == 'Mohamed' else paleta_evelyn
    
    filtered_df = df_pessoa[df_pessoa['Mês'] == selected_month] if selected_month else df_pessoa
    if selected_category != 'Todas':
        filtered_df = filtered_df[filtered_df['Categoria'] == selected_category]
    
    if selected_pessoa == 'Comparativo':
        fig_pizza = px.pie(filtered_df, 
                           values='Valor (R$)', 
                           names='Despesas', 
                           title=f'Distribuição dos Gastos em {selected_month} (Comparativo)',
                           color='Pessoa',
                           color_discrete_map=cores)
    else:
        fig_pizza = px.pie(filtered_df, 
                           values='Valor (R$)', 
                           names='Despesas', 
                           title=f'Distribuição dos Gastos em {selected_month} ({selected_pessoa})',
                           color_discrete_sequence=cores)
    
    if selected_pessoa == 'Comparativo':
        fig_barras = px.bar(filtered_df, 
                            x='Categoria', 
                            y='Valor (R$)', 
                            color='Pessoa', 
                            barmode='group',
                            title=f'Gastos por Categoria em {selected_month} (Comparativo)',
                            color_discrete_map=cores)
    else:
        fig_barras = px.bar(filtered_df, 
                            x='Categoria', 
                            y='Valor (R$)', 
                            color='Despesas', 
                            title=f'Gastos por Categoria em {selected_month} ({selected_pessoa})',
                            barmode='stack',
                            color_discrete_sequence=cores)
    
    if selected_pessoa == 'Comparativo':
        gastos_totais = df_pessoa.groupby(['Mês', 'Pessoa'])['Valor (R$)'].sum().reset_index()
        fig_linha = px.line(gastos_totais, 
                            x='Mês', 
                            y='Valor (R$)', 
                            color='Pessoa',
                            title='Tendência dos Gastos Totais em 2025 (Comparativo)',
                            markers=True,
                            color_discrete_map=cores)
    else:
        gastos_totais = df_pessoa.groupby('Mês')['Valor (R$)'].sum().reset_index()
        fig_linha = px.line(gastos_totais, 
                            x='Mês', 
                            y='Valor (R$)', 
                            title=f'Tendência dos Gastos Totais em 2025 ({selected_pessoa})',
                            markers=True,
                            color_discrete_sequence=cores)
    fig_linha.update_layout(xaxis={'categoryorder': 'array', 'categoryarray': gastos_totais['Mês'].tolist()})
    
    if selected_pessoa == 'Comparativo':
        fig_total_meses = px.bar(gastos_totais, 
                                 x='Mês', 
                                 y='Valor (R$)', 
                                 color='Pessoa',
                                 barmode='group',
                                 title='Comparativo de Gastos Totais por Mês (Comparativo)',
                                 text=gastos_totais['Valor (R$)'].apply(lambda x: f'R$ {x:.2f}'),
                                 color_discrete_map=cores)
    else:
        fig_total_meses = px.bar(gastos_totais, 
                                 x='Mês', 
                                 y='Valor (R$)', 
                                 title=f'Comparativo de Gastos Totais por Mês ({selected_pessoa})',
                                 color='Mês',
                                 text=gastos_totais['Valor (R$)'].apply(lambda x: f'R$ {x:.2f}'),
                                 color_discrete_sequence=cores)
    fig_total_meses.update_traces(textposition='auto')
    fig_total_meses.update_layout(xaxis={'categoryorder': 'array', 'categoryarray': gastos_totais['Mês'].tolist()})
    
    if selected_despesa:
        despesa_df = df_pessoa[df_pessoa['Despesas'] == selected_despesa]
        if selected_pessoa == 'Comparativo':
            fig_comparativo = px.bar(despesa_df, 
                                     x='Mês', 
                                     y='Valor (R$)', 
                                     color='Pessoa',
                                     barmode='group',
                                     title=f'Comparativo de "{selected_despesa}" em Todos os Meses (Comparativo)',
                                     text=despesa_df['Valor (R$)'].apply(lambda x: f'R$ {x:.2f}'),
                                     color_discrete_map=cores)
        else:
            fig_comparativo = px.bar(despesa_df, 
                                     x='Mês', 
                                     y='Valor (R$)', 
                                     title=f'Comparativo de "{selected_despesa}" em Todos os Meses ({selected_pessoa})',
                                     color='Mês',
                                     text=despesa_df['Valor (R$)'].apply(lambda x: f'R$ {x:.2f}'),
                                     color_discrete_sequence=cores)
        fig_comparativo.update_traces(textposition='auto')
        fig_comparativo.update_layout(xaxis={'categoryorder': 'array', 'categoryarray': despesa_df['Mês'].tolist()})
    else:
        fig_comparativo = px.bar()
        fig_comparativo.update_layout(
            title="Comparativo (Selecione uma despesa para visualizar)",
            showlegend=False,
            xaxis={'visible': False},
            yaxis={'visible': False},
            annotations=[dict(text="Nenhuma despesa selecionada", showarrow=False, font={'size': 16})]
        )
    
    tabela_dados = filtered_df[['Pessoa', 'Mês', 'Categoria', 'Despesas', 'Valor (R$)']].to_dict('records')
    
    return fig_pizza, fig_barras, fig_linha, fig_total_meses, fig_comparativo, tabela_dados

# Callback para exportar a tabela como CSV
@app.callback(
    Output("download-dataframe-csv", "data"),
    [Input("btn_csv", "n_clicks"),
     Input("tabela-dados", "data")],
    prevent_initial_call=True,
)
def download_csv(n_clicks, tabela_dados):
    if n_clicks is None:
        return None
    
    df_to_export = pd.DataFrame(tabela_dados)
    csv_string = df_to_export.to_csv(index=False, encoding='utf-8')
    return dcc.send_data_frame(df_to_export.to_csv, "gastos_exportados.csv", index=False)

# Rodar o app
if __name__ == '__main__':
    app.run_server(host='0.0.0.0', port=8050, debug=False)