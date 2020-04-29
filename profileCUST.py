

#This program runs on Python 3.6.5
#Please import the following modules for the program to use

#import math
#import os
import pyodbc
import datetime as dt
import pandas as pd
import flask
import plotly.graph_objs as go
import dash
import dash_html_components as html
import dash_core_components as dcc
from dash.dependencies import Input, Output, State
from dash.exceptions import PreventUpdate

# Connect to the dB

con = pyodbc.connect(r'DRIVER={ODBC Driver 17 for SQL Server};'
                     r'SERVER=mdldbbidev01.devxpo.pvt; '
                     r'DATABASE=RampForecasting;'
                     r'UID=IMBU;'
                     r'PWD=tPQC0CrD9e1PcXQpqaVJ;')

#con = pyodbc.connect(r'DRIVER={ODBC Driver 13 for SQL Server};'
#                     r'SERVER=mdldbbidev01.devxpo.pvt; PORT=2301;'
#                     r'DATABASE=RampForecasting;'
#                     r'trusted_connection=yes;'
#                     r'MARS_Connection=yes;'
#                     r'UID=me;'
#                     r'PWD=pass;')

cursor = con.cursor()

#------------------------------------------------------------------------------------------------------------------------------------------------
# 	READING IN THE DATA

df_customers = pd.read_sql(sql = "select distinct [CustomerCode], [CustomerName] from [dbo].[ipHistoryNForecast]", con=con);
df_customers = df_customers.astype({"CustomerCode": str, 'CustomerName': str});
#print(df_customers.head())

CustomerCodes = df_customers["CustomerCode"].tolist(); CustomerCodes.sort(); #print(CustomerCodes)
CustomerNames = df_customers["CustomerName"].tolist(); CustomerNames.sort(); #print(CustomerNames)

cursor.close()
del cursor

#------------------------------------------------------------------------------------------------------------------------------------------------
customerprofile_layout = html.Div([
#Input Block : FIRSTROW
    html.Div([
#Column 0
        html.Div([
            html.Div([html.H6("CUSTOMER", style={'color': 'blue'})
            ]),
            html.Div([
	    		dcc.Dropdown(
                    id='CustomerName3',
                    options = [{'label': i, 'value': i} for i in CustomerNames],
					value='TOYOTA MOTOR (TMMBC)-331156'
                )
            ], style={'width': '40%', 'display': 'inline-block'}),
             html.Div([html.Div(id='CustomerCode3' )
            ],style={'width': '20%', 'display': 'inline-block', 'verticalAlign': 'center', 'textAlign': 'center'}),
            html.Div([
                    html.Div(html.Br()),
                    html.Div(html.Button('Submit', id='cust_profile_button')),
                    html.Div(html.Br())
                ], style={'width': '20%', 'display': 'inline-block', 'verticallAlign': 'top'}),
        ], style={'width': '100%', 'display': 'inline-block', 'verticalAlign': 'top'}),
    ]),

#GRAPH : SECONDROW
    html.Div([
	    html.Div([
            html.H4('Shipments By Lane', style={'color': 'blue'}),
            html.A(html.Button('Export to Excel'), id='data_customerprofileship', href=f'/export/excel/customerprofileship'),
            dcc.Graph(id = 'customerprofileship')
        ],style={'width': '50%', 'display': 'inline-block'}),
        html.Div([
            html.H4('Forecast By Lane', style={'color': 'blue'}),
            html.A(html.Button('Export to Excel'), id='data_customerprofilefcst', href=f'/export/excel/customerprofilefcst'),
            dcc.Graph(id = 'customerprofilefcst')
        ],style={'width': '50%', 'display': 'inline-block'})
    ],style={'width': '100%', 'display': 'inline-block'}),
])	
#------------------------------------------------------------------------------------------------------------------------------------------------

