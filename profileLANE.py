

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

df_flows = pd.read_sql(sql = "SELECT distinct [CustomerName], [FromMarket], [ToMarket] FROM [dbo].[ipHistoryNForecast] where ([ACTdepLD] + [ACTdepFR]) > 0 ", con = con)
df_flows = df_flows.astype({"CustomerName": str, "FromMarket": str, "ToMarket": str});
df_flows["Lane"] = df_flows.apply( lambda row: row.FromMarket + '-' + row.ToMarket, axis = 1);
#print(df_flows.head());

Lanes = df_flows["Lane"].unique().tolist();

cursor.close()
del cursor

#------------------------------------------------------------------------------------------------------------------------------------------------
laneprofile_layout = html.Div([
#Input Block : FIRSTROW
    html.Div([
#Column 0
        html.Div([
            html.Div([html.H6("LANE", style={'color': 'blue'})
            ]),
            html.Div([
	    		dcc.Dropdown(
                    id='Lane',
                    options = [{'label': i, 'value': i} for i in Lanes],
					value='CHI-LAX'
                )
            ], style={'width': '40%', 'display': 'inline-block'}),
            html.Div([
                    html.Div(html.Br()),
                    html.Div(html.Button('Submit', id='lane_profile_button')),
                    html.Div(html.Br())
                ], style={'width': '20%', 'display': 'inline-block', 'verticallAlign': 'bottom', 'horizontalAlign': 'center'}),
            html.Div([
                html.A(html.Button('Export to Excel'), id='data_laneprofile', href=f'/export/excel/laneprofile'),
            ],style={'width': '20%', 'display': 'inline-block', 'verticalAlign': 'bottom', 'textAlign': 'right'}),
        ], style={'width': '100%', 'display': 'inline-block', 'verticalAlign': 'top'}),
    ]),
#GRAPH : THIRDDROW
	html.Div([dcc.Graph(id = 'laneprofiletbl', style = {'height': '2000', 'width': '2000', 'display': 'block'})]),
])	
#------------------------------------------------------------------------------------------------------------------------------------------------

