


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

df_markets = pd.read_sql(sql = "SELECT [Market], [MarketCity], [MarketState] from [dbo].[ipMarketMaster]", con=con);
df_markets = df_markets.astype({"Market": str, "MarketCity": str, "MarketState": str});
#print(df_markets)
Markets = df_markets["Market"].tolist(); Markets.sort(); #print(Markets);  

df_horizon = pd.read_sql(sql = "select distinct [RampDate] from [dbo].[ipHistoryNForecast]", con=con);
df_horizon = df_horizon.astype({"RampDate": str}); df_horizon["RampDate"] = pd.to_datetime(df_horizon["RampDate"]);  

earliest = df_horizon["RampDate"].min(); #print(earliest);
latest = df_horizon["RampDate"].max(); #print(latest);

cursor.close()
del cursor

#------------------------------------------------------------------------------------------------------------------------------------------------
lane_layout = html.Div([
#Input Block : FIRSTROW
	html.Div([
#Column 31
        html.Div([
            html.Div([
                html.H6("ORIGIN", style={'color': 'blue', 'width': '50%', 'display': 'inline-block', 'verticalAlign': 'top'}),
                html.H6('DESTINATION', style={'color': 'blue', 'width': '50%', 'display': 'inline-block', 'verticalAlign': 'top'})
            ], style={'width': '100%', 'display': 'inline-block', 'verticalAlign': 'top'}),
            html.Div([
	    		dcc.Dropdown(
                    id='FromMarket',
                    options = [{'label': i, 'value': i} for i in Markets],
					value='CHI',
                    style={'width': '20%', 'display': 'inline-block'}
                ),
                html.Div(id='FromCity', style={'width': '15%', 'display': 'inline-block', 'verticalAlign': 'bottom', 'textAlign': 'left'}),
                html.Div(id='FromState', style={'width': '5%', 'display': 'inline-block', 'verticalAlign': 'bottom', 'textAlign': 'left'}),
                dcc.Dropdown(
                    id='ToMarket',
                    options = [{'label': i, 'value': i} for i in Markets],
					value='LAX',
                    style={'width': '20%', 'display': 'inline-block'}
                ),
                html.Div(id='ToCity', style={'width': '15%', 'display': 'inline-block', 'verticalAlign': 'bottom', 'textAlign': 'left'}),
                html.Div(id='ToState', style={'width': '5%', 'display': 'inline-block', 'verticalAlign': 'bottom', 'textAlign': 'left'}),
                html.Div([
                    html.Div(html.Br()),
                    html.Div(html.Button('Submit Lane', id='pub_lane_button')),
                    html.Div(html.Br())
                ], style={'width': '10%', 'display': 'inline-block', 'verticallAlign': 'bottom'}),
                html.Div([
                    html.A(html.Button('Export to Excel'), id='data_bylane', href=f'/export/excel/bylane')
                ],style={'width': '20%', 'display': 'inline-block', 'verticalAlign': 'bottom', 'textAlign': 'right'}),
            ], style={'width': '100%', 'display': 'inline-block', 'verticalAlign': 'top'})
        ]),
    ]),
#GRAPH : SECONDROW
	html.Div([dcc.Graph(id='publane')]),
#UPDATE : THIRDROW
    html.Div([
        html.H4('To override the forecst please enter week of, an override volume, your name and the reason', style={'color':'red'}),
        html.Label([
            html.H4('WeekBeginning', style={'color':'green'}),
            dcc.DatePickerSingle(
                id = 'ovr_lane_week',
                min_date_allowed = dt.datetime.now().date() + dt.timedelta(1),
                max_date_allowed = latest, 
                initial_visible_month = latest,
                date = latest
            )
		], style={'width': '10%', 'display': 'inline-block', 'verticalAlign': 'bottom'}),
        html.Label([
            html.H4('Over Ride Volume', style={'color':'green'}),
            dcc.Input(id='ovr_lane_volume', min='0', max='5000', step='1', value='0', type='number')
        ], style={'width': '10%', 'display': 'inline-block', 'verticalAlign': 'bottom'}),
        html.Label([
            html.H4('Proposed By', style={'color':'green'}),
            dcc.Input(id = 'ovr_lane_prop_by', value = '', type = 'str')
        ],style={'width': '20%', 'display': 'inline-block', 'verticalAlign': 'bottom'}),
        html.Label([
            html.H4('Reason for Over Ride', style={'color':'green'}),
            dcc.Input(id = 'ovr_lane_reason', value = '', type = 'str')
        ],style={'width': '40%', 'display': 'inline-block', 'verticalAlign': 'bottom'}),
        html.Div([
            html.Div(html.Br()),
            html.Div(html.Button('Submit Over Ride', id='ovr_lane_button')),
            html.Div(html.Br())
        ], style={'width': '20%', 'display': 'inline-block', 'verticalAlign': 'bottom'}),
    ],style={'width': '100%', 'display': 'inline-block'}),
])	
#------------------------------------------------------------------------------------------------------------------------------------------------

