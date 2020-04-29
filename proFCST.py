


#This program runs on Python 3.6.5
#Please import the following modules for the program to use

import math
import os
import io
import datetime as dt
import pandas as pd
import pyodbc
import flask
import plotly.graph_objs as go
import plotly.express as px
import dash
import dash_html_components as html
import dash_core_components as dcc
from dash.dependencies import Input, Output, State
from dash.exceptions import PreventUpdate

import pubCUST
import pubLANE
import pubCUSTLANE
import profileCUST
import profileLANE

#Connect to the dB

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

#Initialize the app

external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']
app = dash.Dash(__name__, external_stylesheets=external_stylesheets)
server=app.server

#Define favicon.ico - if the data required to populate the page takes too long to load, flask keeps the page alive by displaying favicon.ico

@server.route('/favicon.ico')
def favicon():
    return flask.send_from_directory(os.path.join(server.root_path, 'static'),'favicon.ico') 

#External Dash Library 
#app.css.append_css({"external_url": "https://codepen.io/chriddyp/pen/bWLwgP.css"})

#====================================================================================================================================================

df_markets = pd.read_sql(sql = "SELECT [Market], [MarketCity], [MarketState] from [dbo].[ipMarketMaster]", con=con);
df_markets = df_markets.astype({"Market": str, "MarketCity": str, "MarketState": str});
#print(df_markets)

Markets = df_markets["Market"].tolist(); Markets.sort(); #print(Markets);  

df_flows = pd.read_sql(sql = "SELECT [RampDate], [FromMarket], [ToMarket], [CustomerCode], [CustomerName], [ACTdepLD], [ACTdepFR], [ACTdepMT], [StatFcst], [ConsensusFcst] FROM [dbo].[ipHistoryNForecast] where ([ACTdepLD] + [ACTdepFR] + [ACTdepMT] + [StatFcst] + [ConsensusFcst]) > 0 order by  [RampDate], [FromMarket], [ToMarket]", con = con)
df_flows = df_flows.astype({"RampDate": str, "FromMarket": str, "ToMarket": str, "CustomerCode": str, "CustomerName": str, "ACTdepLD": int, "ACTdepFR": int, "ACTdepMT": int, "StatFcst": int, "ConsensusFcst": int});
df_flows["RampDate"] = pd.to_datetime(df_flows["RampDate"]);  
#print(df_flows);

earliest = df_flows["RampDate"].min(); #print(earliest);
latest = df_flows["RampDate"].max(); #print(latest);

df_flows["Lane"] = df_flows.apply( lambda row: row.FromMarket + '-' + row.ToMarket, axis = 1);
df_flows["Load"] = df_flows.apply( lambda row: row.ACTdepLD + row.ACTdepFR, axis = 1);

Lanes = df_flows["Lane"].unique().tolist();

df_customers = pd.read_sql(sql = "select distinct [CustomerCode], [CustomerName] from [dbo].[ipHistoryNForecast]", con=con);
df_customers = df_customers.astype({"CustomerCode": str, 'CustomerName': str});
#print(df_customers.head())

CustomerCodes = df_flows["CustomerCode"].unique().tolist(); CustomerCodes.sort(); #print(CustomerCodes)
CustomerNames = df_flows["CustomerName"].unique().tolist(); CustomerNames.sort(); #print(CustomerNames)

#=================================================================================================================================================

app = dash.Dash()

app.config['suppress_callback_exceptions'] = True

app.layout = html.Div([
    html.H1('RAMP FORECASTING'),
    dcc.DatePickerRange(
        id='interval',
        min_date_allowed=earliest,
        max_date_allowed=latest,
        initial_visible_month=earliest,
        start_date = earliest,
        end_date=latest
    ),
    dcc.Tabs(id="byfcst", value='bycust', children=[
        dcc.Tab(label='Forecast by Customer', value='bycust'),
        dcc.Tab(label='Forecast by Lane', value='bylane'),
        dcc.Tab(label='Forecast by Customer by Lane', value='bycustlane'),
        dcc.Tab(label='Customer Profile', value='custprofile'),
        dcc.Tab(label='LaneProfile', value='laneprofile')
    ]),
    html.Div(id='rampforecastby')
])
    
@app.callback(
    Output('rampforecastby', 'children'),
    [Input('byfcst', 'value')]
)
def render_content(byfcst) :
    if byfcst == 'bycust':
        return pubCUST.cust_layout
    elif byfcst == 'bylane':
        return pubLANE.lane_layout
    elif byfcst == 'bycustlane':
        return pubCUSTLANE.custlane_layout
    elif byfcst == 'custprofile' :
        return profileCUST.customerprofile_layout
    elif byfcst == 'laneprofile' :
        return profileLANE.laneprofile_layout

#============================================================================================================================================

#update CustomerCode1
@app.callback(
    Output('CustomerCode1', 'value'),
    [Input('select-all-1', 'value')],
    [State('CustomerCode1', 'options')]
)
def test(selectALL, options) :
    if selectALL == 'SA' :
        return[]
    else :
        raise PreventUpdate()
  		

#------------------------------------------------------------------------------------------------------------------------------------------------------

#The pubcustomer graph
@app.callback(dash.dependencies.Output('pubcustomer', 'figure'),
              [dash.dependencies.Input('pub_cust_button', 'n_clicks')],
              [dash.dependencies.State('CustomerCode1', 'value'),
			   dash.dependencies.State('select-all-1', 'value'),
               dash.dependencies.State('interval', 'start_date'),
               dash.dependencies.State('interval', 'end_date')]
)
def update_bycustomer(n_clicks, CustomerCode1, selectALL, start_date, end_date) :
   
    #print(start_date); print(end_date);

    df_bycustomer = df_flows[(df_flows['RampDate'] >= start_date) & (df_flows['RampDate'] <= end_date)] ;
    if selectALL == 'SC' :
        df_bycustomer = df_bycustomer[df_bycustomer['CustomerCode'].isin(CustomerCode1)]; 
            
    df_bycustomer = df_bycustomer.groupby(['RampDate'])[['ACTdepLD','ACTdepFR','ACTdepMT','StatFcst', 'ConsensusFcst']].sum().reset_index();
        #print(df_bycustomer);
    
    filename = '/apps/IMO/Forecasting/PublishForecast/pubcustomer_graph_data.csv'
    if os.path.exists(filename) :
        os.remove(filename);
    df_bycustomer.to_csv(filename);

    horizon = df_bycustomer['RampDate'];

    xaxis=dict(
        rangeslider=dict(
            visible=True
        ),
        type="date"
    )

    # Casts the time to a datetime object as the numpy Timestamp didn't work well with the graphs
    datetime_horizon = []
    for ele in horizon:
        datetime_ele = ele.to_pydatetime()
        datetime_horizon.append(datetime_ele)
		
    # Create and style traces (data series)   [ACTdepLD], [ACTdepFR], [ACTdepMT], [StatFcst] 
    trace0 = go.Scatter(
        x=datetime_horizon,
        y=df_bycustomer['ACTdepFR'],
        name='ActualDeparturesFreeRunners',
        fill='tozeroy',
		line=dict(
		    shape='linear',
            color='pink',
            width=4)
    )
    trace1 = go.Scatter(
        x=datetime_horizon,
        y=df_bycustomer['ACTdepFR'] + df_bycustomer['ACTdepLD'],
        name='ActualDeparturesLoaded',
        fill='tonexty',
		line=dict(
		    shape='linear',
            color='lightblue',
            width=4)
    )
    trace2 = go.Bar(
        x=datetime_horizon,
        y=df_bycustomer['StatFcst'],
        name='StatisticalForecast',
        marker=dict(
		    color='lightgreen')
    )
    trace3 = go.Scatter(
        x=datetime_horizon,
        y=df_bycustomer['ConsensusFcst'],
        name='ConsensusForecast',
        fill = None,
        line=dict(
		    shape='linear',
            color='seagreen',
            width=4)
    )
    data_bycustomer = [trace0, trace1, trace2, trace3]; #print(data_bylane)
    layout_bycustomer = dict(title='History & Forecact By Customer',
                  xaxis=dict(title='Date'),
                  yaxis=dict(title='Number'),
                  )
    return {
        'data': data_bycustomer,
        'layout': layout_bycustomer
    }

#----------------------------------------------------------------------------------------------------------------------------------------------------

@app.server.route('/export/excel/bycustomer')
def export_excel_bycustomer():
    
    option_df = pd.read_csv('/apps/IMO/Forecasting/PublishForecast/pubcustomer_graph_data.csv')

    xlsx_io = io.BytesIO()
    writer = pd.ExcelWriter(xlsx_io, engine='xlsxwriter')
    option_df.to_excel(writer, sheet_name='bycustomer')
    writer.save()
    xlsx_io.seek(0)

    return flask.send_file(
        xlsx_io,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        attachment_filename=f'export_bycustomer.xlsx',
        as_attachment=True,
        cache_timeout=0
    )

#----------------------------------------------------------------------------------------------------------------------------------------------------

# callback update
@app.callback([dash.dependencies.Output('ovr_cust_week', 'date'),
               dash.dependencies.Output('ovr_cust_volume', 'value'),
               dash.dependencies.Output('ovr_cust_prop_by', 'value'),
               dash.dependencies.Output('ovr_cust_reason', 'value')],
              [dash.dependencies.Input('ovr_cust_button', 'n_clicks')],
              [dash.dependencies.State('ovr_cust_week', 'date'),
               dash.dependencies.State('CustomerNameA', 'value'),
               dash.dependencies.State('ovr_cust_volume', 'value'),
               dash.dependencies.State('ovr_cust_prop_by', 'value'),
               dash.dependencies.State('ovr_cust_reason', 'value')]
)
def update_ovr_by_cust(n_clicks, ovr_cust_week, CustomerNameA, ovr_cust_volume, ovr_cust_prop_by, ovr_cust_reason) :
    print(ovr_cust_week, ovr_cust_volume, ovr_cust_prop_by, ovr_cust_reason);
    if ((len(ovr_cust_reason) > 2)  & (len(ovr_cust_prop_by) > 2)) :
        print(ovr_cust_reason, ovr_cust_prop_by);
        con.execute("insert into [dbo].[WeeklyFcstOverRideByCustomer] ([WeekOf], [CustomerName], [OverRideVolume], [ProposedBy], [Comments], [ProposedOn], [ConsensusFcst], [Approved], [ApprovedOn]) values(?,?,?,?,?,?,?,?,?)",ovr_cust_week, CustomerNameA, ovr_cust_volume, ovr_cust_prop_by, ovr_cust_reason,  dt.datetime.now(), 0, 0, latest);
        con.commit();
    return (latest, 0, '', '')


#================================================================================================================================================

#update FromCity
@app.callback(
    dash.dependencies.Output('FromCity', 'children'),
	[dash.dependencies.Input('FromMarket', 'value')]
)
def update_FromCity(FromMarket) :
    FromCity = df_markets[df_markets.Market == FromMarket].MarketCity.values[0]
    return(FromCity.upper())

#update FromState
@app.callback(
    dash.dependencies.Output('FromState', 'children'),
   	[dash.dependencies.Input('FromMarket', 'value')]
)
def update_FromState(FromMarket) :
    FromState = df_markets[df_markets.Market == FromMarket].MarketState.values[0]; 
    return(FromState.upper())

#update ToCity
@app.callback(
    dash.dependencies.Output('ToCity', 'children'),
	[dash.dependencies.Input('ToMarket', 'value')]
)
def update_ToCity(ToMarket) :
    ToCity = df_markets[df_markets.Market == ToMarket].MarketCity.values[0]
    return(ToCity.upper())

#update ToState
@app.callback(
    dash.dependencies.Output('ToState', 'children'),
   	[dash.dependencies.Input('ToMarket', 'value')]
)
def update_ToState(ToMarket) :
    ToState = df_markets[df_markets.Market == ToMarket].MarketState.values[0]; 
    return(ToState.upper())

#------------------------------------------------------------------------------------------------------------------------------------------------------

#The bylane graph
@app.callback(dash.dependencies.Output('publane', 'figure'),
              [dash.dependencies.Input('pub_lane_button', 'n_clicks')],
              [dash.dependencies.State('FromMarket', 'value'),
               dash.dependencies.State('ToMarket', 'value'),
               dash.dependencies.State('interval', 'start_date'),
               dash.dependencies.State('interval', 'end_date')]
)
def update_bylane(n_clicks, FromMarket, ToMarket, start_date, end_date) :

    #print(start_date); print(end_date);

    df_bylane = df_flows[(df_flows['FromMarket'] == FromMarket) & (df_flows['ToMarket'] == ToMarket) & (df_flows['RampDate'] >= start_date) & (df_flows['RampDate'] <= end_date)];
    
    df_bylane =  df_bylane.groupby(['RampDate'])[['ACTdepLD','ACTdepFR','ACTdepMT','StatFcst', 'ConsensusFcst']].sum().reset_index(); #print(df_bylane.head())
    
    filename = '/apps/IMO/Forecasting/PublishForecast/publane_graph_data.csv'
    df_bylane.to_csv(filename);
	
    horizon = df_bylane['RampDate'];

    xaxis=dict(
        rangeslider=dict(
            visible=True
        ),
        type="date"
    )

    # Casts the time to a datetime object as the numpy Timestamp didn't work well with the graphs
    datetime_horizon = []
    for ele in horizon:
        datetime_ele = ele.to_pydatetime()
        datetime_horizon.append(datetime_ele)
		
    # Create and style traces (data series)   [ACTdepLD], [ACTdepFR], [ACTdepMT], [StatFcst] 
    trace0 = go.Scatter(
        x=datetime_horizon,
        y=df_bylane['ACTdepFR'],
        name='ActualDeparturesFreeRunners',
        fill='tozeroy',
		line=dict(
		    shape='linear',
            color='pink',
            width=4)
    )
    trace1 = go.Scatter(
        x=datetime_horizon,
        y=df_bylane['ACTdepLD'],
        name='ActualDeparturesLoaded',
        fill='tonexty',
		line=dict(
		    shape='linear',
            color='lightblue',
            width=4)
    )
    trace2 = go.Bar(
        x=datetime_horizon,
        y=df_bylane['StatFcst'],
        name='StatisticalForecast',
        marker=dict(
		    color='lightgreen')
    )
    trace3 = go.Scatter(
        x=datetime_horizon,
        y=df_bylane['ConsensusFcst'],
        name='ConsensusForecast',
        fill = None,
        line=dict(
		    shape='linear',
            color='seagreen',
            width=4)
    )
    data_bylane = [trace0, trace1, trace2, trace3]; #print(data_bylane) 
    layout_bylane = dict(title='History & Forecact By Lane',
                  xaxis=dict(title='Date'),
                  yaxis=dict(title='Number'),
                  )
    return {
        'data': data_bylane,
        'layout': layout_bylane
    }

#----------------------------------------------------------------------------------------------------------------------------------------------------

@app.server.route('/export/excel/bylane')
def export_excel_bylane():
    
    option_df = pd.read_csv('/apps/IMO/Forecasting/PublishForecast/publane_graph_data.csv')

    xlsx_io = io.BytesIO()
    writer = pd.ExcelWriter(xlsx_io, engine='xlsxwriter')
    option_df.to_excel(writer, sheet_name='bylane')
    writer.save()
    xlsx_io.seek(0)

    return flask.send_file(
        xlsx_io,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        attachment_filename=f'export_bylane.xlsx',
        as_attachment=True,
        cache_timeout=0
    )

#----------------------------------------------------------------------------------------------------------------------------------------------------

# callback update
@app.callback([dash.dependencies.Output('ovr_lane_week', 'date'),
               dash.dependencies.Output('ovr_lane_volume', 'value'),
               dash.dependencies.Output('ovr_lane_prop_by', 'value'),
               dash.dependencies.Output('ovr_lane_reason', 'value')],
              [dash.dependencies.Input('ovr_lane_button', 'n_clicks')],
              [dash.dependencies.State('ovr_lane_week', 'date'),
               dash.dependencies.State('FromMarket', 'value'),
               dash.dependencies.State('ToMarket', 'value'),
               dash.dependencies.State('ovr_lane_volume', 'value'),
               dash.dependencies.State('ovr_lane_prop_by', 'value'),
               dash.dependencies.State('ovr_lane_reason', 'value')]
)
def update_ovr_by__lane(n_clicks, ovr_lane_week, FromMarket, ToMarket, ovr_lane_volume, ovr_lane_prop_by, ovr_lane_reason) :
    print(ovr_lane_week, ovr_lane_volume, ovr_lane_prop_by, ovr_lane_reason);
    if ((len(ovr_lane_reason) > 2)  & (len(ovr_lane_prop_by) > 2)) :
        print(ovr_lane_reason, ovr_lane_prop_by);
        con.execute("insert into [dbo].[WeeklyFcstOverRideByLane] ([WeekOf], [FromMarket], [ToMarket], [OverRideVolume], [ProposedBy], [Comments], [ProposedOn], [ConsensusFcst], [Approved], [ApprovedOn]) values(?,?,?,?,?,?,?,?,?,?)",ovr_lane_week, FromMarket, ToMarket, ovr_lane_volume, ovr_lane_prop_by, ovr_lane_reason,  dt.datetime.now(), 0, 0, latest);
        con.commit();
    return (latest, 0, '', '')

#===============================================================================================================================================================

#update CustomerCode2
@app.callback(
    dash.dependencies.Output('CustomerCode2', 'children'),
    [dash.dependencies.Input('CustomerName2', 'value')]
)
def update_CustomerCode2(CustomerName2) :
    CustomerCode2 = df_customers[df_customers.CustomerName == CustomerName2].CustomerCode.values[0]
    return(CustomerCode2.upper())
  		
#-----------------------------------------------------------------------------------------------------------------------------------------------

#update FromCity2
@app.callback(
    dash.dependencies.Output('FromCity2', 'children'),
	[dash.dependencies.Input('FromMarket2', 'value')]
)
def update_FromCity2(FromMarket2) :
    FromCity2 = df_markets[df_markets.Market == FromMarket2].MarketCity.values[0]
    return(FromCity2.upper())

#update FromState2
@app.callback(
    dash.dependencies.Output('FromState2', 'children'),
   	[dash.dependencies.Input('FromMarket2', 'value')]
)
def update_FromState2(FromMarket2) :
    FromState2 = df_markets[df_markets.Market == FromMarket2].MarketState.values[0]; 
    return(FromState2.upper())

#update ToCity2
@app.callback(
    dash.dependencies.Output('ToCity2', 'children'),
	[dash.dependencies.Input('ToMarket2', 'value')]
)
def update_ToCity2(ToMarket2) :
    ToCity2 = df_markets[df_markets.Market == ToMarket2].MarketCity.values[0]
    return(ToCity2.upper())

#update ToState2
@app.callback(
    dash.dependencies.Output('ToState2', 'children'),
   	[dash.dependencies.Input('ToMarket2', 'value')]
)
def update_ToState2(ToMarket2) :
    ToState2 = df_markets[df_markets.Market == ToMarket2].MarketState.values[0]; 
    return(ToState2.upper())

#------------------------------------------------------------------------------------------------------------------------------------------------------

#The pubcustomerlane graph
@app.callback(dash.dependencies.Output('pubcustomerlane', 'figure'),
              [dash.dependencies.Input('pub_cust_lane_button', 'n_clicks')],
              [dash.dependencies.State('CustomerName2', 'value'),
               dash.dependencies.State('FromMarket2', 'value'),
               dash.dependencies.State('ToMarket2', 'value'),
               dash.dependencies.State('interval', 'start_date'),
               dash.dependencies.State('interval', 'end_date')]
)
def update_bycustlane(n_clicks, CustomerName2, FromMarket2, ToMarket2, start_date, end_date) :

    #print(start_date); print(end_date);

    df_bycustlane = df_flows[(df_flows["CustomerName"] == CustomerName2) & (df_flows['FromMarket'] == FromMarket2) & (df_flows['ToMarket'] == ToMarket2) & (df_flows['RampDate'] >= start_date) & (df_flows['RampDate'] <= end_date)];
    
    df_bycustlane =  df_bycustlane.groupby(['RampDate'])[['ACTdepLD','ACTdepFR','ACTdepMT','StatFcst', 'ConsensusFcst']].sum().reset_index();
    
    filename = '/apps/IMO/Forecasting/PublishForecast/pubcustlane_graph_data.csv'
    df_bycustlane.to_csv(filename);
	
    horizon = df_bycustlane['RampDate'];

    xaxis=dict(
        rangeslider=dict(
            visible=True
        ),
        type="date"
    )

    # Casts the time to a datetime object as the numpy Timestamp didn't work well with the graphs
    datetime_horizon = []
    for ele in horizon:
        datetime_ele = ele.to_pydatetime()
        datetime_horizon.append(datetime_ele)
		
    # Create and style traces (data series)   [ACTdepLD], [ACTdepFR], [ACTdepMT], [StatFcst] 
    trace0 = go.Scatter(
        x=datetime_horizon,
        y=df_bycustlane['ACTdepFR'],
        name='ActualDeparturesFreeRunners',
        fill='tozeroy',
		line=dict(
		    shape='linear',
            color='pink',
            width=4)
    )
    trace1 = go.Scatter(
        x=datetime_horizon,
        y=df_bycustlane['ACTdepLD'],
        name='ActualDeparturesLoaded',
        fill='tonexty',
		line=dict(
		    shape='linear',
            color='lightblue',
            width=4)
    )
    trace2 = go.Bar(
        x=datetime_horizon,
        y=df_bycustlane['StatFcst'],
        name='StatisticalForecast',
        marker=dict(
		    color='lightgreen')
    )
    trace3 = go.Scatter(
        x=datetime_horizon,
        y=df_bycustlane['ConsensusFcst'],
        name='ConsensusForecast',
        fill = None,
        line=dict(
		    shape='linear',
            color='seagreen',
            width=4)
    )
    data_bycustlane = [trace0, trace1, trace2, trace3]; #print(data_bylane) 
    layout_bycustlane = dict(title='History & Forecact By Customer By Lane',
                  xaxis=dict(title='Date'),
                  yaxis=dict(title='Number'),
                  )
    return {
        'data': data_bycustlane,
        'layout': layout_bycustlane
    }

#----------------------------------------------------------------------------------------------------------------------------------------------------

@app.server.route('/export/excel/bycustlane')
def export_excel_bycustlane():
    
    option_df = pd.read_csv('/apps/IMO/Forecasting/PublishForecast/pubcustlane_graph_data.csv')

    xlsx_io = io.BytesIO()
    writer = pd.ExcelWriter(xlsx_io, engine='xlsxwriter')
    option_df.to_excel(writer, sheet_name='bycustlane')
    writer.save()
    xlsx_io.seek(0)

    return flask.send_file(
        xlsx_io,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        attachment_filename=f'export_bycustlane.xlsx',
        as_attachment=True,
        cache_timeout=0
    )
		
#--------------------------------------------------------------------------------------------------------------------------------------------------

# callback update
@app.callback([dash.dependencies.Output('ovr_cust_lane_week', 'date'),
               dash.dependencies.Output('ovr_cust_lane_volume', 'value'),
               dash.dependencies.Output('ovr_cust_lane_prop_by', 'value'),
               dash.dependencies.Output('ovr_cust_lane_reason', 'value')],
              [dash.dependencies.Input('ovr_cust_lane_button', 'n_clicks')],
              [dash.dependencies.State('ovr_cust_lane_week', 'date'),
               dash.dependencies.State('CustomerName2', 'value'),
               dash.dependencies.State('FromMarket2', 'value'),
               dash.dependencies.State('ToMarket2', 'value'),
               dash.dependencies.State('ovr_cust_lane_volume', 'value'),
               dash.dependencies.State('ovr_cust_lane_prop_by', 'value'),
               dash.dependencies.State('ovr_cust_lane_reason', 'value')]
)
def update_ovr_by_cust_by_lane(n_clicks, ovr_cust_lane_week, CustomerName2, FromMarket2, ToMarket2, ovr_cust_lane_volume, ovr_cust_lane_prop_by, ovr_cust_lane_reason) :
    print(ovr_cust_lane_week, ovr_cust_lane_volume, ovr_cust_lane_prop_by, ovr_cust_lane_reason);
    if ((len(ovr_cust_lane_reason) > 2)  & (len(ovr_cust_lane_prop_by) > 2)) :
        print(ovr_cust_lane_reason, ovr_cust_lane_prop_by);
        con.execute("insert into [dbo].[WeeklyFcstOverRideByCustomerByLane] ([WeekOf], [CustomerName], [FromMarket], [ToMarket], [OverrideVolume], [ProposedBy], [Comments], [ProposedOn], [ConsensusFcst], [Approved], [ApprovedOn]) values(?,?,?,?,?,?,?,?,?,?,?)", ovr_cust_lane_week, CustomerName2, FromMarket2, ToMarket2, ovr_cust_lane_volume, ovr_cust_lane_prop_by, ovr_cust_lane_reason, dt.datetime.now(), 0, 0, latest);
        con.commit();
    return (latest, 0, '', '')

#========================================================================================================================================================================

#update CustomerCode3
@app.callback(
    dash.dependencies.Output('CustomerCode3', 'children'),
    [dash.dependencies.Input('CustomerName3', 'value')]
)
def update_CustomerCode3(CustomerName3) :
    CustomerCode3 = df_customers[df_customers.CustomerName == CustomerName3].CustomerCode.values[0]
    return(CustomerCode3.upper())

#---------------------------------------------------------------------------------------------------------------------------------------------------

#The customerprofileship graph
@app.callback(dash.dependencies.Output('customerprofileship', 'figure'),
              [dash.dependencies.Input('cust_profile_button', 'n_clicks')],
              [dash.dependencies.State('CustomerName3', 'value'),
               dash.dependencies.State('interval', 'start_date'),
               dash.dependencies.State('interval', 'end_date')]
)

def update_customerprofileship(n_clicks, CustomerName3, start_date, end_date) :

    df_customerprofileship = df_flows[(df_flows["CustomerName"] == CustomerName3) & (df_flows['RampDate'] >= start_date) & (df_flows['RampDate'] <= end_date)];
    df_customerprofileship = df_customerprofileship.groupby(['Lane'])[['Load']].sum().reset_index();
    df_customerprofileship = df_customerprofileship[df_customerprofileship['Load'] >= 1];
    #print(df_customerprofileship);

    filename = '/apps/IMO/Forecasting/PublishForecast/customerprofileship_graph_data.csv'
    df_customerprofileship.to_csv(filename);
    
    piechartship = px.pie(
        df_customerprofileship,
        names = 'Lane',
        values = 'Load'
        )

    return piechartship

#--------------------------------------------------------------------------------------------------------------------------------------------------

@app.server.route('/export/excel/customerprofileship')
def export_excel_customerprofileship():
    
    option_df = pd.read_csv('/apps/IMO/Forecasting/PublishForecast/customerprofileship_graph_data.csv')

    xlsx_io = io.BytesIO()
    writer = pd.ExcelWriter(xlsx_io, engine='xlsxwriter')
    option_df.to_excel(writer, sheet_name='customerprofileship')
    writer.save()
    xlsx_io.seek(0)

    return flask.send_file(
        xlsx_io,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        attachment_filename=f'export_customerprofileship.xlsx',
        as_attachment=True,
        cache_timeout=0
    )

#---------------------------------------------------------------------------------------------------------------------------------------------------

#The customerprofilefcst graph
@app.callback(dash.dependencies.Output('customerprofilefcst', 'figure'),
              [dash.dependencies.Input('cust_profile_button', 'n_clicks')],
              [dash.dependencies.State('CustomerName3', 'value'),
               dash.dependencies.State('interval', 'start_date'),
               dash.dependencies.State('interval', 'end_date')]
)

def update_customerprofilefcst(n_clicks, CustomerName3, start_date, end_date) :

    df_customerprofilefcst = df_flows[(df_flows["CustomerName"] == CustomerName3) & (df_flows['RampDate'] >= start_date) & (df_flows['RampDate'] <= end_date)];
    df_customerprofilefcst = df_customerprofilefcst.groupby(['Lane'])[['StatFcst']].sum().reset_index();
    df_customerprofilefcst = df_customerprofilefcst[df_customerprofilefcst['StatFcst'] >= 1];
    #print(df_customerprofilefcst);

    filename = '/apps/IMO/Forecasting/PublishForecast/customerprofilefcst_graph_data.csv'
    df_customerprofilefcst.to_csv(filename);
    
    piechartfcst = px.pie(
        df_customerprofilefcst,
        names = 'Lane',
        values = 'StatFcst'
        )

    return piechartfcst
    
#--------------------------------------------------------------------------------------------------------------------------------------------------

@app.server.route('/export/excel/customerprofilefcst')
def export_excel_customerprofilefcst():
    
    option_df = pd.read_csv('/apps/IMO/Forecasting/PublishForecast/customerprofilefcst_graph_data.csv')

    xlsx_io = io.BytesIO()
    writer = pd.ExcelWriter(xlsx_io, engine='xlsxwriter')
    option_df.to_excel(writer, sheet_name='customerprofilefcst')
    writer.save()
    xlsx_io.seek(0)

    return flask.send_file(
        xlsx_io,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        attachment_filename=f'export_customerprofilefcst.xlsx',
        as_attachment=True,
        cache_timeout=0
    )

#========================================================================================================================================================================

#The laneprofiletbl graph
@app.callback(dash.dependencies.Output('laneprofiletbl', 'figure'),
              [dash.dependencies.Input('lane_profile_button', 'n_clicks')],
              [dash.dependencies.State('Lane', 'value'),
               dash.dependencies.State('interval', 'start_date'),
               dash.dependencies.State('interval', 'end_date')]
)
def update_laneprofiletbl(n_clicks, lane, start_date, end_date) :

    df_laneprofiletbl = df_flows[(df_flows["Lane"] == lane) & (df_flows['RampDate'] >= start_date) & (df_flows['RampDate'] <= end_date)];
    df_laneprofiletbl = df_laneprofiletbl.groupby(['CustomerName', 'CustomerCode'])[['ACTdepLD','ACTdepFR', 'ACTdepMT', 'StatFcst', 'ConsensusFcst']].sum().reset_index();
    df_laneprofiletbl = df_laneprofiletbl[(df_laneprofiletbl['ACTdepLD'] >= 1) | (df_laneprofiletbl['ACTdepFR'] >= 1) | (df_laneprofiletbl['ACTdepMT'] >=1) | (df_laneprofiletbl['StatFcst'] >= 1)];
    #print(df_laneprofiletbl);
    	
    filename = '/apps/IMO/Forecasting/PublishForecast/laneprofile_graph_data.csv'
    df_laneprofiletbl.to_csv(filename);

    table = go.Figure(
        data=[go.Table(
            header = dict( values = list(df_laneprofiletbl.columns), fill_color = 'paleturquoise', align = 'left'),
		    cells  = dict( values = [df_laneprofiletbl.CustomerName, df_laneprofiletbl.CustomerCode, df_laneprofiletbl.ACTdepLD, df_laneprofiletbl.ACTdepFR, df_laneprofiletbl.ACTdepMT, df_laneprofiletbl.StatFcst, df_laneprofiletbl.ConsensusFcst], fill_color = 'lavender', align = 'left')
        )],
        layout=dict(
            title = 'LOADS BY CUSTOMER',
            height = 1500,
            width  = 2000,
        ),
    )

    return table

#--------------------------------------------------------------------------------------------------------------------------------------------------

@app.server.route('/export/excel/laneprofile')
def export_excel_laneprofile():
    
    option_df = pd.read_csv('/apps/IMO/Forecasting/PublishForecast/laneprofile_graph_data.csv')

    xlsx_io = io.BytesIO()
    writer = pd.ExcelWriter(xlsx_io, engine='xlsxwriter')
    option_df.to_excel(writer, sheet_name='laneprofile')
    writer.save()
    xlsx_io.seek(0)

    return flask.send_file(
        xlsx_io,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        attachment_filename=f'export_laneprofile.xlsx',
        as_attachment=True,
        cache_timeout=0
    )

#==============================================================================================================================================================

if __name__ == '__main__':
    app.run_server(debug=False, host = "0.0.0.0", port=8052) # host = "0.0.0.0"

cursor.close()
del cursor