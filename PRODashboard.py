#!/usr/bin/env python
# coding: utf-8

# In[1]:


import dash
import plotly.express as px
import pandas as pd
import numpy as np
from datetime import datetime as dt


# In[2]:


df = pd.read_excel("PRO_DB_Mock.xlsx",  sheet_name = ['PRO DB'])

df = df['PRO DB']


# In[3]:


df.head()


# In[4]:


df.drop(range(0,3))


# In[5]:


header = df.iloc[3]

df = df[4:]

df.head()


# In[6]:


df = df.rename(columns=header)

df = df.dropna(axis='columns',how='all')


# In[7]:


df.columns.tolist()


# In[8]:


df = df.drop(['RPT ID',
              'Asset Family Title',
              'Is this a formatting project?',
              'Primary Service Area',
              'Project Owner',
              'Lead Author',
              'Supporting Author(s)',
              'C=CORE\nPC=Priority CORE\nNC=Non-CORE\nND=Non-deliverable',
              'Is Penang Research Operations (PRO) involved?',
              'PRO Lead Analyst',
              'PRO Supporting Analyst(s)',
              'Check column',
              'Check  column 2',
              'Pillar Analyst estimated effort(days)',
              'Is there any dependency to start project?',
              'Is this project deadline estimated or actual?',
              'Submission Status',
              'Actual Submission Date',
              'Revised Submission Date',
              'PRO / Pillar Analyst delay',
              'Reason for delay/revision',
              'Is this project to be submitted directly to E&P?',
              'PLANNED SUBMISSION DATE TO E&P',
              'CURRENT PLANNED PUB DATE',
              'Is peer review required?',
              'Is peer review done by PIC or QC Checker?',
              'Name of person doing QC',
              'Peer Review Effort Estimated (Days)',
              'Is there an SOP available for this project?',
              'Current SOP Status',
              'What is the SOP completion quarter?',
              'RPT Matching Notes',
              'RPT ID for additional piece ',
              'Notes'], axis=1)


# In[9]:


df = df.reset_index(drop=True)

df.head()


# In[10]:


df['PRO estimated effort (days)'] = df['PRO estimated effort (days)'].fillna(0)
df['Project No'] = df['Project No'].fillna(0)
df['Original  Submission Date'] = pd.to_datetime(df['Original  Submission Date'], errors='coerce')
df = df.dropna(subset=['Original  Submission Date'])

df.head()


# In[11]:


df1 = pd.melt(df, 
              id_vars = ['Project No','Deliverable title', 'PRO Manager','Pillar', 'Original  Submission Date'], 
              value_vars=['Data Collection','Secondary Research', 'Data cleansing', 'Content writing', 'Analysis', 
                          'Primary research ', 'Workbook updating', 'Report formatting', 'Data Visualization', 
                          'Data Transformation', 'Process Automation'],
              var_name='Activities', value_name='value')

df1.head()


# In[12]:


df1.to_excel(r'C:\Users\GurmeetSingS\OneDrive - Informa plc\Personal\File Name.xlsx', index = False)


# In[13]:


colors = ['rgb(68,114,196)', 'rgb(237,125,49)', 'rgb(165,165,165)', 'rgb(255,192,0)', 'rgb(95,155,213)',
          'rgb(112,173,71)', 'rgb(31,78,120)', 'rgb(112,173,71)' , 'rgb(112,2,71)', 'rgb(0,173,71)']


# In[14]:

from dash import dcc
import dash_html_components as html
from dash.dependencies import Output, Input, State
import plotly.graph_objects as go


# In[15]:


app = dash.Dash(__name__)

external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

all = np.sort(df.Pillar.unique())

all_1 = np.sort(df['PRO Manager'].unique())

app.layout=html.Div([
    html.Div([
        html.H1("PRO Project Management dashboard"),
        dcc.Tabs(id="tabs-example", value='tab-example', children=[
        dcc.Tab(label='Tab One', value='tab-1-example'), 
        dcc.Tab(label='Tab Two', value='tab-2-example'),
    ]),
    html.Div(id='tabs-content-example')
    ]),
        
])


# In[16]:


@app.callback(Output('tabs-content-example','children'),
             Input('tabs-example','value'))

def render_content(tab):
    if tab == 'tab-1-example':
        return  html.Div([
        html.Div([
            html.Div([
                html.P('Pillar'),
                dcc.Dropdown(id='pillar-choice', options=[{'label':x, 'value':x} for x in all] + [{'label':'Select All' , 'value': 'all'}], value=['all'], multi=True, clearable = True),
                    ],className='four columns'), 
            
            html.Div([
                html.P('Manager'),
                dcc.Dropdown(id='manager-choice', options=[{'label':k, 'value':k} for k in all_1], value=[], multi=True),
                    ], className='four columns'),
            
            html.Div([
                html.P('Date Range'),
                dcc.DatePickerRange(
                id='my-date-picker-range', calendar_orientation = 'horizontal', start_date_placeholder_text="Start Period",
                    end_date_placeholder_text="End Period", min_date_allowed=(2020, 11, 31),
                    max_date_allowed=(2021, 11, 31), start_date= dt(2021,1,1).date(), end_date=dt(2021,12,31).date()),
                 html.H3(id='output-container-date-picker-range')
                    ], className='four columns'),
                    ], className='row'),   
            
        html.Div([
            dcc.Graph(id='graph1', style={'display':'inline-block', 'width' :'33%'}),
            dcc.Graph(id='graph2', style={'display':'inline-block', 'width' :'33%'}),
            dcc.Graph(id='graph3', style={'display':'inline-block', 'width' :'33%'})
            ], className='row'),
            
        html.Div([
            dcc.Graph(id='graph4', style={'display':'inline-block', 'width' :'33%'}),
            dcc.Graph(id='graph5', style={'display':'inline-block', 'width' :'33%'}),
            dcc.Graph(id='graph6', style={'display':'inline-block', 'width' :'33%'})
        ])
        
        ])
    
    elif tab == 'tab-2-example':
        return html.Div([
        html.Div([
            html.Div([
                html.P('Pillar'),
                dcc.Dropdown(id='pillar-choice', options=[{'label':x, 'value':x} for x in all] + [{'label':'Select All' , 'value': 'all'}], value=['all'], multi=True, clearable = True),
                    ],className='four columns'), 
            
            html.Div([
                html.P('Manager'),
                dcc.Dropdown(id='manager-choice', options=[{'label':k, 'value':k} for k in all_1], value=[], multi=True, clearable=False),
                    ], className='four columns'),
            
            html.Div([
                html.P('Date Range'),
                dcc.DatePickerRange(
                id='my-date-picker-range', calendar_orientation = 'horizontal', start_date_placeholder_text="Start Period",
                    end_date_placeholder_text="End Period", min_date_allowed=(2020, 11, 31),
                    max_date_allowed=(2021, 11, 31), start_date= dt(2021,1,1).date(), end_date=dt(2021,12,31).date()),
                 html.H3(id='output-container-date-picker-range')
                    ], className='four columns'),
                    ], className='row'),   
    
            
            html.Div([
            dcc.Graph(id='graph7', style={'display':'inline-block', 'width' :'33%'})
            ])
        ])


# In[17]:


@app.callback(
    Output(component_id='graph1', component_property='figure'),
    [Input(component_id='pillar-choice', component_property='value'),
     Input(component_id='manager-choice', component_property='value'),
     Input(component_id='my-date-picker-range', component_property='start_date'),
     Input(component_id='my-date-picker-range', component_property='end_date'),]
)

def update_graph1(value_pillar, value_manager, start_date, end_date):
        if ((value_pillar == ['all']) and (len(value_manager) == 0)):
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
            
        elif ((value_pillar == ['all']) and (len(value_manager) != 0)):
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date) & 
                  df['PRO Manager'].isin(value_manager)]
            
        elif ((len(value_pillar) >= 1) and (len(value_manager) >= 1)):
            dff = df[df.Pillar.isin(value_pillar) & df['PRO Manager'].isin(value_manager) &
                     (df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
        
        elif ((len(value_pillar) == 0) and (len(value_manager) >= 1)):
            dff =df[ df['PRO Manager'].isin(value_manager) & (df['Original  Submission Date']>=start_date) & 
                    (df['Original  Submission Date']<=end_date)]
           
        elif ((len(value_pillar) >= 1) and (len(value_manager) == 0)):
            dff =df[df['Pillar'].isin(value_pillar) & (df['Original  Submission Date']>=start_date) & 
                    (df['Original  Submission Date']<=end_date) ]
            
        else:
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
        
        
        fig = go.Figure(data = [go.Pie(labels=dff['Pillar'], values = dff['PRO estimated effort (days)'], marker=dict(colors=colors), sort=False)])
        fig.update_layout (title = {'text' : "Estimated effort (days) by Pillar", 'x' : 0.4, 'y' : 0.95, 'xanchor' : 'center', 
                                    'yanchor' : 'top',
                                   'font_size' : 18})
        return fig


# In[18]:


@app.callback(
    Output(component_id='graph2', component_property='figure'),
    [Input(component_id='pillar-choice', component_property='value'),
     Input(component_id='manager-choice', component_property='value'),
     Input(component_id='my-date-picker-range', component_property='start_date'),
     Input(component_id='my-date-picker-range', component_property='end_date'),]
)

def update_graph2(value_pillar, value_manager, start_date, end_date):
        if ((value_pillar == ['all']) and (len(value_manager) == 0)):
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
            
        elif ((value_pillar == ['all']) and (len(value_manager) != 0)):
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date) & 
                  df['PRO Manager'].isin(value_manager)]
            
        elif ((len(value_pillar) >= 1) and (len(value_manager) >= 1)):
            dff = df[df.Pillar.isin(value_pillar) & df['PRO Manager'].isin(value_manager) &
                     (df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
        
        elif ((len(value_pillar) == 0) and (len(value_manager) >= 1)):
            dff =df[ df['PRO Manager'].isin(value_manager) & (df['Original  Submission Date']>=start_date) & 
                    (df['Original  Submission Date']<=end_date)]
           
        elif ((len(value_pillar) >= 1) and (len(value_manager) == 0)):
            dff =df[df['Pillar'].isin(value_pillar) & (df['Original  Submission Date']>=start_date) & 
                    (df['Original  Submission Date']<=end_date) ]
            
        else:
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
        
        
        fig = go.Figure(data = [go.Pie(labels=dff['Cadence'], values = dff['PRO estimated effort (days)'], marker=dict(colors=colors), sort=False)])
        fig.update_layout (title = {'text' : "Estimated effort (days) by Cadence", 'x' : 0.4, 'y' : 0.95, 'xanchor' : 'center', 
                                    'yanchor' : 'top',
                                   'font_size' : 18})
        return fig


# In[19]:


@app.callback(
    Output(component_id='graph3', component_property='figure'),
    [Input(component_id='pillar-choice', component_property='value'),
     Input(component_id='manager-choice', component_property='value'),
     Input(component_id='my-date-picker-range', component_property='start_date'),
     Input(component_id='my-date-picker-range', component_property='end_date'),]
)

def update_graph3(value_pillar, value_manager, start_date, end_date):
        if ((value_pillar == ['all']) and (len(value_manager) == 0)):
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
            
        elif ((value_pillar == ['all']) and (len(value_manager) != 0)):
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date) & 
                  df['PRO Manager'].isin(value_manager)]
            
        elif ((len(value_pillar) >= 1) and (len(value_manager) >= 1)):
            dff = df[df.Pillar.isin(value_pillar) & df['PRO Manager'].isin(value_manager) &
                     (df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
        
        elif ((len(value_pillar) == 0) and (len(value_manager) >= 1)):
            dff =df[ df['PRO Manager'].isin(value_manager) & (df['Original  Submission Date']>=start_date) & 
                    (df['Original  Submission Date']<=end_date)]
           
        elif ((len(value_pillar) >= 1) and (len(value_manager) == 0)):
            dff =df[df['Pillar'].isin(value_pillar) & (df['Original  Submission Date']>=start_date) & 
                    (df['Original  Submission Date']<=end_date) ]
            
        else:
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
        
        
        fig = go.Figure(data = [go.Pie(labels=dff['Project Type'], values = dff['PRO estimated effort (days)'], marker=dict(colors=colors), sort=False)])
        fig.update_layout (title = {'text' : "Estimated effort (days) by Project Type", 'x' : 0.4, 'y' : 0.95, 'xanchor' : 'center', 
                                    'yanchor' : 'top',
                                   'font_size' : 18})
        return fig
    


# In[20]:


@app.callback(
    Output(component_id='graph4', component_property='figure'),
    [Input(component_id='pillar-choice', component_property='value'),
     Input(component_id='manager-choice', component_property='value'),
     Input(component_id='my-date-picker-range', component_property='start_date'),
     Input(component_id='my-date-picker-range', component_property='end_date'),]
)

def update_graph4(value_pillar, value_manager, start_date, end_date):
        if ((value_pillar == ['all']) and (len(value_manager) == 0)):
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
            
        elif ((value_pillar == ['all']) and (len(value_manager) != 0)):
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date) & 
                  df['PRO Manager'].isin(value_manager)]
            
        elif ((len(value_pillar) >= 1) and (len(value_manager) >= 1)):
            dff = df[df.Pillar.isin(value_pillar) & df['PRO Manager'].isin(value_manager) &
                     (df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
        
        elif ((len(value_pillar) == 0) and (len(value_manager) >= 1)):
            dff =df[ df['PRO Manager'].isin(value_manager) & (df['Original  Submission Date']>=start_date) & 
                    (df['Original  Submission Date']<=end_date)]
           
        elif ((len(value_pillar) >= 1) and (len(value_manager) == 0)):
            dff =df[df['Pillar'].isin(value_pillar) & (df['Original  Submission Date']>=start_date) & 
                    (df['Original  Submission Date']<=end_date) ]
            
        else:
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
        
        
        fig = go.Figure(data = [go.Pie(labels=dff['Pillar'], values = dff['Project No'], marker=dict(colors=colors), sort=False)])
        fig.update_layout (title = {'text' : "Number of projects by Pillar", 'x' : 0.4, 'y' : 0.95, 'xanchor' : 'center', 
                                    'yanchor' : 'top',
                                   'font_size' : 18})
        return fig
    


# In[21]:


@app.callback(
    Output(component_id='graph5', component_property='figure'),
    [Input(component_id='pillar-choice', component_property='value'),
     Input(component_id='manager-choice', component_property='value'),
     Input(component_id='my-date-picker-range', component_property='start_date'),
     Input(component_id='my-date-picker-range', component_property='end_date'),]
)

def update_graph5(value_pillar, value_manager, start_date, end_date):
        if ((value_pillar == ['all']) and (len(value_manager) == 0)):
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
            
        elif ((value_pillar == ['all']) and (len(value_manager) != 0)):
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date) & 
                  df['PRO Manager'].isin(value_manager)]
            
        elif ((len(value_pillar) >= 1) and (len(value_manager) >= 1)):
            dff = df[df.Pillar.isin(value_pillar) & df['PRO Manager'].isin(value_manager) &
                     (df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
        
        elif ((len(value_pillar) == 0) and (len(value_manager) >= 1)):
            dff =df[ df['PRO Manager'].isin(value_manager) & (df['Original  Submission Date']>=start_date) & 
                    (df['Original  Submission Date']<=end_date)]
           
        elif ((len(value_pillar) >= 1) and (len(value_manager) == 0)):
            dff =df[df['Pillar'].isin(value_pillar) & (df['Original  Submission Date']>=start_date) & 
                    (df['Original  Submission Date']<=end_date) ]
            
        else:
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
        
        
        fig = go.Figure(data = [go.Pie(labels=dff['Cadence'], values = dff['Project No'], marker=dict(colors=colors), sort=False)])
        fig.update_layout (title = {'text' : "Number of projects by Cadence", 'x' : 0.4, 'y' : 0.95, 'xanchor' : 'center', 
                                    'yanchor' : 'top',
                                   'font_size' : 18})
        return fig
    


# In[22]:


@app.callback(
    Output(component_id='graph6', component_property='figure'),
    [Input(component_id='pillar-choice', component_property='value'),
     Input(component_id='manager-choice', component_property='value'),
     Input(component_id='my-date-picker-range', component_property='start_date'),
     Input(component_id='my-date-picker-range', component_property='end_date'),]
)

def update_graph6(value_pillar, value_manager, start_date, end_date):
        if ((value_pillar == ['all']) and (len(value_manager) == 0)):
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
            
        elif ((value_pillar == ['all']) and (len(value_manager) != 0)):
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date) & 
                  df['PRO Manager'].isin(value_manager)]
            
        elif ((len(value_pillar) >= 1) and (len(value_manager) >= 1)):
            dff = df[df.Pillar.isin(value_pillar) & df['PRO Manager'].isin(value_manager) &
                     (df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
        
        elif ((len(value_pillar) == 0) and (len(value_manager) >= 1)):
            dff =df[ df['PRO Manager'].isin(value_manager) & (df['Original  Submission Date']>=start_date) & 
                    (df['Original  Submission Date']<=end_date)]
           
        elif ((len(value_pillar) >= 1) and (len(value_manager) == 0)):
            dff =df[df['Pillar'].isin(value_pillar) & (df['Original  Submission Date']>=start_date) & 
                    (df['Original  Submission Date']<=end_date) ]
            
        else:
            dff = df[(df['Original  Submission Date']>=start_date) & (df['Original  Submission Date']<=end_date)]
        
        
        fig = go.Figure(data = [go.Pie(labels=dff['Project Type'], values = dff['Project No'], marker=dict(colors=colors), sort=False)])
        fig.update_layout (title = {'text' : "Number of projects by Project Type", 'x' : 0.4, 'y' : 0.95, 'xanchor' : 'center', 
                                    'yanchor' : 'top',
                                   'font_size' : 18})
        return fig
    


# In[23]:


@app.callback(
    Output(component_id='graph7', component_property='figure'),
    [Input(component_id='pillar-choice', component_property='value'),
     Input(component_id='manager-choice', component_property='value'),
     Input(component_id='my-date-picker-range', component_property='start_date'),
     Input(component_id='my-date-picker-range', component_property='end_date'),]
)

def update_graph7(value_pillar, value_manager, start_date, end_date):
        if ((value_pillar == ['all']) and (len(value_manager) == 0)):
            dff = df1[(df1['Original  Submission Date']>=start_date) & (df1['Original  Submission Date']<=end_date)]
            
        elif ((value_pillar == ['all']) and (len(value_manager) != 0)):
            dff = df1[(df1['Original  Submission Date']>=start_date) & (df1['Original  Submission Date']<=end_date) & 
                  df1['PRO Manager'].isin(value_manager)]
            
        elif ((len(value_pillar) >= 1) and (len(value_manager) >= 1)):
            dff = df1[df1.Pillar.isin(value_pillar) & df1['PRO Manager'].isin(value_manager) &
                     (df1['Original  Submission Date']>=start_date) & (df1['Original  Submission Date']<=end_date)]
        
        elif ((len(value_pillar) == 0) and (len(value_manager) >= 1)):
            dff =df1[df1['PRO Manager'].isin(value_manager) & (df1['Original  Submission Date']>=start_date) & 
                    (df1['Original  Submission Date']<=end_date)]
           
        elif ((len(value_pillar) >= 1) and (len(value_manager) == 0)):
            dff =df1[df1['Pillar'].isin(value_pillar) & (df1['Original  Submission Date']>=start_date) & 
                    (df1['Original  Submission Date']<=end_date) ]
            
        else:
            dff = df1[(df1['Original  Submission Date']>=start_date) & (df1['Original  Submission Date']<=end_date)]
        
        
        fig = go.Figure(data = [go.Pie(labels=dff['Activities'], values = dff['value'], marker=dict(colors=colors),
                                sort=False)])
        fig.update_layout (title = {'text' : "Estimated effort (days) by Activities", 'x' : 0.4, 'y' : 0.95, 
                                    'xanchor' : 'center', 'yanchor' : 'top', 'font_size' : 18})
        
        return fig


# In[ ]:


server = app.server

if __name__=='__main__':
    app.run_server()

