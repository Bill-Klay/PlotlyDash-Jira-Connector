import dash
import json
from dash.dependencies import Input, Output, State
from dash.exceptions import PreventUpdate
import dash_core_components as dcc
import dash_html_components as html
import dash_table
import warnings
from jira import JIRA
import pandas as pd
from datetime import datetime
from datetime import date
from datetime import timedelta
import pyodbc
import plotly.graph_objs as go
from fuzzywuzzy import process
import numpy as np
import win32com.client as win32
import time

warnings.filterwarnings('ignore')

app = dash.Dash(__name__)

#Global variables for passing into multiple callbacks
jiraData = pd.DataFrame()
workLog = pd.DataFrame()
workLogAlpha = pd.DataFrame()
option = []
filename = ""
startDate = date.today()
endDate = date.today()

#The default app/dashboard layout (Screenshot can be observed on readme)
app.layout = html.Div([
    
    html.Div([html.Span('Jira Connector', className='circle-sketch-highlight', style={'font-size': '2.5em', 'font-family':'Script MT', 'marginLeft':'20px'}),
              html.Span(dcc.Loading(children=[html.Div(id='loading')], type='circle'), style={'marginLeft': '10%'})
    ]),
    html.Span('Review team worklog', style={'marginLeft': '20px', 'font-family':'Script MT'}),
    
    html.Hr(),

    # Slider date input for as far as 30 days in the past
    #dcc.Slider(
    #    id='daysSlider',
    #    min=0,
    #    max=30,
    #    value=14,
    #    step=1,
    #    marks={
    #        0: {'label':'0 Days', 'style':{'color': 'purple', 'font-face': 'bold'}},
    #        5: {'label': '5 Days'},
    #        10: {'label': '10 Days'},
    #        15: {'label': '15 Days'},
    #        20: {'label':'20 Days'},
    #        25: {'label':'25 Days'},
    #        30: {'label':'30 Days', 'style':{'color': 'purple', 'font-face': 'bold'}}
    #    }
    #),

    html.Div('Enter Start and End Date for Work Log', style={'marginLeft': '20px', 'fontWeight': 'bold', 'fontStyle': 'italic'}),

    html.Div(children=[dcc.DatePickerRange(
                            id='datePickerRange',
                            min_date_allowed= date.today() - timedelta(days=int(365)),
                            max_date_allowed= date.today(),
                            initial_visible_month=date.today(),
                            end_date=date.today(),
                            start_date = date.today() - timedelta(days=int(2))
                            ),
                       html.Span(id='days', style={'font-size': '1.2em', 'font-family':'Calibri', 'marginLeft':'20px'}),
                       ],
             style={'marginLeft': '20px'}),

    html.Br(),

    html.Div(children=[html.Button('Switch Table', id='convert'), 
                       html.Span(id='ticketCount', style={'font-size': '1.2em', 'font-family':'Calibri', 'marginLeft':'20px'})
                       ], 
             style={'marginLeft': '20px'}),
    
    html.Br(),
    
    html.Div(id='workLogAlpha-dataTable', children=[], style={'marginLeft': '20px', 'marginRight': '20px'}),

    html.Div(children=[
        dash_table.DataTable(
            id='dashboard',
            data=[],
            columns=[
                     {'name': 'Issue key', 'id': 'Issue key'},
                     {'name': 'Summary', 'id': 'Summary'},
                     {'name': 'Assignee name', 'id': 'Assignee name'},
                     {'name': 'Estimate (hrs)', 'id': 'Estimate (hrs)'},
                     {'name': 'Time spent (hrs)', 'id': 'Time spent (hrs)'},
                     {'name': 'Logged Date Time', 'id': 'Logged Date Time'}
                    ],
            style_cell={'textAlign': 'left', 'textOverflow': 'ellipsis', 'minWidth': '20px', 'maxWidth': '400px'},
            style_table={'overflowX': 'auto', 'overflowY': 'auto'},
            page_size=10,
            style_header={'fontWeight': 'bold', 'backgroundColor': '#c799b9', 'fontFamily': 'Calibri', 'fontStyle': 'italic'},
            style_data={'fontFamily': 'Calibri'},
            style_data_conditional=[{'if': {'row_index': 'odd'}, 'backgroundColor': 'rgb(248, 248, 248)'}]
    )],
            id='workLog-dataTable',
            style={'marginLeft': '20px', 'marginRight': '20px'}
    ),

    html.Br(),

    html.Div([html.Span([html.Button("Download Summary", id="summaryFile"), dcc.Download(id="downloadSummaryFile")], style={'marginLeft': '10%'})]),

    html.Br(),
    html.Br(),

    html.Div([
        dcc.Dropdown(
            id = 'names', 
            options = option, 
            placeholder = 'Select user', 
            multi = True
        )
    ],
        style = {'display' : 'inline-block', 'verticalAlign' : 'top', 'width' : '30%', 'marginLeft': '20px'}
    ),
    
    html.Div([
        dcc.Dropdown(
            id = 'graphs',
            options = [{'label': 'Line plot', 'value': 'Line'}, {'label': 'Bubble plot', 'value': 'Bubble'}, {'label': 'Bar plot', 'value': 'Bar'}, {'label': 'Heatmap', 'value': 'Heatmap'}],
            placeholder = 'Select graph',
            multi = False,
            value = 'Bar'
        )
    ],
        style = {'display': 'inline-block', 'verticalAlign': 'top', 'width': '30%', 'marginLeft': '30px'}
    ),
    
    html.Div(children=[dcc.Graph(id='overall-graph')], style={'marginLeft': '20px'}, id='graphsDiv'),

    html.Br(),
    
    html.Div(children=[html.Div(dcc.Graph(id='grouping-graph'), style={'marginLeft': '20px', 'width': '50%', 'float': 'left'}),
                       html.Div(dash_table.DataTable(id='teamTable',
                            data=[],
                            columns=[
                                    {'name': 'Team', 'id': 'Team'},
                                    {'name': 'Allocation (%)', 'id': 'Allocation (%)'},
                                    {'name': 'Utilization (%)', 'id': 'Utilization (%)'}
                                    ],
                            style_cell={'textAlign': 'left', 'textOverflow': 'ellipsis', 'minWidth': '20px', 'maxWidth': '400px'},
                            style_table={'overflowX': 'auto', 'overflowY': 'auto'},
                            page_size=10,
                            style_header={'fontWeight': 'bold', 'backgroundColor': '#c799b9', 'fontFamily': 'Calibri', 'fontStyle': 'italic'},
                            style_data={'fontFamily': 'Calibri'},
                            style_data_conditional=[{'if': {'row_index': 'odd'}, 'backgroundColor': 'rgb(248, 248, 248)'}]
                        ), style={'marginRight': '20px', 'width': '50%', 'float': 'left', 'marginTop': '120px'})
                    ], style={'display': 'flex'}),
    html.Br()

])

#Merge default and worklog dataframe
def fuzzy_merge(df_1, df_2, key1, key2, threshold=90, limit=2):
    """
    :param df_1: the left table to join
    :param df_2: the right table to join
    :param key1: key column of the left table
    :param key2: key column of the right table
    :param threshold: how close the matches should be to return a match, based on Levenshtein distance
    :param limit: the amount of matches that will get returned, these are sorted high to low
    :return: dataframe with boths keys and matches
    """
    s = df_2[key2].tolist()
    
    m = df_1[key1].apply(lambda x: process.extract(x, s, limit=limit))    
    df_1['matches'] = m
    
    m2 = df_1['matches'].apply(lambda x: ', '.join([i[0] for i in x if i[1] >= threshold]))
    df_1['matches'] = m2
    
    df_1[key1] = df_1['matches']
    
    return df_1

#Color Cells
def colorCells(worksheet, color_range, badRed, goodGreen, neutralYellow, overGrey, missingBlue):
    worksheet.conditional_format(color_range, {'type': 'cell',
                                            'criteria': '<',
                                            'value': 0.500,
                                            'format': badRed})
    worksheet.conditional_format(color_range, {'type': 'cell',
                                            'criteria': 'between',
                                            'minimum': 0.800,
                                            'maximum': 1.000,
                                            'format': goodGreen})
    worksheet.conditional_format(color_range, {'type': 'cell',
                                            'criteria': 'between',
                                            'minimum': 0.500,
                                            'maximum': 0.799,
                                            'format': neutralYellow})
    worksheet.conditional_format(color_range, {'type': 'text',
                                            'criteria': 'containing',
                                            'value': 'Estimates Missing',
                                            'format': missingBlue})
    worksheet.conditional_format(color_range, {'type': 'cell',
                                            'criteria': '>',
                                            'value': 1.000,
                                            'format': overGrey})
    worksheet.conditional_format(color_range, {'type': 'text',
                                            'criteria': 'containing',
                                            'value': 'On Leaves',
                                            'format': missingBlue})
    
def generalFormat(style, color_range, worksheet):
    worksheet.conditional_format(color_range, {'type': 'cell',
                                            'criteria': '>',
                                            'value': 0,
                                            'format': style})

#Adjust the past days slider
#@app.callback([
#    Output('days', 'children'),
#    Input('daysSlider', 'value')
#])
#def sliderUpdate(days):
#    return ['{} Days'.format(days)]

#Adjust the past days slider
@app.callback([
    Output('days', 'children'),
    Input('datePickerRange', 'start_date'),
    Input('datePickerRange', 'end_date')
])
def sliderUpdate(start_Date, end_Date):
    startDate = date.fromisoformat(start_Date)
    endDate = date.fromisoformat(end_Date)
    diff = endDate - startDate
    return ['{} Days'.format(diff.days)]

#Download Summary file
@app.callback([
    Output('downloadSummaryFile', 'data'),
    Input('summaryFile', 'n_clicks')
])
def fileDownload(click):
    global filename
    if click is None:
        raise PreventUpdate
    else:
        return [dcc.send_file(".\Output\{}".format(filename))]

#This callback works for the switch table button, there two views to the main data table
#Shows all tickets retrieved and summary of estimation and utilization
#Returns:
    #The new data table with respect to the click
    #Disables the previous Div to hide the old data table
#Input:
    #Click on the switch data table button
@app.callback([
    Output('workLogAlpha-dataTable', 'children'),
    Output('workLog-dataTable', 'children'),
    Input('convert', 'n_clicks')
])
def changeTable(click):
    global workLogAlpha, workLog

    if click is None:
        raise PreventUpdate
    elif click%2 == 0:
        return [html.Div(style={'display': 'None'}), html.Div([dash_table.DataTable(
                id='dashboard',
                data=jiraData.to_dict('records'),
                columns=[{'name': i, 'id': i} for i in jiraData.columns],
                style_cell={'textAlign': 'left', 'textOverflow': 'ellipsis', 'minWidth': '20px', 'maxWidth': '400px'},
                style_table={'overflowX': 'auto', 'overflowY': 'auto'},
                page_size=10,
                style_header={'fontWeight': 'bold', 'backgroundColor': '#c799b9', 'fontFamily': 'Calibri', 'fontStyle': 'italic'},
                style_data={'fontFamily': 'Calibri'},
                style_data_conditional=[{'if': {'row_index': 'odd'}, 'backgroundColor': 'rgb(248, 248, 248)'}]
            )])
        ]
    else:
        return [html.Div([dash_table.DataTable(
                id='dashboard1',
                data=workLogAlpha.to_dict('records'),
                columns=[{'name': i, 'id': i} for i in workLogAlpha.columns],
                style_cell={'textAlign': 'left', 'textOverflow': 'ellipsis', 'minWidth': '20px', 'maxWidth': '400px'},
                style_table={'overflowX': 'auto', 'overflowY': 'auto'},
                page_size=10,
                style_header={'fontWeight': 'bold', 'backgroundColor': '#c799b9', 'fontFamily': 'Calibri', 'fontStyle': 'italic'},
                style_data={'fontFamily': 'Calibri'},
                style_data_conditional=[{'if': {'row_index': 'odd'}, 'backgroundColor': 'rgb(248, 248, 248)'}]
            )]), html.Div(style={'display': 'None'})
        ]

#This callback deals with the changes on the main overall graph
#It also deals with selecting specific names and graphs for viewing
#Returs:
    #Changes in the overall grahp Div wrt the names and graph selected
#Inputs:
    #Names from the drop down list
    #Graph for viewing
@app.callback(Output('graphsDiv', 'children'),
              Input('names', 'value'),
              Input('graphs', 'value'),
              prevent_initial_call=True
)
def selectName (names, graph):
    global workLogAlpha

    mode = 'lines+markers'
    marker = {'size': 12, 'symbol': 'circle-dot', 'color': 'darkviolet'}

    #If nothing is selected then return the default overall bar graph
    trace1 = go.Bar(
                    x = workLogAlpha['Assignee'],
                    y = workLogAlpha['Allocation (%)'],
                    marker = {'color': 'rgb(199, 153, 185)'},
                    name = 'Allocation'
    )
    trace2 = go.Bar(
                    x = workLogAlpha['Assignee'],
                    y = workLogAlpha['Utilization (%)'],
                    marker = {'color': 'darkviolet'},
                    name = 'Utilization'
    )
    data = [trace1, trace2]
    
    #If bubble graph is selected
    if graph == 'Bubble':
        mode = 'markers'
        marker = dict(size=1.5*workLogAlpha['Utilization (%)'])
        data = [
            go.Scatter(
                x = workLogAlpha['Assignee'],
                y = workLogAlpha['Utilization (%)'],
                mode = mode,
                marker = marker
            )
        ]
    #Bar graph
    elif graph == 'Bar':
         trace1 = go.Bar(
                    x = workLogAlpha['Assignee'],
                    y = workLogAlpha['Allocation (%)'],
                    marker = {'color': 'rgb(199, 153, 185)'},
                    name = 'Estimate'
             )
         trace2 = go.Bar(
                        x = workLogAlpha['Assignee'],
                        y = workLogAlpha['Utilization (%)'],
                        marker = {'color': 'darkviolet'},
                        name = 'Time spent'
                 )
         data = [trace1, trace2]

    #Heatmap
    elif graph == 'Heatmap':
        data = [
            go.Heatmap(
            x = workLogAlpha['Assignee'],
            y = workLogAlpha['Utilization (%)'],
            z = workLogAlpha['Utilization (%)'],
            colorscale = 'Jet',
            zmin = 5, zmax = 40 # add max/min color values to make each plot consistent
            )
        ] 
    #Line graph
    elif graph == 'Line':
        trace1 = go.Scatter(
                    x = workLogAlpha['Assignee'],
                    y = workLogAlpha['Utilization (%)'],
                    mode = mode,
                    marker = marker,
                    name = 'Utilization'
                )
        trace2 = go.Scatter(
                    x = workLogAlpha['Assignee'],
                    y = workLogAlpha['Allocation (%)'],
                    mode = mode,
                    marker = {'size': 12, 'symbol': 'circle-dot', 'color': 'rgb(199, 153, 185)'},
                    name = 'Allocation'
                )
        data = [trace1, trace2]

    figure = {
            'data': data,
            'layout': go.Layout(
                title = 'Overall Time Spent',
                xaxis = {'title': 'Assignee'},
                yaxis = {'title': 'Utilization'},
                hovermode='closest'
            )
        }

    #In addition to selecting a specific graph from the list if there is specific name selected then display data relative to only that list
    if names is not None:
        if len(names):
            # traces = []
            data.clear()
            
            for name in names:
                tempNameList = workLogAlpha[workLogAlpha.Assignee == name]
                if graph == 'Bubble':
                    mode = 'markers'
                    marker = dict(size=2*tempNameList['Utilization (%)'])
                    data.append(go.Scatter(
                        x = tempNameList.Assignee,
                        y = tempNameList['Utilization (%)'],
                        mode = mode,
                        name = name,
                        marker = marker
                    ))
                elif graph == 'Bar':
                    data.append(go.Bar(
                        x = tempNameList.Assignee,
                        y = tempNameList['Utilization (%)'],
                        name = name
                    ))
                elif graph == 'Heatmap':
                    data.append(go.Heatmap(
                        x=tempNameList.Assignee,
                        y=tempNameList['Utilization (%)'],
                        z=tempNameList['Utilization (%)'],
                        colorscale='Jet',
                        zmin = 40, zmax = 100,
                        name = name
                    ))
                elif graph == 'Line':
                    data.append(go.Scatter(
                        x = tempNameList.Assignee,
                        y = tempNameList['Utilization (%)'],
                        name = name,
                        mode = mode,
                        marker = marker
                    ))
                else:
                    data.append(go.Bar(
                        x = tempNameList.Assignee,
                        y = tempNameList['Utilization (%)'],
                        name = name,
                    ))
            fig = {'data': data, 'layout': go.Layout(title = 'Overall Time Spent', xaxis = {'title': 'Assignee'}, yaxis = {'title': 'Utilization'}, hovermode='closest')}
            return [dcc.Graph(id='overall-graph', figure=fig)]
        else:
            return [dcc.Graph(id='overall-graph', figure=figure)]
    else:
        return [dcc.Graph(id='overall-graph', figure=figure)]

#Main app callback that returns the default format after dataframe calculation
#Returns:
    #The main dash data table
    #The days count visible at the range selector
    #Main overall bar graph
    #Names list drop down options
    #Grouped team graph
    #Grouped team data table
    #Number of tickets retrieved
#Input:
    #Number of past days to retrieve tickets from (1-30)
@app.callback([
    Output('dashboard', 'data'),
    Output('loading', 'children'),
    Output('overall-graph', 'figure'),
    Output('names', 'options'),
    Output('grouping-graph', 'figure'),
    Output('teamTable', 'data'),
    Output('ticketCount', 'children'),
    Input('datePickerRange', 'start_date'),
    Input('datePickerRange', 'end_date')
])
def jiraConnector(start_Date, end_Date):
    global jiraData, workLog, workLogAlpha, option, filename, startDate, endDate

    if ((startDate == date.fromisoformat(start_Date)) and endDate == date.fromisoformat(end_Date)):
        raise PreventUpdate
    #else:
    #dayCount = days
    print("-----------------------------------------------------------------")

    initial = 0
    size = 100
    
    startDate = date.fromisoformat(start_Date)
    #startDate = startDate.strftime('%Y-%m-%d')
    startDate = date.fromisoformat(start_Date)
    #endDate = endDate.strftime('%Y-%m-%d')
    endDate = date.fromisoformat(end_Date)
    jql = 'worklogDate >= ' + start_Date + ' AND worklogDate <= ' + end_Date
    print(jql)

    jira = JIRA(options={'server': 'https://qordatainc.atlassian.net/'}, basic_auth=('uzair.islam@qordata.com', 'uxeZzpGkzK6FQvbieIgY1C70')) #Connecting to Jira cloud
    #jql='worklogDate >= -'+str(days)+'d' #The JQL on which the whole data is retrieved

    data_jira = []
    work_log = []
    option.clear()

    #Retrieving tickets from jira cloud and appending to a list
    #Since the maximum count of ticket retrieved at a time in 100 we loop to get all the tickets or specific start and end count
    t0 = time.time()
    while True:
        start = initial * size #Initial start of ticket count
        #Fields to get
        jira_search = jira.search_issues(jql, startAt=start, maxResults=size, 
                                         fields = "key, summary, issuetype, assignee, reporter, status, created, resolutiondate, workratio, timespent, timeoriginalestimate")
        if(len(jira_search) == 0): 
            break

        #Geting all the specificed fields of the tickets in a list
        for issue in jira_search:
            timeSpent = 0
            issue_logDateTime = datetime.today()

            #Retrieving the worklog only in the current range
            #This makes a significant impact on the programs execution time
            for w in jira.worklogs(issue.key):
                started = datetime.strptime(w.started[:-5], '%Y-%m-%dT%H:%M:%S.%f')
                if not (startDate <= started.date() <= endDate):
                    continue
                timeSpent = timeSpent + w.timeSpentSeconds
                issue_logDateTime = datetime.strptime(w.started[:-5], '%Y-%m-%dT%H:%M:%S.%f')

            issue_key = issue.key

            issue_summary = issue.fields.summary #Summary

            request_type = str(issue.fields.issuetype) #Issue type

            datetime_creation = issue.fields.created
            if datetime_creation is not None:
                datetime_creation = datetime.strptime(datetime_creation[:19], "%Y-%m-%dT%H:%M:%S") #Creation datetime

            datetime_resolution = issue.fields.resolutiondate
            if datetime_resolution is not None:
                datetime_resolution = datetime.strptime(datetime_resolution[:19], "%Y-%m-%dT%H:%M:%S") #End datetime

            #Reporter specification
            reporter_login = None
            reporter_name = None
            reporter = issue.raw['fields'].get('reporter', None)
            if reporter is not None:
                reporter_login = reporter.get('key', None)
                reporter_name = reporter.get('displayName', None)

            assignee_login = None
            assignee_name = None
            assignee = issue.raw['fields'].get('assignee', None)
            if assignee is not None:
                assignee_login = assignee.get('key', None)
                assignee_name = assignee.get('displayName', None)

            status = None
            st = issue.fields.status
            if st is not None:
                status = st.name

            #Work ratio and time logged
            issue_workratio = issue.fields.workratio
            issue_timespent = timeSpent
            issue_estimate = issue.fields.timeoriginalestimate

            data_jira.append((issue_key, issue_summary, request_type, datetime_creation, datetime_resolution, reporter_login, reporter_name, assignee_login, assignee_name, status, issue_estimate, issue_timespent, issue_logDateTime))

        initial = initial + 1 #Get the next 100 count

    #List to dataframe
    executionTime = time.time() - t0
    jiraData = pd.DataFrame(data_jira, columns=['Issue key', 'Summary', 'Request type', 'Datetime creation', 'Datetime resolution', 'Reporter login', 'Reporter name', 'Assignee login', 'Assignee name', 'Status', 'Estimate', 'Time spent', 'Logged Date Time'])
    del data_jira

    jira.close()
    
    #Preparing a dataframe (workLogAlpha) for downloading the summary report that shows the estimates and utilization of users in percentage
    workLog = jiraData.groupby(['Assignee name']).sum()
    jiraData[['Time spent [Hours]', 'Estimate [Hours]']] = jiraData[['Time spent', 'Estimate']].div(3600).round(2)
    jiraData = jiraData[['Reporter name', 'Assignee name', 'Time spent', 'Time spent [Hours]', 'Estimate', 'Estimate [Hours]', 'Summary', 'Issue key', 'Logged Date Time']]
    workLog[['Time spent', 'Estimate']] = workLog[['Time spent', 'Estimate']].div(3600)
    workLog.reset_index(inplace=True)
    workLog['Assignee name'] = workLog['Assignee name'].str.replace('.', ' ')
    workLog['Assignee name'] = workLog['Assignee name'].str.title()
    jiraData['Assignee name'] = jiraData['Assignee name'].str.replace('.', ' ')
    jiraData['Assignee name'] = jiraData['Assignee name'].str.title()
    
    #Leaves Data
    leaves = pd.read_excel('.\Input\Attendence.xlsx', na_values=['Present', 'Absent', 'Present, Time out is missing', 
                                                     'Present, Request Pending', 'Holiday', 'Absent, Request Pending'])
    leaves = leaves[['Employee', 'Schedule Date IN', 'Remarks']]
    leaves.dropna(inplace=True)
    gazettedHoliday = leaves[leaves['Employee'] == 'Muhammad Bilal Khan']
    gazettedHoliday = gazettedHoliday[gazettedHoliday['Remarks'] == 'GazettedHoliday']['Schedule Date IN'].tolist()
      
    onLeaves = leaves[leaves['Remarks'] == 'On Leave']
    onLeaves.drop('Remarks', axis=1, inplace=True)
    leavesDict = {}
    for index, row in onLeaves.iterrows():
        leavesDict[row.Employee] = onLeaves[onLeaves['Employee'] == row.Employee]
    groups = leaves.groupby('Employee').count().mul(8)
    groups.reset_index(inplace=True)
    groups.rename(columns={'Employee': 'Assignee name', 'Remarks': 'Leaves'}, inplace=True)

    #Using PMs list for names and departments
    defaultDf = pd.read_excel('.\Input\Employee List PM.xlsx')
    defaultDf.rename(columns = {'Name': 'Assignee name'}, inplace=True)
    workLog = fuzzy_merge(workLog, defaultDf, 'Assignee name', 'Assignee name', 80, 1)
    defaultDf = defaultDf.merge(workLog, on='Assignee name', how='left')
    groups = fuzzy_merge(groups, defaultDf, 'Assignee name', 'Assignee name', 90, 1)
    defaultDf = defaultDf.merge(groups, on='Assignee name', how='left')
    defaultDf.fillna(0, inplace = True)
    defaultDf['EndDate'] = startDate
    defaultDf['StartDate'] = endDate
    del workLog

    #Calculating Allocation and Utilization
    defaultDf['Time spent'] = defaultDf['Time spent'].div(80-defaultDf.Leaves).round(4)
    defaultDf['Estimate'] = defaultDf['Estimate'].div(80-defaultDf.Leaves).round(4)
    defaultDf.replace({'Time spent': np.inf}, 0, inplace=True)
    defaultDf.replace({'Estimate': np.inf}, 0, inplace=True)
    defaultDf.rename(columns = {'Estimate':'Allocation', 'Time spent': 'Utilization'}, inplace=True)

    #Sanity Checks
    print("No leaves for: ", len(defaultDf[defaultDf['Leaves'] == 0.00]))
    print("No name matched for: ", len(defaultDf[defaultDf['matches_x'] == 0]))

    #Department wise segregation
    departments = {}
    departments['BA'] = defaultDf[['Assignee name', 'Allocation', 'Utilization']][defaultDf['Department'] == 'BA']
    departments['BI Engineer'] = defaultDf[['Assignee name', 'Allocation', 'Utilization']][defaultDf['Department'] == 'BI Engineer']
    departments['Data Engineer'] = defaultDf[['Assignee name', 'Allocation', 'Utilization']][defaultDf['Department'] == 'Data Engineer']
    departments['Dev Ops'] = defaultDf[['Assignee name', 'Allocation', 'Utilization']][defaultDf['Department'] == 'DevOps']
    departments['Project Manager'] = defaultDf[['Assignee name', 'Allocation', 'Utilization']][defaultDf['Department'] == 'PM']
    departments['SQA'] = defaultDf[['Assignee name', 'Allocation', 'Utilization']][defaultDf['Department'] == 'SQA']
    departments['Support'] = defaultDf[['Assignee name', 'Allocation', 'Utilization']][defaultDf['Department'] == 'Support']
    departments['Web Dev'] = defaultDf[['Assignee name', 'Allocation', 'Utilization']][defaultDf['Department'] == 'Web Dev']
    groups = defaultDf.groupby(['Department']).mean()
    groups.drop(['Leaves', 'Schedule Date IN'], axis=1, inplace=True)
    groups.reset_index(inplace=True)
    for team in departments:
        departments[team].replace({'Allocation': 0}, 'Estimates Missing', inplace=True)
        departments[team].replace({'Utilization': 0}, 'Estimates Missing', inplace=True)

    #Excel engine for report generation
    
    #Setting filename and dataframes
    filename = 'Delivery Report ' + str(datetime.strftime(startDate, "%d%b%Y")) + '.xlsx'
    df = pd.DataFrame()
    writer = pd.ExcelWriter('.\Output\{}'.format(filename), engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Delivery Report')
    jiraData.to_excel(writer, index=False, sheet_name='Jira Data')
    workbook = writer.book
    worksheet = writer.sheets['Delivery Report']
    jiraSheet = writer.sheets['Jira Data']

    #Preparing formats
    number_rows = len(defaultDf)
    indent_fmt = workbook.add_format()
    indent_fmt.set_indent(1)
    total_fmt = workbook.add_format({'bg_color': '#d2fcf9', 'bold': True, 'border' : 1, 'num_format': '0.00%'})
    subheader_fmt = workbook.add_format({'bg_color' : '#7030a0', 'bold': True, 'border' : 1, 'font_color': 'white'})
    avg_fmt = workbook.add_format({'bg_color' : '#7030a0', 'bold': True, 'border' : 1, 'font_color': 'white', 'align': 'center'})
    merge_format_header = workbook.add_format({'align': 'center', 'font_size': 22, 'bg_color' : '#7030a0', 'bold': True, 'border' : 1, 'font_color': 'white'})
    hour_fmt = workbook.add_format({'bg_color': 'yellow'})
    merge_format_team = workbook.add_format({'bold': True, 'border' : 1})
    newline_fmt = workbook.add_format({'text_wrap': True})
    # Add a format. Light red fill with dark red text.
    badRed = workbook.add_format({'bg_color': '#ff0000', 'font_color': 'white', 'border': 1, 'num_format': '0.00%'})#FFC7CE
    # Add a format. Green fill with dark green text.
    goodGreen = workbook.add_format({'bg_color': '#92d050', 'font_color': 'black', 'border': 1, 'num_format': '0.00%'})#C6EFCE
    # Add a format. Yellow fill with dark yellow text.
    neutralYellow = workbook.add_format({'bg_color': '#ffff00', 'font_color': 'black', 'border': 1, 'num_format': '0.00%'})#ffeb9c
    # Add a fromat. Blue fill with white text.
    missingBlue = workbook.add_format({'bg_color': '#333f4f', 'font_color': 'white', 'border': 1})
    # Add a format. Grey fill with black text
    overGrey = workbook.add_format({'bg_color': '#aeaaaa', 'font_color': 'black', 'border': 1, 'num_format': '0.00%'})
    # General format
    general_fmt = workbook.add_format({'border': 1, 'num_format': '0.00%'})
    # General format border
    general_border_fmt = workbook.add_format({'border': 1})
    bottom_border_fmt = workbook.add_format()
    bottom_border_fmt.set_bottom(1)
    bottom_border_fmt.set_indent(1)

    worksheet.write('A3', 'Resource', subheader_fmt)
    worksheet.write('B3', 'Allocation', subheader_fmt)
    worksheet.write('C3', 'Utilization', subheader_fmt)
    worksheet.write('E4', 'Team', subheader_fmt)
    worksheet.write('F4', 'Allocation', subheader_fmt)
    worksheet.write('G4', 'Utilization', subheader_fmt)
    worksheet.merge_range('E17:G17', 'No. of leaves b/w ' + str(datetime.strftime(startDate, "%b %d")) + ' - ' + str(datetime.strftime(endDate, "%b %d")), avg_fmt)
    worksheet.write('E18', 'Name', subheader_fmt)
    worksheet.merge_range('F18:G18', 'No. of Leaves', subheader_fmt)

    #Writing dataframes and formats
    worksheet.merge_range('A1:G1', 'Allocation & Utilization: ' + str(datetime.strftime(startDate, "%b %d")) + ' - ' + str(datetime.strftime(endDate, "%b %d, %Y")), merge_format_header)
    currentRow = 4
    for team in departments:
        worksheet.merge_range('A{}:C{}'.format(currentRow, currentRow), team , merge_format_team)
        departments[team].to_excel(writer, index=False, header=False, sheet_name='Delivery Report', startrow=currentRow)
        allocation_range = "B{}:B{}".format(currentRow+1, len(departments[team])+currentRow)
        utilization_range = "C{}:C{}".format(currentRow+1, len(departments[team])+currentRow)
        team_range = "A{}:A{}".format(currentRow+1, len(departments[team])+currentRow)
        colorCells(worksheet, allocation_range, badRed, goodGreen, neutralYellow, overGrey, missingBlue)
        colorCells(worksheet, utilization_range, badRed, goodGreen, neutralYellow, overGrey, missingBlue)
        generalFormat(indent_fmt, team_range, worksheet)
        currentRow = currentRow + len(departments[team])+1
    generalFormat(bottom_border_fmt, "A{}:A{}".format(currentRow-1, currentRow-1), worksheet)
    generalFormat(hour_fmt, "D1:D1", jiraSheet)
    generalFormat(hour_fmt, "F1:F1", jiraSheet)

    worksheet.merge_range('E3:G3', 'Average', avg_fmt)
    allocation_total = defaultDf['Allocation'].mean()
    utilization_total = defaultDf['Utilization'].mean()
    allocation_range = "F5:F{}".format(len(groups)+4)
    utilization_range = "G5:G{}".format(len(groups)+4)
    team_range = "E5:E{}".format(len(groups)+4)
    generalFormat(general_fmt, utilization_range, worksheet)
    generalFormat(general_fmt, allocation_range, worksheet)
    generalFormat(merge_format_team, team_range, worksheet)
    worksheet.write('E{}'.format(len(groups)+5), 'Total', total_fmt)
    worksheet.write('F{}'.format(len(groups)+5), allocation_total, total_fmt)
    worksheet.write('G{}'.format(len(groups)+5), utilization_total, total_fmt)
    groups.to_excel(writer, index=False, header=False, sheet_name='Delivery Report', startrow=4, startcol=4)

    currentRow = 19

    if len(gazettedHoliday) > 0:    
        combineString = ""
        for x in gazettedHoliday:
            combineString = combineString + x + "\n "
        worksheet.write("E{}".format(currentRow), "Gazetted Holiday", merge_format_team)
    #     worksheet.write("F{}".format(currentRow), combineString, newline_fmt)
        worksheet.merge_range("F{}:G{}".format(currentRow, currentRow), len(gazettedHoliday), general_border_fmt)
        currentRow = currentRow + 1

    for employee in leavesDict:
        leaveDates = leavesDict[employee]['Schedule Date IN'].tolist()
        worksheet.write("E{}".format(currentRow), employee, merge_format_team)
        combineString = ""
        for x in leaveDates:
            combineString = combineString + x + "\n "    
    #     worksheet.write("F{}".format(currentRow), combineString, newline_fmt)
        worksheet.merge_range("F{}:G{}".format(currentRow, currentRow), len(leaveDates), general_border_fmt)
        currentRow = currentRow + 1

    currentRow = currentRow + 3
    worksheet.merge_range('E{}:F{}'.format(currentRow, currentRow), 'Allocation/Utilization - Legend', subheader_fmt)
    worksheet.write('G{}'.format(currentRow), 'Range(%)', subheader_fmt)
    currentRow = currentRow + 1
    worksheet.merge_range('E{}:F{}'.format(currentRow, currentRow), 'Under Allocated/Utilized', badRed)
    worksheet.write('G{}'.format(currentRow), '<50', general_border_fmt)
    currentRow = currentRow + 1
    worksheet.merge_range('E{}:F{}'.format(currentRow, currentRow), 'Average Allocated/Utilized', neutralYellow)
    worksheet.write('G{}'.format(currentRow), '50-80', general_border_fmt)
    currentRow = currentRow + 1
    worksheet.merge_range('E{}:F{}'.format(currentRow, currentRow), 'Balanced Allocated/Utilized', goodGreen)
    worksheet.write('G{}'.format(currentRow), '81-100', general_border_fmt)
    currentRow = currentRow + 1
    worksheet.merge_range('E{}:F{}'.format(currentRow, currentRow), 'Over Allocated/Utilized', overGrey)
    worksheet.write('G{}'.format(currentRow), '>100', general_border_fmt)
    currentRow = currentRow + 1
    worksheet.merge_range('E{}:F{}'.format(currentRow, currentRow), 'Estimates Missing/On Leaves', missingBlue)
    worksheet.write('G{}'.format(currentRow), '0', general_border_fmt)
    writer.save()

    #Print filename
    print(filename)

    #Python windows api to expand the sheet to adjust to the text
    #excel = win32.gencache.EnsureDispatch('Excel.Application')
    #wb = excel.Workbooks.Open('D:\My Projects\Jira Connector\Jira Connector Application\{}'.format(filename))
    #ws = wb.Worksheets("Delivery Report")
    #ws.Columns.AutoFit()
    #ws = wb.Worksheets("Jira Data")
    #ws.Columns.AutoFit()
    #wb.Save()
    #excel.Application.Quit()

    #Database schema
    #CREATE TABLE [dbo].[DeliveryReport](
	   # [ID] [int] IDENTITY(1,1) NOT NULL,
	   # [Team] [varchar](20) NOT NULL,
	   # [Name] [varchar](50) NOT NULL,
	   # [Allocation] [float] NULL,
	   # [Utilization] [float] NULL,
	   # [StartDate] [date] NULL,
	   # [EndDate] [date] NULL,
    #PRIMARY KEY CLUSTERED 
    #(
	   # [ID] ASC
    #)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
    #) ON [PRIMARY]
    #GO

    #Database entry for PowerBI dashboards
    #defaultDf[['Allocation', 'Utilization']] = defaultDf[['Allocation', 'Utilization']].mul(100)
    #server = 'PROD-LPT-69\SQL' 
    #database = 'Cost' 
    #conn = pyodbc.connect("DRIVER={ODBC Driver 17 for SQL Server};SERVER="+server+';DATABASE='+database+';Trusted_Connection=yes;')
    #cursor = conn.cursor()
    #cursor.execute("TRUNCATE TABLE DeliveryReport")

    #for index, row in defaultDf.iterrows():
    #    cursor.execute("INSERT INTO DeliveryReport (Team, Name, Allocation, Utilization, StartDate, EndDate) values(?,?,?,?,?,?);", 
    #                   str(row.Department), str(row['Assignee name']), str(row.Allocation), str(row.Utilization), row.StartDate, row.EndDate)
    #conn.commit()
    #cursor.close()

    #Preparing multiple traces for default graphs
    workLogAlpha = defaultDf[['Assignee name', 'Department', 'Leaves', 'Allocation', 'Utilization']]
    workLogAlpha[['Allocation', 'Utilization']] = workLogAlpha[['Allocation', 'Utilization']].mul(100).round(2)
    workLogAlpha['Leaves'] = workLogAlpha['Leaves'].div(8)
    workLogAlpha.rename(columns = {'Assignee name': 'Assignee', 'Department': 'Team', 'Allocation': 'Allocation (%)', 'Utilization': 'Utilization (%)'}, inplace=True)
    workLogAlpha['Assignee'] = workLogAlpha['Assignee'].str.replace('Muhammad ', '') #just replacing too many similar sir names 
    jiraData = jiraData[['Issue key', 'Summary', 'Assignee name', 'Estimate [Hours]', 'Time spent [Hours]', 'Logged Date Time']]
    jiraData.rename(columns = {'Estimate [Hours]': 'Estimate (hrs)', 'Time spent [Hours]': 'Time spent (hrs)'}, inplace = True)

    #Main overall bar graph
    trace1 = go.Bar(
                    x = workLogAlpha['Assignee'],
                    y = workLogAlpha['Allocation (%)'],
                    marker = {'color': 'rgb(199, 153, 185)'},
                    name = 'Estimate'
             )
    trace2 = go.Bar(
                    x = workLogAlpha['Assignee'],
                    y = workLogAlpha['Utilization (%)'],
                    marker = {'color': 'darkviolet'},
                    name = 'Time spent'
             )
    figure = {
            'data': [trace1, trace2],
            'layout': go.Layout(
                title = 'Overall Time Spent',
                xaxis = {'title': 'Assignee'},
                yaxis = {'title': 'Utilization'},
                hovermode='closest'
            )
        }

    #Teamwise segregated pie chart
    team = groups.Department.tolist()
    groups[['Allocation', 'Utilization']] = groups[['Allocation', 'Utilization']].mul(100).round(2)
    groups.rename(columns = {'Allocation': 'Allocation (%)', 'Utilization': 'Utilization (%)', 'Department': 'Team'}, inplace = True)
    utilize = groups['Utilization (%)'].tolist()
    trace3 = go.Pie(labels=team,
                    values=utilize
            )
    pie_figure = {
            'data': [trace3],
            'layout': go.Layout(
                title = 'Team Segregation',
                hovermode='closest'
            )
        }

    #groups.rename(columns = {'Department': 'Team'}, inplace = True)
    groups.reset_index(inplace=True)

    #Preparing options (list of usernames in the drop down)
    for name in workLogAlpha.Assignee:
        nameDict = {}
        nameDict['label'] = name
        nameDict['value'] = name
        option.append(nameDict)
    
    print('{} tickets retrieved in {} seconds'.format(len(jiraData), round(executionTime, 2)))
    print('Refreshed on: ', datetime.today())

    #Returning everything prepared so far to the dashboard
    return [jiraData.to_dict('records'), html.Div(style={'display': 'None'}), figure, option, pie_figure, groups.to_dict('rows'), '{} tickets retrieved in {} seconds'.format(len(jiraData), round(executionTime, 2))]


if __name__ == '__main__':
    app.run_server(debug=True)


# Current known issue is the leaves data not matching the entered date