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

warnings.filterwarnings('ignore')

app = dash.Dash(__name__)

jiraData = pd.DataFrame()
workLog = pd.DataFrame()
workLogAlpha = pd.DataFrame()
dayCount = 0
option = []

app.layout = html.Div([
    
    html.Div([html.Span('Jira Connector', className='circle-sketch-highlight', style={'font-size': '2.5em', 'font-family':'Script MT', 'marginLeft':'20px'}),
              html.Span(dcc.Loading(children=[html.Div(id='loading')], type='circle'), style={'marginLeft': '10%'})
    ]),
    html.Span('Review team worklog', style={'marginLeft': '20px', 'font-family':'Script MT'}),
    
    html.Hr(),

    html.Div(id='days', style={'font-size': '1.2em', 'font-family':'Calibri', 'marginLeft':'20px'}),

    dcc.Slider(
        id='daysSlider',
        min=0,
        max=30,
        value=14,
        step=1,
        marks={
            0: {'label':'0 Days', 'style':{'color': 'purple', 'font-face': 'bold'}},
            5: {'label': '5 Days'},
            10: {'label': '10 Days'},
            15: {'label': '15 Days'},
            20: {'label':'20 Days'},
            25: {'label':'25 Days'},
            30: {'label':'30 Days', 'style':{'color': 'purple', 'font-face': 'bold'}}
        }
    ),

    html.Br(),

    html.Div(children=[html.Button('Switch Table', id='convert'), html.Span(id='ticketCount', style={'font-size': '1.2em', 'font-family':'Calibri', 'marginLeft':'20px'})], 
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
                     {'name': 'Assignee', 'id': 'Assignee'},
                     {'name': 'Estimate', 'id': 'Estimate'},
                     {'name': 'Time spent', 'id': 'Time spent'}
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

    html.Div([html.Span([html.Button("Download Raw File", id="rawFile"), dcc.Download(id="downloadRawFile")], style={'marginLeft': '20px'}),
              html.Span([html.Button("Download Log File", id="logFile"), dcc.Download(id="downloadLogFile")], style={'marginLeft': '10%'}),
              html.Span([html.Button("Download Summary", id="summaryFile"), dcc.Download(id="downloadSummaryFile")], style={'marginLeft': '10%'})
    ]),

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
            multi = False
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
                                    {'name': 'Estimate', 'id': 'Estimate'},
                                    {'name': 'Utilization', 'id': 'Utilization'}
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

@app.callback([
    Output('days', 'children'),
    Input('daysSlider', 'value')
])
def sliderUpdate(days):
    return ['{} Days'.format(days)]

@app.callback([
    Output('downloadRawFile', 'data'),
    Input('rawFile', 'n_clicks')
])
def fileDownload(click):
    if click is None:
        raise PreventUpdate
    else:
        return [dcc.send_data_frame(jiraData.to_excel, "Jira Data.xlsx")]

@app.callback([
    Output('downloadLogFile', 'data'),
    Input('logFile', 'n_clicks')
])
def fileDownload(click):
    if click is None:
        raise PreventUpdate
    else:
        return [dcc.send_data_frame(workLog.to_excel, "Work Log.xlsx")]

@app.callback([
    Output('downloadSummaryFile', 'data'),
    Input('summaryFile', 'n_clicks')
])
def fileDownload(click):
    if click is None:
        raise PreventUpdate
    else:
        return [dcc.send_file(".\Summary.xlsx")]

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
                data=workLog.to_dict('records'),
                columns=[{'name': i, 'id': i} for i in workLog.columns],
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

@app.callback(Output('graphsDiv', 'children'),
              Input('names', 'value'),
              Input('graphs', 'value'),
              prevent_initial_call=True
)
def selectName (names, graph):
    global workLogAlpha

    mode = 'lines+markers'
    marker = {'size': 12, 'symbol': 'circle-dot', 'color': 'darkviolet'}

    trace1 = go.Bar(
                    x = workLogAlpha['Assignee'],
                    y = workLogAlpha['Estimate'],
                    marker = {'color': 'rgb(199, 153, 185)'},
                    name = 'Estimate'
    )
    trace2 = go.Bar(
                    x = workLogAlpha['Assignee'],
                    y = workLogAlpha['Time spent'],
                    marker = {'color': 'darkviolet'},
                    name = 'Time spent'
    )
    data = [trace1, trace2]
    
    if graph == 'Bubble':
        mode = 'markers'
        marker = dict(size=3*workLogAlpha['Time spent'])
        data = [
            go.Scatter(
                x = workLogAlpha['Assignee'],
                y = workLogAlpha['Time spent'],
                mode = mode,
                marker = marker
            )
        ]
    elif graph == 'Bar':
        data = [
            go.Bar(
                x = workLogAlpha['Assignee'],
                y = workLogAlpha['Time spent'],
                marker = {'color': 'darkviolet'}
            )
        ]
    elif graph == 'Heatmap':
        data = [
            go.Heatmap(
            x = workLogAlpha['Assignee'],
            y = workLogAlpha['Time spent'],
            z = workLogAlpha['Time spent'],
            colorscale = 'Jet',
            zmin = 5, zmax = 40 # add max/min color values to make each plot consistent
            )
        ] 
    elif graph == 'Line':
        data = [
            go.Scatter(
                x = workLogAlpha['Assignee'],
                y = workLogAlpha['Time spent'],
                mode = mode,
                marker = marker
            )]

    figure = {
            'data': data,
            'layout': go.Layout(
                title = 'Overall Time Spent',
                xaxis = {'title': 'Assignee'},
                yaxis = {'title': 'Utilization'},
                hovermode='closest'
            )
        }

    if names is not None:
        if len(names):
            # traces = []
            data.clear()
            
            for name in names:
                tempNameList = workLogAlpha[workLogAlpha.Assignee == name]
                if graph == 'Bubble':
                    mode = 'markers'
                    marker = dict(size=3*tempNameList['Time spent'])
                    data.append(go.Scatter(
                        x = tempNameList.Assignee,
                        y = tempNameList['Time spent'],
                        mode = mode,
                        name = name,
                        marker = marker
                    ))
                elif graph == 'Bar':
                    data.append(go.Bar(
                        x = tempNameList.Assignee,
                        y = tempNameList['Time spent'],
                        name = name
                    ))
                elif graph == 'Heatmap':
                    data.append(go.Heatmap(
                        x=tempNameList.Assignee,
                        y=tempNameList['Time spent'],
                        z=tempNameList['Time spent'],
                        colorscale='Jet',
                        zmin = 40, zmax = 100,
                        name = name
                    ))
                elif graph == 'Line':
                    data.append(go.Scatter(
                        x = tempNameList.Assignee,
                        y = tempNameList['Time spent'],
                        name = name,
                        mode = mode,
                        marker = marker
                    ))
                else:
                    data.append(go.Bar(
                        x = tempNameList.Assignee,
                        y = tempNameList['Time spent'],
                        name = name,
                    ))
            fig = {'data': data, 'layout': go.Layout(title = 'Overall Time Spent', xaxis = {'title': 'Assignee'}, yaxis = {'title': 'Utilization'}, hovermode='closest')}
            return [dcc.Graph(id='overall-graph', figure=fig)]
        else:
            return [dcc.Graph(id='overall-graph', figure=figure)]
    else:
        return [dcc.Graph(id='overall-graph', figure=figure)]

@app.callback([
    Output('dashboard', 'data'),
    Output('loading', 'children'),
    Output('overall-graph', 'figure'),
    Output('names', 'options'),
    Output('grouping-graph', 'figure'),
    Output('teamTable', 'data'),
    Output('ticketCount', 'children'),
    Input('daysSlider', 'value')
])
def jiraConnector(days):
    global jiraData, workLog, workLogAlpha, dayCount, option

    if days == dayCount:
        raise PreventUpdate
    else:
        dayCount = days

    initial = 0
    size = 100

    jira = JIRA(options={'server': 'https://qordatainc.atlassian.net/'}, basic_auth=('bilal.khan@qordata.com', 'yaeSBwTTcPI3CdpPUG2G2D28'))
    jql='worklogDate >= -'+str(days)+'d'

    data_jira = []
    work_log = []
    option.clear()
    defaultDf = pd.read_excel('Members.xlsx')
    defaultDf['Assignee'] = defaultDf['Assignee'].str.replace('Muhammad ', '')
    defaultDf['Assignee'] = defaultDf['Assignee'].str.title()

    while True:
        start = initial * size
        jira_search = jira.search_issues(jql, startAt=start, maxResults=size, 
                                         fields = "key, summary, issuetype, assignee, reporter, status, created, resolutiondate, workratio, timespent, timeoriginalestimate")
        if(len(jira_search) == 0): 
            break

        for issue in jira_search:
            issue_key = issue.key

            issue_summary = issue.fields.summary

            request_type = str(issue.fields.issuetype)

            datetime_creation = issue.fields.created
            if datetime_creation is not None:
                datetime_creation = datetime.strptime(datetime_creation[:19], "%Y-%m-%dT%H:%M:%S")

            datetime_resolution = issue.fields.resolutiondate
            if datetime_resolution is not None:
                datetime_resolution = datetime.strptime(datetime_resolution[:19], "%Y-%m-%dT%H:%M:%S")

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

            issue_workratio = issue.fields.workratio
            issue_timespent = issue.fields.timespent
            issue_estimate = issue.fields.timeoriginalestimate

            data_jira.append((issue_key, issue_summary, request_type, datetime_creation, datetime_resolution, reporter_login, reporter_name, assignee_login, assignee_name, status, issue_workratio, issue_timespent, issue_estimate))
            work_log.append((issue_key, issue_summary, assignee_name, issue_estimate, issue_timespent))

        initial = initial + 1

    jiraData = pd.DataFrame(data_jira, columns=['Issue key', 'Summary', 'Request type', 'Datetime creation', 'Datetime resolution', 'Reporter login', 'Reporter name', 'Assignee login', 'Assignee name', 'Status', 'Work raito', 'Time spent', 'Estimate'])
    workLog = pd.DataFrame(work_log, columns=['Issue key', 'Summary', 'Assignee', 'Estimate', 'Time spent'])
    
    jiraData.to_excel('Raw Data.xlsx', header=True, index=True)
    workLog.to_excel('Work Log.xlsx', header=True, index=True)
    jira.close()

    workLogAlpha = workLog.groupby(['Assignee']).sum()
    workLogAlpha[['Time spent', 'Estimate']] = workLogAlpha[['Time spent', 'Estimate']].div(3600) 
    workLogAlpha['Utilization'] = workLogAlpha['Time spent'].div(80).round(4)
    workLogAlpha['Estimate'] = workLogAlpha['Estimate'].div(80).round(4)
    workLogAlpha.reset_index(inplace=True)
    workLogAlpha['Assignee'] = workLogAlpha['Assignee'].str.replace('.', ' ')
    writer = pd.ExcelWriter('Summary.xlsx', engine='xlsxwriter')
    workLogAlpha.to_excel(writer, index=False, sheet_name='Report')
    workbook = writer.book
    worksheet = writer.sheets['Report']
    number_rows = len(workLogAlpha)
    percent_fmt = workbook.add_format({'num_format': '0.00%', 'bold': True})
    color_range = "D2:D{}".format(number_rows+1)
    badRed = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    goodGreen = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    neutralYellow = workbook.add_format({'bg_color': '#ffeb9c', 'font_color': '#9c6500'})
    worksheet.conditional_format(color_range, {'type': 'cell', 'criteria': '<=', 'value': 0.50, 'format': badRed})
    worksheet.conditional_format(color_range, {'type': 'cell', 'criteria': '>=', 'value': 0.80, 'format': goodGreen})
    worksheet.conditional_format(color_range, {'type': 'cell', 'criteria': 'between', 'minimum': 0.51, 'maximum': 0.79, 'format': neutralYellow})
    worksheet.set_column(color_range, number_rows+1, percent_fmt)
    writer.save()
    
    workLogAlpha['Utilization'] = workLogAlpha['Utilization'].mul(100).round(2)
    #workLogAlpha['Utilization'] = workLogAlpha['Utilization'].astype(str) + '%'
    workLogAlpha['Estimate'] = workLogAlpha['Estimate'].mul(100).round(2)
    #workLogAlpha['Estimate'] = workLogAlpha['Estimate'].astype(str) + '%'
    workLogAlpha['Time spent'] = workLogAlpha['Time spent'].round(2)
    workLogAlpha['Assignee'] = workLogAlpha['Assignee'].str.title()
    workLogAlpha['Assignee'] = workLogAlpha['Assignee'].str.replace('Muhammad ', '')
    defaultDf = workLogAlpha.set_index('Assignee').combine_first(defaultDf.drop_duplicates().set_index('Assignee')).reset_index()
    defaultDf.drop(['Time spent'], axis=1, inplace=True)
    defaultDf['End Date'] = date.today()
    defaultDf['Start Date'] = date.today() - timedelta(days=int(days))
    defaultDf.to_excel('Database Entry.xlsx')
    workLog['Assignee'] = workLog['Assignee'].str.replace('.', ' ')
    workLog['Assignee'] = workLog['Assignee'].str.title()

    server = 'PROD-LPT-69' 
    database = 'Cost' 
    conn = pyodbc.connect("DRIVER={ODBC Driver 17 for SQL Server};SERVER="+server+';DATABASE='+database+';Trusted_Connection=yes;')
    cursor = conn.cursor()
    cursor.execute("TRUNCATE TABLE DeliveryTeamReportV2")
    for index, row in defaultDf.iterrows():
        cursor.execute("INSERT INTO DeliveryTeamReportV2 (Team, Name, Allocation, Utilization, StartDate, EndDate) values(?,?,?,?,?,?);", 
                       str(row.Team), str(row.Assignee), float(row['Estimate']), float(row['Utilization']), row['Start Date'], row['End Date'])
    conn.commit()
    cursor.close()

    trace1 = go.Bar(
                    x = workLogAlpha['Assignee'],
                    y = workLogAlpha['Estimate'],
                    marker = {'color': 'rgb(199, 153, 185)'},
                    name = 'Estimate'
             )
    trace2 = go.Bar(
                    x = workLogAlpha['Assignee'],
                    y = workLogAlpha['Time spent'],
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
    groupDf = defaultDf.groupby(['Team']).sum()
    team = groupDf.index.tolist()
    utilize = groupDf.Utilization.tolist()
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
    groupDf['Utilization'] = groupDf['Utilization'].round(2)
    groupDf['Estimate'] = groupDf['Estimate'].round(2)
    groupDf.reset_index(inplace=True)

    for name in workLogAlpha.Assignee:
        nameDict = {}
        nameDict['label'] = name
        nameDict['value'] = name
        option.append(nameDict)

    return [workLog.to_dict('records'), html.Div(style={'display': 'None'}), figure, option, pie_figure, groupDf.to_dict('rows'), '{} tickets retrieved'.format(len(workLog))]


if __name__ == '__main__':
    app.run_server(debug=True)

    