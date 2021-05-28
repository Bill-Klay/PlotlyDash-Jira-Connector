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

warnings.filterwarnings('ignore')

app = dash.Dash(__name__)

jiraData = pd.DataFrame()
workLog = pd.DataFrame()
workLogAlpha = pd.DataFrame()
dayCount = 0

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

    html.Div([html.Button('Switch Table', id='convert')], style={'marginLeft': '20px'}),
    
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
    ])

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

@app.callback([
    Output('dashboard', 'data'),
    Output('loading', 'children'),
    Input('daysSlider', 'value')
])
def jiraConnector(days):
    global jiraData, workLog, workLogAlpha, dayCount

    if days == dayCount:
        raise PreventUpdate
    else:
        dayCount = days

    initial = 0
    size = 100

    jira = JIRA(options={'server': 'https://qordatainc.atlassian.net/'}, basic_auth=('bilal.khan@qordata.com', 'yaeSBwTTcPI3CdpPUG2G2D28'))
    jql='worklogDate >= -'+str(days)+'d'

    # Container for Jira's data
    data_jira = []
    work_log = []

    while True:
        start = initial * size
        jira_search = jira.search_issues(jql, startAt=start, maxResults=size, 
                                         fields = "key, summary, issuetype, assignee, reporter, status, created, resolutiondate, workratio, timespent, timeoriginalestimate")
        if(len(jira_search) == 0): 
            break

        # Iterate over the issues
        for issue in jira_search:
            # Get issue key
            issue_key = issue.key

            # Get issue summary
            issue_summary = issue.fields.summary

            # Get request type
            request_type = str(issue.fields.issuetype)

            # Get datetime creation
            datetime_creation = issue.fields.created
            if datetime_creation is not None:
                # Interested in only seconds precision, so slice unnecessary part
                datetime_creation = datetime.strptime(datetime_creation[:19], "%Y-%m-%dT%H:%M:%S")

            # Get datetime resolution
            datetime_resolution = issue.fields.resolutiondate
            if datetime_resolution is not None:
                # Interested in only seconds precision, so slice unnecessary part
                datetime_resolution = datetime.strptime(datetime_resolution[:19], "%Y-%m-%dT%H:%M:%S")

            # Get reporter’s login and name
            reporter_login = None
            reporter_name = None
            reporter = issue.raw['fields'].get('reporter', None)
            if reporter is not None:
                reporter_login = reporter.get('key', None)
                reporter_name = reporter.get('displayName', None)

            # Get assignee’s login and name
            assignee_login = None
            assignee_name = None
            assignee = issue.raw['fields'].get('assignee', None)
            if assignee is not None:
                assignee_login = assignee.get('key', None)
                assignee_name = assignee.get('displayName', None)

            # Get status
            status = None
            st = issue.fields.status
            if st is not None:
                status = st.name

            # Time logging
            issue_workratio = issue.fields.workratio
            issue_timespent = issue.fields.timespent
            issue_estimate = issue.fields.timeoriginalestimate

            # Add data to data frame
            data_jira.append((issue_key, issue_summary, request_type, datetime_creation, datetime_resolution, reporter_login, reporter_name, assignee_login, assignee_name, status, issue_workratio, issue_timespent, issue_estimate))
            work_log.append((issue_key, issue_summary, assignee_name, issue_estimate, issue_timespent))

        initial = initial + 1
        print(len(data_jira))

    jiraData = pd.DataFrame(data_jira, columns=['Issue key', 'Summary', 'Request type', 'Datetime creation', 'Datetime resolution', 'Reporter login', 'Reporter name', 'Assignee login', 'Assignee name', 'Status', 'Work raito', 'Time spent', 'Estimate'])
    workLog = pd.DataFrame(work_log, columns=['Issue key', 'Summary', 'Assignee', 'Estimate', 'Time spent'])
    
    jiraData.to_excel('Raw Data.xlsx', header=True, index=True)
    workLog.to_excel('Work Log.xlsx', header=True, index=True)
    jira.close()

    workLogAlpha = workLog.groupby(['Assignee']).sum()
    workLogAlpha[['Time spent', 'Estimate']] = workLogAlpha[['Time spent', 'Estimate']].div(3600) 
    workLogAlpha['Utilization'] = workLogAlpha['Time spent'].div(80).round(4)    
    workLogAlpha.reset_index(inplace=True)
    workLogAlpha['Assignee'] = workLogAlpha['Assignee'].str.replace('.', ' ')
    workLogAlpha['Assignee'] = workLogAlpha['Assignee'].str.capitalize()
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
    workLogAlpha['Utilization'] = workLogAlpha['Utilization'].astype(str) + '%'
    workLogAlpha['Estimate'] = workLogAlpha['Estimate'].round(2)
    workLogAlpha['Time spent'] = workLogAlpha['Time spent'].round(2)
    workLog['Assignee'] = workLog['Assignee'].str.replace('.', ' ')
    workLog['Assignee'] = workLog['Assignee'].str.capitalize()
    #workLog[['Time spent', 'Estimate']] = workLog[['Time spent', 'Estimate']].div(3600)

    return [workLog.to_dict('records'), html.Div(style={'display': 'None'})]


if __name__ == '__main__':
    app.run_server(debug=True)

    