<!-- PROJECT LOGO -->
<br />
<p align="center">
  <a href="https://github.com/Bill-Klay/PlotlyDash-Jira-Connector">
    <img src="./images/Logo.png" alt="Logo" width="247" height="92">
  </a>

  <h3 align="center">Jira Connector</h3>

  <p align="center">
    Review your teams Jira work log with a python dashboard
    <br />
    <a href="https://github.com/Bill-Klay/PlotlyDash-Jira-Connector"><strong>Explore the docs »</strong></a>
    <br />
    <a href="https://github.com/Bill-Klay/PlotlyDash-Jira-Connector/issues">Report Bug</a>
    ·
    <a href="https://github.com/Bill-Klay/PlotlyDash-Jira-Connector/issues">Request Feature</a>
  </p>
</p>

***

<!-- TABLE OF CONTENTS -->
<details open="open">
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#roadmap">Roadmap</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>
    <li><a href="#acknowledgements">Acknowledgements</a></li>
  </ol>
</details>



<!-- ABOUT THE PROJECT -->
## About The Project :book:

![Screenshot2](https://github.com/Bill-Klay/PlotlyDash-Jira-Connector/blob/master/images/Screenshot2.png)

This project was made as an inhouse improvement for making a publicly viewed dashboard that shows team engagements and what is being curated inside the production environment. 
It uses **Python** as the base environment making use of the **plotly dash framework** for hosting plotly enabled dashboards. 
The dashboard provides users with acces to donwloading all the tickets or specified tickets from **Jira Cloud** based on the entered JQL. 
The dashboards presents with an interactive data table for viewing the results from the **JQL**, plus a summary of the **Estimates and Utilization** of the work logged. 
An overall graphs shows all the users with there estimates and utilization in a **bar graph** with more graphs and individual person view available from the drop down. 
A **teamwise segregation** is also availabe at the end for which there first needs to be an excel present with all the team members name and there department. 

### Built With :computer:

This project uses Python, more specifically it uses the Plotly/Dash Framework for making web apps using integrated plotly graphs.
* [Plotly Dash](https://plotly.com/dash/)
* [Python](https://www.python.org/)
* [Visual Studio 2019](https://visualstudio.microsoft.com/vs/)
<!-- GETTING STARTED -->
## Getting Started :construction_worker:

The *ReadMe* documents provides users with a basic introduction to setting up the requirements and get the project running and how improvements can be added. 

### Prerequisites

Things you need to know and have before executing the project on your local environment.
* Python (This project was made and tested on **Python 3.9**, theoratically this should work on anything greater than Python 2.*).
* Have a basic understanding of Python's Plotly Dash framework.
* Python Jira library
* Python Pandas library
* Basic level HTML and CSS
* A Jira cloud account with an API token

### Installation

1. Install and setup [Python](https://www.python.org/) environment 
2. Setup your Python code with a connection to your organizations Jira cloud. Log in to your Jira account and get a free API Key/token at [Atlassian Support](https://support.atlassian.com/atlassian-account/docs/manage-api-tokens-for-your-atlassian-account/)
3. Install all the dependencies from the reqiurements file provided
   ```sh
   pip install -r requirements.txt
   ```

<!-- USAGE EXAMPLES -->
## Usage ![build](https://img.shields.io/badge/Build-Tested-green) ![build mode](https://img.shields.io/badge/Build-Debug%20Mode-orange)

The source code is available for adusting to your own need, besides that to use as is remeber to enter your own Atlassian email id and the API key obtained from the Jira cloud account. 
*Replace the holders with your own email and token*. 
After that the *JQL* needs to be customized to obtained results relevant to one's need. 
The code itself will analyse the results and prepare the results from the data retrieved. 
To get a teamwise segregated result an excel needs to be prepared (default name used in the project is *Members.xlsx*) and kept along with the *app.py* file which holds the value in the following format:
* Team
* Assignee
* Estimate
* Utilization
* Start Date
* End Date

Where *Team* is the department of the team member, *Assignee* holds all the names, *Estimate* has the default allocation time (initially to be filled with 0), *Utilization* has the calculated time spent (initially to be filled with 0), and *Start and End Date* are placeholders for the dataframe to be loaded correctly these can be left or regardless of the initial value these will be ovreridden. 
Adjust the sliding date range to your need and the JQL entered will retrieve tickets back to that number of days (default is 2 weeks). 
All the tickets retrieved will be displayed in a **data table**. \
![data table](https://github.com/Bill-Klay/PlotlyDash-Jira-Connector/blob/master/images/Screenshot1.png) \
A **graph** ![dashboard](https://github.com/Bill-Klay/PlotlyDash-Jira-Connector/blob/master/images/Screenshot4.png) underneath the data table will show the data retrived in a bar graph by default, but can be customized from the drop down available and team members names can be individually chosen from another drop down. 
*Initially the dashboard supports Line graph, bar graph, bubble graph and heatmap*. 
If the `Members.xlsx` excel file is created then the distributed **pie chart** \
![pie chart](https://github.com/Bill-Klay/PlotlyDash-Jira-Connector/blob/master/images/Screenshot3.png) \
will be dispalyed.
*Work log adjustment* made the program a lot slower than intended. Ususally 624 tickets were retrieved and processed in under 10 sec with a complexity of O(n), 
but since the retrieving a work log brings in the whole time logged rather than just the work that was logged in the data range provided *jira.worklog* was used to get the worklog list and traverse manually. 
This increased the complexity to O(n*m) taking around 264 sec upto 4-5 mins to retrive and process the average of 620 tickets.\
![flow diagram](https://github.com/Bill-Klay/PlotlyDash-Jira-Connector/blob/master/images/Flow_Diagram.png) \
The flow of the application can be observed in the flow chart described above. The employee list and leaves are a inhouse dependency can be opted out or adjusted accordingly.

<!-- ROADMAP -->
## Roadmap

See the [open issues](https://github.com/Bill-Klay/PlotlyDash-Jira-Connector/issues) for a list of proposed features (and known issues).



<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to be learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

Thanks for visiting :green_heart:

<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE` for more information.

<!-- CONTACT -->
## Contact - Bilal Khan :email:

* [email](mailto:bilal.khan2998@gmail.com) 
* [LinkedIn](https://www.linkedin.com/in/bilalkhan29/)
* [Project Link](https://github.com/Bill-Klay/PlotlyDash-Jira-Connector)

<!-- ACKNOWLEDGEMENTS -->
## Acknowledgements :octocat:
* [GitHub Emoji Cheat Sheet](https://www.webpagefx.com/tools/emoji-cheat-sheet)
* [Img Shields](https://shields.io)
* [Choose an Open Source License](https://choosealicense.com)
* [Plotly Dash](https://plotly.com/dash/)
* [Python](https://www.python.org/)
* [Jira timesheet PDF](https://github.com/jordanjambazov/jira-timesheet-pdf)
