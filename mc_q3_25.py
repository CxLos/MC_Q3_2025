# =================================== IMPORTS ================================= #
import csv, sqlite3
import numpy as np 
import pandas as pd 
import seaborn as sns 
import matplotlib.pyplot as plt 
import plotly.figure_factory as ff
import plotly.graph_objects as go
from geopy.geocoders import Nominatim
from folium.plugins import MousePosition
import plotly.express as px
from datetime import datetime
import folium
import os
import sys
import dash
from dash import dcc, html
from dash.dependencies import Input, Output, State
from dash.development.base_component import Component

# Google Web Credentials
import json
import base64
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# 'data/~$bmhc_data_2024_cleaned.xlsx'
# print('System Version:', sys.version)
# -------------------------------------- DATA ------------------------------------------- #

current_dir = os.getcwd()
current_file = os.path.basename(__file__)
script_dir = os.path.dirname(os.path.abspath(__file__))
# data_path = 'data/MarCom_Responses.xlsx'
# file_path = os.path.join(script_dir, data_path)
# data = pd.read_excel(file_path)
# df = data.copy()

# Define the Google Sheets URL
sheet_url = "https://docs.google.com/spreadsheets/d/1EFbKxXM_qBrD6PkxoYrOoZSnIfFpsaNY1NOIHrg_x0g/edit?resourcekey=&gid=1782637761#gid=1782637761"

# Define the scope
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Load credentials
encoded_key = os.getenv("GOOGLE_CREDENTIALS")

if encoded_key:
    json_key = json.loads(base64.b64decode(encoded_key).decode("utf-8"))
    creds = ServiceAccountCredentials.from_json_keyfile_dict(json_key, scope)
else:
    creds_path = r"C:\Users\CxLos\OneDrive\Documents\BMHC\Data\bmhc-timesheet-4808d1347240.json"
    if os.path.exists(creds_path):
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
    else:
        raise FileNotFoundError("Service account JSON file not found and GOOGLE_CREDENTIALS is not set.")

# Authorize and load the sheet
client = gspread.authorize(creds)
sheet = client.open_by_url(sheet_url)
worksheet = sheet.get_worksheet(0)  # ✅ This grabs the first worksheet
data = pd.DataFrame(worksheet.get_all_records())
# data = pd.DataFrame(client.open_by_url(sheet_url).get_all_records())
df = data.copy()

# Trim leading and trailing whitespaces from column names
df.columns = df.columns.str.strip()

# Define a discrete color sequence
color_sequence = px.colors.qualitative.Plotly

# Filtered df where 'Date of Activity:' is between 2025-01-01 and 2025-03-31
df['Date of Activity'] = pd.to_datetime(df['Date of Activity'], errors='coerce')
df = df[(df['Date of Activity'] >= '2025-1-01') & (df['Date of Activity'] <= '2025-3-31')]
df['Month'] = df['Date of Activity'].dt.month_name()

df_1 = df[df['Month'] == 'January']
df_2 = df[df['Month'] == 'February']
df_3 = df[df['Month'] == 'March']

# print(df_1.head())
# print(df_2.head())
# print(df_3.head())

# Get the reporting quarter:
def get_custom_quarter(date_obj):
    month = date_obj.month
    if month in [10, 11, 12]:
        return "Q1"
    elif month in [1, 2, 3]:
        return "Q2"
    elif month in [4, 5, 6]:
        return "Q3"
    elif month in [7, 8, 9]:
        return "Q4"

# Example usage:
report_date = datetime(2025, 3, 1)
report_year = datetime(2025, 3, 1).year
current_quarter = get_custom_quarter(report_date)
print(f"Reporting Quarter: {current_quarter}")

# print(df_m.head())
# print('Total Marketing Events: ', len(df))
# print('Column Names: \n', df.columns)
# print('DF Shape:', df.shape)
# print('Dtypes: \n', df.dtypes)
# print('Info:', df.info())
# print("Amount of duplicate rows:", df.duplicated().sum())

# print('Current Directory:', current_dir)
# print('Script Directory:', script_dir)
# print('Path to data:',file_path)

# ================================= Columns ================================= #

columns =[
    'Timestamp',
    'Date of Activity', 
    'Person submitting this form:',
       'Activity Duration (Minutes)', 
       'Total travel time (Minutes):',
       'What type of MARCOM activity are you reporting?', 
       'BMHC Activity:',
       'Care Network Activity:', 
       'Brief activity description:',
       'Activity Status', 
       'Community Outreach Activity:',
       'Community Education Activity:',
       'Any recent or planned changes to BMHC lead services or programs?',
       'Entity Name:', 
       'Email Address', 
       'Month'
]

# =============================== Missing Values ============================ #

# missing = df.isnull().sum()
# print('Columns with missing values before fillna: \n', missing[missing > 0])

#  Please provide public information:    137
# Please explain event-oriented:        13

# ============================== Data Preprocessing ========================== #

# Check for duplicate columns
# duplicate_columns = df.columns[df.columns.duplicated()].tolist()
# print(f"Duplicate columns found: {duplicate_columns}")
# if duplicate_columns:
#     print(f"Duplicate columns found: {duplicate_columns}")

df.rename(
    columns={
        "What type of MARCOM activity are you reporting?": "MarCom Activity",
        "BMHC Activity:": "Purpose",
        # "Purpose of the activity (please only list one):": "Purpose",
        "Care Network Activity:": "Product Type",
        "Activity Duration (Minutes)": "Activity duration",
        "Total travel time (Minutes):": "Travel Time",
    }, 
inplace=True)

# Fill Missing Values
# df['Please provide public information:'] = df['Please provide public information:'].fillna('N/A')
# df['Please explain event-oriented:'] = df['Please explain event-oriented:'].fillna('N/A')

# print(df.dtypes)

# ------------------------------- MarCom Events DF ---------------------------------- #

marcom_events = len(df)

# --------------------------------- MarCom Hours DF --------------------------------- #

df_mc_hours = df[['Activity duration', 'Date of Activity']]
# print(df_mc_hours.head())

# print("MC Hours Unique Before: \n", df['Activity duration'].unique().tolist())

df['Activity duration'] = (
    df['Activity duration']
    .astype(str)
    .str.strip()
    .replace({
        '3 hours': 3,
        '4 hours': 4,
        '1 hour': 1,
        '0 - 1 hour': 0.5,
        '6 hours': 6,
        '5 hours': 5,
        '2 hours': 2,
        '8 hours': 8,
        '3.5 hour': 3.5,
        '0 - hour': 0.5,
        '0 - 1 our': 0.5,
        '0 - 1 hours': 0.5,
        '': None
    })
    .replace(' hours', '', regex=True)
    .replace(' hour', '', regex=True)
    .astype(float)
)

df['Activity duration'] = df['Activity duration'].fillna(0)
df['Activity duration'] = pd.to_numeric(df['Activity duration'], errors='coerce')

# print("MC Hours Unique After: \n", df['Activity duration'].unique().tolist())

marcom_hours = df.groupby('Activity duration').size().reset_index(name='Count')
marcom_hours = df['Activity duration'].sum()
marcom_hours = round(marcom_hours)
print('Q2 MarCom hours', marcom_hours, 'hours')

# Calculate total hours for each month
hours_oct = df[df['Month'] == 'January']['Activity duration'].sum()
hours_oct = round(hours_oct) 
print('MarCom hours in January:', hours_oct, 'hours')

hours_nov = df[df['Month'] == 'February']['Activity duration'].sum()
hours_nov = round(hours_nov)
print('MarCom hours in February:', hours_nov, 'hours')  

hours_dec = df[df['Month'] == 'March']['Activity duration'].sum()
hours_dec = round(hours_dec)
print('MarCom hours in March:', hours_dec, 'hours')

# Create df for MarCom Hours
df_hours = pd.DataFrame({
    'Month': ['January', 'February', 'March'],
    'Hours': [hours_oct, hours_nov, hours_dec]
})

# Bar chart for MarCom Hours
hours_fig = px.bar(
    df_hours,
    x='Month',
    y='Hours',
    color="Month",
    text='Hours',
    title= f'{current_quarter} MarCom Hours by Month',
    labels={
        'Hours': 'Number of Hours',
        'Month': 'Month'
    },
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Hours',
    height=900,  # Adjust graph height
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    xaxis=dict(
        tickfont=dict(size=18),  # Adjust font size for the month labels
        tickangle=-25,  # Rotate x-axis labels for better readability
        title=dict(
            text=None,
            font=dict(size=20),  # Font size for the title
        ),
    ),
    yaxis=dict(
        title=dict(
            text='Number of Hours',
            font=dict(size=22),  # Font size for the title
        ),
    ),
    bargap=0.08,  # Reduce the space between bars
).update_traces(
    texttemplate='%{text}',  # Display the count value above bars
    textfont=dict(size=20),  # Increase text size in each bar
    textposition='auto',  # Automatically position text above bars
    textangle=0, # Ensure text labels are horizontal
    hovertemplate=(  # Custom hover template
        '<b>Hours</b>: %{y}<extra></extra>'  
    ),
).add_annotation(
    x='October',  # Specify the x-axis value
    y=df_hours.loc[df_hours['Month'] == 'January', 'Hours'].values[0] + 100,  # Position slightly above the bar
    text='',  # Annotation text
    showarrow=False,  # Hide the arrow
    font=dict(size=30, color='red'),  # Customize font size and color
    align='center',  # Center-align the text
)

# Pie Chart MarCom Hours
hours_pie = px.pie(
    df_hours,
    names='Month',
    values='Hours',
    color='Month',
    height=800
).update_layout(
    title=dict(
        x=0.5,
        text=f'{current_quarter} MarCom Hours by Month',  # Title text
        font=dict(
            size=35,  # Increase this value to make the title bigger
            family='Calibri',  # Optional: specify font family
            color='black'  # Optional: specify font color
        ),
    ),  # Center-align the title
        margin=dict(
        l=0,  # Left margin
        r=0,  # Right margin
        t=100,  # Top margin
        b=0   # Bottom margin
    )  # Add margins around the chart
).update_traces(
    rotation=180,  # Rotate pie chart 90 degrees counterclockwise
    textfont=dict(size=19),  # Increase text size in each bar
    textinfo='value+percent',
    # texttemplate='<br>%{percent:.0%}',  # Format percentage as whole numbers
    hovertemplate='<b>%{label}</b>: %{value}<extra></extra>'
)

# ------------------------------- MarCom Travel ------------------------------ #

# Extracting "Travel" and "Date of activity:"
df_travel = df[['Travel Time', 'Date of Activity']]
# print(df_travel.head())

df['Activity duration'] = (
    df['Activity duration']
    .astype(str)
    .str.strip()
    .replace({
        ' hours' : '',
        ' hour' : '',
    })
)

df['Travel Time'] = pd.to_numeric(df['Travel Time'], errors='coerce')
travel_time = df['Travel Time'].sum()/60
travel_time = round(travel_time)
# print('Total Travel Time:', travel_time, 'hours')

# ------------------------------ MarCom Activity DF ------------------------------- #

# Extracting "MarCom Activity" and "Date of activity:"
df_activities = df[['MarCom Activity', 'Month']]
# print(df_activities.head())

# Group the data by "Month" and "MarCom Activity" to count occurrences
df_activity_counts = (
    df_activities.groupby(['Month', 'MarCom Activity'], sort=True)
    .size()
    .reset_index(name='Count')
)

# Sort months in the desired order
month_order = ['January', 'February', 'March']

df_activity_counts['Month'] = pd.Categorical(
    df_activity_counts['Month'],
    categories=month_order,
    ordered=True
)

df_activity_counts = df_activity_counts.sort_values(by=['Month', 'MarCom Activity'])

# Create a grouped bar chart
activity_fig = px.bar(
    df_activity_counts,
    x='Month',
    y='Count',
    color='MarCom Activity',
    barmode='group',
    text='Count',
    title= f'{current_quarter} MarCom Activities by Month',
    labels={
        'Count': 'Number of Activities',
        'Month': 'Month',
        'MarCom Activity': 'Activity Category'
    },
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    height=900,  # Adjust graph height
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    xaxis=dict(
        tickmode='array',
        tickvals=df_activity_counts['Month'].unique(),
        tickangle=-35  # Rotate x-axis labels for better readability
    ),
    legend=dict(
        title='Activity Category',
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        xanchor="left",  # Anchor legend to the left
        y=1,  # Position legend at the top
        yanchor="top"  # Anchor legend at the top
    ),
    hovermode='x unified',  # Unified hover display
    bargap=0.08,  # Reduce the space between bars
    bargroupgap=0,  # Reduce space between individual bars in groups
).update_traces(
    textposition='auto',  # Place count values outside bars
    textfont=dict(size=30),  # Increase text size in each bar
    hovertemplate=(
         '<br>'
        '<b>Count: </b>%{y}<br>'  # Count
    ),
    customdata=df_activity_counts['MarCom Activity'].values.tolist()  # Custom data for hover display
).add_vline(
    x=0.5,  # Adjust the position of the line
    line_dash="dash",
    line_color="gray",
    line_width=2
).add_vline(
    x=1.5,  # Position of the second line
    line_dash="dash",
    line_color="gray",
    line_width=2
)

# Group by "MarCom Activity" to count occurrences
df_mc_activity = df_activities.groupby('MarCom Activity').size().reset_index(name='Count')

#  Pie chart:
activity_pie = px.pie(
    df_mc_activity,
    names='MarCom Activity',
    values='Count',
    color=  'MarCom Activity',
    height=800
).update_layout(
    title=dict(
        x=0.5,
        text= f'{current_quarter} MarCom Activities',  # Title text
        font=dict(
            size=35,  # Increase this value to make the title bigger
            family='Calibri',  # Optional: specify font family
            color='black'  # Optional: specify font color
        ),
    ),
    margin=dict(
        l=0,  # Left margin
        r=0,  # Right margin
        t=100,  # Top margin
        b=0   # Bottom margin
    )  # Add margins around the chart
).update_traces(
    rotation=0,  # Rotate pie chart 90 degrees counterclockwise
    textfont=dict(size=19),  # Increase text size in each bar
    textinfo='value+percent',
    texttemplate='%{value}<br>%{percent:.0%}',  # Format percentage as whole numbers
    insidetextorientation='horizontal',  # Horizontal text orientation
    hovertemplate='<b>%{label}</b>: %{value}<extra></extra>'
)

# ---------------------------- Product DF ------------------------ #

data = [
    '', 'Administrative Task', 'Announcement', 'Branding', 'Editing/ Proofing/ Writing', 'Internal Communications, \tMeetings with Internal BMHC Teammember or Team\t1\tOrganizational Activity\tOrganizational Efficiency\tMeeting - Communications\t\t\t\t\t\tMeetings with Internal BMHC Teammember or Team\t1\tOrganizational Activity\tOrganizational Efficiency\tMeeting- Communications', 'MARCOM Check in meeting', 'Marketing', 'Meeting', 'Meeting, Presentation', 'Newsletter', 'Newsletter Archive', 'Newsletter, Writing, Editing, Proofing', 'No Product', 'No Product - Event Support', 'No Product - Internal Communications', 'No product', 'No product - Board Support', 'No product - organizational efficiency', 'No product - organizational strategy', 'No product - organizational strategy meeting', 'No product- organizational strategy', 'No product- troubleshooting', 'Non product - Board Support', 'Organizational Support', 'Presentation', 'Press Release PDF Folder', 'Social Media', 'Social Media Post', 'Update Center of Excellence for Youth Logo', 'Updates', 'Website Updates/Newsletter Archive', 'Website updates', 'newsletter archive', 'sent logo to Director of Outreach', 'website updates'
        ]


df['Product Type'] = (
    df['Product Type']
    .str.strip()
    .replace(
        {
            # No Product
            'No Product': "No Product",
            'No Product - Board Support': "No Product",
            'No Product - Branding Activity': "No Product",
            'No Product - Co-branding in General': "No Product",
            'No Product - Event Support': "No Product",
            'No Product - Internal Communications': "No Product",
            'No product': "No Product",
            'No product - Board Support': "No Product",
            'No product - Communications Support': "No Product",
            'No product - Community Collaboration': "No Product",
            'No product - Event Support': "No Product",
            'No product - Gathering Testimonials': "No Product",
            'No product - Human Resources Training for Organizational Efficiency': "No Product",
            'No product - Human Resources for Efficiency': "No Product",
            'No product - Organizational Efficiency': "No Product",
            'No product - Organizational Strategy': "No Product",
            'No product - organizational efficiency': "No Product",
            'No product - organizational strategy': "No Product",
            'No product - organizational strategy meeting': "No Product",
            'No product- organizational strategy': "No Product",
            'No product- troubleshooting': "No Product",
            'Non product - Board Support': "No Product",
            '': "No Product",  # blank entries

            # Meeting
            'Meeting': "Meeting",
            'Meeting - Communications': "Meeting",
            'Meeting - Impact Report': "Meeting",
            'Meeting - Social Media': "Meeting",
            'Meeting with Areebah': "Meeting",
            'Meeting with Director Pamela Parker': "Meeting",
            'MarCom Impact Report Meeting': "Meeting",
            'meeting with Pamela': "Meeting",
            'BMHC Board Meeting': "Meeting",
            'Key Leader Huddle': "Meeting",
            'Key Leaders Huddle': "Meeting",
            'Key Leaders Meeting': "Meeting",
            'Quarterly Team Meeting': "Meeting",
            'BMHC and Americorp Logo Requirements meeting': "Meeting",
            'MARCOM Check in meeting': "Meeting",
            'Meeting, Presentation': "Meeting",
            'Internal Communications, \tMeetings with Internal BMHC Teammember or Team\t1\tOrganizational Activity\tOrganizational Efficiency\tMeeting - Communications\t\t\t\t\t\tMeetings with Internal BMHC Teammember or Team\t1\tOrganizational Activity\tOrganizational Efficiency\tMeeting- Communications': "Meeting",

            # Newsletter
            'Newsletter': "Newsletter",
            'Newsletter,': "Newsletter",
            'Newsletter Archive': "Newsletter",
            'newsletter archive': "Newsletter",
            'Website Updates/Newsletter Archive': "Newsletter",
            'Newsletter, Writing, Editing, Proofing': "Newsletter",
            'Newsletter, Started Social Media and Newsletter Benchmarking': "Newsletter",
            'Newsletter, edit Social Media and Newsletter Benchmarking': "Newsletter",

            # Presentation
            'Presentation': "Presentation",
            'Presentation, Started Impact Report Presentation': "Presentation",
            'Marcom Report': "Presentation",
            'Sent Pamela Parker Board Member and Staff Diagram slides in PNG': "Presentation",

            # Scheduling
            'Scheduled ACC Tax help with Areebah': "Scheduling",
            'Scheduled Open Board Appointments and Man in Me Posts with Areebah': "Scheduling",

            # Updates
            'Updates': "Updates",
            'Website updates': "Updates",
            'Website Updates': "Updates",
            'updated - Board RFIs - January 2025': "Updates",
            'updated Board Due Outs file': "Updates",
            'Updated Marcom Data for December': "Updates",
            'Updated verbiage for MLK Social Media post': "Updates",
            'Update BMHC Service Webpage images': "Updates",
            'Updated ACC Student Video Post': "Updates",
            'Updated Felicia Chandler headshot in Photoshop for Website': "Updates",
            'Updated and Approve ACC Student Video Post': "Updates",
            'update and approved Organization chart': "Updates",
            'updated Red Card': "Updates",
            'updated and approved Red Card': "Updates",
            'website updates': "Updates",
            'Intranet Updates': "Updates",

            # Student-related activities
            'Came up with social media verbiage for Student Videos': "Student-related activity",
            'Reviewed Student Videos': "Student-related activity",
            'Worked on Video Inquiry and verbiage for ACC and UT': "Student-related activity",
            'Started SQL Certificates': "Student-related activity",

            # Social Media
            'Social Media': "Social Media",
            'Social Media Post': "Social Media",
            'approved and scheduled UT PhARM Social Media Post': "Social Media",
            'sent Areebah a schedule posts list from the Newsletter': "Social Media",
            'Man and Me schedule and post': "Social Media",
            'BMHC PSA Videos Project': "Social Media",
            'Video': "Social Media",
            'Social Media Post, reviewed partner list social media': "Social Media",

            # Editing/ Proofing/ Writing
            'Editing/ Proofing/ Writing': "Editing/ Proofing/ Writing",
            'Writing, Editing, Proofing': "Editing/ Proofing/ Writing",
            'created and updated Center of Excellence for Youth Mental Health logo': "Editing/ Proofing/ Writing",
            'Gathered and sent Previous Meeting Minutes': "Editing/ Proofing/ Writing",
            'Sustainability Binder': "Editing/ Proofing/ Writing",
            'MarCom Playbook': "Editing/ Proofing/ Writing",
            'provided board minutes for audit': "Editing/ Proofing/ Writing",

            # Branding
            'Branding': "Branding",
            'AmeriCorp Logo': "Branding",
            'AmeriCorps Responsibility': "Branding",
            'Co-branding in general': "Branding",
            'Update Center of Excellence for Youth Logo': "Branding",
            'sent logo to Director of Outreach': "Branding",
            'create and update Americorp Logo options': "Branding",
            'updated Center of Excellence Logo': "Branding",
            'Co-Branding, Flyer': "Branding",
            'Co-Branding, Presentation': "Branding",
            'updated Americorp logo': "Branding",
            'updated and finalized Americorp logo': "Branding",
            'created and updated Overcoming Mental Hellness Logo': "Branding",
            'updated Overcoming Mental Hellness Logo': "Branding",

            # Marketing
            'Marketing': "Marketing",
            'Announcement': "Marketing",
            'Community Radio PSA/Promos': "Marketing",
            'Flyer': "Marketing",
            'Please Push - Board Member Garza': "Marketing",
            'Please Push - Community Health Worker position': "Marketing",
            'Press Release': "Marketing",
            'Press Release PDF Folder': "Marketing",
            'Sent Harvard Press Release Links': "Marketing",

            # Administrative Task
            'Administrative Task': "Administrative Task",
            'Created Certificate Order Guide': "Administrative Task",
            'Move all files to BMHC Canva': "Administrative Task",
            'Organizational Support': "Administrative Task",
            'Organizational Efficiency': "Administrative Task",
            'Timesheet': "Administrative Task",
            'Organization': "Administrative Task",

            # Other
            'created Red Card': "Design",
            'Report': "Report",
            'Impact Report': "Impact Report",
        }
    )
)


# List of allowed final values
allowed_categories = [
    "No Product",
    "Meeting",
    "Newsletter",
    "Presentation",
    "Scheduling",
    "Updates",
    "Student-related activity",
    "Social Media",
    "Editing/ Proofing/ Writing",
    "Branding",
    "Marketing",
    "Administrative Task",
    "Design",
    "Report",
    "Impact Report"
]

# Find any remaining unmatched purposes
unmatched_products = df[~df['Product Type'].isin(allowed_categories)]['Product Type'].unique().tolist()
# print("Unmatched Products: \n", unmatched_products)

# Product Type dataframe:
df_product_type = df.groupby('Product Type').size().reset_index(name='Count')
# print("Product Unique", df_product_type["Product Type"].unique().tolist())

product_bar=px.bar(
    df_product_type,
    x='Product Type',
    y='Count',
    color='Product Type',
    text='Count',
).update_layout(
    height=990, 
    width=1700,
    title=dict(
        text= f'{current_quarter} Product Type',
        x=0.5, 
        font=dict(
            size=25,
            family='Calibri',
            color='black',
            )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        tickangle=-15,  # Rotate x-axis labels for better readability
        tickfont=dict(size=18),  # Adjust font size for the tick labels
        title=dict(
            # text=None,
            text="Product",
            font=dict(size=20),  # Font size for the title
        ),
        showticklabels=False  # Hide x-tick labels
        # showticklabels=True  # Hide x-tick labels
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20),  # Font size for the title
        ),
    ),
    legend=dict(
        # title='Support',
        title_text='',
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        y=1,  # Position legend at the top
        xanchor="left",  # Anchor legend to the left
        yanchor="top",  # Anchor legend to the top
        # visible=False
        visible=True
    ),
    hovermode='closest', # Display only one hover label per trace
    bargap=0.08,  # Reduce the space between bars
    bargroupgap=0,  # Reduce space between individual bars in groups
).update_traces(
    textposition='auto',
    hovertemplate='<b>Name:</b> %{label}<br><b>Count</b>: %{y}<extra></extra>'
)

# Pie Chart
product_pie=px.pie(
    df_product_type,
    names="Product Type",
    values='Count'  # Specify the values parameter
).update_layout(
    height=900,
    width=1700,
    title= f'{current_quarter} Product Type',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    margin=dict(
        l=0,  # Left margin
        r=0,  # Right margin
        t=100,  # Top margin
        b=0   # Bottom margin
    )  # Add margins around the chart
).update_traces(
    rotation=180,
    textposition='auto',
    # textinfo='none',
    textinfo='value+percent',
    hovertemplate='<b>%{label} Status</b>: %{value}<extra></extra>',
)

# ---------------------------- Purpose DF ------------------------ #

# print("Purpose Unique:", df["Purpose"].unique().tolist())

purpose_unique = [
'Key Leaders Huddle', 'sent Areebah a schedule posts list from the Newsletter ', 'edit Newsletter', 'Organizational Efficiency', 'Impact Metrics', 'Organizational Strategy', 'Came up with social media verbaige for Student Videos', 'provided board mintues for audit', 'Health Awareness & ED Public Information', 'Marketing Promotion', 'Scheduled ACC Tax help with Areebah ', 'Reviewed Student Videos', 'Training', 'Organizational Collaboration', 'MarCom Impact Report Meeting', 'BMHC Board Meeting', 'edit Black History Month visuals', 'Updated Board Presentation Slides', 'Worked on Video Inquiry and verbiage for ACC and UT ', ' edit Black History Month visuals', 'Scheduled Open Board Appointments and Man in Me Posts with Areebah', 'Sent out Newsletter plan', 'created and sent Know Your Number Day Event flyer to Steve ', 'meeting with Pamela', 'created Newsletter plan for 1/31', 'created and updated Center of Excellence for Youth Mental Health logo', 'edit Black History Month visuals ', 'Timesheet ', 'Updated Visual for BBRA Co-Flyer', 'sent Pamela CommUnityCare TV ads', 'Communications Support', 'Key or Special Event Support', 'edit CommUnityCare TV ads', 'edit CommUnityCare TV ads ', 'Key Leader Huddle ', 'edit and approve Larry Wallace Congratulations Visual ', 'edit and send Newsletter', 'updated Board Due Outs file', 'Updated verbiage for MLK Social Media post', 'edit Larry Wallace Congratulations Visual', 'updated Board Due Outs file ', 'Onboarding or Hiring Staff', 'Edit and approved MLK Visual', 'started Larry Wallace Congratulations Visual', ' edit Impact Report Presentation', 'Updated Marcom Data for December', 'started MLK Visual', 'edit Impact Report Presentation', 'Sent out Newsletter plan to team', 'updated Newsletter', 'Gathered BMHC Event Flyer and sent to Cameron', 'Key Leaders Huddle ', 'updated - Board RFIs - January 2025', 'Quarterly Team Meeting ', 'Meeting with Director Pamela Parker ', 'Started Social Media and Newsletter Analytic Slide Show ', 'Worked on Newsletter Plan with Pamela', 'approved and scheduled UT PhARM Social Media Post ', 'Meeting with Areebah', 'edit Social Media and Newsletter Benchmarking', 'Newsletter meeting', 'Started Social Media and Newsletter Benchmarking', 'Started Impact Report Presentation', 'Key Leaders Meeting', 'Started BMHC Quarterly Health Presentation', 'Gathered and sent Previous Meeting Minutes', 'Updated Digital Analytics for Social Media and Mailchimp', 'Health Education', 'Historical Purposes and Organizational Efficiency', 'Edit Seton Flyer', 'Man and Me schedule and post', 'Organization', 'Updated Black History Month Visuals ', 'created Red Card', 'Started SQL Certificates', 'updated the Board slide on all powerpoints ', 'updated Red Card', 'Timesheet', 'Updated Black History Month Visual ', 'sent logo to Director of Outreach', 'updated Behavioral Health Therapy Co-Branded Flyer', 'updated and approved Red Card', 'Key Leader Huddle', 'Updated Felicia Chandler headshot in Photoshop for Website', ' Impact Report', 'Please Push - Board Member Garza', 'Created and Sent Newsletter plan for 2/14', 'Updated Newsletter ', 'started and updated BMHC + CFV Flyer', 'BMHC and Americorp Logo Requirements meeting ', 'update Organization chart', 'Update Newsletter', 'MARCOM Check in meeting', 'Updated BMHC + CFV Flyer', 'Update Newsletter ', 'update and approved Organization chart', "Valentine's Post and Verbiage", 'Please Push - Community Health Worker position', 'Update BMHC Service Webpage images', "Create Women's History Visuals", 'Create Diabetes class Visual', 'Created Certificate Order Guide', "Updated Women's History Visuals", 'Updated Diabetes class Visual', ' Updated ACC Student Video Post', 'Updated Newsletter', 'Updated and Approve ACC Student Video Post', 'Move all files to BMHC Canva', 'AmeriCorp Logo', 'BMHC Branding', 'BMHC Co-Branding', 'Special Event Execution', 'Organizational Activity', 'Marketing Analysis', 'Community Awareness', 'Health Awareness & Ed Public Information', 'Adding Content', 'Uploading PDFs to Drive', 'Adding/reviewing content', 'Add/Review Content', 'Adding/Reviewing Content', 'Adding/reviewing Content', 'BMHC Board Advisory Committee ', 'Organizational Activity - Board Support', 'General Communications Emails with Social Media Team - providing social links to Bristol Myers about the Harvard Law Press Release', 'Website Troubleshooting', 'Working with IT to get content up', 'Care Network Related Strategy', 'Research Website plugins', 'Organization Strategy', 'Sustainability Binder', 'Add/ Review Content', 'Scheduled Military Affiliated Job Post', 'Marcom Presentation Meeting', 'Updated Marcom Presentation', 'Schedule Measle Post', 'create and update Americorp Logo options', 'update the Student  Video Google Doc', 'Scheduled CTOSH sign up Post ', 'Schedule BMHC SPECIAL ANNOUNCEMENT Post ', 'Updated and Sent CommUnityCare Lobby TVs Ads for review', 'updated Military slides', 'Schedule CommUnity Care Measles Vaccine Clinic Post', 'updated Center of Excellence Logo', 'updated Marcom presentation', 'create and update Americorp Logo options ', 'Updated board member slides on all active presentations', 'created and updated MARCOM slide', 'updated MARCOM Presentation', 'Updated MARCOM Presentation', "Meeting with Dr. Wallace to discuss Overcoming Mental Health logo and Bartley's Method ", 'posted Harvard Press Release', 'updated Americorp logo', 'Sent Harvard Press Release Links', 'updated and finalized Americorp logo', 'sent Areebah posts for social media', 'Sent Newsletter Plan to the team', 'Updated Military Slides', 'Updated and sent Dr. Wallace Military slides', 'Sent Pamela Parker Board Member and Staff Diagram slides in PNG', 'created and updated Overcoming Mental Hellness Logo', 'updated Overcoming Mental Hellness Logo', 'Posted Board Member Special Announcement ', 'Meeting with Areebah to discuss social Media and reposting', 'reviewed partner list social media', 'BMHC PSA Videos Project', 'General Health Awareness Activity', 'Impact Metrics '
]

df['Purpose'] = (
    df['Purpose']
    .str.strip()
    .replace({
        # --- Add/Review Content ---
        'Add/ Review Content': 'Add/Review Content',
        'Adding Content': 'Add/Review Content',
        'Adding/Reviewing Content': 'Add/Review Content',
        'Adding/reviewing Content': 'Add/Review Content',
        'Adding/reviewing content': 'Add/Review Content',
           'Updated Digital Analytics for Social Media and Mailchimp': 'Impact & Metrics',

        # --- BMHC Activity ---
        'BMHC Board Advisory Committee': 'BMHC Activity',
        'BMHC Co-Branding': 'BMHC Activity',
        'BMHC PSA Videos Project': 'BMHC Activity',
        'BMHC and Americorp Logo Requirements meeting': 'BMHC Activity',
        'BMHC Branding': 'BMHC Activity',

        # --- Communications ---
        'Communications Support': 'Communications',
        'General Communications Emails with Social Media Team - providing social links to Bristol Myers about the Harvard Law Press Release': 'Communications',
        'sent Areebah posts for social media': 'Communications',
        'Sent Harvard Press Release Links': 'Communications',
        'posted Harvard Press Release': 'Communications',
        'reviewed partner list social media': 'Communications',
        'Meeting with Areebah to discuss social Media and reposting': 'Communications',
            'sent Areebah a schedule posts list from the Newsletter': 'Communications',
    'sent Pamela CommUnityCare TV ads': 'Communications',
    'edit CommUnityCare TV ads': 'Communications',

        # --- Health Awareness ---
        'General Health Awareness Activity': 'Health Awareness',
        'Health Awareness & ED Public Information': 'Health Awareness',
        'Health Awareness & Ed Public Information': 'Health Awareness',
        'Health Education': 'Health Awareness',
        'Create Diabetes class Visual': 'Health Awareness',
        'Created Certificate Order Guide': 'Health Awareness',
        "Create Women's History Visuals": 'Health Awareness',
        "Updated Women's History Visuals": 'Health Awareness',
        'Updated Diabetes class Visual': 'Health Awareness',
        "Valentine's Post and Verbiage": 'Health Awareness',

        # --- Impact & Metrics ---
        'Impact Metrics': 'Impact & Metrics',
        'Impact Report': 'Impact & Metrics',
        'Started Impact Report Presentation': 'Impact & Metrics',
        'edit Impact Report Presentation': 'Impact & Metrics',
        'Updated Marcom Data for December': 'Impact & Metrics',

        # --- Leadership Meeting ---
        'Key Leader Huddle': 'Leadership Meeting',
        'Key Leaders Huddle': 'Leadership Meeting',
        'Key Leaders Meeting': 'Leadership Meeting',
        'Meeting with Areebah': 'Leadership Meeting',
        'Meeting with Director Pamela Parker': 'Leadership Meeting',
        'meeting with Pamela': 'Leadership Meeting',
        'MarCom Impact Report Meeting': 'Leadership Meeting',
        'Quarterly Team Meeting': 'Leadership Meeting',
        'BMHC Board Meeting': 'Leadership Meeting',
        'Newsletter meeting': 'Leadership Meeting',
        'Meeting with Dr. Wallace to discuss Overcoming Mental Health logo and Bartley\'s Method': 'Leadership Meeting',
           'MARCOM Check in meeting': 'Leadership Meeting',
    'Marcom Presentation Meeting': 'Leadership Meeting',

        # --- Special Event Support ---
        'Key or Special Event Support': 'Special Event Support',
        'Special Event Execution': 'Special Event Support',
        'Scheduled CTOSH sign up Post': 'Special Event Support',
        'Schedule BMHC SPECIAL ANNOUNCEMENT Post': 'Special Event Support',
        'Posted Board Member Special Announcement': 'Special Event Support',
        'created and sent Know Your Number Day Event flyer to Steve': 'Special Event Support',

        # --- Marketing ---
        'Marketing Analysis': 'Marketing',
        'Marketing Promotion': 'Marketing',
        'approved and scheduled UT PhARM Social Media Post': 'Marketing',
        'edit Social Media and Newsletter Benchmarking': 'Marketing',
        'Started Social Media and Newsletter Benchmarking': 'Marketing',
        'Started Social Media and Newsletter Analytic Slide Show': 'Marketing',
        'Came up with social media verbaige for Student Videos': 'Marketing',

        # --- HR / Staffing ---
        'Onboarding or Hiring Staff': 'HR / Staffing',
        'Please Push - Community Health Worker position': 'HR / Staffing',

        # --- Organizational Strategy ---
        'Organization': 'Organizational Strategy',
        'Organization Strategy': 'Organizational Strategy',
        'update Organization chart': 'Organizational Strategy',
        'update and approved Organization chart': 'Organizational Strategy',
        'Care Network Related Strategy': 'Organizational Strategy',

        # --- Organizational Activity ---
        'Organizational Strategy': 'Organizational Strategy',
        'Organizational Activity': 'Organizational Activity',
        'Organizational Activity - Board Support': 'Organizational Activity',
        'Organizational Collaboration': 'Organizational Activity',
        'Organizational Efficiency': 'Organizational Activity',
        'provided board mintues for audit': 'Organizational Activity',
        'updated - Board RFIs - January 2025': 'Organizational Activity',
        'Updated Board Presentation Slides': 'Organizational Activity',
        'updated the Board slide on all powerpoints': 'Organizational Activity',
        'updated Board Due Outs file': 'Organizational Activity',
        'Gathered and sent Previous Meeting Minutes': 'Organizational Activity',
        'Updated board member slides on all active presentations': 'Organizational Activity',
        'Sent Pamela Parker Board Member and Staff Diagram slides in PNG': 'Organizational Activity',

        # --- Web & Tech ---
        'Website Troubleshooting': 'Web & Tech',
        'Working with IT to get content up': 'Web & Tech',
        'Uploading PDFs to Drive': 'Web & Tech',
        'Research Website plugins': 'Web & Tech',
        'Update BMHC Service Webpage images': 'Web & Tech',

        # --- Scheduling ---
        'Schedule Measle Post': 'Scheduling',
        'Scheduled Military Affiliated Job Post': 'Scheduling',
        'Scheduled ACC Tax help with Areebah': 'Scheduling',
        'Scheduled Open Board Appointments and Man in Me Posts with Areebah': 'Scheduling',
        'Schedule CommUnity Care Measles Vaccine Clinic Post': 'Scheduling',

        # --- Admin / Documentation ---
        'Sustainability Binder': 'Admin/Documentation',
        'Timesheet': 'Admin/Documentation',
        'Please Push - Board Member Garza': 'Admin/Documentation',
        'Updated Felicia Chandler headshot in Photoshop for Website': 'Admin/Documentation',

        # --- Newsletter ---
        'Update Newsletter': 'Newsletter',
        'Updated Newsletter': 'Newsletter',
        'Sent out Newsletter plan': 'Newsletter',
        'edit Newsletter': 'Newsletter',
        'edit and send Newsletter': 'Newsletter',
        'Sent out Newsletter plan to team': 'Newsletter',
        'Created and Sent Newsletter plan for 2/14': 'Newsletter',
        'created Newsletter plan for 1/31': 'Newsletter',
        'Sent Newsletter Plan to the team': 'Newsletter',
        'updated Newsletter': 'Newsletter',
        'Worked on Newsletter Plan with Pamela': 'Newsletter',
            'started and updated BMHC + CFV Flyer': 'Branding',
    'Updated BMHC + CFV Flyer': 'Branding',

        # --- Training ---
        'Training': 'Training',
        'Reviewed Student Videos': 'Training',
        'Started SQL Certificates': 'Training',

        # --- Branding ---
        'sent logo to Director of Outreach': 'Branding',
        'create and update Americorp Logo options': 'Branding',
        'updated and finalized Americorp logo': 'Branding',
        'updated Americorp logo': 'Branding',
        'AmeriCorp Logo': 'Branding',
        'created and updated Overcoming Mental Hellness Logo': 'Branding',
        'updated Overcoming Mental Hellness Logo': 'Branding',
        'create and update Americorp Logo options ': 'Branding',
        'created and updated MARCOM slide': 'Branding',
        'updated MARCOM Presentation': 'Branding',
        'Updated MARCOM Presentation': 'Branding',
        'updated Marcom presentation': 'Branding',

        # --- Community Awareness ---
        'Community Awareness': 'Community Awareness',
        'Gathered BMHC Event Flyer and sent to Cameron': 'Community Awareness',
        'Man and Me schedule and post': 'Community Awareness',
        'Updated Black History Month Visual': 'Community Awareness',
        'Updated Black History Month Visuals': 'Community Awareness',
        'edit Black History Month visuals': 'Community Awareness',
        'edit Larry Wallace Congratulations Visual': 'Community Awareness',
        'edit and approve Larry Wallace Congratulations Visual': 'Community Awareness',
        'Edit and approved MLK Visual': 'Community Awareness',
        'Updated verbiage for MLK Social Media post': 'Community Awareness',
        'started Larry Wallace Congratulations Visual': 'Community Awareness',
        'started MLK Visual': 'Community Awareness',
        'updated Behavioral Health Therapy Co-Branded Flyer': 'Community Awareness',
        'updated and approved Red Card': 'Community Awareness',
        'created Red Card': 'Community Awareness',
        'updated Red Card': 'Community Awareness',
        'Updated and Approve ACC Student Video Post': 'Community Awareness',
        'Updated ACC Student Video Post': 'Community Awareness',
        'Community Awareness': 'Community Awareness',
            'Community Awareness': 'Community Awareness',

        # --- Other ---
        'Historical Purposes and Organizational Efficiency': 'Organizational Activity',
        'Edit Seton Flyer': 'Marketing',
        'Move all files to BMHC Canva': 'Admin/Documentation',
        'update the Student  Video Google Doc': 'Admin/Documentation',
        'Started BMHC Quarterly Health Presentation': 'Health Awareness',
        'Updated and Sent CommUnityCare Lobby TVs Ads for review': 'Marketing',
        
            # --- Content / Media Creation ---
    'Worked on Video Inquiry and verbiage for ACC and UT': 'Add/Review Content',
    'created and updated Center of Excellence for Youth Mental Health logo': 'Branding',
    'Updated Visual for BBRA Co-Flyer': 'Branding',
    
        # --- Military Slides / Presentations ---
    'updated Military slides': 'Organizational Activity',
    'Updated Military Slides': 'Organizational Activity',
    'Updated and sent Dr. Wallace Military slides': 'Organizational Activity',
    
        # --- Branding ---
    'updated Center of Excellence Logo': 'Branding',
        'Updated Marcom Presentation': 'Branding',
    })
)


# List of allowed final values
allowed_categories = [
    'Add/Review Content', 'BMHC Activity', 'Communications', 'Health Awareness',
    'Impact & Metrics', 'Leadership Meeting', 'Special Event Support', 'Marketing',
    'HR / Staffing', 'Organizational Strategy', 'Organizational Activity',
    'Web & Tech', 'Scheduling', 'Admin/Documentation', 'Newsletter',
    'Training', 'Branding'
]

# Find any remaining unmatched purposes
unmatched_purposes = df[~df['Purpose'].isin(allowed_categories)]['Purpose'].unique().tolist()
# print("Unmatched Purposes: \n", unmatched_purposes)

df_purpose = df.groupby('Purpose').size().reset_index(name='Count')

# Purpose Bar Chart
purpose_bar = px.bar(
    df_purpose,
    x='Purpose',
    y='Count',
    color='Purpose',
    text='Count',
).update_layout(
    height=990,
    width=1700,
    title=dict(
        text= f'{current_quarter} Purpose Chart',
        x=0.5,
        font=dict(
            size=25,
            family='Calibri',
            color='black',
        )
    ),
    font=dict(
        family='Calibri',
        size=18,
        color='black'
    ),
    xaxis=dict(
        tickangle=-15,
        tickfont=dict(size=18),
        title=dict(
            text="Purpose",
            font=dict(size=20),
        ),
        showticklabels=False
    ),
    yaxis=dict(
        title=dict(
            text='Count',
            font=dict(size=20),
        ),
    ),
    legend=dict(
        title_text='',
        orientation="v",
        x=1.05,
        y=1,
        xanchor="left",
        yanchor="top",
        visible=True
    ),
    hovermode='closest',
    bargap=0.08,
    bargroupgap=0,
).update_traces(
    textposition='auto',
    hovertemplate='<b>Name:</b> %{label}<br><b>Count</b>: %{y}<extra></extra>'
)

# Purpose Pie Chart
purpose_pie = px.pie(
    df_purpose,
    names="Purpose",
    values='Count'
).update_layout(
    height=900,
    width=1700,
    title= f'{current_quarter} Purpose',
    title_x=0.5,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    margin=dict(
        l=0,  # Left margin
        r=0,  # Right margin
        t=100,  # Top margin
        b=0   # Bottom margin
    )  # Add margins around the chart
).update_traces(
    rotation=160,
    textposition='auto',
    textinfo='value+percent',
    hovertemplate='<b>%{label} Status</b>: %{value}<extra></extra>',
)

# ---------------------------- MarCom Person Completing DF ------------------------ #

# Extracting "Person completing this form:" and "Date of activity:"
df_person = df[['Person submitting this form:', 'Month']]

# print("Person submitting before: \n", df['Person submitting this form:'].unique().tolist())

# Excludde rows with NaN or empty values in "Person completing this form:" and "Month"
df_person = df_person[
    df_person['Person submitting this form:'].notna() &
    df_person['Person submitting this form:'].str.strip().ne('') &
    df_person['Month'].notna() &
    df_person['Month'].astype(str).str.strip().ne('')
]

# Remove leading and trailing whitespaces and perform the replacements
df_person['Person submitting this form:'] = (
    df_person['Person submitting this form:']
    .str.strip()
    .replace({
        'Felicia Chanlder': 'Felicia Chandler',
        'Felicia Banks': 'Felicia Chandler',
    })
)

# print("Person submitting after: \n", df['Person submitting this form:'].unique().tolist())

# Group the data by "Month" and "Person completing this form:" to count occurrences
df_person_counts = (
    df_person.groupby(['Month', 'Person submitting this form:'], 
    sort=True) # Sort the values
    .size()
    .reset_index(name='Count')
)

# Sort months in the desired order
month_order = ['January', 'February', 'March']
df_person_counts['Month'] = pd.Categorical(
    df_person_counts['Month'],
    categories=month_order,
    ordered=True
)

# Sort df
df_person_counts = df_person_counts.sort_values(by=['Month', 'Person submitting this form:'])

# Create a grouped bar chart
person_fig = px.bar(
    df_person_counts,
    x='Month',
    y='Count',
    color='Person submitting this form:',
    barmode='group',
    text='Count',
    title= f'{current_quarter} Forms Completed by Month',
    labels={
        'Count': 'Number of Forms',
        'Month': 'Month',
        'Person submitting this form:': 'Person'
    },
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    height=900,  # Adjust graph height
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    xaxis=dict(
        tickmode='array',
        tickvals=df_person_counts['Month'].unique(),
        tickangle=-35  # Rotate x-axis labels for better readability
    ),
    legend=dict(
        title='Person',
        orientation="v",  # Vertical legend
        x=1.05,  # Position legend to the right
        xanchor="left",  # Anchor legend to the left
        y=1,  # Position legend at the top
        yanchor="top"  # Anchor legend at the top
    ),
    hovermode='x unified',  # Unified hover display
    bargap=0.08,  # Reduce the space between bars
    bargroupgap=0,  # Reduce space between individual bars in groups
).update_traces(
    textposition='auto',  # Place count values outside bars
    textfont=dict(size=30),  # Increase text size in each bar
    hovertemplate=(
        '<br>'
        '<b>Count: </b>%{y}<br>'  # Count
    ),
    customdata=df_person_counts['Person submitting this form:'].values.tolist()  # Custom data for hover display
).add_vline(
    x=0.5,  # Adjust the position of the line
    line_dash="dash",
    line_color="gray",
    line_width=2
).add_vline(
    x=1.5,  # Position of the second line
    line_dash="dash",
    line_color="gray",
    line_width=2
)

df_pf = df[['Person submitting this form:', 'Month']]

# Filter out nulls and empty strings
df_pf = df_pf[
    df_pf['Person submitting this form:'].notna() &
    df_pf['Person submitting this form:'].astype(str).str.strip().ne('')
]

# Remove leading and trailing whitespaces and perform the replacements
df_pf['Person submitting this form:'] = (
    df_pf['Person submitting this form:']
    .str.strip()
    .replace({
        'Felicia Chanlder': 'Felicia Chandler',
        'Felicia Banks': 'Felicia Chandler'
    })
)

# Group the data by "Person completing this form:" to count occurrences
df_pf = df_pf.groupby('Person submitting this form:').size().reset_index(name="Count")

# Bar chart for  Totals:
person_totals_fig = px.bar(
    df_pf,
    x='Person submitting this form:',
    y='Count',
    color='Person submitting this form:',
    text='Count',
).update_layout(
    height=850,  # Adjust graph height
    title=dict(
        x=0.5,
        text='Total Q1 Form Submissions by Person',  # Title text
        font=dict(
            size=35,  # Increase this value to make the title bigger
            family='Calibri',  # Optional: specify font family
            color='black'  # Optional: specify font color
        )
    ),
    xaxis=dict(
        tickfont=dict(size=18),  # Adjust font size for the month labels
        tickangle=-25,  # Rotate x-axis labels for better readability
        title=dict(
            text=None,
            font=dict(size=20),  # Font size for the title
        ),
    ),
    yaxis=dict(
        title=dict(
            text='Number of Submissions',
            font=dict(size=22),  # Font size for the title
        ),
    ),
    bargap=0.08,  # Reduce the space between bars
).update_traces(
    texttemplate='%{text}',  # Display the count value above bars
    textfont=dict(size=20),  # Increase text size in each bar
    textposition='auto',  # Automatically position text above bars
    textangle=0, # Ensure text labels are horizontal
    hovertemplate=(  # Custom hover template
        '<b>Name</b>: %{label}<br><b>Count</b>: %{y}<extra></extra>'  
    ),
)

pf_pie = px.pie(
    df_pf,
    names='Person submitting this form:',
    values='Count',
    color='Person submitting this form:',
    height=800
).update_layout(
    title=dict(
        x=0.5,
        text= f'{current_quarter} People Completing Forms',  # Title text
        font=dict(
            size=35,  # Increase this value to make the title bigger
            family='Calibri',  # Optional: specify font family
            color='black'  # Optional: specify font color
        ),
    ),  # Center-align the title
    margin=dict(
        t=150,  # Adjust the top margin (increase to add more padding)
        l=20,   # Optional: left margin
        r=20,   # Optional: right margin
        b=20    # Optional: bottom margin
    )
).update_traces(
    rotation=0,  # Rotate pie chart 90 degrees counterclockwise
    textfont=dict(size=19),  # Increase text size in each bar
    textinfo='value+percent',
    insidetextorientation='horizontal',  # Horizontal text orientation
    texttemplate='%{value}<br>%{percent:.0%}',  # Format percentage as whole numbers
    hovertemplate='<b>%{label}</b>: %{value}<extra></extra>'
)

# ---------------------------- MarCom Activity Status DF -------------------------- #

# "Activity Status" dataframe:
df_activity_status = df.groupby('Activity Status').size().reset_index(name='Count')

# print("Activity Status Unique before: \n", df['Activity Status'].unique().tolist())

mode = df['Activity Status'].mode()[0]
df['Activity Status'] = df['Activity Status'].fillna(mode)

df_activity_status['Activity Status'] = (
    df_activity_status['Activity Status']
    .str.strip()
    .replace({
        '': mode,
    })
)

df_activity_status_counts = (
    df.groupby(['Month', 'Activity Status'], sort=False)
    .size()
    .reset_index(name='Count')
)

df_activity_status_counts['Month'] = pd.Categorical(
    df_activity_status_counts['Month'],
    categories=month_order,
    ordered=True
)

# print("Activity Status Unique after: \n", df['Activity Status'].unique().tolist())

status_fig = px.bar(
    df_activity_status_counts,
    x='Month',
    y='Count',
    color='Activity Status',
    barmode='group',
    text='Count',
    labels={
        'Count': 'Number of Submissions',
        'Month': 'Month',
        'Activity Status': 'Activity Status'
    }
).update_layout(
    title_x=0.5,
    xaxis_title='Month',
    yaxis_title='Count',
    height=900,
    font=dict(
        family='Calibri',
        size=17,
        color='black'
    ),
    title=dict(
        text= f'{current_quarter} Activity Status by Month',
        x=0.5, 
        font=dict(
            size=22,
            family='Calibri',
            color='black',
            )
    ),
    xaxis=dict(
        tickmode='array',
        tickvals=df_activity_status_counts['Month'].unique(),
        tickangle=-35
    ),
    legend=dict(
        title='',
        orientation="v",
        x=1.05,
        xanchor="left",
        y=1,
        yanchor="top"
    ),
    hovermode='x unified'
).update_traces(
    textfont=dict(size=25),  # Increase text size in each bar
    textposition='outside',
    hovertemplate='<br><b>Count: </b>%{y}<br>',
    customdata=df_activity_status_counts['Activity Status'].values.tolist()
)

status_pie = px.pie(
        df_activity_status,
        values='Count',
        names='Activity Status',
        color='Activity Status',
    ).update_layout(
        title= f'{current_quarter} MarCom Activity Status',
        title_x=0.5,
        height=500,
        font=dict(
            family='Calibri',
            size=17,
            color='black'
        )
    ).update_traces(
        textposition='inside',
        textinfo='percent+label',
        hoverinfo='label+percent',
         texttemplate='%{value}<br>%{percent:.0%}',  # Format percentage as whole numbers
        hole=0.2
    )

# # ========================== MarCom DataFrame Table ========================== #

# MarCom Table
marcom_table = go.Figure(data=[go.Table(
    # columnwidth=[50, 50, 50],  # Adjust the width of the columns
    header=dict(
        values=list(df.columns),
        fill_color='paleturquoise',
        align='center',
        height=30,  # Adjust the height of the header cells
        # line=dict(color='black', width=1),  # Add border to header cells
        font=dict(size=12)  # Adjust font size
    ),
    cells=dict(
        values=[df[col] for col in df.columns],
        fill_color='lavender',
        align='left',
        height=25,  # Adjust the height of the cells
        # line=dict(color='black', width=1),  # Add border to cells
        font=dict(size=12)  # Adjust font size
    )
)])

marcom_table.update_layout(
    margin=dict(l=50, r=50, t=30, b=40),  # Remove margins
    height=400,
    # width=1500,  # Set a smaller width to make columns thinner
    paper_bgcolor='rgba(0,0,0,0)',  # Transparent background
    plot_bgcolor='rgba(0,0,0,0)'  # Transparent plot area
)

# ------------------------------- MarCom Products Table ----------------------------#

df_product_type = df.groupby('Product Type').size().reset_index(name='Count')

# Products Table
products_table = go.Figure(data=[go.Table(
    # columnwidth=[50, 50, 50],  # Adjust the width of the columns
    header=dict(
        values=list(df_product_type.columns),
        fill_color='paleturquoise',
        align='center',
        height=30,  # Adjust the height of the header cells
        # line=dict(color='black', width=1),  # Add border to header cells
        font=dict(size=12)  # Adjust font size
    ),
    cells=dict(
        values=[df_product_type[col] for col in df_product_type.columns],
        fill_color='lavender',
        align='left',
        height=25,  # Adjust the height of the cells
        # line=dict(color='black', width=1),  # Add border to cells
        font=dict(size=12)  # Adjust font size
    )
)])

products_table.update_layout(
    margin=dict(l=50, r=50, t=30, b=40),  # Remove margins
    height=900,
    # width=1500,  # Set a smaller width to make columns thinner
    paper_bgcolor='rgba(0,0,0,0)',  # Transparent background
    plot_bgcolor='rgba(0,0,0,0)'  # Transparent plot area
)

# ----------------------------- MarCom Purpose Table ----------------------------#

df_purpose = df.groupby('Purpose').size().reset_index(name='Count')

purpose_table = go.Figure(data=[go.Table(
    # columnwidth=[50, 50, 50],  # Adjust the width of the columns
    header=dict(
        values=list(df_purpose.columns),
        fill_color='paleturquoise',
        align='center',
        height=30,  # Adjust the height of the header cells
        # line=dict(color='black', width=1),  # Add border to header cells
        font=dict(size=12)  # Adjust font size
    ),
    cells=dict(
        values=[df_purpose[col] for col in df_purpose.columns],
        fill_color='lavender',
        align='left',
        height=25,  # Adjust the height of the cells
        # line=dict(color='black', width=1),  # Add border to cells
        font=dict(size=12)  # Adjust font size
    )
)])

purpose_table.update_layout(
    margin=dict(l=50, r=50, t=30, b=40),  # Remove margins
    height=900,
    # width=1500,  # Set a smaller width to make columns thinner
    paper_bgcolor='rgba(0,0,0,0)',  # Transparent background
    plot_bgcolor='rgba(0,0,0,0)'  # Transparent plot area
)

# ============================== Dash Application ========================== #

app = dash.Dash(__name__)
server= app.server

app.layout = html.Div(
    children=[ 
    html.Div(
        className='divv', 
        children=[ 
        html.H1(
            f'BMHC MarCom Report {current_quarter} 2025', 
            className='title'),
        html.H2(
            '01/01/2025 - 3/31/2025', 
            className='title2'),
        html.Div(
            className='btn-box', 
            children=[
                html.A(
                'Repo',
                href='https://github.com/CxLos/MC_Q2_2025',
                className='btn'),
            ]),
    ]),    

# Data Table
# html.Div(
#     className='row0',
#     children=[
#         html.Div(
#             className='table',
#             children=[
#                 html.H1(
#                     className='table-title',
#                     children='Data Table'
#                 )
#             ]
#         ),
#         html.Div(
#             className='table2', 
#             children=[
#                 dcc.Graph(
#                     className='data',
#                     figure=marcom_table
#                 )
#             ]
#         )
#     ]
# ),

# ROW 1
html.Div(
    className='row0',
    children=[
        html.Div(
            className='graph11',
            children=[
            html.Div(
                className='high1',
                children=[f'{current_quarter} MarCom Events']
            ),
            html.Div(
                className='circle1',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high3',
                            children=[marcom_events]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
        html.Div(
            className='graph22',
            children=[
            html.Div(
                className='high2',
                children=[f'{current_quarter} MarCom Hours']
            ),
            html.Div(
                className='circle2',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high4',
                            children=[marcom_hours]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
    ]
),

# ROW 1
html.Div(
    className='row0',
    children=[
        html.Div(
            className='graph11',
            children=[
            html.Div(
                className='high1',
                children=[f'{current_quarter} MarCom Travel Time']
            ),
            html.Div(
                className='circle1',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high3',
                            children=[travel_time]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
        html.Div(
            className='graph22',
            children=[
            html.Div(
                className='high2',
                children=[f'{current_quarter} Blank']
            ),
            html.Div(
                className='circle2',
                children=[
                    html.Div(
                        className='hilite',
                        children=[
                            html.H1(
                            className='high4',
                            children=[]
                    ),
                        ]
                    )
 
                ],
            ),
            ]
        ),
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=hours_fig
                )
            ]
        )
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=hours_pie
                )
            ]
        )
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=activity_fig
                )
            ]
        )
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=activity_pie
                )
            ]
        )
    ]
),

# # ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=product_pie
                )
            ]
        )
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=purpose_pie
                )
            ]
        )
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=person_fig
                )
            ]
        )
    ]
),

# ROW 
html.Div(
    className='row3',
    children=[
        html.Div(
            className='graph0',
            children=[
                dcc.Graph(
                    figure=pf_pie
                )
            ]
        )
    ]
),

# ROW 2
html.Div(
    className='row1',
    children=[
        html.Div(
            className='graph1',
            children=[
                html.Div(
                    className='table',
                    children=[
                        html.H1(
                            className='table-title',
                            children='Products Table'
                        )
                    ]
                ),
                html.Div(
                    className='table2', 
                    children=[
                        dcc.Graph(
                            className='data',
                            figure=products_table
                        )
                    ]
                )
            ]
        ),
        html.Div(
            className='graph2',
            children=[                
              html.Div(
                    className='table',
                    children=[
                        html.H1(
                            className='table-title',
                            children='Purpose Table'
                        )
                    ]
                ),
                html.Div(
                    className='table2', 
                    children=[
                        dcc.Graph(
                            className='data',
                            figure=purpose_table
                        )
                    ]
                )
   
            ]
        )
    ]
),

# ROW 4
html.Div(
    className='row1',
    children=[
        html.Div(
            className='graph1',
            children=[
                dcc.Graph(
                  figure = status_fig
                )
            ]
        ),
        html.Div(
            className='graph2',
            children=[                
                dcc.Graph(
                    figure = status_pie
                )
            ]
        )
    ]
),
])

print(f"Serving Flask app '{current_file}'! 🚀")

if __name__ == '__main__':
    app.run_server(debug=
                   True)
                #    False)
# =================================== Updated Database ================================= #

# updated_path = f'data/MarCom_{current_quarter}_{report_year}.xlsx'
# data_path = os.path.join(script_dir, updated_path)
# df.to_excel(data_path, index=False)
# print(f"DataFrame saved to {data_path}")

# updated_path1 = 'data/service_tracker_q4_2024_cleaned.csv'
# data_path1 = os.path.join(script_dir, updated_path1)
# df.to_csv(data_path1, index=False)
# print(f"DataFrame saved to {data_path1}")

# -------------------------------------------- KILL PORT ---------------------------------------------------

# netstat -ano | findstr :8050
# taskkill /PID 24772 /F
# npx kill-port 8050

# ---------------------------------------------- Host Application -------------------------------------------

# 1. pip freeze > requirements.txt
# 2. add this to procfile: 'web: gunicorn impact_11_2024:server'
# 3. heroku login
# 4. heroku create
# 5. git push heroku main

# Create venv 
# virtualenv venv 
# source venv/bin/activate # uses the virtualenv

# Update PIP Setup Tools:
# pip install --upgrade pip setuptools

# Install all dependencies in the requirements file:
# pip install -r requirements.txt

# Check dependency tree:
# pipdeptree
# pip show package-name

# Remove
# pypiwin32
# pywin32
# jupytercore

# ----------------------------------------------------

# Name must start with a letter, end with a letter or digit and can only contain lowercase letters, digits, and dashes.

# Heroku Setup:
# heroku login
# heroku create mc-impact-11-2024
# heroku git:remote -a mc-impact-11-2024
# git push heroku main

# Clear Heroku Cache:
# heroku plugins:install heroku-repo
# heroku repo:purge_cache -a mc-impact-11-2024

# Set buildpack for heroku
# heroku buildpacks:set heroku/python

# Heatmap Colorscale colors -----------------------------------------------------------------------------

#   ['aggrnyl', 'agsunset', 'algae', 'amp', 'armyrose', 'balance',
            #  'blackbody', 'bluered', 'blues', 'blugrn', 'bluyl', 'brbg',
            #  'brwnyl', 'bugn', 'bupu', 'burg', 'burgyl', 'cividis', 'curl',
            #  'darkmint', 'deep', 'delta', 'dense', 'earth', 'edge', 'electric',
            #  'emrld', 'fall', 'geyser', 'gnbu', 'gray', 'greens', 'greys',
            #  'haline', 'hot', 'hsv', 'ice', 'icefire', 'inferno', 'jet',
            #  'magenta', 'magma', 'matter', 'mint', 'mrybm', 'mygbm', 'oranges',
            #  'orrd', 'oryel', 'oxy', 'peach', 'phase', 'picnic', 'pinkyl',
            #  'piyg', 'plasma', 'plotly3', 'portland', 'prgn', 'pubu', 'pubugn',
            #  'puor', 'purd', 'purp', 'purples', 'purpor', 'rainbow', 'rdbu',
            #  'rdgy', 'rdpu', 'rdylbu', 'rdylgn', 'redor', 'reds', 'solar',
            #  'spectral', 'speed', 'sunset', 'sunsetdark', 'teal', 'tealgrn',
            #  'tealrose', 'tempo', 'temps', 'thermal', 'tropic', 'turbid',
            #  'turbo', 'twilight', 'viridis', 'ylgn', 'ylgnbu', 'ylorbr',
            #  'ylorrd'].

# rm -rf ~$bmhc_data_2024_cleaned.xlsx
# rm -rf ~$bmhc_data_2024.xlsx
# rm -rf ~$bmhc_q4_2024_cleaned2.xlsx