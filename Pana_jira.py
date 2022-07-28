# -*- coding: utf-8 -*- 

###########################################################################
## Created by Ramesh Kankatala for extracting the worklog from pana jira
###########################################################################
import keyword
from pickle import FALSE, TRUE
from sched import scheduler
from timeit import repeat
from jinja2 import Undefined
from pip import main
from schedule import every,repeat,run_pending
import schedule
from turtle import width
from openpyxl import load_workbook
import win32com.client
import wx
import wx.xrc
import calendar
from datetime import date, datetime, timedelta
import matplotlib.pyplot as plt
import copy
import datetime
import hashlib
import imghdr
import json
import logging
import mimetypes
import os
import re
import sys
import time
from apscheduler.schedulers.background import BackgroundScheduler
import warnings
import numpy as np
import pandas as pd
import xlsxwriter
import requests
from time import sleep
import win32com.client as win32
from pkg_resources import parse_version
from requests import Response
from requests.auth import AuthBase
from requests.utils import get_netrc_auth
from jira import JIRA, JIRAError
from jira import __version__
from jira.resources import (
    Attachment,
    Board,
    Comment,
    Component,
    Customer,
    CustomFieldOption,
    Dashboard,
    Filter,
    GreenHopperResource,
    Group,
    Issue,
    IssueLink,
    IssueLinkType,
    IssueType,
    Priority,
    Project,
    RemoteLink,
    RequestType,
    Resolution,
    Resource,
    Role,
    SecurityLevel,
    ServiceDesk,
    Sprint,
    Status,
    User,
    Version,
    Votes,
    Watchers,
    Worklog,
)
from jira.utils import CaseInsensitiveDict, json_loads, threaded_requests
global_user_list = ['BP11115','BP14199','BP11185','BP15119','BP14858','BP11168','BP14504','BP11579','BP11897','BP14507','BP15253','BP10893']

def connecttoPANAjira():
    
    global jira_data
    # Using current time
    ini_time_for_now = datetime.datetime.now()
    # printing initial_date
    #print("initial_date", str(ini_time_for_now))
    present_date,present_time = str(ini_time_for_now).split(" ")
    new_final_time = ini_time_for_now + \
    timedelta(days=-7)
    # printing new final_date
    #print("new_final_time", str(new_final_time))
    previous_date,previous_time = str(new_final_time).split(" ")
    jira_username = 'BP14507'
    jira_password = 'Ficosavldc@001'
    #jira_username = input("\nEnter  global ID:")
    #jira_password = input("\nEnter password for jira:")
    jira = JIRA('https://jira.pase.panasonic.de/', auth=(jira_username, jira_password))
    projects = jira.projects()
    jql_query = "project in ('DAI05') AND issuetype in ('ERR', 'Task', 'Sub-task','ERR ext') AND created > '-1000d'"
    jira_search = jira.search_issues(jql_query, startAt=0,maxResults=80000,fields="issuetype,worklog,created")
    df_data = [];
    for each_issue in jira_search:
        log_entry_count = len(each_issue.fields.worklog.worklogs)
        for each_entry in range(log_entry_count):
            str_user_convert = str(each_issue.fields.worklog.worklogs[each_entry].author)
            issue_key = str(each_issue.key)
            author = str(each_issue.fields.worklog.worklogs[each_entry].author)
            dateoflogged = str(each_issue.fields.worklog.worklogs[each_entry].updated)
            noofhourslogged = str(each_issue.fields.worklog.worklogs[each_entry].timeSpent)
            logged_date,extra_field = dateoflogged.split('T')
            #df_data.append([author, issue_key, dateoflogged, noofhourslogged])
            
            for user in global_user_list:
                if user in str_user_convert and (logged_date >= previous_date and logged_date <=present_date):
                    jira_data = author + "," + issue_key + "," + dateoflogged + "," + noofhourslogged
                    df_data.append([author, issue_key, dateoflogged, noofhourslogged])
    df = pd.DataFrame(data = df_data, columns=["Author", "IsueKey", "Date of Logged", "No of hours Logged"])
    user_names = list(df['Author'].unique())
    writer = pd.ExcelWriter('userdata.xlsx', engine='xlsxwriter')
    for user in global_user_list:
        df_user = df.loc[df['Author'].str.contains(user),:]
        df_user.to_excel(writer, sheet_name=user, index=False)
        print("Workbook is created with sheet name as :",user)
    writer.save()

    schedule.every().friday.at('06:00').do(connecttoPANAjira)
    return


def main():
    connecttoPANAjira()

if __name__ == "__main__":
   main()

while True:
    schedule.run_pending()
    time.sleep(1)