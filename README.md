# Reading Progress Tracking using Google Sheet

The goal of this project is to track the reading progress. All users fill a google form. And the data is submitted to a google sheet. We then run the local post-processing code to analyze the data and then push it back to the google sheet. 

The data we collect from the users have the following fields: 

* 您的姓名 -- Name (e.g., 张三)

*    召会 -- Church (e.g., Chicago)

* 文集年限 -- year of the book (e.g., 1955)

* 文集卷目 -- volumn (e.g., 1)

* 阅读页码 -- range of the pages you read (10-25)

* 信息题目及享受的点 -- the title of the message and what you have enjoyed. 

The post processing tools is basically fetch the data and calculate the number of pages that each user read, then push the summary information back to google drive by creating a worksheet in the same google spreadsheet. The summary information contains the number of pages each user read in each month of the current year. 

To enable this I have performed the following work to

* Enabling Google Drive API and creating a crediencial following the example. 
https://towardsdatascience.com/accessing-google-spreadsheet-data-using-python-90a5bc214fd2

* Install google API python library 
$ pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

The whole project was created for tracking the reading of "Collective Work of Witness Lee". https://www.livingstream.com/en/33-collected-works-of-witness-lee. But it could be easily generalized to other form tracking or data processing based on Google Drive. 

I execute the python script periodically to update the data. I used adscheduler python module. 

I create a Automator job following the instruction: https://superuser.com/questions/229773/run-command-on-startup-login-mac-os-x

* Start Automator.app;
* Select "Application";
* Click "Show library" in the toolbar (if hidden);
* Add "Run shell script" (from the Actions/Utilities);
"#!/bin/bash
python progress_update.py
"
* Test it;
* Save it somewhere: a file called your_name.app will be created);
* Depending your MacOSX version:
* Old versions: Go to System Preferences → Accounts → Login items, or
* New version: Go to System Preferences → Users and Groups → Login items (top right);
* Add this newly-created app;
Log off, log back in, and you should be done. ;)
