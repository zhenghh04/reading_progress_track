# Reading Progress Tracking using Google Sheet

The goal of this project is to track the reading progress. All users fill a google form. And the data is submitted to a google sheet. We then run the local post-processing code to analyze the data and then push it back to the google sheet. 

In the google form, there are data as:

* 您的姓名 -- Name (e.g., 张三)

* 文集年限 -- year of the book (e.g., 1955)

* 文集卷目 -- volumn (e.g., 1)

* 阅读页码 -- range of the pages you read (10-25)

* 信息题目及享受的点 -- the title of the message and what you have enjoyed. 

The post processing tools is basically fetch the data and calculate the number of pages that each user read, then push the summary information back to google drive by creating a worksheet in the same google spreadsheet. 

To enable this I have perform the following work to

* Enabling Google Drive API and creating a crediencial following the example. 
https://towardsdatascience.com/accessing-google-spreadsheet-data-using-python-90a5bc214fd2

* Install google API python library 
$ pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib


