# PCS Email Service

This program is a part of the Loopy system. It handles the monitoring service for the PCS source files and sending the reporting emails to DDOT and other data subscribers. 

### Dependencies 

  - [Python 2.7]
  - [HTML.py] - Formatting html tables for email reports
  - [xlrd] - Parsing excel sheets

### Note 

The scripts should be scheduled daily. The results for the monitor process would be stored on DDOT drives, and the email service will send out the results to the subscribers. 

[Python 2.7]: <https://www.python.org/download/releases/2.7/>
[HTML.py]: <http://www.decalage.info/python/html>
[xlrd]: <https://pypi.python.org/pypi/xlrd>