import pandas as pd

import numpy as np
[262]
#selecting the file to open

viewed_content = pd.read_excel('../Viewed Content/ViewedContent_Feb23.xlsx')

viewed_content.head()

[263]
#eliminating unnecessary columns

viewed_content.drop(['Id', 'Url', 'Last Viewed Date (Coordinated Universal Time)', 'Group Id', 'Application Id'], inplace=True, axis=1)

viewed_content.head()

[264]
#renaming column(s)

viewed_content.columns=['Title', 'User Id', 'Username', 'Type', 'date', 'Total View Count', 'Group Name', 'Application Name']

viewed_content.head()

[317]
#filtering by Group Name (in this case, I'm filtering by a Group Name containing 'Developers')

devs_views = viewed_content[viewed_content['Group Name'].str.contains('.*Developers', na=False)]

devs_views

[318]
#using the describe function to find out the quartiles.
#calculating the IQR (Interquartile range) Q3 - Q1: 15 - 29 - 7 = 22
#multiplying the IQR by 1.5: 22 * 1.5 = 33
#finding out the outlier limit: 33 + 29 = 62

devs_views.describe()

[319]
#filtering out all rows with a Total View Count above 62, according to the calculations above

devs_views = devs_views[(devs_views['Total View Count']<=62)]

devs_views

[320]
#generating the Excel file

writer = pd.ExcelWriter('Views_devsFeb2023.xlsx',
                       engine='xlsxwriter',
                       engine_kwargs={'options': {'strings_to_urls':False}}
                       )
devs_views.to_excel(writer)

writer.close()