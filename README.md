Resources Used / Sources Cited

Credit Card Activity from Class - The solved credit card script was a great resource to reference for tackling some of the tasks at hand with this challenge and provided a great starting point.
Xpert AI - I utilized this resource a lot to help with troubleshooting typos, understanding error messages, etc.
Megan Romano - T.A. - Megan was of great assistance through office hours and Slack messaging. She helped me better understand the tasks at hand and guided me through understanding troubleshooting my issues.
AskBCS - I tried AskBCS a couple of times for assistance. It was mostly unhelpful, though one LA (I believe his name was Mohamed) took the time to get on Zoom and walking through my code with him helped me further understand what I was trying to accomplish.



The following links provided information that helped me complete this task:

https://www.thespreadsheetguru.com/last-row-column-vba/
# Dim LastRow As Long
# LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
## these lines were provided by this link


https://corporatefinanceinstitute.com/resources/excel/vba-variable-types/#:~:text=Long%3A%20The%20Long%20data%20type,range%20of%20%2D2%2C147%2C483%2C648%20to%202%2C147%2C483%2C648.
https://corporatefinanceinstitute.com/resources/excel/vba-variables-dim/
https://stackoverflow.com/questions/39134913/handling-big-number-in-vba
# between these 3 links, I was able to determine the issue I was having with the volume variables needing to be "Double" not "Long"
