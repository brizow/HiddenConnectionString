Why I Used This:
I needed this little file for a project while working with our Lead Production 
scheduler. We leveraged an Excel spreadsheet she was already using to connect
to SQl, pull down job status data, and populate which process the scheduled
jobs were in. This job status data came from the Job Tracking System I 
developed. 

This would help her drop jobs off the scheduling sheet without 
having to run to the floor to see where the job packets (paper) where physically.
Sometimes the walk about sessions would take more than an hour as she would 
have to check with each operator or cell lead to determine where that particular
job made it after 3rd shift. 

How To Use:
Details of how this file works available here: http://goo.gl/oy6Brz

Short Details:
This little C# module allows you to connect a VBA macro to a database using an 
ADODB connection string. There is no way to hide connection string data from 
within VBA so, leveraging this you can keep things secure and simple. 

Calling the mycn.GetConnection method allows you to connect to whichever server
you have specified from within the method. It is a .dll file, which is quite
challenging to decode as well. An added layer of security for you.
