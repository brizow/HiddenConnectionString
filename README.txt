Detailed, details of how this file works available here: http://goo.gl/oy6Brz

This little C# module allows you to connect a VBA macro to a database using an 
ADODB connection string. There is no way to hide connection string data from 
within VBA so, leveraging this you can keep things secure and simple. 

Calling the mycn.GetConnection method allows you to connect to whichever server
you have specified from within the method. It is a .dll file, which is quite
challenging to decode as well. An added layer of security for you.
