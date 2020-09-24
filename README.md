<div align="center">

## Start your application when windows starts


</div>

### Description

Many people want to know how to start there application when windows starts and not simply by putting a shortcut in the windows startup folder.This article will show you how to do this using vb and a few api's
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Venkat Mani](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/venkat-mani.md)
**Level**          |Beginner
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/venkat-mani-start-your-application-when-windows-starts__1-23276/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Start Application as windows starts</title>
</head>
<body>
<p>We shall use a few api function which are&nbsp;&nbsp;<br>
<br>
<br>
<b>Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long<br>
<br>
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long</b><br>
<br>
<br>
<br>
<br>
These api's are used for the following purpose<br>
<br>
<br>
1.RegOpenKey - To open a key for reading/writing values&nbsp;<br>
2.RegSetValue - To write values into a key<br>
<br>
<br>
We shall also use a few constants&nbsp;<br>
<br>
Private Const HKEY_CURRENT_USER = &amp;H80000001<br>
<br>
Private Const REG_SZ = 1<br>
<br>
<br>
In order to make our applications start when windows starts we have to add an entry in the&nbsp;<br>
<br>
HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run<br>
<br>
<br>
<br>
To make an entry into a key we need to get the 'handle' or a unique identifier for that key.We get this 'unique id by opening the key.We do this in the following way<br>
<br>
<b>Dim result As Long<br>
Dim keyres As Long<br>
result=RegOpenKey(HKEY_CURRENT_USER,&quot;Software\Microsoft\Windows\CurrentVersion\Run&quot;,keyres)<br>
</b><br>
<br>
If the function executed correctly we will get 0 as result and keyres will contain the unique id for that key.<br>
<br>
After opening the key we will put in a value into it .We can do it this way<br>
<br>
<br>
<b>Dim file As String<br>
Dim entry as string<br>
entry = "Myprog"<br>
file = "c:\myprog\myprog.exe"<br>
result = RegSetValueEx(keyres, entry, 0, REG_SZ, ByVal file, Len(file))</b><br>
<br>
<br>
you can input your program's path into file and run the function.And entry is the name you give to the value you are trying to put in.If the function executed successfully result will contain 0 .Restart your computer and your program should start as soon as windows starts.Send your comments to<a href="mailto:venky_dude@yahoo.com">
venky_dude@yahoo.com </a> .Visit my <a href="http://www.geocities.com/venky_dude/venkwork.htm"> homepage</a>
for some cool vb codes. <br>
<br>
</p>
</body>
</html>

