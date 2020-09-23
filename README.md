<div align="center">

## Compact Access DB using ASP


</div>

### Description

This code snippet will compact your Access DB online. You can do this via a web browser!
 
### More Info
 
Path to database


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Tropeano](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-tropeano.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-tropeano-compact-access-db-using-asp__4-7403/archive/master.zip)





### Source Code

```
<%
Option Explicit
Const THEJETVAR= 4
Function Squish(thePathDB, boolIs97)
Dim fso, Engine, strThePathDB
strThePathDB = left(thePathDB,instrrev(ThePathDB,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(thePathDB) Then
   Set Engine = CreateObject("JRO.JetEngine")
   If boolIs97 = "True" Then
       Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & thePathDB, _
       "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strThePathDB & "temp.mdb;" _
       & "Jet OLEDB:Engine Type=" & JET_3X
   Else
 Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & thePathDB, _
 "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strThePathDB & "temp.mdb"
    End If
    fso.CopyFile strThePathDB & "temp.mdb",thePathDB
    fso.DeleteFile(strThePathDB & "temp.mdb")
    Set fso = nothing
    Set Engine = nothing
    Squish = "Your database, " & thePathDB & ", has been Compacted" & vbCrLf
Else
    Squish = "The database name or path has not been found. Try Again" & vbCrLf
End If
End Function
%>
<html><head><title>Compact Database</title></head><body>
<h2 align="center"> Compacting Dealer database</h2>
<p align="center">
<form action=compact.asp>
Enter relative path to the database, including database name.<br><br>
<input type="text" name="thePathDB" value="/data/dealers.mdb">
<br><br>
<input type="checkbox" name="boolIs97" value="True"> Check if Access 97 database
<br><i> (default is Access 2000)</i><br><br>
<input type="submit">
<form>
<br><br>
<%
Dim thePathDB,boolIs97
thePathDB = request("thePathDB")
boolIs97 = request("boolIs97")
If thePathDB <> "" Then
    thePathDB = server.mappath(thePathDB)
    response.write(Squish(thePathDB,boolIs97))
End If
%>
</p></body></html>
```

