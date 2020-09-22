<div align="center">

## Access' queries like stored procedures


</div>

### Description

Here is HOWTO call access' select,insert,update and delete queries from asp.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ATM](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/atm.md)
**Level**          |Intermediate
**User Rating**    |4.3 (56 globes from 13 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/atm-access-queries-like-stored-procedures__4-6127/archive/master.zip)





### Source Code

```
'Create new Access Database, C:\test.mdb
'
'Create table:
'TestTable
'--------------------
'id   autonumber
'field1 text    len=50
'field2 number
'--------------------
'
'CreateFour queries:
'Query1: DeleteQuery, Type:Delete
'--------------------------------
'DELETE [testTable].[field1], [testTable].[field2]
'FROM testTable
'WHERE id=[:param1];
'--------------------------------
'Query2: InsertQuery, Type:Insert
'--------------------------------
'INSERT INTO testTable ( field1, field2 )
'VALUES ([:param1], [:param2]);
'--------------------------------
'Query3: SelectByid, Type:Normal
'--------------------------------
'SELECT [testTable].[field1], [testTable].[field2], [testTable].[id]
'FROM testTable
'WHERE ((([testTable].[id])>=[:param1] And ([testTable].[id])<=[:param2]));
'-------------------------------
'Query4: UpdateByid, Type:Update
'--------------------------------
'UPDATE testTable SET field1 = [:param1], field2 = [:param2]
'WHERE id=[:param3];
'Calls:
'accessTest.asp?defMode=0 - for test select
'
'accessTest.asp?defMode=1 - for test insert
'
'accessTest.asp?defMode=2 - for test delete
'
'accessTest.asp?defMode=3 - for test update
'
'
'Here is accessTest.asp.
'---------------------------------------
<!--#include file="adovbs.inc"-->
<%
'for one query we give max 60sec
execTimeout=60
'we have 1 query will be executed
execCount=1
'max +5 sec for script to execute
Server.ScriptTimeout=execCount*execTimeout+5
'connection parameters
dbPath="DBQ=C:\Test.mdb;"
userData="UID=;PWD="
conn_string="PROVIDER=MSDASQL;" & _
       "DRIVER={Microsoft Access Driver (*.mdb)};" &_
       dbPath &_
       userData
'connection object
Set connObj=Server.CreateObject("ADODB.Connection")
'make connection
connObj.Open conn_string
'command object
Set commandObj=Server.CreateObject("ADODB.Command")
commandObj.ActiveConnection=connObj
commandObj.CommandTimeout=execTimeout
commandObj.CommandType=adCmdStoredProc
Select Case request("defMode")
'test Select
case "0"
'------------------------------------------------------------------------
'our select query is
'Name: SelectByid
commandObj.CommandText="SelectByid"
'
'SELECT testTable.field1, testTable.field2, testTable.id
'FROM testTable
'WHERE (((testTable.id)=>:param1)) and (((testTable.id)<=:param2));
'create parameters for query
commandObj.Parameters.Append commandObj.CreateParameter("param1", _
                            adInteger, _
                            adParamInput, _
                            10, _
                            1)
commandObj.Parameters.Append commandObj.CreateParameter("param2", _
                            adInteger, _
                            adParamInput, _
                            10, _
                            100)
'create recordset object
Set rsObj=Server.CreateObject("ADODB.Recordset")
rsObj.CursorType=1 'forwardonly
'run query
rsObj.Open commandObj
response.write("<TABLE>")
response.write("<TR><TD>ID</TD><TD>Field1</TD><TD>Field2</TD></TR>")
If not(rsObj.EOF) then
 Do While Not(rsObj.EOF)
  response.write("<TR><TD>")
  response.write(rsObj("id"))
  response.write("<TD>")
  response.write(rsObj("field1"))
  response.write("</TD><TD>")
  response.write(rsObj("field2"))
  response.write("</TD></TR>")
  rsObj.MoveNext
 Loop
 'close recordset
 rsObj.Close
End if
response.write("</TABLE>")
'deallocate rs object
Set rsObj=Nothing
'delete allocated parameters
commandObj.Parameters.Delete "param2"
commandObj.Parameters.Delete "param1"
'------------------------------------------------------------------------
'test Insert
case "1"
'------------------------------------------------------------------------
'Name: Insert
commandObj.CommandText="InsertQuery"
'INSERT INTO testTable ( field1, field2 )
'VALUES ([:param1], [:param2]);
commandObj.Parameters.Append commandObj.CreateParameter(":param1", _
                            adVarchar, _
                            adParamInput, _
                            50, _
                            "1")
commandObj.Parameters.Append commandObj.CreateParameter(":param2", _
                            adInteger, _
                            adParamInput, _
                            10, _
                            2)
i=0
Do While i<5
 commandObj(":param1")=chr(65+i)
 commandObj(":param2")=(65+i)
 commandObj.Execute
 i=i+1
Loop
'delete allocated parameters
commandObj.Parameters.Delete ":param2"
commandObj.Parameters.Delete ":param1"
'------------------------------------------------------------------------
'test Delete
case "2"
'------------------------------------------------------------------------
'Name: DeleteQuery
commandObj.CommandText="DeleteQuery"
'DELETE FROM testTable
'WHERE id=[:param1];
commandObj.Parameters.Append commandObj.CreateParameter(":param1", _
                            adInteger, _
                            adParamInput, _
                            10, _
                            "1")
i=0
Do While i<5
 '!!!id can be different!!!
 commandObj(":param1")=14
 commandObj.Execute
 i=i+1
Loop
'delete allocated parameters
commandObj.Parameters.Delete ":param1"
'------------------------------------------------------------------------
'test Update
case "3"
'------------------------------------------------------------------------
'Name: UpdateByid
commandObj.CommandText="UpdateByid"
'UPDATE testTable SET field1 = [:param1], field2 = [:param2]
'WHERE id=[:param3];
commandObj.Parameters.Append commandObj.CreateParameter(":param1", _
                            adVarchar, _
                            adParamInput, _
                            50, _
                            "Z")
commandObj.Parameters.Append commandObj.CreateParameter(":param2", _
                            adInteger, _
                            adParamInput, _
                            10, _
                            0)
commandObj.Parameters.Append commandObj.CreateParameter(":param3", _
                            adInteger, _
                            adParamInput, _
                            10, _
                            0)
i=0
Do While i<5
 commandObj(":param1")="14"
 commandObj(":param2")=14
 '!!!id can be different!!!
 commandObj(":param3")=15
 commandObj.Execute
 i=i+1
Loop
'delete allocated parameters
commandObj.Parameters.Delete ":param3"
commandObj.Parameters.Delete ":param2"
commandObj.Parameters.Delete ":param1"
'------------------------------------------------------------------------
End Select
'deallocate commandObj
Set commandObj=Nothing
'close connection
connObj.Close
'deallocate connection object
Set connObj=Nothing
response.end
%>
'--------------------------------------
```

