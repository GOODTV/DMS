<%Response.ContentType="text/html; charset=utf-8"%>
<%
Sub Message()
  If session("errnumber")=0 then
    Response.Write "<center>"&session("msg")&"</center>"
  Else
    Response.Write "<script Language='JavaScript'> alert('" & session("msg")&"')</script>"
  End If
  session("msg")=""
  session("errnumber")=0
End Sub
Function Check_Post()
  Check_Post = True
  URL1 = Cstr(Request.ServerVariables("HTTP_REFERER"))
  URL2 = Cstr(Request.ServerVariables("SERVER_NAME"))
  If Mid(URL1,8,Len(URL2))<>URL2 Then Check_Post = False
End Function

Function Sys_Log (User_ID,User_Name,Dept_Id,Log_Type,Log_Desc)
  On Error Resume Next
  SQL1="Sys_Log"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  RS1.Addnew
  RS1("Log_Time")=Now()
  RS1("Log_Date")=Date()
  RS1("User_Id")=User_ID
  RS1("User_Name")=User_Name
  RS1("Dept_Id")=Dept_Id
  RS1("Log_Type")=Log_Type
  RS1("Log_Desc")=Log_Desc
  RS1("IP_Address")=Request.ServerVariables("REMOTE_ADDR")
  RS1.Update
  RS1.Close
  Set RS1=Nothing
end Function

Function User_Lock (User_ID)
  On Error Resume Next
  SQL1="Select * From userfile Where user_id='"&User_ID&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  If Not RS.EOF Then
    RS1("user_lock")="Y"
    RS1.Update
  End If
  RS1.Close
  Set RS1=Nothing
End Function
%>