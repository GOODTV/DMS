<%
'搜尋條件
SQL_Where=""
If Request("Dept_Id")<>"" Then
  SQL_Where=SQL_Where&"And Donate.Dept_Id = '"&Request("Dept_Id")&"' "
Else
  SQL_Where=SQL_Where&"And Donate.Dept_Id In ("&Session("all_dept_type")&") "
End If
If Request("Donate_Date_Begin")<>"" Then SQL_Where=SQL_Where&"And Donate_Date>='"&Request("Donate_Date_Begin")&"' "
If Request("Donate_Date_End")<>"" Then SQL_Where=SQL_Where&"And Donate_Date<='"&Request("Donate_Date_End")&"' "
If Request("Act_Id")<>"" Then SQL_Where=SQL_Where&"And Donate.Act_Id='"&Request("Act_Id")&"' "

Session("action")=request("action")
Session("Col")=request("Col")
Session("keyword")=request("keyword")
Session("maxlen")=request("maxlen")
Session("Dept_Id")=request("Dept_Id")
Session("Donate_Date_Begin")=request("Donate_Date_Begin")
Session("Donate_Date_End")=request("Donate_Date_End")
Session("Act_Id")=request("Act_Id")
Session("Name_Type")=request("Name_Type")
Session("Amt_Type")=request("Amt_Type")
Session("Purpose_Type")=request("Purpose_Type")
Session("Seq_Type")=request("Seq_Type")
Session("SQL_Where")=SQL_Where
If request("action")="report" Then

ElseIf request("action")="export" Then

ElseIf request("action")="txt" Then
  Response.Redirect "madazine_list_rpt_style_"&Request("Amt_Type")&"_"&Request("Purpose_Type")&".asp"
End If  
%>