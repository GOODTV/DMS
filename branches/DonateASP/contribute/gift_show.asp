<!--#include file="../include/dbfunctionJ.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>物品查詢</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=tool>
<%
  SQL1="Select * From LINKED Where Linked_Type='gift' " & _
       "Order By LINKED.Linked_Seq "	   
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  If Not RS1.EOF Then
    Response.Write "<table border='0' width='100%'><tr><td height='25'><a href='gift_show.asp?LinkId="&request("LinkId")&"&LinkName="&Request("LinkName")&"&LinkUnit="&Request("LinkUnit")&";self.close()'>全部</a>‧"
    While Not RS1.EOF
      Response.Write "<a href='gift_show.asp?goods_type="&RS1("Linked_Name")&"&LinkId="&Request("LinkId")&"&LinkName="&Request("LinkName")&"&LinkUnit="&Request("LinkUnit")&";self.close()'>"&RS1("Linked_Name")&"</a>‧"
      Response.Flush
      Response.Clear
      RS1.MoveNext
    Wend	
    Response.Write "</td></tr></table>"
    
    SQL2="Select 物品代號=Linked_Id,物品名稱=Linked2_Name " & _
         "From LINKED2 Join Linked On Linked.Ser_No=LINKED2.Linked_Id Where Linked.Linked_Type='gift' "    	
    If request("Goods_Type")<>"" Then SQL2=SQL2&"And Linked_Name Like '%"&request("Goods_Type")&"%' "    	
    SQL2=SQL2&"Order By Linked2_Seq"
    Set RS2 = Server.CreateObject("ADODB.RecordSet")
    RS2.Open SQL2,Conn,1,1
    If Not RS2.EOF Then
      Response.Write "<table border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
      'Response.Write "<a href='javascript:opener.document.form."&LinkField&".value="&RS2("Goods_Id")&";self.close()'><tr>"
      For I = 0 To RS2.Fields.Count-1
        Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS2(I).Name & "</span></font></td>"
      Next
      Response.Write "</tr>"	
      While Not RS2.EOF
	      Response.Write "<a href='javascript:opener.document.form."&request("LinkId")&".value="""&RS2(0)&""";opener.document.form."&request("LinkName")&".value="""&RS2(1)&""";opener.document.form."&request("LinkUnit")&".value=""1"";self.close()'><tr style='cursor:hand' bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
	      For I = 0 To RS2.Fields.Count-1
	        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&RS2(I)&"</span></td>"
        Next
        Response.Write "</tr></a>"
        Response.Flush
        Response.Clear
        RS2.MoveNext
      Wend
      Response.Write "</table>"
      RS2.Close
      Set RS2=Nothing
    End If
  Else
    Response.Write "物品主檔尚未建立  ！"	
  End If
  RS1.Close
  Set RS1=Nothing
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->