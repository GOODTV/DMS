<!--#include file="../include/dbfunctionJ.asp"-->
<%
  SQL1="Select CONTRIBUTE_ISSUEDATA.*,Storage_Qty=Isnull(GOODS.Goods_Qty,0),Goods_IsStock=GOODS.Goods_IsStock From CONTRIBUTE_ISSUEDATA Left Join GOODS On CONTRIBUTE_ISSUEDATA.Goods_Id=GOODS.Goods_Id Where CONTRIBUTE_ISSUEDATA.Ser_No='"&request("ser_no")&"'"
  Call QuerySQL(SQL1,RS1)
  If Not RS1.EOF Then
    Goods_Name=RS1("Goods_Name")
    If RS1("Contribute_IsStock")="Y" Then
      SQL2="Select * From GOODS Where Goods_Id='"&RS1("Goods_Id")&"'"
      Set RS2 = Server.CreateObject("ADODB.RecordSet")
      RS2.Open SQL2,Conn,1,3
      If RS2("Goods_IsStock")="Y" Then
        RS2("Goods_Qty")=Cdbl(RS2("Goods_Qty"))+Cdbl(RS1("Goods_Qty"))
        RS2.Update
      End If
      RS2.Close
      Set RS2=Nothing
    End If
    SQL="Delete CONTRIBUTE_ISSUEDATA Where Ser_No='"&request("Ser_No")&"'"
    Set RS=Conn.Execute(SQL)
    session("errnumber")=1
    session("msg")="『"&Goods_Name&"』 刪除成功 ！"
  End If
  RS1.Close
  Set RS1=Nothing
  response.redirect "contribute_issue_edit.asp?Issue_Id="&request("issue_id")&""
%><!--#include file="../include/dbclose.asp"-->