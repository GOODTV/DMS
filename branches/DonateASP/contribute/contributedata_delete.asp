<!--#include file="../include/dbfunctionJ.asp"-->
<%
  SQL1="Select CONTRIBUTEDATA.*,Storage_Qty=Isnull(GOODS.Goods_Qty,0),Goods_IsStock=GOODS.Goods_IsStock From CONTRIBUTEDATA Left Join GOODS On CONTRIBUTEDATA.Goods_Id=GOODS.Goods_Id Where CONTRIBUTEDATA.Ser_No='"&request("ser_no")&"'"
  Call QuerySQL(SQL1,RS1)
  If Not RS1.EOF Then
    Donor_Id=RS1("Donor_Id")
    Goods_Name=RS1("Goods_Name")
    Check_Delete=True
    If RS1("Contribute_IsStock")="Y" Then
      If Cdbl(RS1("Goods_Qty"))>Cdbl(RS1("Storage_Qty")) Then
        If RS1("Goods_IsStock")="Y" Then
          Check_Delete=False
          session("errnumber")=1
          session("msg")="『"&Goods_Name&"』 庫存量不足無法刪除 ！"
        End If
      Else
        SQL2="Select * From GOODS Where Goods_Id='"&RS1("Goods_Id")&"'"
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,3
        If RS2("Goods_IsStock")="Y" Then
          RS2("Goods_Qty")=Cdbl(RS2("Goods_Qty"))-Cdbl(RS1("Goods_Qty"))
          RS2.Update
        End If
        RS2.Close
        Set RS2=Nothing
      End If
    End If
    If Check_Delete Then
      SQL="Delete CONTRIBUTEDATA Where Ser_No='"&request("Ser_No")&"'"
      Set RS=Conn.Execute(SQL)
      
      SQL="Update CONTRIBUTE Set Contribute_Amt=(Select Isnull(Sum(Goods_Amt),0) From CONTRIBUTEDATA Where Contribute_Id='"&request("contribute_id")&"') Where Contribute_Id='"&request("contribute_id")&"'"
      Set RS=Conn.Execute(SQL)
      
      '修改捐贈人捐物紀錄
      call Declare_DonorId (Donor_Id)
    
      session("errnumber")=1
      session("msg")="『"&Goods_Name&"』 刪除成功 ！"
    End If
  End If
  RS1.Close
  Set RS1=Nothing
  response.redirect "contribute_edit.asp?ctype="&request("ctype")&"&contribute_id="&request("contribute_id")&""
%><!--#include file="../include/dbclose.asp"-->