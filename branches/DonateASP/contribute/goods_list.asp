<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="goods"%>
<!--#include file="../include/head.asp"-->
<body class=tool>
<%
If request("SQL")="" Then 
  SQL="Select Goods_Id,物品代號=Goods_Id,物品名稱=Goods_Name,物品性質=Goods_Property,物品類別=Goods_Type,現有庫存量=Goods_Qty,庫存單位=Goods_Unit,庫存管理=(Case When Goods_IsStock='Y' Then 'V' Else '' End) " & _
      "From GOODS Join LINKED On GOODS.Goods_Property=LINKED.Linked_Name Join LINKED2 On GOODS.Goods_Type=LINKED2.Linked2_Name Where LINKED.Linked_Type='contribute' "
  If request("Goods_Id")<>"" Then SQL=SQL&"And Goods_Id Like '%"&request("Goods_Id")&"%' "
  If request("Goods_Name")<>"" Then SQL=SQL&"And Goods_Name Like '%"&request("Goods_Name")&"%' "
  If request("Goods_Property")<>"" Then SQL=SQL&"And Goods_Property Like '%"&request("Goods_Property")&"%' "
  If request("Goods_Type")<>"" Then SQL=SQL&"And Goods_Type Like '%"&request("Goods_Type")&"%' "
  If request("Goods_IsStock")<>"" Then SQL=SQL&"And Goods_IsStock='"&request("Goods_IsStock")&"' "
  If request("Goods_Qty")<>"" Then SQL=SQL&"And Goods_Qty>0 "
  SQL=SQL&"Order By LINKED.Linked_Seq,Linked2_Seq,Goods_Id Desc"
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="goods_list"
HLink="goods_edit.asp?goods_id="
LinkParam="goods_id"
LinkTarget="main"
AddLink="goods_add.asp"
Hit_Nemu=""
If request("action")="stop" Then
  call GridList_S (AddLink)
Else
  If request("action")="report" Or request("action")="export" Then Server.Transfer "goods_rpt.asp"
  call GridListHit (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink,Hit_Nemu)
End If  
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->