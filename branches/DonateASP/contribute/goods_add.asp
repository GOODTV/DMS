<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
  SQL1="Select * From GOODS Where Goods_Id='"&request("Goods_Id")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  If RS1.EOF then
    SQL2="GOODS"
    Set RS2 = Server.CreateObject("ADODB.RecordSet")
    RS2.Open SQL2,Conn,1,3
    RS2.Addnew
    RS2("Goods_Id")=request("Goods_Id")
    RS2("Goods_Name")=request("Goods_Name")
    RS2("Goods_Property")=request("Goods_Property")
    RS2("Goods_Type")=request("Goods_Type")
    RS2("Goods_IsStock")=request("Goods_IsStock")
    RS2("Goods_Qty")="0"
    RS2("Goods_Unit")=request("Goods_Unit")
    RS2("Remark")=request("Remark")
    RS2.Update
    RS2.Close
    Set RS1=Nothing
    session("errnumber")=1
    session("msg")="物品主檔資料新增成功 ！"
    response.redirect "goods.asp"  
  Else
    session("errnumber")=1
    session("msg")="您輸入的物品代號已經存在 ！"
  End If
End If

Menu1_Value=request("Goods_Property")
Menu2_Value=request("Goods_Type")

Str_Menu2=""
Sel_Menu1="<Select name='Goods_Property' Size='1' style='font-size: 9pt; font-family: 新細明體' OnChange='JavaScript:ChgProperty(this.value,document.form.Goods_Type);'><option value=''></option>"
SQL1="Select Ser_No,Menu1=Linked_Name From LINKED Where Linked_Type In ('contribute') Order By Linked_Seq"
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,1
While Not RS1.EOF
  If Cstr(RS1("Menu1"))=Cstr(Menu1_Value) Then
    Sel_Menu1=Sel_Menu1&"<option value='"&RS1("Menu1")&"' Selected >"&RS1("Menu1")&"</option>"
  Else
    Sel_Menu1=Sel_Menu1&"<option value='"&RS1("Menu1")&"'>"&RS1("Menu1")&"</option>"
  End If
  If Str_Menu2="" Then
    Str_Menu2="'"&RS1("Menu1")&"'"&","
  Else
    Str_Menu2=Str_Menu2&","&"'"&RS1("Menu1")&"'"&","
  End If
  Str_Menu2=Str_Menu2&"["
  Menu2_Row=0
  SQL2="Select Menu2=Linked2_Name From LINKED2 Where Linked_Id='"&RS1("Ser_No")&"' Order By Linked2_Seq"
  Set RS2 = Server.CreateObject("ADODB.RecordSet")
  RS2.Open SQL2,Conn,1,1
  While Not RS2.EOF
    If Menu2_Row=0 Then
      Str_Menu2=Str_Menu2&"'"&RS2("Menu2")&"'"
    Else
      Str_Menu2=Str_Menu2&","&"'"&RS2("Menu2")&"'"
    End If
    Menu2_Row=Menu2_Row+1
    RS2.MoveNext
  Wend
  RS2.Close
  Set RS2=Nothing
  Str_Menu2=Str_Menu2&"]"
  Response.Flush
  Response.Clear
  RS1.MoveNext
Wend
RS1.Close
Set RS1=Nothing
Sel_Menu1=Sel_Menu1&"</Select>"

Sel_Menu2="<Select name='Goods_Type' Size='1' style='font-size: 9pt; font-family: 新細明體'><option value=''></option>"
%>
<%Prog_Id="goods"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3" size="20">
        <tr>
          <td>
            <table width="780" border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
              <tr>
                <td width="5%"></td>
                <td width="95%">
  		            <table width="40%"  border="0" cellspacing="0" cellpadding="0">
		                <tr>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                      <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                    </tr>
                    <tr>
                      <td class="table62-bg">&nbsp;</td>
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;"><%=Prog_Desc%>【新增】</td>
                      <td class="table63-bg">&nbsp;</td>
                    </tr>  
    	            </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td> 
	          <table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
              </tr>
              <tr>
                <td class="table62-bg">&nbsp;</td>
                <td valign="top">
                  <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                    <tr>
                      <td class="td02-c" width="100%">
                        <table border="0" cellpadding="1" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
                            <td align="right" width="12%">物品代號：</td>
                            <td align="left" width="88%"><input type="text" name="Goods_Id" size="15" class="font9" maxlength="10" value="<%=request("Goods_Id")%>" onkeyup="javascript:UCaseGoodsId();"></td>
                          </tr>
                          <tr>
                            <td align="right">物品名稱：</td>
                            <td align="left"><input type="text" name="Goods_Name" size="35" class="font9" maxlength="50" value="<%=request("Goods_Name")%>"></td>
                          </tr>
                          <tr>
                            <td align="right">物品性質：</td>
                            <td align="left"><%=Sel_Menu1%></td>
                          </tr>
                          <tr>
                            <td align="right">物品類別：</td>
                            <td align="left"><%=Sel_Menu2%></td>
                          </tr>
                          <tr>
                            <td align="right">庫存單位：</td>
                            <td align="left"><input type="text" name="Goods_Unit" size="10" class="font9" maxlength="10" value="<%=request("Goods_Unit")%>"></td>
                          </tr>
                          <tr>
                            <td align="right">是否庫存管理：</td>
                            <td align="left">
                              <input type="radio" name="Goods_IsStock" value="Y" <%If request("Goods_IsStock")<>"N" Then%>checked<%End If%> >是 
					                    <input type="radio" name="Goods_IsStock" value="N" <%If request("Goods_IsStock")="N" Then%>checked<%End If%> >否
                            </td>
                          </tr>
                          <tr>
                            <td align="right">備註：</td>
                            <td align="left">
                            	<textarea name="Remark" rows="4" cols="35" class="font9"><%=request("Remark")%></textarea>
                            </td>
                          </tr>
                          <tr>
                            <td align="center" colspan="8">
                               <input type="button" value=" 存 檔 " name="save" class="cbutton" style="cursor:hand" onClick="Save_OnClick()">
                               &nbsp;&nbsp;
                               <input type="button" value=" 取 消 " name="cancel" class="cbutton" style="cursor:hand" onClick="Cancel_OnClick()">
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
                <td class="table63-bg">&nbsp;</td>
              </tr>
              <tr>
                <td style="background-color:#EEEEE3"><img src="../images/table06_06.gif" width="10" height="10"></td>
                <td class="table64-bg"><img src="../images/table06_07.gif" width="1" height="10"></td>
                <td style="background-color:#EEEEE3"><img src="../images/table06_08.gif" width="10" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>   
    </form>
  </center></div></p>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Window_OnLoad(){
  document.form.Goods_Id.focus();
}	
function UCaseGoodsId(){
  document.form.Goods_Id.value=document.form.Goods_Id.value.toUpperCase();
}
function ChgProperty(Menu1Value,Menu2Obj){
  ClearOption(Menu2Obj);
  var AryMenu2=new Array(<%=Str_Menu2%>);
  for(var i=0;i<=AryMenu2.length-1;i=i+2){
    if(AryMenu2[i]==Menu1Value){
      var VarMenu2=AryMenu2[i+1];
      var k=1;
      for(var j=0;j<=VarMenu2.length-1;j++){
        Menu2Obj.options[k]=new Option(VarMenu2[j],VarMenu2[j]);
        k++;
      }
    }
  }
}
function ClearOption(SelObj){
  for(var i=SelObj.length-1;i>=0;i--){
    SelObj.options[i]=null;
  }
}
function Save_OnClick(){
  <%call CheckStringJ("Goods_Id","物品代號")%>
  <%call ChecklenJ("Goods_Id",10,"物品代號")%>
  <%call CheckStringJ("Goods_Name","物品名稱")%>
  <%call ChecklenJ("Goods_Name",50,"物品名稱")%>
  <%call CheckStringJ("Goods_Property","物品性質")%>
  <%call ChecklenJ("Goods_Property",20,"物品性質")%>
  <%call CheckStringJ("Goods_Type","物品類別")%>
  <%call ChecklenJ("Goods_Type",20,"物品類別")%>
  <%call CheckStringJ("Goods_Unit","物品單位")%>
  <%call ChecklenJ("Goods_Unit",10,"物品單位")%>  
  <%call SubmitJ("save")%>
}
function Cancel_OnClick(){
  location.href='goods.asp';
}
--></script>