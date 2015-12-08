<!--#include file="../include/dbfunctionJ.asp"-->
<%
Menu1_Value=""
Menu2_Value=""

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
    <form name="form" method="post" action="goods_list.asp" target="detail">
      <input type="hidden" name="action">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="780"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
              <tr>
                <td width="5%"> </td>
                <td width="95%">
  		            <table width="40%"  border="0" cellspacing="0" cellpadding="0">
		                <tr>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                      <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                    </tr>
                    <tr>
                      <td class="table62-bg">&nbsp;</td>
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;"><%=Prog_Desc%></td>
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
                  <table width="100%" border="0" cellpadding="1" cellspacing="1">
                    <tr>
                      <td class="td02-c" width="70" align="right">物品代號：</td>
                      <td class="td02-c" width="550">
                        <input type="text" name="Goods_Id" size=9" class="font9" maxlength="10" onkeyup="javascript:UCaseGoodsId();" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
                        &nbsp;
                        物品名稱：
                        <input type="text" name="Goods_Name" size="24" class="font9" maxlength="40" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
                        &nbsp;
                        庫存管理：
                        <Select name='Goods_IsStock' Size='1' style='font-size: 9pt; font-family: 新細明體'>
                        	<option value=''></option>
                        	<option value='Y'>是</option>
				                  <option value='N'>否</option>
                        </Select>
                        &nbsp;
                        <input type="checkbox" name="Goods_Qty" value='Y'>現有庫存量>0
                      </td>
                      <td class="td02-c" width="180">
                        <input type="button" value=" 查 詢 " name="query" class="cbutton" style="cursor:hand" onClick="Query_OnClick()">&nbsp;
                        <input type="button" value=" 列 表 " name="report" class="cbutton" style="cursor:hand" onClick="Report_OnClick()">&nbsp;
                        <input type="button" value=" 匯 出 " name="report" class="cbutton" style="cursor:hand" onClick="Export_OnClick()">	
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">物品性質：</td>
                      <td class="td02-c">
                        <%=Sel_Menu1%>
                      	&nbsp;
                      	物品類別：
                      	<%=Sel_Menu2%>
                      </td>
                      <td class="td02-c"> </td>
                    </tr>
                    <tr>
                      <td class="td02-c" colspan="3" width="100%"><iframe name="detail" height="380" width="100%" frameborder="0" scrolling="auto" src="goods_list.asp"></iframe></td>
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
function Query_OnClick(){
  document.form.target="detail"
  document.form.action.value="query"
  document.form.submit();
}
function Report_OnClick(){
  if(confirm('您是否確定要將查詢結果列表？')){
    document.form.target="report"
    document.form.action.value="report"
    window.open('','report','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=800,height=600')
    document.form.submit();
  }
}

function Export_OnClick(){
  if(confirm('您是否確定要將查詢結果匯出？')){
    document.form.target='main';
    document.form.action.value='export';
    document.form.submit();
  }
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
--></script>