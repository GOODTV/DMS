<!--#include file="../include/dbfunctionJ.asp"-->
<%
Gift_Date=Date()
If request("action")="save" Then
'  Check_Close=Get_Close("2","C001",Cstr(request("Gift_Date")),Cstr(Session("user_id")))
'  If Check_Close Then     
'    If Ckeck_GoodID Then
      '取捐物編號
'      Invoice_Pre=""
'      Invoice_No=""
'      If request("Issue_Type")="M" And request("Invoice_No")<>"" Then
'        Invoice_No=Trim(request("Invoice_No"))
'      Else
'        Act_id=""
'        If request("Act_Id")<>"" Then Act_id=Cstr(request("Act_Id"))
'        InvoiceNo=Get_Invoice_No2("2",Cstr(request("Dept_Id")),Cstr(request("Contribute_Date")),Cstr(request("Invoice_Type")),Act_id)
'        If InvoiceNo<>"" Then
'          Invoice_Pre=Split(InvoiceNo,"/")(0)
'          Invoice_No=Split(InvoiceNo,"/")(1)
'        End If
'      End If
    
      '新增公關贈品資料
      SQL1="GIFT"
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,3
      RS1.Addnew
      RS1("Donor_id")=request("Donor_Id")
      RS1("Gift_Date")=request("Gift_Date")
      RS1("Gift_Payment")=request("Gift_Payment")      
      RS1("Dept_Id")=request("Dept_Id")       
      If request("Act_id")<>"" Then
       RS1("Act_id")=request("Act_id")
      Else
        RS1("Act_id")=null
      End If
      RS1("Comment")=request("Comment")            
      RS1("Create_Date")=Date()
      RS1("Create_User")=session("user_name")
      RS1("Create_IP")=Request.ServerVariables("REMOTE_HOST")
      RS1.Update
      RS1.Close
      Set RS1=Nothing
    
      SQL1="Select @@IDENTITY As Gift_Id"
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      Gift_id=RS1("Gift_Id")
      RS1.Close
      Set RS1=Nothing
       
      '新增公關贈品明細
      For I=1 To Cint(request("Goods_Max"))
        If request("Goods_Name_"&I)<>"" Then
          SQL1="GIFTDATA"
          Set RS1 = Server.CreateObject("ADODB.RecordSet")
          RS1.Open SQL1,Conn,1,3
          RS1.Addnew
          RS1("Gift_Id")=Gift_Id
          RS1("Donor_id")=request("Donor_Id")
          RS1("Goods_Id")=request("Goods_Id_"&I)
          RS1("Goods_Name")=request("Goods_Name_"&I)
          If request("Goods_Qty_"&I)<>"" Then
            RS1("Goods_Qty")=request("Goods_Qty_"&I)
          Else
            RS1("Goods_Qty")="1"
          End If          
          RS1("Goods_Comment")=request("Goods_Comment_"&I)         
          RS1("Create_Date")=Date()
          RS1("Create_User")=session("user_name")
          RS1("Create_IP")=Request.ServerVariables("REMOTE_HOST")
          RS1.Update
          RS1.Close
          Set RS1=Nothing                  
        End If
      Next     
      
      session("errnumber")=1
      session("msg")="資料新增成功 ！"
      If request("ctype")="gift_data" Then
        Response.Redirect "contribute_data.asp?donor_id="&request("donor_id")
      ElseIf request("ctype")="member_contribute_data" Then
        Response.Redirect "member_contribute_data.asp?donor_id="&request("donor_id")  
      Else
        Response.Redirect "contribute_detail.asp?contribute_id="&contribute_id&"&ctype="&request("ctype")
      End If
'    End If
'  Else
'    session("errnumber")=1
'    session("msg")="您輸入的捐贈日期『 "&Cstr(request("Contribute_Date"))&"』 已關帳無法新增 ！"
'  End If  
End If

SQL1="Select * From DEPT Where Dept_Id='"&session("dept_id")&"'"
Call QuerySQL(SQL1,RS1)
Goods_Max=RS1("Goods_Max")
'IsStock=RS1("IsStock")
RS1.Close
Set RS1=Nothing

SQL1="Select Goods_Name From dbo.GIFTDATA Where Donor_Id='"&request("donor_id")&"' Group By Goods_Name"
Call QuerySQL(SQL1,RS1)
GoodsName_List = ""
While Not RS1.EOF
	GoodsName_List = GoodsName_List & RS1("Goods_Name") & "|"
	RS1.MoveNext
Wend
RS1.Close
Set RS1=Nothing

SQL="Select *,Address2=(Case When DONOR.City='' Then DONOR.Address Else Case When A.mValue<>B.mValue Then DONOR.ZipCode+A.mValue+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End), " & _
    "Invoice_Address2=(Case When DONOR.Invoice_City='' Then DONOR.Invoice_Address Else Case When C.mValue<>B.mValue Then DONOR.Invoice_ZipCode+C.mValue+D.mValue+DONOR.Invoice_Address Else C.mValue+DONOR.Invoice_ZipCode+DONOR.Invoice_Address End End) " & _
    "From DONOR " & _
    "Left Join CodeCity As A On DONOR.City=A.mCode " & _
    "Left Join CodeCity As B On DONOR.Area=B.mCode " & _ 
    "Left Join CodeCity As C On DONOR.Invoice_City=C.mCode " & _
    "Left Join CodeCity As D On DONOR.Invoice_Area=D.mCode " & _
    "Where Donor_Id='"&request("donor_id")&"'"
Call QuerySQL(SQL,RS)
%>
<%Prog_Id="gift"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ctype" value="<%=request("ctype")%>">
      <input type="hidden" name="Goods_Max" value="<%=Goods_Max%>">      
      <input type="hidden" name="Donor_Id" value="<%=RS("donor_id")%>">
      <input type="hidden" name="DonorName" value="<%=RS("Donor_Name")%>">
      <input type="hidden" name="DonorIDNo" value="<%=RS("IDNo")%>">
      <input type="hidden" name="GoodsName_List" value="<%=GoodsName_List%>">
      
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="780" border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">公關贈品輸入</td>
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
                        <table border="0" cellpadding="2" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
                            <td align="right" width="15%">捐款人：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Donor_Name" size="30" class="font9t" readonly value="<%=Data_Minus(RS("Donor_Name"))%>&nbsp;<%=RS("Title")%>">                              
                            </td>
                          </tr>                          
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td width="12%" align="right">日期：</td>
                            <td width="16%" align="left"><%call Calendar("Gift_Date",Gift_Date)%></td>
                            <input type="hidden" name="Gift_Payment" value="公關贈品">                                                      
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <input type="hidden" name="Dept_ID" value="C001">                          
                          <tr>
                            <td align="right">備註：</td>
                            <td align="left" colspan="3"><textarea rows="3" name="Comment" cols="45" class="font9"><%=request("Comment")%></textarea></td>                            
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td width="100%" colspan="8"> 
                              <table border="0" width="100%" cellpadding="2" cellspacing="2">
                                <tr>                                	
                                	<td width="20%">名稱</td>
                                	<td width="11%">數量</td>                                	
                                	<td width="24%">備註</td>
                                	<td width="5%" align="center">清除</td>
                                </tr>
                                <%For I=1 To Goods_Max%>
                                <tr>
                                	<input type="hidden" name="Goods_Id_<%=I%>" value="<%=request("Goods_Id_"&I)%>">                                	
                                	<!--20131003 Modify by GoodTV Tanya:修改windows.open位置-->
                                	<td><input type="text" name="Goods_Name_<%=I%>" size="18" maxlength="50" class="font9" value="<%=request("Goods_Name_"&I)%>">&nbsp;<a href onClick="window.open('gift_show.asp?LinkId=Goods_Id_<%=I%>&LinkName=Goods_Name_<%=I%>&LinkUnit=Goods_Unit_<%=I%>','','status=no,scrollbars=yes,top=200,left=600,width=500,height=450')" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17"></a></td>
                                	<td><input type="text" name="Goods_Qty_<%=I%>" size="10" maxlength="7" class="font9" value="<%=request("Goods_Qty_"&I)%>"></td>                                	
                                	<td><input type="text" name="Goods_Comment_<%=I%>" size="26" maxlength="100" class="font9" value="<%=request("Goods_Comment_"&I)%>"></td>
                                  <td align="center"><img border="0" src="../images/toobar_cancle.gif" width="20" onClick="Contribute_Cancel_OnClick(<%=I%>)" style="cursor:hand"></td>
                                </tr>
                                <%Next%>
                              </table>	
                            </td>
                          </tr>
                          <!--#include file="../include/calendar2.asp"-->
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
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
	for(i=1;i<=Number(document.form.Goods_Max.value);i++){
		document.getElementById('Goods_Name_'+i).style.backgroundColor='#ffffcc';
  	document.getElementById('Goods_Name_'+i).readOnly=true;
  }
      
  document.form.Gift_Date.focus();
}
function Contribute_Cancel_OnClick(i){
  if(document.getElementById('Goods_Name_'+i).value!=''){
    if(confirm('您是否確定要清除『 '+document.getElementById('Goods_Name_'+i).value+' 』？')){
      document.getElementById('Goods_Id_'+i).value='';
      document.getElementById('Goods_Name_'+i).value='';
      document.getElementById('Goods_Qty_'+i).value='';      
      document.getElementById('Goods_Comment_'+i).value='';
    }
  }else{
    if(confirm('您是否確定要清除？')){
      document.getElementById('Goods_Id_'+i).value='';
      document.getElementById('Goods_Name_'+i).value='';
      document.getElementById('Goods_Qty_'+i).value='';      
      document.getElementById('Goods_Comment_'+i).value='';
    }
  }
}
function Save_OnClick(){
  <%call CheckStringJ("Gift_Date","日期")%>
  <%call CheckDateJ("Gift_Date","日期")%>    
  
  var GoodsName_List = document.form.GoodsName_List.value;
  var Goods_Name = "";  
  for(i=1;i<=Number(document.form.Goods_Max.value);i++){    	
  	Goods_Name = document.getElementById('Goods_Name_'+i).value;  	  	
  	if (Goods_Name!='' && GoodsName_List.indexOf(Goods_Name) > -1)
  	{  		
  		if (!confirm('『'+Goods_Name+'』公關品已重複是否確定要新增？'))
  			return false;
  	}
  }  
  
  <%call SubmitJ("save")%>
}
function Cancel_OnClick(){
  if(document.form.ctype.value=='gift_data'){
    location.href='contribute_data.asp?donor_id='+document.form.Donor_Id.value+'';
  }else if(document.form.ctype.value=='member_contribute_data'){
    location.href='member_contribute_data.asp?donor_id='+document.form.Donor_Id.value+'';    
  }else{
    location.href='contribute.asp';
  }
}
--></script>