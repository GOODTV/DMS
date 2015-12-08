<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function GiftDataGrid (SQL,HLink,HLink2,LinkParam)
  Row=0
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table id=grid border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
  Response.Write "<tr>"
  Response.Write "<td align=""center"" bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>序號</span></font></td>"
  For I = 1 To FieldsCount
	  Response.Write "<td align=""center"" bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  Response.Write "<td align=""center"" bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>編輯</span></font></td>"
  Response.Write "<td align=""center"" bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>刪除</span></font></td>"
  Response.Write "</tr>"
  While Not RS1.EOF
	  Response.Write "<tr style='cursor:hand' bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
	  For I = 1 To FieldsCount
	    If I=1 Then
        Row=Row+1
        Response.Write "<td align=""center""><span style='font-size: 9pt; font-family: 新細明體'>" & Row & "</span></td>"
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I) & "</span></td>"
      ElseIf I=2 Then
        Response.Write "<td align=""right""><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(I),0) & "</span></td>"          
      Else
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I) & "</span></td>"
      End If
    Next
    Response.Write "<td align=""center""><a href='#' onclick=""window.open('"&HLink&RS1(LinkParam)&"','','status=no,scrollbars=no,top=100,left=120,width=400,height=170')""><span style='font-size: 9pt; font-family: 新細明體'>編輯</span></a></td>"
    Response.Write "<td align=""center""><a href='JavaScript:if(confirm(""是否確定要刪除『"&RS1(1)&"』 ?"")){window.location.href="""&HLink2&RS1(LinkParam)&""";}' target='main'><img src='../images/x5.gif' border=0 width='16' height='14' alt='刪除'></a></td>" 
    RS1.MoveNext
    Response.Write "</tr>"
  Wend  
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function

If request("action")="update" Then
'  Check_Close=Get_Close("2","C001",Cstr(request("Contribute_Date")),Cstr(Session("user_id")))
'  If Check_Close Then
    Ckeck_GoodID=True
    For I=1 To Cint(request("Goods_Max"))
      If request("Goods_Name_"&I)<>"" Then      
        SQL1="Select * From GIFTDATA Where Gift_Id='"&request("Gift_Id")&"' And Goods_Id='"&request("Goods_Id_"&I)&"' And Goods_Name='"&request("Goods_Name_"&I)&"' "
        Call QuerySQL(SQL1,RS1)
        If Not RS1.EOF Then
          Ckeck_GoodID=False
          session("errnumber")=1
          session("msg")="您輸入的『 "&request("Goods_Name_"&I)&" 』 物品名稱重覆出現"
          Exit For
        End If
        RS1.Close
        Set RS1=Nothing
      End If
    Next

    If Ckeck_GoodID Then
      '修改公關品資料
      SQL1="Select * From GIFT Where Gift_Id='"&request("Gift_Id")&"'"
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,3
      RS1("Donor_id")=request("Donor_Id")
      RS1("Gift_Date")=request("Gift_Date")
      RS1("Gift_Payment")=request("Gift_Payment")                     
      RS1("Comment")=request("Comment")           
      RS1.Update
      RS1.Close
      Set RS1=Nothing
      
      '新增公關品明細
      For I=1 To Cint(request("Goods_Max"))
        If request("Goods_Name_"&I)<>"" Then
          SQL1="GIFTDATA"
          Set RS1 = Server.CreateObject("ADODB.RecordSet")
          RS1.Open SQL1,Conn,1,3
          RS1.Addnew
          RS1("Gift_Id")=request("Gift_Id")
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
      session("msg")="公關品資料修改成功 ！"                 
    End If
'  Else
'    session("errnumber")=1
'    session("msg")="您輸入的捐贈日期『 "&Cstr(request("Donate_Date"))&"』 已關帳無法修改 ！"
'  End If  
End If

If request("action")="delete" Then
'  Check_Close=Get_Close("2","C001",Cstr(request("Contribute_Date")),Cstr(Session("user_id")))
'  If Check_Close Then
    
		SQL="Delete From GIFT Where Gift_Id='"&request("Gift_Id")&"' " & _
		    "Delete From GIFTDATA Where Gift_Id='"&request("Gift_Id")&"' "
		Set RS=Conn.Execute(SQL)  

  	session("errnumber")=1
  	session("msg")="資料刪除成功 ！"
	  If request("ctype")="gift_data" Then
	    Response.Redirect "contribute_data.asp?donor_id="&request("donor_id")
	  ElseIf request("ctype")="member_contribute_data" Then
	    Response.Redirect "member_contribute_data.asp?donor_id="&request("donor_id")  
	  Else
	    Response.Redirect "contribute.asp"
	  End If
   
'  Else
'    session("errnumber")=1
'    session("msg")="您輸入的捐贈日期『 "&Cstr(request("Donate_Date"))&"』 已關帳無法刪除 ！"
'  End If  
End If

SQL="Select GIFT.*,Donor_Name=DONOR.Donor_Name,Title=DONOR.Title,Donor_Type=DONOR.Donor_Type,IDNo=DONOR.IDNo, " & _
    "Address2=(Case When DONOR.City='' Then DONOR.Address Else Case When A.mValue<>B.mValue Then DONOR.ZipCode+A.mValue+B.mValue+DONOR.Address Else DONOR.ZipCode+A.mValue+DONOR.Address End End) " & _
    "From GIFT " & _
    "Join DONOR On GIFT.Donor_Id=DONOR.Donor_Id " & _
    "Left Join CodeCity As A On DONOR.City=A.mCode " & _
    "Left Join CodeCity As B On DONOR.Area=B.mCode " & _
    "Where Gift_Id='"&request("gift_id")&"' "
Call QuerySQL(SQL,RS)

SQL1="Select * From DEPT Where Dept_Id='"&RS("Dept_Id")&"'"
Call QuerySQL(SQL1,RS1)
Goods_Max=RS1("Goods_Max")
RS1.Close
Set RS1=Nothing

SQL1="Select Goods_Name From dbo.GIFTDATA Where Donor_Id='"&RS("donor_id")&"' AND Gift_Id <> '"&request("gift_id")&"' Group By Goods_Name"
Call QuerySQL(SQL1,RS1)
GoodsName_List = ""
While Not RS1.EOF
	GoodsName_List = GoodsName_List & RS1("Goods_Name") & "|"
	RS1.MoveNext
Wend
RS1.Close
Set RS1=Nothing
%>
<%Prog_Id="gift"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ctype" value="<%=request("ctype")%>">
      <input type="hidden" name="Goods_Max" value="<%=Goods_Max%>">           
      <input type="hidden" name="Dept_Id" value="<%=RS("Dept_Id")%>">		
      <input type="hidden" name="Gift_Id" value="<%=RS("Gift_Id")%>">	
      <input type="hidden" name="Donor_Id" value="<%=RS("donor_id")%>">
      <input type="hidden" name="DonorName" value="<%=RS("Donor_Name")%>">
      <input type="hidden" name="DonorIDNo" value="<%=RS("IDNo")%>">      
      <input type="hidden" name="Gift_Payment" value="公關贈品">
      <input type="hidden" name="Dept_ID" value="C001">      
      <input type="hidden" name="GoodsName_List" value="<%=GoodsName_List%>">  
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">公關贈品【修改】</td>
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
                            <td align="right">捐款人：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Donor_Name" size="30" class="font9t" readonly value="<%=Data_Minus(RS("Donor_Name"))%>&nbsp;<%=RS("Title")%>">    
                              &nbsp;身分證/統編：
                              <input type="text" name="IDNo" size="10" class="font9t" readonly value="<%=RS("IDNo")%>">                          
                              &nbsp;身份別：
                              <input type="text" name="Donor_Type" size="40" class="font9t" readonly value="<%=RS("Donor_Type")%>">
                            </td>
                          </tr>                         
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td width="10%" align="right">贈送日期：</td>
                            <td width="16%" align="left" colspan="3"><%call Calendar("Gift_Date",RS("Gift_Date"))%></td>                            
                          </tr>                         
                          <tr>
                            <td align="right">備註：</td>
                            <td align="left" colspan="5"><textarea rows="3" name="Comment" cols="45" class="font9"><%=RS("Comment")%></textarea></td>                            
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="right">贈品內容：</td>
                            <td align="left" colspan="7">
                            <%
                              SQL="Select Ser_No,物品名稱=Goods_Name,數量=Goods_Qty,備註=Goods_Comment " & _
                                  "From GIFTDATA Where Gift_Id='"&RS("gift_id")&"' Order By Ser_No "
                              HLink="giftdata_edit.asp?ctype="&request("ctype")&"&gift_id="&RS("Gift_Id")&"&ser_no="
                              HLink2="giftdata_delete.asp?ctype="&request("ctype")&"&gift_id="&RS("Gift_Id")&"&ser_no="
                              LinkParam="Ser_No"
                              Call GiftDataGrid (SQL,HLink,HLink2,LinkParam)
                            %>
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td width="100%" colspan="9"> 
                              <table border="0" width="100%" cellpadding="2" cellspacing="2">
                                <tr>
                                	<td width="20%"><font color="#ff0000">續增公關品</font></td>
                                	<td width="11%">物品數量</td>                                	
                                	<td width="24%">備註</td>
                                	<td width="5%" align="center">清除</td>
                                </tr>
                                <%For I=1 To Goods_Max%>
                                <tr>
                                	<input type="hidden" name="Goods_Id_<%=I%>" value="<%=request("Goods_Id_"&I)%>">                                	
                                	<!--20131003 Modify by GoodTV Tanya:修改windows.open位置-->
                                	<td><input type="text" name="Goods_Name_<%=I%>" size="18" maxlength="50" class="font9">&nbsp;<a href onClick="window.open('gift_show.asp?LinkId=Goods_Id_<%=I%>&LinkName=Goods_Name_<%=I%>&LinkUnit=Goods_Unit_<%=I%>','','status=no,scrollbars=yes,top=200,left=600,width=500,height=450')" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17"></a></td>
                                	<td><input type="text" name="Goods_Qty_<%=I%>" size="10" maxlength="7" class="font9"></td>                                	
                                	<td><input type="text" name="Goods_Comment_<%=I%>" size="26" maxlength="100" class="font9"></td>
                                  <td align="center"><img border="0" src="../images/toobar_cancle.gif" width="20" onClick="Gift_Cancel_OnClick(<%=I%>)" style="cursor:hand"></td>
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
                               <input type="button" value=" 修 改 " name="save" class="cbutton" style="cursor:hand" onClick="Update_OnClick()">
                               &nbsp;&nbsp;
                               <input type="button" value=" 刪 除 " name="save" class="cbutton" style="cursor:hand" onClick="Delete_OnClick()">
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
  
  document.form.Contribute_Date.focus();
}

function Gift_Cancel_OnClick(i){
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

function Update_OnClick(){
  <%call CheckStringJ("Gift_Date","贈送日期")%>
  <%call CheckDateJ("Gift_Date","贈送日期")%>  
  
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
  
  <%call SubmitJ("update")%>
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
    location.href='gift_detail.asp?ctype='+document.form.ctype.value+'&gift_id='+document.form.Gift_Id.value+'';
}
--></script>