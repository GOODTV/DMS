<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="member" Then
  If request("IsMemberPre")<>"" Then
    SQL1="Select TOP 1 Member_No From DONOR Where Member_No Like '"&request("IsMemberPre")&"%' Order By Member_No Desc"
  Else
    SQL1="Select TOP 1 Member_No From DONOR Order By Member_No Desc"
  End If
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  If Not RS1.EOF Then
    If request("IsMemberPre")<>"" Then Member_No=Replace(RS1("Member_No"),request("IsMemberPre"),"")
    Member_No=Cstr(Cdbl(Member_No)+1)
    Member_No=request("IsMemberPre")&Left("00000",5-Len(Member_No))&Member_No
  Else
    Member_No=request("IsMemberPre")&"00001"
  End If
  RS1.Close
  Set RS1=Nothing
  SQL="Update DONOR Set Member_No='"&Member_No&"',IsMember='Y',Member_Status='服務中' Where Donor_Id='"&request("donor_id")&"'"
  Set RS=Conn.Execute(SQL)
  
  '確認會員編號無重覆
  Member_No_Old=Member_No
  Check_MemberNo=False
  While Check_MemberNo=False
    SQL1="Select * From DONOR Where Member_No='"&Member_No_Old&"' And Donor_Id<>'"&request("donor_id")&"' "
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    If RS1.EOF Then Check_MemberNo=True
    RS1.Close
    Set RS1=Nothing
    If Check_MemberNo=Flase Then
      If request("IsMemberPre")<>"" Then
        SQL1="Select TOP 1 Member_No From DONOR Where Member_No Like '"&request("IsMemberPre")&"%' Order By Member_No Desc"
      Else
        SQL1="Select TOP 1 Member_No From DONOR Order By Member_No Desc"
      End If
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      If Not RS1.EOF Then
        If request("IsMemberPre")<>"" Then Member_No_Old=Replace(RS1("Member_No"),request("IsMemberPre"),"")
        Member_No_Old=Cstr(Cdbl(Member_No_Old)+1)
        Member_No_Old=request("IsMemberPre")&Left("00000",5-Len(Member_No_Old))&Member_No_Old
      Else
        Member_No_Old=request("IsMemberPre")&"00001"
      End If
      RS1.Close
      Set RS1=Nothing
      SQL="Update DONOR Set Member_No='"&Member_No_Old&"',IsMember='Y' Where Donor_Id='"&request("donor_id")&"'"
      Set RS=Conn.Execute(SQL)
    End If
  Wend
  session("errnumber")=1
  session("msg")="已將捐款人轉為本會會員 ！"
  Response.Redirect "../member/member_edit.asp?donor_id="&request("donor_id")
End If

If request("action")="update" Then
  Address=""
  If request("IsAbroad")="N" Then
    Address=request("Street")
    If request("SectionType")<>"" Then Address=Address&request("SectionType")&"段"
    If request("Lane")<>"" Then Address=Address&request("Lane")&"巷"
    If request("Alley")<>"" Then Address=Address&request("Alley")&request("AlleyType")
    If request("HouseNo")<>"" Then
      If request("HouseNoSub")<>"" Then
        Address=Address&request("HouseNo")&"之"&request("HouseNoSub")&"號"
      Else
        Address=Address&request("HouseNo")&"號"
      End If
    End If
    If request("Floor")<>"" Then
      If request("FloorSub")<>"" Then
        Address=Address&request("Floor")&"樓之"&request("FloorSub")
      Else
        Address=Address&request("Floor")&"樓"
      End If
    End If
    If request("Room")<>"" Then Address=Address&" "&request("Room")&"室"
  Else
    Address=Data_Plus(request("OverseasAddress"))&" "&Data_Plus(request("OverseasCountry"))
  End If

  Invoice_Address=""
  If request("IsAbroad_Invoice")="N" Then
    Invoice_Address=request("Invoice_Street")
    If request("Invoice_Section")<>"" Then Invoice_Address=Invoice_Address&request("Invoice_Section")&"段"
    If request("Invoice_Lane")<>"" Then Invoice_Address=Invoice_Address&request("Invoice_Lane")&"巷"
    If request("Invoice_Alley")<>"" Then Invoice_Address=Invoice_Address&request("Invoice_Alley")&request("Invoice_AlleyType")
    If request("Invoice_HouseNo")<>"" Then
      If request("Invoice_HouseNoSub")<>"" Then
        Invoice_Address=Invoice_Address&request("Invoice_HouseNo")&"之"&request("Invoice_HouseNoSub")&"號"
      Else
        Invoice_Address=Invoice_Address&request("Invoice_HouseNo")&"號"
      End If
    End If
    If request("Invoice_Floor")<>"" Then
      If request("Invoice_FloorSub")<>"" Then
        Invoice_Address=Invoice_Address&request("Invoice_Floor")&"樓之"&request("Invoice_FloorSub")
      Else
        Invoice_Address=Invoice_Address&request("Invoice_Floor")&"樓"
      End If
    End If
    If request("Invoice_Room")<>"" Then Invoice_Address=Invoice_Address&" "&request("Invoice_Room")&"室"
  Else
    Invoice_Address=Data_Plus(request("Invoice_OverseasAddress"))&" "&Data_Plus(request("Invoice_OverseasCountry"))
  End If

  SQL1="Select * From DONOR Where Donor_id='"&request("donor_id")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  If request("Introducer_Id")<>"" Then
    RS1("Introducer_Id")=request("Introducer_Id")
    RS1("Introducer_Name")=Data_Plus(request("Introducer_Name"))
  Else
    RS1("Introducer_Name")=""
  End If    
  RS1("Donor_Name")=Data_Plus(request("Donor_Name"))
  RS1("Sex")=request("Sex")
  RS1("Title")=request("Title")
  RS1("Category")=request("Category")
  RS1("Donor_Type")=request("Donor_Type")
  RS1("IDNo")=request("IDNo")
  If request("Birthday")<>"" Then
    RS1("Birthday")=request("Birthday")
  Else
    RS1("Birthday")=null
  End If
  RS1("Education")=request("Education")
  RS1("Occupation")=request("Occupation")
  RS1("Marriage")=request("Marriage")
  RS1("Religion")=request("Religion")
  RS1("ReligionName")=request("ReligionName")
  RS1("Cellular_Phone")=request("Cellular_Phone")
  RS1("Tel_Office")=request("Tel_Office")
  RS1("Tel_Office_Loc")=request("Tel_Office_Loc")
  RS1("Tel_Office_Ext")=request("Tel_Office_Ext")
  RS1("Tel_Home")=request("Tel_Home")
  RS1("Fax")=request("Fax")
  RS1("Fax_Loc")=request("Fax_Loc")
  RS1("Email")=request("Email")
  RS1("Contactor")=Data_Plus(request("Contactor"))
  RS1("OrgName")=request("OrgName")
  RS1("JobTitle")=request("JobTitle")
  RS1("IsAbroad")=request("IsAbroad")
  '20130724 Modify by GoodTV Tanya:修改文宣品項目
  '20130912 Modify by GoodTV Tanya:修改紙本月刊格式
  If request("IsSendNewsNum")<>"" Then
  	If request("IsSendNewsNum") > 0 Then
    	RS1("IsSendNews")="Y"
     	RS1("IsSendNewsNum")=request("IsSendNewsNum")
    Else
     	RS1("IsSendNews")="N"
     	RS1("IsSendNewsNum")="0"
    End If
  End If
  If request("IsDVD")<>"" Then
    RS1("IsDVD")="Y"
  Else
    RS1("IsDVD")="N"
  End If
  If request("IsSendEpaper")<>"" Then
    RS1("IsSendEpaper")="Y"
  Else
    RS1("IsSendEpaper")="N"
  End If
'  If request("IsSendYNews")<>"" Then
'    RS1("IsSendYNews")="Y"
'  Else
'    RS1("IsSendYNews")="N"
'  End If
'  If request("IsBirthday")<>"" Then
'    RS1("IsBirthday")="Y"
'  Else
'    RS1("IsBirthday")="N"
'  End If  
'  If request("IsXmas")<>"" Then
'    RS1("IsXmas")="Y"
'  Else
'    RS1("IsXmas")="N"
'  End If 
  '20130725 Modify by GoodTV Tanya:增加「不主動聯絡-錯址」
  '20130923 Modify by GoodTV Tanya:拆開「不主動聯絡」及「錯址」
  If request("IsErrAddress")<>"" Then
    RS1("IsErrAddress")="Y"
  Else
    RS1("IsErrAddress")="N"
  End If
  If request("IsContact")<>"" Then
    RS1("IsContact")="N"
  Else
    RS1("IsContact")="Y"
  End If
  RS1("Invoice_Type")=request("Invoice_Type")
'  If request("IsAnonymous")<>"" Then
'    RS1("IsAnonymous")="Y"
'  Else
'    RS1("IsAnonymous")="N"
'  End If
  RS1("NickName")=Data_Plus(request("NickName"))
  RS1("Title2")=request("Title2")
  RS1("Invoice_Title")=Data_Plus(request("Invoice_Title"))
  RS1("Invoice_IDNo")=request("Invoice_IDNo")
  RS1("IsAbroad_Invoice")=request("IsAbroad_Invoice")
  If request("IsFdc")<>"" Then
    RS1("IsFdc")=request("IsFdc")
  Else
    RS1("IsFdc")="N"
  End If  
  RS1("Remark")=request("Remark")
  RS1("LastUpdate_Dept_Id")=Session("dept_id")
  RS1("LastUpdate_Date")=Date()
  RS1("LastUpdate_DateTime")=Now()
  RS1("LastUpdate_User")=session("user_name")
  RS1("LastUpdate_IP")=Request.ServerVariables("REMOTE_HOST")   

  If request("IsAbroad")="N" Then  
    RS1("ZipCode")=request("ZipCode")
    RS1("City")=request("City")
    RS1("Area")=request("Area")
    RS1("Street")=request("Street")
    RS1("Section")=request("SectionType")
    RS1("Lane")=request("Lane")
    RS1("Alley")=request("Alley")
    RS1("HouseNo")=request("HouseNo")
    RS1("HouseNoSub")=request("HouseNoSub")
    RS1("Floor")=request("Floor")
    RS1("FloorSub")=request("FloorSub")
    RS1("Room")=request("Room")
    RS1("OverseasCountry")=""
    RS1("OverseasAddress")=""
    RS1("Address")=Address
  Else
    RS1("ZipCode")=""
    RS1("City")=""
    RS1("Area")=""
    RS1("Street")=""
    RS1("Section")=""
    RS1("Lane")=""
    RS1("Alley")=""
    RS1("HouseNo")=""
    RS1("HouseNoSub")=""
    RS1("Floor")=""
    RS1("FloorSub")=""
    RS1("Room")=""
    RS1("OverseasCountry")=request("OverseasCountry")
    RS1("OverseasAddress")=request("OverseasAddress")
    RS1("Address")=Address
  End If
  If request("IsAbroad_Invoice")="N" Then  
    RS1("Invoice_ZipCode")=request("Invoice_ZipCode")
    RS1("Invoice_City")=request("Invoice_City")
    RS1("Invoice_Area")=request("Invoice_Area")
    RS1("Invoice_Street")=request("Invoice_Street")
    RS1("Invoice_Section")=request("Invoice_Section")
    RS1("Invoice_Lane")=request("Invoice_Lane")
    RS1("Invoice_Alley")=request("Invoice_Alley")
    RS1("Invoice_HouseNo")=request("Invoice_HouseNo")
    RS1("Invoice_HouseNoSub")=request("Invoice_HouseNoSub")
    RS1("Invoice_Floor")=request("Invoice_Floor")
    RS1("Invoice_FloorSub")=request("Invoice_FloorSub")
    RS1("Invoice_Room")=request("Invoice_Room")
    RS1("Invoice_OverseasCountry")=""
    RS1("Invoice_OverseasAddress")=""
    RS1("Invoice_Address")=Invoice_Address
  Else
    RS1("Invoice_ZipCode")=""
    RS1("Invoice_City")=""
    RS1("Invoice_Area")=""
    RS1("Invoice_Street")=""
    RS1("Invoice_Section")=""
    RS1("Invoice_Lane")=""
    RS1("Invoice_Alley")=""
    RS1("Invoice_HouseNo")=""
    RS1("Invoice_HouseNoSub")=""
    RS1("Invoice_Floor")=""
    RS1("Invoice_FloorSub")=""
    RS1("Invoice_Room")=""
    RS1("Invoice_OverseasCountry")=request("Invoice_OverseasCountry")
    RS1("Invoice_OverseasAddress")=request("Invoice_OverseasAddress")
    RS1("Invoice_Address")=Invoice_Address
  End If   
  '20130724 Mark by GoodTV Tanya:隱藏「徵信錄原則」、「群組」及「收據抬頭人數」
  '20130917 Modify by GoodTV Tanya:開啟「徵信錄原則」
  RS1("Report_Type")=Request("Report_Type")
  If Request("Report_Type")="不刊登" Then RS1("IsAnonymous")="Y" Else RS1("IsAnonymous")="N" End If
'  RS1("Report_Group")=Request("Report_Group")
'  if Request("Invoice_Title_Man")="" then 
'  	RS1("Invoice_Title_Man")=null
'  Else
'  	RS1("Invoice_Title_Man")=Request("Invoice_Title_Man")
'  End If
'  If Request("Report_Type")="刊登其他名稱" Then RS1("IsAnonymous")="Y"
  RS1("ToGOODTV")=request("ToGOODTV") 
  RS1("Attn")=request("Attn") 
  RS1("Invoice_Attn")=request("Invoice_Attn")
  RS1.Update
  RS1.CLose
  Set RS1=Nothing
  
  SQL="Update DONOR Set Introducer_Name='"&Data_Plus(request("Donor_Name"))&"' Where Introducer_Id='"&request("donor_id")&"'"
  Set RS=Conn.Execute(SQL)

  '修改捐款人捐款紀錄
  call Declare_DonorId (request("donor_id"))

  session("errnumber")=1
  session("msg")="捐款人資料修改成功 ！"
  'session("msg")="捐款人資料已修改\n您可以新增捐款資料 ！"
  'Response.Redirect "donate_add.asp?ctype=donate_data&donor_id="&request("donor_id")
End If

If request("action")="delete" Then
  SQL1="Select * From DONATE Where Donor_id='"&request("donor_id")&"'"
  Call QuerySQL(SQL1,RS1)
  If Not RS1.EOF Then
    session("errnumber")=1
    session("msg")="此捐款人尚有捐款記錄存在，請先刪除捐款明細，才能刪除捐款人！"
  Else
    SQL2="Select * From DONOR Where Donor_id='"&request("donor_id")&"'"
    Call QuerySQL(SQL2,RS2)  
    If Not RS2.EOF Then
      log_type="刪除捐款人"
      log_desc=CStr(RS2("Donor_id"))&RS2("Donor_Name")
      call syslog(log_type,log_desc)
      SQL="Delete From DONOR Where Donor_id='"&request("donor_id")&"'"
      Call ExecSQL(SQL)
      session("errnumber")=1
      session("msg")="捐款人資料刪除成功 ！"
      Response.Redirect "donor.asp"
    End If
    RS2.Close
    Set RS2=Nothing
  End If
  RS1.Close
  Set RS1=Nothing
End If

SQL_IS="Select * From DEPT Where Dept_Id='"&Session("dept_id")&"' "
Set RS_IS=Server.CreateObject("ADODB.Recordset")
RS_IS.Open SQL_IS,Conn,1,1
IsContribute=RS_IS("IsContribute")
IsMember=RS_IS("IsMember")
IsMemberPre=RS_IS("IsMemberPre")
IsShopping=RS_IS("IsShopping")
Donate_IsFdc=RS_IS("Donate_IsFdc")
RS_IS.Close
Set RS_IS=Nothing

SQL="Select * From DONOR Where Donor_id='"&request("donor_id")&"'"
Call QuerySQL(SQL,RS)
%>
<%Prog_Id="donor"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <%If request("ctype")<>"" Then%>
      <input type="hidden" name="ctype" value="<%=request("ctype")%>">
      <%Else%>
      <input type="hidden" name="ctype" value="donor_edit">
      <%End If%>
      <input type="hidden" name="donor_id" value="<%=request("donor_id")%>">
      <input type="hidden" name="Introducer_Id" value="<%=RS("Introducer_Id")%>">
      <input type="hidden" name="IsMemberPre" value="<%=IsMemberPre%>">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;"><%=Prog_Desc%>【修改】</td>
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
                      <td width="99%">
                        <table border="1" cellpadding="2" style="border-collapse: collapse" width="100%" height="25" cellspacing="1">
                          <tr>
                            <td class="button2-bg"><img border="0" src="../images/red_arrow.gif" align="texttop">捐款人資料</td>
                            <td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="donate_data.asp?donor_id=<%=request("donor_id")%>">捐款記錄</a></td>
                            <td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="pledge_data.asp?donor_id=<%=request("donor_id")%>">轉帳授權書記錄</a></td>
                            <%If IsContribute="Y" Then%><td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="../contribute/contribute_data.asp?donor_id=<%=request("donor_id")%>">公關贈品記錄</a></td><%End If%>
                            <%If IsMember="Y" Then%><td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="../member/signup_data.asp?donor_id=<%=request("donor_id")%>">活動報名記錄</a></td><%End If%>
                            <%If IsShopping="Y" Then%><td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="../shopping/shopping_data.asp?donor_id=<%=request("donor_id")%>">商品義賣記錄</a></td><%End If%>
                          </tr>
                        </table>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" width="100%">
                        <table border="0" cellpadding="1" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
                            <td align="center" colspan="8">
                               <input type="button" value=" 修 改 " name="save" class="cbutton" style="cursor:hand" onClick="Update_OnClick()">
                               &nbsp;&nbsp;
                               <input type="button" value=" 刪 除 " name="save" class="cbutton" style="cursor:hand" onClick="Delete_OnClick()">	
                               &nbsp;&nbsp;
                               <input type="button" value=" 取 消 " name="cancel" class="cbutton" style="cursor:hand" onClick="Cancel_OnClick()">
                               &nbsp;&nbsp;
                               <input type="button" value="新增本人捐款資料" name="Add" class="addbutton" style="cursor:hand" onClick="Add_OnClick()">
                               <%If IsMember="Y" And RS("IsMember")="N" Then%>
                               &nbsp;&nbsp;
                               <input type="button" value="將捐款人轉為本會會員" name="Add" class="delbutton" style="cursor:hand" onClick="IsMember_OnClick()">	
                               <%End If%>
                            </td>
                          </tr>
						   <tr>
                            <td width="99%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                          </tr>
						  <tr>
                          	<td align="right">捐款人編號：</td>
                            <td align="left" colspan="7">&nbsp;<%=RS("Donor_ID")%></td>
                          </tr>
                          <tr>
                          	<!--20131106 Modify by GoodTV Tanya:增加「必輸欄位」註記和判斷-->
                          	<!--20140114 Modify by GoodTV Tanya:增加「字數」限制提醒-->
                          	<!--20140210 Modify by GoodTV Tanya:移除「字數」限制提醒-->
                            <td align="right"><font color="red">*</font>捐款人：</td>
                            <td align="left" colspan="3"><input type="text" name="Donor_Name" size="46" class="font9" maxlength="100" value="<%=Data_Minus(RS("Donor_Name"))%>"></td>
                            <td align="left" colspan="4">
                              <!--類別：
                              <%
                                'SQL="Select Category=CodeDesc From CASECODE Where codetype='Category' Order By Seq"
                                'FName="Category"
                                'Listfield="Category"
                                'menusize="1"
                                'BoundColumn=RS("category")
                                'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>-->
                              &nbsp;性別：
                              <%
                                SQL="Select Sex=CodeDesc From CASECODE Where codetype='Sex' Order By Seq"
                                FName="Sex"
                                Listfield="Sex"
                                menusize="1"
                                BoundColumn=RS("Sex")
                                '20130725 Modify by GoodTV Tanya:增加性別與稱謂連動
                                call CodeSex (SQL,FName,Listfield,BoundColumn,menusize) 
                                'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                              &nbsp;稱謂：
                              <%
                                SQL="Select Title=CodeDesc From CASECODE Where codetype='Title' Order By Seq"
                                FName="Title"
                                Listfield="Title"
                                menusize="1"
                                BoundColumn=RS("Title")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                            </td>
                          </tr>
                          <tr>
                            <td align="right"><font color="red">*</font>身份別：</td>
                            <td align="left" colspan="7">
                              <%
                                SQL="Select Donor_Type=CodeDesc From CASECODE Where codetype='DonorType' Order By Seq"
                                FName="Donor_Type"
                                Listfield="Donor_Type"
                                BoundColumn=RS("Donor_Type")
                                call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                              %>
                            </td>
                          </tr>
                          <tr>
                            <td align="right" width="12%">身分證/統編：</td>
                            <td align="left" width="15%"><input type="text" name="IDNo" size="15" class="font9" maxlength="10" onKeyUp="javascript:UCaseIDNO();" onblur="<%call CheckIDJ("IDNo")%>" value="<%=RS("IDNo")%>"></td>
                            <td align="right" width="10%">出生日期：</td>
                            <td align="left" width="15%"><%call Calendar("Birthday",RS("Birthday"))%></td>
                            <input type="hidden" name="Education" size="15" class="font9" maxlength="10" value="<%=RS("Education")%>">
                            <input type="hidden" name="Occupation" size="15" class="font9" maxlength="10" value="<%=RS("Occupation")%>">
                            <input type="hidden" name="Marriage" size="15" class="font9" maxlength="10" value="<%=RS("Marriage")%>">
                            <input type="hidden" name="Religion" size="15" class="font9" maxlength="10" value="<%=RS("Religion")%>">
                            <input type="hidden" name="ReligionName" size="15" class="font9" maxlength="10" value="<%=RS("ReligionName")%>">
                            <td align="right">手機：</td>
                            <td align="left"><input type="text" name="Cellular_Phone" size="15" class="font9" maxlength="40" value="<%=RS("Cellular_Phone")%>"></td>
                            <!--<td align="right" width="10%">教育程度：</td>
                            <td align="left" width="15%">
                            <%
                              'SQL="Select Education=CodeDesc From CASECODE Where codetype='Education' Order By Seq"
                              'FName="Education"
                              'Listfield="Education"
                              'menusize="1"
                              'BoundColumn=RS("Education")
                              'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                            <td align="right" width="8%">職業別：</td>
                            <td align="left" width="15%">
                            <%
                              'SQL="Select Occupation=CodeDesc From CASECODE Where codetype='Occupation' Order By Seq"
                              'FName="Occupation"
                              'Listfield="Occupation"
                              'menusize="1"
                              'BoundColumn=RS("Occupation")
                              'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>-->
                          </tr>
                          <!--<tr>
                            <td align="right">婚姻狀況：</td>
                            <td align="left">
                            <%
                              'SQL="Select Marriage=CodeDesc From CASECODE Where codetype='Marriage' Order By Seq"
                              'FName="Marriage"
                              'Listfield="Marriage"
                              'menusize="1"
                              'BoundColumn=RS("Marriage")
                              'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                            <td align="right">宗教信仰：</td>
                            <td align="left">
                            <%
                              'SQL="Select Religion=CodeDesc From CASECODE Where codetype='Religion' Order By Seq"
                              'FName="Religion"
                              'Listfield="Religion"
                              'menusize="1"
                              'BoundColumn=RS("Religion")
                              'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                            <td colspan="4">
                            	&nbsp;&nbsp;&nbsp;&nbsp;道場/教會名稱：
                              <input type="text" name="ReligionName" size="39" class="font9" maxlength="40" value="<%'=Data_Minus(RS("ReligionName"))%>">
                            </td>
                          </tr>-->                            
                          <tr>                            
                            <td align="right">電話：</td>
                            <td align="left" colspan="3"><input type="text" name="Tel_Office_Loc" size="5" class="font9" maxlength="5" value="<%=RS("Tel_Office_Loc")%>">-<input type="text" name="Tel_Office" size="15" class="font9" maxlength="40" value="<%=RS("Tel_Office")%>">-<input type="text" name="Tel_Office_Ext" size="5" class="font9" maxlength="5" value="<%=RS("Tel_Office_Ext")%>"></td>
                            <input type="hidden" name="Tel_Home" size="15" class="font9" maxlength="10" value="<%=RS("Tel_Home")%>">
                            <!--<td align="right">電話(夜)：</td>
                            <td align="left"><input type="text" name="Tel_Home" size="15" class="font9" maxlength="40" value="<%'=RS("Tel_Home")%>"></td>-->
                            <td align="right">傳真：</td>
                            <td align="left"><input type="text" name="Fax_Loc" size="5" class="font9" maxlength="5" value="<%=RS("Fax_Loc")%>">-<input type="text" name="Fax" size="15" class="font9" maxlength="40" value="<%=RS("Fax")%>"></td>
                          </tr>
                          <tr>
                            <td align="right">E-Mail：</td>
                            <td align="left"><input type="text" name="Email" size="25" class="font9" maxlength="80" value="<%=RS("Email")%>"></td>
                            <td align="right">聯絡人：</td>
                            <td align="left"><input type="text" name="Contactor" size="15" class="font9" maxlength="40" value="<%=Data_Minus(RS("Contactor"))%>"></td>
                            <input type="hidden" name="OrgName" size="15" class="font9" maxlength="10" value="<%=RS("OrgName")%>">
                            <!--<td align="right">服務單位：</td>
                            <td align="left"><input type="text" name="OrgName" size="15" class="font9" maxlength="40" value="<%'=RS("OrgName")%>"></td>-->
                            <!--20130724 Modify by GoodTV Tanya:修改欄位名稱為「聯絡人職稱」-->
                            <!--td align="right">聯絡人職稱：</td-->
                            <!--td align="left"><input type="text" name="JobTitle" size="15" class="font9" maxlength="40" value="<%=RS("JobTitle")%>"></td-->
                          </tr>
                          <tr>
                            <td align="right"><font color="red">*</font>通訊地址：</td>
                            <td align="left" colspan="7">
                              <input type="radio" name="IsAbroad" id="IsAbroadY" value="N" <%If RS("IsAbroad")="N" Or RS("IsAbroad")="" Then Response.Write "checked" End If%> >台灣本島
                              <%call CodeCity2 ("form","ZipCode",RS("ZipCode"),"City",RS("City"),"Area",RS("Area"),"Y")%>
                            	<input type='text' class='font9' name='Street' size='25' maxlength='40' value='<%=RS("Street")%>'>大道/路/街/部落
                              <br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                              <%
                                SQL="Select SectionType=CodeDesc From CASECODE Where codetype='SectionType' Order By Seq"
                                FName="SectionType"
                                Listfield="SectionType"
                                menusize="1"
                                BoundColumn=RS("Section")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>段
                            	<input type='text' class='font9' name='Lane' size='5' maxlength='10' value='<%=RS("Lane")%>'>巷
                            	<input type='text' class='font9' name='Alley' size='5' maxlength='10' value='<%=RS("Alley")%>'>
                              <%
                                SQL="Select AlleyType=CodeDesc From CASECODE Where codetype='AlleyType' Order By Seq"
                                FName="AlleyType"
                                Listfield="AlleyType"
                                menusize="1"
                                If RS("Address")<>"" Then
                                  If Instr(RS("Address"),"弄")>0 Then
                                    BoundColumn="弄"
                                  ElseIf Instr(RS("Address"),"衖")>0 Then
                                    BoundColumn="衖"
                                  Else
                                    BoundColumn=""
                                  End If
                                Else
                                  BoundColumn=""
                                End If
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                            	<input type='text' class='font9' name='HouseNo' size='5' maxlength='10' value='<%=RS("HouseNo")%>'>之<input type='text' class='font9' name='HouseNoSub' size='5' maxlength='10' value='<%=RS("HouseNoSub")%>'>號
                            	<input type='text' class='font9' name='Floor' size='5' maxlength='10' value='<%=RS("Floor")%>'>樓之<input type='text' class='font9' name='FloorSub' size='5' maxlength='10' value='<%=RS("FloorSub")%>'>
                            	<input type='text' class='font9' name='Room' size='5' maxlength='10' value='<%=RS("Room")%>'>室
                                Attn：<input type='text' class='font9' name='Attn' size='20' maxlength='50' value='<%=RS("Attn")%>'>
                              <br />
                              <input type="radio" name="IsAbroad" id="IsAbroadN" value="Y" <%If RS("IsAbroad")="Y" Then Response.Write "checked" End If%> >海外地址
                              <input type='text' class='font9' name='OverseasAddress' size='50' maxlength='100' value='<%=RS("OverseasAddress")%>'>
                              
                              國家/省城市/區
							  <input type='text' class='font9' name='OverseasCountry' size='20' maxlength='50' value='<%=RS("OverseasCountry")%>'>
                              
                            </td>
                          </tr>
                          <tr>
                            <td align="right">文宣品：</td>
                            <td align="left" colspan="3">
                            	<!--20130724 Modify by GoodTV Tanya:修改文宣品選項及增加「不主動聯絡-錯址」-->
                            	   紙本月刊本數<input type="text" class="font9" name="IsSendNewsNum" size='3' maxlength='3' value='<%=RS("IsSendNewsNum")%>'>
                            	<input type="checkbox" name="IsDVD" value="Y" <%If RS("IsDVD")="Y" Then Response.Write "checked" End If%> >DVD
                              <input type="checkbox" name="IsSendEpaper" value="Y" <%If RS("IsSendEpaper")="Y" Then Response.Write "checked" End If%> >電子文宣
                            	<!--<input type="checkbox" name="IsSendNews" value="Y" <%If RS("IsSendNews")="Y" Then Response.Write "checked" End If%> >會訊
                              <input type="checkbox" name="IsSendEpaper" value="Y" <%If RS("IsSendEpaper")="Y" Then Response.Write "checked" End If%> >電子報
                              <input type="checkbox" name="IsSendYNews" value="Y" <%If RS("IsSendYNews")="Y" Then Response.Write "checked" End If%> >年報
                              <input type="checkbox" name="IsBirthday" value="Y" <%If RS("IsBirthday")="Y" Then Response.Write "checked" End If%> >生日卡
                              <input type="checkbox" name="IsXmas" value="Y" <%If RS("IsXmas")="Y" Then Response.Write "checked" End If%> >賀卡-->
                            </td>
                            <!--20130725 Modify by GoodTV Tanya:增加「不主動聯絡-錯址」-->
                            <!--20130923 Modify by GoodTV Tanya:拆開「不主動聯絡」及「錯址」-->
                            <td align="left" colspan="1">
                            	<input type="checkbox" name="IsContact" value="N" <%If RS("IsContact")="N" Then Response.Write "checked" End If%>>不主動聯絡
                            </td>
                            <td align="left" colspan="3">
                            	<input type="checkbox" name="IsErrAddress" value="Y" <%If RS("IsErrAddress")="Y" Then Response.Write "checked" End If%>>錯址
                            </td>
                          </tr>
                          <!--#include file="../include/calendar2.asp"-->
                          <tr>
                            <td width="99%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                          </tr>
                          <tr>
                            <td align="right"><font color="red">*</font>收據開立：</td>
                            <td align="left" colspan="7">
                              <%
                                SQL="Select Invoice_Type=CodeDesc From CASECODE Where codetype='InvoiceType' Order By Seq"
                                FName="Invoice_Type"
                                Listfield="Invoice_Type"
                                menusize="1"
                                BoundColumn=RS("Invoice_Type")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                              <input type="checkbox" name="IsSameAddress" value="Y" OnClick="SameAddress_OnClick()"><font color="#ff0000">收據資料同上</font>	
                              <!--20130917 Mark by Tanya -->
                              <!--<input type="checkbox" name="IsAnonymous" value="Y" <%If RS("IsAnonymous")="Y" Then Response.Write "checked" End If%> >徵信錄匿名-->
                            	<!--&nbsp;匿名：
                            	<input type="text" name="NickName" size="13" class="font9" maxlength="20" value="<%=RS("NickName")%>">-->
                                <input type="hidden" name="NickName" size="13" class="font9" maxlength="20" value="<%=RS("NickName")%>">
                            	&nbsp;收據稱謂：
                              <%
                                SQL="Select Title2=CodeDesc From CASECODE Where codetype='Title' Order By Seq"
                                FName="Title2"
                                Listfield="Title2"
                                menusize="1"
                                BoundColumn=RS("Title2")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                            </td>	
                          </tr>
                          <tr>
                            <td align="right">收據地址：</td>
                            <td align="left" colspan="7">
                              <input type="radio" name="IsAbroad_Invoice" id="IsAbroad_InvoiceY" value="N" <%If RS("IsAbroad_Invoice")="N" Or RS("IsAbroad_Invoice")="" Then Response.Write "checked" End If%> >台灣本島
                            	<%call CodeCity2 ("form","Invoice_ZipCode",RS("Invoice_ZipCode"),"Invoice_City",RS("Invoice_City"),"Invoice_Area",RS("Invoice_Area"),"N")%>
                            	<!--20130801 Modify by GoodTV Tanya:修改欄位長度同通訊地址-->
                            	<input type='text' class='font9' name='Invoice_Street' size='25' maxlength='40' value='<%=RS("Invoice_Street")%>'>大道/路/街/部落
                              <br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                              <%
                                SQL="Select Invoice_Section=CodeDesc From CASECODE Where codetype='SectionType' Order By Seq"
                                FName="Invoice_Section"
                                Listfield="Invoice_Section"
                                menusize="1"
                                BoundColumn=RS("Invoice_Section")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>段
                            	<input type='text' class='font9' name='Invoice_Lane' size='5' maxlength='10' value='<%=RS("Invoice_Lane")%>'>巷
                            	<input type='text' class='font9' name='Invoice_Alley' size='5' maxlength='10' value='<%=RS("Invoice_Alley")%>'>
                              <%
                                SQL="Select Invoice_AlleyType=CodeDesc From CASECODE Where codetype='AlleyType' Order By Seq"
                                FName="Invoice_AlleyType"
                                Listfield="Invoice_AlleyType"
                                menusize="1"
                                If RS("Invoice_Address")<>"" Then
                                  If Instr(RS("Invoice_Address"),"弄")>0 Then
                                    BoundColumn="弄"
                                  ElseIf Instr(RS("Invoice_Address"),"衖")>0 Then
                                    BoundColumn="衖"
                                  Else
                                    BoundColumn=""
                                  End If
                                Else
                                  BoundColumn=""
                                End If
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                            	<input type='text' class='font9' name='Invoice_HouseNo' size='5' maxlength='10' value='<%=RS("Invoice_HouseNo")%>'>之<input type='text' class='font9' name='Invoice_HouseNoSub' size='5' maxlength='10' value='<%=RS("Invoice_HouseNoSub")%>'>號
                            	<input type='text' class='font9' name='Invoice_Floor' size='5' maxlength='10' value='<%=RS("Invoice_Floor")%>'>樓之<input type='text' class='font9' name='Invoice_FloorSub' size='5' maxlength='10' value='<%=RS("Invoice_FloorSub")%>'>
                            	<input type='text' class='font9' name='Invoice_Room' size='5' maxlength='10' value='<%=RS("Invoice_Room")%>'>室
                            		Attn：<input type='text' class='font9' name='Invoice_Attn' size='20' maxlength='50' value='<%=RS("Invoice_Attn")%>'>
                              <br />
                              <input type="radio" name="IsAbroad_Invoice" id="IsAbroad_InvoiceN" value="Y" <%If RS("IsAbroad_Invoice")="Y" Then Response.Write "checked" End If%> >海外地址
                              <input type='text' class='font9' name='Invoice_OverseasAddress' size='50' maxlength='100' value='<%=RS("Invoice_OverseasAddress")%>'>
                              
                              國家/省城市/區
							  <input type='text' class='font9' name='Invoice_OverseasCountry' size='20' maxlength='50' value='<%=RS("Invoice_OverseasCountry")%>'>
                              
                            </td>
                          </tr>
                          <tr>
                            <td align="right">收據抬頭：</td>
                            <!--20140114 Modify by GoodTV Tanya:增加「字數」限制提醒-->
                            <td align="left" colspan="3"><input type="text" name="Invoice_Title" size="46" class="font9" maxlength="100" value="<%=Data_Minus(RS("Invoice_Title"))%>" onblur="javascript:CheckDonorNameLen('2',this.value);"></td>
                            <td align="left" colspan="4">
                              收據身分證/統編：
                              <input type="text" name="Invoice_IDNo" size="15" class="font9" maxlength="10" onKeyUp="javascript:UCaseInvoiceIDNo();" onblur="<%call CheckIDJ("Invoice_IDNo")%>" value="<%=RS("Invoice_IDNo")%>">
                            </td>
                          </tr>
                          <!--20130724 Modify by GoodTV Tanya:隱藏「徵信錄原則」、「群組」及「收據抬頭人數」 -->
                          <!--20130918 Modify by GoodTV Tanya:開啟「徵信錄原則」 --> 
                          <tr> 							
                            <td align="right"> 徵信錄原則：</td>
                            <td align="left" >
                            <%
                              SQL="Select Report_Type=CodeDesc From CASECODE Where codetype='ReportType' Order By Seq"
                              FName="Report_Type"
                              Listfield="Report_Type"
                              menusize="1"
                              BoundColumn=RS("Report_Type")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
						
                            <!--<td align="right">群組：</td>
                            <td align="left" >
                            <%
                              SQL="Select Report_Group=CodeDesc From CASECODE Where codetype='ReportGroup' Order By Seq"
                              FName="Report_Group"
                              Listfield="Report_Group"
                              menusize="1"
                              BoundColumn=RS("Report_Group")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
						
                            <td align="right">收據抬頭人數：</td>
                            <td align="left" ><input type="text" name="Invoice_Title_Man" size="2" class="font9" maxlength="2" value="<%=RS("Invoice_Title_Man")%>"></td>-->
                          </tr>
                          <%If Donate_IsFdc="Y" Then%>
                          <tr>
                            <td align="left" colspan="8">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="IsFdc" value="Y" <%If RS("IsFdc")="Y" Then Response.Write "checked" End If%> ><font color="#ff0000">本捐款人願意將捐款資料上傳至國稅局以供報稅用</font>&nbsp;(&nbsp;請填寫捐款人收據身分證/統編&nbsp;)</td>
                          </tr>
                          <%Else%>
                          <input type="hidden" name="IsFdc" value="<%=RS("IsFdc")%>">
                          <%End If%>
                          <tr>
                            <td align="right">捐款人備註：</td>
                            <td align="left" colspan="7">
                            	<textarea name="Remark" rows="5" cols="90" class="font9"><%=RS("Remark")%></textarea>
                            </td>
                          </tr>
						  						<tr>
                            <td align="right">想對GOOD TV說的話：</td>
                            <td align="left" colspan="7">
                            	<textarea name="ToGOODTV" rows="3" cols="90" class="font9"><%=RS("ToGOODTV")%></textarea>
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                          </tr>
                          <tr>  
                            <td align="left" colspan="8">
                              <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                                <tr>
                                	<!--20140106 Modify by GoodTV Tanya:首次捐款日期改抓Begin_DonateDate,最近捐款日期改抓Last_DonateDate,累計捐款次數改抓Donate_No,累計捐款金額改抓Donate_Total-->
                          	      <td align="right">首次捐款日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDate" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDate")%>"></td>
                                  <td align="right">最近捐款日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDate" size="13" class="font9t" readonly value="<%=RS("Last_DonateDate")%>"></td>
                                  <td align="right">累計捐款次數：</td>
                                  <td align="left"><input type="text" name="Donate_No" size="13" class="font9t" readonly value="<%if RS("Donate_No")<>"" And isnull(RS("Donate_No"))=false Then Response.write FormatNumber(RS("Donate_No"),0)%>" style="text-align: right"></td>
                            	    <td align="right">累計捐款金額：</td>
                            	    <td align="left"><input type="text" name="Donate_Total" size="13" class="font9t" readonly value="<%if RS("Donate_Total")<>"" And isnull(RS("Donate_Total"))=false Then Response.write FormatNumber(RS("Donate_Total"),0)%>" style="text-align: right"></td>
                                </tr>
                                <%If IsContribute="Y" Then%>
                                <tr>
                          	      <td align="right">首次捐物日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateC" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateC")%>"></td>
                                  <td align="right">最近捐物日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateC" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateC")%>"></td>
                                  <td align="right">累計捐物次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoC" size="13" class="font9t" readonly value="<%if RS("Donate_NoC")<>"" And isnull(RS("Donate_NoC"))=false Then Response.write FormatNumber(RS("Donate_NoC"),0)%>" style="text-align: right"></td>
                            	    <td align="right">累計折合現金：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalC" size="13" class="font9t" readonly value="<%if RS("Donate_TotalC")<>"" And isnull(RS("Donate_TotalC"))=false Then Response.write FormatNumber(RS("Donate_TotalC"),0)%>" style="text-align: right"></td>
                                </tr>
                                <%End If%>
                                <%If IsMember="Y" Then%>
                                <tr>
                          	      <td align="right">首次繳費日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateM" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateM")%>"></td>
                                  <td align="right">最近繳費日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateM" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateM")%>"></td>
                                  <td align="right">累計繳費次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoM" size="13" class="font9t" readonly value="<%if RS("Donate_NoM")<>"" And isnull(RS("Donate_NoM"))=false Then Response.write FormatNumber(RS("Donate_NoM"),0)%>" style="text-align: right"></td>
                            	    <td align="right">累計繳費金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalM" size="13" class="font9t" readonly value="<%if RS("Donate_TotalM")<>"" And isnull(RS("Donate_TotalM"))=false Then Response.write FormatNumber(RS("Donate_TotalM"),0)%>" style="text-align: right"></td>
                                </tr>
                                <tr>
                          	      <td align="right">首次報名日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateA" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateA")%>"></td>
                                  <td align="right">最近報名日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateA" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateA")%>"></td>
                                  <td align="right">累計報名次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoA" size="13" class="font9t" readonly value="<%if RS("Donate_NoA")<>"" And isnull(RS("Donate_NoA"))=false Then Response.write FormatNumber(RS("Donate_NoA"),0)%>" style="text-align: right"></td>
                            	    <td align="right">累計報名金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalA" size="13" class="font9t" readonly value="<%if RS("Donate_TotalA")<>"" And isnull(RS("Donate_TotalA"))=false Then Response.write FormatNumber(RS("Donate_TotalA"),0)%>" style="text-align: right"></td>
                                </tr>
                                <%End If%>
                                <%If IsShopping="Y" Then%>
                                <tr>
                          	      <td align="right">首次購物日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateS" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateS")%>"></td>
                                  <td align="right">最近購物日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateS" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateS")%>"></td>
                                  <td align="right">累計購物次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoS" size="13" class="font9t" readonly value="<%if RS("Donate_NoS")<>"" And isnull(RS("Donate_NoS"))=false Then Response.write FormatNumber(RS("Donate_NoS"),0)%>" style="text-align: right"></td>
                            	    <td align="right">累計購物金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalS" size="13" class="font9t" readonly value="<%if RS("Donate_TotalS")<>"" And isnull(RS("Donate_TotalS"))=false Then Response.write FormatNumber(RS("Donate_TotalS"),0)%>" style="text-align: right"></td>
                                </tr>
                                <%End If%>
                                <%If IsMember="Y" Or IsShopping="Y" Then%>
                                <tr>
                          	      <td align="right">首次非捐日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateND" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateND")%>"></td>
                                  <td align="right">最近非捐日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateND" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateND")%>"></td>
                                  <td align="right">累計非捐次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoND" size="13" class="font9t" readonly value="<%if RS("Donate_NoND")<>"" And isnull(RS("Donate_NoND"))=false Then Response.write FormatNumber(RS("Donate_NoND"),0)%>" style="text-align: right"></td>
                            	    <td align="right">累計非捐金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalND" size="13" class="font9t" readonly value="<%if RS("Donate_TotalND")<>"" And isnull(RS("Donate_TotalND"))=false Then Response.write FormatNumber(RS("Donate_TotalND"),0)%>" style="text-align: right"></td>
                                </tr>
                                <%End If%>
                                <tr>
                                  <td align="right">資料建檔日期：</td>
                                  <td align="left"><input type="text" name="Create_Date" size="13" class="font9t" readonly value="<%=RS("Create_Date")%>"></td>
                                  <td align="right">資料建檔人員：</td>
                                  <td align="left"><input type="text" name="Create_User" size="13" class="font9t" readonly value="<%=RS("Create_User")%>"></td>
                            	    <td align="right">最後異動日期：</td>
                                  <td align="left"><input type="text" name="LastUpdate_Date" size="13" class="font9t" readonly value="<%=RS("LastUpdate_Date")%>"></td>
                                  <td align="right">最後異動人員：</td>
                                  <td align="left"><input type="text" name="LastUpdate_User" size="13" class="font9t" readonly value="<%=RS("LastUpdate_User")%>"></td>
                                </tr>
                              </table>
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
<script language="JavaScript"><!--
function Window_OnLoad(){
  document.form.Donor_Name.focus();
}
function CheckDonorNameLen(Type,objValue){
	var regx = /^[u4E00-u9FA5]+$/;
	
	if (Type == '1'){		
		if(!regx.test(objValue)){
			if(objValue.length > 11)
				alert("「捐款人」中文超過11字，列印收據會被截斷!");
		}
		else{
			if(objValue.length > 22)
				alert("「捐款人」英數超過22字，列印收據會被截斷!");
		}
	}
	else if(Type == '2'){	
		if(!regx.test(objValue)){
			if(objValue.length > 20)
				alert("「收據抬頭」中文超過20字，列印收據會被截斷!");
		}
		else{
			if(objValue.length > 40)
				alert("「收據抬頭」英數超過40字，列印收據會被截斷!");
		}
	}	
}
function Donor_Category_OnChange(Category){
  if(Category.indexOf('學生')>-1||Category.indexOf('個人')>-1){
    donorsex.style.display='block';
    donortitle.style.display='block';
  }else{
    donorsex.style.display='none';
    donortitle.style.display='none';
    document.form.Sex.value='';
    document.form.Title.value='';
  }
}
function UCaseMemberNo(){
  document.form.Member_No.value=document.form.Member_No.value.toUpperCase();
}
function Member_Type_OnClick(){
  if(document.form.Member_Type.checked){
    document.form.Member_No.style.backgroundColor='#ffffff';
    document.form.Member_No.readOnly=false;
    document.form.Member_No.focus();
  }else{
    document.form.Member_No.style.backgroundColor='#ffffcc';
    document.form.Member_No.readOnly=true;
    document.form.Member_No.value=document.form.Member_No_Old.value;
  }
}
function Get_Member_OnClick(){
  if(confirm('您是否確定要執行會員編號取號？')){
    document.form.action.value='get';
    document.form.submit();
  }  
}
function UCaseIDNO(){
  document.form.IDNo.value=document.form.IDNo.value.toUpperCase();
}

function UCaseInvoiceIDNo(){
  document.form.Invoice_IDNo.value=document.form.Invoice_IDNo.value.toUpperCase();
}
function IsAbroad_OnClick(){
  if(document.form.IsAbroad.checked){
    donor_address.style.display='none';
    document.form.ZipCode.value='';
    document.form.City.value='';
    ClearOption(document.form.Area);
    document.form.Area.options[0] = new Option('鄉鎮市區','');
  }else{
    donor_address.style.display='block';
  }
}
function IsAbroad_Invoice_OnClick(){
  if(document.form.IsAbroad_Invoice.checked){
    donor_invoic_address.style.display='none';
    document.form.Invoice_ZipCode.value='';
    document.form.Invoice_City.value='';
    ClearOption(document.form.Invoice_Area);
    document.form.Invoice_Area.options[0] = new Option('鄉鎮市區','');
  }else{
    donor_invoic_address.style.display='block';
  }
}
function SameAddress_OnClick(){
  if(document.form.IsSameAddress.checked){
    document.form.Invoice_Title.value=document.form.Donor_Name.value;
    document.form.Invoice_IDNo.value=document.form.IDNo.value;
    document.form.Title2.value=document.form.Title.value;
    document.form.Invoice_ZipCode.value='';
    document.form.Invoice_City.value='';
    ClearOption(document.form.Invoice_Area);
    document.form.Invoice_Area.options[0] = new Option('鄉鎮市區','');
    if(document.form.IsAbroad[0].checked){
      document.form.IsAbroad_Invoice[0].checked=true;
      document.form.Invoice_City.value=document.form.City.value;
      ChgCity(document.form.Invoice_City.value,document.form.Invoice_Area,document.form.Invoice_ZipCode);  
      document.form.Invoice_ZipCode.value=document.form.ZipCode.value;  
      document.form.Invoice_Area.value=document.form.Area.value;
      document.form.Invoice_Street.value=document.form.Street.value;
      document.form.Invoice_Section.value=document.form.SectionType.value;
      document.form.Invoice_Lane.value=document.form.Lane.value;
      document.form.Invoice_Alley.value=document.form.Alley.value;
      document.form.Invoice_AlleyType.value=document.form.AlleyType.value;
      document.form.Invoice_HouseNo.value=document.form.HouseNo.value;
      document.form.Invoice_HouseNoSub.value=document.form.HouseNoSub.value;
      document.form.Invoice_Floor.value=document.form.Floor.value;
      document.form.Invoice_FloorSub.value=document.form.FloorSub.value;
      document.form.Invoice_Room.value=document.form.Room.value;
	  document.form.Invoice_Attn.value=document.form.Attn.value;'added by Samuel Lin 2014/01/20'
      document.form.Invoice_OverseasCountry.value='';
      document.form.Invoice_OverseasAddress.value='';
      document.form.Invoice_Title.value=document.form.Donor_Name.value;
      document.form.Invoice_IDNo.value=document.form.IDNo.value;
      document.form.Title2.value=document.form.Title.value;
    }else{
      document.form.IsAbroad_Invoice[1].checked=true;
      document.form.Invoice_OverseasCountry.value=document.form.OverseasCountry.value;
      document.form.Invoice_OverseasAddress.value=document.form.OverseasAddress.value;
      document.form.Invoice_City.value='';
      document.form.Invoice_Street.value='';
      document.form.Invoice_Section.value='';
      document.form.Invoice_Lane.value='';
      document.form.Invoice_Alley.value='';
      document.form.Invoice_AlleyType.value='弄';
      document.form.Invoice_HouseNo.value='';
      document.form.Invoice_HouseNoSub.value='';
      document.form.Invoice_Floor.value='';
      document.form.Invoice_FloorSub.value='';
      document.form.Invoice_Room.value='';
    }
  }
}
function Update_OnClick(){
	//20131106 Modify by GoodTV Tanya:調整畫面欄位驗證項目「身份別」、「通訊地址」、「生日」、「電話類」、「E-mail」及「收據開立」
  <%call CheckStringJ("Donor_Name","捐款人")%>
  <%call ChecklenJ("Donor_Name",100,"捐款人")%>
  <%'call ChecklenJ("Category",20,"類別")%>
  <%call ChecklenJ("Sex",2,"性別")%>
  <%call ChecklenJ("Title",20,"稱謂")%>
  <%call CheckStringCheckBoxJ("Donor_Type","身份別")%>
  <%call ChecklenCheckBoxJ("Donor_Type",100,"身份別")%>  
  <%call ChecklenJ("IDNo",10,"身分證/統編")%>
  if(document.form.Birthday.value!=''){
    <%call CheckDateJ("Birthday","出生日期")%>
  }  
  <%call ChecklenJ("Education",20,"教育程度")%>
  <%call ChecklenJ("Occupation",20,"職業別")%>
  <%call ChecklenJ("Marriage",20,"婚姻狀況")%>
  <%call ChecklenJ("Religion",20,"宗教信仰")%>
  <%call ChecklenJ("ReligionName",40,"道場/教會名稱")%>  
  <%call ChecklenJ("Cellular_Phone",40,"手機")%>
  <%call ChecklenJ("Cellular_Phone",40,"手機")%>  
  <%call ChecklenJ("Tel_Office",40,"電話")%>  
  <%call ChecklenJ("Tel_Office_Ext",40,"電話(夜)")%>
  <%call ChecklenJ("Fax",40,"傳真")%>
  <%call ChecklenJ("Email",80,"E-Mail")%>  
  if(document.form.Email.value!=''){
  	<%call CheckEmailJ("Email","E-Mail")%>
  }  
  <%call ChecklenJ("Contactor",100,"聯絡人")%>
  <%call ChecklenJ("OrgName",40,"服務單位")%>
  <!--20130724 Modify by GoodTV Tanya:修改欄位名稱為「聯絡人職稱」-->
  <!--20140212 Modify by GoodTV Tanya:修改「鄉鎮市區」非必輸欄位-->
  <%'call ChecklenJ("JobTitle",40,"聯絡人職稱")%>
  if (document.form.IsAbroadY.checked)
  {  
  	<%call CheckStringJ("City","通訊地址-台灣本島(縣市)")%>  	
  	<%'call CheckStringJ("Area","通訊地址-台灣本島(鄉鎮市區)")%>  	
  	<%call CheckStringJ("Street","通訊地址-台灣本島(大道/路/街/部落 )")%>
  }
	if (document.form.IsAbroadN.checked)
	{
		<%call CheckStringJ("OverseasCountry","通訊地址-海外地址(國家/省城市/區)")%>  	
		<%call CheckStringJ("OverseasAddress","通訊地址-海外地址明細")%>  	
	}
  <%call ChecklenJ("Street",40,"大道/路/街/部落")%>
  <%call ChecklenJ("Lane",10,"巷")%>
  <%call ChecklenJ("Alley",10,"弄/衖")%>
  <%call ChecklenJ("HouseNo",10,"號")%>
  <%call ChecklenJ("HouseNoSub",10,"號之")%>
  <%call ChecklenJ("Floor",10,"樓")%>
  <%call ChecklenJ("FloorSub",10,"樓之")%>
  <%call ChecklenJ("Room",10,"室")%>  
  <%call ChecklenJ("OverseasCountry",50,"通訊地址國家/省城市/區")%>
  <%call ChecklenJ("OverseasAddress",100,"通訊地址")%>
  <%call CheckStringJ("Invoice_Type","收據開立")%>
  <%call ChecklenJ("Invoice_Type",20,"收據開立")%>
  <!--20130917 Mark by GoodTV Tanya -->
  //if(document.form.IsAnonymous.checked){
    <%'call CheckStringJ("NickName","匿名")%>
  //}
  <%'call ChecklenJ("NickName",20,"匿名")%>
  if(document.form.Invoice_Type.value.indexOf('不')==-1){
    <%call CheckStringJ("Invoice_Title","收據抬頭")%>
  }
  <%call ChecklenJ("Invoice_Street",40,"大道/路/街/部落")%>
  <%call ChecklenJ("Invoice_Lane",10,"巷")%>
  <%call ChecklenJ("Invoice_Alley",10,"弄/衖")%>
  <%call ChecklenJ("Invoice_HouseNo",10,"號")%>
  <%call ChecklenJ("Invoice_HouseNoSub",10,"號之")%>
  <%call ChecklenJ("Invoice_Floor",10,"樓")%>
  <%call ChecklenJ("Invoice_FloorSub",10,"樓之")%>
  <%call ChecklenJ("Invoice_Room",10,"室")%>
  <%call ChecklenJ("Invoice_OverseasCountry",50,"收據地址國家/省城市/區")%>
  <%call ChecklenJ("Invoice_OverseasAddress",100,"收據地址")%>
  <%call ChecklenJ("Invoice_Title",100,"收據抬頭")%>
  <%call ChecklenJ("Invoice_IDNo",10,"收據身分證/統編")%>
  if(document.form.IsFdc.checked){
    <%call CheckStringJ("Invoice_IDNo","收據身分證/統編")%>
  }
  <%call CheckNumberJ("Cellular_Phone","手機")%>
  <%call CheckNumberJ("Tel_Office","電話")%>
  <%call CheckNumberJ("Fax","傳真")%>
  <%call SubmitJ("update")%>
}  
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
  location.href='donor.asp';
}
function Add_OnClick(){
  location.href='donate_add.asp?ctype='+document.form.ctype.value+'&donor_id='+document.form.donor_id.value+'';
}
function IsMember_OnClick(){
  if(confirm('您是否確定要將捐款人轉為本會會員？')){
    document.form.action.value='member';
    document.form.submit();
  }
}
--></script>	