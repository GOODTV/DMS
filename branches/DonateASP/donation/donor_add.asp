<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
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

  SQL1="DONOR"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  RS1.Addnew
  RS1("Dept_Id")=Session("dept_id")
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
  '20130731 Modify by GoodTV Tanya:修改文宣品項目
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
  '20130731 Modify by GoodTV Tanya:增加「不主動聯絡-錯址」
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
  RS1("Invoice_Type")=request("Invoice_Type")
'20130917 Mark by GoodTV Tanya 
'  If request("IsAnonymous")<>"" Then
'    RS1("IsAnonymous")="Y"
'  Else
'    RS1("IsAnonymous")="N"
'  End If
'  RS1("NickName")=Data_Plus(request("NickName"))
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
  RS1("IsThanks")="0"
  RS1("IsThanks_Add")="0"
  RS1("Donate_No")="0"
  RS1("Donate_Total")="0"
  RS1("Donate_NoD")="0"
  RS1("Donate_TotalD")="0"
  RS1("Donate_NoC")="0"
  RS1("Donate_TotalC")="0"
  RS1("Donate_NoM")="0"
  RS1("Donate_TotalM")="0"
  RS1("Donate_NoS")="0"
  RS1("Donate_TotalS")="0"
  RS1("Donate_NoA")="0"
  RS1("Donate_TotalA")="0"
  RS1("Donate_NoND")="0"
  RS1("Donate_TotalND")="0"
  RS1("Create_Date")=Date()
  RS1("Create_DateTime")=Now()
  RS1("Create_User")=session("user_name")
  RS1("Create_IP")=Request.ServerVariables("REMOTE_HOST")
  RS1("IsMember")="N"
  RS1("Member_No")=""
  RS1("Member_Type")=""
  RS1("Member_Status")=""
  RS1("IsVolunteer")="N"
  RS1("Volunteer_Type")=""
  RS1("Volunteer_Status")=""
  RS1("IsGroup")="N"
  RS1("Group_No")=""
  RS1("Group_Licence")=""
  RS1("Group_CreateDate")=null
  RS1("Group_Person")=""
  RS1("Group_Person_JobTitle")=""
  RS1("Group_WebUrll")=""
  RS1("Group_Mission")=""
  RS1("Group_Service")=""
  RS1("Group_Other")=""
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
  '20130731 Mark by GoodTV Tanya:隱藏「徵信錄原則」、「群組」及「收據抬頭人數」
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
  '20130731 Add by GoodTV Tanya:增加「收據地址-Attn」
  RS1("Invoice_Attn")=request("Invoice_Attn")
  RS1.Update
  RS1.Close
  Set RS1=Nothing
  session("errnumber")=1
  session("msg")="捐款人資料新增成功 ！"
  SQL="Select @@IDENTITY as ser_no"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then donor_id=RS("ser_no")

  If request("ctype")="donate" Then
    session("msg")="捐款人資料已建立\n您可以新增捐款資料 ！"
    Response.Redirect "donate_add.asp?donor_id="&donor_id
  ElseIf request("ctype")="donate2" Then
    session("msg")="捐款人資料已建立\n您可以新增捐物資料 ！"
    Response.Redirect "donate2_add.asp?donor_id="&donor_id
  ElseIf request("ctype")="pledge" Then
    session("msg")="捐款人資料已建立\n您可以新增轉帳授權書資料 ！"
    Response.Redirect "pledge_add.asp?donor_id="&donor_id
  ElseIf request("ctype")="contribute" Then
    session("msg")="捐款人資料已建立\n您可以新增捐物資料 ！"
    Response.Redirect "../contribute/contribute_add.asp?donor_id="&donor_id
  Else
    session("msg")="捐款人資料已建立\n您可以新增捐款資料 ！"
    Response.Redirect "donate_add.asp?donor_id="&donor_id
  End If
End If

Donate_IsFdc="Y"
SQL1="Select Donate_IsFdc From DEPT Where Dept_Id='"&session("dept_id")&"'"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then
  If RS1("Donate_IsFdc")<>"" Then Donate_IsFdc=RS1("Donate_IsFdc")
End If
RS1.Close
Set RS1=Nothing

If Request("Invoice_Title_Man")<>"" Then
	Invoice_Title_Man=Request("Invoice_Title_Man")
Else
	Invoice_Title_Man="1"
End if
%>
<%Prog_Id="donor"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ctype" value="<%=request("ctype")%>">
      <input type="hidden" name="Introducer_Id" value="<%=request("Introducer_Id")%>">			
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
                          	<!--20131106 Modify by GoodTV Tanya:增加「必輸欄位」註記和判斷-->
                          	<!--20140114 Modify by GoodTV Tanya:增加「字數」限制提醒-->
                          	<!--20140210 Modify by GoodTV Tanya:移除「字數」限制提醒-->
                            <td align="right"><font color="red">*</font>捐款人：</td>
                            <td align="left" colspan="3"><input type="text" name="Donor_Name" size="46" class="font9" maxlength="100" value="<%=request("donor_name2")%>"></td>
                            <td align="left" colspan="4">
                            	<!--20130731 Mark by GoodTV Tanya-->
                              <!--類別：
                              <%
                                SQL="Select Category=CodeDesc From CASECODE Where codetype='Category' Order By Seq"
                                FName="Category"
                                Listfield="Category"
                                menusize="1"
                                BoundColumn=request("category2")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>-->
                              &nbsp;性別：
                              <%
                                SQL="Select Sex=CodeDesc From CASECODE Where codetype='Sex' Order By Seq"
                                FName="Sex"
                                Listfield="Sex"
                                menusize="1"
                                BoundColumn=""
                                '20130731 Modify by GoodTV Tanya:增加性別與稱謂連動
                                call CodeSex (SQL,FName,Listfield,BoundColumn,menusize) 
                                'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                              &nbsp;稱謂：
                              <%
                                SQL="Select Title=CodeDesc From CASECODE Where codetype='Title' Order By Seq"
                                FName="Title"
                                Listfield="Title"
                                menusize="1"
                                'BoundColumn=""
                                '20130731 Modify by GoodTV Tanya:稱謂預設為「先生/小姐/寶號」
                                BoundColumn="先生/小姐/寶號"
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
                                BoundColumn=request("donor_type2")
                                call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                              %>
                            </td>
                          </tr>
                          <tr>
                            <td align="right" width="12%">身分證/統編：</td>
                            <td align="left" width="15%"><input type="text" name="IDNo" size="15" class="font9" maxlength="10" onKeyUp="javascript:UCaseIDNO();" onblur="<%call CheckIDJ("IDNo")%>" value="<%=request("idno2")%>"></td>
                            <td align="right" width="10%">出生日期：</td>
                            <td align="left" width="15%"><%call Calendar("Birthday","")%></td>
                            <input type="hidden" name="Education" size="15" class="font9" maxlength="10" value="">
                            <input type="hidden" name="Occupation" size="15" class="font9" maxlength="10" value="">
                            <input type="hidden" name="Marriage" size="15" class="font9" maxlength="10" value="">
                            <input type="hidden" name="Religion" size="15" class="font9" maxlength="10" value="">
                            <input type="hidden" name="ReligionName" size="15" class="font9" maxlength="10" value="">
                            <td align="right">手機：</td>
                            <td align="left"><input type="text" name="Cellular_Phone" size="15" class="font9" maxlength="40" value="<%=Cellular_Phone%>"></td>
                            <!--<td align="right" width="10%">教育程度：</td>
                            <td align="left" width="15%">
                            <%
                              'SQL="Select Education=CodeDesc From CASECODE Where codetype='Education' Order By Seq"
                              'FName="Education"
                              'Listfield="Education"
                              'menusize="1"
                              'BoundColumn=""
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
                              'BoundColumn=""
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
                              'BoundColumn=""
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
                              'BoundColumn=""
                              'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                            <td colspan="4">
                            	&nbsp;&nbsp;&nbsp;&nbsp;道場/教會名稱：
                              <input type="text" name="ReligionName" size="39" class="font9" maxlength="40">
                            </td>
                          </tr>-->
                          <tr>
                            <%
                              Cellular_Phone=""
                              Tel_Office=""
                              If request("Tel_Office2")<>"" Then
                                If Len(request("Tel_Office2"))>=2 Then
                                  If Left(request("Tel_Office2"),2)="09" Then
                                    Cellular_Phone=request("Tel_Office2")
                                  Else
                                    Tel_Office=request("Tel_Office2")
                                  End If
                                Else
                                  Tel_Office=request("Tel_Office2")
                                End If
                              End If
                            %>                            
                            <td align="right">電話：</td>
                            <td align="left" colspan="3"><input type="text" name="Tel_Office_Loc" size="5" class="font9" maxlength="5" value="">-<input type="text" name="Tel_Office" size="15" class="font9" maxlength="40" value="<%=Tel_Office%>">-<input type="text" name="Tel_Office_Ext" size="5" class="font9" maxlength="5" value=""></td>
                            <input type="hidden" name="Tel_Home" size="15" class="font9" maxlength="10" value="">
                            <!--<td align="right">電話(夜)：</td>
                            <td align="left"><input type="text" name="Tel_Home" size="15" class="font9" maxlength="40"></td>-->
                            <td align="right">傳真：</td>
                            <td align="left"><input type="text" name="Fax_Loc" size="5" class="font9" maxlength="5">-<input type="text" name="Fax" size="15" class="font9" maxlength="40"></td>
                          </tr>
                          <tr>
                            <td align="right">E-Mail：</td>
                            <td align="left"><input type="text" name="Email" size="25" class="font9" maxlength="80"></td>
                            <td align="right">聯絡人：</td>
                            <td align="left"><input type="text" name="Contactor" size="15" class="font9" maxlength="40"></td>
                            <input type="hidden" name="OrgName" size="15" class="font9" maxlength="10" value="">
                            <!--<td align="right">服務單位：</td>
                            <td align="left"><input type="text" name="OrgName" size="15" class="font9" maxlength="40"></td>-->
                            <!--20130731 Modify by GoodTV Tanya:修改欄位名稱為「聯絡人職稱」-->
                            <!--td align="right">聯絡人職稱：</td>
                            <td align="left"><input type="text" name="JobTitle" size="15" class="font9" maxlength="40"></td-->
                          </tr>
                          <tr>
                            <td align="right"><font color="red">*</font>通訊地址：</td>
                            <td align="left" colspan="7">
                              <input type="radio" name="IsAbroad" id="IsAbroadY" value="N" <%If request("IsAbroad2")="N" Or request("IsAbroad2")="" Then Response.Write "checked" End If%> >台灣本島
                              <%
                            	  ZipCode=""
                            	  If request("area2")<>"" Then ZipCode=Left(request("area2"),3)
                            	  call CodeCity2 ("form","ZipCode",ZipCode,"City",request("city2"),"Area",request("area2"),"Y")
                            	%>
                            	<input type='text' class='font9' name='Street' size='25' maxlength='40'>大道/路/街/部落
                              <br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                              <%
                                SQL="Select SectionType=CodeDesc From CASECODE Where codetype='SectionType' Order By Seq"
                                FName="SectionType"
                                Listfield="SectionType"
                                menusize="1"
                                BoundColumn=""
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>段
                            	<input type='text' class='font9' name='Lane' size='5' maxlength='10'>巷
                            	<input type='text' class='font9' name='Alley' size='5' maxlength='10'>
                              <%
                                SQL="Select AlleyType=CodeDesc From CASECODE Where codetype='AlleyType' Order By Seq"
                                FName="AlleyType"
                                Listfield="AlleyType"
                                menusize="1"
                                BoundColumn="弄"
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                            	<input type='text' class='font9' name='HouseNo' size='5' maxlength='10'>之<input type='text' class='font9' name='HouseNoSub' size='5' maxlength='10'>號
                            	<input type='text' class='font9' name='Floor' size='5' maxlength='10'>樓之<input type='text' class='font9' name='FloorSub' size='5' maxlength='10'>
                            	<input type='text' class='font9' name='Room' size='5' maxlength='10'>室
                                Attn：<input type='text' class='font9' name='Attn' size='20' maxlength='50'>
                              <br />
                              <input type="radio" name="IsAbroad" id="IsAbroadN" value="Y" <%If request("IsAbroad2")="Y" Then Response.Write "checked" End If%> >海外地址
                              <input type='text' class='font9' name='OverseasAddress' size='50' maxlength='100'>
                              國家/省城市/區
                              <input type='text' class='font9' name='OverseasCountry' size='20' maxlength='50'>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">文宣品：</td>
                            <td align="left" colspan="3">
                            	<!--20130731 Modify by GoodTV Tanya:修改文宣品選項-->                            	
                            	   紙本月刊本數<input type="text" class="font9" name="IsSendNewsNum" size='3' maxlength='3' value='1'>
                            	<input type="checkbox" name="IsDVD" value="Y" checked>DVD
                              <input type="checkbox" name="IsSendEpaper" value="Y" checked>電子文宣
                            	<!--<input type="checkbox" name="IsSendNews" value="Y" <%If request("IsSendNews2")="Y" Then Response.Write "checked" End If%> >會訊
                              <input type="checkbox" name="IsSendEpaper" value="Y" <%If request("IsSendEpaper2")="Y" Then Response.Write "checked" End If%> >電子報
                              <input type="checkbox" name="IsSendYNews" value="Y" <%If request("IsSendYNews2")="Y" Then Response.Write "checked" End If%> >年報
                              <input type="checkbox" name="IsBirthday" value="Y" <%If request("IsBirthday2")="Y" Then Response.Write "checked" End If%> >生日卡
                              <input type="checkbox" name="IsXmas" value="Y" <%If request("IsXmas2")="Y" Then Response.Write "checked" End If%> >賀卡-->
                            </td>
                            <!--20130731 Modify by GoodTV Tanya:增加「不主動聯絡-錯址」-->
                            <!--20130923 Modify by GoodTV Tanya:拆開「不主動聯絡」及「錯址」-->
                            <td align="left" colspan="1">
                            	<input type="checkbox" name="IsContact" value="N">不主動聯絡
                            </td>
                            <td align="left" colspan="3">
                            	<input type="checkbox" name="IsErrAddress" value="Y">錯址
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
                                BoundColumn="年度證明及收據"
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                              <input type="checkbox" name="IsSameAddress" value="Y" OnClick="SameAddress_OnClick()"><font color="#ff0000">收據資料同上</font>	
                              <!--<input type="checkbox" name="IsAnonymous" value="Y">徵信錄匿名-->
                            	<!--&nbsp;匿名：
                            	<input type="text" name="NickName" size="13" class="font9" maxlength="20">-->
                            	&nbsp;收據稱謂：
                              <%
                                SQL="Select Title2=CodeDesc From CASECODE Where codetype='Title' Order By Seq"
                                FName="Title2"
                                Listfield="Title2"
                                menusize="1"
                                BoundColumn=""
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                            </td>	
                          </tr>
                          <tr>
                            <td align="right">收據地址：</td>
                            <td align="left" colspan="7">
                              <input type="radio" name="IsAbroad_Invoice" id="IsAbroad_InvoiceY" value="N" <%If request("IsAbroad2")="N" Or request("IsAbroad2")="" Then Response.Write "checked" End If%> >台灣本島
                            	<%
                            	  Invoice_ZipCode=""
                            	  If request("area2")<>"" Then Invoice_ZipCode=Left(request("area2"),3)
                            	  call CodeCity2 ("form","Invoice_ZipCode",Invoice_ZipCode,"Invoice_City",request("city2"),"Invoice_Area",request("area2"),"N")
                            	%>
                            	<!--20130801 Modify by GoodTV Tanya:修改欄位長度同通訊地址-->
                            	<input type='text' class='font9' name='Invoice_Street' size='25' maxlength='40'>大道/路/街/部落
                              <br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                              <%
                                SQL="Select Invoice_Section=CodeDesc From CASECODE Where codetype='SectionType' Order By Seq"
                                FName="Invoice_Section"
                                Listfield="Invoice_Section"
                                menusize="1"
                                BoundColumn=""
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>段
                            	<input type='text' class='font9' name='Invoice_Lane' size='5' maxlength='10'>巷
                            	<input type='text' class='font9' name='Invoice_Alley' size='5' maxlength='10'>
                              <%
                                SQL="Select Invoice_AlleyType=CodeDesc From CASECODE Where codetype='AlleyType' Order By Seq"
                                FName="Invoice_AlleyType"
                                Listfield="Invoice_AlleyType"
                                menusize="1"
                                BoundColumn="弄"
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                            	<input type='text' class='font9' name='Invoice_HouseNo' size='5' maxlength='10'>之<input type='text' class='font9' name='Invoice_HouseNoSub' size='5' maxlength='10'>號
                            	<input type='text' class='font9' name='Invoice_Floor' size='5' maxlength='10'>樓之<input type='text' class='font9' name='Invoice_FloorSub' size='5' maxlength='10'>
                            	<input type='text' class='font9' name='Invoice_Room' size='5' maxlength='10'>室
                            		<!--20130731 Add by GoodTV Tanya:增加「收據地址-Attn」-->
                            		Attn：<input type='text' class='font9' name='Invoice_Attn' size='20' maxlength='50'>
                              <br />
                              <input type="radio" name="IsAbroad_Invoice" id="IsAbroad_InvoiceN" value="Y" <%If request("IsAbroad2")="Y" Then Response.Write "checked" End If%> >海外地址
                              <input type='text' class='font9' name='Invoice_OverseasAddress' size='50' maxlength='100'>
                              國家/省城市/區
                              <input type='text' class='font9' name='Invoice_OverseasCountry' size='20' maxlength='50'>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">收據抬頭：</td>
                            <!--20140114 Modify by GoodTV Tanya:增加「字數」限制提醒-->
                            <td align="left" colspan="3"><input type="text" name="Invoice_Title" size="46" class="font9" maxlength="100" value="<%=request("donor_name2")%>" onblur="javascript:CheckDonorNameLen('2',this.value);"></td>
                            <td align="left" colspan="4">
                              收據身分證/統編：
                              <input type="text" name="Invoice_IDNo" size="15" class="font9" maxlength="10" onKeyUp="javascript:UCaseInvoiceIDNo();" onblur="<%call CheckIDJ("Invoice_IDNo")%>" value="<%=request("idno2")%>">
                            </td>
                          </tr>
                          <!--20130731 Modify by GoodTV Tanya:隱藏「徵信錄原則」、「群組」及「收據抬頭人數」 --> 
                          <!--20130917 Modify by GoodTV Tanya:開啟「徵信錄原則」 --> 
                          <tr>                           	                         	
                            <td align="right">徵信錄原則：</td>
                            <td align="left" >
                            <%
                              SQL="Select Report_Type=CodeDesc From CASECODE Where codetype='ReportType' Order By Seq"
                              FName="Report_Type"
                              Listfield="Report_Type"
                              menusize="1"
                              BoundColumn="不刊登"
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
                              BoundColumn=Request("Report_Group")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                            <td align="right">收據抬頭人數：</td>
                            <td align="left" ><input type="text" name="Invoice_Title_Man" size="2" class="font9" maxlength="2" value="<%=Invoice_Title_Man%>"></td>-->
                          </tr>
                          <%If Donate_IsFdc="Y" Then%>
                          <tr>
                            <td align="left" colspan="8">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="IsFdc" value="Y" ><font color="#ff0000">本捐款人願意將捐款資料上傳至國稅局以供報稅用</font>&nbsp;(&nbsp;請填寫捐款人收據身分證/統編&nbsp;)</td>
                          </tr>
                          <%Else%>
                          <input type="hidden" name="IsFdc" value="N">
                          <%End If%>
                          <tr>
                            <td align="right">捐款人備註：</td>
                            <td align="left" colspan="7">
                            	<textarea name="Remark" rows="5" cols="90" class="font9"></textarea>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">想對GOOD TV說的話：</td>
                            <td align="left" colspan="7">
                            	<textarea name="ToGOODTV" rows="3" cols="90" class="font9"></textarea>
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
  if(document.form.City.value!=''&&document.form.Area.value==''){
    ChgCity(document.form.City.value,document.form.Area,document.form.ZipCode);
  }
  if(document.form.Invoice_City.value!=''&&document.form.Invoice_Area.value==''){
    ChgCity(document.form.Invoice_City.value,document.form.Invoice_Area,document.form.Invoice_ZipCode);
  }
  document.form.Donor_Name.focus();
}
function CheckDonorNameLen(Type,objValue){
	var regx = /^[u4E00-u9FA5]+$/;
	
	if (Type == '1'){		
		if(!regx.test(objValue)){
			if(objValue.length > 11)
				alert("「捐款人」中文字超過11字，列印收據會被截斷!");
		}
		else{
			if(objValue.length > 22)
				alert("「捐款人」英數字超過22字，列印收據會被截斷!");
		}
	}
	else if(Type == '2'){	
		if(!regx.test(objValue)){
			if(objValue.length > 20)
				alert("「收據抬頭」中文字超過20字，列印收據會被截斷!");
		}
		else{
			if(objValue.length > 40)
				alert("「收據抬頭」英數字超過40字，列印收據會被截斷!");
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
function Save_OnClick(){
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
  <%call ChecklenJ("Tel_Office",40,"電話(日)")%>
  <%call ChecklenJ("Tel_Home",40,"電話(夜)")%>
  <%call ChecklenJ("Fax",40,"傳真")%>
  <%call ChecklenJ("Email",80,"E-Mail")%>
  if(document.form.Email.value!=''){
  	<%call CheckEmailJ("Email","E-Mail")%>
  } 
  <%call ChecklenJ("Contactor",40,"聯絡人")%>
  <%call ChecklenJ("OrgName",40,"服務單位")%>
  <!--20130731 Modify by GoodTV Tanya:修改欄位名稱為「聯絡人職稱」-->
  <%'call ChecklenJ("JobTitle",40,"聯絡人職稱")%>
  if (document.form.IsAbroadY.checked)
  <!--20140212 Modify by GoodTV Tanya:修改「鄉鎮市區」非必輸欄位-->
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
  <%call SubmitJ("save")%>
}
function Cancel_OnClick(){
  if(document.form.ctype.value=='donate'){
    location.href='donate_input.asp';
  }else if(document.form.ctype.value=='pledge'){
    location.href='pledge_input.asp';
  }else if(document.form.ctype.value=='contribute'){
    location.href='contribute_input.asp';
  }else{
    location.href='donor.asp';
  }
}
--></script>