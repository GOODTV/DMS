<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="get" Then
  Member_No=""
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
  SQL="Update DONOR Set Member_No='"&Member_No&"' Where Donor_Id='"&request("donor_id")&"'"
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
      SQL="Update DONOR Set Member_No='"&Member_No_Old&"' Where Donor_Id='"&request("donor_id")&"'"
      Set RS=Conn.Execute(SQL)
    End If
  Wend
  session("errnumber")=1
  session("msg")="會員編號取號成功 ！"
End If

If request("action")="update" Then
	'20130912Add by GoodTV-Tanya:修改地址格式同捐款人
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
  
  Check_MemberNo=True
  If request("Member_No")<>request("Member_No_Old") Then
    SQL1="Select * From DONOR Where Member_No='"&request("Member_No")&"'"
    Call QuerySQL(SQL1,RS1)
    If Not RS1.EOF Then
      Check_MemberNo=False
      session("errnumber")=1
      session("msg")="您輸入的會員編號已經存在 ！"
    End If
    RS1.Close
    Set RS1=Nothing
  End If
  
  If Check_MemberNo Then
    SQL1="Select * From DONOR Where Donor_id='"&request("donor_id")&"'"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,3
    RS1("Member_No")=request("Member_No")
    If request("Introducer_Name")<>"" Then
'      RS1("Introducer_Id")=request("Introducer_Id")
      RS1("Introducer_Name")=request("Introducer_Name")
    Else
      RS1("Introducer_Name")=""
    End If
    RS1("Donor_Name")=request("Donor_Name")
    RS1("Sex")=request("Sex")
    RS1("Title")=request("Title")
    '20130911Mark by GoodTV-Tanya
    'RS1("Category")=request("Category")
    RS1("Donor_Type")=request("Donor_Type")
    RS1("IDNo")=request("IDNo")
    If request("Birthday")<>"" Then
      RS1("Birthday")=request("Birthday")
    Else
      RS1("Birthday")=null
    End If
'    RS1("Education")=request("Education")
'    RS1("Occupation")=request("Occupation")
'    RS1("Marriage")=request("Marriage")
'    RS1("Religion")=request("Religion")
'    RS1("ReligionName")=request("ReligionName")
    RS1("Cellular_Phone")=request("Cellular_Phone")
    RS1("Tel_Office")=request("Tel_Office")
    RS1("Tel_Office_Loc")=request("Tel_Office_Loc")
  	RS1("Tel_Office_Ext")=request("Tel_Office_Ext")
'    RS1("Tel_Home")=request("Tel_Home")
    RS1("Fax")=request("Fax")
    RS1("Fax_Loc")=request("Fax_Loc")
    RS1("Email")=request("Email")
    RS1("Contactor")=request("Contactor")
'    RS1("OrgName")=request("OrgName")
'    RS1("JobTitle")=request("JobTitle")
    '20130912Modify by GoodTV-Tanya:修改地址格式同捐款人
    RS1("IsAbroad")=request("IsAbroad")    
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
	    RS1("Attn")=request("Attn") 
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
	    RS1("Attn")=""
	    RS1("OverseasCountry")=request("OverseasCountry")
	    RS1("OverseasAddress")=request("OverseasAddress")
	    RS1("Address")=Address
	  End If
'    If request("IsAbroad")<>"" Then
'      RS1("IsAbroad")="Y"
'    Else
'      RS1("IsAbroad")="N"
'    End If
'    RS1("ZipCode")=request("ZipCode")
'    RS1("City")=request("City")
'    RS1("Area")=request("Area")
'    RS1("Address")=request("Address")
    '20130911 Modify by GoodTV Tanya:修改文宣品項目
    If request("IsSendNewsNum")<>"" Then
    	If request("IsSendNewsNum") > 0 Then
      	RS1("IsSendNews")="Y"
      	RS1("IsSendNewsNum")=request("IsSendNewsNum")
    	Else
      	RS1("IsSendNews")="N"
      	RS1("IsSendNewsNum")="0"
      End If
    End If
    If request("IsSendEpaper")<>"" Then
      RS1("IsSendEpaper")="Y"
    Else
      RS1("IsSendEpaper")="N"
    End If
	  '20130911 Modify by GoodTV Tanya:增加「不主動聯絡-錯址」
	  If request("IsErrAddress")<>"" Then
	    RS1("IsErrAddress")="Y"
	  Else
	    RS1("IsErrAddress")="N"
	  End If
    '20130911Mark by GoodTV-Tanya
'    If request("IsSendYNews")<>"" Then
'      RS1("IsSendYNews")="Y"
'    Else
'      RS1("IsSendYNews")="N"
'    End If
'    If request("IsBirthday")<>"" Then
'      RS1("IsBirthday")="Y"
'    Else
'      RS1("IsBirthday")="N"
'    End If  
'    If request("IsXmas")<>"" Then
'      RS1("IsXmas")="Y"
'    Else
'      RS1("IsXmas")="N"
'    End If    
'    RS1("Invoice_Type")=request("Invoice_Type")
'    If request("IsAnonymous")<>"" Then
'      RS1("IsAnonymous")="Y"
'    Else
'      RS1("IsAnonymous")="N"
'    End If
'    RS1("NickName")=request("NickName")
'    RS1("Title2")=request("Title2")
'    RS1("Invoice_Title")=request("Invoice_Title")
'    RS1("Invoice_IDNo")=request("Invoice_IDNo")
'    If request("IsAbroad_Invoice")<>"" Then
'      RS1("IsAbroad_Invoice")="Y"
'    Else
'      RS1("IsAbroad_Invoice")="N"
'    End If
'    RS1("Invoice_ZipCode")=request("Invoice_ZipCode")
'    RS1("Invoice_City")=request("Invoice_City")
'    RS1("Invoice_Area")=request("Invoice_Area")
'    RS1("Invoice_Address")=request("Invoice_Address")
'    If request("IsFdc")<>"" Then
'      RS1("IsFdc")=request("IsFdc")
'    Else
'      RS1("IsFdc")="N"
'    End If
    RS1("Remark")=request("Remark")
    RS1("LastUpdate_Dept_Id")=Session("dept_id")
    RS1("LastUpdate_Date")=Date()
    RS1("LastUpdate_DateTime")=Now()
    RS1("LastUpdate_User")=session("user_name")
    RS1("LastUpdate_IP")=Request.ServerVariables("REMOTE_HOST")
    '20130911Mark by GoodTV-Tanya
'    If request("Member_JoinDate")<>"" Then
'      RS1("Member_JoinDate")=request("Member_JoinDate")
'    Else
'      RS1("Member_JoinDate")=null
'    End If
'    RS1("Member_Type")=request("Member_Type")
    RS1("Member_Status")=request("Member_Status")
    '20130911Mark by GoodTV-Tanya
    'RS1("Member_Group")=request("Member_Group")
    'RS1("IsVolunteer")=request("IsVolunteer")
    'RS1("Volunteer_Service")=request("Volunteer_Service")
    'RS1("Volunteer_Other")=request("Volunteer_Other")
    RS1.Update
    RS1.CLose
    Set RS1=Nothing
    
    SQL="Update DONOR Set Introducer_Name='"&request("Donor_Name")&"' Where Introducer_Id='"&request("donor_id")&"'"
    Set RS=Conn.Execute(SQL)
  
    '修改捐款人捐款紀錄
    call Declare_DonorId (request("donor_id"))
      
    session("errnumber")=1
    session("msg")="讀者資料修改成功 ！"
  End If
End If

If request("action")="delete" Then
  SQL1="Select * From DONATE Where Donor_id='"&request("donor_id")&"'"
  Call QuerySQL(SQL1,RS1)
  If Not RS1.EOF Then
    session("errnumber")=1
    session("msg")="此會員尚有繳費記錄存在，請先刪除繳費明細，才能刪除會員！"
  Else
    SQL2="Select * From DONOR Where Donor_id='"&request("donor_id")&"'"
    Call QuerySQL(SQL2,RS2)  
    If Not RS2.EOF Then
      log_type="刪除讀者"
      log_desc=CStr(RS2("Donor_id"))&RS2("Donor_Name")
      call syslog(log_type,log_desc)
      SQL="Delete From DONOR Where Donor_id='"&request("donor_id")&"'"
      Call ExecSQL(SQL)
      session("errnumber")=1
      session("msg")="讀者資料刪除成功 ！"       
      Response.Redirect "member.asp"
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
IsDonation=RS_IS("IsDonation")
IsContribute=RS_IS("IsContribute")
IsMember=RS_IS("IsMember")
IsMemberNo=RS_IS("IsMemberNo")
IsMemberPre=RS_IS("IsMemberPre")
IsShopping=RS_IS("IsShopping")
'20130911Mark by GoodTV-Tanya
'Donate_IsFdc=RS_IS("Donate_IsFdc")
RS_IS.Close
Set RS_IS=Nothing

SQL="Select * From DONOR Where Donor_id='"&request("donor_id")&"'"
Call QuerySQL(SQL,RS)
%>
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>讀者資料維護</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <%If request("ctype")<>"" Then%>
      <input type="hidden" name="ctype" value="<%=request("ctype")%>">
      <%Else%>
      <input type="hidden" name="ctype" value="member_edit">
      <%End If%>      	
      <input type="hidden" name="donor_id" value="<%=request("donor_id")%>">
      <input type="hidden" name="IsMember" value="<%=IsMember%>">
      <input type="hidden" name="IsMemberNo" value="<%=IsMemberNo%>">
      <input type="hidden" name="IsMemberPre" value="<%=IsMemberPre%>">      	
      <input type="hidden" name="Member_No_Old" value="<%=RS("Member_No")%>">
      <input type="hidden" name="Introducer_Id" value="<%=RS("Introducer_Id")%>">		
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
                      <td class="table62-bg">　</td>
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">讀者資料維護</td>
                      <td class="table63-bg">　</td>
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
                <td class="table62-bg">　</td>
                <td valign="top">
                  <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                    <tr>
                      <td width="99%">
                        <table border="1" cellpadding="2" style="border-collapse: collapse" width="100%" height="25" cellspacing="1">
                          <tr>
                            <td class="button2-bg"><img border="0" src="../images/red_arrow.gif" align="texttop">讀者資料</td>
                            <!--<td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="member_donate_data.asp?donor_id=<%=request("donor_id")%>">繳費記錄</a></td>
                            <td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="member_pledge_data.asp?donor_id=<%=request("donor_id")%>">轉帳授權書記錄</a></td>
                            <%If IsContribute="Y" Then%><td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="../contribute/member_contribute_data.asp?donor_id=<%=request("donor_id")%>">物品捐贈記錄</a></td><%End If%>
                            <%If IsMember="Y" Then%><td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="member_signup_data.asp?donor_id=<%=request("donor_id")%>">活動報名記錄</a></td><%End if%>
                            <%If IsShopping="Y" Then%><td class="button-bg"><img border="0" src="../images/red_arrow.gif" align="texttop"><a class="tool" href="../shopping/mrmber_shopping_data.asp?donor_id=<%=request("donor_id")%>">商品義賣記錄</a></td><%End If%>-->
                          </tr>
                        </table>
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" width="100%">
                        <table border="0" cellpadding="1" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <%
                            If IsMember="Y" Then
                              Response.Write "<tr>"
                              Response.Write "  <td align=""right"">會員編號：</td>"
                              If IsMemberNo="Y" Then
                                Response.Write "<td align=""left"" colspan=""7"">"
                                Response.Write "	<input type=""text"" name=""Member_No"" size=""10"" maxlength=""20"" class=""font9t"" readonly onkeyup=""javascript:UCaseMemberNo();"" value="""&RS("Member_No")&""">"
                                If RS("Member_No")="" Then
                                  Response.Write "&nbsp;<input type=""button"" value=""會員編號取號"" name=""get"" class=""cbutton"" style=""cursor:hand"" onClick=""Get_Member_OnClick()"">"
                                Else
                                  Response.Write "&nbsp;<input type=""checkbox"" name=""Issue_Type"" value=""M"" OnClick=""Issue_Type_OnClick()"">修改會員編號"
                                End If
                                Response.Write "</td>"
                              Else
                                Response.Write "<td align=""left"" colspan=""7"">"
                                Response.Write "	<input type=""text"" name=""Member_No"" size=""10"" maxlength=""20"" class=""font9"" onkeyup=""javascript:UCaseMemberNo();"" value="""&RS("Member_No")&""">"
                                Response.Write "</td>"
                              End If
                              Response.Write "</tr>"
                            Else
                              Response.Write "<input type=""hidden"" name=""Member_No"" vlaue="""&RS("Member_No")&""">"
                            End If
                          %>
                          <tr>
                            <td align="right">姓名：</td>
                            <td align="left" colspan="3"><input type="text" name="Donor_Name" size="46" class="font9" maxlength="100" value="<%=RS("Donor_Name")%>"></td>
                            <td align="left" colspan="4">
                              <table width="100%" border="0" cellpadding="0" cellspacing="0" >
                                <tr>
                                	<!--20130911Mark by GoodTV-Tanya-->
                                	<!--<td width="90">
                                    類別：
                                    <%
                                      category=""
                                      If RS("category")<>"" Then category=RS("category")
                                      SQL1="Select Category=CodeDesc From CASECODE Where codetype='Category' Order By Seq"
                                      Set RS1 = Server.CreateObject("ADODB.RecordSet")
                                      RS1.Open SQL1,Conn,1,1
                                      Response.Write "<SELECT Name='Category' size='1' style='font-size: 9pt; font-family: 新細明體' OnChange=""Donor_Category_OnChange(this.value)"">"
                                      Response.Write "<OPTION>" & " " & "</OPTION>"
                                      While Not RS1.EOF
                                        If Cstr(category)=Cstr(RS1("Category")) Then
                                        	Response.Write "<OPTION value='"&RS1("Category")&"' selected >"&RS1("Category")&"</OPTION>"
                                        Else
                                        	Response.Write "<OPTION value='"&RS1("Category")&"'>"&RS1("Category")&"</OPTION>"
                                        End If
                                        RS1.MoveNext
                                      Wend
                                      Response.Write "</SELECT>"
                                      RS1.Close
                                      Set RS1=Nothing
                                      
                                      display="block"
                                      If Cstr(category)<>"" Then
                                        If Instr(Cstr(category),"個人")>0 Then display="block"
                                      End If
                                    %>
                                	</td>-->
                                	<td width="80" id="donorsex" style="display:<%=display%>">
                                    性別：
                                    <%
                                      SQL="Select Sex=CodeDesc From CASECODE Where codetype='Sex' Order By Seq"
                                      FName="Sex"
                                      Listfield="Sex"
                                      menusize="1"
                                      BoundColumn=RS("Sex")
                                      
                                      '20130911 Modify by GoodTV Tanya:增加性別與稱謂連動
                                      call CodeSex (SQL,FName,Listfield,BoundColumn,menusize)
                                      'call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                                    %>
                                	</td>
                                	<td width="145" id="donortitle" style="display:<%=display%>">
                                    稱謂：
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
                              </table>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">身份別：</td>
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
                          	<td align="right">狀態：</td>
                            <td align="left" colspan="7">    	 			                            
                              <%
                                SQL="Select Member_Status=CodeDesc From CASECODE Where codetype='MemberStatus' Order By Seq"
                                FName="Member_Status"
                                Listfield="Member_Status"
                                menusize="1"
                                BoundColumn=RS("Member_Status")
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                            </td>
                          </tr>
                          <!--20130912Mark by GoodTV-Tanya-->
                          <!--
                          <tr>
                            <td align="right">會員別：</td>
                            <td align="left">
                            <%
                              SQL="Select Member_Type=CodeDesc From CASECODE Where codetype='MemberType' Order By Seq"
                              FName="Member_Type"
                              Listfield="Member_Type"
                              menusize="1"
                              BoundColumn=RS("Member_Type")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>                            
                            <td align="right">加入日期：</td>
                            <td align="left"><%call Calendar("Member_JoinDate",RS("Member_JoinDate"))%></td>                            
                           
                            <td align="right">組別：</td>
                            <td align="left"><input type="text" name="Member_Group" size="15" class="font9" maxlength="20" value="<%=RS("Member_Group")%>"></td>
                          </tr>
                          <tr>
                            <td align="right">擔任志工：</td>
                            <td align="left" colspan="7"><input type="checkbox" name="IsVolunteer" value="Y" <%If RS("IsVolunteer")="Y" Then Response.Write "checked" End If%> >有志工證</td>
                          </tr>
                          <tr>
                            <td align="right">志工項目：</td>
                            <td align="left" colspan="7">
                              <%
                                SQL="Select Volunteer_Service=CodeDesc From CASECODE Where codetype='VolunteerService' Order By Seq"
                                FName="Volunteer_Service"
                                Listfield="Volunteer_Service"
                                BoundColumn=RS("Volunteer_Service")
                                call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                              %>
                              <input type="text" name="Volunteer_Other" size="15" class="font9" maxlength="100" value="<%=RS("Volunteer_Other")%>">
                            </td>
                          </tr>-->
                          <tr>
                            <td width="99%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                          </tr>
                          <tr>
                            <td align="right" width="12%">身分證/統編：</td>
                            <td align="left" width="15%"><input type="text" name="IDNo" size="15" class="font9" maxlength="10" onKeyUp="javascript:UCaseIDNO();" value="<%=RS("IDNo")%>"></td>
                            <td align="right" width="10%">出生日期：</td>
                            <td align="left" width="15%"><%call Calendar("Birthday",RS("Birthday"))%></td>
                            <!--20140124 Mark by GoodTV Tanya:隱藏部份項目及修改電話及地址格式-->
                            <td align="right">手機：</td>
                            <td align="left"><input type="text" name="Cellular_Phone" size="15" class="font9" maxlength="40" value="<%=RS("Cellular_Phone")%>"></td>                            
                            <!--<td align="right" width="10%">教育程度：</td>
                            <td align="left" width="15%">
                            <%
                              SQL="Select Education=CodeDesc From CASECODE Where codetype='Education' Order By Seq"
                              FName="Education"
                              Listfield="Education"
                              menusize="1"
                              BoundColumn=RS("Education")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                            <td align="right" width="8%">職業別：</td>
                            <td align="left" width="15%">
                            <%
                              SQL="Select Occupation=CodeDesc From CASECODE Where codetype='Occupation' Order By Seq"
                              FName="Occupation"
                              Listfield="Occupation"
                              menusize="1"
                              BoundColumn=RS("Occupation")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>-->
                          <!--</tr>
                          <tr>
                            <td align="right">婚姻狀況：</td>
                            <td align="left">
                            <%
                              SQL="Select Marriage=CodeDesc From CASECODE Where codetype='Marriage' Order By Seq"
                              FName="Marriage"
                              Listfield="Marriage"
                              menusize="1"
                              BoundColumn=RS("Marriage")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                            <td align="right">宗教信仰：</td>
                            <td align="left">
                            <%
                              SQL="Select Religion=CodeDesc From CASECODE Where codetype='Religion' Order By Seq"
                              FName="Religion"
                              Listfield="Religion"
                              menusize="1"
                              BoundColumn=RS("Religion")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                            <td colspan="4">
                            	&nbsp;&nbsp;&nbsp;&nbsp;所屬教會：
                              <input type="text" name="ReligionName" size="39" class="font9" maxlength="40" value="<%=RS("ReligionName")%>">
                            </td>
                          </tr>-->
                          <tr>                            
                            <td align="right">電話：</td>
                            <td align="left" colspan="3">
                            	<input type="text" name="Tel_Office_Loc" size="5" class="font9" maxlength="5" value='<%=RS("Tel_Office_Loc")%>'>-<input type="text" name="Tel_Office" size="15" class="font9" maxlength="40" value="<%=RS("Tel_Office")%>">-<input type="text" name="Tel_Office_Ext" size="5" class="font9" maxlength="5" value="<%=RS("Tel_Office_Ext")%>">
                            </td>
                            <!--<td align="right">電話(夜)：</td>
                            <td align="left"><input type="text" name="Tel_Home" size="15" class="font9" maxlength="40" value="<%=RS("Tel_Home")%>"></td>-->
                            <td align="right">傳真：</td>
                            <td align="left"><input type="text" name="Fax_Loc" size="5" class="font9" maxlength="5" value="<%=RS("Fax_Loc")%>">-<input type="text" name="Fax" size="15" class="font9" maxlength="40" value="<%=RS("Fax")%>"></td>
                          </tr>
                          <tr>
                            <td align="right">E-Mail：</td>
                            <td align="left"><input type="text" name="Email" size="25" class="font9" maxlength="80" value="<%=RS("Email")%>"></td>
                            <td align="right">聯絡人：</td>
                            <td align="left"><input type="text" name="Contactor" size="15" class="font9" maxlength="40" value="<%=RS("Contactor")%>"></td>
                            <!--<td align="right">服務單位：</td>
                            <td align="left"><input type="text" name="OrgName" size="15" class="font9" maxlength="40" value="<%=RS("OrgName")%>"></td>
                            <td align="right">職稱：</td>
                            <td align="left"><input type="text" name="JobTitle" size="15" class="font9" maxlength="40" value="<%=RS("JobTitle")%>"></td>-->
                          </tr>
                          <tr>
                            <td align="right">通訊地址：</td>
                            <td align="left" colspan="7">
                            	<table width="100%" border="0" cellpadding="0" cellspacing="0" >
                                <tr>
                                	<!--Modify by GoodTV-Tanya:修改地址格式同捐款人-->
                                	<td align="left">
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
                            	</table>   
                            </td>                            
                          </tr>
                          <tr>
                          	<td align="right">介紹人：</td>
                            <td align="left" colspan="7">
                            	<input type="text" name="Introducer_Name" size="12" maxlength="80" class="font9" value="<%=RS("Introducer_Name")%>"><!--&nbsp;<a href onClick="window.open('member_show.asp?LinkId=Introducer_Id&LinkName=Introducer_Name&Donor_Id=<%=request("donor_id")%>','','status=yes,scrollbars=yes,top=100,left=120,width=550,height=550')" style="cursor:hand"><img border="0" src="../images/personal.gif" width="19"></a>--></td>
                          </tr>
                          <tr>
                            <td align="right">文宣品：</td>
                            <td align="left" colspan="4">
                            	<!--20130911 Modify by GoodTV Tanya:修改文宣品選項及增加「不主動聯絡-錯址」-->                            	
                            	  紙本月刊本數<input type="text" class="font9" name="IsSendNewsNum" size='3' maxlength='3' value='<%=RS("IsSendNewsNum")%>'>                            	
                              <input type="checkbox" name="IsSendEpaper" value="Y" <%If RS("IsSendEpaper")="Y" Then Response.Write "checked" End If%> >電子報                              
                            	<!--<input type="checkbox" name="IsSendNews" value="Y" <%If RS("IsSendNews")="Y" Then Response.Write "checked" End If%> >會訊                              
                              <input type="checkbox" name="IsSendYNews" value="Y" <%If RS("IsSendYNews")="Y" Then Response.Write "checked" End If%> >年報
                              <input type="checkbox" name="IsBirthday" value="Y" <%If RS("IsBirthday")="Y" Then Response.Write "checked" End If%> >生日卡
                              <input type="checkbox" name="IsXmas" value="Y" <%If RS("IsXmas")="Y" Then Response.Write "checked" End If%> >賀卡-->
                            </td>                            
                            <td align="left" colspan="3">
                            	<input type="checkbox" name="IsErrAddress" value="Y" <%If RS("IsErrAddress")="Y" Then Response.Write "checked" End If%>>不主動聯絡-錯址
                            </td>
                          </tr>
                          <!--#include file="../include/calendar2.asp"-->
                          <!--20130911Mark by GoodTV-Tanya-->
                          <!--<tr>
                            <td width="99%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                          </tr>
                          <tr>
                            <td align="right">收據開立：</td>
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
                              <input type="checkbox" name="IsAnonymous" value="Y" <%If RS("IsAnonymous")="Y" Then Response.Write "checked" End If%> >徵信錄匿名
                            	&nbsp;匿名：
                            	<input type="text" name="NickName" size="13" class="font9" maxlength="20" value="<%=RS("NickName")%>">
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
                          </tr>-->
                          <tr>
                            <!--<td align="right">收據地址：</td>-->
                            <td align="left" colspan="7">
                            	<table width="100%" border="0" cellpadding="0" cellspacing="0" >
                                <tr>
                                	<!--<%
                                	  IsAbroad_Invoice_Checked=""
                                	  donor_invoic_address="block"
                                	  If RS("IsAbroad_Invoice")="Y" Then 
                                	    IsAbroad_Invoice_Checked="checked"
                                	    donor_invoic_address="none"
                                	  End If
                                	%>-->
                                	<!--<td width="70"><input type="checkbox" name="IsAbroad_Invoice" value="Y" OnClick='IsAbroad_Invoice_OnClick()' <%=IsAbroad_Invoice_Checked%> >海外地址</td>-->
                            	    <!--<td width="175" id="donor_invoic_address" style="display:<%=donor_invoic_address%>">
                            	    <!--<%
                            	      call CodeCity2 ("form","Invoice_ZipCode",RS("Invoice_ZipCode"),"Invoice_City",RS("Invoice_City"),"Invoice_Area",RS("Invoice_Area"),"Y")
                            	    %>
                            	    </td>
                            	    <!--<td><input type='text' class='font9' name='Invoice_Address' size='38' maxlength='100' value='<%=RS("Invoice_Address")%>'></td>-->
                            	  </tr>
                            	</table>
                            </td>
                          </tr>
                          <!--<tr>
                            <td align="right">收據抬頭：</td>
                            <td align="left" colspan="3"><input type="text" name="Invoice_Title" size="46" class="font9" maxlength="100" value="<%=RS("Invoice_Title")%>"></td>
                            <td align="left" colspan="4">
                              收據身分證/統編：
                              <input type="text" name="Invoice_IDNo" size="15" class="font9" maxlength="10" onKeyUp="javascript:UCaseInvoiceIDNo();" value="<%=RS("Invoice_IDNo")%>">
                            </td>
                          </tr>
                          <%If Donate_IsFdc="Y" Then%>
                          <tr>
                            <td align="left" colspan="8">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="IsFdc" value="Y" <%If RS("IsFdc")="Y" Then Response.Write "checked" End If%> ><font color="#ff0000">本捐款人願意將捐款資料上傳至國稅局以供報稅用</font>&nbsp;(&nbsp;請填寫款人收據身分證/統編&nbsp;)</td>
                          </tr>
                          <%Else%>
                          <input type="hidden" name="IsFdc" value="<%=RS("IsFdc")%>">
                          <%End If%>-->
                          <tr>
                            <td align="right">備註：</td>
                            <td align="left" colspan="7">
                            	<textarea name="Remark" rows="6" cols="81" class="font9"><%=RS("Remark")%></textarea>
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                          </tr>
                          <tr>  
                            <td align="left" colspan="8">
                              <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                              	<!--20130911Mark by GoodTV-Tanya-->
                                <!--<tr>
                          	      <td align="right">首次捐款日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateD" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateD")%>"></td>
                                  <td align="right">最近捐款日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateD" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateD")%>"></td>
                                  <td align="right">累計捐款次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoD" size="13" class="font9t" readonly value="<%=RS("Donate_NoD")%>" style="text-align: right"></td>
                            	    <td align="right">累計捐款金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalD" size="13" class="font9t" readonly value="<%=RS("Donate_TotalD")%>" style="text-align: right"></td>
                                </tr>
                                <%If IsContribute="Y" Then%>
                                <tr>
                          	      <td align="right">首次捐物日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateC" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateC")%>"></td>
                                  <td align="right">最近捐物日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateC" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateC")%>"></td>
                                  <td align="right">累計捐物次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoC" size="13" class="font9t" readonly value="<%=RS("Donate_NoC")%>" style="text-align: right"></td>
                            	    <td align="right">累計折合現金：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalC" size="13" class="font9t" readonly value="<%=RS("Donate_TotalC")%>" style="text-align: right"></td>
                                </tr>
                                <%End If%>
                                <tr>
                          	      <td align="right">首次繳費日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateM" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateM")%>"></td>
                                  <td align="right">最近繳費日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateM" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateM")%>"></td>
                                  <td align="right">累計繳費次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoM" size="13" class="font9t" readonly value="<%=RS("Donate_NoM")%>" style="text-align: right"></td>
                            	    <td align="right">累計繳費金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalM" size="13" class="font9t" readonly value="<%=RS("Donate_TotalM")%>" style="text-align: right"></td>
                                </tr>
                                <tr>
                          	      <td align="right">首次報名日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateA" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateA")%>"></td>
                                  <td align="right">最近報名日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateA" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateA")%>"></td>
                                  <td align="right">累計報名次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoA" size="13" class="font9t" readonly value="<%=RS("Donate_NoA")%>" style="text-align: right"></td>
                            	    <td align="right">累計報名金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalA" size="13" class="font9t" readonly value="<%=RS("Donate_TotalA")%>" style="text-align: right"></td>
                                </tr>
                                <%If IsShopping="Y" Then%>
                                <tr>
                          	      <td align="right">首次購物日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateS" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateS")%>"></td>
                                  <td align="right">最近購物日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateS" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateS")%>"></td>
                                  <td align="right">累計購物次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoS" size="13" class="font9t" readonly value="<%=RS("Donate_NoS")%>" style="text-align: right"></td>
                            	    <td align="right">累計購物金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalS" size="13" class="font9t" readonly value="<%=RS("Donate_TotalS")%>" style="text-align: right"></td>
                                </tr>
                                <%End If%>
                                <tr>
                          	      <td align="right">首次非捐日期：</td>
                                  <td align="left"><input type="text" name="Begin_DonateDateND" size="13" class="font9t" readonly value="<%=RS("Begin_DonateDateND")%>"></td>
                                  <td align="right">最近非捐日期：</td>
                                  <td align="left"><input type="text" name="Last_DonateDateND" size="13" class="font9t" readonly value="<%=RS("Last_DonateDateND")%>"></td>
                                  <td align="right">累計非捐次數：</td>
                                  <td align="left"><input type="text" name="Donate_NoND" size="13" class="font9t" readonly value="<%=RS("Donate_NoND")%>" style="text-align: right"></td>
                            	    <td align="right">累計非捐金額：</td>
                            	    <td align="left"><input type="text" name="Donate_TotalND" size="13" class="font9t" readonly value="<%=RS("Donate_TotalND")%>" style="text-align: right"></td>
                                </tr>-->
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
                          <tr>
                            <td align="center" colspan="8">
                               <input type="button" value=" 修 改 " name="update" class="cbutton" style="cursor:hand" onClick="Update_OnClick()">
                               &nbsp;&nbsp;
                               <input type="button" value=" 刪 除 " name="delete" class="cbutton" style="cursor:hand" onClick="Delete_OnClick()">	
                               &nbsp;&nbsp;
                               <input type="button" value=" 取 消 " name="cancel" class="cbutton" style="cursor:hand" onClick="Cancel_OnClick()">
                               &nbsp;&nbsp;
                               <!--20130912Mark by GoodTV-Tanya-->
                               <!--<input type="button" value="新增本人繳費資料" name="Add" class="addbutton" style="cursor:hand" onClick="Add_OnClick()">	-->
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
                <td class="table63-bg">　</td>
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
//20130911Mark by GoodTV-Tanya
//function Donor_Category_OnChange(Category){
//  if(Category.indexOf('學生')>-1||Category.indexOf('個人')>-1){
//    donorsex.style.display='block';
//    donortitle.style.display='block';
//  }else{
//    donorsex.style.display='none';
//    donortitle.style.display='none';
//    document.form.Sex.value='';
//    document.form.Title.value='';
//  }
//}
function UCaseMemberNo(){
  document.form.Member_No.value=document.form.Member_No.value.toUpperCase();
}
function Issue_Type_OnClick(){
  if(document.form.Issue_Type.checked){
    document.form.Member_No.style.backgroundColor='#ffffff';
    document.form.Member_No.readOnly=false;
    document.form.Member_No.value='';
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
//20130911Mark by GoodTV-Tanya
//function UCaseInvoiceIDNo(){
//  document.form.Invoice_IDNo.value=document.form.Invoice_IDNo.value.toUpperCase();
//}
//function IsAbroad_OnClick(){
//  if(document.form.IsAbroad.checked){
//    donor_address.style.display='none';
//    document.form.ZipCode.value='';
//    document.form.City.value='';
//    ClearOption(document.form.Area);
//    document.form.Area.options[0] = new Option('鄉鎮市區','');
//  }else{
//    donor_address.style.display='block';
//  }
//}
//20130911Mark by GoodTV-Tanya
//function IsAbroad_Invoice_OnClick(){
//  if(document.form.IsAbroad_Invoice.checked){
//    donor_invoic_address.style.display='none';
//    document.form.Invoice_ZipCode.value='';
//    document.form.Invoice_City.value='';
//    ClearOption(document.form.Invoice_Area);
//    document.form.Invoice_Area.options[0] = new Option('鄉鎮市區','');
//  }else{
//    donor_invoic_address.style.display='block';
//  }
//}
//20130911Mark by GoodTV-Tanya
//function SameAddress_OnClick(){
//  if(document.form.IsSameAddress.checked){
//    if(document.form.IsAbroad.checked){
//      donor_invoic_address.style.display='none';
//      document.form.Invoice_ZipCode.value='';
//      document.form.Invoice_City.value='';
//      ClearOption(document.form.Invoice_Area);
//      document.form.Invoice_Area.options[0] = new Option('鄉鎮市區','');
//    }else{
//      donor_invoic_address.style.display='block';
//    }
//    document.form.IsAbroad_Invoice.checked=document.form.IsAbroad.checked;
//    document.form.Invoice_Title.value=document.form.Donor_Name.value;
//    document.form.Invoice_IDNo.value=document.form.IDNo.value;
//    document.form.Invoice_City.value=document.form.City.value;
//    ChgCity(document.form.Invoice_City.value,document.form.Invoice_Area,document.form.Invoice_ZipCode);  
//    document.form.Invoice_ZipCode.value=document.form.ZipCode.value;  
//    document.form.Invoice_Area.value=document.form.Area.value;
//    document.form.Invoice_Address.value=document.form.Address.value;
//    document.form.Title2.value=document.form.Title.value;
//  }
//}
function Update_OnClick(){	
  if(document.form.IsMember.value=='Y'){
    if(document.form.IsMemberNo.value!='Y'){
      <%call CheckStringJ("Member_No","會員編號")%>
      <%call ChecklenJ("Member_No",20,"會員編號")%>
    }
  }
  <%call CheckStringJ("Donor_Name","會員姓名")%>
  <%call ChecklenJ("Donor_Name",100,"會員姓名")%>
  //20130911Mark by GoodTV-Tanya
  <%'call ChecklenJ("Category",20,"類別")%>
  <%call ChecklenJ("Sex",2,"性別")%>
  <%call ChecklenJ("Title",20,"稱謂")%>
  <%call ChecklenCheckBoxJ("Donor_Type",100,"身份別")%>
  
  //20130911Mark by GoodTV-Tanya
  <%'call ChecklenJ("Member_Type",20,"會員別")%>  
  //if(document.form.Member_JoinDate.value!=''){
    <%'call CheckDateJ("Member_JoinDate","加入日期")%>
  //}
  <%call ChecklenJ("Member_Status",20,"狀態")%>
  //20130911Mark by GoodTV-Tanya
  <%'call ChecklenJ("Member_Group",20,"組別")%>
  <%'call ChecklenCheckBoxJ("Volunteer_Service",100,"志工項目")%>
  <%'call ChecklenJ("Volunteer_Other",100,"其他志工項目")%>
  
  <%call ChecklenJ("IDNo",10,"身分證/統編")%>
  if(document.form.Birthday.value!=''){
    <%call CheckDateJ("Birthday","出生日期")%>
  }
  <%'call ChecklenJ("Education",20,"教育程度")%>
  <%'call ChecklenJ("Occupation",20,"職業別")%>
  <%'call ChecklenJ("Marriage",20,"婚姻狀況")%>
  <%'call ChecklenJ("Religion",20,"宗教信仰")%>
  <%'call ChecklenJ("ReligionName",40,"道場/教會名稱")%>  
  <%call ChecklenJ("Cellular_Phone",40,"手機")%>
  <%call ChecklenJ("Tel_Office",40,"電話")%>
  <%'call ChecklenJ("Tel_Home",40,"電話(夜)")%>
  <%call ChecklenJ("Fax",40,"傳真")%>
  <%call ChecklenJ("Email",80,"E-Mail")%>
  <%call ChecklenJ("Contactor",100,"聯絡人")%>
  <%'call ChecklenJ("OrgName",40,"服務單位")%>
  <%'call ChecklenJ("JobTitle",40,"職稱")%>
  <!--20130912Modify by GoodTV-Tanya:修改地址格式同捐款人-->
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
  <%'call ChecklenJ("Address",100,"通訊地址")%>
  <!--20130911Mark by GoodTV-Tanya-->
  <%'call CheckStringJ("Invoice_Type","收據開立")%>
  <%'call ChecklenJ("Invoice_Type",20,"收據開立")%>
  //if(document.form.IsAnonymous.checked){
    <%'call CheckStringJ("NickName","匿名")%>
  //}
  <%'call ChecklenJ("NickName",20,"匿名")%>
  //if(document.form.Invoice_Type.value.indexOf('不')==-1){
    <%'call CheckStringJ("Invoice_Title","收據抬頭")%>
    //if(document.form.IsAbroad.checked==false){
      <%'call CheckStringJ("Invoice_City","收據地址縣市")%>
      <%'call CheckStringJ("Invoice_Area","收據地址鄉鎮市區")%>
    //}
    <%'call CheckStringJ("Invoice_Address","收據地址路段號")%>
  //}
  <%'call ChecklenJ("Invoice_Title",100,"收據抬頭")%>
  <%'call ChecklenJ("Invoice_IDNo",10,"收據身分證/統編")%>
  <%'call ChecklenJ("Invoice_Address",100,"收據地址路段號")%>
  //if(document.form.IsFdc.checked){
    <%'call CheckStringJ("Invoice_IDNo","收據身分證/統編")%>
  //}
  if(document.form.IsMemberNo.value=='自動編號'){
    document.form.Member_No.value='';
  }
  <%call SubmitJ("update")%>
}  
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
  location.href='member.asp';
}
//20130912Mark by GoodTV-Tanya
//function Add_OnClick(){
//  location.href='member_donate_add.asp?ctype='+document.form.ctype.value+'&donor_id='+document.form.donor_id.value+'';
//}
--></script>	