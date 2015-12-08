<!--#include file="../include/dbfunctionJ.asp"-->
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>讀者名冊</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
    <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
      <tr>
        <td>
          <table width="800" border="0" cellspacing="0" cellpadding="0" align="center" style="background-color:#EEEEE3">
            <tr>
              <td width="5%"></td>
              <td width="95%">
  		          <table width="40%" border="0" cellspacing="0" cellpadding="0">
		              <tr>
                    <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                    <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                    <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                  </tr>
                  <tr>
                    <td class="table62-bg">　</td>
                    <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">讀者名冊</td>
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
	        <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
            <tr>
              <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
              <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
              <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
            </tr>
            <tr>
              <td class="table62-bg">　</td>
              <td valign="top">
                <table width="100%"  border="0" cellspacing="0" style="border-collapse: collapse">
                  <tr>
                    <td class="td02-c" width="100%">
                      <table border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                        <tr>
                          <td>
                  	      <%
                            If request("SQL")="" Then
   						                Join_Type="Left Join"
'   						                If Request("Donate_Date_Begin")<>"" Or Request("Donate_Date_End")<>"" Or Request("Donate_Total_Begin")<>"" Or Request("Donate_Total_End")<>"" Or Request("Act_Id")<>"" Or Request("Donate_No_Begin")<>"" Or Request("Donate_No_End")<>"" Then Join_Type="Join"
															'20131002 Modify by GoodTV Tanya:拿掉「類別」、「會員別」、「加入日期」、「組別」、「收據地址」、「捐款資料」
   						                SQL="Select Distinct Donor.Donor_Id,編號=Donor.Donor_Id,讀者姓名=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),狀態=Member_Status,性別=Sex,身份別=Donor_Type,聯絡電話日=Tel_Office,手機=Cellular_Phone,聯絡電話夜=Tel_Home,電子信箱=Email,通訊地址=(Case When DONOR.City='' Then Address Else Case When B.mValue<>C.mValue Then B.mValue+DONOR.ZipCode+C.mValue+Address Else B.mValue+DONOR.ZipCode+Address End End) " & _
                                  "From Donor "
'                              If Request("Dept_Id")<>"" Then
'                                SQL=SQL&"Where Dept_Id = '"&Request("Dept_Id")&"' And Issue_Type<>'D' "
'                              Else
'                                SQL=SQL&"Where Dept_Id In ("&Session("all_dept_type")&") And Issue_Type<>'D' "
'                              End If
'                              If Request("Donate_Date_Begin")<>"" Then SQL=SQL&"And Donate_Date>='"&Request("Donate_Date_Begin")&"' "
'                              If Request("Donate_Date_End")<>"" Then SQL=SQL&"And Donate_Date<='"&Request("Donate_Date_End")&"' "
'                              If Request("Act_Id")<>"" Then SQL=SQL&"And Act_Id='"&Request("Act_Id")&"' "
                              SQL=SQL&"Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode " & _
                                      "Where IsMember='Y' "
                              If Request("Donor_Name")<>"" Then SQL=SQL&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Donor.Invoice_Title Like '%"&request("Donor_Name")&"%') "
                              '20131002 Modify by Tanya:改查詢編號Donor_Id
                              If Request("Member_No_Begin")<>"" Then SQL=SQL&"And Donor.Donor_Id>='"&Request("Member_No_Begin")&"' "
                              If Request("Member_No_End")<>"" Then SQL=SQL&"And Donor.Donor_Id<='"&Request("Member_No_End")&"' "
                              If Request("Birthday_Month")<>"" Then SQL=SQL&"And Month(Birthday)='"&Request("Birthday_Month")&"' "
'                              If Request("Donate_Over")<>"" Then SQL=SQL&"And Last_DonateDate<'"&DateAdd("M",Cint(Request("Donate_Over"))*-1,Date())&"' "
                              '20131002 Mark by GoodTV Tanya
'                              If Request("Category")<>"" Then
'                                donorcategory=""
'                                Ary_Category=Split(Request("Category"),",")
'                                If UBound(Ary_Category)=0 Then
'                                  donorcategory=donorcategory&"('"&Trim(Ary_Category(0))&"')"
'                                Else
'                                  For I = 0 To UBound(Ary_Category)
'                                    If I=0 Then
'                                      donorcategory=donorcategory&"('"&Trim(Ary_Category(I))&"',"
'                                    ElseIf I=UBound(Ary_Category) Then 
'                                      donorcategory=donorcategory&"'"&Trim(Ary_Category(I))&"')" 
'                                    Else
'                                      donorcategory=donorcategory&"'"&Trim(Ary_Category(I))&"',"
'                                    End If
'                                  Next
'                                End If
'                                If donorcategory<>"" Then SQL=SQL&"And Category In "&donorcategory&" "
'                              End If
                              If Request("Donor_Type")<>"" Then
                                donortype=""
                                Ary_Donor_Type=Split(Request("Donor_Type"),",")
                                If UBound(Ary_Donor_Type)=0 Then
                                  donortype=donortype&"And Donor_Type Like '%"&Trim(Ary_Donor_Type(0))&"%' "
                                Else
                                  For I = 0 To UBound(Ary_Donor_Type)
                                    If I=0 Then
                                      donortype=donortype&"And (Donor_Type Like '%"&Trim(Ary_Donor_Type(I))&"%' "
                                    ElseIf I=UBound(Ary_Donor_Type) Then 
                                      donortype=donortype&"Or Donor_Type Like '%"&Trim(Ary_Donor_Type(I))&"%') "
                                    Else
                                      donortype=donortype&"Or Donor_Type Like '%"&Trim(Ary_Donor_Type(I))&"%' "
                                    End If
                                  Next
                                End If
                                If donortype<>"" Then SQL=SQL&donortype
                              End If
                              
                              '20131002 Mark by GoodTV Tanya
'                              If Request("Member_Type")<>"" Then SQL=SQL&"And Member_Type='"&Request("Member_Type")&"' "
                              If Request("Member_Status")<>"" Then SQL=SQL&"And Member_Status='"&Request("Member_Status")&"' "
'                              If Request("Member_JoinYear")<>"" Then SQL=SQL&"And Year(Member_JoinDate)='"&Request("Member_JoinYear")&"' "
'                              If Request("Member_Group")<>"" Then SQL=SQL&"And Member_Group Like '%"&Request("Member_Group")&"%' "
                              If Request("City")<>"" Then SQL=SQL&"And City='"&Request("City")&"' "
                              If Request("Area")<>"" Then SQL=SQL&"And Area='"&Request("Area")&"' "
'                              If Request("Invoice_City")<>"" Then SQL=SQL&"And City='"&Request("Invoice_City")&"' "
'                              If Request("Invoice_Area")<>"" Then SQL=SQL&"And Invoice_Area='"&Request("Invoice_Area")&"' "
                              If Request("IsAbroad")<>"" Then SQL=SQL&"And (IsAbroad = 'Y' Or IsAbroad_Invoice = 'Y') "
                              If Request("IsSendNews")<>"" Then SQL=SQL&"And IsSendNews='Y' "
                              If Request("IsSendEpaper")<>"" Then SQL=SQL&"And IsSendEpaper='Y' "
'                              If Request("IsSendYNews")<>"" Then SQL=SQL&"And IsSendYNews='Y' "
'                              If Request("IsBirthday")<>"" Then SQL=SQL&"And IsBirthday='Y' "
'                              If Request("IsXmas")<>"" Then SQL=SQL&"And IsXmas='Y' "
'                              If Request("Donate_Total_Begin")<>"" Then SQL=SQL&"And CONVERT(int,A.Donate_Total)>='"&Request("Donate_Total_Begin")&"' "
'                              If Request("Donate_Total_End")<>"" Then SQL=SQL&"And CONVERT(int,A.Donate_Total)<='"&Request("Donate_Total_End")&"' "
'                              If Request("Donate_No_Begin")<>"" Then SQL=SQL&"And A.Donate_No>='"&Request("Donate_No_Begin")&"' "
'                              If Request("Donate_No_End")<>"" Then SQL=SQL&"And A.Donate_No<='"&Request("Donate_No_End")&"' "
                              SQL=SQL&"Order By Donor.Donor_Id Desc "
                            Else
                              SQL=request("SQL")
                            End If
                            Session("SQL")=SQL
                            '20131002 Mark by GoodTV Tanya
'                            Session("Donate_Date_Begin")=Request("Donate_Date_Begin")
'                            Session("Donate_Date_End")=Request("Donate_Date_End")
'                            Session("Act_Id")=Request("Act_Id")
                            PageSize=20
                            If request("nowPage")="" Then
                              nowPage=1
                            Else
                              nowPage=request("nowPage")
                            End If
                            ProgID="member_name_qry"
                            HLink=""
                            LinkParam="Donor_Id"
                            LinkTarget="main"
                            AddLink="member_name_list.asp"
                            Session("SQL")=SQL
                            If request("action")="report" Then
					                    Response.Redirect "../include/movebar.asp?URL=../member/member_name_rpt.asp"
					                  ElseIf request("action")="export" then
					                    Response.Redirect "member_name_rpt.asp?action=export"
					                  End If 
                            call GridList_Q (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)  
                  	      %>
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
  </center></div></p>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->
