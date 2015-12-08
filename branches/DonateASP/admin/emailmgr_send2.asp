<!--#include file="../include/dbfunctionJ.asp"-->
<%
SQL="Select Donor.Donor_Id,Donor.Email From Donor Left Join (Select Donor_Id,Donate_No=Count(*),Donate_Total=Sum(Donate_Amt) From Donate Where 1=1 "
If Request("Dept_Id")<>"" Then SQL=SQL&"And Dept_Id = '"&Request("Dept_Id")&"' "
If Request("Donate_Date_Begin")<>"" Then SQL=SQL&"And Donate_Date>='"&Request("Donate_Date_Begin")&"' "
If Request("Donate_Date_End")<>"" Then SQL=SQL&"And Donate_Date<='"&Request("Donate_Date_End")&"' "
If Request("Act_Id")<>"" Then SQL=SQL&"And Act_Id='"&Request("Act_Id")&"' "
SQL=SQL&"Group By Donor_Id) As A On Donor.Donor_Id=A.Donor_Id " & _
        "Left Join CODECITY As B On DONOR.City=B.mCode Left Join CODECITY As C On DONOR.Area=C.mCode " & _
        "Left Join CODECITY As D On DONOR.Invoice_City=D.mCode Left Join CODECITY As E On DONOR.Invoice_Area=E.mCode Where Email<>'' "
If Request("Category")<>"" Then
  donorcategory=""
  Ary_Category=Split(Request("Category"),",")
  If UBound(Ary_Category)=0 Then
    donorcategory=donorcategory&"('"&Trim(Ary_Category(0))&"')"
  Else
    For I = 0 To UBound(Ary_Category)
      If I=0 Then
        donorcategory=donorcategory&"('"&Trim(Ary_Category(I))&"',"
      ElseIf I=UBound(Ary_Category) Then 
        donorcategory=donorcategory&"'"&Trim(Ary_Category(I))&"')" 
      Else
        donorcategory=donorcategory&"'"&Trim(Ary_Category(I))&"',"
      End If
    Next
  End If
  If donorcategory<>"" Then SQL=SQL&"And Category In "&donorcategory&" "
End If
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
If Request("Donor_Name")<>"" Then SQL_Where=SQL_Where&"And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Donor.Invoice_Title Like '%"&request("Donor_Name")&"%') "
If Request("Donor_Id_Begin")<>"" Then SQL=SQL&"And Donor_Id>='"&Request("Donor_Id_Begin")&"' "
If Request("Donor_Id_End")<>"" Then SQL=SQL&"And Donor_Id<='"&Request("Donor_Id_End")&"' "
If Request("Birthday_Month")<>"" Then SQL=SQL&"And Month(Birthday)='"&Request("Birthday_Month")&"' "
If Request("Donate_Over")<>"" Then SQL=SQL&"And Last_DonateDate<'"&DateAdd("M",Cint(Request("Donate_Over"))*-1,Date())&"' "
If Request("City")<>"" Then SQL=SQL&"And City='"&Request("City")&"' "
If Request("Area")<>"" Then SQL=SQL&"And Area='"&Request("Area")&"' "
If Request("Invoice_City")<>"" Then SQL=SQL&"And City='"&Request("Invoice_City")&"' "
If Request("Invoice_Area")<>"" Then SQL=SQL&"And Invoice_Area='"&Request("Invoice_Area")&"' "
If Request("IsAbroad")<>"" Then SQL=SQL&"And IsAbroad='Y' "
If Request("IsSendNews")<>"" Then SQL=SQL&"And IsSendNews='Y' "
If Request("IsSendEpaper")<>"" Then SQL=SQL&"And IsSendEpaper='Y' "
If Request("IsSendYNews")<>"" Then SQL=SQL&"And IsSendYNews='Y' "
If Request("Donate_Total")<>"" Then SQL=SQL&"And CONVERT(numeric,A.Donate_Total)>='"&Request("Donate_Total")&"' "
If Request("Donate_No")<>"" Then SQL=SQL&"And A.Donate_No>='"&Request("Donate_No")&"' "
SQL=SQL&"Order By Donor.Donor_Id Desc "
TolRow=0
Call QuerySQL(SQL,RS)
If Not RS.EOF Then TolRow=RS.Recordcount
Session("SQL")=SQL
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>郵件管理 【新增資料】</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="post" action="emailmgr_send2.asp">
      <input type="hidden" name="action">
      <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">	
      <table border="0" width="830" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
	        <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25" colspan="2"><font color="#FFFFFF"><b>&nbsp;&nbsp;郵件管理 </b> <font size="2">【新增資料】</font></font></td>
        </tr>
        <tr>
	        <td width="100%" class="font11" valign="top">
            <div align="center"><center>
              <table border="0" width="100%" cellspacing="0">
   		          <tr>
   		            <td>
                    <table width="100%" border="1" cellspacing="0" style="border-collapse: collapse" cellpadding="2" bordercolor="#E1DFCE">
                      <tr>
                        <td noWrap align="right"><font color="#000080">機構：</font></td>
                        <td colspan="3">
                        <%
                          If Session("comp_label")="1" Then
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where Comp_Label>='"&Session("comp_label")&"' Order By Comp_Label,Dept_Id"
                          Else
                            SQL="Select Dept_Id,Comp_ShortName From DEPT Where (Dept_Id='"&Session("dept_id")&"' Or Comp_Label>'"&Session("comp_label")&"') Order By Comp_Label,Dept_Id"
                          End If
                          FName="Dept_Id"
                          Listfield="Comp_ShortName"
                          menusize="1"
                          BoundColumn=Session("dept_id")
                          call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"Y")
                        %>
                        </td>
                      </tr>
                      <tr>
                        <td noWrap align="right"><font color="#000080">類別：</font></td>
                        <td colspan="3">
                        <%
                          SQL="Select Category=CodeDesc From CASECODE Where CodeType='Category' Order By Seq"
                          FName="Category"
                          Listfield="Category"
                          BoundColumn=""
                          call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                        %>
                        </td>
                      </tr>
                      <tr>
                        <td noWrap align="right"><font color="#000080">身份別：</font></td>
                        <td colspan="3">
                        <%
                          SQL="Select Donor_Type=CodeDesc From CASECODE Where CodeType='DonorType' Order By Seq"
                          FName="Donor_Type"
                          Listfield="Donor_Type"
                          BoundColumn=""
                          call CheckBoxList (SQL,FName,Listfield,BoundColumn)
                        %>
                        </td>
                      </tr>
                      <tr>
                        <td noWrap align="right" width="12%"><font color="#000080">通訊地址：</font></td>
                        <td width="25%"><%call CodeArea ("form","City","","Area","","Y")%></td>
                        <td noWrap align="right" width="12%"><font color="#000080">收據地址：</font></td>
                        <td width="51%"><%call CodeArea ("form","Invoice_City","","Invoice_Area","","N")%></td>
                      </tr>
                      <tr>
                        <td noWrap align="right"><font color="#000080">捐款日期：</font></td>
                        <td><%call Calendar("Donate_Date_Begin","")%> ~ <%call Calendar("Donate_Date_End","")%></td>
                        <td noWrap align="right"><font color="#000080">捐款總金額：</font></td>
                        <td><input type="text" name="Donate_Total" size="9" class="font9">&nbsp;元以上</td>
                      </tr>
                      <tr>
                        <td noWrap align="right"><font color="#000080">募款活動：</font></td>
                        <td>
                        <%
                          SQL="Select Act_Id,Act_ShortName From ACT Order By Act_id Desc"
                          FName="Act_Id"
                          Listfield="Act_ShortName"
                          menusize="1"
                          BoundColumn=""
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>
                        </td>
                        <td noWrap align="right"><font color="#000080">捐款總次數：</font></td>
                        <td><input type="text" name="Donate_No" size="9" class="font9">&nbsp;次以上</td>
                      </tr>
                      <tr>
                        <td noWrap align="right"><font color="#000080">捐款人姓名：</font></td>
                        <td><input type="text" name="Donor_Name" size="19" class="font9"></td>
                        <td noWrap align="right"><font color="#000080">捐款人編號：</font></td>
                        <td><input type="text" name="Donor_Id_Begin" size="9" class="font9"> ~ <input type="text" name="Donor_Id_End" size="9" class="font9"></td>
                      </tr>
                      <tr>
                        <td noWrap align="right"><font color="#000080">生日月份：</font></td>
                        <td>
                        <%
                          Response.Write "<SELECT Name='Birthday_Month' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                          Response.Write "<OPTION> </OPTION>"
                          For I=1 To 12
                            Response.Write "<OPTION value='"&I&"'>"&I&"月</OPTION>"
                          Next
                          Response.Write "</SELECT>"
                        %>	
                        </td>
                        <td noWrap align="right"><font color="#000080">多久未捐款：</font></td>
                        <td>
                        <%
                          Response.Write "<SELECT Name='Donate_Over' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                          Response.Write "<OPTION> </OPTION>"
                          Response.Write "<OPTION value='3'>3個月</OPTION>"
                          Response.Write "<OPTION value='6'>6個月</OPTION>"
                          Response.Write "<OPTION value='9'>9個月</OPTION>"
                          Response.Write "<OPTION value='12'>一年</OPTION>"
                          Response.Write "<OPTION value='24'>二年</OPTION>"
                          Response.Write "<OPTION value='36'>三年以上</OPTION>"
                          Response.Write "</SELECT>"
                        %>
                        </td>
                      </tr>
                      <tr>
                        <td noWrap align="right"><font color="#000080">其他：</font></td>
                        <td colspan="3">
                          <input type="checkbox" name="IsAbroad" value="Y">海外地址
                          <input type="checkbox" name="IsSendNews" value="Y">會訊
                          <input type="checkbox" name="IsSendEpaper" value="Y">電子報
                          <input type="checkbox" name="IsSendYNews" value="Y">年報
                        </td>
                      </tr>
                      <tr> 
                        <td align="right" height=2 class="td2"></td>
                        <td height=2 class="td2" colspan="3"></td>
                      </tr>
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 class="font9" colspan="4">
                        <font color="#ff0000">請篩選您要寄發的對象&nbsp;</font>
                        <%
                          Response.Write "<button id='next' style='position:relative;left:0;width:75;height:25;font-size:9pt;cursor:hand' class='button3-bg' onClick='Next_OnClick()'> <img src='../images/xmenu02.gif' width='20' height='20' align='absmiddle'> 下一步 </button>&nbsp;"
                          Response.Write "<button id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt;cursor:hand' class='button3-bg' onClick='Cancel_OnClick()'> <img src='../images/icon6.gIf' width='20' height='15' align='absmiddle'> 離開 </button>&nbsp;"
                        %>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </center></div>
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
function Next_OnClick(){
  <%call CheckStringJ("EmailMgr_Subject","訊息標題")%>
  <%call ChecklenJ("EmailMgr_Subject",100,"訊息標題")%>
  <%call SubmitJ("save")%>
}
function Cancel_OnClick(){
  location.href='emailmgr_edit.asp?ser_no='+document.form.ser_no.value+'';
}
--></script>