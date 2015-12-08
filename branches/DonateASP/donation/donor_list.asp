<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="donor"%>
<%
Function GridList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  If Not RS1.EOF Then 
    FieldsCount = RS1.Fields.Count-1
    totRec=RS1.Recordcount
    If totRec>0 Then 
      RS1.PageSize=PageSize
      If nowPage="" or nowPage=0 Then
        nowPage=1
      ElseIf cint(nowPage) > RS1.PageCount Then 
        nowPage=RS1.PageCount 
      End If
      session("nowPage")=nowPage
      RS1.AbsolutePage=nowPage
      totPage=RS1.PageCount
      SQL=server.URLEncode(SQL)
    End If
    Response.Write "<Form action='' id=form1 name=form1>" 
    Response.Write "<table border=0 cellspacing='0' cellpadding='1' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='30%'></td>"
    Response.Write "<td width='55%'><span style='font-size: 9pt; font-family: 新細明體'>第" & nowPage & "/" & totPage & "頁&nbsp;&nbsp;</span>"
    Response.Write "<span style='font-size: 9pt; font-family: 新細明體'>共" & FormatNumber(totRec,0) & "筆&nbsp;&nbsp;</span>"
    If cint(nowPage) <>1 Then
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage-1)&"&SQL="&SQL&"'>上一頁</a>" 
    End If
    If cint(nowPage)<>RS1.PageCount and cint(nowPage)<RS1.PageCount Then 
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage+1)&"&SQL="&SQL&"'>下一頁</a>" 
    End If
    Response.Write " |&nbsp;<span style='font-size: 9pt; font-family: 新細明體'> 跳至第<select name=GoPage size='1' style='font-size: 9pt; font-family: 新細明體' onchange='GoPage_OnChange(this.value)'>"
    For iPage=1 to totPage
      If iPage=cint(nowPage) Then
        strSelected = "selected"
      Else
	      strSelected = ""
      End If
      Response.Write "<option value='"&iPage&"'" & strSelected & ">" & iPage & "</option>"
    Next
    Response.Write "</select>頁</span></td>" 
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
    End If   
    Response.Write "</tr></table>"
    Dim I
    Dim J
    Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
    Response.Write "<tr>"
    For J = 1 To FieldsCount
      Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J).Name & "</span></font></td>"
    Next
    Response.Write "</tr>"
    I = 1
    While Not RS1.EOF And I <= RS1.PageSize
      If Hlink<>"" Then
        Response.Write "<a href='" & HLink & RS1(LinkParam) & "' target='" & LinkTarget &"'>"
        showhand="style='cursor:hand'"
      End If
      Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
      For J = 1 To FieldsCount
	  	if J=11 Then		
			if isnull(RS1(J))=false And RS1(J)<>"" Then
				Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&FormatNumber(RS1(J),0)&"</span></td>"
			Else
				Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>0</span></td>"
			End If
		Else
	        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&Data_Minus(RS1(J))&"</span></td>"
		End If
      Next
      I = I + 1
      RS1.MoveNext
      Response.Write "</tr>"
      Response.Write "</a>"
    Wend
    Response.Write "</table>"
    Response.Write "</form>"
  Else
    Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='85%' align='center' style='color:#ff0000'>** 沒有符合條件的資料 **</td>"
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
    End If
    Response.Write "</table>"
  End If
  RS1.Close
  Set RS1=Nothing
End Function
%>
<!--#include file="../include/head.asp"-->
<body class=tool>
<%
'20130913Modify by GoodTV-Tanya:Add IsMember<>'Y'只顯示天使的資料
'20131008Modify by GoodTV Tanya:Add 查詢「舊編號」
'20131105Modify by GoodTV Tanya:1.可查詢讀者，並增加「讀者」註記 2.修改文宣品查詢 3.增加「雷同資料查詢」
'20131111Modify by GoodTV Tanya:page_load不帶全部資料
If request("action") = "" and request("SQL")="" Then
  Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
  Response.Write "  <tr>"
  Response.Write "    <td width='100%' align='center' style='color:#ff0000'>** 請先輸入查詢條件 **</td>"	  
  Response.Write "  </tr>"
  Response.Write "</table>"  
ElseIf request("SQL")="" Then 
	'20131112 Modify by GoodTV Tanya:修正海外通訊地址無法顯示問題&增加海外地址查詢欄位	
	'20131212 Modify by GoodTV Tanya:修正身份證/統編查詢為模糊查詢	
	'20140106 Modify by GoodTV Tanya:修改捐款累計紀錄查詢結果-首捐日=Begin_DonateDate,累積次數=Donate_No,累計捐款金額=Donate_Total
	SQL="Select Distinct Donor_Id,編號=(Case When Member_No<>'' Then Member_No Else CONVERT(nvarchar(20),Donor_Id) End),捐款人=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),IDNo AS '身份證/統編',讀者=(Case When IsMember='Y' Then 'V' ELse '' End),身份別=Donor_Type,聯絡電話=(Case When Tel_Office<>'' Then Tel_Office Else Tel_Home End),手機號碼=Cellular_Phone,通訊地址=(Case When ISNULL(DONOR.City,'')='' Then Address Else Case When A.mValue<>B.mValue Then ISNULL(DONOR.ZipCode,'')+A.mValue+B.mValue+Address Else ISNULL(DONOR.ZipCode,'')+A.mValue+Address End End),首捐日=CONVERT(VarChar,Begin_DonateDate,111),末捐日=CONVERT(VarChar,Last_DonateDate,111),累積次數=Donate_No,累計捐款金額=Donate_Total,資料建檔人員=Create_User,舊編號=IsNull(Donor_Id_Old,'') " & _
	      "From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode "
	      
	If request("action") = "fuzzy_query" Then		
		Dim SQLWhere
		If request("Donor_Name")<>"" Then SQLWhere=SQLWhere & "Or (Donor_Name Like N'%"&request("Donor_Name")&"%' Or NickName Like N'%"&request("Donor_Name")&"%' Or Contactor Like N'%"&request("Donor_Name")&"%' ) "
	  If request("Member_No")<>"" Then SQLWhere=SQLWhere & "Or Donor_Id Like '%"&request("Member_No")&"%' "
	  If request("Donor_Id_Old")<>"" Then SQLWhere=SQLWhere & "Or Donor_Id_Old Like '%"&request("Donor_Id_Old")&"%' "  	
	  If request("Tel_Office")<>"" Then SQLWhere=SQLWhere & "Or (Tel_Office Like '%"&request("Tel_Office")&"%' Or Tel_Home Like '%"&request("Tel_Office")&"%' Or Cellular_Phone Like '%"&request("Tel_Office")&"%') "
	  If request("Category")<>"" Then SQLWhere=SQLWhere & "Or Category = '"&request("Category")&"' "
	  If request("Donor_Type")<>"" Then SQLWhere=SQLWhere & "Or Donor_Type Like '%"&request("Donor_Type")&"%' "
	  If request("IDNo")<>"" Then SQLWhere=SQLWhere & "Or IDNo Like N'%"&request("IDNo")&"%' "
	  If request("Invoice_Title")<>"" Then SQLWhere=SQLWhere & "Or Invoice_Title Like '%"&request("Invoice_Title")&"%' "
	  If request("IsAbroad")="" And request("City")<>"" Then SQLWhere=SQLWhere & "Or (City = '"&request("City")&"' Or Invoice_City = '"&request("City")&"') "
	  If request("IsAbroad")="" And request("Area")<>"" Then SQLWhere=SQLWhere & "Or (Area = '"&request("Area")&"' Or Invoice_Area = '"&request("Area")&"') "
	  If request("IsAbroad")="" And request("Address")<>"" Then SQLWhere=SQLWhere & "Or (Address Like N'%"&request("Address")&"%' Or Invoice_Address Like N'%"&request("Address")&"%') "
	  If request("IsAbroad")="" And request("Street")<>"" Then SQLWhere=SQLWhere & "Or (Street='"&request("Street")&"' Or Invoice_Street='"&request("Street")&"') " 
	  If request("IsAbroad")="" And request("SectionType")<>"" Then SQLWhere=SQLWhere & "Or (Section='"&request("SectionType")&"' Or Invoice_Section='"&request("SectionType")&"') " 
	  If request("IsAbroad")="" And request("Lane")<>"" Then SQLWhere=SQLWhere & "Or (Lane='"&request("Lane")&"' Or Invoice_Lane='"&request("Lane")&"') " 
	  If request("IsAbroad")="" And request("Alley")<>"" Then SQLWhere=SQLWhere & "Or (Alley='"&request("Alley")&"' Or Invoice_Alley='"&request("Alley")&"') " 
	  If request("IsAbroad")="" And request("HouseNo")<>"" Then SQLWhere=SQLWhere & "Or (HouseNo='"&request("HouseNo")&"' Or Invoice_HouseNo='"&request("HouseNo")&"') "
	  If request("IsAbroad")="" And request("HouseNoSub")<>"" Then SQLWhere=SQLWhere & "Or (HouseNoSub='"&request("HouseNoSub")&"' Or Invoice_HouseNoSub='"&request("HouseNoSub")&"') "
	  If request("IsAbroad")="" And request("Floor")<>"" Then SQLWhere=SQLWhere & "Or (Floor='"&request("Floor")&"' Or Invoice_Floor='"&request("Floor")&"') "
	  If request("IsAbroad")="" And request("FloorSub")<>"" Then SQLWhere=SQLWhere & "Or (FloorSub='"&request("FloorSub")&"' Or Invoice_FloorSub='"&request("FloorSub")&"') "
	  If request("IsAbroad")="" And request("Room")<>"" Then SQLWhere=SQLWhere & "Or (Room='"&request("Room")&"' Or Invoice_Room='"&request("Room")&"') "
	  If request("IsAbroad")<>"" Then SQLWhere=SQLWhere & "Or (IsAbroad = 'Y' Or IsAbroad_Invoice = 'Y') "
    If request("IsAbroad")<>"" And request("IsAbroad_Address")<>"" Then SQLWhere=SQLWhere & "Or Address LIKE N'%"&request("IsAbroad_Address")&"%'"	 		  	
'	  If request("IsSendNews")<>"" Then SQLWhere=SQLWhere & "Or IsSendNews = 'Y' "
    If request("IsSendNewsNum")<>"" Then SQLWhere=SQLWhere & "Or (IsSendNews = 'Y' And IsSendNewsNum='"&request("IsSendNewsNum")&"')"	  	
	  If request("IsSendEpaper")<>"" Then SQLWhere=SQLWhere & "Or IsSendEpaper = 'Y' "
	  If request("IsDVD")<>"" Then SQLWhere=SQLWhere & "Or IsDVD = 'Y' "
'	  If request("IsSendYNews")<>"" Then SQLWhere=SQLWhere & "Or IsSendYNews = 'Y' "
'	  If request("IsBirthday")<>"" Then SQLWhere=SQLWhere & "Or IsBirthday = 'Y' "
'	  If request("IsXmas")<>"" Then SQLWhere=SQLWhere & "Or IsXmas = 'Y' "
		
		If SQLWhere <> "" Then
			SQL = SQL & " Where " & Mid(SQLWhere,4)
		End If
	Else	  
		SQL=SQL & " Where 1=1 "
		'20140122 Modify by GoodTV Tanya:捐款人編號改為精準查詢
        '20140620 捐款人舊編號改為精準查詢
	  If request("Donor_Name")<>"" Then SQL=SQL & "And (Donor_Name Like N'%"&request("Donor_Name")&"%' Or NickName Like N'%"&request("Donor_Name")&"%' Or Contactor Like N'%"&request("Donor_Name")&"%' ) "
	  If request("Member_No")<>"" Then SQL=SQL & "And Donor_Id = "&request("Member_No")&" "
	  If request("Donor_Id_Old")<>"" Then SQL=SQL & "And Donor_Id_Old = '"&request("Donor_Id_Old")&"' "  	
	  If request("Tel_Office")<>"" Then SQL=SQL & "And (Tel_Office Like '%"&request("Tel_Office")&"%' Or Tel_Home Like '%"&request("Tel_Office")&"%' Or Cellular_Phone Like '%"&request("Tel_Office")&"%') "
	  If request("Category")<>"" Then SQL=SQL & "And Category = '"&request("Category")&"' "
	  If request("Donor_Type")<>"" Then SQL=SQL & "And Donor_Type Like '%"&request("Donor_Type")&"%' "
	  If request("IDNo")<>"" Then SQL=SQL & "And IDNo Like N'%"&request("IDNo")&"%' "
	  If request("Invoice_Title")<>"" Then SQL=SQL & "And Invoice_Title Like '%"&request("Invoice_Title")&"%' "
	  If request("IsAbroad")="" And request("City")<>"" Then SQL=SQL & "And (City = '"&request("City")&"' Or Invoice_City = '"&request("City")&"') "
	  If request("IsAbroad")="" And request("Area")<>"" Then SQL=SQL & "And (Area = '"&request("Area")&"' Or Invoice_Area = '"&request("Area")&"') "
	  If request("IsAbroad")="" And request("Address")<>"" Then SQL=SQL & "And (Address Like N'%"&request("Address")&"%' Or Invoice_Address Like N'%"&request("Address")&"%') "
	  If request("IsAbroad")="" And request("Street")<>"" Then SQL=SQL & "And (Street='"&request("Street")&"' Or Invoice_Street='"&request("Street")&"') " 
	  If request("IsAbroad")="" And request("SectionType")<>"" Then SQL=SQL & "And (Section='"&request("SectionType")&"' Or Invoice_Section='"&request("SectionType")&"') " 
	  If request("IsAbroad")="" And request("Lane")<>"" Then SQL=SQL & "And (Lane='"&request("Lane")&"' Or Invoice_Lane='"&request("Lane")&"') " 
	  If request("IsAbroad")="" And request("Alley")<>"" Then SQL=SQL & "And (Alley='"&request("Alley")&"' Or Invoice_Alley='"&request("Alley")&"') " 
	  If request("IsAbroad")="" And request("HouseNo")<>"" Then SQL=SQL & "And (HouseNo='"&request("HouseNo")&"' Or Invoice_HouseNo='"&request("HouseNo")&"') "
	  If request("IsAbroad")="" And request("HouseNoSub")<>"" Then SQL=SQL & "And (HouseNoSub='"&request("HouseNoSub")&"' Or Invoice_HouseNoSub='"&request("HouseNoSub")&"') "
	  If request("IsAbroad")="" And request("Floor")<>"" Then SQL=SQL & "And (Floor='"&request("Floor")&"' Or Invoice_Floor='"&request("Floor")&"') "
	  If request("IsAbroad")="" And request("FloorSub")<>"" Then SQL=SQL & "And (FloorSub='"&request("FloorSub")&"' Or Invoice_FloorSub='"&request("FloorSub")&"') "
	  If request("IsAbroad")="" And request("Room")<>"" Then SQL=SQL & "And (Room='"&request("Room")&"' Or Invoice_Room='"&request("Room")&"') "
	  If request("IsAbroad")<>"" Then SQL=SQL & "And (IsAbroad = 'Y' Or IsAbroad_Invoice = 'Y') "
	  If request("IsAbroad")<>"" And request("IsAbroad_Address")<>"" Then SQL=SQL & "And Address Like N'%"&request("IsAbroad_Address")&"%'"
'	  If request("IsSendNews")<>"" Then SQL=SQL & "And IsSendNews = 'Y' "
		If request("IsSendNewsNum")<>"" Then SQL=SQL & "And (IsSendNews = 'Y' And IsSendNewsNum='"&request("IsSendNewsNum")&"')"
	  If request("IsSendEpaper")<>"" Then SQL=SQL & "And IsSendEpaper = 'Y' "
	  If request("IsDVD")<>"" Then SQL=SQL & "Or IsDVD = 'Y' "
'	  If request("IsSendYNews")<>"" Then SQL=SQL & "And IsSendYNews = 'Y' "
'	  If request("IsBirthday")<>"" Then SQL=SQL & "And IsBirthday = 'Y' "
'	  If request("IsXmas")<>"" Then SQL=SQL & "And IsXmas = 'Y' "
	End If
	  SQL=SQL&"Order By Donor_Id Desc"
	  'response.write SQL
	  'response.end()
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="donor_list"
HLink="donor_edit.asp?donor_id="
LinkParam="donor_id"
LinkTarget="main"
AddLink="donor_add.asp?donor_name2="&request("Donor_Name")&"&tel_office2="&request("Tel_Office")&"&category2="&request("Category")&"&donor_type2="&request("Donor_Type")&"&idno2="&request("IDNo")&"&city2="&request("City")&"&area2="&request("Area")&"&address2="&request("Address")&"&isabroad2="&request("IsAbroad")&"&issendnews2="&request("IsSendNews")&"&issendepaper2="&request("IsSendEpaper")&"&issendynews2="&request("IsSendYNews")&""
If request("action")="stop" Then
  call GridList_S (AddLink)
Else
  If request("action")="report" Or request("action")="export" Or request("action")="mobile" Or request("action")="email" Then Server.Transfer "donor_rpt.asp"
	If request("action")<>"" Or SQL<>"" Then call GridList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
End If  
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->