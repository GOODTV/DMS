<!--#include file="../include/dbfunctionJ.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>物品查詢</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=tool>
  <form name="form" method="post" action="">
    <table width="100%" border="0" cellpadding="1" cellspacing="1">
      <tr>
        <td class="td02-c" width="500">
        	&nbsp;
        	捐款人：
        	<input type="text" name="Donor_Name" size="13" class="font9" value="<%=request("Donor_Name")%>" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
          &nbsp;
          會員編號：
		      <input type="text" name="Member_No" size="13" class="font9" value="<%=request("Member_No")%>" onkeyup="javascript:UCaseMemberNo();" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>
		      &nbsp;
		      身份別：
          <%
            SQL="Select Donor_Type=CodeDesc From CASECODE Where codetype='DonorType' Order By Seq"
            FName="Donor_Type"
            Listfield="Donor_Type"
            menusize="1"
            BoundColumn=request("Donor_Type")
            call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
          %>
          &nbsp;
          <input type="button" value=" 查 詢 " name="query" class="cbutton" style="cursor:hand" onClick="Query_OnClick()">
        </td>
      </tr>
    </table><br />
<%
  If request("Donor_Name")<>"" Or request("Member_No")<>"" Or request("Donor_Type")<>"" Then
    SQL1="Select Distinct Donor_Id,編號=(Case When Member_No<>'' Then Member_No Else CONVERT(nvarchar(20),Donor_Id) End),捐款人=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),類別=Category,身份別=Donor_Type,聯絡電話=(Case When Cellular_Phone<>'' Then Cellular_Phone Else Case When Tel_Office<>'' Then Tel_Office Else Tel_Home End End),通訊地址=(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) " & _
         "From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode Where 1=1 "    
    If request("Donor_Id")<>"" Then SQL1=SQL1 & "And Donor_Id <> '"&request("Donor_Id")&"' "
    If request("Donor_Name")<>"" Then SQL1=SQL1 & "And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Invoice_Title Like '%"&request("Donor_Name")&"%') "
    If request("Member_No")<>"" Then SQL1=SQL1 & "And (Member_No Like '%"&request("Member_No")&"%') "
    If request("Donor_Type")<>"" Then SQL1=SQL1 & "And Donor_Type Like '%"&request("Donor_Type")&"%' "
    SQL1=SQL1&"Order By (Case When Member_No<>'' Then Member_No Else CONVERT(nvarchar(20),Donor_Id) End) Desc"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    If Not RS1.EOF Then
      Response.Write "<table border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
      For I = 1 To RS1.Fields.Count-1
        Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
      Next
      Response.Write "</tr>"	
      While Not RS1.EOF
	      Response.Write "<a href='javascript:opener.document.form."&request("LinkId")&".value="""&RS1(0)&""";opener.document.form."&request("LinkName")&".value="""&RS1(2)&""";self.close()'><tr style='cursor:hand' bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
	      For I = 1 To RS1.Fields.Count-1
	        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&RS1(I)&"</span></td>"
        Next
        Response.Write "</tr></a>"
        Response.Flush
        Response.Clear
        RS1.MoveNext
      Wend
      Response.Write "</table>"
      RS1.Close
      Set RS1=Nothing
    End If
  Else
    Response.Write "<p align='center'><font color='#ff0000'>** 請先輸入查詢條件 **</font></p>"
  End If  
%>
  </form>  
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function UCaseMemberNo(){
  document.form.Member_No.value=document.form.Member_No.value.toUpperCase();
}	
function Query_OnClick(){
  if(document.form.Donor_Name.value!=''||document.form.Member_No.value!=''||document.form.Donor_Type.value!=''){
    document.form.submit();  
  }else{
    document.form.Donor_Name.focus();
    alert('請輸入查詢條件！');
  }
}
--></script>