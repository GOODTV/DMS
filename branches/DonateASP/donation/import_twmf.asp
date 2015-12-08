<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="upload" Then
  FilePath="../upload/"
  call DundasUpLoad(FilePath,UploadName,UploadSize,objUpload,"12")
  If UploadName<>"" Then
    Dept_Id=objUpload.form("Dept_Id")
    If objUpload.form("Date_Type")="1" Then
      Dept_Date=""
    Else
      Dept_Date=objUpload.form("Donate_Date")
    End If
    Sheet_Name=objUpload.form("Sheet_Name")

    InvoiceTypeY="年度收據"
    SQL1="Select InvoiceTypeY=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Like '%年%' Order By Seq"
    Set RS1=Server.CreateObject("ADODB.Recordset")
    RS1.Open SQL1,Conn,1,1
    If Not RS1.EOF Then InvoiceTypeY=RS1("InvoiceTypeY")
    RS1.Close
    Set RS1=Nothing

    InvoiceTypeN="不寄收據"
    SQL1="Select InvoiceTypeN=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Like '%不%' Order By Seq"
    Set RS1=Server.CreateObject("ADODB.Recordset")
    RS1.Open SQL1,Conn,1,1
    If Not RS1.EOF Then InvoiceTypeN=RS1("InvoiceTypeN")
    RS1.Close
    Set RS1=Nothing

    InvoiceTypeT="單次開立"
    SQL1="Select InvoiceTypeT=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Not In ('"&InvoiceTypeY&"','"&InvoiceTypeN&"') Order By Seq"
    Set RS1=Server.CreateObject("ADODB.Recordset")
    RS1.Open SQL1,Conn,1,1
    If Not RS1.EOF Then InvoiceTypeT=RS1("InvoiceTypeT")
    RS1.Close
    Set RS1=Nothing

    Row=0
    Total=0
    ExcelFile	= Server.MapPath("../upload/"&UploadName&"")
    Set ConnX = Server.CreateObject("ADODB.Connection")
    Driver = "Driver={Microsoft Excel Driver (*.xls)};"
    DBPath = "DBQ=" & ExcelFile
    ConnX.Open Driver & DBPath
    SQLX="Select * From ["&Sheet_Name&"$]"
    Set RSX=ConnX.Execute(SQLX)
    While Not RSX.EOF
      If Trim(RSX(0))<>"" Then
        Remark=""
        Donate_Payment="台哥大5180"
        Excel_Type="twmf"
        Excel_No=Trim(RSX(5))&" "&Trim(RSX(6))
        Cellular_Phone=""
        If Trim(RSX(3))<>"" Then Cellular_Phone=Trim(RSX(3))
        Donate_Amt="0"
        If Trim(RSX(4))<>"" Then Donate_Amt=Trim(RSX(4))
        Donate_Date=""
        If Dept_Date<>"" Then
         Donate_Date=Dept_Date
        Else
          If Trim(RSX(5))<>"" Then Donate_Date=Trim(RSX(5))
        End If        
        Remark=Remark&"帳單月份："&Trim(RSX(7))&vbcrlf
        InvoiceYN=Trim(RSX(8))
        Donor_Name=""
        If Trim(RSX(9))<>"" Then Donor_Name=Trim(RSX(9))
        Title=""
        If Trim(RSX(10))<>"" Then Title=Trim(RSX(10))
        Sex=""
        If Trim(RSX(10))="小姐" Then Sex="女"
        If Trim(RSX(10))="先生" Then Sex="男"
        Remark=Remark&"帳單週期："&Trim(RSX(12))&vbcrlf
   
        ZipCode=""
        City=""
        Area=""
        Address=Trim(RSX(11))
        If Address<>"" Then
          Address=Replace(Address,"臺","台")
          Address=Replace(Address,"台北縣","新北市")
          Address=Replace(Address,"台中縣","台中市")
          Address=Replace(Address,"台南縣","台南市")
          Address=Replace(Address,"高雄縣","高雄市")
          If Instr(Address,"新竹市")>0 Then
            ZipCode="300"
            City="F"
            Area="300"
            Address=Replace(Address,"新竹市","")
            If Instr(Address,"北區")>0 Then
              Area="3001"
              Address=Replace(Address,"北區","")
            End If
            If Instr(Address,"香山區")>0 Then
              Area="3002"
              Address=Replace(Address,"香山區","")
            End If
          ElseIf Instr(Address,"嘉義市")>0 Then
            ZipCode="600"
            City="N"
            Area="600"
            Address=Replace(Address,"嘉義市","")
            If Instr(Address,"西區")>0 Then
              Area="6001"
              Address=Replace(Address,"西區","")
            End If
          Else
            SQL1="Select City=B.mCode,Area=A.mCode,City_Name=B.mValue,Area_Name=A.mValue From CODECITY A Join CODECITY B On Substring(A.codeMetaID,7,1)=B.mCode Where B.mValue='"&Left(Address,3)&"' And (A.mValue='"&Mid(Address,4,2)&"' Or A.mValue='"&Mid(Address,4,3)&"' Or A.mValue='"&Mid(Address,4,4)&"') "
            Set RS1 = Server.CreateObject("ADODB.RecordSet")
            RS1.Open SQL1,Conn,1,1
            If Not RS1.EOF Then 
              ZipCode=RS1("Area")
              City=RS1("City")
              Area=RS1("Area")
              Address=Replace(Address,RS1("City_Name"),"")
              Address=Replace(Address,RS1("Area_Name"),"")
            End If
            RS1.Close
            Set RS1=Nothing
          End If
        End If
        
        Invoice_ZipCode=ZipCode
        Invoice_City=City
        Invoice_Area=Area
        Invoice_Address=Address
        Invoice_Title=Donor_Name
        Title2=Title
        IDNo=""
        Invoice_IDNo=""
        Email=""
        
        '捐款人資料(DONOR)
        Donor_Id=0
        SQL1="Select * From DONOR Where Donor_Name='"&Donor_Name&"' "
        WhereSQL=""
        If IDNO<>"" Then
          If WhereSQL="" Then
            WhereSQL=WhereSQL&"And (IDNo='"&IDNO&"' "
          Else
             WhereSQL=WhereSQL&"Or IDNo='"&IDNO&"' "
          End If
        End If
        If ZipCode<>"" Then
          If WhereSQL="" Then
            WhereSQL=WhereSQL&"And (ZipCode='"&ZipCode&"' Or Invoice_ZipCode='"&ZipCode&"' "
          Else
             WhereSQL=WhereSQL&"Or ZipCode='"&ZipCode&"' Or Invoice_ZipCode='"&ZipCode&"' "
          End If
        End If        
        If Address<>"" Then
          If WhereSQL="" Then
            WhereSQL=WhereSQL&"And (Address='"&Address&"' Or Invoice_Address='"&Address&"' "
          Else
             WhereSQL=WhereSQL&"Or Address='"&Address&"' Or Invoice_Address='"&Address&"' "
          End If
        End If
        If Cellular_Phone<>"" Then
          If WhereSQL="" Then
            WhereSQL=WhereSQL&"And (Cellular_Phone='"&Cellular_Phone&"' "
          Else
             WhereSQL=WhereSQL&"Or Cellular_Phone='"&Cellular_Phone&"' "
          End If
        End If
        If Tel_Office<>"" Then
          If WhereSQL="" Then
            WhereSQL=WhereSQL&"And (Tel_Office='"&Tel_Office&"' Or Tel_Home='"&Tel_Office&"' "
          Else
             WhereSQL=WhereSQL&"Or Tel_Office='"&Tel_Office&"' Or Tel_Home='"&Tel_Office&"' "
          End If
        End If
        If Email<>"" Then
          If WhereSQL="" Then
            WhereSQL=WhereSQL&"And (Email='"&Email&"' "
          Else
             WhereSQL=WhereSQL&"Or Email='"&Email&"' "
          End If
        End If
        If WhereSQL<>"" Then WhereSQL=WhereSQL&")"
        SQL1=SQL1&WhereSQL
        Set RS1=Server.CreateObject("ADODB.RecordSet")
        RS1.Open SQL1,Conn,1,3
        If RS1.EOF Then
          RS1.Addnew
          RS1("Donor_Name")=Donor_Name
          RS1("Dept_Id")=Dept_Id
          RS1("Title")=Title
          RS1("Title2")=Title2
          If IDNO<>"" Then
            If Len(IDNO)=8 Then
              RS1("Category")="團體"
              RS1("Sex")=""
            Else
              RS1("Category")="個人"
              If Mid(IDNO,2,1)="1" Then RS1("Sex")="男"
              If Mid(IDNO,2,1)="2" Then RS1("Sex")="女"
            End If
          Else
            RS1("Sex")=""
            RS1("Category")="個人"
          End If
          RS1("Donor_type")=""
          RS1("IDNo")=IDNo
          RS1("Birthday")=null
          RS1("Education")=""
          RS1("Occupation")=""
          RS1("Marriage")=""
          RS1("Religion")=""
          RS1("ReligionName")=""
          RS1("Cellular_Phone")=Cellular_Phone
          RS1("Tel_Office")=Tel_Office
          RS1("Tel_Home")=""
          RS1("Fax")=""
          RS1("Email")=Email
          RS1("Contactor")=""
          RS1("OrgName")=""
          RS1("JobTitle")=""
          RS1("IsAbroad_Invoice")="N"
          RS1("ZipCode")=ZipCode
          RS1("City")=City
          RS1("Area")=Area
          RS1("Address")=Address
          RS1("IsSendNews")="N"
          RS1("IsSendEpaper")="N"
          RS1("IsSendYNews")="N"
          RS1("IsBirthday")="N"
          RS1("IsXmas")="N"
          If InvoiceYN="Y" Then
            RS1("Invoice_type")=InvoiceTypeT
          Else
            RS1("Invoice_type")=InvoiceTypeN
          End If
          RS1("IsAnonymous")="N"
          RS1("NickName")=""
          If Invoice_Title<>"" Then
            RS1("Invoice_Title")=Invoice_Title
          Else
           RS1("Invoice_Title")=Donor_Name
          End If
          If Invoice_IDNo<>"" Then
           RS1("Invoice_IDNo")=Invoice_IDNo
          Else
            RS1("Invoice_IDNo")=IDNo
          End If
          If Invoice_ZipCode<>"" Then
            RS1("Invoice_ZipCode")=Invoice_ZipCode
            RS1("Invoice_City")=Invoice_City
            RS1("Invoice_Area")=Invoice_Area
            RS1("Invoice_Address")=Invoice_Address
          Else
            RS1("Invoice_ZipCode")=ZipCode
            RS1("Invoice_City")=City
            RS1("Invoice_Area")=Area
            RS1("Invoice_Address")=Address
          End If
          RS1("Remark")=""
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
          RS1("Create_User")=Session("user_name")
          RS1("Create_IP")=Request.ServerVariables("REMOTE_ADDR")
          RS1("IsMember")="N"
          RS1("Member_No")=""
          RS1("Member_Type")=""
          RS1("Member_Status")=""
          RS1("IsVolunteer")="N"
          RS1("Volunteer_Type")=""
          RS1("Volunteer_Status")=""
          RS1("IsGroup")="N"
          RS1("Group_Licence")=""
          RS1("Group_CreateDate")=null
          RS1("Group_Person")=""
          RS1("Group_Person_JobTitle")=""
          RS1("Group_WebUrll")=""
          RS1("Group_Mission")="" 
          RS1.Update
          RS1.Close
          Set RS1=Nothing
        
          '取捐款人編號
          SQL1="Select @@IDENTITY As Donor_Id"
          Set RS1=Server.CreateObject("ADODB.RecordSet")
          RS1.Open SQL1,Conn,1,1
          If Not RS1.EOF Then Donor_Id=RS1("Donor_Id")
          RS1.Close
          Set RS1=Nothing
        Else
          Donor_Id=Cdbl(RS1("Donor_Id"))
          If IDNO<>"" Then
            If Len(IDNO)=8 Then
              If RS1("Category")="" Then RS1("Category")="團體"
            Else
              If RS1("Category")="" Then RS1("Category")="個人"
              If RS1("Sex")="" And Mid(IDNO,2,1)="1" Then RS1("Sex")="男"
              If RS1("Sex")="" And Mid(IDNO,2,1)="2" Then RS1("Sex")="女"
            End If
            RS1("IDNo")=IDNo
          End If
          If Cellular_Phone<>"" Then RS1("Cellular_Phone")=Cellular_Phone
          If Tel_Office<>"" Then RS1("Tel_Office")=Tel_Office
          If Address<>"" Then
            RS1("ZipCode")=ZipCode
            RS1("City")=City
            RS1("Area")=Area
            RS1("Address")=Address
          End If
          If Invoice_Title<>"" Then RS1("Invoice_Title")=Invoice_Title
          If Invoice_IDNo<>"" Then RS1("Invoice_IDNo")=Invoice_IDNo
          RS1("IsAbroad_Invoice")="N"
          If Invoice_ZipCode<>"" Then
            RS1("Invoice_ZipCode")=Invoice_ZipCode
            RS1("Invoice_City")=Invoice_City
            RS1("Invoice_Area")=Invoice_Area
            RS1("Invoice_Address")=Invoice_Address
          End If
          RS1.Update
          RS1.Close
          Set RS1=Nothing
        End If

        '捐款紀錄
        SQL1="Select * From DONATE Where Excel_Type='"&Excel_Type&"' And Excel_No='"&Excel_No&"' And Donor_Id='"&Donor_Id&"'"
        Set RS1=Server.CreateObject("ADODB.RecordSet")
        RS1.Open SQL1,Conn,1,3
        If RS1.EOF Then
          Invoice_Pre=""
          Invoice_No=""
          If InvoiceYN="Y" Then
            InvoiceNo=Get_Invoice_No2("1",Dept_Id,Donate_Date,InvoiceTypeT,"")
            If InvoiceNo<>"" Then
              Invoice_Pre=Split(InvoiceNo,"/")(0)
              Invoice_No=Split(InvoiceNo,"/")(1)
            End If
          End If  
          RS1.Addnew
          RS1("Excel_Type")=Excel_Type
          RS1("Excel_No")=Excel_No
          RS1("Donor_Id")=Donor_Id
          RS1("Donate_Date")=Donate_Date
          RS1("Donate_Amt")=Donate_Amt
          RS1("Donate_AmtD")=Donate_Amt
          RS1("Donate_Fee")="0"
          RS1("Donate_FeeD")="0"
          RS1("Donate_Accou")=Donate_Amt
          RS1("Donate_AccouD")=Donate_Amt
          RS1("Donate_AmtM")="0"
          RS1("Donate_FeeM")="0"
          RS1("Donate_AccouM")="0"
          RS1("Donate_AmtA")="0"
          RS1("Donate_FeeA")="0"
          RS1("Donate_AccouA")="0"
          RS1("Donate_AmtS")="0"
          RS1("Donate_RateS")="0"
          RS1("Donate_FeeS")="0"
          RS1("Donate_AccouS")="0"
          RS1("Donate_Payment")=Donate_Payment
          RS1("Donate_Purpose")="一般捐款"
          RS1("Donate_Purpose_Type")="D"
          RS1("Donate_Type")="單次捐款"
          RS1("Donate_Forign")=""
          RS1("Donate_Desc")=""
          RS1("IsBeductible")="N"
          RS1("Donate_Amt2")="0"
          RS1("Card_Bank")=""
          RS1("Card_Type")=""
          RS1("Account_No")=""
          RS1("Valid_Date")=""
          RS1("Card_Owner")=""
          RS1("Owner_IDNo")=""
          RS1("Relation")=""
          RS1("Authorize")=""
          RS1("Check_No")=""
          RS1("Check_ExpireDate")=null
          RS1("Post_Name")=""
          RS1("Post_IDNo")=""
          RS1("Post_SavingsNo")=""
          RS1("Post_AccountNo")=""
          RS1("Dept_Id")=Dept_Id
          RS1("Invoice_Title")=Invoice_Title
          RS1("Invoice_Pre")=Invoice_Pre
          RS1("Invoice_No")=Invoice_No
          RS1("Invoice_Print")="0"
          RS1("Invoice_Print_Add")="0"
          RS1("Invoice_Print_Yearly")="0"
          RS1("Invoice_Print_Yearly_Add")="0"
          RS1("Request_Date")=null
          RS1("Accoun_Bank")=""
          RS1("Accoun_Date")=null
          If Address<>"" Then
            RS1("Invoice_type")=InvoiceTypeT
          Else
            RS1("Invoice_type")=InvoiceTypeN
          End If
          RS1("Accounting_Title")=""
          RS1("Act_id")=null
          RS1("Comment")=Remark
          RS1("Issue_Type")=""
          RS1("Issue_Type_Keep")=""
          RS1("Export")="N"
          RS1("Create_Date")=Date()
          RS1("Create_User")=Session("user_name")
          RS1("Create_IP")=Request.ServerVariables("REMOTE_ADDR")
          RS1.Update
          RS1.Close
          Set RS1=Nothing
        
          '取捐款PK
          SQL1="Select @@IDENTITY As donate_id"
          Set RS1 = Server.CreateObject("ADODB.RecordSet")
          RS1.Open SQL1,Conn,1,1
          donate_id=RS1("donate_id")
          RS1.Close
          Set RS1=Nothing

          '確認收據編號無重覆
          If Invoice_No<>"" Then
            Invoice_Pre_Old=Invoice_Pre
            Invoice_No_Old=Invoice_No
            Check_InvoiceNo=False
            While Check_InvoiceNo=False
              Check_Donate=False
              SQL1="Select * From DONATE Where Invoice_Pre='"&Invoice_Pre_Old&"' And Invoice_No='"&Invoice_No_Old&"' And Donate_Id<>'"&donate_id&"' "
              Call QuerySQL(SQL1,RS1)
              If RS1.EOF Then Check_Donate=True
              RS1.Close
              Set RS1=Nothing

              Check_Contribute=False
              If Check_Donate Then
                SQL1="Select * From CONTRIBUTE Where Invoice_Pre='"&Invoice_Pre_Old&"' And Invoice_No='"&Invoice_No_Old&"'"
                Call QuerySQL(SQL1,RS1)
                If RS1.EOF Then Check_Contribute=True
                RS1.Close
                Set RS1=Nothing
              End If

              If Check_Contribute And Check_Donate Then
                Check_InvoiceNo=True
              Else
                InvoiceNo=Get_Invoice_No2("1",Dept_Id,Donate_Date,InvoiceTypeT,"")
                If InvoiceNo<>"" Then
                  Invoice_Pre_Old=Split(InvoiceNo,"/")(0)
                  Invoice_No_Old=Split(InvoiceNo,"/")(1)
                  SQL="Update DONATE Set Invoice_Pre='"&Invoice_Pre_Old&"',Invoice_No='"&Invoice_No_Old&"' Where Donate_Id='"&donate_id&"' "
                  Set RS=Conn.Execute(SQL)
                End If
              End If
            Wend
          End If
        Else
          RS1("Donate_Date")=Donate_Date
          RS1("Donate_Amt")=Donate_Amt
          RS1("Donate_AmtD")=Donate_Amt
          RS1.Update
          RS1.Close
          Set RS1=Nothing
        End If

        '修改捐款人捐款紀錄
        call Declare_DonorId (Donor_Id)

        Row=Row+1
        Total=Cdbl(Total)+Cdbl(Donate_Amt)
        Response.Flush 
        Response.Clear
      End If
      RSX.MoveNext 
    Wend
    RSX.Close
    Set RSX=Nothing
    ConnX.Close
    Set ConnX=Nothing
    session("errnumber")=1
    session("msg")="台哥大捐款資料匯入成功 ！\n\n共計："&FormatNumber(Row,0)&"筆，金額："&FormatNumber(Total,0)&"元。"
  End If
End If
%>
<%Prog_Id="import_twmf"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form  name="form" action="import_twmf.asp?action=upload" method="post" enctype="multipart/form-data">	
      <input type="hidden" name="action">
      <table width="700" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="100%"  border="0" cellspacing="0" cellpadding="0" align="center" style="background-color:#EEEEE3">
              <tr>
                <td width="5%"></td>
                <td width="95%">
  		            <table width="60%"  border="0" cellspacing="0" cellpadding="0">
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
	          <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
              </tr>
              <tr>
                <td class="table62-bg">&nbsp;</td>
                <td valign="top">
                  <table width="100%" border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111">
                    <tr>
                      <td class="td02-c" align="right">機構名稱：</td>
                      <td class="td02-c">
                      <%
                        If Session("comp_label")="1" Then
                          SQL="Select Dept_Id,Comp_ShortName From DEPT Where Comp_Label>='"&Session("comp_label")&"' And Dept_Id In ("&Session("all_dept_type")&") Order By Comp_Label,Dept_Id"
                        Else
                          SQL="Select Dept_Id,Comp_ShortName From DEPT Where (Dept_Id In ("&Session("all_dept_type")&") Or Comp_Label>'"&Session("comp_label")&"') Order By Comp_Label,Dept_Id"
                        End If
                        FName="Dept_Id"
                        Listfield="Comp_ShortName"
                        menusize="1"
                        BoundColumn=Session("dept_id")
                        call optionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                      %>
                  　  </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">捐款日期：</td>
                      <td class="td02-c">
                      	<input type="radio" name="Date_Type" value="1" checked >依檔案內的『&nbsp;<b>捐款日期</b>&nbsp;』
                        <input type="radio" name="Date_Type" value="2" >另訂捐款日期為：<%call Calendar("Donate_Date",Date())%>
                      </td>	
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">工作表名稱：</td>
                      <td class="td02-c"><input type="text" name="Sheet_Name" size="11" class="font9" maxlength="60" value="Sheet1"></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">上傳檔案：</td>
                      <td class="td02-c">
                      	<input name="AttachFile" size="48" type="file" class="font9">&nbsp;
                        <input type="button" value="匯入『&nbsp;台哥大捐款&nbsp;』資料" name="input" class="addbutton" style="cursor:hand" onClick="Input_OnClick()">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right" width="150">※&nbsp;注意事項：</td>
                      <td class="td02-c" width="630">
                      	1.您選擇的『&nbsp;<font color="#FF0000">機構名稱</font>&nbsp;』及『&nbsp;<font color="#FF0000">捐款日期</font>&nbsp;』會影響&nbsp;<font color="#FF0000">收據編號</font>&nbsp;取號規則，上傳前請您務必確認。<br />
                      	2.工作表名稱必須與檔案的工作表名稱一致否則將導致作業失敗。<br />
                      	3.檔案必須為&nbsp;Excel&nbsp;格式，如您的檔案來源自網站下載請務必&nbsp;<font color="#FF0000">另存新檔</font>&nbsp;，檔案大小請勿超過10M。<br />
                      	4.相關欄位說明請下載&nbsp;<a href='import/台哥大參考範例.xls'>參考範例</a>&nbsp;。
                      </td>
                    </tr>
                    <!--#include file="../include/calendar2.asp"-->
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
  document.form.Sheet_Name.focus();
}	
function Input_OnClick(){
  <%call CheckStringJ("Dept_Id","機構名稱")%>
  <%call CheckStringJ("Sheet_Name","工作表名稱")%>
  <%call CheckStringJ("AttachFile","檔案位置")%>
  var ExtName=document.form.AttachFile.value.substr(document.form.AttachFile.value.lastIndexOf('.')+1,document.form.AttachFile.value.length).toLowerCase();
  if(ExtName!='xls'){
    alert('捐款匯入只接受xls檔案格式');
    document.form.AttachFile.focus();
    return;
  }
  if(document.form.Date_Type[1].checked){
    <%call CheckStringJ("Donate_Date","另訂捐款日期")%>
  }
  if(document.form.Donate_Date.value!=''){
    <%call CheckDateJ("Donate_Date","另訂捐款日期")%>
  }  
  if(confirm('您是否確定要匯入『台哥大捐款』資料？\n\n※請注意匯入過程中請勿關閉視窗')){
    document.form.submit();
  }
}
--></script>