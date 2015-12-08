<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="upload" Then
  FilePath="../upload/"
  call DundasUpLoad(FilePath,UploadName,UploadSize,objUpload,"12")
  If UploadName<>"" Then
    Dept_Id=objUpload.form("Dept_Id")
    Donate_Date=objUpload.form("Donate_Date")

    StrCity=""
    SQL_City = "Select mValue,mCode From CODECITY Where codeMetaID = 'Addr0' Order By Seq,mSortValue"
    Set RS_City = Server.CreateObject("ADODB.RecordSet")
    RS_City.Open SQL_City,Conn,1,1
    While Not RS_City.EOF
      If StrCity="" Then
        StrCity=RS_City("mValue")&","&RS_City("mCode")
      Else
        StrCity=StrCity&","&RS_City("mValue")&","&RS_City("mCode")
      End If
      Row_Area=0
      SQL_Area = "Select mValue,mCode From CodeCity Where codeMetaID = 'Addr0R"&RS_City("mCode")&"' Order By mSortValue"
      Set RS1_Area = Server.CreateObject("ADODB.RecordSet")
      RS1_Area.Open SQL_Area,Conn,1,1
      While Not RS1_Area.EOF
        If Row_Area=0 Then
          StrCity=StrCity&","&RS1_Area("mValue")&"/"&RS1_Area("mCode")
        Else
          StrCity=StrCity&"/"&RS1_Area("mValue")&"/"&RS1_Area("mCode")
        End If
        Row_Area=Row_Area+1
        RS1_Area.MoveNext
	      Response.Flush
        Response.Clear
      Wend
      RS1_Area.Close
	    Set RS1_Area=Nothing
      RS_City.MoveNext
	  Wend
	  RS_City.Close
	  Set RS_City=Nothing
	  AryCity=Split(StrCity,",")
	  
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
    
	  UploadFile="../upload/"&UploadName&""
	  Set FS=Server.CreateObject("Scripting.FileSystemObject")
	  Set Txt=FS.OpenTextFile(Server.MapPath(UploadFile))
	  While Not Txt.atEndOfStream
	    Line=Txt.ReadLine
	    If Line<>"" Then
	      AryLine=Split(Line,",")
	      Donate_Payment="花旗銀行"
        Excel_Type="citi"
	      IDNO=Trim(Cstr(AryLine(0)))
	      Invoice_IDNo=Trim(Cstr(AryLine(0)))
	      Donor_Name=Trim(Cstr(AryLine(1)))
	      Invoice_Title=Trim(Cstr(AryLine(1)))
	      Tel_Office=Trim(Cstr(AryLine(2)))
	      Donate_Amt=Trim(Cstr(AryLine(3)))
        ZipCode=""
        City=""
        Area=""
	      Address=Trim(Cstr(AryLine(4)))
	      If Address<>"" Then
          For I = 0 To UBound(AryCity) Step 3
            If Instr(Address,AryCity(I))>0 Then
              City=AryCity(I+1)
              Address=Replace(Address,AryCity(I),"")
              AryArea=Split(AryCity(I+2),"/")
              For J = 0 To UBound(AryArea) Step 2
                If Instr(Address,AryArea(J))>0 Then 
                  ZipCode=AryArea(J+1)
                  Area=AryArea(J+1)
                  Address=Replace(Address,AryArea(J),"")
                  Exit For
                End If  
              Next
              Exit For
            End If
          Next
	      End If
        Invoice_ZipCode=ZipCode
        Invoice_City=City
        Invoice_Area=Area
        Invoice_Address=Address

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
        If Tel_Office<>"" Then
          If WhereSQL="" Then
            WhereSQL=WhereSQL&"And (Cellular_Phone='"&Cellular_Phone&"' Or Tel_Office='"&Tel_Office&"' Or Tel_Home='"&Tel_Office&"' "
          Else
             WhereSQL=WhereSQL&"Or Cellular_Phone='"&Cellular_Phone&"' Or Tel_Office='"&Tel_Office&"' Or Tel_Home='"&Tel_Office&"' "
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
          RS1("Title")="君"
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
          RS1("Cellular_Phone")=""
          RS1("Tel_Office")=Tel_Office
          RS1("Tel_Home")=""
          RS1("Fax")=""
          RS1("Email")=""
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
          If Address<>"" Then
            RS1("Invoice_type")=InvoiceTypeT
          Else
            RS1("Invoice_type")=InvoiceTypeN
          End If
          RS1("IsAnonymous")="N"
          RS1("NickName")=""
          RS1("Invoice_Title")=Invoice_Title
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
        SQL1="Select * From DONATE Where Excel_Type='"&Excel_Type&"' And Donor_Id='"&Donor_Id&"' And Donate_Date='"&Donate_Date&"' And CONVERT(numeric,Donate_amt)='"&Donate_amt&"'"
        Set RS1=Server.CreateObject("ADODB.RecordSet")
        RS1.Open SQL1,Conn,1,3
        If RS1.EOF Then
          Invoice_Pre=""
          Invoice_No=""
          InvoiceNo=Get_Invoice_No2("1",Dept_Id,Donate_Date,InvoiceTypeT,"")
          If InvoiceNo<>"" Then
            Invoice_Pre=Split(InvoiceNo,"/")(0)
            Invoice_No=Split(InvoiceNo,"/")(1)
          End If
          
          RS1.Addnew
          RS1("Excel_Type")=Excel_Type
          RS1("Excel_No")=""
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
          RS1("Comment")=""
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
      End If
	    Response.Flush
      Response.Clear
	  Wend
    session("errnumber")=1
    session("msg")="花旗銀行捐款資料匯入成功 ！\n\n共計："&FormatNumber(Row,0)&"筆，金額："&FormatNumber(Total,0)&"元。"	  
  End If
End If
%>
<%Prog_Id="citibank"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form  name="form" action="import_citibank.asp?action=upload" method="post" enctype="multipart/form-data">	
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
                      <td class="td02-c"><%call Calendar("Donate_Date",Date())%>
                      </td>	
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">上傳檔案：</td>
                      <td class="td02-c">
                      	<input name="AttachFile" size="48" type="file" class="font9">&nbsp;
                        <input type="button" value="匯入『&nbsp;花旗捐款&nbsp;』資料" name="input" class="addbutton" style="cursor:hand" onClick="Input_OnClick()">
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right" width="150">※&nbsp;注意事項：</td>
                      <td class="td02-c" width="630">
                      	1.您選擇的『&nbsp;<font color="#FF0000">機構名稱</font>&nbsp;』及『&nbsp;<font color="#FF0000">捐款日期</font>&nbsp;』會影響&nbsp;<font color="#FF0000">收據編號</font>&nbsp;取號規則，上傳前請您務必確認。<br />
                      	2.相關欄位說明請下載&nbsp;<a href='import/花旗銀行參考範例.csv'>參考範例</a>&nbsp;。
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
  document.form.Donate_Date.focus();
}	
function Input_OnClick(){
  <%call CheckStringJ("Dept_Id","機構名稱")%>
  <%call CheckStringJ("Donate_Date","捐款日期")%>
  <%call CheckDateJ("Donate_Date","另訂捐款日期")%>
  <%call CheckStringJ("AttachFile","檔案位置")%>
  if(confirm('您是否確定要匯入『花旗銀行捐款』資料？\n\n※請注意匯入過程中請勿關閉視窗')){
    document.form.submit();
  }
}
--></script>