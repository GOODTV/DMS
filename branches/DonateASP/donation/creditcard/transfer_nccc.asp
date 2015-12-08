<%Response.ContentType="text/html; charset=utf-8"%>
<%
Response.Buffer =true
Response.Expires=-1
If session("user_id")="" Then
  Session.Abandon
  Response.Redirect("../../sysmgr/timeout.asp")
End If
session.Timeout=60
Server.ScriptTimeout=600
Set Conn=server.createobject("ADODB.Connection")
Conn.Connectiontimeout=600
Conn.commandtimeout=600
Conn.Provider="sqloledb"
Conn.open "server="&session("server")&";uid="&session("uid")&";pwd="&session("pwd")&";database="&session("database")

SQL="Select * From DONATE_TRANSFER Where Transfer_AspName='"&Session("Transfer_AspName")&"'"
Set RS=Server.CreateObject("ADODB.Recordset")
RS.Open SQL,Conn,1,1
Transfer_StoreCode=RS("Transfer_StoreCode")
Transfer_ServerCode=RS("Transfer_ServerCode")
Transfer_FileName=RS("Transfer_FileName")
RS.Close
Set RS=Nothing

Response.AddHeader "Content-disposition","attachment; filename="&Transfer_FileName&""
Response.ContentType = "application/vnd.ms-txt"

'授權筆數
Row=0
'授權金額
Danate_Total=0
'20131216 Modify by GoodTV Tanya:交易日期為畫面上的「扣款日期」
'交易日期
'Donate_Year_AD=Cstr(Year(Date()))
Donate_Year=Cstr(Year(Date())-1911)
If Session("Donate_Date")="" Then 
	Donate_Year_AD=Cstr(Year(Date()))
	Donate_Month=Cstr(Month(Date()))
	Donate_Day=Cstr(Day(Date()))
Else 
	Donate_Year_AD=Cstr(Year(CDate(Session("Donate_Date")))) 
	Donate_Month=Cstr(Month(CDate(Session("Donate_Date"))))
	Donate_Day=Cstr(Day(CDate(Session("Donate_Date"))))
End If

If Len(Donate_Year)=2 Then Donate_Year="0"&Donate_Year
If Len(Donate_Month)=1 Then Donate_Month="0"&Donate_Month
If Len(Donate_Day)=1 Then Donate_Day="0"&Donate_Day	
'必要條件(卡號不可為空白)
WhereSQL="And Account_No<>''"

Ary_Pledge_Id=Split(request("pledge_id"),",")
For I = 0 To UBound(Ary_Pledge_Id)
  SQL="Select Pledge_Id,PLEDGE.Donor_Id,Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, " & _
      "Post_SavingsNo,Post_AccountNo,Post_IDNo,Donate_Amt=(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),DONOR.Donor_Name,Donate_ToDate " & _
      "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Pledge_Id='"&Trim(Ary_Pledge_Id(I))&"' And Status='授權中' "&WhereSQL&" "
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then
    If Row>0 Then Response.Write vbcrlf
    '1.商店代號
    StoreLen=10
    If Len(Transfer_StoreCode)>=StoreLen Then
      Response.Write Left(Transfer_StoreCode,StoreLen)
    Else
      Response.Write Transfer_StoreCode&Space(StoreLen-Len(Transfer_StoreCode))
    End If
    '2.端末機代號
    ServereLen=8
    If Len(Transfer_ServerCode)>=ServereLen Then
      Response.Write Left(Transfer_ServerCode,ServereLen)
    Else
      Response.Write Transfer_ServerCode&Space(ServereLen-Len(Transfer_ServerCode))
    End If
    '3.卡號
    CardLen=16
    account_no= RS("account_no")
    If Len(account_no)>=CardLen Then
      Response.Write Left(account_no,CardLen)
    Else
      Response.Write account_no&Space(CardLen-Len(account_no))
    End If
    '4.金額
    AmtLen=11
    Donate_Amt=RS("Donate_Amt")
    For J= Len(Donate_Amt) To AmtLen-1
      Donate_Amt="0"&Donate_Amt
    Next
    Response.Write Donate_Amt
    '5.空白(交易碼8)
    SpaceLen=8
    Response.Write Space(SpaceLen)
	  '6.交易日期
    Response.Write "00"&Right(Donate_Year_AD,2)&Donate_Month&Donate_Day
	  '7.信用卡有效年月
    ValidDateLen=4
    Valid_Date=RS("Valid_Date")
    If Len(Valid_Date)>=ValidDateLen Then
      Response.Write Left(Valid_Date,ValidDateLen)
    Else
      Response.Write Valid_Date&Space(ValidDateLen-Len(Valid_Date))
    End If   
    '8.空白(回應碼3,回應訊息6,處理日期6,處理時間6)
    SpaceLen=31
    Response.Write Space(SpaceLen)
    '9.備註(商家自用)
    RemarkLen=22
    Donor_Remark=""
    Donor_Name=""
    If Instr(RS("Donor_Name"),"&#")>0 Then
      For J=1 To Len(RS("Donor_Name"))
        If Mid(RS("Donor_Name"),J,1)="&" And J<Len(RS("Donor_Name")) Then
          If Mid(RS("Donor_Name"),J+1,1)="#" Then
            Donor_Name=Donor_Name&"*"
            J=J+7
          Else
            Donor_Name=Donor_Name&Mid(RS("Donor_Name"),J,1)
          End If
        Else
          Donor_Name=Donor_Name&Mid(RS("Donor_Name"),J,1)
        End If
      Next
    Else
      Donor_Name=RS("Donor_Name")
    End If
	Donor_Name=""
    DonorRemark=Left("00000",5-Len(RS("Pledge_Id")))&RS("Pledge_Id")&Donor_Name&RS("Donor_Id")
    iLen=0
    For J = 1 To Len(DonorRemark)
      If Asc(Mid(DonorRemark,J,1))<0 Then
    	  sLen=2
      Else
        sLen=1
      End If
      If sLen+iLen<=RemarkLen Then
        Donor_Remark=Donor_Remark&Mid(DonorRemark,J,1)
        iLen=iLen+sLen
      Else
        Exit For
      End If
    Next
    Response.Write Donor_Remark
	  If iLen<RemarkLen Then Response.Write Space(RemarkLen-iLen)
    '10.取消旗號
    Response.Write "0"
    '11.交易旗號
    Response.Write "2"
    Row=Row+1
    Donate_Total=Cdbl(Danate_Total)+Cdbl(RS("Donate_Amt"))
    Response.Flush
    Response.Clear
  End If
  RS.Close
  Set RS=Nothing
Next
Conn.Close
Set Conn=Nothing
Session.Contents.Remove("Transfer_AspName")
%>
