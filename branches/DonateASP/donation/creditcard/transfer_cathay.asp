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
Donate_Year_AD=Cstr(Year(Date()))
If Session("Donate_Date")="" Then 
	Donate_Year=Cstr(Year(Date())-1911)
	Donate_Month=Cstr(Month(Date()))
	Donate_Day=Cstr(Day(Date()))
Else 
	Donate_Year=Cstr(Year(CDate(Session("Donate_Date")))-1911) 
	Donate_Month=Cstr(Month(CDate(Session("Donate_Date"))))
	Donate_Day=Cstr(Day(CDate(Session("Donate_Date"))))
End If

If Len(Donate_Year)=2 Then Donate_Year="0"&Donate_Year
If Len(Donate_Month)=1 Then Donate_Month="0"&Donate_Month
If Len(Donate_Day)=1 Then Donate_Day="0"&Donate_Day
'必要條件(局帳號不可為空白)
WhereSQL="And Account_No<>''"

Ary_Pledge_Id=Split(request("pledge_id"),",")
For I = 0 To UBound(Ary_Pledge_Id)
  SQL="Select Pledge_Id,Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, " & _
      "Post_SavingsNo,Post_AccountNo,Post_IDNo,Donate_Amt=(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),DONOR.Donor_Name,Donate_ToDate " & _
      "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Status='授權中' And Pledge_Id='"&Ary_Pledge_Id(I)&"' "&WhereSQL&" "
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then
    If Row>0 Then Response.Write vbcrlf
    '1.商家代碼
    StoreLen=15
    If Len(Transfer_StoreCode)>=StoreLen Then
      Response.Write Left(Transfer_StoreCode,StoreLen)
    Else
      Response.Write Transfer_StoreCode&Space(StoreLen-Len(Transfer_StoreCode))
    End If
    '2.訂單編號
    OrderLen=20
    Donate_Order=Donate_Year&Donate_Month&Donate_Day&Left("0000",4-Len(Cstr(Row+1)))&Cstr(Row+1)
    Response.Write Space(OrderLen-Len(Donate_Order))&Donate_Order
    '3.卡號
    CardLen=16
    Account_No=RS("Account_No")
    If Len(Account_No)>=CardLen Then
      Response.Write Left(Account_No,CardLen)
    Else
      Response.Write Account_No&Space(CardLen-Len(Account_No))
    End If
    '4.信用卡CSV2
    AuthorizeLen=3
    Authorize=RS("Authorize")
	if isnull(RS("Authorize")) Then Authorize="   "
    If Len(Authorize)>=AuthorizeLen Then
      Response.Write Left(Authorize,AuthorizeLen)
    Else
      Response.Write Authorize&Space(AuthorizeLen-Len(Authorize))
    End If
    '5.有效年月(西元年月)
    ValidDateLen=6
    Valid_Date="20"&RS("Valid_Date")
    If Len(Valid_Date)>=ValidDateLen Then
      Response.Write Left(Valid_Date,ValidDateLen)
    Else
      Response.Write Valid_Date&Space(ValidDateLen-Len(Valid_Date))
    End If
    '6.金額
    AmtLen=8
    Donate_Amt=RS("Donate_Amt")
    For J= Len(Donate_Amt) To AmtLen-1
      Donate_Amt="0"&Donate_Amt
    Next
    Response.Write Donate_Amt
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
