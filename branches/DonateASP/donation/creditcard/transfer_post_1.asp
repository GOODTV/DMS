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
'交易日期
Donate_Year_AD=Cstr(Year(Date()))
Donate_Year=Cstr(Year(Date())-1911)
If Len(Donate_Year)=2 Then Donate_Year="0"&Donate_Year
Donate_Month=Cstr(Month(Date()))
If Len(Donate_Month)=1 Then Donate_Month="0"&Donate_Month
Donate_Day=Cstr(Day(Date()))
If Len(Donate_Day)=1 Then Donate_Day="0"&Donate_Day
'必要條件(局帳號不可為空白)
WhereSQL="And Post_SavingsNo<>'' And Post_AccountNo<>''"

Ary_Pledge_Id=Split(request("pledge_id"),",")
For I = 0 To UBound(Ary_Pledge_Id)	
  SQL="Select Pledge_Id,Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, " & _
      "Post_SavingsNo,Post_AccountNo,Post_IDNo,Donate_Amt=(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),DONOR.Donor_Name,Donate_ToDate " & _
      "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Status='授權中' And Pledge_Id='"&Ary_Pledge_Id(I)&"' "&WhereSQL&" "
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then
    '郵局明細資料
    If Row>0 Then Response.Write vbcrlf
    '1.資料別(明細為1)
    Response.Write "1"
    '2.存款別(存簿為P,劃撥為G)
    Response.Write "P"
    '3.事業單位代號
    Response.Write Transfer_StoreCode
    '4.區處站所
    SpaceLen=4
    Response.Write Space(SpaceLen)
    '5.繳費日期(民國年月日)
    Response.Write Donate_Year&Donate_Month&Donate_Day
    '6.保留欄
    SpaceLen=3
    Response.Write Space(SpaceLen)
    '7.1局帳
    SavingsNoLen=7
    Post_SavingsNo=RS("Post_SavingsNo")
    If Len(SavingsNoLen)>=SavingsNoLen Then
      Response.Write Left(Post_SavingsNo,SavingsNoLen)
    Else
      Response.Write Post_SavingsNo&Space(SavingsNoLen-Len(Post_SavingsNo))
    End If
    '7.2帳帳
    AccountNoLen=7
    Post_AccountNo=RS("Post_AccountNo")
    If Len(AccountNoLen)>=AccountNoLen Then
      Response.Write Left(Post_AccountNo,AccountNoLen)
    Else
      Response.Write Post_AccountNo&Space(AccountNoLen-Len(Post_AccountNo))
    End If
    '7.3劃撥固定值
    'Response.Write "000000"
    '7.4劃撥帳號
    'SpaceLen=8
    'Response.Write Space(SpaceLen)
    '8.身分證號
    PostIDNoLen=10
    Post_IDNo=RS("Post_IDNo")
    If Len(Post_IDNo)>=PostIDNoLen Then
      Response.Write Left(Post_IDNo,PostIDNoLen)
    Else
      Response.Write Post_IDNo&Space(PostIDNoLen-Len(Post_IDNo))
    End If
    '9.繳費金額
    DonateAmtLen=9
    Donate_Amt=RS("Donate_Amt")
    If Len(Donate_Amt)>=DonateAmtLen Then
      Response.Write Left(Donate_Amt,DonateAmtLen)&"00"
    Else
      StrDonateAmt=""
      For J=1 To DonateAmtLen
        StrDonateAmt=StrDonateAmt&"0"
      Next
      Response.Write Left(StrDonateAmt,DonateAmtLen-Len(Donate_Amt))&Donate_Amt&"00"
    End If
    '10用戶編號欄(商家自用)
    RemarkLen=10
    donor_remark=""
    donorremark=RS("Pledge_Id")
    iLen=0
    For J = 1 To Len(donorremark)
      If Asc(Mid(donorremark,J,1))<0 Then
    	  iLen=iLen+2
      Else
        iLen=iLen+1
      End If
      If iLen<=RemarkLen Then
        donor_remark=donor_remark&Mid(donorremark,J,1)
      Else
        Exit For
      End If
    Next
    Response.Write donor_remark
	  If iLen<RemarkLen Then Response.Write Space(RemarkLen-iLen)
    '11.事業單位使用欄
    Response.Write Left("0000000000",10-Len(Transfer_StoreCode))&Transfer_StoreCode
    '12.用戶編號列印記號
    Response.Write "1"
    '13.保留欄
    SpaceLen=1
    Response.Write Space(SpaceLen)
    '14.非連線記號
    SpaceLen=1
    Response.Write Space(SpaceLen)
    '15.變更存簿局號記號
    SpaceLen=1
    Response.Write Space(SpaceLen)
    '16.狀況代號
    SpaceLen=2
    Response.Write Space(SpaceLen)
    '17.繳費月份(民國年月)
    Response.Write Donate_Year&Donate_Month
    '18.保留欄
    SpaceLen=5
    Response.Write Space(SpaceLen)
    Row=Row+1
    Donate_Total=Cdbl(Donate_Total)+Cdbl(RS("Donate_Amt"))
    Response.Flush
    Response.Clear
  End If
  RS.Close
  Set RS=Nothing 
Next
'郵局總數資料
Response.Write vbcrlf
'1.資料別(總數)
Response.Write "2"
'2.存款別
SpaceLen=1
Response.Write Space(SpaceLen)
'3.事業單位代號
Response.Write Transfer_StoreCode
'4.區處站所
SpaceLen=4
Response.Write Space(SpaceLen)
'5.繳費日期
Response.Write Donate_Year&Donate_Month&Donate_Day
'6.保留欄
SpaceLen=3
Response.Write Space(SpaceLen)
'7.總件數
DonateRowLen=7
If Len(Cstr(Row))>=DonateRowLen Then
  Response.Write Left(Cstr(Row),DonateRowLen)
Else
  StrDonateRow=""
  For J=1 To DonateRowLen
    StrDonateRow=StrDonateRow&"0"
  Next
  Response.Write Left(StrDonateRow,DonateRowLen-Len(Cstr(Row)))&Row
End If
'8.總金額
DonateTotalLen=13
If Len(Cstr(Row))>=DonateTotalLen Then
  Response.Write Left(Cstr(Donate_Total),DonateTotalLen)
Else
  StrDonateTotal=""
  For J=1 To DonateTotalLen
    StrDonateTotal=StrDonateTotal&"00"
  Next
  Response.Write Left(StrDonateTotal,DonateTotalLen-Len(Cstr(Donate_Total)))&Donate_Total
End If
'9.保留欄
SpaceLen=16
Response.Write Space(SpaceLen)
'10.成功件數
DonateRowLen=7
For J=1 To DonateRowLen
  Response.Write "0"
Next
'11.成功金額
DonateTotalLen=13
For J=1 To DonateTotalLen
  Response.Write "0"
Next
'12.保留欄
SpaceLen=15
Response.Write Space(SpaceLen)
Conn.Close
Set Conn=Nothing
Session.Contents.Remove("Transfer_AspName")
%>
