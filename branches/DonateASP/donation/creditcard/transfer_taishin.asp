<%Response.ContentType="text/html; charset=utf-8"%>
<%
Response.Buffer =true
Response.Expires=-1
If session("user_id")="" Then
  Session.Abandon
  Response.Redirect("../sysmgr/timeout.asp")
End If
session.Timeout=60
Server.ScriptTimeout=600
Set Conn=server.createobject("ADODB.Connection")
Conn.Connectiontimeout=600
Conn.commandtimeout=600
Conn.Provider="sqloledb"
Conn.open "server="&session("server")&";uid="&session("uid")&";pwd="&session("pwd")&";database="&session("database")

'1.授權總筆數
Row=0
'2.授權總金額
Donate_Total=0
'3.授權日期
Donate_Year_AD=Cstr(Year(Date()))
Donate_Year_ROC=Cstr(Year(Date())-1911)
If Len(Donate_Year)=2 Then Donate_Year_ROC="0"&Donate_Year_ROC
Donate_Month=Cstr(Month(Date()))
If Len(Donate_Month)=1 Then Donate_Month="0"&Donate_Month
Donate_Day=Cstr(Day(Date()))
If Len(Donate_Day)=1 Then Donate_Day="0"&Donate_Day

SQL="Select * From DONATE_TRANSFER Where Transfer_AspName='"&Session("Transfer_AspName")&"'"
Set RS=Server.CreateObject("ADODB.Recordset")
RS.Open SQL,Conn,1,3
'4.Merchant Number(商店代號) 
StoreLen=15
Merchant_Number=Left(RS("Transfer_StoreCode")&Space(StoreLen-Len(RS("Transfer_StoreCode"))),StoreLen)
'5.MCC Code(終端機代號)
ServerLen=4
MCC_Code=Left(RS("Transfer_ServerCode")&Space(ServerLen-Len(RS("Transfer_ServerCode"))),ServerLen)
'6.Transfer_FileName (匯出檔案名稱)
OrderNo=""
If RS("Transfer_OrderNo")<>"" Then
  If Instr(RS("Transfer_OrderNo"),Donate_Year_AD&Donate_Month&Donate_Day)>0 Then
    OrderNo=Donate_Year_AD&Donate_Month&Donate_Day&Left("00",2-Len(Cint(Right(RS("Transfer_OrderNo"),2))+1))&Cint(Right(RS("Transfer_OrderNo"),2))+1
  Else
    OrderNo=Donate_Year_AD&Donate_Month&Donate_Day&"01"
  End If
Else
  OrderNo=Donate_Year_AD&Donate_Month&Donate_Day&"01"
End If
Transfer_FileName=Merchant_Number&OrderNo&".txt"
RS("Transfer_OrderNo")=OrderNo
RS.Update
RS.Close
Set RS=Nothing
'7.匯出欄位格式
Response.AddHeader "Content-disposition","attachment; filename="&Transfer_FileName&""
Response.ContentType = "application/vnd.ms-txt"

'查詢必要條件(卡號不可為空白)
WhereSQL="And Account_No<>''"
Ary_Pledge_Id=Split(request("pledge_id"),",")
For I = 0 To UBound(Ary_Pledge_Id)
  SQL="Select Pledge_Id,PLEDGE.Donor_Id,Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, " & _
      "Post_SavingsNo,Post_AccountNo,Post_IDNo,Donate_Amt=(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),DONOR.Donor_Name,Donate_ToDate " & _
      "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Pledge_Id='"&Trim(Ary_Pledge_Id(I))&"' And Status='授權中' "&WhereSQL&" "
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then
    'Data Record
    If Row>0 Then Response.Write vbcrlf
    '1.Record type
    Response.Write "D"
    '2.Transaction type
    Response.Write Space(2)
    '3.Merchant Number
    Response.Write Merchant_Number
    '4.MCC Code
    Response.Write MCC_Code
    '5.卡號
    CardLen=19
    Response.Write Left(RS("account_no")&Space(CardLen-Len(RS("account_no"))),CardLen)
    '6.信用卡有效年月
    ValidDateLen=4
    Valid_Date=Left(RS("Valid_Date"),2)&Right(RS("Valid_Date"),2)
    Response.Write Left(Valid_Date&Space(ValidDateLen-Len(Valid_Date)),ValidDateLen)
    '7.金額
    Response.Write Left("0000000000",10-Len(RS("Donate_Amt")))&RS("Donate_Amt")&"00"
    '8.Response Code
    Response.Write Space(2)
    '9.Approval Code
    Response.Write Space(6)
    '10.交易日期
    Response.Write Donate_Year_AD&Donate_Month&Donate_Day
    '11.SSL / SET flag
    Response.Write Space(1)
    '12.CVC2
    Response.Write Space(3)
    '13.SET XID
    Response.Write Space(40)
    '14.備註(商家自用)
    RemarkLen=60
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
    '15.交易類別Message Type
    Response.Write "00"
    '16.分期期數Installment Period
    Response.Write "00"
    '17.訂單編號Installment OrderNo
    Response.Write Space(12)
    '18.首期應繳金額First Period Amount
    Response.Write "000000000000"
    '19.每期應繳金額Period Amount
    Response.Write "000000000000"
    '20.手續費Process Fee
    Response.Write "000000000000"
    '21.首期手續費First Period Process Fee
    Response.Write "000000000000"
    '22.每期手續費Period Process Fee
    Response.Write "000000000000"
    '23.訊息來源Message Source
    Response.Write "1"&Space(1)
    '24.強制授權(郵購用)Strong Authorization
    Response.Write Space(1)
    '25.郵購商品代碼MO PRODUCTNO
    Response.Write "000000000000000"
    '26.郵購商品數量MO_Quantity
    Response.Write "000"
    '27.System Trace Audit No.
    Response.Write Space(6)
    '28.Transmission Date
    Response.Write Space(8)
    '29.Transmission Time
    Response.Write Space(6)
    '30.Retrieval Reference Number
    Response.Write Space(12)
    '31.Filler
    Response.Write Space(24)
    Row=Row+1
    Donate_Total=Cdbl(Donate_Total)+Cdbl(RS("Donate_Amt"))
    Response.Flush
    Response.Clear
  End If
  RS.Close
  Set RS=Nothing
Next
Response.Write vbcrlf

'Trailer Record
'1.Record type
Response.Write "T"
'2.Merchant Number
Response.Write Merchant_Number
'3.Total 05 items扣款
Response.Write Left("000000",6-Len(Row))&Cstr(Row)
'4.Total 05 amount扣款
Response.Write Left("0000000000",10-Len(Donate_Total))&Cstr(Donate_Total)&"00"
'5.Total 06 items刷退
Response.Write "000000"
'6.Total 06 amount刷退
Response.Write "000000000000"
Response.Write vbcrlf
Conn.Close
Set Conn=Nothing
Session.Contents.Remove("Transfer_AspName")
%>
