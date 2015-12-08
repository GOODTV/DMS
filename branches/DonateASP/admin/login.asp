<%
login_return="01000011000"
NextPage=""
Set FS = Server.CreateObject("Scripting.FileSystemObject")
If FS.FileExists(lcase(Replace(Server.Mappath("create.asp"),"admin","upload"))) Then
  NextPage="../upload/create.asp"
ElseIf FS.FileExists(lcase(Replace(Server.Mappath("license.asp"),"admin","upload"))) Then
  NextPage="../upload/license.asp"
ElseIf FS.FileExists(lcase(Replace(Server.Mappath("expiredate.asp"),"admin","upload"))) Then
  NextPage="../upload/expiredate.asp"
ElseIf FS.FileExists(lcase(Replace(Server.Mappath("delay.asp"),"admin","upload"))) Then
  NextPage="../upload/delay.asp"  
End If
Set FS = Nothing
If NextPage<>"" Then Response.Redirect NextPage
%>
<!--#include file="../include/check_query_string.asp"-->
<!--#include file="../include/md5.asp"-->
<!--#include file="../global.asp"-->
<%
If Request("user_id") = "" Then
  Session("errnumber")=1
  Session("msg")="帳號 欄位不可為空白 ！"
  Response.Redirect "index.asp"
ElseIf Request("password") = "" then
  Session("errnumber")=1
  Session("msg")="密碼 欄位不可為空白 ！"
  Response.Redirect "index.asp"
Else
  If Check_Post() Then
    call check_string(Request("user_id"),"Y","string","20")
    call check_string(Request("password"),"Y","string","20")
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Provider="sqloledb"
    Conn.open "server="&session("server")&";uid="&session("uid")&";pwd="&session("pwd")&";database="&session("database")
    
    Check_MD5=True
    Set FS = Server.CreateObject("Scripting.FileSystemObject")
    If Cstr(Request.ServerVariables("SERVER_NAME"))="web.npois.com.tw" Then
      InetpubPath = lcase(Server.MapPath("../index.asp"))
    Else
      InetpubPath = lcase(Replace(Server.MapPath("index.asp"),"admin\index.asp",""))
    End If
    DrivePath = FS.GetDriveName(InetpubPath)
    Set Drive = FS.GetDrive(DrivePath)
    DriveID = Trim(Drive.SerialNumber)
    Set Drive = Nothing
    Set FS = Nothing
    
    SQL="Select Style_Name,Style_Value From WEB_STYLE Where Style_Type='MD5' Order By Style_Seq"
    Set RS=Server.CreateObject("ADODB.Recordset")
    RS.Open SQL,Conn,1,1
    If Not RS.EOF Then
      While Not RS.EOF
        If RS("Style_Name")="KEY_1" Then KEY_1=RS("Style_Value")
        If RS("Style_Name")="KEY_2" Then KEY_2=RS("Style_Value")
        If RS("Style_Name")="KEY_3" Then KEY_3=RS("Style_Value")
        If RS("Style_Name")="KEY_4" Then KEY_4=RS("Style_Value")
        If RS("Style_Name")="KEY_5" Then KEY_5=RS("Style_Value")
        If RS("Style_Name")="KEY_6" Then KEY_6=RS("Style_Value")
        If RS("Style_Name")="KEY_7" Then KEY_7=RS("Style_Value")
        If RS("Style_Name")="KEY_8" Then KEY_8=RS("Style_Value")
        If RS("Style_Name")="KEY_9" Then KEY_9=RS("Style_Value")
        RS.MoveNext
		  Wend
		  
		  Test_Limit=""
		  If KEY_3<>"" And KEY_4<>"" And KEY_5<>"" Then
		    For I=2009 To 2099
		      If MD5(Cstr(I))=Cstr(KEY_3) Then 
		        Test_Limit=Cstr(I)
		        Exit For
		      End If  
		    Next
		    For I=1 To 12
		      If MD5(Cstr(I))=Cstr(KEY_4) Then 
		        Test_Limit=Test_Limit&"/"&Cstr(I)
		        Exit For
		      End If
		    Next
		    For I=1 To 31
		      If MD5(Cstr(I))=Cstr(KEY_5) Then 
		        Test_Limit=Test_Limit&"/"&Cstr(I)
		        Exit For
		      End If
		    Next
		  End If
		  
		  Expire_Limit=""
		  If KEY_6<>"" And KEY_7<>"" And KEY_8<>"" Then
		    For I=2009 To 2099
		      If MD5(Cstr(I))=Cstr(KEY_6) Then 
		        Expire_Limit=Cstr(I)
		        Exit For
		      End If
		    Next
		    For I=1 To 12
          If MD5(Cstr(I))=Cstr(KEY_7) Then 
            Expire_Limit=Expire_Limit&"/"&Cstr(I)
		        Exit For
		      End If
		    Next
		    For I=1 To 31
          If MD5(Cstr(I))=Cstr(KEY_8) Then 
            Expire_Limit=Expire_Limit&"/"&Cstr(I)
		        Exit For
		      End If
		    Next
		  End If
		  Delay_Limit=0
		  For I=1 To 100
        If MD5(Cstr(I))=Cstr(KEY_9) Then 
          Delay_Limit=Cint(I)
		      Exit For
		    End If
		  Next		  

		  If MD5(InetpubPath)<>KEY_1 Or MD5(DriveID)<>KEY_2 Then Check_MD5=true
		  If Test_Limit="" And Expire_Limit="" Then Check_MD5=true
    Else
      Check_MD5=true
    End If
    
    If Check_MD5=Flase Then
      Session("errnumber")=1
      Session("msg")="License設定錯誤\n\n詳情請洽網軟公司"
      Response.Redirect "index.asp"
    Else
      'If Expire_Limit<>"" Then
      '  If Date()>CDate(DateAdd("D",Cint(Delay_Limit),Expire_Limit)) Then
      '    Session("errnumber")=1
      '    Session("msg")="網站代管期限期已經逾期\n\n至  "&CDate(Expire_Limit)&"  止\n\n請盡速聯絡網軟公司以免影響貴會權益"
      '    Response.Redirect "index.asp"
      '  End If
      'ElseIf Test_Limit<>"" Then
      '  If Date()>CDate(Test_Limit) Then
      '    Session("errnumber")=1
      '    Session("msg")="網站測試有效期限已經逾期\n\n至  "&CDate(Test_Limit)&"  止\n\n請盡速聯絡網軟公司以免影響貴會權益"
      '    Response.Redirect "index.asp"
      '  End If
      'End If 
    End If
        
    SQL="Select user_id,user_name,user_group,userfile.dept_id,comp_label,comp_name,comp_shortname,dept_desc,supervisor,pwd_validDate,password,sys_name,password_day,user_lock,pwd_validDate,BranchType,Client_IP,dept_type " & _
        "FROM userfile Join Dept On userfile.dept_id=Dept.dept_id Where user_id='"&Request("user_id")&"'"
    Set RS=Server.CreateObject("ADODB.Recordset")
    RS.Open SQL,Conn,1,1
    If Not RS.EOF Then
      If RS.Recordcount=1 Then
        '帳號是否封鎖
        If RS("user_lock") = "Y" Then
          Session("User_Lock")="Y"
          Call Sys_Log (Request("user_id"),"","","系統登入",Request("user_id")&"系統登入帳號已封鎖")
          Set RS=Nothing
          Conn.Close
          Set Conn=Nothing
          Session("errnumber")=1
          Session("msg")="帳號已封鎖 ！"
          Response.Redirect "index.asp"
        Else
          Session("User_Lock")="N"
        End If
        
        '帳號是否逾期
        If datediff("d",date(),RS("pwd_validDate"))<0 Then
          Session("User_Valid")="Y"
          Call Sys_Log (Request("user_id"),"","","系統登入",Request("user_id")&"系統登入帳號已逾期")
          Set RS=Nothing
          Conn.Close
          Set Conn=Nothing
          Session("errnumber")=1
          Session("msg")="帳號已逾期 ！"
          Response.Redirect "index.asp"
        Else
          Session("User_Valid")="N"
        End If	
        
        '密碼是否錯誤
        If RS("password") <> Request("password") Then
          Session("LoginErrTime")=Session("LoginErrTime")+1
          Call Sys_Log (Request("user_id"),"","","系統登入",Request("user_id")&"系統登入密碼錯誤")
          If Cint(Session("LoginErrTime"))>=6 Then Call User_Lock (Request("user_id"))
          RS.Close
          Set RS=Nothing
          Conn.Close
          Set Conn=Nothing
          Session("errnumber")=1
          Session("msg")="密碼 欄位輸入錯誤 ！"
          Response.Redirect "index.asp"
        End If
        
        'Check IP
        If Request("user_id")<>"npois" And Mid(login_return,1,1)="1" Then
          Check_IP="N"
          Ary_Client_IP=Split(RS("Client_IP"),";")
          Ary_Login_IP=Split(Request.ServerVariables("REMOTE_ADDR"),".")
          For I = 0 To Ubound(Ary_Client_IP)
            Client_IP=Split(Trim(Ary_Client_IP(I)),".")
            If Ary_Login_IP(0)=Client_IP(0) And Ary_Login_IP(1)=Client_IP(1) And Ary_Login_IP(2)=Client_IP(2) Then
              If Client_IP(3)="*" Or (Ary_Login_IP(3)=Client_IP(3)) Then
                Check_IP="Y"
                Exit For
              End If
            End If
          Next
          
          If Check_IP="N" Then
            Call Sys_Log (Request("user_id"),"","","系統登入",Request("user_id")&"系統登入IP錯誤")
            If Cint(Session("LoginErrTime"))>=6 Then Call User_Lock (Request("user_id"))
            RS.Close
            Set RS=Nothing
            Conn.Close
            Set Conn=Nothing
            Session("errnumber")=1
            Session("msg")="IP錯誤 ！"
            Response.Redirect "index.asp"
          End If
        End If
        
        'donate_return
        If Mid(login_return,2,1)="1" Then
          SQL1="Select * From DONATE_WEB Where Donate_Update='N' Order By Ser_No"
          Set RS1=Server.CreateObject("ADODB.Recordset")
          RS1.Open SQL1,Conn,1,3
          While Not RS1.EOF
            SQL2="Select * From DONOR Where Donor_Id='"&RS1("Donor_Id")&"'"
            Set RS2 = Server.CreateObject("ADODB.RecordSet")
            RS2.Open SQL2,Conn,1,3
            If Not RS2.EOF Then
              RS2("Donor_Name")=RS1("Donate_DonorName")
              RS2("Sex")=Request("Donate_Sex")
              RS2("IDNo")=RS1("Donate_IDNO")
              If RS1("Donate_Birthday")<>"" Then 
                RS2("Birthday")=RS1("Donate_Birthday")
              Else
                RS2("Birthday")=null
              End If
              RS2("Education")=RS1("Donate_Education")
              RS2("Occupation")=RS1("Donate_Occupation")
              RS2("Marriage")=RS1("Donate_Marriage")
              RS2("Cellular_Phone")=RS1("Donate_CellPhone")
              Tel_Office=RS1("Donate_TelOffice")
              If RS1("Donate_TelOffice_Region")<>"" Then Tel_Office="("&RS1("Donate_TelOffice_Region")&")"&Tel_Office
              If RS1("Donate_TelOffice_Ext")<>"" Then Tel_Office=Tel_Office&"#"&RS1("Donate_TelOffice_Ext")
              RS2("Tel_Office")=Tel_Office
              RS2("ZipCode")=RS1("Donate_ZipCode")  
              RS2("City")=RS1("Donate_CityCode")
              RS2("Area")=RS1("Donate_AreaCode")
              RS2("Address")=RS1("Donate_Address")
              RS2("Email")=RS1("Donate_Email")
              RS2("Invoice_type")=RS1("Donate_Invoice_Type")
              RS2("Invoice_Title")=RS1("Donate_Invoice_Title")
              RS2("Invoice_IDNo")=RS1("Donate_Invoice_IDNo")
              RS2("Invoice_ZipCode")=RS1("Donate_Invoice_ZipCode")
              RS2("Invoice_City")=RS1("Donate_Invoice_CityCode")
              RS2("Invoice_Area")=RS1("Donate_Invoice_AreaCode")
              RS2("Invoice_Address")=RS1("Donate_Invoice_Address")
              RS2.Update

              SQL3="Update DONOR Set Introducer_Name='"&RS1("Donate_DonorName")&"' Where Introducer_Id='"&RS1("Donor_Id")&"'"
              Set RS3=Conn.Execute(SQL3)

              SQL3="DECLARE @last_donatedate datetime " & _
                   "DECLARE @begin_donatedate datetime " & _
                   "DECLARE @donate_no numeric " & _
                   "DECLARE @donate_total numeric " & _
                   "Select Top 1 @last_donatedate=Donate_Date From DONATE Where Donor_Id='"&RS1("Donor_Id")&"' And Issue_Type<>'D' And Donate_Amt>0 Order By Donate_Date Desc " & _
                   "Select Top 1 @begin_donatedate=Donate_Date From DONATE Where Donor_Id='"&RS1("Donor_Id")&"' And Issue_Type<>'D' And Donate_Amt>0 Order By Donate_Date " & _
                   "Select @donate_no=Count(*) From DONATE Where Donor_Id='"&RS1("Donor_Id")&"' And Issue_Type<>'D' And Donate_Amt>0 " & _
                   "Select @donate_total=IsNull(Sum(Donate_Amt),0) From DONATE Where Donor_Id='"&RS1("Donor_Id")&"' And Issue_Type<>'D' And Donate_Amt>0 " & _
                   "Update DONOR Set Last_DonateDate=@last_donatedate,Begin_DonateDate=@begin_donatedate,Donate_No=@donate_no,Donate_Total=@donate_total Where Donor_Id='"&RS1("Donor_Id")&"'"
              Set RS3=Conn.Execute(SQL3)
            End If
            RS2.Close
            Set RS2=Nothing
            RS1("Donate_Update")="Y"
            RS1.Update
            RS1.MoveNext
		      Wend
          RS1.Close
          Set RS1=Nothing
        End If

        'epaper_return
        If Mid(login_return,3,1)="1" Then
          SQL1="Select * From EPAPER_RETURN Order By ePaper_Return_Id"
          Set RS1=Server.CreateObject("ADODB.Recordset")
          RS1.Open SQL1,Conn,1,1
          While Not RS1.EOF
            SQL2="Select ePaper_Return From EPAPER_USER Where Ser_No='"&RS1("ePaper_Return_Id")&"'"
            Set RS2 = Server.CreateObject("ADODB.RecordSet")
            RS2.Open SQL2,Conn,1,3
            If Not RS2.EOF Then
              RS2("ePaper_Return")="Y"
              RS2.Update
            End If
            RS2.Close
            Set RS2=Nothing
            SQL3="Delete EPAPER_RETURN Where ePaper_Return_Id='"&RS1("ePaper_Return_Id")&"'"
            Set RS3=Conn.Execute(SQL3)
            RS1.MoveNext
		      Wend
          RS1.Close
          Set RS1=Nothing
          
          SQL1="Select * From EPAPER_READ Where Read_Type='E' Order By Read_RegDate"      
          Set RS1=Server.CreateObject("ADODB.Recordset")
          RS1.Open SQL1,Conn,1,1
          While Not RS1.EOF  
            If RS1("Read_UserID")<>"" Then
              SQL2="Select * From EPAPER_USER Where Ser_No='"&RS1("Read_UserID")&"'"
            Else
              SQL2="Select * From EPAPER_USER Where ePaper_Email='"&RS1("Read_EMail")&"'"
            End If
            Set RS2 = Server.CreateObject("ADODB.RecordSet")
            RS2.Open SQL2,Conn,1,3
            If Not RS2.EOF Then
              RS2("Last_ReadNo")=RS1("Read_ePaperNo")
              RS2("Last_ReadDate")=RS1("Read_RegDate")
              RS2.Update
            End If
            RS2.Close
            Set RS2=Nothing
            SQL3="Delete EPAPER_READ Where Read_ePaperNo='"&RS1("Read_ePaperNo")&"' And Read_AddIP='"&RS1("Read_AddIP")&"' And Read_Type='E' "
            If RS1("Read_UserID")<>"" Then SQL3=SQL3&"And Read_UserID='"&RS1("Read_UserID")&"' "
            If RS1("Read_EMail")<>"" Then SQL3=SQL3&"And Read_EMail='"&RS1("Read_EMail")&"' "
            Set RS3=Conn.Execute(SQL3)
            RS1.MoveNext
		      Wend
          RS1.Close
          Set RS1=Nothing
          
          SQL1="Select * From EPAPER_READ Where Read_Type='C' Order By Read_RegDate"      
          Set RS1=Server.CreateObject("ADODB.Recordset")
          RS1.Open SQL1,Conn,1,1
          While Not RS1.EOF  
            If RS1("Read_UserID")<>"" Then
              SQL2="Select * From MEMBER Where Ser_No='"&RS1("Read_UserID")&"'"
            Else
              SQL2="Select * From MEMBER Where Member_EMail='"&RS1("Read_EMail")&"'"
            End If
            Set RS2 = Server.CreateObject("ADODB.RecordSet")
            RS2.Open SQL2,Conn,1,3
            If Not RS2.EOF Then
              RS2("Last_ReadNo")=RS1("Read_ePaperNo")
              RS2("Last_ReadDate")=RS1("Read_RegDate")
              RS2.Update
            End If
            RS2.Close
            Set RS2=Nothing
            SQL3="Delete EPAPER_READ Where Read_ePaperNo='"&RS1("Read_ePaperNo")&"' And Read_AddIP='"&RS1("Read_AddIP")&"' And Read_Type='C' And (Read_UserID='"&RS1("Read_UserID")&"' Or Read_EMail='"&RS1("Read_EMail")&"') "
            Set RS3=Conn.Execute(SQL3)
            RS1.MoveNext
		      Wend
          RS1.Close
          Set RS1=Nothing
        End If

        'info_return
        If Mid(login_return,4,1)="1" Then
          SQL1="Select * From INFO_READ Order By Read_RegDate"      
          Set RS1=Server.CreateObject("ADODB.Recordset")
          RS1.Open SQL1,Conn,1,1
          While Not RS1.EOF  
            If RS1("Read_MemberNo")<>"" Then
              SQL2="Select * From MEMBER Where Ser_No='"&RS1("Read_MemberNo")&"'"
            Else
              SQL2="Select * From MEMBER Where Member_EMail='"&RS1("Read_EMail")&"'"
            End If
            Set RS2 = Server.CreateObject("ADODB.RecordSet")
            RS2.Open SQL2,Conn,1,3
            If Not RS2.EOF Then
              RS2("Last_Info_ReadNo")=RS1("Read_InfoNo")
              RS2("Last_Info_ReadDate")=RS1("Read_RegDate")
              RS2.Update
            End If
            RS2.Close
            Set RS2=Nothing
            SQL3="Delete INFO_READ Where Read_InfoNo='"&RS1("Read_InfoNo")&"' And Read_AddIP='"&RS1("Read_AddIP")&"' And (Read_MemberNo='"&RS1("Read_MemberNo")&"' Or Read_EMail='"&RS1("Read_EMail")&"') "
            Set RS3=Conn.Execute(SQL3)
            RS1.MoveNext
		      Wend
          RS1.Close
          Set RS1=Nothing
        End If

        'member_return
        If Mid(login_return,5,1)="1" Then
          SQL1="Select * From MEMBER_LIST Order By Member_DateTime"
          Set RS1=Server.CreateObject("ADODB.Recordset")
          RS1.Open SQL1,Conn,1,1
          If Not RS1.EOF Then
            SQL2="Select * From MEMBE Where Ser_No='"&RS1("Member_id")&"'"
            Set RS2 = Server.CreateObject("ADODB.RecordSet")
            RS2.Open SQL2,Conn,1,3
            If Not RS2.EOF Then
              RS2("Member_Id")=RS1("Member_Id")
              RS2.Update
            End If
            RS2.Close
            Set RS2=Nothing
            SQL3="Delete From MEMBER_LIST Where od_sob='"&RS1("od_sob")&"' "
            Set RS3=Conn.Execute(SQL3) 
          End If
          RS1.Close
          Set RS1=Nothing
        End If

        'signup_return
        If Mid(login_return,6,1)="1" Then

        End If

        'Delete Shopping Car
        If Mid(login_return,7,1)="1" Then
          SQL1="Delete DONATE_SHOPPING_CART Where Create_Date<='"&DateAdd("D",-2,Date())&"'"
          Set RS1=Conn.Execute(SQL1)
        End If

        'Delete DONATE_OD_SOB
        If Mid(login_return,7,1)="1" Then
          SQL1="Delete DONATE_OD_SOB Where Create_Date<='"&DateAdd("D",-30,Date())&"'"
          Set RS1=Conn.Execute(SQL1)
        End If
     
        '登入成功
        Call Sys_Log (Request("user_id"),RS("user_name"),RS("dept_id"),"系統登入",Request("user_id")&"系統登入成功")
        Session("LoginErrTime")=0
        Session("user_id")=RS("user_id")
        Session("menu_user_id")=RS("user_id")
        Session("user_name")=RS("user_name")
        Session("user_group")=RS("user_group")
        Session("sys_name")=RS("sys_name")
        Session("dept_id")=RS("dept_id")
        Session("dept_id_login")=RS("dept_id")
        Session("menu_dept_id")=RS("dept_id")
        Session("dept_desc")=RS("dept_desc")
        Session("comp_label")=RS("comp_label")
        Session("comp_name")=RS("comp_name")
        Session("comp_ShortName")=RS("comp_ShortName")
        Session("supervisor")=RS("supervisor")
        Session("dept_type")=RS("dept_type")
        Session("all_dept_type")="'"&Replace(RS("dept_type"),", ","','")&"'"
        pwd_validDate=CDate(RS("pwd_validDate"))
        If RS("password_day")<>"" Then
          password_day=Cint(RS("password_day"))
        Else
          password_day=0
        End If
        Session("BranchType")=RS("BranchType")
        If session("menu_lock")="N" Then
          Session("Menu_Edit")="Y"
        Else
          If Session("user_id")="npois" And (Request.ServerVariables("REMOTE_ADDR")="127.0.0.1" Or Request.ServerVariables("REMOTE_ADDR")="60.250.147.33") Then
            Session("Menu_Edit")="Y"
          Else
            Session("Menu_Edit")="N"
          End If        
        End If
        Session("msg")=""
        Session("errnumber")=0
        
        RS.Close
        Set RS=Nothing
        Conn.Close
        Set Conn=Nothing
        Response.Redirect "../sysmgr/main.asp"
      Else
        Session("LoginErrTime")=Session("LoginErrTime")+1
        RS.Close
        Set RS=Nothing
        Conn.Close
        Set Conn=Nothing
        Session("errnumber")=1
        Session("msg")="帳號/密碼 輸入錯誤 ！"
        Response.Redirect "index.asp"
      End If
    Else
      Session("LoginErrTime")=Session("LoginErrTime")+1
      RS.Close
      Set RS=Nothing
      Conn.Close
      Set Conn=Nothing
      Session("errnumber")=1
      Session("msg")="帳號/密碼 輸入錯誤 ！"
      Response.Redirect "index.asp"
    End If  
  End If
End If

Function Sys_Log (User_ID,User_Name,Dept_Id,Log_Type,Log_Desc)
  On Error Resume Next
  SQL1="sys_log"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  RS1.Addnew
  RS1("Log_Time")=now()
  RS1("Log_Date")=date()
  RS1("User_Id")=User_ID
  RS1("User_Name")=User_Name
  RS1("Dept_Id")=Dept_Id
  RS1("Log_Type")=Log_Type
  RS1("Log_Desc")=Log_Desc
  RS1("IP_Address")=Request.ServerVariables("REMOTE_ADDR")
  RS1.Update
  RS1.Close
  Set RS1=Nothing
end Function

Function User_Lock (User_ID)
  On Error Resume Next
  SQL1="Select * From userfile Where user_id='"&User_ID&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  If Not RS.EOF Then
    RS1("user_lock")="Y"
    RS1.Update
  End If
  RS1.Close
  Set RS1=Nothing
End Function
%>