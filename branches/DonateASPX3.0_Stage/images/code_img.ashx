<%@ WebHandler Language="VB" Class="code_img" %>

Imports System
Imports System.Web
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Web.SessionState


Public Class code_img : Implements IHttpHandler, IRequiresSessionState
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        
        context.Response.ContentType = "image/Png"

        'RndNum是一個自定義函數
        Dim VNum As String = RndNum(4)
        'Session(sessionName) = VNum
        context.Session("CheckCode") = VNum
        context.Response.BinaryWrite(ValidateCode(VNum).ToArray)
        'context.Response.Write("Hello World")
        
    End Sub
    
    
    '生成圖像驗證碼函數 
    Function ValidateCode(ByVal VNum As Object) As MemoryStream
        
        Dim Img As System.Drawing.Bitmap
        Dim g As Graphics
        Dim ms As MemoryStream
        'Dim gwidth As Integer = Int(Len(VNum) * 18)
        'Dim gwidth As Integer = 80
        Dim gwidth As Integer = 70
        'gheight為圖片寬度,根據字符長度自動更改圖片寬度 
        Img = New Bitmap(gwidth, 25)
        g = Graphics.FromImage(Img)
        'g.Clear(Color.Black) '背景
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        'g.DrawRectangle (New Pen(Color.gray ),new Rectangle (0,0,Img.Width-1 ,Img.Height-1 ))
        g.DrawString(VNum, (New Font("Arial Black", 12)), (New SolidBrush(Color.Black)), 3, 3) '在矩形內繪製字串（字串，字體，畫筆顏色，左上x.左上y） 
        ms = New MemoryStream()
        Img.Save(ms, ImageFormat.Png)
        'HttpResponse.ClearContent() '需要輸出圖像信息 要修改HTTP頭 
        'Response.ContentType = "image/Png"
        'Response.BinaryWrite(ms.ToArray())
        'g.Dispose()
        'Img.Dispose()
        'Response.End()
        
        Return ms
        
    End Function
    
    
    '-------------------------------------------- 
    '函數名稱:RndNum 
    '函數參數:VcodeNum--設定返回隨機字符串的位數 
    '函數功能:產生數字和字符混合的隨機字符串 
    Function RndNum(ByVal VcodeNum As Integer) As String
        Dim Vchar As String = "0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,W,X,Y,Z"
        Dim VcArray() As String = Split(Vchar, ",") '將字符串生成數組 
        Dim VNum As String = ""
        Dim i As Byte
        For i = 1 To VcodeNum
            Randomize()
            VNum = VNum & VcArray(Int(35 * Rnd())) '數組一般從0開始讀取，所以這裏為35*Rnd 
        Next
        Return VNum
    End Function
    
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class