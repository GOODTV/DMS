<%
Class IMAGE
	'公用存取變數
	Dim dataFile		'計數器資料檔
	Dim imgType			'輸出圖形類型
	Dim jpgQuality	'品質 0~100
	'顏色屬性
	Dim fgColor			'前景色
	Dim bgColor			'背景色
	Dim bdColor			'邊框色
	Dim transparent	'通透色
	'字型屬性
	Dim fontSize		'字型大小
	Dim fontColor		'字型顏色
	Dim fontFace		'字體
	Dim fontBold		'粗體
	Dim fontItalic	'斜體
	'寬高
	Dim width				'圖寬
	Dim height			'圖高
	Dim width_plus	'加寬量
	Dim height_plus	'加高量
	Dim bdSize			'邊框粗細

	'內部存取屬性,方法
	Private img			'內部用物件變數


  '開始時執行的工作
  Private Sub Class_Initialize()
		dataFile = "./counter.dat"
		imgType = 2
		jpgQuality = 100
		fontSize = 10
		fontColor = "#000000"
		bgColor = "#FFFFFF"
		fgColor = "#FFFFFF"
		bdSize = 0
		width_plus = 0
		height_plus = 0
		transparent = False
		set img = server.createobject("Overpower.ImageLib") 'Create一個物件 
  End Sub

  '結束時執行的工作
  Private Sub Class_Terminate()
		'Nothing to do
		Set img = Nothing
  End Sub  

	'產生數字
	Function Counter()		
		Dim cnt
		cnt = CStr(Read_Cnt_Data())
		ImageText(cnt)
	End Function

	'圖形字
	Function ImageText(char)
		Dim imgwidth, imgheight
    SetFont() 
		imgwidth = img.GetTextWidth(char)+width_plus
		imgheight = img.GetTextHeight(char)+height_plus

		img.width = imgwidth '設定此物件的寬度 
		img.height = imgheight '設定此物件的高度 

		img.BrushColor = bgColor
		img.FillRect 0, 0, imgheight, imgwidth
		'(x1,y1) - (y2,x2)

		If bdColor<>Empty Then
			img.PenColor = bdColor
			img.PenWidth = bdSize
			pQty=1
			If bdSize Mod 2 = 1 Then pQty=0
			img.Box Int(bdSize/2), Int(bdSize/2), imgwidth-Int(bdSize/2)+pQty, imgheight-Int(bdSize/2)+pQty
			'(x1,y1) - (x2,y2)
		End If

		AddText char, width_plus/2, height_plus/2

		'文字, X, Y
		WriteImage("")
	End Function
 

	'顯示圖片
	Function ShowImage(imgFile, w, h)
		Dim imgwidth, imgheight
		'set img = server.createobject("Overpower.ImageLib") 'Create一個物件 
		GetImageSize(Server.MapPath(imgFile))

		imgwidth = w
		imgheight = height * w/width
		If h<>Empty Then imgheight=h
                
                If w<>"" Then 
		  imgwidth = w
		  imgheight = height * w/width
		  If h<>Empty Then imgheight=h
		Else
		  imgheight = h
		  imgwidth = width * h/height
		  If w<>Empty Then imgheight=w
		End If				
		img.width = imgwidth '設定此物件的寬度 
		img.height = imgheight '設定此物件的高度 

		FileName=Server.MapPath(imgFile)
		img.InsertPicture Filename, 0, 0, True, imgwidth, imgheight

		SetFont() 
		AddText text, 2, 2

		WriteImage("")
	End Function

	'轉換圖檔
	Function TransImage(sf, df, w, h)
		Dim imgwidth, imgheight
		GetImageSize(Server.MapPath(sf))

                If w<>"" Then 
		  imgwidth = w
		  imgheight = height * w/width
		  If h<>Empty Then imgheight=h
		Else
		  imgheight = h
		  imgwidth = width * h/height
		  If w<>Empty Then imgheight=w
		End If
		img.width = imgwidth '設定此物件的寬度 
		img.height = imgheight '設定此物件的高度 

		FileName=Server.MapPath(sf)
		img.InsertPicture Filename, 0, 0, True, imgwidth, imgheight

		WriteImage(Server.MapPath(df))
	End Function
	
	'輸出或存檔
	Private Function WriteImage(df)
		Dim transColor
		transColor = ""
	  If transparent<>Empty Then
			transColor = bgColor
		End If

		If df<>Empty Then
			imgType = Right(df, 3)
			ImageType()
			img.SavePicture df, imgType, jpgQuality, transColor '直接輸出圖檔, 不存檔 
		Else
			ImageType()
			img.PictureBinaryWrite imgType, jpgQuality, transColor '直接輸出圖檔, 不存檔 
			'物件.方法 檔案類型, 壓縮品質, 省略通透顏色值 
		End IF

	End Function

	
	'加文字
	Private Function AddText(char, x, y)
	  img.FontColor=fontColor
		img.TextOut char, x, y
	End Function
	
	'轉換圖片型態成數字
	Private Function ImageType()
    If LCase(imgType)="bmp" Then 
			imgType=1
		ElseIf LCase(imgType)="gif" Then 
			imgType=2
		ElseIf LCase(imgType)="jpg" Then 
			imgType=3
		End If
		ImageType = imgType
	End Function

  '設定字型屬性
	Private Function setFont
		img.fontSize   = fontSize
		img.fontColor  = fontColor
		img.fontFace   = fontFace
		img.fontBold   = fontBold
		img.fontItalic = fontItalic
	End Function
	
	'讀數字&累加
	Private Function Read_Cnt_Data()
			Dim fs, txtf, fileName, cnt
			fileName = Server.MapPath(dataFile)
			Set fs = Server.CreateObject("Scripting.FileSystemObject")
			
			Set txtf = fs.OpenTextFile(fileName, 1, True)
			If Not txtf.atEndOfStream Then cnt = txtf.ReadLine
			txtf.Close

			cnt = Clng(cnt)+1

			Set txtf = fs.OpenTextFile(fileName, 2, True)
			txtf.WriteLine cnt
			txtf.Close

			Set fs = Nothing
			Read_Cnt_Data = cnt
	End Function

	'取得圖檔尺寸
	Private Function GetImageSize(imgFile)
		img.PictureSize imgFile, width, height
	End Function
End Class
%>