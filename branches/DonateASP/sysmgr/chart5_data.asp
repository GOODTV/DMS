<%@ language="vbscript" %>
<%
Set cd = CreateObject("ChartDirector.API")
year_now=year(Date)

' The currently selected year
selectedYear = Request("year")
If Request("year") = "" Then
    selectedYear = year_now
Else
    selectedYear = CInt(Request("year"))
End If

'
' The following code generates the <option> tags for the HTML select box, with the
' <option> tag representing the currently selected year marked as selected.
'

ReDim optionTags(year_now - 1999)
For i = 1999 To year_now
    If i = selectedYear Then
        optionTags(i - 1999) = "<option value='" & i & "' selected>" & i & _
            "</option>"
    Else
        optionTags(i - 1999) = "<option value='" & i & "'>" & i & "</option>"
    End If
Next
selectYearOptions = Join(optionTags, "")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<body style="margin:5px 0px 0px 5px">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
 <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<form>
    <font size="2">捐款年度：</font>
    <select name="year">
        <%=selectYearOptions%>
    </select>
    <input type="submit" value="查詢">
</form>
</div>

<img src="chart5.asp?year=<%=selectedYear%>">

</body>
</html>