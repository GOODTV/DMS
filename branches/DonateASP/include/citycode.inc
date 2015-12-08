<Script Language="VBScript">
sub xaddOption(xlist,name,value)
	set xOption = document.createElement("OPTION")
	xOption.text=name
	xOption.value=value
	xlist.add(xOption)
end sub

sub removeOption(xlist)
	for i=xlist.options.length-1 to 0 step -1
		xlist.options.remove(i)
	next
	xlist.selectedIndex = -1
end sub
sub city_onChange
	set xsrc = document.all("area")
	removeOption(xsrc)

	set oXML = createObject("Microsoft.XMLDOM")
	oXML.async = false

	xURI = "../include/cityDiv.asp?cityCode=Addr0R" & document.all("city").value
	oXML.load(xURI)

	xaddOption xsrc, "區域", ""

	set pckItemList = oXML.selectNodes("divList/row")
	for each pckItem in pckItemList
		xaddOption xsrc, pckItem.selectSingleNode("mValue").text, pckItem.selectSingleNode("mCode").text
	next	
	xsrc.selectedIndex = 0
end sub

sub area_onChange
	document.all("ZipCode").value = left(document.all("area").value,3)
end sub
</Script>