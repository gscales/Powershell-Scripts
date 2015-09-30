dim objspcarray(9)
dim objlabelarray(9)
set rec = createobject("ADODB.Record")
Set cnvt = CreateObject("ADs.ArrayConvert")
objspcarray(0) = "0000"
objspcarray(1) = "0000"
objspcarray(2) = "0000"
objspcarray(3) = "0000"
objspcarray(4) = "0000"
objspcarray(5) = "0000"
objspcarray(6) = "0000"
objspcarray(7) = "0000"
objspcarray(8) = "0000"
objspcarray(9) = "0000"
objlabelarray(0) = ""
objlabelarray(1) = "External"
objlabelarray(2) = "Internal"
objlabelarray(3) = ""
objlabelarray(4) = ""
objlabelarray(5) = ""
objlabelarray(6) = ""
objlabelarray(7) = ""
objlabelarray(8) = ""
objlabelarray(9) = ""
dstring = "0000"
for i = lbound(objspcarray) to ubound(objspcarray)
	if objlabelarray(i) <> "" then
		objtooct = cnvt.CvStr2vOctetStr(objlabelarray(i))
		objtohex = cnvt.CvOctetStr2vHexStr(objtooct)
		for h = 1 to len(objtohex)/2
			dstring = dstring & mid(objtohex,((h*2)-1),2) & "00"
		next
	end if
	dstring = dstring & objspcarray(i)
next
set convobj = CreateObject("Msxml2.DOMDocument.4.0")
Set oRoot = convobj.createElement("test")
oRoot.dataType = "bin.base64"
oRoot.nodeTypedValue = cnvt.CvHexStr2vOctetStr(dstring)
wscript.echo oRoot.text


