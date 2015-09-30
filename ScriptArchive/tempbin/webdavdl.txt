server = "servername"
mailbox = "mailbox"
set fso = createobject("Scripting.FileSystemObject")
strURL = "http://" & server & "/exchange/" & mailbox & "/inbox/"
strURL1 = "http://" & server & "/exchange/" & mailbox & "/sent items/"
strQuery = "<?xml version=""1.0""?><D:searchrequest xmlns:D = ""DAV:"" >"
strQuery = strQuery & "<D:sql>SELECT  ""DAV:displayname"", ""urn:schemas:httpmail:subject"""
strQuery = strQuery & " FROM scope('shallow traversal of """
strQuery = strQuery & strURL & """') Where ""DAV:ishidden"" = False AND ""DAV:isfolder"" = False AND "
strQuery = strQuery & """urn:schemas:httpmail:read"" = false AND "
strQuery = strQuery & """urn:schemas:httpmail:hasattachment"" = True </D:sql></D:searchrequest>"
set req = createobject("microsoft.xmlhttp")
req.open "SEARCH", strURL, false
req.setrequestheader "Content-Type", "text/xml"
req.setRequestHeader "Translate","f"
req.send strQuery
If req.status >= 500 Then
ElseIf req.status = 207 Then
   set oResponseDoc = req.responseXML
   set oNodeList = oResponseDoc.getElementsByTagName("a:displayname")
   set oNodeList1 = oResponseDoc.getElementsByTagName("a:href")
   For i = 0 To (oNodeList.length -1)
	set oNode = oNodeList.nextNode
	set oNode1 = oNodeList1.nextNode
	wscript.echo oNode.Text
	embedattach(oNode1.Text)
	updateunread(oNode1.Text)
   Next	
Else
End If


function embedattach(objhref)
req.open "X-MS-ENUMATTS", objhref, false, "", ""
req.send
If req.status > 207 Or req.status < 207 Then
    wscript.echo "Status: " & req.status
    wscript.echo "Status text: " & req.statustext
    wscript.echo "Response text: " & req.responsetext
Else
    wscript.echo ""
    wscript.echo "Attachment"
    set resDoc1 = req.responseXML
    Set objPropstatNodeList1 = resDoc1.getElementsByTagName("a:propstat")
    Set objHrefNodeList1 = resDoc1.getElementsByTagName("a:href")
    If objPropstatNodeList1.length > 0 Then
         wscript.echo objPropstatNodeList1.length & " attachments found..."
    For f = 0 To (objPropstatNodeList1.length -1)
        set objPropstatNode1 = objPropstatNodeList1.nextNode
        set objHrefNode1 = objHrefNodeList1.nextNode
        wscript.echo "Attachment: " &  objHrefNode1.Text
        set objNodef = objPropstatNode1.selectSingleNode("a:prop/d:x37050003")
        wscript.echo "Attachment Method: " & objNodef.Text
        set objNodef2 = objPropstatNode1.selectSingleNode("a:prop/f:cn")
        wscript.echo "CN: " & objNodef2.Text
        if objNodef.Text = 5 then
            embedattach(objHrefNode1.Text)
       else
            set objNode1f = objPropstatNode1.selectSingleNode("a:prop/d:x3704001f")
            wscript.echo "Attachment name: " & objNode1f.Text
            req.open "GET", objHrefNode1.Text, false, "", ""
	    req.send
	    set stm = createobject("ADODB.Stream")
	    stm.open
            msgstring = req.responsetext
	    stm.type = 2
	    stm.Charset = "x-ansi"
	    stm.writetext msgstring,0
	    stm.Position = 0
	    stm.type = 1
	    stm.savetofile "c:\temp\" & objNode1f.Text,2
	    set stm = nothing
       end if
    next
Else
     wscript.echo "No file attachments found..."
End If
End If
wscript.echo 
end function

function updateunread(objhref)
req.open "PROPPATCH", objhref, False
xmlstr = "<?xml version=""1.0"" encoding=""UTF-8"" ?>"
xmlstr = xmlstr & "<a:propertyupdate xmlns:a=""DAV:"" xmlns:dt=""urn:uuid:c2f41010-65b3-11d1-a29f-00aa00c14882/"" xmlns:d=""urn:schemas:httpmail:"">" 
xmlstr = xmlstr &  "<a:set>"
xmlstr = xmlstr &  "<a:prop>" 
xmlstr = xmlstr &  "<d:read dt:dt=""boolean"">1</d:read>" 
xmlstr = xmlstr &  "</a:prop>" 
xmlstr = xmlstr &  "</a:set>" 
xmlstr = xmlstr &  "</a:propertyupdate>" 
req.setRequestHeader "Content-Type", "text/xml;"
req.setRequestHeader "Translate", "f"
req.setRequestHeader "Content-Length:", Len(xmlstr)
req.send(xmlstr)
wscript.echo req.status
end function
