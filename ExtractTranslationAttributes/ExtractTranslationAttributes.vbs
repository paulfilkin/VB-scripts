Public Sub ExtractSDLXLIFFNamespaceAware()
    Dim fName As Variant
    Dim xmlDoc As Object
    Dim ws As Worksheet
    Dim row As Long
    
    ' Pick file
    fName = Application.GetOpenFilename( _
                "XLIFF Files (*.sdlxliff;*.xliff),*.sdlxliff;*.xliff", , _
                "Select SDLXLIFF file")
    If fName = False Then Exit Sub
    
    ' Create XML document
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    xmlDoc.resolveExternals = False
    xmlDoc.preserveWhiteSpace = True
    
    ' Load file
    If Not xmlDoc.Load(fName) Then
        MsgBox "Error loading file: " & xmlDoc.parseError.reason
        Exit Sub
    End If
    
    ' Set up namespaces for XPath
    xmlDoc.SetProperty "SelectionLanguage", "XPath"
    xmlDoc.SetProperty "SelectionNamespaces", _
        "xmlns:xliff='urn:oasis:names:tc:xliff:document:1.2' " & _
        "xmlns:sdl='http://sdl.com/FileTypes/SdlXliff/1.0'"
    
    ' First, let's try to find trans-units with XPath
    Dim transUnits As Object
    
    ' Try different XPath expressions
    On Error Resume Next
    Set transUnits = xmlDoc.SelectNodes("//xliff:trans-unit")
    If transUnits Is Nothing Or transUnits.Length = 0 Then
        Set transUnits = xmlDoc.SelectNodes("//trans-unit")
    End If
    If transUnits Is Nothing Or transUnits.Length = 0 Then
        Set transUnits = xmlDoc.SelectNodes("//*[local-name()='trans-unit']")
    End If
    On Error GoTo 0
    
    ' If XPath fails, try getElementsByTagName
    If transUnits Is Nothing Or transUnits.Length = 0 Then
        Set transUnits = xmlDoc.getElementsByTagName("trans-unit")
    End If
    
    If transUnits Is Nothing Or transUnits.Length = 0 Then
        MsgBox "No trans-units found in file. The file may not be a valid SDLXLIFF.", vbExclamation
        Exit Sub
    End If
    
    ' Load comment definitions using XPath
    Dim commentDict As Object
    Set commentDict = CreateObject("Scripting.Dictionary")
    
    Dim cmtDefs As Object
    Set cmtDefs = xmlDoc.SelectNodes("//xliff:doc-info/xliff:cmt-defs/xliff:cmt-def")
    If cmtDefs Is Nothing Or cmtDefs.Length = 0 Then
        Set cmtDefs = xmlDoc.SelectNodes("//doc-info/cmt-defs/cmt-def")
    End If
    If cmtDefs Is Nothing Or cmtDefs.Length = 0 Then
        Set cmtDefs = xmlDoc.SelectNodes("//*[local-name()='cmt-def']")
    End If
    
    If Not cmtDefs Is Nothing Then
        Dim i As Long
        For i = 0 To cmtDefs.Length - 1
            Dim cmtDef As Object
            Set cmtDef = cmtDefs.Item(i)
            Dim cmtId As String
            cmtId = GetAttributeValue(cmtDef, "id")
            
            If cmtId <> "" Then
                Dim commentInfo As Object
                Set commentInfo = CreateObject("Scripting.Dictionary")
                
                ' Look for Comments/Comment structure
                Dim commentsNode As Object
                Set commentsNode = cmtDef.SelectSingleNode("sdl:Comments")
                If commentsNode Is Nothing Then
                    Set commentsNode = cmtDef.SelectSingleNode("Comments")
                End If
                If commentsNode Is Nothing Then
                    Set commentsNode = cmtDef.SelectSingleNode(".//*[local-name()='Comments']")
                End If
                
                If Not commentsNode Is Nothing Then
                    ' Now get the Comment child
                    Dim commentNode As Object
                    Set commentNode = commentsNode.SelectSingleNode("sdl:Comment")
                    If commentNode Is Nothing Then
                        Set commentNode = commentsNode.SelectSingleNode("Comment")
                    End If
                    If commentNode Is Nothing Then
                        Set commentNode = commentsNode.SelectSingleNode(".//*[local-name()='Comment']")
                    End If
                    
                    If Not commentNode Is Nothing Then
                        commentInfo.Add "text", commentNode.text
                        commentInfo.Add "user", GetAttributeValue(commentNode, "user")
                        commentInfo.Add "date", GetAttributeValue(commentNode, "date")
                        commentInfo.Add "severity", GetAttributeValue(commentNode, "severity")
                        
                        commentDict.Add cmtId, commentInfo
                    End If
                End If
            End If
        Next i
    End If
    
    ' Load revision definitions
    Dim revDict As Object
    Set revDict = CreateObject("Scripting.Dictionary")
    
    Dim revDefs As Object
    Set revDefs = xmlDoc.SelectNodes("//xliff:doc-info/xliff:rev-defs/xliff:rev-def")
    If revDefs Is Nothing Or revDefs.Length = 0 Then
        Set revDefs = xmlDoc.SelectNodes("//doc-info/rev-defs/rev-def")
    End If
    If revDefs Is Nothing Or revDefs.Length = 0 Then
        Set revDefs = xmlDoc.SelectNodes("//*[local-name()='rev-def']")
    End If
    
    If Not revDefs Is Nothing Then
        For i = 0 To revDefs.Length - 1
            Dim revDef As Object
            Set revDef = revDefs.Item(i)
            Dim revId As String
            revId = GetAttributeValue(revDef, "id")
            
            If revId <> "" Then
                Dim revInfo As Object
                Set revInfo = CreateObject("Scripting.Dictionary")
                
                revInfo.Add "author", GetAttributeValue(revDef, "author")
                revInfo.Add "date", GetAttributeValue(revDef, "date")
                revInfo.Add "type", GetAttributeValue(revDef, "type")
                
                revDict.Add revId, revInfo
            End If
        Next i
    End If
    
    ' Create extraction sheet
    Application.ScreenUpdating = False
    Set ws = Worksheets.Add
    ws.Name = "SDLXLIFF_" & Format(Now, "hhmmss")
    
    ' Headers
    Dim headers As Variant
    headers = Array("Trans-Unit ID", "Source Text", "Target Text", "Segment ID", _
                   "Status", "Confirmation Level", "Origin", "Origin System", _
                   "Percent Match", "Text Match", "Struct Match", "Context Match", _
                   "Created Date", "Created By", "Modified Date", "Modified By", _
                   "Comment Text", "Comment Author", "Comment Date", "Comment Severity", _
                   "Has Track Changes", "Revision Author", "Revision Date", _
                   "Deleted Text", "Added Text", "Locked")
    
    ' Set headers
    Dim col As Long
    For col = 1 To UBound(headers) + 1
        ws.Cells(1, col).Value = headers(col - 1)
    Next col
    ws.Range("A1:Y1").Font.Bold = True
    
    row = 2
    
    ' Process each trans-unit
    For i = 0 To transUnits.Length - 1
        Dim transUnit As Object
        Set transUnit = transUnits.Item(i)
        
        ' Skip if translate="no"
        If GetAttributeValue(transUnit, "translate") <> "no" Then
            
            ' Trans-unit ID
            ws.Cells(row, 1).Value = GetAttributeValue(transUnit, "id")
            
            ' Get source text
            Dim sourceNode As Object
            Set sourceNode = transUnit.SelectSingleNode("xliff:seg-source")
            If sourceNode Is Nothing Then
                Set sourceNode = transUnit.SelectSingleNode("seg-source")
            End If
            If sourceNode Is Nothing Then
                Set sourceNode = transUnit.SelectSingleNode("xliff:source")
            End If
            If sourceNode Is Nothing Then
                Set sourceNode = transUnit.SelectSingleNode("source")
            End If
            
            If Not sourceNode Is Nothing Then
                ws.Cells(row, 2).Value = GetCleanText(sourceNode)
            End If
            
            ' Get target text
            Dim targetNode As Object
            Set targetNode = transUnit.SelectSingleNode("xliff:target")
            If targetNode Is Nothing Then
                Set targetNode = transUnit.SelectSingleNode("target")
            End If
            
            If Not targetNode Is Nothing Then
                ws.Cells(row, 3).Value = GetCleanText(targetNode)
            End If
            
            ' Get segment info
            Dim segNode As Object
            Set segNode = transUnit.SelectSingleNode(".//sdl:seg")
            If segNode Is Nothing Then
                Set segNode = transUnit.SelectSingleNode(".//*[local-name()='seg'][@conf]")
            End If
            
            If Not segNode Is Nothing Then
                ws.Cells(row, 4).Value = GetAttributeValue(segNode, "id")
                ws.Cells(row, 5).Value = GetAttributeValue(segNode, "conf")
                
                ' Translate confirmation level
                Dim conf As String
                conf = GetAttributeValue(segNode, "conf")
                Select Case conf
                    Case "Draft": ws.Cells(row, 6).Value = "Draft"
                    Case "Translated": ws.Cells(row, 6).Value = "Translated"
                    Case "RejectedTranslation": ws.Cells(row, 6).Value = "Rejected"
                    Case "ApprovedTranslation": ws.Cells(row, 6).Value = "Approved"
                    Case "RejectedSignOff": ws.Cells(row, 6).Value = "Sign-off Rejected"
                    Case "ApprovedSignOff": ws.Cells(row, 6).Value = "Sign-off Approved"
                    Case Else: ws.Cells(row, 6).Value = conf
                End Select
                
                ws.Cells(row, 7).Value = GetAttributeValue(segNode, "origin")
                ws.Cells(row, 8).Value = GetAttributeValue(segNode, "origin-system")
                ws.Cells(row, 9).Value = GetAttributeValue(segNode, "percent")
                ws.Cells(row, 10).Value = GetAttributeValue(segNode, "text-match")
                ws.Cells(row, 11).Value = GetAttributeValue(segNode, "struct-match")
                ws.Cells(row, 12).Value = GetAttributeValue(segNode, "context-match")
                
                ' Get metadata from value nodes
                Dim valueNodes As Object
                Set valueNodes = segNode.SelectNodes(".//sdl:value")
                If valueNodes Is Nothing Or valueNodes.Length = 0 Then
                    Set valueNodes = segNode.SelectNodes(".//*[local-name()='value']")
                End If
                
                If Not valueNodes Is Nothing Then
                    Dim j As Long
                    For j = 0 To valueNodes.Length - 1
                        Dim valueNode As Object
                        Set valueNode = valueNodes.Item(j)
                        Dim key As String
                        key = GetAttributeValue(valueNode, "key")
                        
                        Select Case key
                            Case "created_on", "CreationDate", "SDL:CreationDate"
                                ws.Cells(row, 13).Value = valueNode.text
                            Case "created_by", "CreationUser", "SDL:CreationUser"
                                ws.Cells(row, 14).Value = valueNode.text
                            Case "last_modified_on", "modified_on", "ModificationDate"
                                ws.Cells(row, 15).Value = valueNode.text
                            Case "last_modified_by", "ModificationUser"
                                ws.Cells(row, 16).Value = valueNode.text
                        End Select
                    Next j
                End If
            End If
            
            ' Look for comments and track changes in target
            If Not targetNode Is Nothing Then
                Dim mrkNodes As Object
                Set mrkNodes = targetNode.SelectNodes(".//xliff:mrk")
                If mrkNodes Is Nothing Or mrkNodes.Length = 0 Then
                    Set mrkNodes = targetNode.SelectNodes(".//mrk")
                End If
                If mrkNodes Is Nothing Or mrkNodes.Length = 0 Then
                    Set mrkNodes = targetNode.SelectNodes(".//*[local-name()='mrk']")
                End If
                
                Dim hasRevision As Boolean
                hasRevision = False
                Dim deletedText As String
                deletedText = ""
                Dim addedText As String
                addedText = ""
                
                If Not mrkNodes Is Nothing Then
                    Dim k As Long
                    For k = 0 To mrkNodes.Length - 1
                        Dim mrkNode As Object
                        Set mrkNode = mrkNodes.Item(k)
                        Dim mtype As String
                        mtype = GetAttributeValue(mrkNode, "mtype")
                        
                        ' Check for comments
                        If mtype = "x-sdl-comment" Then
                            ' Get sdl:cid using namespace-aware method
                            Dim cmtRefId As String
                            Dim cmtRefNode As Object
                            Set cmtRefNode = mrkNode.SelectSingleNode("@sdl:cid")
                            If Not cmtRefNode Is Nothing Then
                                cmtRefId = cmtRefNode.text
                            End If
                            
                            If cmtRefId <> "" And commentDict.Exists(cmtRefId) Then
                                Dim cmtInfo As Object
                                Set cmtInfo = commentDict(cmtRefId)
                                
                                ws.Cells(row, 17).Value = cmtInfo("text")
                                ws.Cells(row, 18).Value = cmtInfo("user")
                                ws.Cells(row, 19).Value = cmtInfo("date")
                                ws.Cells(row, 20).Value = cmtInfo("severity")
                            End If
                        End If
                        
                        ' Check for revisions
                        If mtype = "x-sdl-deleted" Or mtype = "x-sdl-added" Then
                            hasRevision = True
                            
                            ' Get sdl:revid using namespace-aware method
                            Dim revRefId As String
                            Dim revRefNode As Object
                            Set revRefNode = mrkNode.SelectSingleNode("@sdl:revid")
                            If Not revRefNode Is Nothing Then
                                revRefId = revRefNode.text
                            End If
                            
                            If revRefId <> "" And revDict.Exists(revRefId) Then
                                Dim revData As Object
                                Set revData = revDict(revRefId)
                                
                                ws.Cells(row, 22).Value = revData("author")
                                ws.Cells(row, 23).Value = revData("date")
                            End If
                            
                            If mtype = "x-sdl-deleted" Then
                                deletedText = deletedText & mrkNode.text & " "
                            ElseIf mtype = "x-sdl-added" Then
                                addedText = addedText & mrkNode.text & " "
                            End If
                        End If
                    Next k
                End If
                
                ' Set track changes status and text
                ws.Cells(row, 21).Value = IIf(hasRevision, "Yes", "No")
                If deletedText <> "" Then ws.Cells(row, 24).Value = Trim(deletedText)
                If addedText <> "" Then ws.Cells(row, 25).Value = Trim(addedText)
            End If
            
            row = row + 1
        End If
    Next i
    
    ' Format results
    If row > 2 Then
        ' Create table
        On Error Resume Next
        Dim tbl As ListObject
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = "TranslationAttributes"
        tbl.TableStyle = "TableStyleMedium2"
        On Error GoTo 0
        
        ' Format columns
        ws.Columns("A:Y").AutoFit
        ws.Columns("B:C").ColumnWidth = 40
        ws.Columns("B:C").WrapText = True
        ws.Columns("Q").ColumnWidth = 35
        ws.Columns("Q").WrapText = True
        
        ' Add autofilter
        ws.Range("A1").AutoFilter
        
        MsgBox "Extracted " & (row - 2) & " translation units successfully!", vbInformation
    Else
        MsgBox "No translatable segments found in the file.", vbExclamation
    End If
    
    Application.ScreenUpdating = True
End Sub

Private Function GetAttributeValue(node As Object, attrName As String) As String
    On Error Resume Next
    GetAttributeValue = node.getAttribute(attrName)
    If Err.Number <> 0 Then GetAttributeValue = ""
    On Error GoTo 0
End Function

Private Function GetCleanText(node As Object) As String
    On Error Resume Next
    Dim text As String
    text = ""
    
    If node.HasChildNodes Then
        Dim child As Object
        For Each child In node.ChildNodes
            If child.NodeType = 3 Then ' Text node
                text = text & child.NodeValue
            ElseIf child.NodeType = 1 Then ' Element node
                If child.nodeName = "mrk" Then
                    If GetAttributeValue(child, "mtype") <> "x-sdl-deleted" Then
                        text = text & GetCleanText(child)
                    End If
                Else
                    text = text & GetCleanText(child)
                End If
            End If
        Next child
    Else
        text = node.text
    End If
    
    GetCleanText = text
    On Error GoTo 0
End Function

