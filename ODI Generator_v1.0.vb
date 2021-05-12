Sub odic()
    Dim XDoc As Object, root As Object
    Dim op As String
    Dim object_name As String
    Dim object_version As String
    Dim object_desc As String
    Dim subsheet_id As String
    Dim subsheet_name As String
    Dim st_id As String
    Dim end_id As String
    Dim input_name As String
    Dim input_desc As String
    Dim input_type As String
    Dim output_name As String
    Dim output_desc As String
    Dim output_type As String
    
    Dim oid As String
    Dim aid As String
    Dim iid As String
    Dim outid As String
    Dim nid As String
    Dim rid As Integer
    Dim cid As Integer
    Dim s_temp As Integer
    Dim e_temp As Integer
    
    oid = "A"
    aid = "B"
    iid = "C"
    outid = "D"
    nid = "G"
    
    rid = 6
    cid = 1
    
    
    Dim lngCount As Long
    
' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Show
        op = .SelectedItems(1)
    End With
    
    If op = "" Then
        MsgBox "Next time...select valid file :("
    End If
    
    rid = InputBox("Enter Starting row id")
    
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load (op)
    Set root = XDoc.DocumentElement
    object_name = root.Attributes(0).Text
    object_version = root.Attributes(1).Text
    object_desc = root.Attributes(3).Text
    ActiveSheet.Range(oid & rid).Value = object_name & "_v" & object_version
'ActiveSheet.Range("B4").Value = object_desc
'Debug.Print object_name & "_v" & object_version & " __ " & object_desc
'For Each Subsheet ***********************************
    Set subsheets = XDoc.SelectNodes("//subsheet[@type='Normal']")
    For Each Subsheet In subsheets
        subsheet_id = Subsheet.Attributes(0).Text
        subsheet_name = Subsheet.FirstChild.Text
        Debug.Print "____________________________________________________________________________________________________________"
        Debug.Print subsheet_id & " : " & subsheet_name
        ActiveSheet.Range(aid & rid).Value = subsheet_name
        s_temp = rid
'For each start stage in Document
        Debug.Print "Inputs________________"
        Set starts = XDoc.SelectNodes("//stage[@type='Start']")
        For Each s In starts
'Identify correct start stage
            If InStr(s.XML, subsheet_id) > 0 Then
'Grab the identified -  start id
                st_id = s.Attributes(0).Text
                Set inputs = XDoc.SelectNodes("//stage[@stageid='" & st_id & "']/inputs/input")
'For each input of that start id
                For Each input_stage In inputs
'Grab the input details
                    input_name = input_stage.Attributes(1).Text
                    input_desc = input_stage.Attributes(2).Text
                    input_type = input_stage.Attributes(0).Text
                    Debug.Print input_name & ":" & input_desc & ":" & input_type
                    ActiveSheet.Range(iid & rid).Value = input_name
                    ActiveSheet.Range(nid & rid).Value = input_desc & ", type: " & input_type
                    rid = rid + 1
                Next
            End If
        Next
        e_temp = rid
        Debug.Print "Outputs________________"
        rid = s_temp
'For each end stage in Document
        Set ends = XDoc.SelectNodes("//stage[@type='End']")
        For Each e In ends
'Identify correct start stage
            If InStr(e.XML, subsheet_id) > 0 Then
'Grab the identified -  start id
                end_id = e.Attributes(0).Text
                Set outputs = XDoc.SelectNodes("//stage[@stageid='" & end_id & "']/outputs/output")
'For each input of that start id
                For Each output_stage In outputs
'Grab the input details
                    output_name = output_stage.Attributes(1).Text
                    output_desc = output_stage.Attributes(2).Text
                    output_type = output_stage.Attributes(0).Text
                    Debug.Print output_name & ":" & output_desc & ":" & output_type
                    ActiveSheet.Range(outid & rid).Value = output_name
                    ActiveSheet.Range(nid & rid).Value = output_desc & ", type: " & output_type
                    rid = rid + 1
                Next
                
            End If
        Next
        If rid > e_temp Then
            e_temp = rid
        End If
        If s_temp = e_temp Then
            rid = e_temp + 1
        Else
            rid = IIf(s_temp >= e_temp, s_temp, e_temp)
        End If
        
    Next
'get sbs id - Debug.Print subsheet(0).Attributes(0).Text
'Get all xml of subsheet
'Debug.Print subsheet(0).XML
End Sub

