Attribute VB_Name = "GetAccept"
Option Explicit
Private GlobalPersonSourceTab As String
Private GlobalPersonSourceField As String

'#Global variables
Private GlobalDocumentTypeField As String


Declare Function GetSystemMetrics32 Lib "user32" _
    Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Public TokenHandler As String

Public Sub SetTokens(strToken As String)
    On Error GoTo ErrorHandler
    
    'used to combine token between modal and parent actionpad
    TokenHandler = strToken
    If strToken = "-" Then
        TokenHandler = ""
    End If
    
    Exit Sub
ErrorHandler:
    UI.ShowError ("GetAccept.SetTokens")
End Sub

Public Function OpenGetAccept(className As String, personSourceTab As String, personSourceField As String) As String
    On Error GoTo ErrorHandler
    
    Dim oDialog As Lime.Dialog
    Dim oInspector As New Lime.Inspector
    Set oInspector = ThisApplication.ActiveInspector
    
    GlobalPersonSourceTab = personSourceTab
    GlobalPersonSourceField = personSourceField
    
    '#Global variables for fields in Lime CRM
    'Set value to empty if field not excists
    GlobalDocumentTypeField = "type"

    If Globals.VerifyInspector(className, oInspector) And GetAccept.SaveNew() Then
        If Not oInspector.ActiveExplorer Is Nothing Then
            If oInspector.ActiveExplorer.Class.Name = "document" Then
                If oInspector.ActiveExplorer.Selection.Count > 0 Then
                    If oInspector.ActiveExplorer.ActiveItem.Record.Document("document") Is Nothing Then
                        Call Lime.MessageBox(Localize.GetText("GetAccept", "ga_missing_file"))
                        OpenGetAccept = "-1"
                        Exit Function
                    ElseIf Not CheckFileTypes(oInspector.ActiveExplorer.ActiveItem.Record.Document("document").Extension) Then
                        Call Lime.MessageBox(Localize.GetText("GetAccept", "ga_invalid_filetype"))
                        OpenGetAccept = "-1"
                        Exit Function
                    Else
                        Set oDialog = New Lime.Dialog
                        oDialog.Type = lkDialogHTML
                        oDialog.Property("url") = Application.WebFolder & "lbs.html?ap=apps/GetAccept/getaccept&type=tab"
                        oDialog.Property("height") = 530
                        oDialog.Property("width") = 700
                        oDialog.show
                        OpenGetAccept = TokenHandler
                        Exit Function
                End If
                Else
                    Call Lime.MessageBox(Localize.GetText("GetAccept", "i_only_one_document"))
                    OpenGetAccept = "-1"
                    Exit Function
                End If
            Else
                Call Lime.MessageBox(Localize.GetText("GetAccept", "i_no_document_tab_selected"))
            End If
        End If
    End If
    
    GlobalPersonSourceTab = ""
    GlobalPersonSourceField = ""
    
    Exit Function
ErrorHandler:
    UI.ShowError ("GetAccept.OpenGetAccept")
End Function




Public Function CheckFileTypes(fileType As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim vAcceptedFileType As Variant
    Dim colAcceptedFileTypes As Variant
    '''If you need to send more types of document. Just check if GetAccept can handle them and then add it to the list below
    colAcceptedFileTypes = Array("doc", "docx", "pdf", "ppt", "txt")
    
    For Each vAcceptedFileType In colAcceptedFileTypes
        If vAcceptedFileType = fileType Then
            CheckFileTypes = True
            Exit Function
        End If
    Next
    CheckFileTypes = False
    
    Exit Function
ErrorHandler:
    UI.ShowError ("GetAccept.CheckFileType")
End Function

Public Function GetContactList(className As String) As String
    'Get the contacts from the connected company
 
    On Error GoTo ErrorHandler
    
    Dim oRecords As LDE.Records
    Dim oRecord As LDE.Record
    Dim oView As LDE.View
    Dim oFilter As LDE.Filter
    Dim oInspector As Lime.Inspector
    Dim strJSON As String
    Dim i As Integer
    
    Set oInspector = Application.ActiveInspector
    If Globals.VerifyInspector(className, oInspector) And GetAccept.SaveNew() Then
        Set oView = New LDE.View
        Call oView.Add("firstname", lkSortAscending)
        Call oView.Add("lastname")
        Call oView.Add("email")
        Call oView.Add("mobilephone")
        
        If GlobalPersonSourceTab <> "" Then
            If oInspector.Explorers.Exists(GlobalPersonSourceTab) Then
                Set oFilter = New LDE.Filter
                Call oFilter.AddCondition(oInspector.Class.Name, lkOpEqual, oInspector.Record.ID)
                
                If oFilter.HitCount(Database.Classes(GlobalPersonSourceTab)) > 0 Then
                    Set oRecords = New LDE.Records
                    Call oRecords.Open(Database.Classes(GlobalPersonSourceTab), oFilter, oView)
                    strJSON = CreatePersonJSON(oRecords)
                End If
            Else
                Call Lime.MessageBox(Localize.GetText("GetAccept", "i_cant_get_person"))
                
            End If
        End If
        
        If GlobalPersonSourceField <> "" Then
            Set oFilter = New LDE.Filter
            Call oFilter.AddCondition(GlobalPersonSourceField, lkOpEqual, oInspector.Controls.GetValue(GlobalPersonSourceField))
            
            If oFilter.HitCount(Database.Classes("person")) > 0 Then
                Set oRecords = New LDE.Records
                Call oRecords.Open(Database.Classes("person"), oFilter, oView)
                
                strJSON = CreatePersonJSON(oRecords)
            End If
        End If
    End If
    
    GetContactList = strJSON

    Exit Function
ErrorHandler:
    Call UI.ShowError("GetAccept.GetContactList")
    GetContactList = ""
End Function
Public Function GetCoworkerList()
    'Get the coworkers from Coworker tab
 
    On Error GoTo ErrorHandler
    
    Dim oRecords As LDE.Records
    Dim oRecord As LDE.Record
    Dim oView As LDE.View
    Dim oFilter As LDE.Filter
    Dim oInspector As Lime.Inspector
    Dim strJSON As String
    Dim i As Integer
    
    Set oInspector = Application.ActiveInspector
    
        Set oView = New LDE.View
        Call oView.Add("firstname", lkSortAscending)
        Call oView.Add("lastname")
        Call oView.Add("email")
        Call oView.Add("cellphone")
                
            Set oFilter = New LDE.Filter
          
            Call oFilter.AddCondition("inactive", lkOpEqual, False)
            
            If oFilter.HitCount(Database.Classes("coworker")) > 0 Then
                Set oRecords = New LDE.Records
                Call oRecords.Open(Database.Classes("coworker"), oFilter, oView, 10)
                strJSON = CreatePersonJSON(oRecords)
                
            End If
        GetCoworkerList = strJSON

    Exit Function
ErrorHandler:
    Call UI.ShowError("GetAccept.GetContactList")
   
End Function


Public Function CreatePersonJSON(oRecords As LDE.Records) As String
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim oRecord As LDE.Record
    Dim strJSON As String
    i = 0
    strJSON = "{" + """Persons"":[{" _
    
    'loop through the coworkers and build up a JSON
    If oRecords.Class.Name = "coworker" Then
        For Each oRecord In oRecords
            i = i + 1
            strJSON = strJSON + """firstname"":""" & oRecord("firstname") & """," _
            & """lastname"":""" & oRecord("lastname") & """," _
            & """mobilephone"":""" & oRecord("cellphone") & """," _
            & """email"":""" & oRecord("email") & """" _
        
            If i < oRecords.Count Then
                strJSON = strJSON + "},{"
            Else
                strJSON = strJSON + "}"
            End If
        
        Next oRecord
    End If
    
    'loop through the persons and build up a JSON
    
    If oRecords.Class.Name = "person" Then
        For Each oRecord In oRecords
            i = i + 1
            strJSON = strJSON + """firstname"":""" & oRecord("firstname") & """," _
            & """lastname"":""" & oRecord("lastname") & """," _
            & """mobilephone"":""" & oRecord("mobilephone") & """," _
            & """email"":""" & oRecord("email") & """" _
        
            If i < oRecords.Count Then
                strJSON = strJSON + "},{"
            Else
                strJSON = strJSON + "}"
            End If
            
        Next oRecord
    End If
    
    strJSON = strJSON + "]}"
    
    CreatePersonJSON = strJSON
    
    Exit Function
ErrorHandler:
    Call UI.ShowError("GetAccept.CreatePersonJSON")
End Function

Public Function CheckDocuments(activeRecordId As Long, activeClass As String) As String
    On Error GoTo ErrorHandler
    'Check if there are any documents sent with GetAccept connected to the inspector
    Dim oRecords As New LDE.Records
    Dim oRecord As LDE.Record
    Dim oView As New LDE.View
    Dim oFilter As New LDE.Filter
    Dim retval As String
    Dim i As Integer
    i = 0
    
    Call oView.Add("iddocument")
    
    Call oFilter.AddCondition("sent_with_ga", lkOpEqual, 1)
    Call oFilter.AddCondition(activeClass, lkOpEqual, activeRecordId)
    Call oFilter.AddOperator(lkOpAnd)
    
    If activeRecordId > 0 Then
        If oFilter.HitCount(Application.Classes("document")) > 0 Then
            Call oRecords.Open(Database.Classes("document"), oFilter, oView)
            For Each oRecord In oRecords
                i = i + 1
                retval = retval & oRecord.ID
                
                If i < oRecords.Count Then
                    retval = retval & ","
                End If
            Next oRecord
            
            CheckDocuments = retval
            Exit Function
        Else
            CheckDocuments = "False"
            Exit Function
        End If
    Else
        CheckDocuments = "False"
        Exit Function
    End If
    
    Exit Function
ErrorHandler:
    UI.ShowError ("GetAccept.CheckDocuments")
    CheckDocuments = False
End Function

Public Function showList(sType As String) As Boolean
    On Error GoTo ErrorHandler
    'Check if there are any documents sent with GetAccept connected to the inspector
    If Not (ActiveControls.State And lkControlsStateNew) = lkControlsStateNew Then
   
        Dim oFilter As New LDE.Filter
        Call oFilter.AddCondition("sent_with_ga", lkOpEqual, 1)
        Call oFilter.AddCondition(sType, lkOpEqual, ActiveInspector.Record.ID)
        Call oFilter.AddOperator(lkOpAnd)
        
        If oFilter.HitCount(Application.Classes("document")) > 0 Then
            showList = True
            Exit Function
        Else
            showList = False
            Exit Function
        End If
    Else
        showList = False
    End If
    
    
    Exit Function
ErrorHandler:
    UI.ShowError ("GetAccept.showList")
    showList = False
End Function

Public Function GetDocumentData(className As String) As String
    'Collects the document data from the selected document in the table document
    On Error GoTo ErrorHandler
    
    Dim retval As String
    Dim oRecord As LDE.Record
    Dim oView As LDE.View
    Dim oItem As New Lime.ExplorerItem
    Dim oInspector As New Lime.Inspector
    Set oInspector = ThisApplication.ActiveInspector
    retval = "["
    If Globals.VerifyInspector(className, oInspector) And GetAccept.SaveNew() Then
        If Not oInspector.ActiveExplorer Is Nothing Then
            
            If oInspector.ActiveExplorer.Class.Name = "document" Then
                For Each oItem In oInspector.ActiveExplorer.Selection
                    Set oRecord = New LDE.Record
                    Set oView = New LDE.View
                    Call oView.Add("document")
                    Call oView.Add("comment", lkSortAscending)
               
                    Call oRecord.Open(Database.Classes("document"), oItem.Record.ID, oView)
                    retval = retval & " { "
                    retval = retval & " ""file_name"" : """ & oRecord.Value("comment")
                    retval = retval & "."
                    retval = retval & oRecord.Document("document").Extension & ""","
                    retval = retval & " ""file_content"" :  """ & VBA.Replace(VBA.Replace(VBA.Replace(VBA.Replace(EncodeBase64(oRecord.Document("document").Contents), "/", "\/"), """", "\"""), vbLf, ""), vbCr, "") & """ "
                    retval = retval & " },"
                Next
            Else
                Lime.MessageBox (Localize.GetText("GetAccept", "i_no_document_tab_selected"))
            End If
        End If
    End If
    If VBA.Len(retval) > 3 Then
        retval = VBA.Left(retval, VBA.Len(retval) - 1)
    End If
    retval = retval & "]"
    
    GetDocumentData = retval
    
    Exit Function
ErrorHandler:
    UI.ShowError ("GetAccept.GetDocumentData")
    GetDocumentData = ""
End Function

Public Function GetDocumentType() As Boolean
    'returns true if there is a certain documenttype that you choose, can be used to block send outs of certain doc types
    On Error GoTo ErrorHandler
    
    Dim retval As Boolean
    Dim oRecord As LDE.Record
    Dim oView As LDE.View
    Dim oInspector As New Lime.Inspector
    Set oInspector = ThisApplication.ActiveInspector
    ' The user has selected an document
    
        If Not oInspector.ActiveExplorer Is Nothing Then
            If oInspector.ActiveExplorer.Class.Name = "document" Then
                Set oRecord = New LDE.Record
                Set oView = New LDE.View
                
                If GlobalDocumentTypeField <> "" Then
                    Call oView.Add(GlobalDocumentTypeField)
                End If
                
                Call oRecord.Open(Database.Classes("document"), oInspector.ActiveExplorer.Selection.Item(oInspector.ActiveExplorer.Selection.Count).Record.ID, oView)
                
                If GlobalDocumentTypeField <> "" Then
                    If oRecord.text("type") = "Quote" Then
                        retval = True
                    Else
                        retval = False
                    End If
                Else
                    retval = False
                End If
            End If
        End If
   
    
    GetDocumentType = retval
    Exit Function
ErrorHandler:
    UI.ShowError ("GetAccept.GetDocumentType")
End Function

Public Function GetDocumentDescription(className As String) As String
    'returns the document name and file extension
    On Error GoTo ErrorHandler
    
    Dim retval As String
    Dim oRecord As LDE.Record
    Dim oView As LDE.View
    Dim oInspector As New Lime.Inspector
    Set oInspector = ThisApplication.ActiveInspector
    ' The user has selected an document
    If Globals.VerifyInspector(className, oInspector) And GetAccept.SaveNew() Then
        If Not oInspector.ActiveExplorer Is Nothing Then
            If oInspector.ActiveExplorer.Class.Name = "document" Then
                Set oRecord = New LDE.Record
                Set oView = New LDE.View
                Call oView.Add("document")
                Call oView.Add("comment", lkSortAscending)
                
                Call oRecord.Open(Database.Classes("document"), oInspector.ActiveExplorer.Selection.Item(oInspector.ActiveExplorer.Selection.Count).Record.ID, oView)
                If Not oRecord.Document("document") Is Nothing Then
                    retval = retval & oRecord.Value("comment")
                    retval = retval & "."
                    retval = retval & oRecord.Document("document").Extension
                End If
            End If
        End If
    End If
    
    GetDocumentDescription = retval
    Exit Function
ErrorHandler:
    UI.ShowError ("GetAccept.GetDocumentDescription")
End Function

Public Function GetDocumentId(className As String) As String
    'returns the document id
    On Error GoTo ErrorHandler
    
    Dim retval As String
    Dim oInspector As New Lime.Inspector
    Set oInspector = ThisApplication.ActiveInspector
    ' The user has selected an document
    If Globals.VerifyInspector(className, oInspector) And GetAccept.SaveNew() Then
        If Not oInspector.ActiveExplorer Is Nothing Then
            If oInspector.ActiveExplorer.Class.Name = "document" Then
                GetDocumentId = oInspector.ActiveExplorer.Selection.Item(1).Record.ID
            End If
        End If
    End If
    
    Exit Function
ErrorHandler:
    UI.ShowError ("GetAccept.GetDocumentId")
End Function

Public Sub SetDocumentStatus(sStatus As String, className As String)
    'set document sent_with_ga parameter
    On Error GoTo ErrorHandler
    
    Dim retval As String
    Dim oInspector As New Lime.Inspector
    Dim oItem As New Lime.ExplorerItem
    Dim oRecordDocument As LDE.Record
    Set oInspector = ThisApplication.ActiveInspector
    
    ' The user has selected an document
    If Globals.VerifyInspector(className, oInspector) And GetAccept.SaveNew() Then
        If Not oInspector.ActiveExplorer Is Nothing Then
            If oInspector.ActiveExplorer.Class.Name = "document" Then
                
                'If oInspector.ActiveExplorer.Selection.Count = 1 Then
                For Each oItem In oInspector.ActiveExplorer.Selection
                    ' Set sent_with_ga status
                    Set oRecordDocument = New LDE.Record
                    oRecordDocument.Open Classes("document"), oItem.Record.ID
                    oRecordDocument.Value("sent_with_ga") = sStatus
                    Call oRecordDocument.Update
                    
                    ' Create historynote
                    Dim oRecordHistory As New LDE.Record
                    oRecordHistory.Open Classes("history")
                    ' Check that the field with same class name exist on the document which should be connected
                    If oRecordHistory.Fields.Exists(oInspector.Class.Name) Then
                        oRecordHistory.Value(oInspector.Class.Name) = oInspector.Record.ID
                    End If
                    oRecordHistory.Value("type") = Database.Classes("history").Fields("type").Options.Lookup("sentemail", lkLookupOptionByKey).Value
                    oRecordHistory.Value("note") = "Sent with GetAccept"
                    oRecordHistory.Value("date") = VBA.Now
                    If oRecordHistory.Fields.Exists("document") Then
                        oRecordHistory.Value("document") = oInspector.ActiveExplorer.Selection.Item(1).Record.ID
                    End If
                    Call oRecordHistory.Update
                Next
               
                
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    UI.ShowError ("GetAccept.SetDocumentStatus")
End Sub

Public Sub OpenGALink(ByVal sLink As String)

    Call Application.Shell(sLink)
    
    Exit Sub
ErrorHandler:
    UI.ShowError ("GetAccept.OpenGALink")
End Sub

Private Function EncodeBase64(ByRef arrData() As Byte) As String
    On Error GoTo ErrorHandler
    
    Dim objXML As Object
    Dim objNode As Object
    
    Set objXML = VBA.CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text
 
    Set objNode = Nothing
    Set objXML = Nothing
    
    Exit Function
ErrorHandler:
        UI.ShowError ("GetAccept.EncodeBase64")
End Function

' ##SUMMARY Saves changes made in actionpad.
Public Function SaveNew() As Boolean
    On Error GoTo ErrorHandler
    
    Dim oInspector As Lime.Inspector
    
    Set oInspector = Application.ActiveInspector
    
    On Error GoTo ErrorSave
        If (oInspector.Record.State And lkRecordStateNew) = lkRecordStateNew Then
            Call oInspector.Save(True)
        End If
        GoTo SaveOK
ErrorSave:
        Lime.MessageBox (Err.Description)
        SaveNew = False
        Exit Function
SaveOK:
    SaveNew = True

    Exit Function
ErrorHandler:
    Call UI.ShowError("GetAccept.TrySave")
    SaveNew = False
End Function


Public Sub DownloadFile(sLink As String, sFileName As String, className As String, commentField As String)
    On Error GoTo ErrorHandler
    
    ThisApplication.MousePointer = 11
    Dim myURL As String
    myURL = sLink
    
    Dim oInspector As Lime.Inspector
    
    Set oInspector = Application.ActiveInspector
    
    Dim WinHttpReq As Object
    Dim oStream As Object
    Dim sFileLocation As String
    Dim sMapLocation As String
    Dim oRecord As New LDE.Record
    Dim pDocument As New LDE.Document
    
    
    sMapLocation = ThisApplication.TemporaryFolder & "\GetAccept\"
    sFileLocation = sMapLocation & sFileName & ".pdf"
    
    If Len(Dir(sMapLocation, vbDirectory)) = 0 Then
        MkDir sMapLocation
    End If
    
    Set WinHttpReq = VBA.CreateObject("WinHttp.WinHttpRequest.5.1")
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.Send
    
    myURL = WinHttpReq.responseBody
    If WinHttpReq.Status = 200 Then
        Set oStream = VBA.CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile sFileLocation, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
        
        Call pDocument.Load(sFileLocation)
        Call oRecord.Open(Database.Classes("document"))
        oRecord.Value("document") = pDocument
        If oRecord.Fields.Exists("type") Then
            oRecord("type") = Database.Classes("document").Fields("type").Options.Lookup("agreement", lkLookupOptionByKey)
        End If
        If oRecord.Fields.Exists(className) Then
            oRecord(className) = oInspector.Record.ID
        End If
        'connect company if a company field exists on the parent card and the document card.
        If className <> "company" Then 'only done if the parent isnt alreaady the company
            If oRecord.Fields.Exists("company") Then
                If oInspector.Record.Fields.Exists("company") Then
                    oRecord("company") = oInspector.Controls.GetValue("company")
                End If
            End If
        End If
        
        oRecord(commentField) = sFileName & " (" & (Localize.GetText("GetAccept", "ga_signed")) & ")"
        oRecord("sent_with_ga") = 1
        oRecord.Update
         
    Else
        Call Lime.MessageBox(Localize.GetText("GetAccept", "i_download_failed"))
    End If
    
    VBA.Kill (sFileLocation)
    
    ThisApplication.MousePointer = 1
    Exit Sub
ErrorHandler:
    Call UI.ShowError("GetAccept.DownloadFile")
    ThisApplication.MousePointer = 1
End Sub

Private Function AddOrCheckLocalize(sOwner As String, sCode As String, sDescription As String, sEN_US As String, sSV As String, sNO As String, sFI As String, sDA As String) As Boolean
    On Error GoTo ErrorHandler:
    Dim oFilter As New LDE.Filter
    Dim oRecs As New LDE.Records
    
    Call oFilter.AddCondition("owner", lkOpEqual, sOwner)
    Call oFilter.AddCondition("code", lkOpEqual, sCode)
    oFilter.AddOperator lkOpAnd
    
    If oFilter.HitCount(Database.Classes("localize")) = 0 Then
        Debug.Print ("Localization " & sOwner & "." & sCode & " not found, creating new!")
        Dim oRec As New LDE.Record
        Call oRec.Open(Database.Classes("localize"))
        oRec.Value("owner") = sOwner
        oRec.Value("code") = sCode
        oRec.Value("context") = sDescription
        
        'Disable languages below that you do not have your Lime Crm Solution
        oRec.Value("en_us") = sEN_US
        oRec.Value("sv") = sSV
        oRec.Value("no") = sNO
        oRec.Value("fi") = sFI
        oRec.Value("da") = sDA
        Call oRec.Update
    ElseIf oFilter.HitCount(Database.Classes("localize")) = 1 Then
    Debug.Print ("Updating localization " & sOwner & "." & sCode)
        Call oRecs.Open(Database.Classes("localize"), oFilter)
        oRecs(1).Value("owner") = sOwner
        oRecs(1).Value("code") = sCode
        oRecs(1).Value("context") = sDescription
        oRecs(1).Value("sv") = sSV
        oRecs(1).Value("en_us") = sEN_US
        oRecs(1).Value("no") = sNO
        oRecs(1).Value("fi") = sFI
        oRecs(1).Value("da") = sDA
        Call oRecs.Update
        
    Else
        Call MsgBox("There are multiple copies of " & sOwner & "." & sCode & "  which is bad! Fix it", vbCritical, "To many translations makes Jack a dull boy")
    End If
    
    Set Localize.dicLookup = Nothing
    AddOrCheckLocalize = True
    Exit Function
ErrorHandler:
    Debug.Print ("Error while validating or adding Localize")
    AddOrCheckLocalize = False
End Function


Public Sub initGa(personSourceTab As String, personSourceField As String)
On Error GoTo ErrorHandler
    GlobalPersonSourceTab = personSourceTab
    GlobalPersonSourceField = personSourceField
Exit Sub
ErrorHandler:
    Call UI.ShowError("GetAccept.personSourceTab")
End Sub

Public Sub CreateTodo(days As Integer)
On Error GoTo ErrorHandler
    Dim oRecord As New LDE.Record
    Dim oInspector As Lime.Inspector
    Dim sDate As String
    
    Set oInspector = Application.ActiveInspector
    
    If oInspector.Explorers.Exists("todo") Then
        sDate = VBA.DateAdd("d", days, VBA.Date)
        Call oRecord.Open(Database.Classes("todo"))
        oRecord.Value("subject") = "Follow up GA document"
        oRecord.Value("starttime") = sDate
        oRecord.Value(oInspector.Record.Class.Name) = oInspector.Record.ID
        oRecord.Update
    Else
        Lime.MessageBox ("Couldn't create a todo")
    End If
Exit Sub
ErrorHandler:
    Call UI.ShowError("GetAccept.CreateTodo")
End Sub

Public Sub CreateHistory()
On Error GoTo ErrorHandler
    Dim oInspector As Lime.Inspector
    
    Set oInspector = Application.ActiveInspector
    ' Create historynote
    Dim oRecordHistory As New LDE.Record
    oRecordHistory.Open Classes("history")
    ' Check that the field with same class name exist on the document which should be connected
    If oRecordHistory.Fields.Exists(oInspector.Class.Name) Then
        oRecordHistory.Value(oInspector.Class.Name) = oInspector.Record.ID
    End If
    oRecordHistory.Value("type") = Database.Classes("history").Fields("type").Options.Lookup("sentemail", lkLookupOptionByKey).Value
    oRecordHistory.Value("note") = "Sent with GetAccept"
    oRecordHistory.Value("date") = VBA.Now
    oRecordHistory.Value("coworker") = ActiveUser.Record.ID
    If oInspector.ActiveExplorer.Class.Name = "document" Then
        If oInspector.ActiveExplorer.Selection.Count > 0 Then
            If oRecordHistory.Fields.Exists("document") Then
                oRecordHistory.Value("document") = oInspector.ActiveExplorer.Selection.Item(1).Record.ID
            End If
        End If
    End If
    Call oRecordHistory.Update
Exit Sub
ErrorHandler:
    Call UI.ShowError("GetAccept.CreateHistory")
End Sub


Public Function GetDocuments(className As String) As String
    'returns the document name and file extension
    On Error GoTo ErrorHandler
    
    Dim retval As String
    Dim oRecords As New LDE.Records
    Dim oRecord As New LDE.Record
    Dim oPool As New LDE.Pool
    Dim oView As LDE.View
    Dim oInspector As New Lime.Inspector
    Set oInspector = ThisApplication.ActiveInspector
    ' The user has selected an document
    If Globals.VerifyInspector(className, oInspector) And GetAccept.SaveNew() Then
        If Not oInspector.ActiveExplorer Is Nothing Then
            If oInspector.ActiveExplorer.Class.Name = "document" Then
                Set oRecord = New LDE.Record
                Set oView = New LDE.View
                Call oView.Add("document")
                Call oView.Add("comment", lkSortAscending)
                Call oView.Add("iddocument")
                Set oPool = oInspector.ActiveExplorer.Selection.Pool
                
                Call oRecords.Open(Database.Classes("document"), oPool, oView)
                retval = "["
                For Each oRecord In oRecords
                    If Not oRecord.Document("document") Is Nothing Then
                        retval = retval & "{"
                        retval = retval & " ""name"" : """ & oRecord.Value("comment") & "." & oRecord.Document("document").Extension & ""","
                        retval = retval & " ""id"" : """ & oRecord.ID & """"
                        retval = retval & " },"
                    End If
                Next oRecord
            End If
        End If
    End If
    retval = Left(retval, Len(retval) - 1)
    retval = retval & "]"
    
    GetDocuments = retval
    Exit Function
ErrorHandler:
    UI.ShowError ("GetAccept.GetDocuments")
End Function

'Run Installation to get all transalations installed
Public Sub Install()
    On Error GoTo ErrorHandler
    Dim key As Variant
    
    Dim en As Scripting.Dictionary
    Dim sv As Scripting.Dictionary
    Dim no As Scripting.Dictionary
    Dim fi As Scripting.Dictionary
    Dim da As Scripting.Dictionary
    
    Set en = LoadLanguage("apps\GetAccept-v2\Install\Locals\getaccept-xml-archive\res\values\strings.xml")
    Set sv = LoadLanguage("apps\GetAccept-v2\Install\Locals\getaccept-xml-archive\res\values-sv-rSE\strings.xml")
    Set no = LoadLanguage("apps\GetAccept-v2\Install\Locals\getaccept-xml-archive\res\values-no-rNO\strings.xml")
    Set fi = LoadLanguage("apps\GetAccept-v2\Install\Locals\getaccept-xml-archive\res\values-fi-rFI\strings.xml")
    Set da = LoadLanguage("apps\GetAccept-v2\Install\Locals\getaccept-xml-archive\res\values-da-rDK\strings.xml")
    
    For Each key In en
        If en.Exists(key) And sv.Exists(key) And no.Exists(key) And fi.Exists(key) Then
            Call AddOrCheckLocalize("GetAccept", CStr(key), en.Item(key), en.Item(key), sv.Item(key), no.Item(key), fi.Item(key), da.Item(key))
        End If
    Next key
    Debug.Print "----INSTALLATION IS DONE----"
    Exit Sub
ErrorHandler:
    UI.ShowError ("GetAccept.Install")
End Sub

Public Function LoadLanguage(FilePath As String) As Scripting.Dictionary
    On Error GoTo ErrorHandler
    
    Dim bundle As New Scripting.Dictionary
    Dim oXmlFile As MSXML2.DOMDocument60
    Dim oChild As MSXML2.IXMLDOMNode
    
    Set oXmlFile = New MSXML2.DOMDocument60
    oXmlFile.async = False
    
    If oXmlFile.Load(WebFolder + FilePath) Then
        For Each oChild In oXmlFile.childNodes
            If oChild.nodeName = "resources" Then
                Call AddToBundle(bundle, oChild.childNodes)
            End If
        Next oChild
    Else
        Lime.MessageBox ("Could not find language file: '" & FilePath & "'")
    End If
    
    Set LoadLanguage = bundle
    Exit Function
ErrorHandler:
    UI.ShowError ("GetAccept.LoadLanguage")
End Function

Public Sub AddToBundle(ByRef languageBundle As Scripting.Dictionary, ByRef nodes As MSXML2.IXMLDOMNodeList)
    On Error GoTo ErrorHandler
    
    Dim sKey As String
    Dim sValue As String
    Dim xNode As MSXML2.IXMLDOMNode
    Dim oArray As Collection
    
    For Each xNode In nodes
        sKey = xNode.Attributes.getNamedItem("name").text
        'sKey = Replace(sKey, "_", "-")
        sValue = xNode.text
        Call languageBundle.Add(sKey, sValue)
    Next xNode
    Exit Sub
ErrorHandler:
    UI.ShowError ("GetAccept.BuildBundle")
End Sub


Public Function SearchCoworkerByEmail(email As String) As String
    On Error GoTo ErrorHandler
    
    Dim oRecords As New LDE.Records
    Dim oRecord As New LDE.Record
    Dim oFilter As New LDE.Filter
    Dim strJSON As String
    Dim oView As LDE.View
    Dim oInspector As New Lime.Inspector
    Set oInspector = ThisApplication.ActiveInspector
    
    If email <> "" Then
        Set oView = New LDE.View
        Call oView.Add("firstname", lkSortAscending)
        Call oView.Add("lastname")
        Call oView.Add("email")
        Call oView.Add("cellphone")
        
        Set oFilter = New LDE.Filter
          
            Call oFilter.AddCondition("inactive", lkOpEqual, False)
            Call oFilter.AddCondition("email", lkOpLike, email)
            Call oFilter.AddOperator(lkOpAnd)
            
            If oFilter.HitCount(Database.Classes("coworker")) > 0 Then
                Set oRecords = New LDE.Records
                Call oRecords.Open(Database.Classes("coworker"), oFilter, oView)
                strJSON = CreatePersonJSON(oRecords)
                SearchCoworkerByEmail = strJSON
            Else
                SearchCoworkerByEmail = "{}"
            End If
    Else
        SearchCoworkerByEmail = "{}"
    End If
    
    Exit Function
ErrorHandler:
    UI.ShowError ("GetAccept.SearchCoworkerByEmail")
End Function

Public Sub openGaModal()
    On Error GoTo ErrorHandler:
    
    Dim oDialog As Lime.Dialog
    Set oDialog = New Lime.Dialog
    oDialog.Type = lkDialogHTML
    oDialog.Property("url") = Application.WebFolder & "lbs.html?ap=apps/GetAccept-v2/availablefields&type=tab"
    oDialog.Property("height") = 530
    oDialog.Property("width") = 700
    oDialog.show

    Exit Sub
ErrorHandler:
    UI.ShowError ("GetAccept.openGaModal")
End Sub


