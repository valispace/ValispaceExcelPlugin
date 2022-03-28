Attribute VB_Name = "ValiAddon"

Public id_array() As String
Public valis As Object
Public valiUrl As String
Public projectId As String
Public username As String
Public password As String
Public create_links As Boolean
Public cache_valis As Boolean
Public token As String
Public strCookie As String
Public xobj As Object



Private Sub SetVariables()
    valiUrl = GetSetting("ValiAddon", "Settings", "URL")
    projectId = GetSetting("ValiAddon", "Settings", "ProjectID")
    username = GetSetting("ValiAddon", "Settings", "User")
    password = GetSetting("ValiAddon", "Settings", "PW")
    create_links = CBool(GetSetting("ValiAddon", "Settings", "LINKS"))
    cache_valis = CBool(GetSetting("ValiAddon", "Settings", "CACHE"))
End Sub


Private Function Login()
    SetVariables

    On Error GoTo ConnectionFail
    
      'Ignoring Trailing "/" on URL
      If Right(valiUrl, 1) = "/" Then
        valiUrl = Left(valiUrl, Len(valiUrl) - 1)
      End If

      oAuthUrl = valiUrl & "/o/token/"

      Set xobj = CreateObject("WinHttp.WinHttpRequest.5.1") 'New WinHttp.WinHttpRequest

      ' request access token
      xobj.Open "POST", oAuthUrl, False
      xobj.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
      xobj.Send "grant_type=password&client_id=docs.valispace.com/user-guide/addons/#excel" & "&username=" & username & "&password=" & password

      If Not xobj.Status = 200 Then
        GoTo ConnectionFail
      End If

      Dim TokenResponse As Object
      Set TokenResponse = JsonConverter.ParseJson(xobj.ResponseText)
      token = "Bearer " & TokenResponse("access_token")

      Exit Function

ConnectionFail:
    MsgBox ("Connection to " & valiUrl & " could not be established...")
    End
    Resume
End Function


Private Function ValiAPI(Page As String, HttpMethod As String, Optional ByVal Data As String) As String

 'On Error GoTo ConnectionFail

    ' login if necessary
    Login

    requestUrl = valiUrl & "/" & Page

    'MsgBox (requestUrl)

    If HttpMethod = "GET" Then
        xobj.Open "GET", requestUrl, False
    ElseIf HttpMethod = "POST" Then
        xobj.Open "POST", requestUrl, False
    ElseIf HttpMethod = "PATCH" Then
        xobj.Open "PATCH", requestUrl, False
    Else
        MsgBox ("Method not allowed")
    End If


    xobj.SetRequestHeader "Authorization", token
    If Data = "" Then 'IsMissing(Data) Then
        xobj.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        xobj.Send
    Else
        xobj.SetRequestHeader "Content-Type", "application/json"
        'xobj.SetRequestHeader "Content-Length", "300"
        xobj.Send Data
    End If
    ValiAPI = xobj.ResponseText

    'MsgBox (xobj.ResponseText)

    Exit Function


ConnectionFail:
    MsgBox ("Connection to " & valiUrl & " could not be established...")
    End
    Resume

End Function

' Function which fetches all valis from the rest-API and provides them in a dictionary
Private Function getValiDict(Optional ByVal fetch_again As Boolean = False)

    ' only fetch from the server in the following cases:
    ' refresh was clicked --> fetch_again=True
    ' caching is disabled
    ' the valis have not been fetched yet in this session
    If (fetch_again = True) Or (cache_valis = False) Or (valis Is Nothing) Then

        Application.ScreenUpdating = False
        Dim vali_id As String
        Dim content(9) As String

        Set dict = CreateObject("Scripting.Dictionary")

        Dim Json As Object
        Dim Project As Object
        Set Json = JsonConverter.ParseJson(ValiAPI("rest/valis/?project=" & projectId, "GET"))
        Set Project = JsonConverter.ParseJson(ValiAPI("rest/project/" & projectId & "/", "GET"))

        For Each vali In Json
            If Not (vali("is_part_of_linking_matrix")) Then
                vali_id = vali("id")
                content(0) = vali("name")
                content(1) = vali("project")
                content(2) = Replace(vali("value"), ",", ".")
                content(3) = vali("unit")
                content(4) = Replace(vali("value"), ",", ".") & " " & vali("unit")
                content(5) = Replace(vali("margin_plus"), ",", ".") & "%"
                content(6) = Replace(vali("margin_minus"), ",", ".") & "%"
                content(7) = vali("path")
                ' Commented out because it's not clear if this was really needed, will leave it here, shouold the need to use it arise
                'If name_index Then
                '    content(7) = Replace(vali("path"), Left(vali("path"), name_index - 1), Project("name"))
                'Else
                '    content(7) = Project("name") & "." & vali("path")
                'End If
                'content(7) = Vali("minimum")
                'content(8) = Vali("maximum")
                dict(vali_id) = content
            End If
        Next vali

        Set getValiDict = dict

        Application.ScreenUpdating = True

    'in any other case, just return the valis dictionary
    Else
        Set getValiDict = valis
    End If
End Function

' Sub which is called from Ribbon
Sub CtrlInsertVali(ByVal Control As IRibbonControl)
    InsertVali
End Sub


Sub InsertVali()

    If Selection.Count <> 1 Then
        MsgBox ("Plese select one single cell, to insert a Vali.")
        End
    End If

    SetVariables

    'MsgBox (cache_valis)

    Dim str As String
    Set valis = getValiDict()
    ReDim id_array(valis.Count)

    v_items = valis.Items

    AddValiForm.ComboBox1.Clear

    ' save the correct keys in the right order for the dropdown-field in AddValiForm
    For i = 0 To valis.Count - 1
        id_array(i) = valis.Keys()(i)
        str = v_items(i)(0) & " (" & v_items(i)(4) & ")"
        AddValiForm.ComboBox1.AddItem str, i
    Next i

    AddValiForm.Show

    AddValiForm.ComboBox1.SetFocus

End Sub

' Sub which is called from Ribbon
Sub CtrlRefreshAllValis(ByVal Control As IRibbonControl)
    RefreshAllValis
End Sub

Sub RefreshAllValis()

    SetVariables

    Dim valiRange As Range
    Set valis = getValiDict(True)

    Set refsToDelete = CreateObject("System.Collections.ArrayList")

    CleanEmptyCells

    Set nms = ActiveWorkbook.Names

    vURL = GetSetting("ValiAddon", "Settings", "URL")

    Application.ScreenUpdating = False

    For n = 1 To nms.Count
        id = Replace(nms(n).Name, "V_", "")

        content = 2
        scrtip = ""

        If InStr(id, ".margin_plus") Then
            content = 5
            id = Replace(id, ".margin_plus", "")
            scrtip = " --> Margin +"
        End If

        If InStr(id, ".margin_minus") Then
            content = 6
            id = Replace(id, ".margin_minus", "")
            scrtip = " --> Margin -"
        End If



        If valis.Exists(id) And InStr(nms(n).Name, "V_") <> 0 Then

            Set valiRange = Range(nms(n).RefersTo)

            'update the "V_xxx" fields
            For Each rCell In valiRange.Cells
                rCell.FormulaR1C1 = valis(id)(content)
                If create_links = True Then
                    ActiveSheet.Hyperlinks.Add Anchor:=rCell, Address:=vURL & "/components/properties/vali/" & id & "/", ScreenTip:=valis(id)(0) & ": " & valis(id)(4) & scrtip
                End If
            Next
        ElseIf Not valis.Exists(id) And InStr(nms(n).Name, "V_") <> 0 Then
            For Each rCell In Range(nms(n).RefersTo)
                reference = Replace(nms(n).RefersTo, Left(nms(n).RefersTo, InStr(nms(n).RefersTo, "!")), "")
                reference = Replace(reference, "$", "")
                If MsgBox("Refresh Valis Failed: " & vbNewLine & "Vali on cell " & reference & " might have been deleted from Valispace or belongs to another project, retry refresh valis?", vbYesNo, "Confirm") = vbYes Then
                    RefreshAllValis
                    Exit Sub
                Else
                    If MsgBox("Would you like to clear cell " & reference & "?", vbYesNo, "Confirm") = vbYes Then
                        rCell.ClearContents
                        nms(n).Delete
                    End If
                    End
                End If
            Next
        End If
    Next

    'Deleting Broken refs after deleting the cell(s)
    DelRefsFromDeletedCells

    Application.ScreenUpdating = True

End Sub
Sub DelRefsFromDeletedCells()

    Set nms = ActiveWorkbook.Names

    For n = 1 To nms.Count
        If InStr(nms(n).RefersTo, "#REF!") <> 0 Then
            nms(n).Delete 'Deleting the invalid (#REF!) name range
            DelRefsFromDeletedCells 'Restarting Sub because nms.Count changed
            Exit Sub
        End If
    Next
End Sub
' Sub which is called from Ribbon
Sub CtrlValiSettings(ByVal Control As IRibbonControl)
    ValiSettings
End Sub

Sub ValiSettings()
    SettingsForm.Show
End Sub


Sub CtrlAddPushVali(ByVal Control As IRibbonControl)
    AddPushVali
End Sub

Sub AddPushVali()
    If Selection.Count <> 1 Then
        MsgBox ("Plese select one single cell, to define a push-Vali.")
        End
    End If


    SetVariables

    Dim str As String
    Set valis = getValiDict()
    ReDim id_array(valis.Count)

    v_items = valis.Items

    AddPushValiForm.ComboBox1.Clear

    ' save the correct keys in the right order for the dropdown-field in AddValiForm
    For i = 0 To valis.Count - 1
        id_array(i) = valis.Keys()(i)
        str = v_items(i)(0)
        AddPushValiForm.ComboBox1.AddItem str, i
    Next i

    AddPushValiForm.Show
End Sub

Sub CtrlPushValis(ByVal Control As IRibbonControl)
    PushValis
End Sub

Sub PushValis()
    SetVariables

    Set valis = getValiDict(Not cache_valis)

    Set pushDict = CreateObject("Scripting.Dictionary")

    Set nms = ActiveWorkbook.Names
    Dim JSONResponse As Object

    'Fixing Name Range when Cells|Rows|Columns are deleted, removing broken references
    CleanEmptyCells
    For n = 1 To nms.Count
        If Left(nms(n).Name, 2) = "P_" Then 'find all fields which are ready for push
            ValiID = Replace(nms(n).Name, "P_", "")
            ValiValue = Replace(Range(nms(n).RefersTo).Cells, ",", ".")
            Data = "{""formula"":""" & ValiValue & """}"
            Response = ValiAPI("rest/valis/" & ValiID, "PATCH", Data)
            'MsgBox (Response)
            Set JSONResponse = JsonConverter.ParseJson(Response)
            pushDict(JSONResponse("name")) = JSONResponse("value") & " " & JSONResponse("unit") & vbTab & "(before: " & valis(ValiID)(4) & ")"
        End If
    Next

    UpdatedValis = pushDict.Keys
    Message = "Uploaded the following values to Valispace" & vbNewLine & vbNewLine
    Dim i
    For i = 0 To pushDict.Count - 1
        Message = Message & UpdatedValis(i) & ": " & vbTab & pushDict(UpdatedValis(i)) & vbNewLine
    Next

    MsgBox (Message)

End Sub

Private Sub CleanEmptyCells()

    Dim valiRange As Range
    Dim cellRange As Range
    Dim resultRange As Range
    Set nms = ActiveWorkbook.Names

    For n = 1 To nms.Count

        ' Only clean Vali-generated Names
        If ((Left(nms(n).Name, 2) = "V_") Or (Left(nms(n).Name, 2) = "P_")) Then
            invalid_reference = InStr(nms(n).RefersTo, "#REF!")
            If invalid_reference <> 0 Then
                If MsgBox("Broken Reference found on " & nms(n).Name & "." & vbNewLine & "Do you want to delete the reference?(Process will not continue if the broken reference persists)", vbYesNo, "Confirm") = vbYes Then
                    nms(n).Delete
                    CleanEmptyCells
                    Exit Sub
                Else
                    MsgBox ("Execution Stopped, review the valis and references before refreshing or pushing valis")
                    End
                End If
            End If
            Set valiRange = Range(nms(n).RefersTo)
            For Each rCell In valiRange.Cells
                If (rCell.FormulaR1C1 = "") Then
                    Set cellRange = Range(rCell, rCell)
                    ' Delete this named range, if it was the last one. Else, just remove this cell from the range
                    If valiRange.Count = 1 Then

                        ' TBD: check whether the name was "V.xxx" and if so, check whether a "V_" exists and change the RefersTo to this cell. Else delete.

                        ActiveWorkbook.Names(n).Delete
                        CleanEmptyCells ' start again from the beginning
                        Exit Sub ' stop here, since the amount of named ranges has changed
                    Else
                        Set resultRange = RangeManipulation.Subtract(valiRange, cellRange)
                        nms(n).RefersTo = resultRange
                    End If
                End If
            Next
        End If
    Next
End Sub