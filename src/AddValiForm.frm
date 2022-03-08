VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddValiForm
   Caption         =   "Add a Vali!"
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   150
   ClientWidth     =   7560
   OleObjectBlob   =   "AddValiForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddValiForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

' If user wrote bullshit, stop him
If Me.ComboBox1.ListIndex = -1 Then
    MsgBox ("Please Select one of the options.")
    Me.ComboBox1.SetFocus
Else



' Add a Vali
    vURL = GetSetting("ValiAddon", "Settings", "URL")
    create_links = GetSetting("ValiAddon", "Settings", "LINKS")

    Dim found As Boolean
    Dim valiRange As Range

    Set nms = ActiveWorkbook.Names

    're-create the vali-id from the combobox-id
    Dim id, autoname, autoid, vali_comment, extension, scrtip As String
    Dim content As Integer

    If Me.OptionValue = True Then
        extension = ""
        content = 2
        scrtip = ""
    ElseIf Me.OptionMarginPlus = True Then
        extension = ".margin_plus"
        content = 5
        scrtip = " --> Margin +"
    ElseIf Me.OptionMarginMinus = True Then
        extension = ".margin_minus"
        content = 6
        scrtip = " --> Margin -"
    End If

    id = id_array(Me.ComboBox1.ListIndex)
    autoid = "V_" & id & extension
    autoname = "V." & valis(id)(0) & extension
    vali_comment = valis(id)(7)

    found = False

    For n = 1 To nms.Count
        If nms(n).Name = autoid Then ' Name already existed --> update ID-cell-range, so that all values get updated when refreshed; leave named cell, so that it can be accessed in formulas
            found = True
            Set valiRange = Range(nms(n).RefersTo)
            Set valiRange = Union(valiRange, Selection)
            nms(n).RefersTo = valiRange
            For Each rCell In valiRange.Cells
                rCell.FormulaR1C1 = valis(id)(content)
            Next
        End If
    Next

    If found = False Then ' Name is new --> create both, named cell "V.Sat.mass" and ID-cell: "V_123"
        nms.add Name:=autoid, RefersTo:=ActiveCell
        nms.add Name:=autoname, RefersTo:=ActiveCell
        For n = 1 To nms.Count
            If nms(n).Name = autoid Or nms(n).Name = autoname Then
                nms(n).comment = vali_comment
            End If
        Next
        ActiveCell.Formula = valis(id)(content)
    End If

    If create_links = True Then
        ActiveSheet.Hyperlinks.add Anchor:=Selection, Address:=vURL & "/vali/" & id & "/", ScreenTip:=valis(id)(0) & ": " & valis(id)(4) & scrtip
    End If

    Me.Hide

End If

End Sub
