VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddPushValiForm
   Caption         =   "Add a Vali!"
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   150
   ClientWidth     =   7185
   OleObjectBlob   =   "AddPushValiForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddPushValiForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
' Add a Vali
    vURL = GetSetting("ValiAddon", "Settings", "URL")
    create_links = GetSetting("ValiAddon", "Settings", "LINKS")

    Dim found As Boolean
    Dim valiRange As Range

    Set nms = ActiveWorkbook.Names

    're-create the vali-id from the combobox-id
    Dim id, autoid As String

    id = id_array(Me.ComboBox1.ListIndex)
    autoid = "P_" & id


    nms.Add Name:=autoid, RefersTo:=ActiveCell
    For n = 1 To nms.Count
        If nms(n).Name = autoid Then
            nms(n).Comment = Me.ComboBox1.Value
        End If
    Next
    If create_links = True Then
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=vURL & "/project/" & valis(id)(1) & "/components/properties/vali/" & id & "/", ScreenTip:=valis(id)(0)
    End If

    Me.Hide

End Sub
