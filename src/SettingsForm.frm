VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsForm
   Caption         =   "ValiAddon Settings"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   150
   ClientWidth     =   5205
   OleObjectBlob   =   "SettingsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Me.TextBox_URL.Text = GetSetting("ValiAddon", "Settings", "URL", "")
    Me.TextBox_PROJECTID.Text = GetSetting("ValiAddon", "Settings", "ProjectID", "")
    Me.TextBox_USER.Text = GetSetting("ValiAddon", "Settings", "User", "")
    Me.TextBox_PW.Text = GetSetting("ValiAddon", "Settings", "PW", "")

    Me.CheckBox1.Value = CBool(GetSetting("ValiAddon", "Settings", "LINKS", True))
    Me.CheckBox2.Value = CBool(GetSetting("ValiAddon", "Settings", "CACHE", True))
End Sub

Private Sub CommandButton1_Click()
    SaveSetting "ValiAddon", "Settings", "URL", Me.TextBox_URL.Value
    SaveSetting "ValiAddon", "Settings", "ProjectID", Me.TextBox_PROJECTID.Value

    Dim links, cache As Integer

    links = CInt(Me.CheckBox1.Value)
    cache = CInt(Me.CheckBox2.Value)

    SaveSetting "ValiAddon", "Settings", "LINKS", links
    SaveSetting "ValiAddon", "Settings", "CACHE", cache

    SaveSetting "ValiAddon", "Settings", "User", Me.TextBox_USER.Value
    SaveSetting "ValiAddon", "Settings", "PW", Me.TextBox_PW.Value

    Me.Hide
End Sub


