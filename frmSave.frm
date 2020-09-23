VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Map"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2820
      Left            =   60
      Pattern         =   "*.map"
      TabIndex        =   3
      Top             =   60
      Width           =   2115
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FF8080&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   60
      MaxLength       =   80
      TabIndex        =   2
      Top             =   2880
      Width           =   2115
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4140
      TabIndex        =   1
      Top             =   2880
      Width           =   915
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Save"
      Height          =   375
      Left            =   3180
      TabIndex        =   0
      Top             =   2880
      Width           =   915
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdLoad_Click()
Dim Svar
Dim a
Dim Text As String
    Text = txtName.Text
    If Text = "" Then Exit Sub
    For a = 1 To Len(Text)
        If Mid(Text, a, 1) = "." Then MsgBox "Invalid name", vbOKOnly + vbCritical, GameTitle & " - ERROR": Exit Sub
    Next
    If Dir(App.path & "\maps\" & Text & ".map") <> "" Then
        Svar = MsgBox("Game already exsist. Overwrite?", vbOKCancel + vbInformation, GameTitle)
        If Not Svar = vbOK Then Exit Sub
    End If
    Me.Hide
    
    Dirty = False
    FirstTime = False
    
    MapName = Text
    Form1.Caption = AppTitle & " - " & MapName
    SaveMap App.path & "\maps\" & Text & ".map"
End Sub

Private Sub File1_Click()
    txtName.Text = Mid(File1.FileName, 1, Len(File1.FileName) - 4)
End Sub

Private Sub Form_Activate()
    frmNew.Hide
    frmLoad.Hide
    Import.Hide
    File1.path = App.path & "\maps"
    txtName.Text = MapName
    File1.Refresh
End Sub

