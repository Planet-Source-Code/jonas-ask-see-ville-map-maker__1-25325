VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Load Map"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4140
      TabIndex        =   2
      Top             =   2700
      Width           =   915
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   3180
      TabIndex        =   1
      Top             =   2700
      Width           =   915
   End
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
      TabIndex        =   0
      Top             =   60
      Width           =   2115
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdLoad_Click()
    If File1.FileName = "" Then Exit Sub
    If Dirty Then
        Svar = MsgBox("The map is not saved. Continue?", vbOKCancel + vbInformation, GameTitle)
        If Not Svar = vbOK Then Me.Hide: Exit Sub
    End If
    
    Dirty = False
    FirstTime = False
    
    Me.Hide
    MapName = Mid(File1.FileName, 1, Len(File1.FileName) - 4)
    Form1.Caption = AppTitle & " - " & MapName
    OpenMap File1.path & "\" & File1.FileName
End Sub

Private Sub Form_Activate()
    Import.Hide
    frmNew.Hide
    frmSave.Hide
    File1.path = App.path & "\maps"
    File1.Refresh
End Sub

