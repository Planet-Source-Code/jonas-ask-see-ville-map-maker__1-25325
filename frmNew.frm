VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Map"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Map Name"
      Height          =   675
      Left            =   60
      TabIndex        =   11
      Top             =   120
      Width           =   2115
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   120
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   1875
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Proportions "
      Height          =   1215
      Left            =   2580
      TabIndex        =   8
      Top             =   120
      Width           =   1995
      Begin VB.TextBox txtBredde 
         Height          =   315
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "50"
         Top             =   300
         Width           =   675
      End
      Begin VB.TextBox txtHoyde 
         Height          =   315
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "50"
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Map Width"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Map Height"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   780
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Default Ground"
      Height          =   975
      Left            =   60
      TabIndex        =   7
      Top             =   900
      Width           =   2115
      Begin VB.OptionButton Opt2 
         Caption         =   "Water"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Land"
         Height          =   195
         Left            =   60
         TabIndex        =   4
         Top             =   600
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdMake 
      Caption         =   "Make"
      Height          =   375
      Left            =   2700
      TabIndex        =   5
      Top             =   1560
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3660
      TabIndex        =   6
      Top             =   1560
      Width           =   915
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdMake_Click()
    On Error GoTo Error1
    If txtBredde < 17 Then Exit Sub
    If txtHoyde < 17 Then Exit Sub
    If txtBredde > 999 Then Exit Sub
    If txtHoyde > 999 Then Beep: Exit Sub
    If txtName = "" Then Exit Sub
    
    If Dirty Then
        Svar = MsgBox("The map is not saved. Continue?", vbOKCancel + vbInformation, GameTitle)
        If Not Svar = vbOK Then Me.Hide: Exit Sub
    End If
    
    Dirty = False
    FirstTime = True
    
    Hoyde = txtHoyde.Text
    Bredde = txtBredde.Text
    WBredde = 16
    WHoyde = 16
    WstartX = 1
    WstartY = 1
    
    MapName = txtName.Text
    
    Form1.Caption = AppTitle & " - " & MapName
    
    Me.Hide
    
    ReDim Boarddata(1 To Bredde, 1 To Hoyde)
    Select Case Opt1.Value
    Case True
        For Y = 2 To Hoyde - 1
            For X = 2 To Bredde - 1
                Boarddata(X, Y).Ter = 1
                Boarddata(X, Y).TerType = RndTall(1, 4)
            Next X
        Next Y
    Case False
        'Trenger ikke gjøre noe, står på 0 fra før
    End Select
    
    Form1.VScroll.Max = Hoyde - WHoyde
    Form1.HScroll.Max = Bredde - WBredde
    Form1.VScroll.Value = 1
    Form1.HScroll.Value = 1
    PaintGround
    
    Form1.picMM.Cls
    PaintMap 1, Form1.picMM
    
Exit Sub
Error1:
    txtBredde.Text = 50
    txtHoyde.Text = 50
    Beep
End Sub

Private Sub Form_Activate()
    frmSave.Hide
    frmLoad.Hide
    Import.Hide
    
    txtName = "New Map"
    txtBredde.Text = 50
    txtHoyde.Text = 50
    Opt2.Value = True
End Sub

Private Sub txtBredde_GotFocus()
    txtBredde.SelStart = 0
    txtBredde.SelLength = Len(txtBredde.Text)
End Sub
Private Sub txthoyde_GotFocus()
    txtHoyde.SelStart = 0
    txtHoyde.SelLength = Len(txtHoyde.Text)
End Sub
Private Sub txtname_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName.Text)
End Sub
