Attribute VB_Name = "Publics"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Const SRCAND = &H8800C6
Public Const SRCPAINT = &HEE0086
Public Const SRCCOPY = &HCC0020


Public Type Tile
 Ter As Integer
 TerType As Integer
 Build As Integer
 BuildType As Integer
End Type

Public MapName As String

Public OldX, OldY As Integer

Public Boarddata() As Tile

Public Dirty As Boolean
Public FirstTime As Boolean

Public Const AppTitle As String = "CityGame MapMaker"
Public Season As Integer
Public ISize As Integer
Public SelItem As Integer
Public Const Size As Integer = 20
Public Bredde As Integer
Public Hoyde As Integer
Public WstartX As Integer
Public WstartY As Integer
Public WBredde As Integer
Public WHoyde As Integer
Public Sub PaintMapSmall(PS, picMap As PictureBox)
    On Error Resume Next
    For Y = WstartY To WstartY + WHoyde + 5
        For X = WstartX To WstartX + WBredde + 5
            DoEvents
            Select Case Boarddata(X, Y).Ter
            Case 0: Form1.BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), vbBlue, BF
            Case 1
                If Boarddata(X, Y).BuildType = 0 Then
                    Form1.BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), RGB(6, 124, 12), BF
                Else
                    Form1.BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), RGB(60, 171, 60), BF
                End If
            Case Else: Form1.BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), vbBlack, BF
            End Select
        Next X
    Next Y
    picMap.Cls
    BitBlt picMap.hDC, 0, 0, Bredde * PS, Hoyde * PS, Form1.BufferMap.hDC, 0, 0, SRCCOPY
    picMap.Line (WstartX - 2, WstartY - 2)-Step(WBredde, WHoyde), vbWhite, B
    picMap.Refresh
End Sub
Public Sub PaintMap(PS, picMap As PictureBox)
    Form1.BufferMap.Cls
    On Error Resume Next
    For Y = 1 To Hoyde
        For X = 1 To Bredde
            DoEvents
            Select Case Boarddata(X, Y).Ter
            Case 0: Form1.BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), vbBlue, BF
            Case 1
                If Boarddata(X, Y).BuildType = 0 Then
                    Form1.BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), RGB(6, 124, 12), BF
                Else
                    Form1.BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), RGB(60, 171, 60), BF
                End If
            Case Else: Form1.BufferMap.Line ((X * PS) - 1, (Y * PS) - 1)-Step(PS - 1, PS - 1), vbBlack, BF
            End Select
        Next X
    Next Y
    picMap.Cls
    BitBlt picMap.hDC, 0, 0, Bredde * PS, Hoyde * PS, Form1.BufferMap.hDC, 0, 0, SRCCOPY
    picMap.Line (WstartX - 2, WstartY - 2)-Step(WBredde, WHoyde), vbWhite, B
    picMap.Refresh
End Sub
Public Function RndTall(Min, Max)
    Randomize
    RndTall = Int((Rnd * Max) + Min)
End Function

Public Function GetXY(XY)
    GetXY = Int(XY / Size)
End Function

Public Sub PaintGround()
Dim X, Y As Integer
    time1 = GetTickCount
    Form1.BufferG.Cls
    Form1.BufferM.Cls
    Form1.BufferS.Cls
    For Y = WstartY To WstartY + WHoyde
        For X = WstartX To WstartX + WBredde
            BitBlt Form1.BufferG.hDC, ((X - WstartX) * Size), ((Y - WstartY) * Size), Size, Size, Form1.PicGround.Item(Boarddata(X, Y).TerType).hDC, 0, 0, SRCCOPY
            If Boarddata(X, Y).BuildType <> 0 Then
                BitBlt Form1.BufferM.hDC, ((X - WstartX) * Size), ((Y - WstartY) * Size), Size, Size, Form1.PicmTree.Item(Boarddata(X, Y).BuildType - 1).hDC, 0, 0, SRCAND
                BitBlt Form1.BufferS.hDC, ((X - WstartX) * Size), ((Y - WstartY) * Size), Size, Size, Form1.PicTree.Item(Boarddata(X, Y).BuildType - 1).hDC, 0, 0, SRCPAINT
            End If
        Next X
    Next Y
    Form1.MainPic.Cls
    BitBlt Form1.MainPic.hDC, 0, 0, Bredde * Size, Hoyde * Size, Form1.BufferG.hDC, 0, 0, SRCCOPY
    BitBlt Form1.MainPic.hDC, 0, 0, Bredde * Size, Hoyde * Size, Form1.BufferM.hDC, 0, 0, SRCAND
    BitBlt Form1.MainPic.hDC, 0, 0, Bredde * Size, Hoyde * Size, Form1.BufferS.hDC, 0, 0, SRCPAINT
End Sub

Public Sub BuildTree(mX, mY)
    If Not Boarddata(mX, mY).Build = 0 Then Exit Sub
    If Boarddata(mX, mY).Ter = 0 Then Exit Sub
    
    Boarddata(mX, mY).Build = 0
    
    Select Case Boarddata(mX, mY).BuildType
    Case 0
        Boarddata(mX, mY).BuildType = 0 + RndTall(1, 6)
    Case 1 To 4
        Boarddata(mX, mY).BuildType = 4 + RndTall(1, 4)
    Case 5 To 8
        Boarddata(mX, mY).BuildType = 8 + RndTall(1, 4)
    Case 9 To 12
        Boarddata(mX, mY).BuildType = 12 + RndTall(1, 4)
    Case 13 To 16
        Exit Sub
    End Select
    
End Sub

Public Sub Determin(X, Y)
        If X < 1 Then Exit Sub
        If X > Bredde Then Exit Sub
        If Y < 1 Then Exit Sub
        If Y > Hoyde Then Exit Sub
        
        If X = 1 Then Exit Sub
        If X = Bredde Then Exit Sub
        If Y = 1 Then Exit Sub
        If Y = Hoyde Then Exit Sub
        
        Select Case SelItem
        Case 0
            If Boarddata(X, Y).Ter = 1 Then Exit Sub
            Boarddata(X, Y).Ter = 1
            Boarddata(X, Y).TerType = RndTall(1, 4)
        Case 1
            If Boarddata(X, Y).Ter = 0 Then Exit Sub
            Boarddata(X, Y).Ter = 0
            Boarddata(X, Y).TerType = 0
            Boarddata(X, Y).Build = 0
            Boarddata(X, Y).BuildType = 0
        Case 2
            BuildTree X, Y
        Case 3
            If Boarddata(X, Y).BuildType = 0 Then Exit Sub
            Boarddata(X, Y).Build = 0
            Boarddata(X, Y).BuildType = 0
        End Select
        OldX = X
        OldY = Y
End Sub

Public Sub DoTheShit(glX, glY)
        Select Case ISize
        Case 1
            Determin glX, glY
        Case 2
            Determin glX, glY
            Determin glX - 1, glY
            Determin glX, glY + 1
            Determin glX + 1, glY
            Determin glX, glY - 1
        Case 3
            Determin glX, glY
            Determin glX - 1, glY
            Determin glX - 2, glY
            Determin glX - 1, glY + 1
            Determin glX, glY + 1
            Determin glX, glY + 2
            Determin glX + 1, glY + 1
            Determin glX + 1, glY
            Determin glX + 2, glY
            Determin glX + 1, glY - 1
            Determin glX, glY - 1
            Determin glX, glY - 2
            Determin glX - 1, glY - 1
        Case 4
            Determin glX, glY
            Determin glX - 1, glY
            Determin glX - 2, glY
            Determin glX - 3, glY
            Determin glX - 1, glY + 1
            Determin glX - 2, glY + 1
            Determin glX - 3, glY + 1
            Determin glX - 1, glY + 2
            Determin glX - 2, glY + 2
            Determin glX - 1, glY + 3
            Determin glX, glY + 1
            Determin glX, glY + 2
            Determin glX, glY + 3
            Determin glX + 1, glY + 3
            Determin glX + 1, glY + 2
            Determin glX + 2, glY + 2
            Determin glX + 1, glY + 1
            Determin glX + 2, glY + 1
            Determin glX + 3, glY + 1
            Determin glX + 1, glY
            Determin glX + 2, glY
            Determin glX + 3, glY
            Determin glX + 1, glY - 1
            Determin glX + 2, glY - 1
            Determin glX + 3, glY - 1
            Determin glX + 1, glY - 2
            Determin glX + 2, glY - 2
            Determin glX + 1, glY - 3
            Determin glX, glY - 1
            Determin glX, glY - 2
            Determin glX, glY - 3
            Determin glX - 1, glY - 1
            Determin glX - 2, glY - 1
            Determin glX - 3, glY - 1
            Determin glX - 1, glY - 2
            Determin glX - 2, glY - 2
            Determin glX - 1, glY - 3
        End Select
End Sub


Public Sub ImportMap(path)
Dim X, Y
    Import.picMap.AutoSize = True
    Import.picMap.Picture = LoadPicture(path)
    Import.picMap.AutoSize = False
    
    Bredde = Import.picMap.ScaleWidth
    Hoyde = Import.picMap.ScaleHeight
    
    ReDim Boarddata(1 To Bredde, 1 To Hoyde)
    
    For Y = 1 To Hoyde
        For X = 1 To Bredde
            Select Case Import.picMap.Point(X - 1, Y - 1)
            Case 0
                Boarddata(X, Y).Ter = 0
                Boarddata(X, Y).Ter = 0
            Case Else
                Boarddata(X, Y).Ter = 1
                Boarddata(X, Y).TerType = RndTall(1, 4)
            End Select
        Next X
    Next Y

    WstartX = 1
    WstartY = 1
    Form1.VScroll.Max = Hoyde - WHoyde
    Form1.HScroll.Max = Bredde - WBredde
    Form1.VScroll.Value = 1
    Form1.HScroll.Value = 1
    PaintGround
    
    PaintMap 1, Form1.picMM
End Sub
Public Sub SaveMap(path)
Dim FileNum
Dim TempOUT As String
Dim Lengde
Dim Tempdata
    
    For X = 1 To Bredde
        If Boarddata(X, 1).Ter = 1 Then
            Boarddata(X, 1).Ter = 0
            Boarddata(X, 1).TerType = 0
            Boarddata(X, 1).Build = 0
            Boarddata(X, 1).BuildType = 0
        End If
        If Boarddata(X, Hoyde).Ter = 1 Then
            Boarddata(X, Hoyde).Ter = 0
            Boarddata(X, Hoyde).TerType = 0
            Boarddata(X, Hoyde).Build = 0
            Boarddata(X, Hoyde).BuildType = 0
        End If
    Next X
    
    For Y = 1 To Hoyde
        If Boarddata(1, Y).Ter = 1 Then
            Boarddata(1, Y).Ter = 0
            Boarddata(1, Y).TerType = 0
            Boarddata(1, Y).Build = 0
            Boarddata(1, Y).BuildType = 0
        End If
        If Boarddata(Bredde, Y).Ter = 1 Then
            Boarddata(Bredde, Y).Ter = 0
            Boarddata(Bredde, Y).TerType = 0
            Boarddata(Bredde, Y).Build = 0
            Boarddata(Bredde, Y).BuildType = 0
        End If
    Next Y
    
    
    If Dir(path) <> "" Then Kill path
    
    FileNum = FreeFile
    Open path For Random As FileNum Len = 10
    Put FileNum, 1, Bredde
    Put FileNum, 2, Hoyde
    a = 2
    For Y = 1 To Hoyde
        For X = 1 To Bredde
            a = a + 1

            TempOUT = Boarddata(X, Y).Ter
            
            TempOUT = TempOUT & Boarddata(X, Y).TerType
            
            Tempdata = Boarddata(X, Y).BuildType
            Lengde = Len(Tempdata)
            TempOUT = TempOUT & Lengde & Boarddata(X, Y).BuildType
            
            Put FileNum, a, TempOUT
        Next X
    Next Y
    Close FileNum
    
    PaintGround
    PaintMap 1, Form1.picMM
    
End Sub

Public Sub OpenMap(path)
Dim FileNum
Dim TempIN As String
Dim Nowpos
Dim Lengde As Integer

    FileNum = FreeFile
    Open path For Random As FileNum Len = 10
    
    Get FileNum, 1, Bredde
    Get FileNum, 2, Hoyde
    
    ReDim Boarddata(1 To Bredde, 1 To Hoyde)
    
    a = 2
    For Y = 1 To Hoyde
        For X = 1 To Bredde
            a = a + 1
            Get FileNum, a, TempIN
            Nowpos = 1
            
            Boarddata(X, Y).Ter = Mid(TempIN, 1, 1)
            Boarddata(X, Y).TerType = Mid(TempIN, 2, 1)
            Lengde = Mid(TempIN, 3, 1)
            Boarddata(X, Y).BuildType = Mid(TempIN, 4, Lengde)
        Next X
    Next Y
    Close FileNum
    
    
    WstartX = 1
    WstartY = 1
    Form1.VScroll.Max = Hoyde - WHoyde
    Form1.HScroll.Max = Bredde - WBredde
    Form1.VScroll.Value = 1
    Form1.HScroll.Value = 1
    PaintGround
    
    PaintMap 1, Form1.picMM
End Sub

Public Sub SetSeason(Num)
    With Form1
    Select Case Num
    Case 1
        For a = 1 To 4
            BitBlt .PicGround.Item(a).hDC, 0, 0, Size, Size, .PicSpring.Item(a - 1).hDC, 0, 0, SRCCOPY
        Next a
    Case 2
        For a = 1 To 4
            BitBlt .PicGround.Item(a).hDC, 0, 0, Size, Size, .PicSummer.Item(a - 1).hDC, 0, 0, SRCCOPY
        Next a
    Case 3
        For a = 1 To 4
            BitBlt .PicGround.Item(a).hDC, 0, 0, Size, Size, .PicAutumn.Item(a - 1).hDC, 0, 0, SRCCOPY
        Next a
    Case 4
        For a = 1 To 4
            BitBlt .PicGround.Item(a).hDC, 0, 0, Size, Size, .PicWinter.Item(a - 1).hDC, 0, 0, SRCCOPY
        Next a
    End Select
    End With
End Sub
