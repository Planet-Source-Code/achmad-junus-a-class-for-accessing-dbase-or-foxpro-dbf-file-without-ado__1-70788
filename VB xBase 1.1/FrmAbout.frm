VERSION 5.00
Begin VB.Form FrmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0E0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   5325
      TabIndex        =   0
      Top             =   2775
      Width           =   990
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' picture size is 50 x 50 pixel

Private sTitle As String            ' App.Title
Private sVersion As String          ' Version
Private sDescription1 As String     ' Description, 1st line
Private sDescription2 As String     ' Description, 2nd line
Private sCopyright As String        ' Copyright Info
Private sExtra1 As String           ' Extra information 1st line
Private sExtra2 As String           ' Extra information 2nd line


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim sp As String
    DrawBackGround
    FontBold = True
    FontSize = 13
    cPrint 0, 4, sTitle, 2, &HFFEEFF, True, 15, 80
    
    FontSize = 9
    cPrint 0, 25, sVersion, 2, &HFFFF00, False, 15, 80
    FontBold = False
    FontItalic = False
    FontSize = 6.8
    cPrint 6, 135, sExtra1, 0, &HDDFFFF, True
    FontItalic = True
    cPrint 6, 155, sExtra2, 0, &HDDFFFF, True
    FontSize = 8.1
    cPrint 0, 105, sCopyright, 2, &HF0C0F0, False, 15, 15
    
    FontItalic = False
    FontBold = False
    FontSize = 9
    cPrint 80, 60, sDescription1, 2, &HFF00&, True, 20, 80
    cPrint 80, 75, sDescription2, 2, &HFF00&, True, 20, 80
    PaintPicture LoadPicture(App.Path & "\Image01.jpg"), 165, 165
End Sub

Private Sub Form_Load()
    Caption = App.Title
End Sub

' the X, Y, LMargin and RMargin are in Pixel (not Twip)
Private Sub cPrint(X As Long, Y As Long, Text As String, Alignment As Long, Optional TextColor As Long = &H0, Optional Shadow As Boolean = False, Optional LMargin As Long = 2, Optional RMargin As Long = 2) ' Alignment: 0=Left, 1=Center, 2=Right
    Dim L As Long
    
    L = Len(Text)
    If Shadow Then
        ForeColor = &HC0B0A0
        If Alignment = 1 Then       ' Align Center
            CurrentX = ((ScaleWidth - ((RMargin - LMargin) * 15) - TextWidth(Text)) / 2) + 7
        ElseIf Alignment = 2 Then   ' Align Right
            CurrentX = ScaleWidth - (RMargin * 15) - TextWidth(Text) - 8 ' Reserve 1 pixel
        Else                        ' Allign Left
            CurrentX = (X * 15) + 7
        End If
        CurrentY = (Y * 15) + 7
        Print Text
        
        ForeColor = TextColor
        If Alignment = 1 Then       ' Align Center
            CurrentX = (ScaleWidth - ((RMargin - LMargin) * 15) - TextWidth(Text)) / 2
        ElseIf Alignment = 2 Then   ' Align Right
            CurrentX = ScaleWidth - (RMargin * 15) - TextWidth(Text) - 15    ' Reserve 1 pixel
        Else                        ' Allign Left
            CurrentX = X * 15
        End If
        CurrentY = Y * 15
        Print Text
    
    Else
        ForeColor = TextColor
        If Alignment = 1 Then       ' Align Center
            CurrentX = (ScaleWidth - ((RMargin - LMargin) * 15) - TextWidth(Text)) / 2
        ElseIf Alignment = 2 Then   ' Align Right
            CurrentX = ScaleWidth - (RMargin * 15) - TextWidth(Text) - 15    ' Reserve 1 pixel
        Else                        ' Allign Left
            CurrentX = X * 15
        End If
        CurrentY = Y * 15
        Print Text
    End If
    
    
End Sub

Public Property Let Title(StringTitle As String)
    sTitle = StringTitle
End Property

Public Property Let Extra2(sX As String)
    sExtra2 = sX
End Property

Public Property Let Extra1(sX As String)
    sExtra1 = sX
End Property

Public Property Let CopyrightInfo(scpr As String)
    sCopyright = scpr
End Property

Public Property Let Description2(strDescription As String)
    sDescription2 = strDescription
End Property

Public Property Let Description1(strDescription As String)
    sDescription1 = strDescription
End Property

Public Property Let FileVersion(StringVersion As String)
    sVersion = StringVersion
End Property

' creating gradient
Private Sub DrawBackGround()
    Dim i As Long
    Dim dX As Long, dY As Long
    Dim cR As Long, cG As Long, cB As Long
    Dim Interval As Long
    
    Interval = 1995 / 195
    dX = ScaleWidth
    cR = 250
    cG = 255
    cB = 255
    For i = 0 To 1995 Step Interval
        Line (0, i)-(dX, i + Interval), RGB(cR, cG, cB), BF
        cR = 5 + cR - (cR / (Interval * 2))
        cG = 1 + cG - (cG / (Interval * 2))
        cB = cB - 1
    Next i
    
    Interval = (ScaleHeight - 1995) / 60
    For i = 1996 To ScaleHeight Step Interval
        Line (0, i)-(dX, i + Interval), RGB(cR, cG, cB), BF
        cR = cR + 1
        cG = cG + 2
        cB = cB + 2
    Next i
    
    ' draw a blind
    For i = 1996 To ScaleHeight Step Interval + 15
        Line (0, i)-(ScaleWidth, i + 2), RGB(cR, cG, cB), BF
        cR = cR - 1
        cG = cG - 2
        cB = cB - 2
    Next i
    
    ' Draw other Box in the midle, from 1425 to 1995
    Line (60, 1455)-(ScaleWidth - 75, 1935), &HD0B0D0, B
    Line (75, 1470)-(ScaleWidth - 90, 1920), &HF0D0F0, B
    Line (ScaleWidth - 75, 1455)-(ScaleWidth - 75, 1935), &HB000B0, B
    Line (ScaleWidth - 90, 1470)-(ScaleWidth - 90, 1920), &HC0A0C0, B
    Line (75, 1920)-(ScaleWidth - 90, 1920), &HC0A0C0, B
    Line (60, 1935)-(ScaleWidth - 75, 1935), &HB000B0, B
    DoEvents
    
End Sub
