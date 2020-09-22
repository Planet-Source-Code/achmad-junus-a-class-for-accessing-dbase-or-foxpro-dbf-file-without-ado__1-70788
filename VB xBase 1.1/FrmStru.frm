VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmStru 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2640
      Left            =   75
      TabIndex        =   1
      Top             =   975
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4657
      _Version        =   393216
      BorderStyle     =   0
   End
   Begin VB.CommandButton btn 
      Caption         =   "OK"
      Height          =   315
      Left            =   6000
      TabIndex        =   0
      Top             =   3300
      Width           =   690
   End
End
Attribute VB_Name = "FrmStru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nFieldCount As Long

Public Sub Add_Field_Info(FieldName As String, FieldType As String, FieldWidth As String, FieldDecimal As String)
    nFieldCount = nFieldCount + 1
    Grid1.Rows = nFieldCount + 1
    Grid1.TextMatrix(nFieldCount, 0) = FieldName
    Grid1.TextMatrix(nFieldCount, 1) = FieldType
    Grid1.TextMatrix(nFieldCount, 2) = FieldWidth
    Grid1.TextMatrix(nFieldCount, 3) = FieldDecimal

End Sub

Public Property Let RecordCountInfo(lrecCount As Long)
    Dim RecordCount As Long
    Dim srn As String
    srn = "Number of records: " & vbTab & FormatNumber(lrecCount, 0)
    FontBold = False
    FontSize = 8.5
    CurrentX = 120
    CurrentY = 60 + (Me.TextHeight("x") * 3)
    Print srn
End Property
Public Property Let FileName_Info(Name As String)
    Dim L As Long
    Dim ssplit() As String
    Dim sfn As String
    Dim sff As String
    
    L = TextWidth(Name)
    If L + 135 > ScaleWidth Then
        ssplit = Split(Name, "\")
        sfn = ssplit(UBound(ssplit))
    Else
        sfn = Name
    End If
    Caption = App.Title & " - " & Name
    sff = "Structure for file: "
    FontBold = False
    CurrentX = (ScaleWidth - TextWidth(sff)) / 2
    CurrentY = 45
    Print sff
    FontBold = True
    CurrentX = (ScaleWidth - TextWidth(sfn)) / 2
    CurrentY = 105 + Me.TextHeight("x")
    Print sfn
End Property

Private Sub btn_Click()
    Cls
    Grid1.Clear
    Unload Me
End Sub

Private Sub Form_Activate()
    Grid1.Visible = True
    Line (Grid1.Left + 60, Grid1.Top + 60)-(Grid1.Left + Grid1.Width + 60, Grid1.Top + Grid1.Height + 60), RGB(160, 160, 160), BF
    Line (Grid1.Left - 15, Grid1.Top - 15)-(Grid1.Left + Grid1.Width, Grid1.Top + Grid1.Height), RGB(112, 112, 112), B
    Line (btn.Left + 60, btn.Top + 60)-(btn.Left + btn.Width + 60, btn.Top + btn.Height + 60), RGB(160, 160, 160), BF
End Sub

Private Sub Form_Initialize()
    nFieldCount = 0
End Sub

Private Sub Form_Load()
    Dim i As Long
    BackColor = &HE0E0E0
    With Grid1
        .Visible = False
        .BackColorBkg = RGB(196, 196, 196)
        .Rows = 2
        .Cols = 4
        .FixedCols = 0
        .FixedRows = 1
        .TextMatrix(0, 0) = "Field Name"
        .TextMatrix(0, 1) = "Data Type"
        .TextMatrix(0, 2) = "Max Width"
        .TextMatrix(0, 3) = "Decimal"
        .Font = "Tahoma 9pt Normal"
        .ColWidth(0) = Me.TextWidth("XXXXXXXXXMWW")
        For i = 1 To 3
            .ColWidth(i) = Me.TextWidth(UCase$(.TextMatrix(0, i))) + Me.TextWidth("MM")
        Next i
    End With
End Sub
