VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   2055
   ClientTop       =   630
   ClientWidth     =   12660
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkMode        =   1  'Source
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8265
   ScaleWidth      =   12660
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command4 
      Caption         =   "About..."
      Height          =   300
      Left            =   11175
      TabIndex        =   76
      Top             =   435
      Width           =   1335
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   300
      Left            =   3000
      TabIndex        =   57
      Top             =   435
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   6480
      ScaleHeight     =   330
      ScaleWidth      =   3015
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   405
      Width           =   3015
      Begin VB.CommandButton cmdRandom2 
         Caption         =   "Add random data"
         Height          =   300
         Left            =   1440
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   25
         Width           =   1575
      End
      Begin VB.CommandButton cmdCreatField 
         Caption         =   "Create Fields"
         Height          =   300
         Left            =   0
         TabIndex        =   3
         Top             =   25
         Width           =   1335
      End
   End
   Begin VB.PictureBox PicDemoBtn 
      BackColor       =   &H8000000A&
      Height          =   5430
      Left            =   0
      ScaleHeight     =   5370
      ScaleWidth      =   1785
      TabIndex        =   37
      Top             =   1800
      Width           =   1845
      Begin VB.CommandButton cmdSortDemo 
         Caption         =   "Sort Table"
         Height          =   360
         Left            =   45
         TabIndex        =   71
         Top             =   1425
         Width           =   1695
      End
      Begin VB.CommandButton cmdDistinctDemo 
         Caption         =   "Get Distinct Values"
         Height          =   360
         Left            =   45
         TabIndex        =   56
         Top             =   990
         Width           =   1695
      End
      Begin VB.CommandButton cmdFindDemo 
         Caption         =   "Find Text"
         Height          =   360
         Left            =   45
         TabIndex        =   39
         Top             =   540
         Width           =   1695
      End
      Begin VB.CommandButton cmdSearchDemo 
         Caption         =   "Search Record(s)"
         Height          =   360
         Left            =   45
         TabIndex        =   38
         Top             =   90
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog ComDil1 
      Left            =   5925
      Top             =   375
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdStructure 
      Caption         =   "Display Structure"
      Height          =   300
      Left            =   4320
      TabIndex        =   21
      Top             =   435
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11280
      TabIndex        =   17
      Top             =   7815
      Width           =   1215
   End
   Begin VB.CommandButton cmdRandom 
      Caption         =   "Add random data"
      Height          =   300
      Left            =   7920
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   435
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   12540
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1275
      Width           =   12540
      Begin VB.TextBox txtGoRecNo 
         Height          =   300
         Left            =   7200
         TabIndex        =   11
         Text            =   "txtGoRecNo"
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Go"
         Height          =   300
         Left            =   8400
         TabIndex        =   12
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1480
         TabIndex        =   9
         Top             =   0
         Width           =   400
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1080
         TabIndex        =   8
         Top             =   0
         Width           =   400
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   525
         TabIndex        =   7
         Top             =   0
         Width           =   400
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   400
      End
      Begin VB.Label Line2 
         BorderStyle     =   1  'Fixed Single
         Height          =   45
         Left            =   -120
         TabIndex        =   16
         Top             =   405
         Width           =   12735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Go to Rec. #: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   10
         Top             =   45
         Width           =   1335
      End
      Begin VB.Label lblPos 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblPos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   345
         Left            =   1920
         TabIndex        =   15
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdCloseFile 
      Caption         =   "Close File"
      Height          =   300
      Left            =   9660
      TabIndex        =   13
      Top             =   435
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open file..."
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   435
      Width           =   1335
   End
   Begin VB.CommandButton cmdNewFile 
      Caption         =   "New file..."
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   435
      Width           =   1335
   End
   Begin VB.PictureBox PGrid 
      BorderStyle     =   0  'None
      Height          =   4140
      Left            =   1950
      ScaleHeight     =   4140
      ScaleWidth      =   10590
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   3225
      Width           =   10590
      Begin VB.HScrollBar HScroll1 
         Height          =   240
         Left            =   0
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   3900
         Width           =   10365
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3915
         Left            =   10350
         Max             =   0
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Bindings        =   "Form1.frx":0442
         Height          =   3915
         Left            =   0
         TabIndex        =   73
         Top             =   0
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   6906
         _Version        =   393216
         BackColorBkg    =   -2147483636
         GridColor       =   -2147483648
         GridColorFixed  =   -2147483640
         Redraw          =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         ScrollBars      =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   1950
      TabIndex        =   58
      Top             =   1770
      Visible         =   0   'False
      Width           =   10590
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1275
         TabIndex        =   65
         Text            =   "Text3"
         Top             =   375
         Width           =   4815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "To File Name:"
         Height          =   285
         Left            =   0
         TabIndex        =   64
         Top             =   375
         Width           =   1290
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "Sort"
         Height          =   285
         Left            =   9375
         TabIndex        =   68
         Top             =   375
         Width           =   1065
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   6900
         TabIndex        =   67
         Text            =   "Combo1"
         Top             =   375
         Width           =   2340
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Finish"
         Height          =   240
         Left            =   6660
         TabIndex        =   69
         Top             =   750
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Field:  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   6000
         TabIndex        =   66
         Top             =   375
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sort function demo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   1650
         TabIndex        =   70
         Top             =   0
         Width           =   6495
      End
      Begin VB.Label lblWrite 
         BackColor       =   &H00E0E0E0&
         Caption         =   "#"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5325
         TabIndex        =   63
         Top             =   750
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Records writen:"
         Height          =   240
         Index           =   2
         Left            =   4050
         TabIndex        =   62
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Sorting complete"
         Height          =   240
         Index           =   1
         Left            =   2625
         TabIndex        =   61
         Top             =   750
         Width           =   1365
      End
      Begin VB.Label lblRead 
         BackColor       =   &H00E0E0E0&
         Caption         =   "#"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1275
         TabIndex        =   60
         Top             =   750
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Records read: "
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   59
         Top             =   750
         Width           =   1140
      End
   End
   Begin VB.PictureBox PicFindText 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1650
      ScaleHeight     =   735
      ScaleWidth      =   10920
      TabIndex        =   30
      Top             =   1770
      Width           =   10920
      Begin VB.CommandButton cmdFindText1 
         Caption         =   "Find first"
         Height          =   315
         Left            =   8580
         TabIndex        =   41
         Top             =   375
         Width           =   1080
      End
      Begin VB.CommandButton cmdFindText 
         Caption         =   "Find next"
         Height          =   315
         Left            =   9705
         TabIndex        =   36
         Top             =   375
         Width           =   1080
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Match Case"
         Height          =   255
         Left            =   7275
         TabIndex        =   35
         Top             =   385
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   34
         Top             =   360
         Width           =   6225
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Find: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   75
         TabIndex        =   33
         Top             =   375
         Width           =   690
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Find Text function Demo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1920
         TabIndex        =   32
         Top             =   10
         Width           =   6495
      End
   End
   Begin VB.PictureBox PicDistinct 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   1650
      ScaleHeight     =   855
      ScaleWidth      =   10815
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   1770
      Visible         =   0   'False
      Width           =   10815
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   55
         Text            =   "Combo1"
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton cmdDistinct 
         Caption         =   "Search"
         Height          =   315
         Left            =   3840
         TabIndex        =   54
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Field:  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   120
         TabIndex        =   53
         Top             =   520
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Search Distinct Values"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   1920
         TabIndex        =   52
         Top             =   0
         Width           =   6495
      End
   End
   Begin VB.PictureBox PicSearch 
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   1800
      ScaleHeight     =   1200
      ScaleWidth      =   10770
      TabIndex        =   22
      Top             =   1770
      Width           =   10770
      Begin VB.OptionButton Option1 
         Caption         =   "BEGINWITH"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   9255
         TabIndex        =   50
         Top             =   410
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   ">="
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   8415
         TabIndex        =   49
         Top             =   410
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   7695
         TabIndex        =   48
         Top             =   410
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "<="
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   6855
         TabIndex        =   47
         Top             =   410
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6135
         TabIndex        =   46
         Top             =   410
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "<>"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5295
         TabIndex        =   45
         Top             =   410
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4575
         TabIndex        =   44
         Top             =   410
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdSearchAll 
         Caption         =   "Search All"
         Height          =   315
         Left            =   9465
         TabIndex        =   29
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdSeacrhNext 
         Caption         =   "Search Next"
         Height          =   315
         Left            =   8265
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search First"
         Height          =   315
         Left            =   7065
         TabIndex        =   27
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1050
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   840
         Width           =   5910
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1050
         TabIndex        =   24
         Text            =   "Combo1"
         Top             =   360
         Width           =   2265
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "operator: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   3450
         TabIndex        =   43
         Top             =   405
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Search function demo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   1770
         TabIndex        =   31
         Top             =   15
         Width           =   6495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Search:  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Field:  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Label HiddenLabel 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lW"
      Height          =   255
      Left            =   -2000
      TabIndex        =   42
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label lblFileName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblFileName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   120
      TabIndex        =   20
      Top             =   825
      Width           =   12375
   End
   Begin VB.Label lblResult 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblResult"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   45
      TabIndex        =   19
      Top             =   7845
      Width           =   11175
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Record Manager"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   360
      Left            =   45
      TabIndex        =   18
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   12645
   End
   Begin VB.Label Line1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Left            =   0
      TabIndex        =   14
      Top             =   1200
      Width           =   12615
   End
   Begin VB.Menu mnuXBase 
      Caption         =   "XBase"
      Visible         =   0   'False
      Begin VB.Menu mnuSortTest 
         Caption         =   "Sort Test"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const PTOP As Single = 2400
Private Const MAXLONG As Long = 2147483647
Private Const MAXCURRENCY As Currency = 922337203685477@
Private Const MAXDOUBLE As Double = 1.79769313486231E+308
Private Const MAXDECIMAL As Variant = 7.92281625142643E+28

Dim WithEvents DBFTable As VB_xBase
Attribute DBFTable.VB_VarHelpID = -1
Dim xTimer As New CTiming

Dim TableInfo() As Long
Dim FieldInfo() As String

Dim TableFileName As String

Dim MaxView As Long

Dim GridMode As Byte
Dim TableReady As Boolean
Dim CurrentNavPos As Long   ' Used for clicking |<,<,>, and >|

'// Mousewheel
Private Const PM_REMOVE = &H1
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private bCancel As Boolean
Private Const WM_MOUSEWHEEL = 522
'// Mousewheel END


'// Extend VScroll to enable scrolling larger than 32767
'   The trick is scale the 32 bit position to 16 bit scroll position
Private ScrollMin As Long
Private ScrollMax As Long
Private ScrollPos As Long
Private ScrollPageSize As Long
Private ScrollScale As Long
Private LargeScrollScale As Double
Private VScrollOldVal As Long
Private RescaleScroll As Long       ' 1 or -1
'// END OF Extending VScroll & HScroll

Private ArrGridView() As Long
Private BrowseMode As Boolean




Private Declare Sub GetMem1 Lib "Msvbvm60" (ByVal Addr As Long, RetVal As Byte)
Private Declare Sub GetMem2 Lib "Msvbvm60" (ByVal Addr As Long, RetVal As Integer)
Private Declare Sub GetMem4 Lib "Msvbvm60" (ByVal Addr As Long, RetVal As Long)
Private Declare Sub PutMem1 Lib "Msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Private Declare Sub PutMem2 Lib "Msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Sub PutMem4 Lib "Msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)



Private Sub InitTable()
    Set DBFTable = New VB_xBase
End Sub

Private Sub cmdBrowse_Click()
    Dim i As Long
    
    Frame1.Visible = False
    PicSearch.Visible = False
    PicFindText.Visible = False
    PicDistinct.Visible = False
    lblResult.Caption = ""
    
    If TableReady Then
        If Not BrowseMode Then
            CurrentNavPos = 1
            BrowseMode = True
            txtGoRecNo.Text = ""
        End If
        
        If GridMode <> 1 Then AdjustGridTitle
        GridMode = 0
        lblResult.Caption = ""
        DoEvents
        
        ScrollRange TableInfo(0), MaxView
        
        ReDim ArrGridView(0 To TableInfo(0) - 1)
        
        For i = 0 To TableInfo(0) - 1
            ArrGridView(i) = i + 1
        Next i
        
        If RescaleScroll = 1 Then RescaleScroll = -1
        ScrollPos = CurrentNavPos - 1
        VScroll1.Value = (ScrollPos \ ScrollScale)
        
        ViewInGrid 0, ArrGridView
        
    End If
End Sub

Private Sub cmdCloseFile_Click()
    ResetVscroll
    cmdOpen.Enabled = True
    cmdNewFile.Enabled = True
    EmptyLabels
    EmptyTxtBox
    Picture1.Visible = False
    Frame1.Visible = False
    ComDil1.FileName = ""
    TableFileName = ""
    HideGrid
    Set DBFTable = Nothing
    
    Erase TableInfo
    Erase FieldInfo
    Grid1.Clear
    Grid1.Rows = 2
    GridMode = 0
    DoEvents
    PicDemoBtn.Visible = False
    PicFindText.Visible = False
    PicSearch.Visible = False
    PicDistinct.Visible = False
    TableReady = False
    DoEvents
End Sub

Private Sub cmdCreatField_Click()
    Dim Hasil As Long
    
    ' in field name, all case will converted to uppercase, and all space will converted to underscore (_)
'    DBFTable.CreateField "SortKey", "C", 10
'    DBFTable.CreateField "OtherData", "C", 89
    
    ' =====================================
    ' CREATING SOME FIELDS FOR DEMO PROJECT
    ' =====================================
    
    DBFTable.CreateField "SomeNumber", "I"          ' Data type Long
    DBFTable.CreateField "Char_1", "C", 15          ' Data type String, 15 characters max
    DBFTable.CreateField "Float 1", "F", 16, 2      ' Data type Numeric string float, 16 character max with 2 decimal digit
    DBFTable.CreateField "Date 1", "D"              ' Data type date string (YYYYMMDD)
    DBFTable.CreateField "char 2", "C", 50          ' Data type String, 50 characters max
    DBFTable.CreateField "Logical", "L"             ' Data type Logical
    DBFTable.CreateField "double 1", "B", , 6       ' Data type Double, decimal will rounded to 6 digit
    DBFTable.CreateField "Numeric 1", "N", 13, 0    ' Data type Numeric string, 12 characters max without decimal
    DBFTable.CreateField "currency 1", "Y"          ' Data type Currency, default decimal is 4 digit
    DBFTable.CreateField "Int 32", "i"              ' Data type Long
    DBFTable.CreateField "numeric 2", "N", 20, 2    ' Data type Numeric string, 18 characters max with 2 decimal digit
    '/// NEW DATA TYPE THAT AVAILABLE ONLY FOR THIS CLASS:
    '/// -------------------------------------------------
    '/// ** Data type "Z" and "S" **
    '///
    '/// ** These data type will work only with this VB class **
    '///
'    DBFTable.CreateField "Byte_1", "Z"              ' /**Extra**/ Data type 8 bit unsigned integer (Byte)
'    DBFTable.CreateField "Integers", "S"            ' /**Extra**/ Data type 16 bit signed integer (Integer)
    '///
'    DBFTable.CreateField "FINDtest", "C", 35        ' Data type String, 25 characters max
    
    
    DoEvents
    
    ' Finish creating fields
    Hasil = DBFTable.BuildFile
    
    DoEvents
    
    If Hasil > 0 Then
        tDBFGetInfo
    Else
        MsgBox "Error"
        Set DBFTable = Nothing
        Exit Sub
    End If
    
    ' Fill combo box with field names
    FillCombo
    Picture1.Visible = True
    PicDemoBtn.Visible = True
    Grid1.Visible = True
    cmdStructure_Click
End Sub

Private Sub cmdExit_Click()
    Set DBFTable = Nothing
    DoEvents
    End
End Sub

Private Sub cmdFindText_Click()
    Dim MatchCase As Boolean
    Dim Hasil As Long
    Dim X As Double
    
    DoEvents
    Me.MousePointer = 11
    MatchCase = IIf(Check1.Value = 1, True, False)
    xTimer.TReset
    Hasil = DBFTable.FindText(Text1.Text, MatchCase)
    X = xTimer.Elapsed
    Me.MousePointer = 0
    If Hasil > 0 Then
        DBFTable.MoveTo Hasil
        DisplayData Hasil
        lblResult.Caption = "  Found in " & Format(X, "Standard") & " ms."
'        MsgBox "Found!!!"
    Else
        lblResult.Caption = "  Not Found. EOF in " & Format(X, "Standard") & " ms."
        MsgBox "Not found"
    End If
End Sub

Private Sub cmdFindText1_Click()
    Dim MatchCase As Boolean
    Dim Hasil As Long
    Dim X As Double
    
    Me.MousePointer = 11
    DoEvents
    MatchCase = IIf(Check1.Value = 1, True, False)
    xTimer.TReset
    Hasil = DBFTable.FindText(Text1.Text, MatchCase, 1)
    X = xTimer.Elapsed
    Me.MousePointer = 0
    If Hasil > 0 Then
        DBFTable.MoveTo Hasil
        DisplayData Hasil
        lblResult.Caption = "  Found in " & Format(X, "Standard") & " ms."
'        MsgBox "Found!!!"
    Else
        lblResult.Caption = "  Not Found. EOF in " & Format(X, "Standard") & " ms."
        MsgBox "Not found"
    End If
End Sub

Private Sub cmdLast_Click()
    CurrentNavPos = DBFTable.RecordCount
    DisplayData CurrentNavPos
    lblResult.Caption = ""
End Sub

Private Sub cmdNewFile_Click()
    ComDil1.ShowSave
    If Len(ComDil1.FileName) > 0 Then
        TableFileName = ComDil1.FileName
        lblFileName.Caption = " " & TableFileName
        Picture2.Visible = True
        Picture1.Visible = False
        
        InitTable
        DBFTable.CreateTable TableFileName
        Me.Caption = TableFileName
    Else
        lblFileName.Caption = " "
        Picture2.Visible = False
        cmdOpen.Enabled = True
        Picture1.Visible = False
        PicDemoBtn.Visible = False
    End If
End Sub

'Private Sub cmdNext_Click()
'    If CurrentNavPos < TableInfo(0) Then CurrentNavPos = CurrentNavPos + 1
'    DisplayData CurrentNavPos
'    lblResult.Caption = ""
'End Sub


Private Sub cmdNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If CurrentNavPos < TableInfo(0) Then CurrentNavPos = CurrentNavPos + 1
    DisplayData CurrentNavPos
    lblResult.Caption = ""
End Sub

Private Sub cmdOpen_Click()
    ComDil1.ShowOpen
    If Len(ComDil1.FileName) > 0 Then
        TableFileName = ComDil1.FileName
        lblFileName.Caption = " " & TableFileName
        Picture2.Visible = False
        Picture1.Visible = True
        
        InitTable
        DBFTable.DBFOpenFile (TableFileName)
        
        Me.Caption = TableFileName
        PicDemoBtn.Visible = True
        'cmdSearchDemo_Click
        cmdBrowse_Click
    
    Else
        lblFileName.Caption = " "
        Picture2.Visible = False
        cmdNewFile.Enabled = True
        Picture1.Visible = False
        PicDemoBtn.Visible = False
    End If
End Sub

Private Sub cmdPrev_Click()
    If CurrentNavPos > 1 Then CurrentNavPos = CurrentNavPos - 1
    DisplayData CurrentNavPos
    lblResult.Caption = ""
End Sub

Private Sub cmdFirst_Click()
    lblResult.Caption = ""
    CurrentNavPos = 1
    DisplayData CurrentNavPos
End Sub

Private Sub cmdRandom_Click()
    On Error GoTo SALAH
    Dim i, j As Long
    Dim Buanyak As String
    Dim Buanyaaak As Long
    Dim X As Double
    Dim RandomRecord() As String
    
    If TableReady Then
        Buanyak = InputBox("Enter number of random records to add: ")
        If IsNumeric(Buanyak) Then
            Buanyaaak = CLng(Abs(Buanyak))
            If Buanyaaak = 0 Then
                Buanyaaak = 1
            End If
        Else
            Exit Sub
        End If
        
        Me.MousePointer = 11
        xTimer.TReset
        
        ReDim RandomRecord(0 To TableInfo(1) - 1)
        
        Randomize Timer
        For j = 1 To Buanyaaak
            
            For i = 0 To UBound(RandomRecord)
                RandomString FieldInfo(i, 1), RandomRecord(i), CLng(FieldInfo(i, 2)), CLng(FieldInfo(i, 3))
            Next i
            
            ' -------------------------------- add a record from string array
            DBFTable.Append RandomRecord
            
            ' Display caption in label
            If j Mod 1000 = 0 Then
                lblResult.Caption = FormatNumber(j, 0)
                DoEvents
            End If
        Next j
        Erase RandomRecord
        
        tDBFGetInfo
        X = xTimer.Elapsed / 1000
        MsgBox Format(Buanyaaak, "#,###") & " random records has been added in " & Format(X, "Standard") & " seconds"
        
        Me.MousePointer = 0
        lblResult.Caption = ""
    End If
    Exit Sub

SALAH:
    Err.Clear
    Me.MousePointer = 0
    lblResult.Caption = ""
End Sub

Private Sub cmdRandom2_Click()
    On Error GoTo SALAH2
    Dim i, j As Long
    Dim Buanyak As String
    Dim Buanyaaak As Long
    Dim X As Double
    Dim RandomRecord() As String
    
    If TableReady Then
        Buanyak = InputBox("Enter number of random records to add: ")
        If IsNumeric(Buanyak) Then
            Buanyaaak = CLng(Abs(Buanyak))
            If Buanyaaak = 0 Then
                Buanyaaak = 1
            End If
        Else
            Exit Sub
        End If
        
        Me.MousePointer = 11
        
        xTimer.TReset
        ReDim RandomRecord(0 To TableInfo(1) - 1)
        DBFTable.AutoUpdateHeader = True
        
        ' Initialize random number with current date/time
        Randomize ((Year(Date) * 10000) + (Month(Date) * 100) + Day(Date)) + _
                    (Hour(time) * 10000000) + (Minute(time) * 100000) + (Second(time) * 1000)
        
        For j = 1 To Buanyaaak
'            RandomRecord(0) = CStr(TableInfo(0) + j)    ' AutoNumber
            For i = 0 To UBound(RandomRecord)
                RandomString FieldInfo(i, 1), RandomRecord(i), CLng(FieldInfo(i, 2)), CLng(FieldInfo(i, 3))
            Next i
            
            ' -------------------------------- add a record from string array
            DBFTable.Append RandomRecord
            
            ' Display caption in label
            If j Mod 1000 = 0 Then
                lblResult.Caption = FormatNumber(j, 0)
                DoEvents
            End If
        Next j
        
        Erase RandomRecord
        DBFTable.AutoUpdateHeader = True
        DBFTable.UpdateHeader
        tDBFGetInfo
        X = xTimer.Elapsed / 1000
        MsgBox Format(Buanyaaak, "#,###") & " random records has been added in " & Format(X, "#0.000000") & " seconds"
        
        Me.MousePointer = 0
        lblResult.Caption = ""
        Picture2.Visible = False
    End If
    Exit Sub
SALAH2:
    Err.Clear
    Me.MousePointer = 0
    lblResult.Caption = ""
End Sub

Private Sub cmdSearch_Click()
    Dim i As Long
    Dim SearchResult As Long
    Dim X As Double
    Dim mOperator As Integer
    
    lblResult.Caption = ""
    If GridMode <> 1 Then AdjustGridTitle
    DoEvents
    BrowseMode = False
    If Len(Text2.Text) > 0 Then
        ' select which operator will used?
        For i = 0 To 6
            If Option1(i).Value = True Then
                mOperator = 5001 + i
                Exit For
            End If
        Next i
        
        
        xTimer.TReset
        
        SearchResult = DBFTable.SearchRecord(Combo1(0).List(Combo1(0).ListIndex), Text2.Text, 1, mOperator)   ' 1: Start search from begining of file, -1: start from current position
        X = xTimer.Elapsed
        If SearchResult > 0 Then
            DBFTable.MoveTo SearchResult
            DisplayData SearchResult
            lblResult.Caption = "  Found in " & Format(X, "Standard") & " ms."
            ' make a noise
            Beep
        Else
            lblResult.Caption = "  NOT FOUND, and EOF reached in " & Format(X, "Standard") & " ms."
        End If
        
    End If
End Sub

Private Sub cmdSearchAll_Click()
    Dim Hasil As Long
    Dim i, j As Long
    Dim X, Y, Z As Double
    Dim strData() As String
    Dim mOperator As Integer
    Dim arrHasil() As Long

    If GridMode <> 1 Then AdjustGridTitle
    GridMode = 0
    BrowseMode = True
    DoEvents
    ' select which operator will used?
    For i = 0 To 6
        If Option1(i).Value = True Then
            mOperator = 5001 + i
            Exit For
        End If
    Next i
    
    xTimer.TReset
    Hasil = DBFTable.SearchAllRecord(Combo1(0).List(Combo1(0).ListIndex), Text2.Text, arrHasil, mOperator)
    X = xTimer.Elapsed
    
    If Hasil > 0 Then
        
        ReDim ArrGridView(LBound(arrHasil) To UBound(arrHasil))
        For i = LBound(arrHasil) To UBound(arrHasil)
            ArrGridView(i) = arrHasil(i)
        Next i
        ScrollRange Hasil, MaxView
        ViewInGrid 0, ArrGridView
        
        lblResult.Caption = "  Found " & Format(Hasil, "#,###") & " records in " & Format$(X, "Standard") & " ms."
        
'        xTimer.TReset
'        strData = DBFTable.GetRow2DA(arrHasil)  ' Get 2D array of string for records data
'        Z = xTimer.Elapsed
'        lblResult.Caption = "  Found " & Format(Hasil, "#,###") & " records in " & Format$(x, "#,##0.0000") & " + " & Format$(Z, "#,##0.0000") & " msec"
'        DoEvents
'
'        xTimer.TReset
'        MaxView = 325000 \ (TableInfo(1) + 1)
'        MaxView = IIf(Hasil > MaxView - 1, MaxView - 1, Hasil)
'
'        Grid1.Visible = False
'
'        Grid1.Rows = MaxView + 1
'        For i = 0 To MaxView - 1
'            Grid1.TextMatrix((i + 1), 0) = CStr(arrHasil(i))
'            For j = LBound(strData, 1) To UBound(strData, 1)
'                ' Copy text to grid cells
'                Grid1.TextMatrix((i + 1), (j + 1)) = strData(j, i)
'            Next j
'        Next i
'
'        Grid1.Visible = True
'        Y = xTimer.Elapsed
'        lblResult.Caption = lblResult.Caption & ", plus " & Format(Y, "#,##0.0000") & " seconds to populate grid. =TOTAL= " & Format$(x + Y + Z, "#,##0.0000") & " msec"
'        lblPos.Caption = ""
    Else
        MsgBox "No records match with criteria"
    End If
    Me.MousePointer = 0
    ' clean up
    Erase strData
    Erase arrHasil
End Sub

Private Sub cmdDistinct_Click()
    Dim Hasil As Long
    Dim i, j As Long
    Dim X, Y As Double
    Dim strData() As String
    Dim mOperator As Integer
    Dim arrHasil() As Long
    Dim GridFontWidth As Long
    Dim Compares As Double

    Me.MousePointer = 11
    DoEvents
    BrowseMode = False
    ' select which operator will used?
    For i = 0 To 6
        If Option1(i).Value = True Then
            mOperator = 5001 + i
            Exit For
        End If
    Next i
    
    xTimer.TReset       ' calculate time for query
    Compares = DBFTable.GetDistinctValues(Combo1(1).List(Combo1(1).ListIndex), strData)
    X = xTimer.Elapsed
    
    Hasil = UBound(strData) + 1
    
    lblResult.Caption = "  Found " & FormatNumber(Hasil, 0) & " distinct values after doing " & FormatNumber(Compares, 0) & " comparisons, in " & Format(X, "Standard") & " ms."
    lblPos.Caption = ""
    Me.MousePointer = 0
    DoEvents
    
'    MaxView = IIf(Hasil > 324998, 324998, Hasil)
'
'    xTimer.TReset       ' Calculate time for populate grid
'
'    Grid1.Visible = False
'
'    Grid1.Cols = 1
'    GridFontWidth = CalculateFontWidth("lW")
'    Grid1.ColWidth(0) = CLng(FieldInfo(Combo1(1).ListIndex, 2)) * GridFontWidth
'    If Grid1.ColWidth(0) < GridFontWidth * 2 Then Grid1.ColWidth(0) = GridFontWidth * 2
'    Grid1.FixedCols = 0
'    Grid1.Rows = MaxView + 1
'    Grid1.TextMatrix(0, 0) = Trim(Combo1(1).List(Combo1(1).ListIndex))
'
'    For i = 0 To MaxView - 1
'        Grid1.TextMatrix(i + 1, 0) = strData(i)
'    Next i
'
'    Grid1.Visible = True
'
'    Y = xTimer.Elapsed
'
'    lblResult.Caption = lblResult.Caption & " and " & Format(Y, "#,##0.0000") & " msec. to populate grid. TOTAL= " & Format(X + Y, "#,##0.0000") & " msec."
'
'    Me.MousePointer = 0
'    GridMode = 0
End Sub


Private Sub cmdSearchDemo_Click()
    Frame1.Visible = False
    PicSearch.Visible = True
    PicFindText.Visible = False
    PicDistinct.Visible = False
    lblResult.Caption = ""
    txtGoRecNo.Text = ""
End Sub

Private Sub cmdFindDemo_Click()
    Frame1.Visible = False
    PicSearch.Visible = False
    PicFindText.Visible = True
    PicDistinct.Visible = False
    lblResult.Caption = ""
    txtGoRecNo.Text = ""
End Sub

Private Sub cmdDistinctDemo_Click()
    PicDistinct.Visible = True
    PicFindText.Visible = False
    PicSearch.Visible = False
    Frame1.Visible = False
    lblResult.Caption = ""
    txtGoRecNo.Text = ""
End Sub

Private Sub cmdSeacrhNext_Click()
    Dim i As Long
    Dim HasilCari As Long
    Dim X As Double
    Dim mOperator As Integer
    
    lblResult.Caption = ""
    DoEvents
    BrowseMode = False
    If Len(Text2.Text) > 0 Then
        ' select which operator will used?
        For i = 0 To 6
            If Option1(i).Value = True Then
                mOperator = 5001 + i
                Exit For
            End If
        Next i
        xTimer.TReset
        HasilCari = DBFTable.SearchRecord(Combo1(0).List(Combo1(0).ListIndex), Text2.Text, -1, mOperator)   ' -1: Start search from current position
        X = xTimer.Elapsed
        If HasilCari > 0 Then
            DBFTable.MoveTo HasilCari
            DisplayData HasilCari
            lblResult.Caption = "  Found in " & Format(X, "Standard") & " ms."
            
            ' Make a noise
            Beep
        Else
            lblResult.Caption = "  NOT FOUND, and EOF reached in " & Format(X, "Standard") & " ms."
            MsgBox "-- END OF FILE --" & vbCrLf & vbCrLf & "Search next did not find anything", vbInformation, "Search Next"
        End If
    End If

End Sub


Private Sub cmdSort_Click()
    xTimer.TReset
    DBFTable.SortData Text3.Text, Combo1(2).List(Combo1(2).ListIndex)
    lblResult.Caption = "Sorting in " & xTimer.sElapsed
End Sub

Private Sub cmdSortDemo_Click()
    Frame1.Visible = True
    Label5(1).Visible = False
    lblRead.Caption = ""
    lblWrite.Caption = ""
    txtGoRecNo.Text = ""
End Sub

Private Sub cmdStructure_Click()
    Dim i As Long, j As Long
    Dim GridFontWidth As Single     ' Average width
    
    Dim f As New FrmStru
    
    If TableReady Then
        With f
            .FileName_Info = TableFileName
            For i = 0 To UBound(FieldInfo, 1)
                .Add_Field_Info FieldInfo(i, 0), FieldInfo(i, 1), FieldInfo(i, 2), FieldInfo(i, 3)
            Next i
            .RecordCountInfo = TableInfo(0)
            .Show 1, Me
        End With
    End If
'    If TableReady Then
'        BrowseMode = False
'        GridFontWidth = CalculateFontWidth("Mpr:W ?1wixX")
'        Grid1.Clear
'
'        Grid1.Cols = 4
'        Grid1.Rows = UBound(FieldInfo, 1) + 9
'        Grid1.FixedCols = 0
'        Grid1.FixedRows = 1
'
'        ' Adjust column width
'        Grid1.ColWidth(0) = 26 * GridFontWidth
'        Grid1.ColWidth(1) = 11 * GridFontWidth
'        Grid1.ColWidth(2) = 11 * GridFontWidth
'        Grid1.ColWidth(3) = 11 * GridFontWidth
'
'        Grid1.TextMatrix(1, 0) = " Records count:"
'        Grid1.TextMatrix(2, 0) = " Field count:"
'        Grid1.TextMatrix(3, 0) = " Record size:"
'        For i = 1 To 3
'            Grid1.TextMatrix(i, 2) = Format$(TableInfo(i - 1), "#,##0")
'        Next i
'
'        Grid1.TextMatrix(5, 0) = " FIELD INFO:"
'        Grid1.TextMatrix(6, 0) = " Field Name"
'        Grid1.TextMatrix(7, 0) = String$(100, "-")
'        Grid1.TextMatrix(6, 1) = " Data Type"
'        Grid1.TextMatrix(7, 1) = String$(100, "-")
'        Grid1.TextMatrix(6, 2) = " Width"
'        Grid1.TextMatrix(7, 2) = String$(100, "-")
'        Grid1.TextMatrix(6, 3) = " Decimal"
'        Grid1.TextMatrix(7, 3) = String$(100, "-")
'
'
'        For i = 0 To UBound(FieldInfo, 1)
'            Grid1.TextMatrix(8 + i, 0) = " " & CStr(i + 1) & ".  " & FieldInfo(i, 0)
'            Grid1.TextMatrix(8 + i, 1) = FieldInfo(i, 1)
'            Grid1.TextMatrix(8 + i, 2) = FieldInfo(i, 2)
'            Grid1.TextMatrix(8 + i, 3) = FieldInfo(i, 3)
'        Next i
'        GridMode = 0
'        BrowseMode = True
'    End If
End Sub

Private Sub Command1_Click()
    If IsNumeric(txtGoRecNo.Text) Then
        DBFTable.MoveTo CLng(txtGoRecNo.Text)
        DisplayData CLng(txtGoRecNo.Text)
    End If
    lblResult.Caption = ""
End Sub

Private Sub Command2_Click()
    Frame1.Visible = False
    Label5(1).Caption = ""
'    cmdSort.Visible = False
    cmdDistinct.Visible = True
End Sub

Private Sub Command3_Click()
    Dim sFilenama As String
    ComDil1.ShowSave
    sFilenama = ComDil1.FileName
    If Len(sFilenama) < 1 Then
        Text3.Text = TableFileName & ".sorted"
    Else
        Text3.Text = sFilenama
    End If
End Sub

Private Sub Command4_Click()
    Dim fa As New FrmAbout
    With fa
        .Title = App.Title
        .FileVersion = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
        .Description1 = "Demo project for xBase Class Database"
        .Description2 = "The Class is licensed under: LGPL"
        .CopyrightInfo = "Copyright " & Chr(169) & "2008  -  Achmad Junus"
        .Extra1 = "Still under construction. Please leave some comments. I need that to make this things get better."
        .Extra2 = "seruling_m4l4m@yahoo.com"
        .Show 1, Me
    End With
End Sub

Private Sub DBFTable_FileBuilt()
    StatusIsReady
End Sub

Private Sub DBFTable_FileOpened()
    StatusIsReady
End Sub

Private Sub StatusIsReady()
    TableReady = True
    cmdOpen.Enabled = False
    cmdNewFile.Enabled = False
    ShowGrid
    
    tDBFGetInfo
    FillCombo
End Sub

Private Function CalculateFontWidth(strX As String) As Single
    HiddenLabel.FontName = Grid1.CellFontName
    HiddenLabel.FontBold = Grid1.CellFontBold
    HiddenLabel.FontItalic = Grid1.CellFontItalic
    HiddenLabel.FontSize = Grid1.CellFontSize
    HiddenLabel.Caption = strX
    CalculateFontWidth = (HiddenLabel.Width) / (Len(strX))     ' Get average font width with this brutal method. The width will use for adjust grid column width
    HiddenLabel.Visible = False
End Function

Private Sub DBFTable_RecRead(ByVal Total As Long)
    lblRead.Caption = Total
End Sub

Private Sub DBFTable_RecWriten(ByVal Total As Long)
    lblWrite.Caption = Total
End Sub

Private Sub DBFTable_SortCompleted()
    Label5(1).Caption = "Sort completed"
End Sub

Private Sub DBFTable_SortStarted()
    Label5(1).Caption = "Sorting..."
    Label5(1).Visible = True
End Sub

Private Sub Form_Load()
    Me.Caption = App.Title
    lblCaption.Caption = "Demo project: Class for manipulating .dbf files - without ADO -"
    EmptyLabels
    EmptyTxtBox
    Picture2.Visible = False
    Picture1.Visible = False
    PicDemoBtn.Visible = False
    PicFindText.Visible = False
    PicSearch.Visible = False
    HideGrid
    ComDil1.InitDir = App.Path
    ComDil1.DefaultExt = ".dbf"
    ComDil1.Filter = "dBase, FoxPro, XBase Files (*.dbf)|*.dbf| All files (*.*)|*.*"
    BrowseMode = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set DBFTable = Nothing
End Sub

Private Sub Form_Resize()
    If ScaleHeight > 8000 Then
        ' resize grid height
        PGrid.Height = ScaleHeight - 4365
        Grid1.Height = PGrid.Height - 240
        VScroll1.Height = PGrid.Height - 240
        HScroll1.Top = Grid1.Height
        MaxView = (Grid1.Height \ Grid1.RowHeight(1)) - 1
        ' resize picdemobtn
        PicDemoBtn.Height = ScaleHeight - 2835
        ' move lblresult
        lblResult.Top = ScaleHeight - 420
        ' move cmdexit
        cmdExit.Top = lblResult.Top
    End If
    
    If ScaleWidth > 10000 Then
        Picture1.Width = ScaleWidth + 300
        lblCaption.Width = ScaleWidth - (lblCaption.Left * 2)
        Line1.Width = ScaleWidth + 150
        Line2.Width = Picture1.Width
        lblFileName.Width = ScaleWidth - (lblFileName.Left * 2)
        ' resize grid width
        PGrid.Width = ScaleWidth - PGrid.Left - 350
        Grid1.Width = PGrid.Width - 240
        HScroll1.Width = PGrid.Width - 240
        VScroll1.Left = Grid1.Width
        ' resize lblresult
        lblResult.Width = ScaleWidth - 1215 - 285
        ' move cmdexit
        cmdExit.Left = ScaleWidth - 1215 - 150
    End If
    
    If ScrollMax > MaxView Then
        ScrollRange (ScrollMax + ScrollPageSize - 1), MaxView
        VScroll1_Change
    End If
    AdjustHscroll
End Sub

Private Sub Grid1_Click()
    CurrentNavPos = Grid1.RowSel + ScrollPos
    lblPos = "Rec. " & FormatNumber(ArrGridView(ScrollPos + Grid1.RowSel - 1), 0) ' & " of " & FormatNumber(TableInfo(0), 0)
End Sub

Private Sub Grid1_GotFocus()
    bCancel = False
    DoEvents
    ProcessMessages
End Sub

Private Sub Grid1_LostFocus()
    bCancel = True
End Sub

Private Sub HScroll1_Change()
    Grid1.Left = 0 - (HScroll1.Value * 150)
End Sub

Private Sub HScroll1_Scroll()
    Grid1.Left = 0 - (HScroll1.Value * 150)
End Sub

Private Sub mnuSortTest_Click()
    Frame1.Visible = True
'    cmdDistinct.Visible = False
    
'    cmdSort.Visible = True
    
    Label5(1).Visible = False
    lblRead.Caption = ""
    lblWrite.Caption = ""
End Sub

Private Sub Text1_Change()
    lblResult.Caption = ""
End Sub

Private Sub Text2_Change()
    lblResult.Caption = ""
End Sub

Private Sub txtGoRecNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txtGoRecNo.Text) Then
            DBFTable.MoveTo CLng(txtGoRecNo.Text)
            DisplayData CLng(txtGoRecNo.Text)
        End If
        lblResult.Caption = ""
    End If
End Sub

Private Sub EmptyLabels()
    lblFileName.Caption = ""
    lblResult.Caption = ""
    lblPos.Caption = ""
End Sub

Private Sub EmptyTxtBox()
    txtGoRecNo.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Combo1(0).Clear
    Combo1(1).Clear
    Combo1(2).Clear
End Sub

Private Sub tDBFGetInfo()
    ' Get File Info
    DBFTable.GetFileInfo TableInfo
    
    ' Get structure
    DBFTable.GetFieldInfo FieldInfo
    
End Sub

Private Sub FillCombo()
    Dim i  As Long, j As Long
    
    Combo1(0).Clear
    Combo1(1).Clear
    Combo1(2).Clear
    For i = 0 To TableInfo(1) - 1       ' Number of fields
        For j = 0 To 2
            Combo1(j).AddItem FieldInfo(i, 0), i
        Next j
    Next i
    
    ' Adjust Grid width and layout for viewing data
    AdjustGridTitle
End Sub

Private Sub DisplayData(RecNo As Long)
    Dim strData() As String
    Dim i As Long
    
    If GridMode <> 1 Then AdjustGridTitle
    
    If Not BrowseMode Then
        BrowseMode = True
        ScrollRange TableInfo(0), MaxView
        
        ReDim ArrGridView(0 To TableInfo(0) - 1)
        
        For i = 0 To TableInfo(0) - 1
            ArrGridView(i) = i + 1
        Next i
    End If
    
    If RescaleScroll = 1 Then RescaleScroll = -1
    If ScrollMax > RecNo Then
        ScrollPos = RecNo - 1
        VScroll1.Value = (ScrollPos \ ScrollScale)
        ViewInGrid RecNo - 1, ArrGridView
    Else
        ScrollPos = ScrollMax
        VScroll1.Value = VScroll1.Max
        ViewInGrid TableInfo(0) - MaxView, ArrGridView
    End If
    
    lblPos = "Rec. " & FormatNumber(ArrGridView(RecNo - 1), 0) ' & " of " & FormatNumber(TableInfo(0), 0)
    DoEvents
End Sub


Private Sub RandomString(sType As String, ByRef VarName As String, MaxWidth As Long, DecimalPlace As Long)
    Dim TempRnd As Double
    Dim i As Long, j As Long
    Static k As Long
    
    Dim LebarDef As Long
    Dim mLebar As Long
    
    Dim Huruf As String
    Dim sS1 As String
    Dim sS2 As String
    
    Dim RandomNumber1 As Long
    Dim Lebarx As Long
    
    
    If k < 48 Then k = 48
    
    Select Case sType
        Case "N", "F"
            TempRnd = (Rnd * 3141592621#) / 3141592621#
            mLebar = IIf(DecimalPlace = 0, MaxWidth, MaxWidth - (DecimalPlace + 1))
            LebarDef = IIf(mLebar > 13, 13, mLebar)
            If LebarDef > 0 Then
                sS1 = IIf(LebarDef < 10, CStr(Int((10 ^ (LebarDef * Rnd)) * TempRnd)), CStr(CDec(Int((10 ^ (LebarDef * Rnd)) * TempRnd)) + 1))
            Else
                sS1 = ""
            End If
            
            ' Decimal place
            If DecimalPlace > 0 Then
                sS2 = IIf(DecimalPlace < 10, CStr(Int((((10 ^ DecimalPlace) - 1) * TempRnd) + 1)), CStr(CDec(Int((((10 ^ DecimalPlace) - 1) * TempRnd) + 1))))
            End If
            
            VarName = IIf(DecimalPlace = 0, sS1, sS1 & "." & sS2)
            
        Case "I"
            TempRnd = (Rnd * 3141592621#) / 3141592621#
            VarName = CStr(Int(((2 ^ (30 * Rnd)) * TempRnd) + 1))
        
        Case "Z"
            VarName = CStr(Int((256 * Rnd)))
        
        Case "S"
            VarName = CStr(Int((32767 * Rnd)))
        
        Case "Y"
            sS1 = CStr(CCur(Int((10 ^ ((10 * Rnd) + 1)) * ((Rnd * 3141592621#) / 3141592621#))))
            
            sS2 = CStr(Int(((10 ^ (4 * Rnd)) - 1) * Rnd))
            VarName = sS1 & "." & sS2
            
        Case "B"
            RandomNumber1 = Int((20000000 * Rnd) + 1)
            VarName = Trim$(Str((((Rnd * 3141592621#) / 3141592621#) * 100000) ^ (Rnd * 2)))
            If RandomNumber1 Mod 2 = 0 Then ' Negative value as well
                VarName = "-" & VarName
            End If
        Case "D"
            RandomNumber1 = Int((12 * Rnd) + 1) ' Month
            Select Case RandomNumber1           ' Day
                Case 1, 3, 5, 7, 8, 10, 12
                    VarName = IIf(Int(2019 * Rnd) = 1988, " ", RandomNumber1 & "/" & Int((31 * Rnd) + 1) & "/2007")
                Case 4, 6, 9, 11
                    VarName = IIf(Int(2019 * Rnd) = 1988, " ", RandomNumber1 & "/" & Int((30 * Rnd) + 1) & "/2007")
                Case Else
                    VarName = IIf(Int(2019 * Rnd) = 1988, " ", RandomNumber1 & "/" & Int((28 * Rnd) + 1) & "/2007")
            End Select
            
        Case "L"
            VarName = IIf((Int((20000000 * Rnd) + 1)) Mod 2 = 0, "T", "F")
        
        Case "C"
            Huruf = "0123456789 ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz,."
            LebarDef = MaxWidth
            VarName = Space(LebarDef)
            
            RandomNumber1 = CLng(Int(((LebarDef \ 2) * Rnd) + (LebarDef \ 2)))
            mLebar = 1
            Select Case RandomNumber1
                Case Is < 2
                    k = IIf(k > 122, 48, k + 1)
                    Mid(VarName, 1, 1) = Chr$(k)
                
                Case Is > 20
                    For i = 1 To RandomNumber1 - 14
                            If mLebar = 1 Then
                                TempRnd = ((Rnd * 3141592621#) + 663896637#) / 28
                            Else
                                TempRnd = TempRnd / 65
                            End If
                            mLebar = mLebar + 1
                            If mLebar = 5 Then mLebar = 1
                        
                            Mid(VarName, i, 1) = Mid$(Huruf, (TempRnd Mod 65) + 1, 1)
                    
                    Next i
                    k = IIf(k > 122, 48, k + 1)
                    Mid(VarName, RandomNumber1 - 13, 1) = Chr$(k)
                    Mid(VarName, RandomNumber1 - 12) = " RANDOM CHAR"
                
                Case Else
            
                    For i = 1 To RandomNumber1 - 1
                        If mLebar = 1 Then
                            TempRnd = ((Rnd * 3141592621#) + 663896637#) / 28
                        Else
                            TempRnd = TempRnd / 65
                        End If
                        mLebar = mLebar + 1
                        If mLebar = 5 Then mLebar = 1
                        
                        Mid(VarName, i, 1) = Mid$(Huruf, (TempRnd Mod 65) + 1, 1)
                    
                    Next i
                    k = IIf(k > 122, 48, k + 1)
                    Mid(VarName, RandomNumber1, 1) = Chr$(k)
            End Select
            
        Case Else
            VarName = " "   ' a space
    End Select
End Sub


Private Sub AdjustGridTitle()
    Dim i As Long
    Dim GridFontWidth As Single     ' Average width
    Dim TotWidth As Long
    
    With Grid1
        .Visible = False
        .Clear           ' Clearing Grid
        DoEvents
        .Rows = 2
        .Cols = 2
        .FixedCols = 1
        .FixedRows = 1
        .Col = 0
        .Row = 0
        .Text = "Rec. #"
    End With
    
    With Grid1
        GridFontWidth = CalculateFontWidth("lW")
        .Cols = TableInfo(1) + 1
        .ColWidth(0) = 9 * GridFontWidth
        For i = 0 To TableInfo(1) - 1
            .TextMatrix(0, i + 1) = FieldInfo(i, 0)
            Select Case FieldInfo(i, 1)
                Case "C"
                    GridFontWidth = CalculateFontWidth("lW")
                    .ColWidth(i + 1) = CLng(FieldInfo(i, 2)) * GridFontWidth
                    .ColAlignment(i + 1) = flexAlignLeftCenter
                Case "D"
                    GridFontWidth = CalculateFontWidth("999")
                    .ColWidth(i + 1) = 10 * GridFontWidth
                    .ColAlignment(i + 1) = flexAlignCenterCenter
                Case "B", "Y", "N", "F"
                    GridFontWidth = CalculateFontWidth("999")
                    .ColWidth(i + 1) = 20 * GridFontWidth
                    .ColAlignment(i + 1) = flexAlignRightCenter
                Case "I"
                    GridFontWidth = CalculateFontWidth("999")
                    .ColWidth(i + 1) = 10 * GridFontWidth
                    .ColAlignment(i + 1) = flexAlignRightCenter
                Case "S"
                    GridFontWidth = CalculateFontWidth("999")
                    .ColWidth(i + 1) = 5 * GridFontWidth
                    .ColAlignment(i + 1) = flexAlignRightCenter
                Case "Z"
                    GridFontWidth = CalculateFontWidth("999")
                    .ColWidth(i + 1) = 3 * GridFontWidth
                    .ColAlignment(i + 1) = flexAlignRightCenter
                Case Else
                    GridFontWidth = CalculateFontWidth("lW")
                    .ColWidth(i + 1) = CLng(FieldInfo(i, 2)) * GridFontWidth
                    .ColAlignment(i + 1) = flexAlignLeftCenter
            End Select
            ' Keeps so that width not too small
            GridFontWidth = CalculateFontWidth("lW")
            If .ColWidth(i + 1) < GridFontWidth * 2 Then .ColWidth(i + 1) = GridFontWidth * 2
        Next i
    End With
    
    AdjustHscroll
    Grid1.Visible = True
    GridMode = 1
End Sub

Private Sub AdjustHscroll()
    Dim i As Long
    Dim TotWidth As Long
    ' Check total width...
    For i = 0 To Grid1.Cols - 1
        TotWidth = TotWidth + Grid1.ColWidth(i)
    Next i
    Grid1.Width = TotWidth + (Grid1.Cols * 30)
    If Grid1.Width > HScroll1.Width Then
        HScroll1.Max = (Grid1.Width - HScroll1.Width) / 150
        HScroll1.LargeChange = HScroll1.Max \ 4
    Else
        Grid1.Width = HScroll1.Width
        HScroll1.Max = 0
    End If
End Sub
'\'// Mousewheel for grid
Private Sub ProcessMessages()
    Dim Message As Msg

    Do While Not bCancel
        WaitMessage 'Wait For message and...

        If PeekMessage(Message, Grid1.hWnd, WM_MOUSEWHEEL, WM_MOUSEWHEEL, PM_REMOVE) Then  '...when the mousewheel is used...
            If Message.wParam < 0 Then '...scroll up...
                ' Repaint the grid based on ScrollPos value
                If ScrollPos < ScrollMax Then
                    RescaleScroll = 1
                    VScroll1.Value = VScroll1.Value + 1
                End If
                'If VScroll1.Value < VScroll1.Max Then VScroll1.Value = VScroll1.Value + 1
                Grid1.SetFocus
                'Grid1.TopRow = Grid1.TopRow + 1
            Else '... or scroll down
                If ScrollPos > ScrollMin Then
                    RescaleScroll = 1
                    VScroll1.Value = VScroll1.Value - 1
                End If
                'If VScroll1.Value > VScroll1.Min Then VScroll1.Value = VScroll1.Value - 1
                Grid1.SetFocus
'                If Grid1.TopRow > 1 Then
'                    Grid1.TopRow = Grid1.TopRow - 1
'                End If
            End If
            
        End If

        DoEvents
    Loop
End Sub

Private Sub HideGrid()
    PGrid.Visible = False
    VScroll1.Visible = False
    HScroll1.Visible = False
End Sub

Private Sub ShowGrid()
    PGrid.Visible = True
    VScroll1.Visible = True
    HScroll1.Visible = True
End Sub

Private Sub ViewInGrid(TopData As Long, lArray() As Long)
    On Error GoTo E_VIEWINGRID
    Dim naView() As Long
    Dim VStr() As String
    Dim rLeft As Long
    Dim i As Long, j As Long, k As Long
    
    Grid1.Visible = False
    Grid1.Rows = 1
    
    rLeft = UBound(lArray) - TopData + 1
    If rLeft < MaxView Then
        k = rLeft
        Grid1.Rows = k + 1
        For i = 0 To k - 1
            DBFTable.GetRow lArray(TopData + i), VStr
            Grid1.TextMatrix((i + 1), 0) = CStr(lArray(i + TopData))
            ' put to grid
            For j = 1 To TableInfo(1)
                Grid1.TextMatrix(i + 1, j) = VStr(j - 1)
            Next j
        Next i
    Else
        k = MaxView
        Grid1.Rows = k + 1
        For i = 0 To k - 1
            DBFTable.GetRow lArray(TopData + i), VStr
            Grid1.TextMatrix((i + 1), 0) = CStr(lArray(i + TopData))
            ' put to grid
            For j = 1 To TableInfo(1)
                Grid1.TextMatrix(i + 1, j) = VStr(j - 1)
            Next j
        Next i
    End If
    Grid1.Visible = True

E_VIEWINGRID:
End Sub

' Arrow clicked or the bar clicked
Private Sub VScroll1_Change()
    Dim lResult As Long, Hasil As Long
    If ScrollScale > 1 Then
        Hasil = VScroll1.Value - VScrollOldVal
        Select Case Hasil
            
            Case 1      ' Line Down
                If ScrollPos < ScrollMax Then ScrollPos = ScrollPos + 1
            
            Case -1     ' Line Up
                If ScrollPos > ScrollMin Then ScrollPos = ScrollPos - 1
            
            Case VScroll1.LargeChange   ' Page Down
                If ScrollPos + ScrollPageSize < ScrollMax Then ScrollPos = ScrollPos + ScrollPageSize
            
            Case Is = (0 - VScroll1.LargeChange)    ' Page up
                If ScrollPos - ScrollPageSize > ScrollMin Then ScrollPos = ScrollPos - ScrollPageSize
        
        End Select
        
        VScroll1.Value = (ScrollPos \ ScrollScale)
        
        If RescaleScroll <> 0 Then
            RescaleScroll = 0 - RescaleScroll
        End If
    
    Else
        ScrollPos = VScroll1.Value
    End If
    
    VScrollOldVal = VScroll1.Value

    ' Repaint the grid based on ScrollPos value
    If ScrollPos < 0 Then ScrollPos = 0
    If ScrollPos > ScrollMax Then ScrollPos = ScrollMax
    If TableReady Then
        If BrowseMode Then ViewInGrid ScrollPos, ArrGridView
    End If
    lblPos = "Rec. " & FormatNumber(ArrGridView(ScrollPos), 0) ' & " of " & FormatNumber(TableInfo(0), 0)
End Sub

' Scroll Thumb draged to new position
Private Sub VScroll1_Scroll()
    ' Convert current Vscroll position to 32 bit value.
    ScrollPos = VScroll1.Value * ScrollScale
    If ScrollPos < ScrollMin Then ScrollPos = ScrollMin
    If ScrollPos > ScrollMax Then ScrollPos = ScrollMax
    
    ' Save the current value of VScroll
    VScrollOldVal = VScroll1.Value
    lblPos = "Rec. " & FormatNumber(ArrGridView(ScrollPos), 0) ' & " of " & FormatNumber(TableInfo(0), 0)
    
    ' Repaint the grid based on ScrollPos value
    If BrowseMode Then ViewInGrid ScrollPos, ArrGridView
End Sub

' Tricking VScroll range to allow more than 32767
Private Sub ScrollRange(ItemCount As Long, PageSize As Long)
    Dim T As Long
    
    ScrollMin = 0
    ScrollMax = ItemCount - (PageSize - 1)
    ScrollPageSize = PageSize - 1
    If ScrollMax > 32766 Then
        ScrollScale = (ScrollMax \ 32766) + 1
        If PageSize \ ScrollScale < 2 Then
            VScroll1.LargeChange = 2
        Else
            VScroll1.LargeChange = PageSize \ ScrollScale
        End If
        RescaleScroll = 1
        VScroll1.Min = ScrollMin - 1
        VScroll1.Max = (ScrollMax \ ScrollScale) + 1
    Else
        RescaleScroll = 0
        ScrollScale = 1
        LargeScrollScale = 1
        VScroll1.SmallChange = 1
        VScroll1.Min = ScrollMin
        If ScrollMax > 1 Then
            VScroll1.Max = ScrollMax
            VScroll1.LargeChange = PageSize
        Else
            VScroll1.Max = VScroll1.Min
            VScroll1.LargeChange = 1
        End If
    End If

End Sub

Private Sub ResetVscroll()
    ScrollScale = 1
    ScrollMin = 0
    ScrollMax = 0
    ScrollPageSize = 0
    VScroll1.Min = 0
    VScroll1.Max = 0
End Sub
