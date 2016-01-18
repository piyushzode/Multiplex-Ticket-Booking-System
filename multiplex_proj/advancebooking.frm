VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmadvbooking 
   BackColor       =   &H00404040&
   Caption         =   "Advance booking"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15630
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10170
   ScaleWidth      =   15630
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MouseIcon       =   "advancebooking.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CheckBox chkC 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   42
      Top             =   6120
      Width           =   255
   End
   Begin VB.CheckBox chkD 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   41
      Top             =   6600
      Width           =   255
   End
   Begin VB.CheckBox chkE 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   40
      Top             =   7080
      Width           =   255
   End
   Begin VB.CheckBox chkF 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   39
      Top             =   7560
      Width           =   255
   End
   Begin VB.CheckBox chkG 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   38
      Top             =   8520
      Width           =   255
   End
   Begin VB.CheckBox chkA 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   37
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox chkB 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   36
      Top             =   5160
      Width           =   255
   End
   Begin VB.ComboBox cmbScreen 
      Height          =   315
      ItemData        =   "advancebooking.frx":030A
      Left            =   3720
      List            =   "advancebooking.frx":0320
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   960
      Width           =   1935
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      ItemData        =   "advancebooking.frx":0336
      Left            =   3720
      List            =   "advancebooking.frx":0343
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ComboBox cmbShow 
      Height          =   315
      ItemData        =   "advancebooking.frx":035F
      Left            =   3720
      List            =   "advancebooking.frx":0361
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SAVE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      MouseIcon       =   "advancebooking.frx":0363
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      MouseIcon       =   "advancebooking.frx":066D
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   10680
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2775
      Left            =   6360
      ScaleHeight     =   2715
      ScaleWidth      =   3435
      TabIndex        =   2
      Top             =   1080
      Width           =   3495
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MOVIE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label lblmoviename 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   975
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   -3360
      Top             =   5160
   End
   Begin VB.PictureBox Picture2 
      Height          =   1215
      Left            =   5880
      ScaleHeight     =   1155
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Label lblHouseFull 
         Alignment       =   2  'Center
         Caption         =   "HOUSE FULL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5640
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking for"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   6360
      TabIndex        =   35
      Top             =   720
      Width           =   1575
   End
   Begin VB.Image Image4 
      Height          =   990
      Left            =   2760
      Picture         =   "advancebooking.frx":0977
      Top             =   9360
      Width           =   10275
   End
   Begin VB.Label Label26 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "G(121-140)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   34
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label Label25 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "F(101-120)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   33
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label24 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "E(81-100)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   32
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label23 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "D(80-61)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   31
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label22 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "C(41-60)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   30
      Top             =   6120
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   1080
      Picture         =   "advancebooking.frx":20CC
      Top             =   8040
      Width           =   13470
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   1080
      Picture         =   "advancebooking.frx":2EAB
      Top             =   5640
      Width           =   13530
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Screen No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1680
      TabIndex        =   29
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Class"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1680
      TabIndex        =   28
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Time"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1680
      TabIndex        =   27
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Seats"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   1680
      TabIndex        =   26
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblseats 
      BackStyle       =   0  'Transparent
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   25
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblseatsavail1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seats Available"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   3600
      TabIndex        =   24
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblseatsavail 
      BackStyle       =   0  'Transparent
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   5280
      TabIndex        =   23
      Top             =   2760
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   3255
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "&Rate"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   10440
      TabIndex        =   22
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   10440
      TabIndex        =   21
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Entertainment Tax (10%)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   10440
      TabIndex        =   20
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Service Tax (4%)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   10440
      TabIndex        =   19
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblEtax 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   13200
      TabIndex        =   18
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblServicetax 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   13200
      TabIndex        =   17
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   13200
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblAmount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   13200
      TabIndex        =   15
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   10440
      TabIndex        =   14
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   13200
      TabIndex        =   13
      Top             =   2760
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   3135
      Left            =   10200
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label today 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Today"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      Top             =   720
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   1080
      Picture         =   "advancebooking.frx":3C5A
      Top             =   4200
      Width           =   13485
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "B(40-21)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "A(1-20)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   4680
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   13200
      X2              =   14280
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "frmadvbooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer
Dim SQL As String
Dim Srch As String
Dim Booking As Integer
Dim a, a1, c, d, e, b1, c1, d1 As String
Dim b, f, g, h, i, e1, f1, g1, h1, i1, silver, platinum, gold, col, gold1, silver1, platinum1 As Integer
Dim cn, cn1, cn2, cn3, cn4, cn5, cn6, CN7 As ADODB.Connection
Dim rs, rs1, rs2, rs3, rs4, rs5, rs6, RS7 As ADODB.Recordset
Dim z As Variant

Private Sub chkA_Click(Index As Integer)
    Srch = Combine("A", Index)
    If chkA(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Booking = Booking + 1
    End Sub

Private Sub chkB_Click(Index As Integer)
    Srch = Combine("B", Index)
    If chkB(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Booking = Booking + 1
  End Sub

Private Sub chkC_Click(Index As Integer)
    Srch = Combine("C", Index)
    If chkC(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Booking = Booking + 1
    End Sub

Private Sub chkD_Click(Index As Integer)
    Srch = Combine("D", Index)
    If chkD(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Booking = Booking + 1
    End Sub

Private Sub chkE_Click(Index As Integer)
    Srch = Combine("E", Index)
    If chkE(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Booking = Booking + 1
End Sub

Private Sub chkF_Click(Index As Integer)
    Srch = Combine("F", Index)
    If chkF(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
    Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Booking = Booking + 1
    End Sub

Private Sub chkG_Click(Index As Integer)
    Srch = Combine("G", Index)
    If chkG(Index).Value = 1 Then
        Seats = Seats - 1
        SeatNos(X) = Srch
        X = X + 1
     Else
        Seats = Seats + 1
        Call removeItem(Srch)
    End If
    Booking = Booking + 1
    End Sub



Private Sub cmbClass_Click()
Select Case cmbClass.ListIndex
        Case 0:
            Call SetSilver
            Call ClearGold
            Call ClearPlatinum
        Case 1:
            Call SetGold
            Call ClearSilver
            Call ClearPlatinum
        Case 2:
            Call SetPlatinum
            Call ClearGold
            Call ClearSilver
    End Select
    
Set cn = New Connection
Set rs = New Recordset
cn.Open "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
rs.Open "Select * from class", cn, adOpenDynamic, adLockOptimistic

    Seats = 140
    lblseatsavail.Caption = 140
    lblAmount.Caption = ""
    lblServicetax.Caption = ""
    lblEtax.Caption = ""
    lbltotal.Caption = ""
    
    With rs
        .MoveFirst
        While .EOF <> True
            If rs(0) = cmbClass.Text Then
                lblRate.Caption = rs(1)
                Exit Sub
            Else
                .MoveNext
            End If
        Wend
    End With
    rs.Close
    cn.Close
    cmbShow.Clear

End Sub



Private Sub cmbShow_Click()
    
    Dim a As String
    Set cn1 = New Connection
    Set rs1 = New Recordset
    cn1.Open "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
    rs1.Open "Select * from booking", cn1, adOpenDynamic, adLockOptimistic
    cmbScreen.Enabled = False
    cmbClass.Enabled = False
    cmdsave.Enabled = True
    
    
    Booking = 0
    With rs1
             .MoveFirst
        While .EOF <> True
            If rs1(1) = cmbScreen.ListIndex + 1 Then
                If Format(rs1(2), "dd/MMM/yyyy") = Format(AdvBookDate, "dd/MMM/yyyy") Then
                    If rs1(3) = cmbShow.Text Then
                       Call MarkReserved(rs1.Fields(0))
                    End If
                End If
                .MoveNext
            Else
              .MoveNext
          End If
      Wend
    End With
    
    
    Booking = 0
    rs1.Close
    cn1.Close
    Call housefull
    
End Sub

Public Sub MarkReserved(SNo As String)
    
    Dim L As Integer
    Dim Series As String
    Dim Index As Integer
    Dim i As Integer
    
    i = 1
    L = Len(SNo)
    While i <= L
        Series = Mid(SNo, i, 4)
        i = i + 4
        Index = Val(Mid(Series, 2, 3))
        Series = Mid(Series, 1, 1)
        Select Case Series
            Case "A":
                chkA(Index).Enabled = False
                chkA(Index).Value = 1
                platinum1 = platinum1 + 1
            Case "B":
                chkB(Index).Enabled = False
                chkB(Index).Value = 1
                platinum1 = platinum1 + 1
            Case "C":
                chkC(Index).Enabled = False
                chkC(Index).Value = 1
                gold1 = gold1 + 1
            Case "D":
                chkD(Index).Enabled = False
                chkD(Index).Value = 1
                gold1 = gold1 + 1
            Case "E":
                chkE(Index).Enabled = False
                chkE(Index).Value = 1
                gold1 = gold1 + 1
            Case "F":
                chkF(Index).Enabled = False
                chkF(Index).Value = 1
                gold1 = gold1 + 1
            Case "G":
                chkG(Index).Enabled = False
                chkG(Index).Value = 1
                silver1 = silver1 + 1
            End Select
        'Booking = Booking + 1
    Wend
    lblseatsavail.Caption = Seats
    lblAmount.Caption = ""
    lbltotal.Caption = ""
    lblServicetax.Caption = ""
    lblEtax.Caption = ""
    
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    Dim i, col As Integer
    Dim s As String
    Dim tot As Integer
    

    tot = 0
    platinum = 0
    gold = 0
    silver = 0
                    
    Set cn6 = New Connection
    Set rs6 = New Recordset
    cn6.Open "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
    rs6.Open "Select * from ticket", cn6, adOpenDynamic, adLockOptimistic
    cn6.Execute "truncate table ticket"
    cn6.Close
    rs6.Close
    
            
            
            
    For i = 0 To 19
    If chkA(i).Enabled = True And chkA(i).Value = 1 Then
      platinum = platinum + 1
    End If
    
    If chkB(i).Enabled = True And chkB(i).Value = 1 Then
     platinum = platinum + 1
    End If
    
    If chkC(i).Enabled = True And chkC(i).Value = 1 Then
     gold = gold + 1
    End If
        
    If chkD(i).Enabled = True And chkD(i).Value = 1 Then
     gold = gold + 1
    End If
    
    If chkE(i).Enabled = True And chkE(i).Value = 1 Then
     gold = gold + 1
    End If
    
    If chkF(i).Enabled = True And chkF(i).Value = 1 Then
     gold = gold + 1
    End If
    
    If chkG(i).Enabled = True And chkG(i).Value = 1 Then
     silver = silver + 1
    End If
    
    Next i
    
    
    If platinum = 0 And silver = 0 And gold = 0 Then
    MsgBox "Please select atleast one seat"
    Else
    
    Call CalAmt
    Set cn1 = New Connection
    Set rs1 = New Recordset
    cn1.Open "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
    rs1.Open "Select * from booking", cn1, adOpenDynamic, adLockOptimistic
        
    For i = 0 To X - 1
    s = s & SeatNos(i)
    Next
    
   
         a = s
         b = cmbScreen.ListIndex + 1
         c = Format(AdvBookDate, "dd/MMM/yyyy")
         d = cmbShow.Text
         f = lblAmount.Caption
         g = lblEtax.Caption
         h = lblServicetax.Caption
         i = lbltotal.Caption

    cn1.Execute "insert into booking values ('" & a & "'," & b & ",'" & c & "','" & d & "','A'," & Val(f) & "," & Val(g) & "," & Val(h) & "," & Val(i) & ")"
    rs1.Close
    cn1.Close
        
    Set CN7 = New Connection
    Set RS7 = New Recordset
    CN7.Open "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
    RS7.Open "select * from movie where mname='" & lblmoviename.Caption & "'", CN7, adOpenDynamic, adLockOptimistic
        
    col = RS7(4) + Val(lbltotal.Caption)
    CN7.Execute "UPDATE  movie SET  collection=" & col & "  WHERE mname='" & lblmoviename.Caption & "'"
    RS7.Close
    CN7.Close
    
    i = 0
    Set cn3 = New Connection
    Set rs3 = New Connection
    cn3.Open "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
    rs3.Open "select * from ticket"
    
        For i = X - Booking + 1 To X - 1
    
          With rs3
          .MoveNext
            a1 = SeatNos(i)
            b1 = lblmoviename.Caption
            c1 = Format(AdvBookDate, "dd/MMM/yyyy")
            d1 = cmbShow.Text
            e1 = lblRate.Caption
            f1 = Val(lblRate.Caption) * 10 / 100
            g1 = Val(lblRate.Caption) * 4 / 100
            h1 = Val(lblRate.Caption) + Val(lblRate.Caption) * 4 / 100 + Val(lblRate.Caption) * 10 / 100
            i1 = cmbScreen.Text
            cn3.Execute " insert into ticket values ('" & a1 & "','" & b1 & "','" & c1 & "','" & d1 & "'," & Val(e1) & "," & Val(f1) & "," & Val(g1) & "," & Val(h1) & "," & i1 & ")"
         
        End With
    Next i
    
    
    MsgBox "Your seats have been booked..!!! "
    
       
    If isLoad = True Then
        Unload DataReport4
       Unload DataEnvironment1
    Else
        isLoad = True
    End If
    Load DataReport4
    DataReport4.Show
    
    
    rs3.Close
    cn3.Close
    
            
    Set cn5 = New Connection
    Set rs5 = New Recordset
    cn5.Open "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
    rs5.Open "select * from collection", cn5, adOpenDynamic, adLockOptimistic
    
    z = Format(AdvBookDate, "dd/MMM/yyyy")
       
    With rs5
    .MoveFirst
    
    While .EOF <> True
        If Format(rs5(0), "dd/MMM/yyyy") = Format(AdvBookDate, "dd/MMM/yyyy") Then
        tot = rs5(cmbScreen.ListIndex + 1) + Val(lbltotal.Caption)
        
       If cmbScreen.ListIndex = 0 Then
        cn5.Execute "update collection set screen1=" & tot & "where showdate='" & z & "'"
        
       ElseIf cmbScreen.ListIndex = 1 Then
       cn5.Execute "update collection set screen2=" & tot & "where showdate='" & z & "'"
    
       ElseIf cmbScreen.ListIndex = 2 Then
       cn5.Execute "update collection set screen3=" & tot & "where showdate='" & z & "'"
       
       ElseIf cmbScreen.ListIndex = 3 Then
       cn5.Execute "update collection set screen4=" & tot & "where showdate='" & z & "'"
    
       ElseIf cmbScreen.ListIndex = 4 Then
       cn5.Execute "update collection set screen5=" & tot & "where showdate='" & z & "'"
       
       ElseIf cmbScreen.ListIndex = 5 Then
       cn5.Execute "update collection set screen6=" & tot & "where showdate='" & z & "'"
    
    End If
    
    tot = 0
    tot = rs5(1) + rs5(2) + rs5(3) + rs5(4) + rs5(5) + rs5(6) + Val(lbltotal.Caption)
    
    
    cn5.Execute "update collection set total=" & tot & "where showdate='" & z & "'"
    
    .MoveNext
    Unload Me
    Exit Sub
    
    Else
    .MoveNext
   End If
   Wend
   
   cn5.Execute "insert into collection values ('" & Format(AdvBookDate, "dd/MMM/yyyy") & "',0,0,0,0,0,0,0)"
    
   tot = 0
   tot = Val(lbltotal.Caption)
       
   If cmbScreen.ListIndex = 0 Then
   cn5.Execute "update collection set screen1=" & tot & "where showdate='" & z & "'"
        
   ElseIf cmbScreen.ListIndex = 1 Then
   cn5.Execute "update collection set screen2=" & tot & "where showdate='" & z & "'"
    
   ElseIf cmbScreen.ListIndex = 2 Then
   cn5.Execute "update collection set screen3=" & tot & "where showdate='" & z & "'"
       
   ElseIf cmbScreen.ListIndex = 3 Then
   cn5.Execute "update collection set screen4=" & tot & "where showdate='" & z & "'"
    
   ElseIf cmbScreen.ListIndex = 4 Then
   cn5.Execute "update collection set screen5=" & tot & "where showdate='" & z & "'"
     
   ElseIf cmbScreen.ListIndex = 5 Then
   cn5.Execute "update collection set screen6=" & tot & "where showdate='" & z & "'"
    
   End If
    
   tot = rs5(1) + rs5(2) + rs5(3) + rs5(4) + rs5(5) + rs5(6) + Val(lbltotal.Caption)
   cn5.Execute "update collection set total=" & tot & "where showdate='" & z & "'"
    
   End With
   cn5.Close
   rs5.Close
   Unload Me
      End If
      
End Sub

Private Sub Command1_Click()
frmcurrentbooking.Refresh
Unload Me
Load Me

End Sub

Private Sub Form_Load()
Call DispSeats
Call ClearGold
Call ClearPlatinum
Call ClearSilver
        
End Sub
Private Sub Timer1_Timer()
today.Caption = AdvBookDate
End Sub


Public Sub DispSeats()
    Dim i As Integer
    For i = 1 To 19
        Load chkA(i)
        chkA(i).Left = chkA(i - 1).Left + chkA(i - 1).Width + 320
        chkA(i).Visible = True
        
        Load chkB(i)
        chkB(i).Left = chkB(i - 1).Left + chkB(i - 1).Width + 320
        chkB(i).Visible = True

        Load chkC(i)
        chkC(i).Left = chkC(i - 1).Left + chkC(i - 1).Width + 320
        chkC(i).Visible = True
        
        Load chkD(i)
        chkD(i).Left = chkD(i - 1).Left + chkD(i - 1).Width + 320
        chkD(i).Visible = True
        
        Load chkE(i)
        chkE(i).Left = chkE(i - 1).Left + chkE(i - 1).Width + 320
        chkE(i).Visible = True
        
        Load chkF(i)
        chkF(i).Left = chkF(i - 1).Left + chkF(i - 1).Width + 320
        chkF(i).Visible = True
        
        Load chkG(i)
        chkG(i).Left = chkG(i - 1).Left + chkG(i - 1).Width + 320
        chkG(i).Visible = True
       Next i
    End Sub
    
    Public Sub SetPlatinum()
    Dim i As Integer
    For i = 0 To 19
    If chkA(i).Value = 0 Then
    chkA(i).Enabled = True
    End If
    If chkB(i).Value = 0 Then
    chkB(i).Enabled = True
    End If
    
    Next i
    End Sub
    
    Public Sub SetGold()
    Dim i As Integer
    For i = 0 To 19
    If chkC(i).Value = 0 Then
    chkC(i).Enabled = True
    End If
    If chkD(i).Value = 0 Then
    chkD(i).Enabled = True
    End If
    If chkE(i).Value = 0 Then
    chkE(i).Enabled = True
    End If
    If chkF(i).Value = 0 Then
    chkF(i).Enabled = True
    End If
    Next i
    End Sub
    
    Public Sub SetSilver()
    Dim i As Integer
    For i = 0 To 19
    If chkG(i).Value = 0 Then
        chkG(i).Enabled = True
    End If
    
    Next i
    End Sub
    
    Public Sub ClearPlatinum()
    Dim i As Integer
    For i = 0 To 19
        chkA(i).Value = 0
        chkA(i).Enabled = False
        chkB(i).Value = 0
        chkB(i).Enabled = False
    Next i
    End Sub

    Public Sub ClearGold()
    Dim i As Integer
    For i = 0 To 19
        chkC(i).Value = 0
        chkC(i).Enabled = False
        chkD(i).Value = 0
        chkD(i).Enabled = False
        chkE(i).Value = 0
        chkE(i).Enabled = False
        chkF(i).Value = 0
        chkF(i).Enabled = False
        Next i
    End Sub

    Public Sub ClearSilver()
    Dim i As Integer
    For i = 0 To 19
        chkG(i).Value = 0
        chkG(i).Enabled = False
      Next i
    End Sub

    Public Sub ClearAllSeats()
    Dim i As Integer
    For i = 0 To 19
        chkA(i).Value = 0
        chkB(i).Value = 0
        chkC(i).Value = 0
        chkD(i).Value = 0
        chkE(i).Value = 0
        chkF(i).Value = 0
        chkG(i).Value = 0
        chkA(i).Enabled = True
        chkB(i).Enabled = True
        chkC(i).Enabled = True
        chkD(i).Enabled = True
        chkE(i).Enabled = True
        chkF(i).Enabled = True
        chkG(i).Enabled = True
        
    Next i
    End Sub


Public Function Combine(s As String, Index As Integer) As String
    Dim comb As String
    Dim ind As String
    If Index < 10 Then
        ind = "00" & Index
    ElseIf Index < 100 Then
        ind = "0" & Index
    Else
        ind = Index
    End If
    comb = s & ind
    Combine = comb
End Function



Public Sub removeItem(src As String)
    Dim i As Integer
    Dim j As Integer
    For i = 0 To X - 1
        If SeatNos(i) = src Then
            For j = i + 1 To X - 1
                SeatNos(j - 1) = SeatNos(j)
            Next j
            SeatNos(j) = ""
            X = X - 1
        End If
    Next i
End Sub



Public Sub CalAmt()
    Booking = Booking + 1
    lblseatsavail.Caption = Seats
    
    If cmbClass.ListIndex = 0 Then
    lblAmount.Caption = silver * lblRate.Caption
    End If
    If cmbClass.ListIndex = 1 Then
    lblAmount.Caption = gold * lblRate.Caption
    End If
    If cmbClass.ListIndex = 2 Then
    lblAmount.Caption = platinum * lblRate.Caption
    End If
        
    lblEtax.Caption = Val(lblAmount.Caption) * 10 / 100
    lblServicetax.Caption = Val(lblAmount.Caption) * 4 / 100
    lbltotal.Caption = Val(lblAmount.Caption) + Val(lblEtax.Caption) + Val(lblServicetax.Caption)
  
    
End Sub



Private Sub cmbScreen_Click()

    Dim i As Integer
    Set cn2 = New Connection
    Set rs2 = New Recordset
    cn2.Open "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
    rs2.Open "Select * from movie where screen_no=" & cmbScreen.ListIndex + 1 & " order by indate asc", cn2, adOpenDynamic, adLockOptimistic
    rs2.MoveLast
    lblmoviename.Caption = rs2(2)
    rs2.Close
    cn2.Close
    cmbShow.Clear
    
    Set cn3 = New Connection
    Set rs3 = New Recordset
    cn3.Open "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
    rs3.Open "Select * from theater where screen_no=" & cmbScreen.ListIndex + 1, cn3, adOpenDynamic, adLockOptimistic
    For i = 2 To 6
        If rs3(i) <> "12:00:00 AM" Then cmbShow.AddItem rs3(i)
    Next i
    rs3.Close
    cn3.Close
End Sub

Public Sub housefull()
If lblseatsavail.Caption = 0 Then
        Picture2.Visible = True
        cmdsave.Enabled = False
        cmbShow.Enabled = False
        cmbScreen.Enabled = False
        cmbClass.Enabled = False
    Else
        Picture2.Visible = False
    End If
End Sub


Private Sub Timer3_Timer()
lblHouseFull.Visible = Not lblHouseFull.Visible
End Sub

