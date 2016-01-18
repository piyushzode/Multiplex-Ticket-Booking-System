VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmcurrentbooking 
   BackColor       =   &H00404040&
   Caption         =   "Current Booking"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4920
      Top             =   0
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
      Connect         =   "Provider=MSDAORA.1;Password=suman;User ID=system;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=suman;User ID=system;Persist Security Info=True"
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
   Begin VB.CheckBox chkB 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   33
      Top             =   4920
      Width           =   255
   End
   Begin VB.CheckBox chkA 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   16
      Top             =   4440
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   2775
      Left            =   5640
      ScaleHeight     =   2715
      ScaleWidth      =   3435
      TabIndex        =   36
      Top             =   840
      Width           =   3495
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
         Height          =   615
         Left            =   240
         TabIndex        =   38
         Top             =   1440
         Width           =   2895
      End
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
         TabIndex        =   37
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CheckBox chkG 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   35
      Top             =   8280
      Width           =   255
   End
   Begin VB.CheckBox chkF 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   34
      Top             =   7320
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9960
      Top             =   0
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
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SAVE"
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CheckBox chkE 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   19
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox chkD 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   18
      Top             =   6360
      Width           =   255
   End
   Begin VB.CheckBox chkC 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   17
      Top             =   5880
      Width           =   255
   End
   Begin VB.ComboBox cmbShow 
      Height          =   315
      ItemData        =   "ADVBOOKING.frx":0000
      Left            =   3000
      List            =   "ADVBOOKING.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      ItemData        =   "ADVBOOKING.frx":0017
      Left            =   3000
      List            =   "ADVBOOKING.frx":0024
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ComboBox cmbScreen 
      Height          =   315
      ItemData        =   "ADVBOOKING.frx":0040
      Left            =   3000
      List            =   "ADVBOOKING.frx":0056
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
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
      Left            =   840
      TabIndex        =   39
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label2 
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
      Left            =   840
      TabIndex        =   5
      Top             =   4920
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   360
      Picture         =   "ADVBOOKING.frx":006C
      Top             =   3960
      Width           =   13485
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
      Left            =   5520
      TabIndex        =   32
      Top             =   480
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   3135
      Left            =   9480
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   4335
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
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   12480
      TabIndex        =   29
      Top             =   2520
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
      Left            =   9720
      TabIndex        =   28
      Top             =   2520
      Width           =   2295
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
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   12480
      TabIndex        =   27
      Top             =   1080
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
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   12480
      TabIndex        =   26
      Top             =   600
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
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   12480
      TabIndex        =   25
      Top             =   2040
      Width           =   975
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
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   12480
      TabIndex        =   24
      Top             =   1560
      Width           =   975
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
      Left            =   9720
      TabIndex        =   23
      Top             =   2040
      Width           =   2295
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
      Left            =   9720
      TabIndex        =   22
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label11 
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
      Left            =   9720
      TabIndex        =   21
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label10 
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
      Left            =   9720
      TabIndex        =   20
      Top             =   600
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   3255
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   4455
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
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   4560
      TabIndex        =   15
      Top             =   2880
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
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   2880
      Width           =   1575
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
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblseats1 
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
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label9 
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
      Left            =   960
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label8 
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
      Left            =   960
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblscreenno 
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
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   360
      Picture         =   "ADVBOOKING.frx":098E
      Top             =   5400
      Width           =   13530
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   360
      Picture         =   "ADVBOOKING.frx":173D
      Top             =   7800
      Width           =   13470
   End
   Begin VB.Label Label3 
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
      Left            =   720
      TabIndex        =   4
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label4 
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
      Left            =   720
      TabIndex        =   3
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label5 
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
      Left            =   720
      TabIndex        =   2
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label6 
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
      Left            =   720
      TabIndex        =   1
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label7 
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
      Left            =   720
      TabIndex        =   0
      Top             =   8280
      Width           =   975
   End
   Begin VB.Image Image4 
      Height          =   990
      Left            =   2040
      Picture         =   "ADVBOOKING.frx":251C
      Top             =   9120
      Width           =   10275
   End
End
Attribute VB_Name = "frmcurrentbooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer
Dim SQL As String
Dim Srch As String
Dim Booking As Integer
Dim a, c, d, e As String
Dim b, f, g, h, i As Integer
Dim cn, CN1, cn2, cn3 As ADODB.Connection
Dim rs, RS1, rs2, rs3 As ADODB.Recordset

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
    Call CalAmt
    
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
    Call CalAmt
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
    Call CalAmt
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
    Call CalAmt
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
    Call CalAmt
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
    Call CalAmt
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
    Call CalAmt
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
    
    
    
'Adodc1.ConnectionString = "Provider=MSDAORA.1;Password=suman;User ID=system;Persist Security Info=True"
'Adodc1.RecordSource = "Select * from class"
Set cn = New Connection
Set rs = New Recordset
cn.Open "Provider=MSDAORA.1;Password=suman;User ID=system;Persist Security Info=True"
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
End Sub



Private Sub cmbShow_Click()
    'Call ClearAllSeats
    
    Set CN1 = New Connection
    Set RS1 = New Recordset
    CN1.Open "Provider=MSDAORA.1;Password=suman;User ID=system;Persist Security Info=True"
    RS1.Open "Select * from booking", CN1, adOpenDynamic, adLockOptimistic
         
         Label16.Caption = RS1(0)
        ' Call MarkReserved(RS1(0))
                 With RS1
             .MoveFirst
        While .EOF <> True
       
            If RS1(1) = cmbScreen.ListIndex + 1 Then
               ' If Format(RS1(2), "DD/MM/YYYY") = Format(Date, "DD/MM/YYYY") Then
                 '   If RS1(3) = cmbShow.Text Then
                       'Call MarkReserved(RS1(0))
               '     End If
             '   End If
                .MoveNext
            Else
              .MoveNext
          End If
      Wend
    End With
    Booking = 0
    RS1.Close
    CN1.Close
End Sub

Public Sub MarkReserved(SNo As String)
    Dim l As Integer
    Dim Series As String
    Dim Index As Integer
    Dim i As Integer
    i = 1
    l = Len(SNo)
    While i <= l
        Series = Mid(SNo, i, 4)
        i = i + 4
        lblseatsavail.Caption = Seats
        Index = Val(Mid(Series, 2, 3))
        Series = Mid(Series, 1, 1)
        Select Case Series
            Case "A":
                chkA(Index).Enabled = False
                chkA(Index).Value = 1
            Case "B":
                chkB(Index).Enabled = False
                chkB(Index).Value = 1
            Case "C":
                chkC(Index).Enabled = False
                chkC(Index).Value = 1
            Case "D":
                chkD(Index).Enabled = False
                chkD(Index).Value = 1
            Case "E":
                chkE(Index).Enabled = False
                chkE(Index).Value = 1
            Case "F":
                chkF(Index).Enabled = False
                chkF(Index).Value = 1
            Case "G":
                chkG(Index).Enabled = False
                chkG(Index).Value = 1
              End Select
        Booking = Booking + 1
    Wend
    
    lblAmount.Caption = ""
    lbltotal.Caption = ""
    lblServicetax.Caption = ""
    lblEtax.Caption = ""
End Sub




Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    Dim i As Integer
    Dim s As String
    
    Set CN1 = New Connection
    Set RS1 = New Recordset
    CN1.Open "Provider=MSDAORA.1;Password=suman;User ID=system;Persist Security Info=True"
    RS1.Open "Select * from booking", CN1, adOpenDynamic, adLockOptimistic
    Label16.Caption = RS1(2)
    For i = 0 To X - 1
        s = s & SeatNos(i)
    Next
    
    'CN1.Execute "insert into booking values ( 'scasc', " & cmbScreen.ListIndex + 1 & ",' " & Format(Date, "DD/MM/YYYY") & " ','" & cmbShow.Text & "','C'," & lblAmount.Caption & "," & lblEtax.Caption & "," & lblServicetax.Caption & "," & lbltotal.Caption & ")"
    
    'With RS1
       '.AddNew
         a = s
         b = cmbScreen.ListIndex + 1
         c = Date
         d = cmbShow.Text
         e = "C"  'Current Booking
         f = lblAmount.Caption
         g = lblEtax.Caption
         h = lblServicetax.Caption
         i = lbltotal.Caption
        '.Update
        'a = Text1.Text
         Text1.Text = f
      
        CN1.Execute "insert into booking values (' " & a & " '," & b & ",'14-oct-2010',' " & d & " ','C',1,1,1,1)"
        
       'CN1.Execute "insert into booking(seat_no) values (' imba ')"
       
       ', " & b & " ,' " & c & " ',' " & d & " ',' " & e & " ', " & f & " , " & g & " , " & h & " , " & i & "
       ',theater_no,show_date,show_time,booking_type,amount,enter_tax,service_tax,total
   'End With
    RS1.Close
    CN1.Close
    
End Sub

Private Sub Form_Load()
Call DispSeats
End Sub

Private Sub Timer1_Timer()
today.Caption = Now()
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
    chkA(i).Enabled = True
    chkB(i).Enabled = True
    Next i
    End Sub
    
    Public Sub SetGold()
    Dim i As Integer
    For i = 0 To 19
    chkC(i).Enabled = True
    chkD(i).Enabled = True
    chkE(i).Enabled = True
    chkF(i).Enabled = True
    Next i
    End Sub
    
    Public Sub SetSilver()
    Dim i As Integer
    For i = 0 To 19
    chkG(i).Enabled = True
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
    lblAmount.Caption = (140 - Seats) * Val(lblRate.Caption)
    lblEtax.Caption = Val(lblAmount.Caption) * 10 / 100
    lblServicetax.Caption = Val(lblAmount.Caption) * 4 / 100
    lbltotal.Caption = Val(lblAmount.Caption) + Val(lblEtax.Caption) + Val(lblServicetax.Caption)
End Sub



Private Sub cmbScreen_Click()

    Dim i As Integer
    Set cn2 = New Connection
    Set rs2 = New Recordset
    cn2.Open "Provider=MSDAORA.1;Password=suman;User ID=system;Persist Security Info=True"
    rs2.Open "Select * from movie where screen_no=" & cmbScreen.ListIndex + 1 & " order by indate asc", cn2, adOpenDynamic, adLockOptimistic
    rs2.MoveLast
    lblmoviename.Caption = rs2(2)
    rs2.Close
    cn2.Close
    cmbShow.Clear
    
    Set cn3 = New Connection
    Set rs3 = New Recordset
    cn3.Open "Provider=MSDAORA.1;Password=suman;User ID=system;Persist Security Info=True"
    rs3.Open "Select * from theater where screen_no=" & cmbScreen.ListIndex + 1, cn3, adOpenDynamic, adLockOptimistic
    For i = 2 To 6
        If rs3(i) <> "12:00:00AM" Then cmbShow.AddItem rs3(i)
    Next i
    rs3.Close
    cn3.Close
End Sub
