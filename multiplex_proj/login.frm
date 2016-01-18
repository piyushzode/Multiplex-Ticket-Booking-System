VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   Caption         =   "Login"
   ClientHeight    =   10695
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   Picture         =   "login.frx":0000
   ScaleHeight     =   10695
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   4440
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select * from security"
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
   Begin VB.CommandButton signin 
      Height          =   375
      Left            =   7320
      MouseIcon       =   "login.frx":B261
      MousePointer    =   99  'Custom
      Picture         =   "login.frx":B56B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6720
      Width           =   1935
   End
   Begin VB.TextBox pass 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6000
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5640
      Width           =   3255
   End
   Begin VB.TextBox user 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Top             =   4800
      Width           =   3255
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
Adodc1.RecordSource = "select * from security"
End Sub


Private Sub signin_Click()
With Adodc1.Recordset
        .MoveFirst
        If user.Text = .Fields(0) And pass.Text = .Fields(1) Then
            Unload Me
            Load MDIForm1
            MDIForm1.Show
        Else
       MsgBox "Please enter a valid Data..!!", vbCritical + vbOKOnly, "Login Denied"
            user.SetFocus
            user.Text = ""
        End If
End With
End Sub
