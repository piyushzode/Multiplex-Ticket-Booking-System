VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNewMovie 
   Caption         =   "New Movie"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16845
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "frmnewmovie.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   16845
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9120
      Top             =   10080
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "select * from movie"
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
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   15720
      Top             =   3840
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11400
      MouseIcon       =   "frmnewmovie.frx":9D92
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   9000
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      MouseIcon       =   "frmnewmovie.frx":A09C
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   9000
      Width           =   2295
   End
   Begin VB.ComboBox cmbAMPM 
      Height          =   315
      Index           =   4
      ItemData        =   "frmnewmovie.frx":A3A6
      Left            =   14160
      List            =   "frmnewmovie.frx":A3B0
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbAMPM 
      Height          =   315
      Index           =   3
      ItemData        =   "frmnewmovie.frx":A3BC
      Left            =   14160
      List            =   "frmnewmovie.frx":A3C6
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbAMPM 
      Height          =   315
      Index           =   2
      ItemData        =   "frmnewmovie.frx":A3D2
      Left            =   14160
      List            =   "frmnewmovie.frx":A3DC
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbAMPM 
      Height          =   315
      Index           =   1
      ItemData        =   "frmnewmovie.frx":A3E8
      Left            =   14160
      List            =   "frmnewmovie.frx":A3F2
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox cmbAMPM 
      Height          =   315
      Index           =   0
      ItemData        =   "frmnewmovie.frx":A3FE
      Left            =   14160
      List            =   "frmnewmovie.frx":A408
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4920
      Width           =   735
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   13200
      TabIndex        =   12
      Top             =   4920
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox cmbShows 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmnewmovie.frx":A414
      Left            =   8640
      List            =   "frmnewmovie.frx":A427
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   6360
      Width           =   1815
   End
   Begin VB.ComboBox cmbScreen 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmnewmovie.frx":A43F
      Left            =   8640
      List            =   "frmnewmovie.frx":A455
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox txtMovieName 
      Alignment       =   2  'Center
      DataSource      =   "Adodc1"
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
      Left            =   9480
      TabIndex        =   1
      Top             =   2880
      Width           =   5655
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   13200
      TabIndex        =   13
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   13200
      TabIndex        =   14
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   13200
      TabIndex        =   15
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   13200
      TabIndex        =   16
      Top             =   7320
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "hh:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox Text1 
      Height          =   765
      Index           =   0
      Left            =   4320
      TabIndex        =   26
      Text            =   "12:00:00 AM"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   1
      Left            =   4320
      TabIndex        =   27
      Text            =   "12:00:00 AM"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   2
      Left            =   4320
      TabIndex        =   28
      Text            =   "12:00:00 AM"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   3
      Left            =   4320
      TabIndex        =   29
      Text            =   "12:00:00 AM"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   4
      Left            =   4320
      TabIndex        =   30
      Text            =   "12:00:00 AM"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbltoday 
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
      Height          =   495
      Left            =   12720
      TabIndex        =   25
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C0C0C0&
      Height          =   9255
      Left            =   5760
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   10695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   1095
      Left            =   6480
      Shape           =   4  'Rounded Rectangle
      Top             =   8760
      Width           =   9255
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      Height          =   3735
      Left            =   6480
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   4455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Timings"
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
      Height          =   495
      Left            =   12840
      TabIndex        =   22
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      Height          =   3735
      Left            =   11280
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   4455
   End
   Begin VB.Label lblshow 
      BackStyle       =   0  'Transparent
      Caption         =   "Show 5"
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
      Height          =   495
      Index           =   4
      Left            =   12000
      TabIndex        =   11
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblshow 
      BackStyle       =   0  'Transparent
      Caption         =   "Show 4"
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
      Height          =   495
      Index           =   3
      Left            =   12000
      TabIndex        =   10
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblshow 
      BackStyle       =   0  'Transparent
      Caption         =   "Show 3"
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
      Height          =   495
      Index           =   2
      Left            =   12000
      TabIndex        =   9
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblshow 
      BackStyle       =   0  'Transparent
      Caption         =   "Show 2"
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
      Height          =   495
      Index           =   1
      Left            =   12000
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblshow 
      BackStyle       =   0  'Transparent
      Caption         =   "Show &1"
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
      Height          =   495
      Index           =   0
      Left            =   12000
      TabIndex        =   7
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Screen  && Shows"
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
      Height          =   495
      Left            =   7920
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Of Shows"
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
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "S&elect Screen"
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
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   975
      Left            =   6480
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   9135
   End
   Begin VB.Label lblMovieName 
      BackStyle       =   0  'Transparent
      Caption         =   "&Name Of the Movie"
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
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   3000
      Width           =   2055
   End
End
Attribute VB_Name = "frmNewMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn, cn1 As ADODB.Connection
Dim rs, rs1 As ADODB.Recordset
Dim a, b, c As String
Dim d, e As Integer

Private Sub cmbShows_Click()
Dim i As Integer
Dim j As Integer
  cmdsave.Enabled = True
  For i = 1 To cmbShows.Text - 1
  lblshow(i).Visible = True
  MaskEdBox1(i).Visible = True
  cmbAMPM(i).Visible = True
  Next
  
  For j = i To 4
        lblshow(j).Visible = False
        MaskEdBox1(j).Visible = False
        cmbAMPM(j).Visible = False
    Next
    
 End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim showTime As String
    Dim i As Integer
    Dim j As Integer
    Dim flag As Integer
    On Error Resume Next
    
    flag = 0
    For i = 0 To Val(cmbShows.Text) - 1
    If Len(MaskEdBox1(i).ClipText) < 4 Or Len(cmbAMPM(i).Text) = 0 Then
        MsgBox "Please enter show timings"
        MaskEdBox1(i).SetFocus
        flag = 1
    
    End If
    Next i
    
    For i = 0 To Val(cmbShows.Text) - 1
    If Mid(MaskEdBox1(i).ClipText, 1, 2) > 24 Then
    MsgBox "Enter a valid Time"
    MaskEdBox1(i).SetFocus
    flag = 1
    End If
    Next i
    
    If flag = 0 Then
    With rs1
    .MoveLast
        While .BOF <> True
            If rs1(3) = cmbScreen.ListIndex + 1 Then
                rs1(1) = Format(Date, "dd/MMM/yyyy")
                cn1.Execute "UPDATE  movie SET  MNAME='" & txtMovieName.Text & "'  WHERE Screen_NO='" & Val(cmbScreen.Text) & "'"
                GoTo a:
            Else
                .MovePrevious
            End If
        Wend
    End With

a:
        a = Format(Date, "dd/MMM/yyyy")
        b = Format(Date, "dd/MMM/yyyy")
        c = txtMovieName.Text
        d = cmbScreen.ListIndex + 1
        e = 0
     
    cn1.Execute "insert into movie values ('" & a & "','" & b & "','" & c & "'," & Val(d) & "," & e & ")"
    rs1.Close
    cn1.Close
   
    Set cn = New Connection
    Set rs = New Recordset
    cn.Open "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
    rs.Open "select * from theater", cn, adOpenDynamic, adLockOptimistic
    
    
    With rs
        .MoveFirst
        While .EOF <> True
            If rs(0) = cmbScreen.ListIndex + 1 Then
               For i = 0 To cmbShows.ListIndex
                 If cmbAMPM(i).ListIndex = 1 Then
                    If Mid(MaskEdBox1(i).ClipText, 1, 2) <> 12 Then
                         showTime = (12 + Mid(MaskEdBox1(i).ClipText, 1, 2))
                    Else
                        showTime = "12"
                    End If
                    showTime = showTime & ":" & Mid(MaskEdBox1(i).ClipText, 3, 2)
                 Else
                    showTime = MaskEdBox1(i).Text
                 End If
                 'rs.Fields(i + 2) = FormatDateTime(showTime, vbLongTime)
                 Text1(i).Text = FormatDateTime(showTime, vbLongTime)
                 
                 showTime = ""
                
                Next i
                
                cn.Execute "UPDATE  theater SET show1='" & Text1(0).Text & "' WHERE Screen_NO='" & Val(cmbScreen.Text) & "'"
                cn.Execute "UPDATE  theater SET show2='" & Text1(1).Text & "' WHERE Screen_NO='" & Val(cmbScreen.Text) & "'"
                cn.Execute "UPDATE  theater SET show3='" & Text1(2).Text & "' WHERE Screen_NO='" & Val(cmbScreen.Text) & "'"
                cn.Execute "UPDATE  theater SET show4='" & Text1(3).Text & "' WHERE Screen_NO='" & Val(cmbScreen.Text) & "'"
                cn.Execute "UPDATE  theater SET show5='" & Text1(4).Text & "' WHERE Screen_NO='" & Val(cmbScreen.Text) & "'"
                MsgBox "Your movie has been added successfully...!!", vbOKOnly, "ADDED"
                Unload Me
                Exit Sub
            Else
                .MoveNext
            End If
        Wend
    End With
    rs.Close
    cn.Close
    
    MsgBox "Your movie has been added successfully...!!", vbOKOnly, "ADDED"
    Unload Me
    End If
    
   End Sub

Private Sub Form_Load()
lblToday.Caption = Format(Date, "dddd, dd/MMM/yyyy")
    Set cn1 = New Connection
    Set rs1 = New Recordset
    cn1.Open "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
    rs1.Open "select * from movie", cn1, adOpenDynamic, adLockOptimistic

End Sub

Private Sub txtMovieName_KeyPress(KeyAscii As Integer)
If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
If KeyAscii = 32 And Len(txtMovieName.Text) = 0 Then
    MsgBox "Please enter a movie name"
Else
    cmbScreen.Enabled = True
    cmbShows.Enabled = True
End If
End Sub

