VERSION 5.00
Begin VB.Form frmrate 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton cmdok 
         BackColor       =   &H00FFC0C0&
         Caption         =   "OK"
         Height          =   375
         Left            =   720
         MouseIcon       =   "frmMonthly.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cmbClass 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmMonthly.frx":030A
         Left            =   1800
         List            =   "frmMonthly.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFC0C0&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2640
         MouseIcon       =   "frmMonthly.frx":0333
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Fare"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pick a Class"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmrate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim flag As Integer

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdok_Click()
flag = 0
On Error GoTo t1
Set cn = New Connection
Set rs = New Recordset


cn.Open "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
rs.Open "select * from class", cn, adOpenDynamic, adLockOptimistic

cn.Execute "update class set rate=" & Text1.Text & " where classtype='" & cmbClass.Text & "'"
flag = 1

MsgBox "Your Rates have been UPDATED...!!", vbInformation, "MULTIPLEX"
Unload Me

If flag = 0 Then
t1:
MsgBox "Enter valid price"
Text1.Text = ""
Text1.SetFocus
End If


End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If IsNumeric(Text1) = False Then
MsgBox "Enter valid data...", vbCritical, "MULTIPLEX"
Text1 = ""
Cancel = False
End If

End Sub
