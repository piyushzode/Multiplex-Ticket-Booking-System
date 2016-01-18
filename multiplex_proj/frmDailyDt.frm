VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDailyDt 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFC0C0&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdGenereate 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Generate Report"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmDailyDt.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   2535
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   12648447
         ForeColor       =   16711680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmDailyDt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenereate_Click()
    
    If IsDate(MaskEdBox1) = False Then
        MsgBox "Please enter valid date"
        Load Me
    Else
       dt = Format(MaskEdBox1.Text, "dd/MMM/yyyy")
    
    If isLoad = True Then
        Unload DataReport2
        Unload DataEnvironment1
    Else
        isLoad = True
    End If

    
    DataEnvironment1.Connection1.ConnectionString = "Provider=MSDAORA.1;Password=piyush;User ID=system;Persist Security Info=True"
    DataEnvironment1.cmdDaily (dt)
    
    Unload Me
    
    Load DataReport2
    DataReport2.Show
    End If
    
End Sub


