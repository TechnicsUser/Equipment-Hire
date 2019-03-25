VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8985
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   11700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   8175
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   11175
      Begin MSComDlg.CommonDialog cdbEHS 
         Left            =   4680
         Top             =   6240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Timer tmrTimer 
         Interval        =   3000
         Left            =   10680
         Top             =   1560
      End
      Begin VB.Image Image4 
         Height          =   735
         Left            =   960
         Picture         =   "FRMSPL~1.frx":0000
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   9135
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H80000004&
         Caption         =   "info@tool.com"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7800
         TabIndex        =   9
         Top             =   7410
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H80000004&
         Caption         =   "Company: Tools4Hire Inc."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   8
         Top             =   6960
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000004&
         Caption         =   "www.tools.com"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   7
         Top             =   7170
         Width           =   1455
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "LicenseTo: MicroComp Inc"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   6
         Top             =   6480
         Width           =   2745
      End
      Begin VB.Image Image1 
         Height          =   2055
         Left            =   1560
         Picture         =   "FRMSPL~1.frx":385A
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2415
      End
      Begin VB.Image Image2 
         Height          =   2055
         Left            =   4080
         Picture         =   "FRMSPL~1.frx":4700
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2655
      End
      Begin VB.Image Image3 
         Height          =   2055
         Left            =   6840
         Picture         =   "FRMSPL~1.frx":5FD0
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sean O Brien          COM2020"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   6960
         Width           =   2970
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fergal Purcell         COM2019"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   7320
         Width           =   2970
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Derek Stafford        COM2026"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   7680
         Width           =   2970
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dave Hamilton       COM2023"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   6600
         Width           =   2970
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Authors:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   6240
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub tmrTimer_Timer()
frmMainMenu.Show
Unload Me
End Sub

Private Sub Form_Load()
cdbEHS.DialogTitle = "Open database"
cdbEHS.Filter = "Database (*.mdb)|*.mdb|All Files (*.*)|*.*"
On Error GoTo HandleCancelError
cdbEHS.ShowOpen
Do While cdbEHS.FileTitle <> "EHS.mdb"
    On Error GoTo HandleCancelError
    MsgBox "Please select the database EHS.mdb", vbExclamation, "Incorrect database"
    cdbEHS.ShowOpen
Loop
strThePath = cdbEHS.FileName
tmrTimer.Enabled = True
tmrTimer.Interval = 3000
Exit Sub
HandleCancelError:
    End
End Sub
