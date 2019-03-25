VERSION 5.00
Begin VB.Form frmAmendViewCustomer 
   Caption         =   "Amend/View a Customer Screen"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
   LinkTopic       =   "Form2"
   ScaleHeight     =   7500
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      Cancel          =   -1  'True
      Caption         =   "&Back"
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
      Left            =   11520
      TabIndex        =   43
      Top             =   10200
      Width           =   1575
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "&Main Menu"
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
      Left            =   13560
      TabIndex        =   42
      Top             =   10200
      Width           =   1575
   End
   Begin VB.Data datSQL1 
      Caption         =   "SQL1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data datCustomer 
      Caption         =   "Customer"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame frmNameNumOption 
      Caption         =   "Customer Name/ID Number Option"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8775
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Width           =   9255
      Begin VB.CommandButton cmdEdit 
         Caption         =   " &Edit"
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
         Left            =   6000
         TabIndex        =   41
         Top             =   6000
         Width           =   2175
      End
      Begin VB.CommandButton cmdNextCustomer 
         Caption         =   "Ne&xt Customer"
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
         Left            =   6000
         TabIndex        =   40
         Top             =   7360
         Width           =   2175
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   6000
         TabIndex        =   39
         Top             =   6680
         Width           =   2175
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
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
         Left            =   6000
         TabIndex        =   38
         Top             =   8040
         Width           =   2175
      End
      Begin VB.TextBox txtBalance 
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
         Height          =   405
         Left            =   3360
         TabIndex        =   36
         Top             =   8100
         Width           =   1935
      End
      Begin VB.OptionButton optName 
         Caption         =   "Search useing Customers Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   480
         TabIndex        =   24
         Top             =   360
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton optNumber 
         Caption         =   "Search using Customers Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5160
         TabIndex        =   23
         Top             =   360
         Width           =   3255
      End
      Begin VB.ComboBox cboName 
         Height          =   315
         ItemData        =   "frmAmendViewCustomer.frx":0000
         Left            =   5280
         List            =   "frmAmendViewCustomer.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CommandButton cmdSeekCustomer 
         Caption         =   "&Display Customer"
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
         Height          =   495
         Left            =   5880
         TabIndex        =   21
         Top             =   2280
         Width           =   2415
      End
      Begin VB.ComboBox cboNumber 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmAmendViewCustomer.frx":0004
         Left            =   5880
         List            =   "frmAmendViewCustomer.frx":0006
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtStatus 
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
         Height          =   405
         Left            =   3360
         TabIndex        =   19
         Top             =   7560
         Width           =   1935
      End
      Begin VB.TextBox txtNameFound 
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
         Height          =   405
         Left            =   3360
         TabIndex        =   18
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtPhoneNo 
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
         Height          =   405
         Left            =   3360
         TabIndex        =   17
         Top             =   5100
         Width           =   1935
      End
      Begin VB.TextBox txtAddress3 
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
         Height          =   405
         Left            =   3360
         TabIndex        =   16
         Top             =   4500
         Width           =   1935
      End
      Begin VB.TextBox txtAddress2 
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
         Height          =   405
         Left            =   3360
         TabIndex        =   15
         Top             =   3900
         Width           =   1935
      End
      Begin VB.TextBox txtAddress1 
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
         Height          =   405
         Left            =   3360
         TabIndex        =   14
         Top             =   3300
         Width           =   1935
      End
      Begin VB.TextBox txtCreditLimit 
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
         Height          =   405
         Left            =   3360
         TabIndex        =   13
         Top             =   6900
         Width           =   1935
      End
      Begin VB.TextBox txtEmail 
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
         Height          =   405
         Left            =   3360
         TabIndex        =   12
         Top             =   6300
         Width           =   1935
      End
      Begin VB.TextBox txtMobileNo 
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
         Height          =   405
         Left            =   3360
         TabIndex        =   11
         Top             =   5700
         Width           =   1935
      End
      Begin VB.TextBox txtCustID 
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
         Height          =   405
         Left            =   3360
         TabIndex        =   10
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Frame fmeMove 
         Height          =   2535
         Left            =   5880
         TabIndex        =   5
         Top             =   3120
         Visible         =   0   'False
         Width           =   3015
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "Display &Next"
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
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   1800
            Width           =   2295
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "Display Pre&vious"
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
            Height          =   495
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox txtMatches 
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
            Height          =   405
            Left            =   2280
            TabIndex        =   6
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label lblNumMatchs 
            Caption         =   "Number of Customers matching the search critera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   855
            Left            =   240
            TabIndex        =   9
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Label lblBalance 
         AutoSize        =   -1  'True
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2280
         TabIndex        =   37
         Top             =   8145
         Width           =   810
      End
      Begin VB.Label lblNameNumber 
         AutoSize        =   -1  'True
         Caption         =   "Enter Customer Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2760
         TabIndex        =   35
         Top             =   1320
         Width           =   2250
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2640
         TabIndex        =   34
         Top             =   7605
         Width           =   630
      End
      Begin VB.Label lblNameFound 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2640
         TabIndex        =   33
         Top             =   2805
         Width           =   600
      End
      Begin VB.Label lblPhoneNo 
         AutoSize        =   -1  'True
         Caption         =   "Phone Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1800
         TabIndex        =   32
         Top             =   5100
         Width           =   1485
      End
      Begin VB.Label lblAddress3 
         AutoSize        =   -1  'True
         Caption         =   "Address3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2280
         TabIndex        =   31
         Top             =   4500
         Width           =   945
      End
      Begin VB.Label lblAddress2 
         AutoSize        =   -1  'True
         Caption         =   "Address2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2280
         TabIndex        =   30
         Top             =   3900
         Width           =   945
      End
      Begin VB.Label lblAddress1 
         AutoSize        =   -1  'True
         Caption         =   "Address1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2280
         TabIndex        =   29
         Top             =   3345
         Width           =   945
      End
      Begin VB.Label lblCreditLimit 
         AutoSize        =   -1  'True
         Caption         =   "Credit Limit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2040
         TabIndex        =   28
         Top             =   6900
         Width           =   1215
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "e-mail"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2640
         TabIndex        =   27
         Top             =   6300
         Width           =   615
      End
      Begin VB.Label lblMobileNo 
         AutoSize        =   -1  'True
         Caption         =   "Mobile Number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Top             =   5700
         Width           =   1590
      End
      Begin VB.Label lblCustID 
         AutoSize        =   -1  'True
         Caption         =   "Customer ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1920
         TabIndex        =   25
         Top             =   2325
         Width           =   1305
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Main Menu"
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
      Left            =   13320
      TabIndex        =   2
      Top             =   12225
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Back"
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
      Left            =   9120
      TabIndex        =   1
      Top             =   12225
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   """Sub-Menu"""
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
      Left            =   11160
      TabIndex        =   0
      Top             =   12225
      Width           =   1575
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Amend/View a Customer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   825
      Left            =   3885
      TabIndex        =   3
      Top             =   600
      Width           =   7785
   End
End
Attribute VB_Name = "frmAmendViewCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose        This form allows the user to select a customer either by
'               picking a customer ID from a drop down list,picking a
'               customer name from a drop down list or by entering part of
'               a name and searching through any names that match the criteria
'               The Customers details can then be saved or ammended. If a
'               customers details are amended they will be temperarly stored
'               in the associated tag of the relevent text box until the save
'               comand is selected. If the cancel comand is selected then all
'               amendments not saved will be deleted
'Student        David Hamilton
'StudentID      Com2023
'Last Modified  15/3/02

'Variables used
'varName        Variant
'varAdd1        Variant
'varAdd2        Variant
'varAdd3        Variant
'varPhone       Variant      (Declared as a varant because they are being
'varMob         Variant       used with the split Function)
'varEmail       Variant
'varCrLim       Variant
'varNum         Variant
'strName        String
'strID          String
'intID          Integer
'intNumber      Integer
'intSaveCount   Integer(Form level)


Dim intSaveCount As Integer
'Clears the screen, clears the tags and resets the form level variabl
'intSaveCount to 0

Private Sub cmdBack_Click()
frmCustomerFileProcessing.Show
Unload Me
End Sub

Private Sub cmdCancel_Click()

ClearAll 'Procedure to clear Input boxes

If txtNameFound.Tag <> "" Then
    txtNameFound.Tag = ""
    txtAddress1.Tag = ""
    txtAddress2.Tag = ""
    txtAddress3.Tag = ""
    txtPhoneNo.Tag = ""
    txtMobileNo.Tag = ""
    txtEmail.Tag = ""
    txtCreditLimit.Tag = ""
    intSaveCount = 0
End If
End Sub
'Enables relevent text boxes to allow editing
Private Sub cmdEdit_Click()

txtNameFound.Enabled = True
txtAddress1.Enabled = True
txtAddress2.Enabled = True
txtAddress3.Enabled = True
txtPhoneNo.Enabled = True
txtMobileNo.Enabled = True
txtEmail.Enabled = True
txtCreditLimit.Enabled = True

End Sub

Private Sub cmdMainMenu_Click()
frmMainMenu.Show
Unload Me
End Sub

'selecting this button allows you to display another Customer
'The inputs are not saved at this stage but
'instead moved to the tags of the relevent text boxes, the
'text boxes are then cleared to make way for the next inputs
Private Sub cmdNextCustomer_Click()
If txtNameFound.Tag = "" Then
    txtCustID.Tag = txtCustID
    txtNameFound.Tag = txtNameFound
    txtAddress1.Tag = txtAddress1
    txtAddress2.Tag = txtAddress2
    txtAddress3.Tag = txtAddress3
    txtPhoneNo.Tag = txtPhoneNo
    txtMobileNo.Tag = txtMobileNo
    txtEmail.Tag = txtEmail
    txtCreditLimit.Tag = txtCreditLimit
    intSaveCount = 1
Else
    txtCustID.Tag = txtCustID.Tag & Chr(9) & txtCustID
    txtNameFound.Tag = txtNameFound.Tag & Chr(9) & txtNameFound
    txtAddress1.Tag = txtAddress1.Tag & Chr(9) & txtAddress1
    txtAddress2.Tag = txtAddress2.Tag & Chr(9) & txtAddress2
    txtAddress3.Tag = txtAddress3.Tag & Chr(9) & txtAddress3
    txtPhoneNo.Tag = txtPhoneNo.Tag & Chr(9) & txtPhoneNo
    txtMobileNo.Tag = txtMobileNo.Tag & Chr(9) & txtMobileNo
    txtEmail.Tag = txtEmail.Tag & Chr(9) & txtEmail
    txtCreditLimit.Tag = txtCreditLimit.Tag & Chr(9) & txtCreditLimit
    intSaveCount = intSaveCount + 1
End If

ClearAll 'Procedure to clear Input boxes
End Sub
'Allows the user to search through all records which match the initial criteria
Private Sub cmdPrevious_Click()

Dim strName As String

strName = cboName.Text
datCustomer.Recordset.FindPrevious ("Name like '*" & strName & "*'")
    
DisplayCustomer ' Displays the customer that relates to the above criteria
End Sub
'clears the text boxes and sets the "Enabled" settings of relevent input boxes
Private Sub ClearAll()
txtCustID = ""
txtNameFound = ""
txtAddress1 = ""
txtAddress2 = ""
txtAddress3 = ""
txtPhoneNo = ""
txtMobileNo = ""
txtEmail = ""
txtCreditLimit = ""
txtStatus = ""
txtBalance = ""

fmeMove.Visible = False
optName = True
cboName.Text = ""
cboNumber.Text = ""
optName.Enabled = True
optNumber.Enabled = True
cboName.Enabled = True
cmdSeekCustomer.Enabled = True
cboName.SetFocus
End Sub
'Allows the user to search through all records which match the initial criteria
Private Sub cmdNext_Click()
Dim strName As String

strName = cboName.Text
datCustomer.Recordset.FindNext ("Name like '*" & strName & "*'")
    
DisplayCustomer ' Displays the customer that relates to the above criteria
    
End Sub
'Saves the information presently displayed in the relevant
'input boxes as well as the information stored in their associated tags
Private Sub cmdSave_Click()
Dim intNumber As Integer
Dim varName, varAdd1, varAdd2, varAdd3, varPhone, varMob, varEmail, varCrLim, varNum As Variant

intResponse = MsgBox("Click OK to save  all inputted Data", vbOKCancel, "SAVE NOW")
If intResponse = 1 Then
        
    If txtCustID.Text <> "" Then
        intNumber = Int(Val(txtCustID.Text))
        datCustomer.Recordset.FindFirst "Cust_ID = " & intNumber & ""
        datCustomer.Recordset.Edit
        datCustomer.Recordset("Name") = txtNameFound
        datCustomer.Recordset("Address 1") = txtAddress1
        datCustomer.Recordset("Address 2") = txtAddress2
        datCustomer.Recordset("Address 3") = txtAddress3
        datCustomer.Recordset("Phone No") = txtPhoneNo
        datCustomer.Recordset("Mobile No") = txtMobileNo
        datCustomer.Recordset("E-Mail") = txtEmail
        datCustomer.Recordset("Credit Limit") = txtCreditLimit
        datCustomer.Recordset.Update
            
        ClearAll 'Procedure to clear Input boxes
           
    End If
        
    If txtNameFound.Tag <> "" Then
        If InStr(txtNameFound.Tag, Chr(9)) <> 0 Then
            varNum = Split(txtCustID.Tag, Chr(9))
            varName = Split(txtNameFound.Tag, Chr(9))
            varAdd1 = Split(txtAddress1.Tag, Chr(9))
            varAdd2 = Split(txtAddress2.Tag, Chr(9))
            varAdd3 = Split(txtAddress3.Tag, Chr(9))
            varPhone = Split(txtPhoneNo.Tag, Chr(9))
            varMob = Split(txtMobileNo.Tag, Chr(9))
            varEmail = Split(txtEmail.Tag, Chr(9))
            varCrLim = Split(txtCreditLimit.Tag, Chr(9))
                
            For intindex = 0 To intSaveCount - 1
                intNumber = Val(varNum(intindex))
                datCustomer.Recordset.FindFirst "Cust_ID = " & intNumber & ""
                datCustomer.Recordset.Edit
                datCustomer.Recordset("Name") = varName(intindex)
                datCustomer.Recordset("Address 1") = varAdd1(intindex)
                datCustomer.Recordset("Address 2") = varAdd2(intindex)
                datCustomer.Recordset("Address 3") = varAdd3(intindex)
                datCustomer.Recordset("Phone No") = varPhone(intindex)
                datCustomer.Recordset("Mobile No") = varMob(intindex)
                datCustomer.Recordset("E-Mail") = varEmail(intindex)
                datCustomer.Recordset("Credit Limit") = Val(varCrLim(intindex))
                datCustomer.Recordset.Update
            Next intindex
            cmdCancel_Click
            
        Else
            intNumber = Val(txtCustID.Tag)
            datCustomer.Recordset.FindFirst "Cust_ID = " & intNumber & ""
            datCustomer.Recordset.Edit
            datCustomer.Recordset("Name") = txtNameFound.Tag
            datCustomer.Recordset("Address 1") = txtAddress1.Tag
            datCustomer.Recordset("Address 2") = txtAddress2.Tag
            datCustomer.Recordset("Address 3") = txtAddress3.Tag
            datCustomer.Recordset("Phone No") = txtPhoneNo.Tag
            datCustomer.Recordset("Mobile No") = txtMobileNo.Tag
            datCustomer.Recordset("E-Mail") = txtEmail.Tag
            datCustomer.Recordset("Credit Limit") = txtCreditLimit.Tag
            datCustomer.Recordset.Update
            cmdCancel_Click
        End If
    End If
Else
    cmdCancel_Click
End If

        
        
End Sub
'When the form is loaded drop down lists of both cboName and cboNumber are
'filled from the customer table
Private Sub Form_Load()
datCustomer.DatabaseName = strThePath
datCustomer.RecordSource = "Customer"
datCustomer.Refresh

While Not datCustomer.Recordset.EOF
    cboName.AddItem datCustomer.Recordset("Name")
    cboNumber.AddItem datCustomer.Recordset("Cust_ID")
    datCustomer.Recordset.MoveNext
Wend


End Sub

Private Sub optName_Click()

lblNameNumber = "Enter Customers Name"
cboName = ""
cboName.Visible = True
cboNumber.Visible = False
cboName.Enabled = True
cboName.SetFocus

End Sub

Private Sub optNumber_Click()

lblNameNumber = "Enter Customers ID Number"
cboName.Visible = False
cboNumber.Visible = True
cboNumber = ""
cboNumber.Enabled = True
cboNumber.SetFocus

End Sub
Private Sub cboName_Change()
cmdSeekCustomer.Enabled = True
End Sub

Private Sub cboName_Click()
cmdSeekCustomer.Enabled = True
End Sub
Private Sub cboNumber_Change()
cmdSeekCustomer.Enabled = True
End Sub

Private Sub cboNumber_Click()
cmdSeekCustomer.Enabled = True

End Sub
'Using the Name/Number/letters entered a search is made of the
'Customer Table to return the first match, the coresponding
'record is then displayed. If a search is made using a name
'and their is more than one match, the user is then allowed
'step through the records of those with matching names
Private Sub cmdSeekCustomer_Click()
Dim strName, strID As String
Dim intNumber, intID As Integer


cmdPrevious.Enabled = True
cmdNext.Enabled = True

strID = cboName.Text

SqlCount = "SELECT Count(*)AS IDCount FROM  Customer  WHERE Name Like '*" & strID & "*';"


datSQL1.DatabaseName = strThePath
datSQL1.RecordSource = SqlCount
datSQL1.Refresh


intCount = datSQL1.Recordset.Fields("IDCount")

If intCount > 1 And cboName.Visible = True Then  'Displays matching records if their is more than 1
    fmeMove.Visible = True
    txtMatches = intCount
    cmdPrevious.Enabled = True
    cmdNext.Enabled = True
End If
If optName = True Then
    strName = cboName.Text
    If Not Len(strName) = 0 Then
        datCustomer.Recordset.FindFirst "Name like '*" & strName & "*'"
        
        txtCustID = datCustomer.Recordset("Cust_ID")
        txtNameFound = datCustomer.Recordset("Name")
        txtAddress1 = datCustomer.Recordset("Address 1")
        txtAddress2 = datCustomer.Recordset("Address 2")
        txtAddress3 = datCustomer.Recordset("Address 3")
        txtPhoneNo = datCustomer.Recordset("Phone No")
        txtMobileNo = datCustomer.Recordset("Mobile No")
        txtEmail = datCustomer.Recordset("E-Mail")
        txtCreditLimit = datCustomer.Recordset("Credit Limit")
        txtStatus = datCustomer.Recordset("Status")
        txtBalance = datCustomer.Recordset("Balance owed")
    End If
Else
    intNumber = Int(Val(cboNumber.Text))
    
        datCustomer.Recordset.FindFirst "Cust_ID = " & intNumber & ""
 
        txtCustID = datCustomer.Recordset("Cust_ID")
        txtNameFound = datCustomer.Recordset("Name")
        txtAddress1 = datCustomer.Recordset("Address 1")
        txtAddress2 = datCustomer.Recordset("Address 2")
        txtAddress3 = datCustomer.Recordset("Address 3")
        txtPhoneNo = "" & datCustomer.Recordset("Phone No")
        txtMobileNo = "" & datCustomer.Recordset("Mobile No")
        txtEmail = "" & datCustomer.Recordset("E-Mail")
        txtCreditLimit = datCustomer.Recordset("Credit Limit")
        txtStatus = datCustomer.Recordset("Status")
        txtBalance = "" & datCustomer.Recordset("Balance owed")
        
        
End If


End Sub
'A procedure the contents of a related record
Private Sub DisplayCustomer()

txtCustID = datCustomer.Recordset("Cust_ID")
txtNameFound = datCustomer.Recordset("Name")
txtAddress1 = datCustomer.Recordset("Address 1")
txtAddress2 = datCustomer.Recordset("Address 2")
txtAddress3 = datCustomer.Recordset("Address 3")
txtPhoneNo = "" & datCustomer.Recordset("Phone No")
txtMobileNo = "" & datCustomer.Recordset("Mobile No")
txtEmail = "" & datCustomer.Recordset("E-Mail")
txtCreditLimit = datCustomer.Recordset("Credit Limit")
txtStatus = datCustomer.Recordset("Status")
txtBalance = "" & datCustomer.Recordset("Balance owed")

End Sub
Private Sub SetScreen()
fmeMove.Visible = False
optName = True
cboName.Text = ""
cboNumber.Text = ""
optName.Enabled = False
optNumber.Enabled = False
cboName.Enabled = False
cmdSeekCustomer.Enabled = False
End Sub

