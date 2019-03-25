VERSION 5.00
Begin VB.Form frmBrowse 
   Caption         =   "Browse Supplier"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstTable 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1635
      ItemData        =   "frmBrowse.frx":0000
      Left            =   1080
      List            =   "frmBrowse.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Select the person you want"
      Top             =   2160
      Width           =   13455
   End
   Begin VB.Data datTable 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8520
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Left            =   13080
      TabIndex        =   1
      ToolTipText     =   "Back "
      Top             =   9720
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
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
      Left            =   11040
      TabIndex        =   0
      ToolTipText     =   "Click to accept person"
      Top             =   9720
      Width           =   1575
   End
   Begin VB.TextBox txtID 
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
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "The persons ID"
      Top             =   5835
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "E-Mail"
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
      Height          =   255
      Left            =   12120
      TabIndex        =   12
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Address 1"
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
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Address 2"
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
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Address 3"
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
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Phone"
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
      Height          =   255
      Left            =   8520
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Mobile"
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
      Height          =   255
      Left            =   10320
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Supplier ID"
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
      Left            =   5400
      TabIndex        =   3
      Top             =   5880
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Browse Supplier"
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
      Left            =   5220
      TabIndex        =   2
      Top             =   240
      Width           =   5115
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The purpose of this screen is to allow the user to browse for a supplier or a customer.

'It allows the user to browse for the supplier or customer by placing all the persons
'details in a list box. The persons primary key is stored in the ItemData property
'of the list box.
'The user then clicks "Ok" and the primary key is passed back to the previous screen.

'Author Fergal Purcell
'Date   06/03/2002

'Variables used
'strTable   :   String ; This variable stores the name of the table to be browsed,
                         'e.g, "Customer" or "Supplier"


'Objects used
'datTable    :   Data Control; This data control fills the list of customers or suppliers
Option Explicit

'Use a DLL to place tab stops in a list box.
'This piece of code is taken from "Programming in Visual Basic 6.0, by Julia Case Bradley"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
        lParam As Any) As Long
Const LB_SETTABSTOPS = &H192

Private Sub cmdCancel_Click()
If frmBrowse.Tag = "Supplier" Then          'Tag property set to "Supplier" if coming from DeleteaSupplier screen
    frmDeleteaSupplier.Show
ElseIf frmBrowse.Tag = "Customer" Then      'Tag property set to "Customer" if coming from DeleteaSupplier screen
    frmMakeaPayment.Show
Else
    frmOrdersReport.Show
End If
lstTable.Clear
Unload Me
End Sub

Private Sub cmdOk_Click()
If frmBrowse.Tag = "Supplier" Then          'Tag property set to "Supplier" if coming from DeleteaSupplier screen
    frmDeleteaSupplier.txtSearch.Text = txtID.Text
    frmDeleteaSupplier.Show
ElseIf frmBrowse.Tag = "Customer" Then      'Tag property set to "Customer" if coming from DeleteaSupplier screen
    frmMakeaPayment.txtSearch.Text = txtID.Text
    frmMakeaPayment.Show
Else
    frmOrdersReport.txtList.Text = txtID.Text
    frmOrdersReport.Show
End If
frmBrowse.Tag = ""
Unload Me
End Sub

Private Sub Form_Activate()
Dim strTable As String
SetTabs lstTable                        'Set the tabs in the list box
If frmBrowse.Tag = "Report" Then
    strTable = "Supplier"
Else
    strTable = frmBrowse.Tag                'Set which table is to be searched
End If
datTable.DatabaseName = strThePath
datTable.RecordSource = strTable
datTable.Refresh
If strTable = "Supplier" Then
    Label1.Caption = "Browse Supplier"      'Search the Supplier table
    While Not datTable.Recordset.EOF
        If datTable.Recordset.Fields("Deletion") = False Then
            lstTable.AddItem (datTable.Recordset("Supplier Name") & Chr(9) & datTable.Recordset("Address 1") & Chr(9) & datTable.Recordset("Address 2") & Chr(9) & datTable.Recordset("Address 3") & Chr(9) & datTable.Recordset("Phone No") & Chr(9) & datTable.Recordset("Mobile No") & Chr(9) & datTable.Recordset("E-Mail"))
            lstTable.ItemData(lstTable.NewIndex) = (datTable.Recordset("Supplier_ID"))      'Store the suppliers ID
        End If
        datTable.Recordset.MoveNext
    Wend
Else
    Label1.Caption = "Browse Customer"          'Search the Customer table
    Label2.Caption = "Customer ID"
    While Not datTable.Recordset.EOF
        If datTable.Recordset.Fields("Deletion") = False Then
            lstTable.AddItem (datTable.Recordset("Name") & Chr(9) & datTable.Recordset("Address 1") & Chr(9) & datTable.Recordset("Address 2") & Chr(9) & datTable.Recordset("Address 3") & Chr(9) & datTable.Recordset("Phone No") & Chr(9) & datTable.Recordset("Mobile No") & Chr(9) & datTable.Recordset("E-Mail"))
            lstTable.ItemData(lstTable.NewIndex) = (datTable.Recordset("Cust_ID"))          'Store the Customers ID
        End If
        datTable.Recordset.MoveNext
    Wend
End If
End Sub

Private Sub lstTable_Click()
Dim strTheString As String
txtID.Text = lstTable.ItemData(lstTable.ListIndex)    'Get the Id if the person the user selected

End Sub

Private Sub SetTabs(lst As ListBox)
'Set the tab stops in a listbox
ReDim lngTabs(0 To 7) As Long       'Eight tab stops needed
Dim lngRtn As Long                  'DLL function returns a long variable
lngTabs(0) = 70                     'Twips measurement for 1st tab stop
lngTabs(1) = 140
lngTabs(2) = 210
lngTabs(3) = 280
lngTabs(4) = 350
lngTabs(5) = 420
lngTabs(6) = 490
lngRtn = SendMessage(lst.hwnd, LB_SETTABSTOPS, 7, lngTabs(0))   'Set the stops
End Sub
