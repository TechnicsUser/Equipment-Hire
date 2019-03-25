VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEquipmentHiredReport 
   Caption         =   "Equipment Hired Report"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10140
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Height          =   495
      Left            =   11400
      TabIndex        =   22
      Top             =   10200
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid thegrid 
      Bindings        =   "frmEquipmentHiredReport.frx":0000
      Height          =   3255
      Left            =   480
      TabIndex        =   21
      Top             =   5160
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   5741
      _Version        =   393216
      ForeColor       =   8388608
      AllowUserResizing=   3
   End
   Begin VB.Data datCustomer 
      Caption         =   "cutomer"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data datRental 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "datEquipment"
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox cmbmonth 
      Height          =   405
      Left            =   2160
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ComboBox cmbyear 
      Height          =   405
      Left            =   3480
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ComboBox CmbDate 
      Height          =   405
      Left            =   840
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdMainMenu 
      Caption         =   "&Main Menu"
      Height          =   495
      Left            =   13320
      TabIndex        =   2
      Top             =   10200
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Options"
      ForeColor       =   &H00800000&
      Height          =   3135
      Left            =   480
      TabIndex        =   7
      Top             =   1200
      Width           =   14295
      Begin VB.OptionButton OptEQHired 
         Caption         =   "List All Equipment Hired"
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   10680
         TabIndex        =   13
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox CmbEqCust 
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   4920
         TabIndex        =   12
         Top             =   1440
         Width           =   4575
      End
      Begin VB.OptionButton OptEqCust 
         Caption         =   "List All Equipment Hired By A Particular Customer "
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   4920
         TabIndex        =   11
         Top             =   720
         Width           =   4455
      End
      Begin VB.CommandButton cmdDisplayDate 
         Caption         =   "Display"
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   2280
         Width           =   975
      End
      Begin VB.OptionButton Optdate 
         Caption         =   "List Equipment Hired On A Certain Date"
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   480
         TabIndex        =   8
         Top             =   720
         Width           =   3135
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   4215
         Begin VB.Label Label4 
            Caption         =   "Year"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   2760
            TabIndex        =   6
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Month"
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   1440
            TabIndex        =   20
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Day"
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2415
         Left            =   4800
         TabIndex        =   14
         Top             =   360
         Width           =   4815
      End
      Begin VB.Frame Frame4 
         Height          =   2415
         Left            =   9960
         TabIndex        =   15
         Top             =   360
         Width           =   3975
         Begin VB.Frame Frame5 
            Caption         =   "Order"
            ForeColor       =   &H00800000&
            Height          =   1335
            Left            =   240
            TabIndex        =   16
            Top             =   840
            Width           =   3495
            Begin VB.OptionButton Optorderbydate 
               Caption         =   "Hire Date"
               ForeColor       =   &H00800000&
               Height          =   495
               Left            =   240
               TabIndex        =   18
               Top             =   720
               Width           =   1935
            End
            Begin VB.OptionButton optOrdebyrCust 
               Caption         =   "Customer Rental ID "
               ForeColor       =   &H00800000&
               Height          =   375
               Left            =   240
               TabIndex        =   17
               Top             =   360
               Width           =   2895
            End
         End
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label Sample add a bit"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5040
      TabIndex        =   1
      Top             =   1800
      Width           =   2400
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Equipment Hired Report"
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
      Left            =   3900
      TabIndex        =   0
      Top             =   240
      Width           =   7755
   End
End
Attribute VB_Name = "frmEquipmentHiredReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Purpose of screen:


'The purpose of this screen is to display equipment hired in different
'formats such as:
    '1listing equipment hired on a certain date
    '2listing equipment hired by a particular customer
    '3listing all equipment presently on hire  which the user can display either
    'by  date ored or customer number order


'Author: Derek Stafford com2026
'15/03/2002


'variable used                            purpose

'Inindex                 integer;         used to control for loops
'intCustID               integer;         store customer id number
'srtTime1                string;          used to hold date
'srtTime2                string;          used to hold date
      
'objects used

'datRental               datacontrol;     used to retrieve info from rental/return table
'datCustomer             datacontrol;     used to retrieve info from customer table
'thegrid                 flexigrid ;      bound to the datRental datacontrol















Option Explicit
Dim intCustID As Integer
Private Sub CmbEqCust_click()

Dim strmySQl As String
'list equipment hired by a certain customer

intCustID = CmbEqCust.ItemData(CmbEqCust.ListIndex)
    
strmySQl = "SELECT [Rental/Return].[Date/Time out] as [Hire Date],Customer.Cust_ID as [Customer No], " & _
"Customer.Name,Customer.[Address 1],Customer.[Address 2], [Rental/Return].Equipment_ID as [Equipment ID], " & _
"[Rental/Return].Rental_ID as [Rental ID No],[Rental/Return].[Date/Time returned]as [Return Date] " & _
"FROM Customer INNER JOIN [Rental/Return] ON Customer.Cust_ID = [Rental/Return].Cust_ID " & _
"Where (((Customer.Cust_ID) = " & intCustID & "))" & _
"ORDER BY Customer.Cust_ID;"
datRental.RecordSource = strmySQl
datRental.Refresh
End Sub

Private Sub CmbEqCust_DropDown()
datCustomer.RecordSource = "Customer"

'put customer details into combo box

CmbEqCust.Clear

datCustomer.Recordset.MoveFirst
While Not datCustomer.Recordset.EOF
    CmbEqCust.AddItem (datCustomer.Recordset("Name")) + " ," + (datCustomer.Recordset("Address 1")) + ", " + (datCustomer.Recordset("Address 2"))
    CmbEqCust.ItemData(CmbEqCust.NewIndex) = datCustomer.Recordset("Cust_ID")
    datCustomer.Recordset.MoveNext
Wend
End Sub

Private Sub cmbmonth_DropDown()
'fills combo box with  month 0f year

Dim intindex As Integer
cmbmonth.Clear
For intindex = 1 To 12
    cmbmonth.AddItem (intindex)
Next intindex
End Sub
Private Sub cmbDate_DropDown()
'fills combo box with days of month

Dim intindex As Integer
CmbDate.Clear
For intindex = 1 To 31
    CmbDate.AddItem (intindex)
Next intindex
End Sub
Private Sub cmbyear_DropDown()
'fills combo box with yaers 2001-2009

Dim intindex As Integer

For intindex = 2001 To 2009
    cmbyear.AddItem (intindex)
Next intindex
End Sub

Private Sub cmdBack_Click()
frmReportsMenu.Show
Unload Me
End Sub

Private Sub cmdDisplayDate_Click()
'lists equipment hired on a particular date

Dim strTime1, strTime2, strmySQl As String
    
If CmbDate.ListIndex = -1 Or cmbyear.ListIndex = -1 Or cmbmonth.ListIndex = -1 Then
    Call MsgBox("The date entry is not complete or not valid ", vbCritical, "Entry Not Complete")
Else
    strTime1 = CmbDate & " / " & cmbmonth & " / " & cmbyear
    'strTime2 = "#" & CmbDate + 1 & "/" & cmbmonth & "/" & cmbyear & "#"
    
    If IsDate(strTime1) Then
        
        strTime1 = "#" & CmbDate & " / " & cmbmonth & " / " & cmbyear & "#"
        strTime2 = "#" & CmbDate + 1 & "/" & cmbmonth & "/" & cmbyear & "#"
        
        
        strmySQl = "SELECT [Rental/Return].[Date/Time out] as [Hire Date],Customer.Cust_ID as [Customer No], " & _
        "Customer.Name,Customer.[Address 1],Customer.[Address 2], [Rental/Return].Equipment_ID as [Equipment ID], " & _
        "[Rental/Return].Rental_ID as [Rental ID No],[Rental/Return].[Date/Time returned]as [Return Date] " & _
        "FROM Customer INNER JOIN [Rental/Return] ON Customer.Cust_ID = [Rental/Return].Cust_ID " & _
        "Where ((([Rental/Return].[Date/Time out]) >= " & strTime1 & " And ([Rental/Return].[Date/Time out]) <" & strTime2 & "))" & _
        "ORDER BY Customer.Cust_ID;"
     
        datRental.RecordSource = strmySQl
        datRental.Refresh
    
    Else
        Call MsgBox("Invalid date entry", vbCritical, "invalid date")
    End If
End If
End Sub

Private Sub cmdMainMenu_Click()
 frmMainMenu.Show
 Unload Me
End Sub

Private Sub Form_Activate()

Optorderbydate.Enabled = False
optOrdebyrCust.Enabled = False
CmbEqCust.Enabled = False
CmbDate.Enabled = False
cmbyear.Enabled = False
cmbmonth.Enabled = False
cmdDisplayDate.Enabled = False
cmdDisplayDate.ToolTipText = "Displays the equipment hire on the choosen date"
thegrid.Clear
End Sub
Private Sub Form_Load()
'assigns the tables of the data base to the controls

datRental.DatabaseName = "g:\EHS.mdb"
datRental.RecordSource = "Equipment"
datCustomer.DatabaseName = "g:\EHS.mdb"

datCustomer.RecordSource = "Customer"

   
End Sub
Private Sub optDate_Click()
'if date option is clicked
CmbDate.Text = ""
cmbmonth.Text = ""
cmbyear.Text = ""
thegrid.Clear
CmbDate.Enabled = True
cmbmonth.Enabled = True
cmbyear.Enabled = True
cmdDisplayDate.Enabled = True
CmbEqCust.Text = ""
CmbEqCust.Enabled = False
Optorderbydate.Enabled = False
optOrdebyrCust.Enabled = False
End Sub

Private Sub OptEqCust_Click()
cmbmonth.Clear
CmbDate.Clear
cmbyear.Clear

'thegrid.Clear
Optorderbydate.Enabled = False
optOrdebyrCust.Enabled = False
CmbEqCust.Enabled = True
CmbDate.Enabled = False
cmbyear.Enabled = False
cmbmonth.Enabled = False
cmdDisplayDate.Enabled = False
End Sub

Private Sub OptEQHired_Click()
Dim strmySQl As String
'lists all the equipment hired
cmbmonth.Clear
CmbDate.Clear
cmbyear.Clear

CmbEqCust.Text = ""
CmbEqCust.Enabled = False
CmbDate.Enabled = False
cmbyear.Enabled = False
cmbmonth.Enabled = False
cmdDisplayDate.Enabled = False
Optorderbydate.Enabled = True
optOrdebyrCust.Enabled = True
  
strmySQl = "SELECT [Rental/Return].[Date/Time out] as [Hire Date],Customer.Cust_ID as [Customer No], " & _
"Customer.Name,Customer.[Address 1],Customer.[Address 2], [Rental/Return].Equipment_ID as [Equipment ID], " & _
"[Rental/Return].Rental_ID as [Rental ID No],[Rental/Return].[Date/Time returned]as [Return Date] " & _
"FROM [Rental/Return] INNER JOIN Customer ON [Rental/Return].Cust_ID = Customer.Cust_ID " & _
"Where ((([Rental/Return].[Date/Time returned]) Is Null)) " & _
"ORDER BY [Rental/Return].[Date/Time out];"

Optorderbydate.Value = True
datRental.RecordSource = strmySQl
datRental.Refresh
End Sub
Private Sub Optorderbydate_Click()
'order all equipment hired by date

Dim strmySQl As String, intEquipID, intEquipType As Integer
    
strmySQl = "SELECT [Rental/Return].[Date/Time out] as [Hire Date],Customer.Cust_ID as [Customer No], " & _
"Customer.Name,Customer.[Address 1],Customer.[Address 2],[Rental/Return].Equipment_ID as [Equipment ID], " & _
"[Rental/Return].Rental_ID as [Rental ID No],[Rental/Return].[Date/Time returned]as [Return Date] " & _
"FROM [Rental/Return] INNER JOIN Customer ON [Rental/Return].Cust_ID = Customer.Cust_ID " & _
"Where ((([Rental/Return].[Date/Time returned]) Is Null)) " & _
"ORDER BY [Rental/Return].[Date/Time out];"

datRental.RecordSource = strmySQl
datRental.Refresh
    
End Sub

Private Sub optOrdebyrCust_Click()
'order all equipment hired by  customer id

Dim strmySQl As String

strmySQl = "SELECT [Rental/Return].[Date/Time out] as [Hire Date],Customer.Cust_ID as [Customer No], " & _
"Customer.Name,Customer.[Address 1],Customer.[Address 2], [Rental/Return].Equipment_ID as [Equipment ID], " & _
"[Rental/Return].Rental_ID as [Rental ID No],[Rental/Return].[Date/Time returned]as [Return Date] " & _
"FROM Customer INNER JOIN [Rental/Return] ON Customer.Cust_ID = [Rental/Return].Cust_ID " & _
"Where ((([Rental/Return].[Date/Time returned]) Is Null))" & _
"ORDER BY [Rental/Return].Rental_ID;"

datRental.RecordSource = strmySQl
datRental.Refresh

End Sub
