VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_OutStWithClientInvoiceDetails 
   Caption         =   "Invoices outstanding Report (With Client)"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11265
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   11265
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   3615
      Left            =   480
      TabIndex        =   37
      Top             =   5280
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6376
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame2 
      Height          =   6495
      Left            =   120
      TabIndex        =   34
      Top             =   2880
      Width           =   15000
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6135
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   10821
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2805
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbRegion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   4245
      End
      Begin VB.CommandButton CmdMail 
         Caption         =   "Send Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13605
         TabIndex        =   36
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show ClientName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13200
         TabIndex        =   32
         Top             =   1440
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.ComboBox CmbStateID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   3765
      End
      Begin VB.TextBox TxtSearchBy 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   30
         Top             =   2220
         Width           =   3825
      End
      Begin VB.ComboBox CmbSearchBy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmRpt_OutStWithClientInvoiceDetails.frx":0000
         Left            =   9600
         List            =   "FrmRpt_OutStWithClientInvoiceDetails.frx":000D
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   1725
         Width           =   3240
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13020
         TabIndex        =   26
         Top             =   900
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   14220
         TabIndex        =   25
         Top             =   900
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox CmbDPName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1365
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox CmbCNonNcdexClient 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1365
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox CmbClientDP 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5160
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1725
         Width           =   4335
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Show DP "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11640
         TabIndex        =   5
         Top             =   960
         Width           =   1140
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12270
         TabIndex        =   12
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10935
         TabIndex        =   11
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox CmbExchangeTypeID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8250
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   900
         Width           =   3240
      End
      Begin VB.ComboBox CmbLocation 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   900
         Width           =   4245
      End
      Begin VB.TextBox TxtCDPID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1725
         Width           =   1215
      End
      Begin VB.TextBox TxtCDPName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1725
         Width           =   3135
      End
      Begin VB.TextBox TxtCClientName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1725
         Width           =   3495
      End
      Begin VB.TextBox TxtClientID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1725
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Region"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   660
      End
      Begin VB.Label LblStateID 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1470
         TabIndex        =   31
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label lblSearchBy 
         Caption         =   "TI No (Voucher No)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10725
         TabIndex        =   28
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label28 
         Caption         =   "Client ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   420
         TabIndex        =   24
         Top             =   1395
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "Client Name "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   23
         Top             =   1395
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "DP ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5535
         TabIndex        =   22
         Top             =   1395
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "DP Name "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   21
         Top             =   1395
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "ClientDPID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   20
         Top             =   1395
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "_"
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
         Left            =   8940
         TabIndex        =   16
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Total Record Found :-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6540
         TabIndex        =   15
         Top             =   2280
         Width           =   2280
      End
      Begin VB.Label LblExchangeType 
         Caption         =   "Exchange Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   14
         Top             =   615
         Width           =   1920
      End
      Begin VB.Label Label8 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5655
         TabIndex        =   13
         Top             =   615
         Width           =   840
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   120
      TabIndex        =   33
      Top             =   9480
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmRpt_OutStWithClientInvoiceDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCommodityID, mLocationID As Variant
Dim mExchangeTypeID As Double
Dim mClientIDRow As Double
Dim mClientNAVCode As String
Dim mClientRNAVCode As String
Dim mDPNAVCode As String
Dim mStateID, mRegionID As Variant

Private Sub CmbLocation_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Select State!!!"
   CmbStateID.SetFocus
   Exit Sub
End If

tmp = CmbLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where StateID=" & mStateID & " ")
Else
   Call TableMst_Location("Where StateID =" & mStateID & " And " & Gen_WcLocation)
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmbRegion_GotFocus()

tmp = CmbRegion.Text

Call TableMst_Region

CmbRegion.Clear

If RsMst_Region.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Region.RecordCount
    CmbRegion.AddItem RsMst_Region!Region
    RsMst_Region.MoveNext
Next
CmbRegion.Text = tmp


End Sub

Private Sub CmbRegion_LostFocus()

If CmbRegion.Text = "" Then
   Exit Sub
End If
RsMst_Region.MoveFirst
RsMst_Region.Find "Region = '" & CmbRegion.Text & "'"
If RsMst_Region.EOF Then
   MsgBox "Invalid selection "
   CmbRegion.SetFocus
   Exit Sub
End If
mRegionID = RsMst_Region!RegionID

End Sub

Private Sub CmbSearchBy_LostFocus()

If CmbSearchBy.Text = "" Then
   lblSearchBy.Caption = "TI No (Voucher No)"
   TxtSearchBy = ""
   TxtSearchBy.Locked = True
   Exit Sub
End If

If LCase(CmbSearchBy.Text) <> "ti no" And LCase(CmbSearchBy.Text) <> "post invoice no" And LCase(CmbSearchBy.Text) <> "entry no" Then
   MsgBox "Invalid Selection!!!"
   CmbSearchBy.SetFocus
   Exit Sub
End If

lblSearchBy.Caption = CmbSearchBy.Text
TxtSearchBy.Locked = False

End Sub

Private Sub CmbStateID_GotFocus()
If CmbRegion.Text = "" Then
   MsgBox "Select Region!!!"
   CmbRegion.SetFocus
   Exit Sub
End If

tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where RegionID = " & mRegionID & " ANd " & Gen_WcState)
Else
   Call TableMst_State("Where RegionID = " & mRegionID & "")
End If
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub

Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmdClear_Click()
Label2.Caption = ""
CmbRegion.Text = ""
CmbStateID.Text = ""
CmbLocation.Text = ""
Call Grid_Head
'Call Grid_Head1
Label2.Caption = ""
TxtClientID = ""
TxtCClientName = ""
TxtCDPID = ""
TxtCDPName = ""
CmbClientDP.Text = ""
CmbClientDP.Visible = False
CmbExchangeTypeID.Text = ""
mExchangeTypeID = 0
CmbCNonNcdexClient.Text = ""
Check3.Value = 0
CmbDPName.Text = ""
CmdExcel.Enabled = False
CmdMail.Enabled = False
Option1.Value = True
'SSTab1.Tab = 0
CmbSearchBy.Text = ""
TxtSearchBy = ""
TxtSearchBy.Locked = True
CmbCNonNcdexClient.Width = 9400
CmbCNonNcdexClient.Top = TxtClientID.Top
CmbCNonNcdexClient.Left = TxtClientID.Left
CmbCNonNcdexClient.Visible = True
Call Check3_Click
'Check1.Visible = False
'Check1.Value = 1
'If LCase(Gen_UserRole) = "power" Then
'   Check1.Visible = True
'   Check1.Value = 1
'End If
ProgressBar1.Value = 0

End Sub

Private Sub CmdGetReport_Click()

Call Grid_Head
Call Grid_Head1
ProgressBar1.Value = 0
MsgBox "Wait till next message!!!"

'If CmbExchangeTypeID.Text = "" And CmbLocation.Text = "" Then
'   MsgBox "Select either Location or Client!!"
'   Exit Sub
'End If

If CmbExchangeTypeID.Text <> "" Then
   If Check3.Value = 0 Then
'      If CmbClientDP.Visible = True Then
'         MsgBox "Invalid selection !!! "
'         Exit Sub
'      End If
      If CmbCNonNcdexClient.Visible = True Then
         If CmbCNonNcdexClient.Text = "" Then
            MsgBox "Blank Client Name not allowed!!!" '"Invalid selection !!! "
            CmbCNonNcdexClient.SetFocus
            Exit Sub
         End If
      End If
'      If TxtClientID = "" Then
'         MsgBox "Invalid Client ID !!! "
'         Exit Sub
'      End If
   Else
      If CmbDPName.Visible = True Then
         If CmbDPName.Text = "" Then
            MsgBox "Blank DP Name Not allowed !!! "
            CmbDPName.SetFocus
            Exit Sub
         End If
      End If
   End If
End If

''''tmp = " Select TINo,TIDate,TIAmount,(select Sum(RecAmount) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID ) as RecAmount,"
''''tmp = tmp & " (select Sum(RecTDS) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID ) as RecTDSAmount"
''''tmp = tmp & " ,(select Sum(RecOtherDed) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID ) as RecOtherDed"
''''tmp = tmp & " ,TIID,TI.NAVCode,TI.PostInvoiceNo,TI.EntryNo,ML.LocationName,MS.StateName From  Txn_TIDetails TI "
tmp = " Select TINo, TIDate, TIAmount, PIFlag, "
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+ isnull(TI.AdjAmountFAG,0) as RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecTDSAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID= TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID )RecOtherDed"
tmp = tmp & " ,TIID,TI.NAVCode,TI.PostInvoiceNo,TI.EntryNo,ML.LocationID, ML.LocationName,MS.StateName,MR.REgion From  Txn_TIDetails TI "
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID "
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID "
tmp = tmp & " Left Join Mst_Region MR on MS.RegionID=MR.RegionID "

tmp = tmp & " Where  TINo<>'' "


'''updated on 25092013
'If Option1.Value = True Then
'   ''tmp = tmp & " And TIAmount-(select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID)>0 "
'   tmp = tmp & " And TIAmount-((select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y')+"
'   tmp = tmp & " (Select isnull(Sum(RecTDS)+Sum(RecOtherDed),0)  From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID = TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.Flag <> 'R')+AdjAmountFAG)>0"
'End If
''tmp = tmp & " And TIAmount-(select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail Where TIID=TI.TIID)>0 "

tmp = tmp & " And TIAmount-((select isnull(Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed),0) From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.NewCRID= TCRT.NewCRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.ReturnFlag<>'Y')+"
tmp = tmp & " (Select isnull(Sum(RecTDS)+Sum(RecOtherDed),0)  From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.CRTDSAdjustmentID = TCRT.CRTDSAdjustmentID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID=TI.TIID And TCR.Flag <> 'R')+AdjAmountFAG)>0"
'''updated on 25092013


If CmbRegion.Text <> "" Then
   tmp = tmp & " And (MR.Region='" & CmbRegion.Text & "' Or MR.Region is Null ) "
End If

If CmbStateID.Text <> "" Then
   If CmbExchangeTypeID.Text = "" Then
      tmp = tmp & " And (MS.StateID=" & mStateID & " ) "
   Else
      tmp = tmp & " And (MS.StateID=" & mStateID & " Or MS.StateID is Null) "
   End If
Else
   If Gen_WcState <> "" Then
      If CmbStateID.Text = "" Then
         tmp = tmp & " And (MS." & Gen_WcState & " )"
      Else
         tmp = tmp & " And (MS." & Gen_WcState & " Or MS.StateID is Null) "
      End If
   End If
End If

If CmbLocation.Text <> "" Then
   If CmbExchangeTypeID.Text = "" Then
      tmp = tmp & " And (ML.LocationID=" & mLocationID & " ) "
   Else
      tmp = tmp & " And (ML.LocationID=" & mLocationID & " Or ML.LocationID is Null) "
   End If
Else
   If Gen_WcLocation <> "" Then
      If CmbExchangeTypeID.Text = "" Then
         tmp = tmp & " And (ML." & Gen_WcLocation & " )"
      Else
         tmp = tmp & " And (ML." & Gen_WcLocation & " Or ML.LocationID is Null) "
      End If
   End If
End If

'If CmbExchangeTypeID.Text <> "" Then
'    If Check3.Value = 0 Then
'       If mClientRNAVCode <> "" Then
'          tmp = tmp & " And (TI.NAVCode = '" & mClientNAVCode & "' or TI.NAVCode = '" & mClientRNAVCode & "')"
'       Else
'          tmp = tmp & " And TI.NAVCode = '" & mClientNAVCode & "'"
'       End If
'    Else
'       tmp = tmp & " And TI.NAVCode='" & mDPNAVCode & "'"
'    End If
'End If


If CmbExchangeTypeID.Text <> "" Then
    If Check3.Value = 0 Then
       If mClientRNAVCode <> "" Then
          'tmp = tmp & " And (TI.NAVCode = '" & mClientNAVCode & "' or TI.NAVCode = '" & mClientRNAVCode & "')"
          tmp = tmp & " And (TI.NAVCode in (Select distinct GSTNavCode From Mst_ClientStateWiseGST  Where ClientID = " & mClientIDRow & " ) Or TI.NAVCode in (Select distinct RGSTNavCode From Mst_ClientStateWiseGST  Where ClientID = " & mClientIDRow & " ) Or TI.NAVCode in (Select distinct NAVCode From Mst_Client  Where ClientID = " & mClientIDRow & " ) Or TI.NAVCode in (Select distinct RNAVCode From Mst_Client  Where ClientID = " & mClientIDRow & " ))"
       Else
          'tmp = tmp & " And TI.NAVCode = '" & mClientNAVCode & "'"
          tmp = tmp & " And (TI.NAVCode in (Select distinct GSTNavCode From Mst_ClientStateWiseGST  Where ClientID = " & mClientIDRow & " ) Or TI.NAVCode in (Select distinct NAVCode From Mst_Client  Where ClientID = " & mClientIDRow & " ))"
       End If
    Else
       'tmp = tmp & " And TI.NAVCode='" & mDPNAVCode & "'"
       tmp = tmp & " And (TI.NAVCode in (Select distinct GSTNavCode From Mst_DepositoryStateWiseGST  Where DepositoryID = " & mDPID & " ) Or TI.NAVCode in (Select distinct NAVCode From Mst_Depository  Where DepositoryID = " & mDPID & " ))"
    End If
End If


If CmbSearchBy.Text <> "" Then
   If Left(LCase(CmbSearchBy.Text), 1) = "t" Then
      tmp = tmp & " And TINo like '%" & TxtSearchBy & "%'"
   ElseIf Left(LCase(CmbSearchBy.Text), 1) = "p" Then
      tmp = tmp & " And PostInvoiceNo like '%" & TxtSearchBy & "%'"
   ElseIf Left(LCase(CmbSearchBy.Text), 1) = "e" Then
      tmp = tmp & " And EntryNo like '%" & TxtSearchBy & "%'"
   End If
Else
   If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "fagadmin" Then
      If CmbExchangeTypeID.Text = "" And CmbRegion.Text = "" Then
         MsgBox "Select either Region or Client!!"
         Exit Sub
      End If
   End If
End If

''If CmbExchangeTypeID.Text <> "" Then
''   If Check3.Value = 0 Then
''      tmp = tmp & " And TI.ClientIDRow = " & mClientIDRow & ""
''   Else
''      tmp = tmp & " And TI.DepositoryID ='" & TxtCDPID & "'"
''   End If
''End If

tmp = tmp & " Order By TIDate "

Call Tmp3Table

'Call Grid_Head


'tmp = "Select top(1) MC.DepositoryID, MC.DepositoryName, ME.ExchangeType, MC.NAVCode  from Mst_Depository MC"
'tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID = ME.ExchangeTypeID"
'tmp = tmp & " Where MC.NAVCode='" & TmpRs3!NAVCode & "'"
''-- Ullhas Patil 18 Sep 2017 Query change Due to GSTNAV Code
'''tmp = "Select MC.DepositoryID, MC.DepositoryName, ME.ExchangeType, MC.NAVCode  from Mst_Depository MC"
'''tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID = ME.ExchangeTypeID"
tmp = " Select Distinct MD.DepositoryID,MD.DepositoryName,ME.ExchangeType,MDGST.GSTNavCode "
tmp = tmp & " From Mst_DepositoryStateWiseGST MDGST "
tmp = tmp & " Inner Join Mst_Depository MD On MDGST.DepositoryID = MD.DepositoryID And MDGST.ExchangeTypeID = MD.ExchangeTypeID "
tmp = tmp & " Inner join Mst_ExchangeType ME on MD.ExchangeTypeID = ME.ExchangeTypeID "
''----------
Call Tmp4Table

''-- Ullhas Patil 18 Sep 2017 Query change Due to GSTNAV Code
'''tmp = " Select MC.ClientID, MC.ClientName, ME.ExchangeType, MC.NAVCode,  MC.RNAVCode from Mst_Client MC"
'''tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID = ME.ExchangeTypeID"
tmp = " Select Distinct MC.ClientID,MC.ClientName,ME.ExchangeType,MCGST.GSTNavCode,MCGST.RGSTNavCode "
tmp = tmp & " From Mst_ClientStateWiseGST MCGST "
tmp = tmp & " Inner Join Mst_Client MC On MCGST.ClientID = MC.ClientID "
tmp = tmp & " Inner Join Mst_ExchangeType ME on MC.ExchangeTypeID = ME.ExchangeTypeID "
'tmp = tmp & " Where MC.NAVCode='" & TmpRs3!NAVCode & "' or MC.RNAVCode='" & TmpRs3!NAVCode & "'"
''----------
Call Tmp5Table

tmp = " Select ECM.LocationID, EMP2.EmployeeName+'~'+EMP2.EmployeeNo  as 'AMName' , EMP2.EmployeeEmailID as AMEmailID"
tmp = tmp & " , EMP3.EmployeeName +'~'+EMP3.EmployeeNo as 'RMName', EMP3.EmployeeEmailID As RMEmailID,EMP.EmployeeName +'~'+EMP.EmployeeNo as 'CMPMName', EMP.EmployeeEmailID As CMPMEmailID"
tmp = tmp & " From Mst_EMPCMPMapping ECM"
tmp = tmp & " inner Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID And EMP.Flag = 'Y'"
tmp = tmp & " inner Join Mst_Employee EMP2 On EMP.DirectReportingTo = EMP2.EmployeeID And EMP2.Flag = 'Y'"
tmp = tmp & " inner Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID And EMP3.Flag = 'Y'"
tmp = tmp & " Where EMP2.EmployeeName Is Not Null"
tmp = tmp & " Group By ECM.LocationID, EMP2.EmployeeName, EMP2.EmployeeNo,EMP2.EmployeeEmailID, EMP3.EmployeeEmailID"
tmp = tmp & " ,EMP3.EmployeeName ,EMP3.EmployeeNo, EMP.EmployeeName, EMP.EmployeeNo,EMP.EmployeeEmailID"

Call Tmp6Table

'Call Grid_Head1
If TmpRs3.RecordCount > 0 Then
   ProgressBar1.Max = TmpRs3.RecordCount + 3
   For gr = 1 To TmpRs3.RecordCount
       rn = MSHFlexGrid1.Rows - 1

'       MSHFlexGrid1.TextMatrix(rn, 0) = TmpRs3!TINo
'       MSHFlexGrid1.TextMatrix(rn, 1) = IIf(IsNull(TmpRs3!TIDate), "", Format(TmpRs3!TIDate, "dd-MMM-yyyy"))
'       MSHFlexGrid1.TextMatrix(rn, 2) = IIf(IsNull(TmpRs3!TIAmount), 0, TmpRs3!TIAmount)
'       MSHFlexGrid1.TextMatrix(rn, 3) = IIf(IsNull(TmpRs3!RecAmount), 0, TmpRs3!RecAmount)
'       MSHFlexGrid1.TextMatrix(rn, 4) = IIf(IsNull(TmpRs3!RecTDSAmount), 0, TmpRs3!RecTDSAmount)
'       MSHFlexGrid1.TextMatrix(rn, 5) = IIf(IsNull(TmpRs3!RecOtherDed), 0, TmpRs3!RecOtherDed)
'       MSHFlexGrid1.TextMatrix(rn, 6) = Format(Val(MSHFlexGrid1.TextMatrix(rn, 2)) - (Val(MSHFlexGrid1.TextMatrix(rn, 3)) + Val(MSHFlexGrid1.TextMatrix(rn, 4)) + Val(MSHFlexGrid1.TextMatrix(rn, 5))), "#####0.00")
'       MSHFlexGrid1.TextMatrix(rn, 7) = IIf(IsNull(TmpRs3!PostInvoiceNo), 0, TmpRs3!PostInvoiceNo)
'       MSHFlexGrid1.TextMatrix(rn, 8) = IIf(IsNull(TmpRs3!EntryNo), 0, TmpRs3!EntryNo)
'       MSHFlexGrid1.TextMatrix(rn, 9) = IIf(IsNull(TmpRs3!locationname), "", TmpRs3!locationname)
'       MSHFlexGrid1.TextMatrix(rn, 10) = IIf(IsNull(TmpRs3!stateName), "", TmpRs3!stateName)
'       MSHFlexGrid1.TextMatrix(rn, 13) = IIf(IsNull(TmpRs3!NAVCode), "", TmpRs3!NAVCode)
'       MSHFlexGrid1.TextMatrix(rn, 14) = TmpRs3!TIID
       MSHFlexGrid1.TextMatrix(rn, 0) = TmpRs3!TINo
       MSHFlexGrid1.TextMatrix(rn, 1) = IIf(IsNull(TmpRs3!PostInvoiceNo), 0, TmpRs3!PostInvoiceNo)
       MSHFlexGrid1.TextMatrix(rn, 2) = IIf(IsNull(TmpRs3!TIDate), "", Format(TmpRs3!TIDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(rn, 3) = IIf(IsNull(TmpRs3!Locationname), "", TmpRs3!Locationname)
       MSHFlexGrid1.TextMatrix(rn, 13) = IIf(IsNull(TmpRs3!LocationID), "", TmpRs3!LocationID)
       MSHFlexGrid1.TextMatrix(rn, 17) = IIf(IsNull(TmpRs3!StateName), "", TmpRs3!StateName)
       MSHFlexGrid1.TextMatrix(rn, 18) = IIf(IsNull(TmpRs3!Region), "", TmpRs3!Region)
       MSHFlexGrid1.TextMatrix(rn, 8) = "" 'OC / BD
       
       TmpRs6.MoveFirst
       TmpRs6.Filter = "LocationID = " & IIf(MSHFlexGrid1.TextMatrix(rn, 13) = "", 0, MSHFlexGrid1.TextMatrix(rn, 13))
       If TmpRs6.RecordCount > 0 Then
          For j = 1 To TmpRs6.RecordCount
              If MSHFlexGrid1.TextMatrix(rn, 8) = "" Then
                 MSHFlexGrid1.TextMatrix(rn, 8) = IIf(IsNull(TmpRs6!CMPMName), "", TmpRs6!CMPMName)
              Else
                 If InStr(1, MSHFlexGrid1.TextMatrix(rn, 8), IIf(IsNull(TmpRs6!CMPMName), "", TmpRs6!CMPMName)) <= 0 Then
                    MSHFlexGrid1.TextMatrix(rn, 8) = IIf(IsNull(TmpRs6!CMPMName), "", TmpRs6!CMPMName) & " , " & MSHFlexGrid1.TextMatrix(rn, 8)
                 End If
              End If
              
              If MSHFlexGrid1.TextMatrix(rn, 9) = "" Then
                 MSHFlexGrid1.TextMatrix(rn, 9) = IIf(IsNull(TmpRs6!AMname), "", TmpRs6!AMname)
              Else
                 If InStr(1, MSHFlexGrid1.TextMatrix(rn, 9), IIf(IsNull(TmpRs6!AMname), "", TmpRs6!AMname)) <= 0 Then
                    MSHFlexGrid1.TextMatrix(rn, 9) = IIf(IsNull(TmpRs6!AMname), "", TmpRs6!AMname) & " , " & MSHFlexGrid1.TextMatrix(rn, 9)
                 End If
              End If
              If MSHFlexGrid1.TextMatrix(rn, 10) = "" Then
                 MSHFlexGrid1.TextMatrix(rn, 10) = IIf(IsNull(TmpRs6!RMName), "", TmpRs6!RMName)
              Else
                 If InStr(1, MSHFlexGrid1.TextMatrix(rn, 10), IIf(IsNull(TmpRs6!RMName), "", TmpRs6!RMName)) <= 0 Then
                    MSHFlexGrid1.TextMatrix(rn, 10) = IIf(IsNull(TmpRs6!RMName), "", TmpRs6!RMName) & " , " & MSHFlexGrid1.TextMatrix(rn, 10)
                 End If
              End If
              
              If MSHFlexGrid1.TextMatrix(rn, 14) = "" Then
                 MSHFlexGrid1.TextMatrix(rn, 14) = IIf(IsNull(TmpRs6!AMEmailID), "", TmpRs6!AMEmailID)
              Else
                 If InStr(1, MSHFlexGrid1.TextMatrix(rn, 14), IIf(IsNull(TmpRs6!AMEmailID), "", TmpRs6!AMEmailID)) <= 0 Then
                    MSHFlexGrid1.TextMatrix(rn, 14) = IIf(IsNull(TmpRs6!AMEmailID), "", TmpRs6!AMEmailID) & " , " & MSHFlexGrid1.TextMatrix(rn, 14)
                 End If
              End If
              
              If MSHFlexGrid1.TextMatrix(rn, 15) = "" Then
                 MSHFlexGrid1.TextMatrix(rn, 15) = IIf(IsNull(TmpRs6!RMEmailID), "", TmpRs6!RMEmailID)
              Else
                 If InStr(1, MSHFlexGrid1.TextMatrix(rn, 15), IIf(IsNull(TmpRs6!RMEmailID), "", TmpRs6!RMEmailID)) <= 0 Then
                    MSHFlexGrid1.TextMatrix(rn, 15) = IIf(IsNull(TmpRs6!RMEmailID), "", TmpRs6!RMEmailID) & " , " & MSHFlexGrid1.TextMatrix(rn, 15)
                 End If
              End If
              
              If MSHFlexGrid1.TextMatrix(rn, 16) = "" Then
                 MSHFlexGrid1.TextMatrix(rn, 16) = IIf(IsNull(TmpRs6!CMPMEmailID), "", TmpRs6!CMPMEmailID)
              Else
                 If InStr(1, MSHFlexGrid1.TextMatrix(rn, 16), IIf(IsNull(TmpRs6!CMPMEmailID), "", TmpRs6!CMPMEmailID)) <= 0 Then
                    MSHFlexGrid1.TextMatrix(rn, 16) = IIf(IsNull(TmpRs6!CMPMEmailID), "", TmpRs6!CMPMEmailID) & " , " & MSHFlexGrid1.TextMatrix(rn, 16)
                 End If
              End If
              
              TmpRs6.MoveNext
          Next j
       End If
       TmpRs6.Filter = ""
       'Invoice Type
       If IsNull(TmpRs3!PIFlag) = True Then
          MSHFlexGrid1.TextMatrix(rn, 4) = ""
       ElseIf LCase(TmpRs3!PIFlag) = "reservation" Then
          MSHFlexGrid1.TextMatrix(rn, 4) = "Reservation"
       ElseIf LCase(TmpRs3!PIFlag) = "rj" Then
          MSHFlexGrid1.TextMatrix(rn, 4) = "Rejected"
       ElseIf LCase(TmpRs3!PIFlag) = "commtrack" Then
          MSHFlexGrid1.TextMatrix(rn, 4) = "Commtrack"
       ElseIf LCase(TmpRs3!PIFlag) = "nnnr" Then
          MSHFlexGrid1.TextMatrix(rn, 4) = "NNNR"
       ElseIf LCase(TmpRs3!PIFlag) = "dp" Then
          MSHFlexGrid1.TextMatrix(rn, 4) = "DP"
       ElseIf LCase(TmpRs3!PIFlag) = "sd" Then
          MSHFlexGrid1.TextMatrix(rn, 4) = "Seven Days"
       ElseIf LCase(TmpRs3!PIFlag) = "adhoc" Then
          MSHFlexGrid1.TextMatrix(rn, 4) = "Ad-Hoc"
       ElseIf LCase(TmpRs3!PIFlag) = "rrn" Then
          MSHFlexGrid1.TextMatrix(rn, 4) = "RRN Withdrawal"
       End If
       MSHFlexGrid1.TextMatrix(rn, 7) = Format(IIf(IsNull(TmpRs3!TIAmount), 0, TmpRs3!TIAmount) - (IIf(IsNull(TmpRs3!RecAmount), 0, TmpRs3!RecAmount) + IIf(IsNull(TmpRs3!RecTDSAmount), 0, TmpRs3!RecTDSAmount) + IIf(IsNull(TmpRs3!RecOtherDed), 0, TmpRs3!RecOtherDed)), "#####0.00") 'OutStanding Amount
       MSHFlexGrid1.TextMatrix(rn, 11) = IIf(IsNull(TmpRs3!NavCode), "", TmpRs3!NavCode)
       MSHFlexGrid1.TextMatrix(rn, 12) = TmpRs3!TIID

       If InStr(1, TmpRs3!NavCode, "CSWD") > 0 Then
          TmpRs4.MoveFirst
          TmpRs4.Find "GSTNavCode='" & TmpRs3!NavCode & "'"
          If Not TmpRs4.EOF Then
             MSHFlexGrid1.TextMatrix(rn, 5) = IIf(IsNull(TmpRs4!DepositoryID), "", TmpRs4!DepositoryID)
             MSHFlexGrid1.TextMatrix(rn, 6) = IIf(IsNull(TmpRs4!DepositoryName), "", TmpRs4!DepositoryName)
          End If
       Else
          'TmpRs5.Filter = "NavCode=" & TmpRs3!NAVCode & " And RNAVCode=" & TmpRs3!NAVCode
          TmpRs5.MoveFirst
          'RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
          TmpRs5.Find "GSTNavCode ='" & TmpRs3!NavCode & "'"
          If Not TmpRs5.EOF Then
             MSHFlexGrid1.TextMatrix(rn, 5) = IIf(IsNull(TmpRs5!ClientID), "", TmpRs5!ClientID)
             MSHFlexGrid1.TextMatrix(rn, 6) = IIf(IsNull(TmpRs5!ClientName), "", TmpRs5!ClientName)
          Else
             TmpRs5.MoveFirst
             TmpRs5.Find "RGSTNavCode='" & TmpRs3!NavCode & "'"
             If Not TmpRs5.EOF Then
                MSHFlexGrid1.TextMatrix(rn, 5) = IIf(IsNull(TmpRs5!ClientID), "", TmpRs5!ClientID)
                MSHFlexGrid1.TextMatrix(rn, 6) = IIf(IsNull(TmpRs5!ClientName), "", TmpRs5!ClientName)
             End If
          End If
       End If
       
       If Trim(MSHFlexGrid1.TextMatrix(rn, 6)) = "" Then
        GetClientName (rn)
       End If
       
       MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(rn, 7))
       TmpRs3.MoveNext
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   
   ProgressBar1.Value = ProgressBar1.Max
   CmdExcel.Enabled = True
'   If Check1.Value = 1 Then
'      Call FillClientDetails
'   End If
   CmdMail.Enabled = True
Else
   CmdExcel.Enabled = False
   CmdMail.Enabled = False
End If
Label2.Refresh
Label2.Caption = TmpRs3.RecordCount
MsgBox ("Done!!!")
End Sub

Private Sub GetClientName(x_ As Variant)
tmp = " Select ClientID,ClientName From Mst_Client where NAVCode = '" & MSHFlexGrid1.TextMatrix(x_, 11) & "' Or RNAVCode = '" & MSHFlexGrid1.TextMatrix(x_, 11) & "'"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
    MSHFlexGrid1.TextMatrix(x_, 5) = IIf(IsNull(TmpRs1!ClientID), "", TmpRs1!ClientID)
    MSHFlexGrid1.TextMatrix(x_, 6) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
Else
    tmp = " Select DepositoryID,DepositoryName From Mst_Depository where NAVCode = '" & MSHFlexGrid1.TextMatrix(x_, 11) & "' "
    If TmpRs1.RecordCount > 0 Then
        MSHFlexGrid1.TextMatrix(x_, 5) = IIf(IsNull(TmpRs1!DepositoryID), "", TmpRs1!DepositoryID)
        MSHFlexGrid1.TextMatrix(x_, 6) = IIf(IsNull(TmpRs1!DepositoryName), "", TmpRs1!DepositoryName)
    End If
End If
End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "FrmRpt_OutStWithClientInvoiceDetails.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "FrmRpt_OutStWithClientInvoiceDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Invoice Details"
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 3
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    'For c = 0 To MSHFlexGrid1.Cols - 2
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
ProgressBar1.Value = 0
'Set oWS = oWB.Worksheets("Sheet2")
'oWS.Name = "Client Details"
'oWS.ClearArrows
'mRow = 1
'ProgressBar1.Max = MSHFlexGrid2.Rows + 3
'For i = 0 To MSHFlexGrid2.Rows - 1
'    mRow = i + 1 ' + 6
'    For c = 0 To MSHFlexGrid2.Cols - 2
'        If InStr(MSHFlexGrid2.TextMatrix(i, c), "/") > 0 Then
'           If IsDate(MSHFlexGrid2.TextMatrix(i, c)) Then
'              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid2.TextMatrix(i, c), "MM/dd/yyyy")
'           Else
'              oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(i, c)
'           End If
'        Else
'           oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(i, c)
'        End If
'    Next
'    ProgressBar1.Value = ProgressBar1.Value + 1
'Next
'
'ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub

Private Sub CmdMail_Click()
mFlag = True
mAM = ""
mRow = 0
j = 1
Call Grid_Head1
For I = 2 To MSHFlexGrid1.Rows - 1
    mAM = ""
    mFlag = True
    If MSHFlexGrid1.TextMatrix(I, 8) <> "" Then
       mAM = MSHFlexGrid1.TextMatrix(I, 8)
       For j = 1 To MSHFlexGrid2.Rows - 1
           If MSHFlexGrid2.TextMatrix(j, 0) = mAM Then
              mFlag = False
              Exit For
           End If
       Next
       If mFlag = True Then
          mRow = MSHFlexGrid2.Rows - 1
          MSHFlexGrid2.TextMatrix(mRow, 0) = mAM
          MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
        End If
    End If
Next

Call Gen_Mail
MsgBox "Done!!"

End Sub
Private Function Gen_Mail()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg, mMsgSubject, mGridMsgBlank, mGridMsgDP As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String
Dim AM_ As String
Dim mFlag As Boolean
mMsgSubject = "Invoices outstanding with clients in your location"

mMsg = ""
'mMsg = "<tr>BOD Update Details for " & CmbLocationID.Text & "</tr>"
'mMsg = mMsg & "<tr><td> Location </td><td> " & CmbLocationID.Text & "</td></tr>"
AM_ = ""
mFlag = True
mGridMsgBlank = ""
mGridMsgDP = ""
mTotalBlank = 0
mTotalDP = 0
mTotalOther = 0
For I = 1 To MSHFlexGrid2.Rows - 1
    AM_ = ""
    mGridMsg = ""
    TOADD = ""
    TOCC = ""
    mEmailId = ""
    AM_ = MSHFlexGrid2.TextMatrix(I, 0)
    mTotalOther = 0
    For j = 2 To MSHFlexGrid1.Rows - 2
        If MSHFlexGrid1.TextMatrix(j, 8) <> "" Then
           If MSHFlexGrid1.TextMatrix(j, 8) = AM_ Then
              If LCase(MSHFlexGrid1.TextMatrix(j, 4)) <> "dp" And LCase(MSHFlexGrid1.TextMatrix(j, 4)) <> "commtrack" Then
                 mGridMsg = mGridMsg & "<tr><td> " & MSHFlexGrid1.TextMatrix(j, 0) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 1) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 2) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 3) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 4) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 5) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 6) & " </td><td align=right> " & MSHFlexGrid1.TextMatrix(j, 7) & " </td></tr>"
                 
                 If InStr(1, TOADD, MSHFlexGrid1.TextMatrix(j, 16)) = 0 Then
                    If TOADD = "" Then
                       TOADD = TOADD & MSHFlexGrid1.TextMatrix(j, 16)
                    Else
                       TOADD = TOADD & "," & MSHFlexGrid1.TextMatrix(j, 16)
                    End If
                 End If
                 
'Commented on 18 April 2015 On request of Sameer Sir to Mayur Sir for AM Name to be displayed in TO Email ID list and not in CC
                 
'                 If InStr(1, TOCC, MSHFlexGrid1.TextMatrix(j, 14)) = 0 Then
'                    If TOCC = "" Then
'                       TOCC = TOCC & MSHFlexGrid1.TextMatrix(j, 14)
'                    Else
'                       TOCC = TOCC & "," & MSHFlexGrid1.TextMatrix(j, 14)
'                    End If
'                 End If
                 If InStr(1, TOADD, MSHFlexGrid1.TextMatrix(j, 14)) = 0 Then
                    If TOADD = "" Then
                       TOADD = TOADD & MSHFlexGrid1.TextMatrix(j, 14)
                    Else
                       TOADD = TOADD & "," & MSHFlexGrid1.TextMatrix(j, 14)
                    End If
                 End If
                 
                 
                 If InStr(1, TOCC, MSHFlexGrid1.TextMatrix(j, 15)) = 0 Then
                    If TOCC = "" Then
                       TOCC = TOCC & MSHFlexGrid1.TextMatrix(j, 15)
                    Else
                       TOCC = TOCC & "," & MSHFlexGrid1.TextMatrix(j, 15)
                    End If
                 End If
                 mTotalOther = mTotalOther + Val(MSHFlexGrid1.TextMatrix(j, 7))
              Else
                 'mGridMsgDP = mGridMsgDP & "<tr><td> " & MSHFlexGrid1.TextMatrix(j, 0) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 1) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 2) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 3) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 4) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 5) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 6) & " </td></tr>"
                 mGridMsgDP = mGridMsgDP & "<tr><td> " & MSHFlexGrid1.TextMatrix(j, 0) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 1) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 2) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 3) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 4) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 5) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 6) & " </td><td align=right> " & MSHFlexGrid1.TextMatrix(j, 7) & " </td></tr>"
                 mTotalDP = mTotalDP + Val(MSHFlexGrid1.TextMatrix(j, 7))
              End If
           End If
        Else
           If LCase(MSHFlexGrid1.TextMatrix(j, 4)) <> "dp" And LCase(MSHFlexGrid1.TextMatrix(j, 4)) <> "commtrack" Then
              mGridMsgBlank = mGridMsgBlank & "<tr><td> " & MSHFlexGrid1.TextMatrix(j, 0) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 1) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 2) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 3) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 4) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 5) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 6) & " </td><td align=right> " & MSHFlexGrid1.TextMatrix(j, 7) & " </td></tr>"
              mTotalBlank = mTotalBlank + Val(MSHFlexGrid1.TextMatrix(j, 7))
           Else
              mGridMsgDP = mGridMsgDP & "<tr><td> " & MSHFlexGrid1.TextMatrix(j, 0) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 1) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 2) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 3) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 4) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 5) & " </td><td> " & MSHFlexGrid1.TextMatrix(j, 6) & " </td><td align=right> " & MSHFlexGrid1.TextMatrix(j, 7) & " </td></tr>"
              mTotalDP = mTotalDP + Val(MSHFlexGrid1.TextMatrix(j, 7))
           End If
        End If
    Next
    If TOCC <> "" Then
       TOCC = TOCC & ",kuldip.s@ncml.com,basant.a@ncml.com,prakash.je@ncml.com,dipesh.s@ncml.com,sameer.v@ncml.com,ashok.d@ncml.com"
    Else
       TOCC = ",kuldip.s@ncml.com,basant.a@ncml.com,prakash.je@ncml.com,dipesh.s@ncml.com,sameer.v@ncml.com,ashok.d@ncml.com"
    End If
    If mGridMsg <> "" Then
       mMsg = "<tr><td colspan=7> Total Outstanding  Amount</td><td align=right>" & Format(mTotalOther, "#####0.00") & " </td></tr>"
       mMsg = mMsg & "<tr><td> TI No (Voucher No) </td><td> Post Invoice No </td><td>TI Date</td><td>Location</td><td>Invoice Type</td><td>Client ID</td><td>Client Name</td><td align=right>OutStanding Amount</td></tr>"
       
       mMsg = mMsg & mGridMsg
       TOCC = TOCC & "," & Gen_UserEmailID
       Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)
    End If
Next
If mGridMsgBlank <> "" Then
   mMsg = "<tr><td colspan=7> Total Outstanding  Amount</td><td align=right >" & Format(mTotalBlank, "#####0.00") & " </td></tr>"
   mMsg = mMsg & "<tr><td> TI No (Voucher No) </td><td> Post Invoice No </td><td>TI Date</td><td>Location</td><td>Invoice Type</td><td>Client ID</td><td>Client Name</td><td align=right>OutStanding Amount</td></tr>"
   mMsg = mMsg & mGridMsgBlank
   TOADD = "kuldip.s@ncml.com"
   TOCC = "kuldip.s@ncml.com,basant.a@ncml.com,prakash.je@ncml.com,dipesh.s@ncml.com,sameer.v@ncml.com,ashok.d@ncml.com," & Gen_UserEmailID
   Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)
End If
If mGridMsgDP <> "" Then
   mMsg = "<tr><td colspan=7> Total Outstanding  Amount</td><td  align=right>" & Format(mTotalDP, "#####0.00") & " </td></tr>"
   mMsg = mMsg & "<tr><td> TI No (Voucher No) </td><td> Post Invoice No </td><td>TI Date</td><td>Location</td><td>Invoice Type</td><td>Client ID</td><td>Client Name</td><td align=right>OutStanding Amount</td></tr>"
   mMsg = mMsg & mGridMsgDP

   TOADD = "kuldip.s@ncml.com"
   TOCC = "basant.a@ncml.com,prakash.je@ncml.com,dipesh.s@ncml.com,sameer.v@ncml.com,ashok.d@ncml.com," & Gen_UserEmailID
   Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)
End If

End Function
Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = Blue>The following invoices in the name of the clients in your location are outstanding and yet to be recovered. Please initiate steps immediately to collect these outstanding invoices from the clients concerned immediately and report the status to your reporting authority and also to CFO on priority basis.<br><br><table border = 1>"
htmlStringtail = "</table></Font></body></html>"
   
'With oSMTP
'  .Server = servertxt
'  .MailFrom = Gen_UserEmailID 'TOCC_ '"mayur.d@ncml.com"
'  .SendTo = ToADD_
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
'  .BCC = "parag.r@ncml.com"
'  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
objSMTP.FromAddr = Gen_UserEmailID
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If

objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If

End Function

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True

Call CmdClear_Click
CmdExcel.Enabled = False
CmdMail.Enabled = False
End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.Cols = 19 '17
MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(1, 0) = "TI No (Voucher No)"
MSHFlexGrid1.TextMatrix(1, 1) = "Post Invoice No"
MSHFlexGrid1.TextMatrix(1, 2) = "TI Date"
MSHFlexGrid1.TextMatrix(1, 3) = "Location"
MSHFlexGrid1.TextMatrix(1, 4) = "Invoice Type"
MSHFlexGrid1.TextMatrix(1, 5) = "Client ID"
MSHFlexGrid1.TextMatrix(1, 6) = "Client Name"
MSHFlexGrid1.TextMatrix(1, 7) = "OutStanding Amount"
MSHFlexGrid1.TextMatrix(1, 8) = "CMPM"
MSHFlexGrid1.TextMatrix(1, 9) = "AE / AM"
MSHFlexGrid1.TextMatrix(1, 10) = "RM"
MSHFlexGrid1.TextMatrix(1, 11) = "NAVCode"
MSHFlexGrid1.TextMatrix(1, 12) = "TIID"
MSHFlexGrid1.TextMatrix(1, 13) = "LocationID"
MSHFlexGrid1.TextMatrix(1, 14) = "AM Email ID"
MSHFlexGrid1.TextMatrix(1, 15) = "RM Email ID"
MSHFlexGrid1.TextMatrix(1, 16) = "CMPM Email ID"
MSHFlexGrid1.TextMatrix(1, 17) = "State Name"
MSHFlexGrid1.TextMatrix(1, 18) = "Region"

MSHFlexGrid1.ColWidth(0) = 1700
MSHFlexGrid1.ColWidth(1) = 1700
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 1800
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 2500
MSHFlexGrid1.ColWidth(7) = 1500
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(10) = 1400
MSHFlexGrid1.ColWidth(11) = 1800
MSHFlexGrid1.ColWidth(12) = 1000
MSHFlexGrid1.ColWidth(13) = 1500
MSHFlexGrid1.ColWidth(14) = 1500
MSHFlexGrid1.ColWidth(15) = 1500
MSHFlexGrid1.ColWidth(16) = 1500

MSHFlexGrid1.RowHeight(1) = 800

End Sub

Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   mExchangeTypeID = 0
   CmbCNonNcdexClient.Clear
   CmbClientDP.Clear
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If

If mExchangeTypeID <> RsMst_ExchangeType!ExchangeTypeID Then
   CmbCNonNcdexClient.Clear
   CmbClientDP.Clear
   CmbDPName.Clear
   Call Grid_Head
   'Call Grid_Head1
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

If mExchangeTypeID = 1 Then
   CmbCNonNcdexClient.Width = 9400
   CmbCNonNcdexClient.Top = TxtClientID.Top
   CmbCNonNcdexClient.Left = TxtClientID.Left
   CmbCNonNcdexClient.Visible = True
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   Check3.Value = 0
   Check3.Enabled = False
Else
   CmbCNonNcdexClient.Width = 9400
   CmbCNonNcdexClient.Top = TxtClientID.Top
   CmbCNonNcdexClient.Left = TxtClientID.Left
   CmbCNonNcdexClient.Visible = True
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   
''   Label25.Visible = True
''   Label26.Visible = True
''   Label28.Visible = True
''
''   CmbCNonNcdexClient.Visible = False
   Check3.Enabled = True
   Call Check3_Click
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   CmbDPName.Visible = True
   CmbDPName.Width = 9400 'CmbNonNcdexClient.Width
   CmbDPName.Top = TxtClientID.Top 'CmbNonNcdexClient.Top
   CmbDPName.Left = TxtClientID.Left 'CmbNonNcdexClient.Left
   CmbCNonNcdexClient.Visible = False
   CmbDPName.Text = ""
Else
   Label25.Visible = False
   Label26.Visible = False
   Label28.Visible = False
   CmbCNonNcdexClient.Visible = True
   CmbDPName.Visible = False
   CmbCNonNcdexClient.Text = ""
End If

End Sub



Private Sub TxtClientID_LostFocus()
CmbClientDP.Visible = False
If TxtClientID = "" Then
   TxtCClientName = ""
   TxtCDPID = ""
   TxtCDPName = ""
   Exit Sub
   
End If
If CmbExchangeTypeID.Text = "" Then Exit Sub

Dim ans As Variant
ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)

If ans(0, 0) = False Then
   MsgBox "Not Found !!!"
   TxtClientID.SetFocus
   Exit Sub
End If
If ans(3, 0) > 1 Then
   CmbClientDP.Visible = True
   CmbClientDP.SetFocus
End If
CmbClientDP.Text = ""
CmbClientDP.Clear
tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      "Mst_Client.ClientName,Mst_Client.ClientIDRow,Mst_Client.NAVCode,Mst_Client.RNAVCode " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"

Call Tmp1Table

For op = 1 To TmpRs1.RecordCount
    If CmbClientDP.Text = "" Then
       CmbClientDP.Text = TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    End If
    CmbClientDP.AddItem TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    TmpRs1.MoveNext
Next
If ans(3, 0) <= 1 Then
   CmbClientDP_LostFocus
End If
End Sub


Private Sub CmbCNonNcdexClient_GotFocus()

tmp1 = CmbCNonNcdexClient.Text
'Call TableMst_Client("where ExchangeTypeId = " & mExchangeTypeID)

tmp = "Select MC.ClientName,MC.ClientID,MD.DepositoryName,MC.ClientIDRow,MC.NAVCode,MC.RNAVCode From Mst_Client MC"
tmp = tmp & " inner join Mst_Depository MD on MD.DepositoryID=MC.DepositoryID"
tmp = tmp & " Where MC.ExchangeTypeId=" & mExchangeTypeID

Call TmpTable

CmbCNonNcdexClient.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbCNonNcdexClient.AddItem TmpRs!ClientName & "~" & TmpRs!ClientIDRow & "~" & TmpRs!DepositoryName & "~" & TmpRs!ClientID
    TmpRs.MoveNext
Next
CmbCNonNcdexClient.Text = tmp1

End Sub

Private Sub CmbCNonNcdexClient_LostFocus()
If CmbCNonNcdexClient.Text = "" Then
   Exit Sub
End If

strSplitString = Split(CmbCNonNcdexClient, "~")

If UBound(strSplitString) <> 3 Then
   MsgBox "Invalid Selection!!!"
   CmbCNonNcdexClient.SetFocus
   Exit Sub
End If

ClientName_ = Trim(strSplitString(0))
ClientNo_ = Trim(strSplitString(1))
DPName_ = Trim(strSplitString(2))
ClientIDRow_ = Trim(strSplitString(3))

TmpRs.MoveFirst
TmpRs.Find "ClientID = " & ClientIDRow_ & ""

If TmpRs.EOF Then
   MsgBox "Invalid Selection "
   CmbCNonNcdexClient.SetFocus
   Exit Sub
End If
  
mClientIDRow = TmpRs!ClientID
mClientNAVCode = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
mClientRNAVCode = IIf(IsNull(TmpRs!RNAVCode), "", TmpRs!RNAVCode)
End Sub
Private Sub CmbClientDP_LostFocus()
Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   Else
      CmbClientDP.SetFocus
   End If
   Exit Sub
End If
TmpRs1.MoveFirst


TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtCClientName = TmpRs1!ClientName
TxtCDPID = TmpRs1!DepositoryID
TxtCDPName = TmpRs1!DepositoryName
mClientNAVCode = IIf(IsNull(TmpRs1!NavCode), "", TmpRs1!NavCode)
mClientRNAVCode = IIf(IsNull(TmpRs1!RNAVCode), "", TmpRs1!RNAVCode)
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientID

End Sub

Private Sub CmbDPName_GotFocus()
If CmbExchangeTypeID.Text = "" Then
   MsgBox "Please select Exchange Type!!!"
   CmbDPName.Clear
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
tmp = CmbDPName.Text

Call TableMst_Depository("Where ExchangeTypeId = " & mExchangeTypeID & " ")

CmbDPName.Clear
If RsMst_Depository.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Depository.RecordCount
    CmbDPName.AddItem RsMst_Depository!DepositoryName & "~" & RsMst_Depository!DepositoryID
    RsMst_Depository.MoveNext
Next
CmbDPName.Text = tmp

End Sub

Private Sub CmbDPName_LostFocus()

Dim strSplitString() As String
Dim DPID_, DPName_ As String

If CmbDPName.Text = "" Then
   TxtCDPID = ""
   TxtCDPName = ""
   Exit Sub
End If
strSplitString = Split(CmbDPName, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection!!!"
   CmbDPName.SetFocus
   Exit Sub
End If


DPName_ = Trim(strSplitString(0))
DPID_ = Trim(strSplitString(1))

RsMst_Depository.MoveFirst
RsMst_Depository.Find "DepositoryID = '" & DPID_ & "'"
If RsMst_Depository.EOF Then
   MsgBox "Invalid selection "
   CmbDPName.SetFocus
   Exit Sub
End If

mDPID = RsMst_Depository!DepositoryID
TxtCDPID = mDPID
TxtCDPName = DPName_ 'CmbDPName.Text
mDPNAVCode = IIf(IsNull(RsMst_Depository!NavCode), "", RsMst_Depository!NavCode)
End Sub

Private Sub Grid_Head1()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 1
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.WordWrap = True

MSHFlexGrid2.TextMatrix(0, 0) = "AMName"

MSHFlexGrid2.ColWidth(0) = 1800

End Sub

Private Sub FillClientDetails()
Call Grid_Head1
If CmbExchangeTypeID.Text <> "" Then
    If Check3.Value = 1 Then
          tmp = "Select MC.DepositoryID,MC.DepositoryName,ME.ExchangeType,MC.NAVCode from Mst_Depository MC"
          tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID=ME.ExchangeTypeID"
          tmp = tmp & " Where MC.NAVCode='" & mDPNAVCode & "'"
          Call Tmp2Table
          For I = 1 To TmpRs2.RecordCount
             rn = MSHFlexGrid2.Rows - 1
             MSHFlexGrid2.TextMatrix(rn, 0) = IIf(IsNull(TmpRs2!DepositoryID), "", TmpRs2!DepositoryID)
             MSHFlexGrid2.TextMatrix(rn, 1) = IIf(IsNull(TmpRs2!DepositoryName), "", TmpRs2!DepositoryName)
             MSHFlexGrid2.TextMatrix(rn, 2) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
             MSHFlexGrid2.TextMatrix(rn, 3) = IIf(IsNull(TmpRs2!NavCode), "", TmpRs2!NavCode)
             MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
             TmpRs2.MoveNext
          Next
          
       
    Else
          tmp = " Select MC.ClientID,MC.ClientName,ME.ExchangeType,MC.NAVCode,MC.RNAVCode from Mst_Client MC"
          tmp = tmp & " Inner join Mst_ExchangeType ME on MC.ExchangeTypeID=ME.ExchangeTypeID"
          If mClientRNAVCode <> "" Then
             tmp = tmp & " Where MC.NAVCode='" & mClientNAVCode & "' or MC.RNAVCode='" & mClientRNAVCode & "'"
          Else
             tmp = tmp & " Where MC.NAVCode='" & mClientNAVCode & "'"
          End If

          Call Tmp2Table
          For I = 1 To TmpRs2.RecordCount
              rn = MSHFlexGrid2.Rows - 1
              MSHFlexGrid2.TextMatrix(rn, 0) = IIf(IsNull(TmpRs2!ClientID), "", TmpRs2!ClientID)
              MSHFlexGrid2.TextMatrix(rn, 1) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
              MSHFlexGrid2.TextMatrix(rn, 2) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
              MSHFlexGrid2.TextMatrix(rn, 3) = IIf(IsNull(TmpRs2!NavCode), "", TmpRs2!NavCode)
              MSHFlexGrid2.TextMatrix(rn, 4) = IIf(IsNull(TmpRs2!RNAVCode), "", TmpRs2!RNAVCode)
              MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
              TmpRs2.MoveNext
          Next
    End If
End If
        
          
End Sub
