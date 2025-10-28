VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_GSLRectification 
   Caption         =   "GSL Rectification"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   8895
      Left            =   14970
      TabIndex        =   12
      Top             =   255
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   135
      TabIndex        =   0
      Top             =   9780
      Width           =   10935
   End
   Begin VB.PictureBox Frame4 
      Height          =   9540
      Left            =   0
      ScaleHeight     =   9480
      ScaleWidth      =   15135
      TabIndex        =   13
      Top             =   0
      Width           =   15195
      Begin VB.OptionButton Option2 
         Caption         =   "Update Negative GSL"
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
         Left            =   2880
         TabIndex        =   28
         Top             =   120
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "GSL vs WHR Monitoring"
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
         TabIndex        =   27
         Top             =   120
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.Frame Frame11 
         Height          =   1665
         Left            =   45
         TabIndex        =   14
         Top             =   450
         Width           =   14880
         Begin VB.CommandButton CmdSave 
            Caption         =   "Save"
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
            Left            =   6015
            TabIndex        =   7
            Top             =   1260
            Width           =   1095
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
            Height          =   315
            Left            =   7110
            TabIndex        =   8
            Top             =   1260
            Width           =   1095
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
            Height          =   315
            Left            =   3705
            TabIndex        =   5
            Top             =   1260
            Width           =   1215
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
            Height          =   315
            Left            =   4920
            TabIndex        =   6
            Top             =   1260
            Width           =   1095
         End
         Begin VB.ComboBox CmbFlag 
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
            ItemData        =   "Frm_GSLRectification.frx":0000
            Left            =   6240
            List            =   "Frm_GSLRectification.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   420
            Width           =   1965
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
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   420
            Width           =   2325
         End
         Begin VB.ComboBox CmbClientName 
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
            ItemData        =   "Frm_GSLRectification.frx":0026
            Left            =   1680
            List            =   "Frm_GSLRectification.frx":002D
            Sorted          =   -1  'True
            TabIndex        =   4
            Top             =   840
            Width           =   6525
         End
         Begin VB.ComboBox CmbCommodity 
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
            Left            =   2520
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   420
            Width           =   3645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Difference  between GSL WHR and WHR Bags is "
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
            Left            =   8400
            TabIndex        =   26
            Top             =   909
            Width           =   4440
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Difference  between Deposited and WHR Qty is "
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
            Left            =   8400
            TabIndex        =   25
            Top             =   522
            Width           =   4275
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Difference  between Deposited and WHR Bags is "
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
            Left            =   8400
            TabIndex        =   24
            Top             =   135
            Width           =   4455
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Difference  between GSL WHR and WHR Qty is "
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
            Left            =   8400
            TabIndex        =   23
            Top             =   1297
            Width           =   4260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Flag"
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
            Left            =   7020
            TabIndex        =   21
            Top             =   135
            Width           =   405
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Location Name"
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
            Left            =   495
            TabIndex        =   20
            Top             =   135
            Width           =   1575
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Client Name"
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
            Left            =   120
            TabIndex        =   19
            Top             =   900
            Width           =   1275
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Commodity"
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
            Left            =   3757
            TabIndex        =   18
            Top             =   135
            Width           =   1170
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "-"
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
            Left            =   2610
            TabIndex        =   16
            Top             =   1290
            Width           =   90
         End
         Begin VB.Label Label5 
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
            Left            =   120
            TabIndex        =   15
            Top             =   1290
            Width           =   2280
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   45
         TabIndex        =   17
         Top             =   9150
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6945
         Left            =   45
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2160
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   12250
         _Version        =   393216
         Tab             =   2
         TabHeight       =   520
         TabCaption(0)   =   "GSL Detail"
         TabPicture(0)   =   "Frm_GSLRectification.frx":003B
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "TxtEnterData"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "MSHFlexGrid3"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "MSHFlexGrid1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "WHR Detail"
         TabPicture(1)   =   "Frm_GSLRectification.frx":0057
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSHFlexGrid2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Godown Summary"
         TabPicture(2)   =   "Frm_GSLRectification.frx":0073
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "MSHFlexGrid4"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.TextBox TxtEnterData 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   -73200
            TabIndex        =   9
            Top             =   1320
            Visible         =   0   'False
            Width           =   1800
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   825
            Left            =   -74880
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   6000
            Visible         =   0   'False
            Width           =   14640
            _ExtentX        =   25823
            _ExtentY        =   1455
            _Version        =   393216
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   6465
            Left            =   -74880
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   360
            Width           =   14640
            _ExtentX        =   25823
            _ExtentY        =   11404
            _Version        =   393216
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   6465
            Left            =   -74880
            TabIndex        =   10
            Top             =   360
            Width           =   14640
            _ExtentX        =   25823
            _ExtentY        =   11404
            _Version        =   393216
            FixedCols       =   0
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
            Height          =   6465
            Left            =   120
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   360
            Width           =   14640
            _ExtentX        =   25823
            _ExtentY        =   11404
            _Version        =   393216
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
   End
End
Attribute VB_Name = "Frm_GSLRectification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCommodityID, mClientID, mSM_Prefix As Variant
Dim mColsel, mRowsel As Single
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub CmbClientName_GotFocus()
If Trim(CmbLocation.Text) = "" Then
   MsgBox "Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
If Trim(CmbCommodity.Text) = "" Then
   MsgBox "Select Commodity!!!"
   CmbCommodity.SetFocus
   Exit Sub
End If
tmp1 = CmbClientName.Text
tmp = "Select ClientName+'~'+convert(varchar, ClientID)+'~'+convert(varchar, ClientIDRow) as Client, ClientIDRow"
tmp = tmp & " from Mst_Client where exchangetypeid = 1 and ClientIDRow in ( Select Distinct ClientIDRow from Mst_GSL Where SM_Prefix = '" & mSM_Prefix & "' and CommodityID = " & mCommodityID & ")"
Call TmpTable
CmbClientName.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Record found!!! "
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
   CmbClientName.AddItem TmpRs!Client
   TmpRs.MoveNext
Next
CmbClientName.Text = tmp1
End Sub

Private Sub CmbClientName_LostFocus()
If Trim(CmbClientName.Text) = "" Then
   Call Grid_Head
   Exit Sub
End If
If tmp1 <> CmbClientName.Text Then
   Call Grid_Head
End If

TmpRs.MoveFirst
TmpRs.Find "Client = '" & CmbClientName.Text & "'"
If TmpRs.EOF Then
   MsgBox "Invalid selection!!!"
   CmbClientName.SetFocus
   Exit Sub
End If
mClientID = TmpRs!ClientIDRow
End Sub

Private Sub CmbCommodity_GotFocus()
If Trim(CmbLocation.Text) = "" Then
   MsgBox "Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
tmp = CmbCommodity.Text
Call TableMst_Commodity(" Where CommodityID in (Select distinct CommodityID from Mst_GSL Where SM_Prefix = '" & mSM_Prefix & "')")
CmbCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
   CmbCommodity.AddItem RsMst_Commodity!Commodity
   RsMst_Commodity.MoveNext
Next
CmbCommodity.Text = tmp
End Sub

Private Sub CmbCommodity_LostFocus()
If Trim(CmbCommodity.Text) = "" Then
   Call Grid_Head
   CmbClientName.Text = ""
   Exit Sub
End If
If tmp <> CmbCommodity.Text Then
   Call Grid_Head
   CmbClientName.Text = ""
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection!!!"
   CmbCommodity.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity!CommodityID
End Sub

Private Sub CmbFlag_GotFocus()
tmp = CmbFlag.Text
End Sub

Private Sub CmbFlag_LostFocus()
If tmp <> CmbFlag.Text Then
   Call Grid_Head
End If
If Trim(CmbFlag.Text) = "" Then Exit Sub
If LCase(CmbFlag.Text) <> "non transfer" And LCase(CmbFlag.Text) <> "transfer" Then
   MsgBox "Invalid Selection!!!"
   CmbFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" where " & Gen_WcLocation)
Else
   Call TableMst_Location("")
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbLocation.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If Trim(CmbLocation.Text) = "" Then
   Call Grid_Head
   CmbCommodity.Text = ""
   CmbClientName.Text = ""
   Exit Sub
End If
If tmp <> CmbLocation.Text Then
   Call Grid_Head
   CmbCommodity.Text = ""
   CmbClientName.Text = ""
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName= '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
CmbLocation.Text = ""
CmbCommodity.Text = ""
CmbFlag.Text = ""
CmbClientName.Text = ""
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GSLvsWHRMonitoringMIS.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Single
Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GSLvsWHRMonitoringMIS.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
   For c = 1 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, c) = MSHFlexGrid1.TextMatrix(I, c)
   Next
   mRow = mRow + 1
Next
Set oWS = oWB.Worksheets("Sheet2")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
   For c = 0 To MSHFlexGrid2.Cols - 1
      oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
   Next
   mRow = mRow + 1
Next

Set oWS = oWB.Worksheets("Sheet3")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid4.Rows - 1
   For c = 0 To MSHFlexGrid4.Cols - 1
      oWS.Cells(mRow, c + 1) = MSHFlexGrid4.TextMatrix(I, c)
   Next
   mRow = mRow + 1
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()

If Trim(CmbLocation.Text) = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
If Trim(CmbCommodity.Text) = "" Then
   MsgBox "Blank Commodity not allowed!!!"
   CmbCommodity.SetFocus
   Exit Sub
End If
If Trim(CmbClientName.Text) = "" Then
   MsgBox "Blank Client Name not allowed!!!"
   CmbClientName.SetFocus
   Exit Sub
End If
Dim wc, wc1 As Variant

If Option1.Value = True Then
   wc = " where mcl.exchangetypeid = 1"
   If CmbLocation.Text <> "" Then
      wc = wc & " and ml.locationid = " & mLocationID
   ElseIf Gen_WcLocation <> "" Then
      wc = wc & " and  ml. " & Gen_WcLocation
   End If
   If CmbClientName.Text <> "" Then
      wc = wc & " and mcl.clientidrow = " & mClientID
   End If
   If CmbCommodity.Text <> "" Then
      wc = wc & " and mcom.commodityid = " & mCommodityID
   End If
   If CmbFlag.Text <> "" Then
      If LCase(CmbFlag.Text) = "transfer" Then
         wc1 = " and gsl.flag in ('Z') "
      ElseIf LCase(CmbFlag.Text) = "non transfer" Then
         wc1 = " and gsl.flag not in ('Z')"
      End If
   End If
   Call Grid_Head
   MsgBox ("Wait till Next Message!!!")
   
   tmp = " select ml.locationname, mcmp.cmpname, mg.godownno, gsl.CDTFDate, mcl.clientname , mcom.Commodity, gsl.gslid, gsl.DepositeWeight, gsl.DepositeBags, "
   tmp = tmp & " gsl.WHRBags, gsl.WHRWeight, gsl.BalanceBags, gsl.BalanceWeight, gsl.DepositeBags - gsl.whrbags differencebags,"
   tmp = tmp & " gsl.DepositeWeight - gsl.WHRWeight differenceweight, gsl.flag, gsl.sm_prefix, gsl.remarks"
   tmp = tmp & " from mst_gsl gsl"
   tmp = tmp & " inner join mst_location ml on ml.sm_prefix = gsl.sm_prefix"
   tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = gsl.commodityid"
   tmp = tmp & " inner join mst_client mcl on mcl.clientidrow = gsl.clientidrow"
   tmp = tmp & " inner join mst_godown mg on mg.godownid = gsl.godownid"
   tmp = tmp & " inner join mst_cmp mcmp on mg.cmpid = mcmp.cmpid"
   tmp = tmp & wc & wc1
   'tmp = tmp & " order by locationname, mcmp.cmpname, mg.godownno, mcl.clientname, mcom.commodity"
   tmp = tmp & " order by mcmp.cmpname, mg.godownno, gsl.CDTFDate"
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
      CmdExcel.Enabled = True
      CmdSave.Enabled = True
      Label22.Caption = TmpRs.RecordCount
      ProgressBar1.Max = TmpRs.RecordCount + 3
      For I = 2 To TmpRs.RecordCount + 1
         MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName) ''IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)  'IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
         MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!CDTFDate), "", Format(TmpRs!CDTFDate, "dd-MMM-yyyy")) 'IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
         MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!GSLID), "", TmpRs!GSLID)
         MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag) '"Flag"
         MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!DepositeBags), "", TmpRs!DepositeBags) '"Deposite Bags"
         MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!WHRBags), "", TmpRs!WHRBags) '"WHR Bags"
         MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!DepositeWeight), "", TmpRs!DepositeWeight) '"Deposite Qty"
         MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!WHRWeight), "", TmpRs!WHRWeight) '"WHR Qty"
         MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags) '"Balance Bags"
         MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight) '"Balance Qty"
         MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs!DifferenceBags), "", TmpRs!DifferenceBags) '"Difference in Bags"
         MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs!DifferenceWeight), "", TmpRs!DifferenceWeight) '"Difference in Qty"
         
         MSHFlexGrid1.TextMatrix(I, 15) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
         MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
         
         MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(MSHFlexGrid1.TextMatrix(0, 6))
         MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(I, 7)) + Val(MSHFlexGrid1.TextMatrix(0, 7))
         MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(I, 8)) + Val(MSHFlexGrid1.TextMatrix(0, 8))
         MSHFlexGrid1.TextMatrix(0, 9) = Val(MSHFlexGrid1.TextMatrix(I, 9)) + Val(MSHFlexGrid1.TextMatrix(0, 9))
         MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(I, 11)) + Val(MSHFlexGrid1.TextMatrix(0, 11))
         MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(I, 12)) + Val(MSHFlexGrid1.TextMatrix(0, 12))
         MSHFlexGrid1.TextMatrix(0, 13) = Val(MSHFlexGrid1.TextMatrix(I, 13)) + Val(MSHFlexGrid1.TextMatrix(0, 13))
         MSHFlexGrid1.TextMatrix(0, 14) = Val(MSHFlexGrid1.TextMatrix(I, 14)) + Val(MSHFlexGrid1.TextMatrix(0, 14))
         MSHFlexGrid1.Row = I
         MSHFlexGrid1.Col = 0
         MSHFlexGrid1.CellFontName = "Wingdings"
         MSHFlexGrid1.CellFontSize = 12
         MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
         MSHFlexGrid1.Text = strUnChecked
   
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
      ProgressBar1.Value = ProgressBar1.Max
      
      tmp = " select"
      tmp = tmp & " ml.locationname, mcmp.cmpname, mg.godownno, mcl.clientname, mcom.commodity, whr.whrno, whr.whrdate, whr.WHRNoOfBags, whr.WHRQuantity, whr.DeliveredQuantity,"
      tmp = tmp & " whr.DeliveredNoOfBags , whr.pledgeid, convert(varchar,whr.cmpid)+'-'+convert(varchar, whr.whrno-whr.cmpid*1000000) whrno1, sum(plg.PledgeNoOfBags) PledgeNoOfBags, sum(plg.pledgeQuantity) pledgeQuantity"
      tmp = tmp & " from txn_whr_detail whr"
      tmp = tmp & " inner join mst_location ml on ml.sm_prefix = whr.sm_prefix"
      tmp = tmp & " inner join mst_cmp mcmp on mcmp.cmpid = whr.cmpid"
      tmp = tmp & " inner join mst_godown mg on mg.godownid = whr.godownid"
      tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = whr.commodityid"
      tmp = tmp & " inner join mst_client mcl on mcl.clientidrow = whr.clientidrow"
      tmp = tmp & " left join txn_pledgecreation_gsl plg on plg.pledgeid = whr.pledgeid and whr.sm_Prefix = plg.sm_prefix"
      tmp = tmp & wc
      tmp = tmp & " group by ml.locationname, mcmp.cmpname, mg.godownno, mcl.clientname, mcom.commodity, whr.whrno, whr.whrdate,"
      tmp = tmp & " whr.WHRNoOfBags , whr.WHRQuantity, whr.DeliveredQuantity, whr.DeliveredNoOfBags, whr.pledgeid, whr.cmpid"
      'tmp = tmp & " order by ml.locationname, mcmp.cmpname, mg.godownno, mcl.clientname, mcom.commodity, whr.whrno"
      tmp = tmp & " order by mcmp.cmpname, mg.godownno, whr.whrno"
      Call TmpTable
      
      If TmpRs.RecordCount > 0 Then
         MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
         ProgressBar1.Value = 0
         ProgressBar1.Max = TmpRs.RecordCount + 3
         For I = 2 To TmpRs.RecordCount + 1
            MSHFlexGrid2.TextMatrix(I, 0) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
            MSHFlexGrid2.TextMatrix(I, 1) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
            MSHFlexGrid2.TextMatrix(I, 2) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
            MSHFlexGrid2.TextMatrix(I, 3) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
            MSHFlexGrid2.TextMatrix(I, 4) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
            MSHFlexGrid2.TextMatrix(I, 5) = IIf(IsNull(TmpRs!whrno), "", TmpRs!whrno)
            MSHFlexGrid2.TextMatrix(I, 6) = IIf(IsNull(TmpRs!WHRDate), "", Format(TmpRs!WHRDate, "dd-MMM-yyyy"))
            MSHFlexGrid2.TextMatrix(I, 7) = IIf(IsNull(TmpRs!WHRNoOfBags), "", TmpRs!WHRNoOfBags)
            MSHFlexGrid2.TextMatrix(I, 8) = IIf(IsNull(TmpRs!WHRQuantity), "", TmpRs!WHRQuantity)
            MSHFlexGrid2.TextMatrix(I, 9) = IIf(IsNull(TmpRs!DeliveredNoOfBags), "", TmpRs!DeliveredNoOfBags)
            MSHFlexGrid2.TextMatrix(I, 10) = IIf(IsNull(TmpRs!DeliveredQuantity), "", TmpRs!DeliveredQuantity)
            MSHFlexGrid2.TextMatrix(I, 11) = IIf(IsNull(TmpRs!PledgeID), "", TmpRs!PledgeID)
            MSHFlexGrid2.TextMatrix(I, 12) = IIf(IsNull(TmpRs!PledgeNoOfBags), "", TmpRs!PledgeNoOfBags)
            MSHFlexGrid2.TextMatrix(I, 13) = IIf(IsNull(TmpRs!PledgeQuantity), "", TmpRs!PledgeQuantity)
            MSHFlexGrid2.TextMatrix(I, 14) = IIf(IsNull(TmpRs!whrno1), "", TmpRs!whrno1)
            
            MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(I, 7)) + Val(MSHFlexGrid2.TextMatrix(0, 7))
            MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(I, 8)) + Val(MSHFlexGrid2.TextMatrix(0, 8))
            MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(I, 9)) + Val(MSHFlexGrid2.TextMatrix(0, 9))
            MSHFlexGrid2.TextMatrix(0, 10) = Val(MSHFlexGrid2.TextMatrix(I, 10)) + Val(MSHFlexGrid2.TextMatrix(0, 10))
            MSHFlexGrid2.TextMatrix(0, 12) = Val(MSHFlexGrid2.TextMatrix(I, 12)) + Val(MSHFlexGrid2.TextMatrix(0, 12))
            MSHFlexGrid2.TextMatrix(0, 13) = Val(MSHFlexGrid2.TextMatrix(I, 13)) + Val(MSHFlexGrid2.TextMatrix(0, 13))
            ProgressBar1.Value = ProgressBar1.Value + 1
            TmpRs.MoveNext
         Next
      End If
      
      Label2.Caption = "Difference between Deposited and WHR Bags is " & Format(Val(MSHFlexGrid1.TextMatrix(0, 6)) - Val(MSHFlexGrid2.TextMatrix(0, 7)), "####0.000")
      Label4.Caption = "Difference between Deposited and WHR Qty is " & Format(Val(MSHFlexGrid1.TextMatrix(0, 8)) - Val(MSHFlexGrid2.TextMatrix(0, 8)), "#####0.000")
      Label7.Caption = "Difference between GSL-WHR and WHR Bags is " & Format(Val(MSHFlexGrid1.TextMatrix(0, 7)) - Val(MSHFlexGrid2.TextMatrix(0, 7)), "#####0.000")
      Label9.Caption = "Difference between GSL-WHR and WHR Qty is " & Format(Val(MSHFlexGrid1.TextMatrix(0, 9)) - Val(MSHFlexGrid2.TextMatrix(0, 8)), "#####0.000")
      
      tmp = " select mcmp.cmpname, mg.godownno, sum(gsl.DepositeWeight) DepositeWeight, sum(gsl.DepositeBags) DepositeBags, "
      tmp = tmp & " sum(gsl.WHRBags) GSLWHRBags, sum(gsl.WHRWeight) GSLWHRQty, "
      tmp = tmp & " (select sum(WHRNoOfBags) from txn_whr_detail where clientidrow = mcl.clientidrow and gsl.sm_prefix = sm_prefix and commodityid = mcom.commodityid and godownid = mg.godownid ) WHRNoOfBags,"
      tmp = tmp & " (select sum(WHRQuantity) from txn_whr_detail where clientidrow = mcl.clientidrow and gsl.sm_prefix = sm_prefix and commodityid = mcom.commodityid and godownid = mg.godownid ) WHRQuantity"
      tmp = tmp & " from mst_gsl gsl"
      tmp = tmp & " inner join mst_location ml on ml.sm_prefix = gsl.sm_prefix"
      tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = gsl.commodityid"
      tmp = tmp & " inner join mst_client mcl on mcl.clientidrow = gsl.clientidrow"
      tmp = tmp & " inner join mst_godown mg on mg.godownid = gsl.godownid"
      tmp = tmp & " inner join mst_cmp mcmp on mg.cmpid = mcmp.cmpid"
      tmp = tmp & wc
      tmp = tmp & " group by mcmp.cmpname, mg.godownno, mcl.clientidrow, gsl.sm_prefix, mcom.commodityid, mg.godownid"
      tmp = tmp & " order by mcmp.cmpname, mg.godownno"
      Call TmpTable
      
      If TmpRs.RecordCount > 0 Then
         MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + TmpRs.RecordCount
         ProgressBar1.Value = 0
         ProgressBar1.Max = TmpRs.RecordCount + 3
         For I = 2 To TmpRs.RecordCount + 1
            MSHFlexGrid4.TextMatrix(I, 0) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
            MSHFlexGrid4.TextMatrix(I, 1) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
            MSHFlexGrid4.TextMatrix(I, 2) = IIf(IsNull(TmpRs!DepositeBags), 0, TmpRs!DepositeBags)
            MSHFlexGrid4.TextMatrix(I, 3) = IIf(IsNull(TmpRs!DepositeWeight), 0, TmpRs!DepositeWeight)
            MSHFlexGrid4.TextMatrix(I, 4) = IIf(IsNull(TmpRs!GSLWHRBags), 0, TmpRs!GSLWHRBags)
            MSHFlexGrid4.TextMatrix(I, 5) = IIf(IsNull(TmpRs!GSLWHRQty), 0, TmpRs!GSLWHRQty)
            MSHFlexGrid4.TextMatrix(I, 6) = IIf(IsNull(TmpRs!WHRNoOfBags), "", TmpRs!WHRNoOfBags)
            MSHFlexGrid4.TextMatrix(I, 7) = IIf(IsNull(TmpRs!WHRQuantity), "", TmpRs!WHRQuantity)
            
            MSHFlexGrid4.TextMatrix(0, 2) = Val(MSHFlexGrid4.TextMatrix(I, 2)) + Val(MSHFlexGrid4.TextMatrix(0, 2))
            MSHFlexGrid4.TextMatrix(0, 3) = Val(MSHFlexGrid4.TextMatrix(I, 3)) + Val(MSHFlexGrid4.TextMatrix(0, 3))
            MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(I, 4)) + Val(MSHFlexGrid4.TextMatrix(0, 4))
            MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(I, 5)) + Val(MSHFlexGrid4.TextMatrix(0, 5))
            MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(I, 6)) + Val(MSHFlexGrid4.TextMatrix(0, 6))
            MSHFlexGrid4.TextMatrix(0, 7) = Val(MSHFlexGrid4.TextMatrix(I, 7)) + Val(MSHFlexGrid4.TextMatrix(0, 7))
            ProgressBar1.Value = ProgressBar1.Value + 1
            TmpRs.MoveNext
         Next
      End If
      
      ProgressBar1.Value = ProgressBar1.Max
      MsgBox "Done!!!"
   Else
      MsgBox "No Record found!!!"
   End If

ElseIf Option2.Value = True Then
   wc = " where gsl.flag <> 'Z'"
   If CmbLocation.Text <> "" Then
      wc = wc & " and ml.locationid = " & mLocationID
   ElseIf Gen_WcLocation <> "" Then
      wc = wc & " and  ml. " & Gen_WcLocation
   End If
   If CmbClientName.Text <> "" Then
      wc = wc & " and mcl.clientidrow = " & mClientID
   End If
   If CmbCommodity.Text <> "" Then
      wc = wc & " and mcom.commodityid = " & mCommodityID
   End If
   Call Grid_Head
   MsgBox ("Wait till Next Message!!!")
   
   tmp = " select ml.locationname, mcl.clientname, mcom.commodity, gsl.gslid, gsl.DepositeWeight, gsl.DepositeBags, "
   tmp = tmp & " gsl.CDTFBags, gsl.CDTFWeight, gsl.BalanceBags, gsl.BalanceWeight, gsl.DepositeBags - gsl.CDTFbags differencebags,"
   tmp = tmp & " gsl.DepositeWeight - gsl.CDTFWeight differenceweight, gsl.flag, gsl.sm_prefix, gsl.remarks"
   tmp = tmp & " from mst_gsl gsl"
   tmp = tmp & " inner join mst_location ml on ml.sm_prefix = gsl.sm_prefix"
   tmp = tmp & " inner join mst_commodity mcom on mcom.commodityid = gsl.commodityid"
   tmp = tmp & " inner join mst_client mcl on mcl.clientidrow = gsl.clientidrow"
   tmp = tmp & wc
   'tmp = tmp & " order by locationname, mcl.clientname, mcom.commodity"
   tmp = tmp & " order by gsl.gslid"
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
      CmdExcel.Enabled = True
      CmdSave.Enabled = True
      Label22.Caption = TmpRs.RecordCount
      ProgressBar1.Max = TmpRs.RecordCount + 3
      For I = 2 To TmpRs.RecordCount + 1
         MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
         MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
         MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!GSLID), "", TmpRs!GSLID)
         MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!Flag), "", TmpRs!Flag) '"Flag"
         MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!DepositeBags), "", TmpRs!DepositeBags) '"Deposite Bags"
         MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!CDTFBags), "", TmpRs!CDTFBags) '"WHR Bags"
         MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs!DepositeWeight), "", TmpRs!DepositeWeight) '"Deposite Qty"
         MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs!CDTFWeight), "", TmpRs!CDTFWeight) '"WHR Qty"
         MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags) '"Balance Bags"
         MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight) '"Balance Qty"
         MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs!DifferenceBags), "", TmpRs!DifferenceBags) '"Difference in Bags"
         MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs!DifferenceWeight), "", TmpRs!DifferenceWeight) '"Difference in Qty"
         MSHFlexGrid1.TextMatrix(I, 15) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
         MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
         
         MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(I, 6)) + Val(MSHFlexGrid1.TextMatrix(0, 6))
         MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(I, 7)) + Val(MSHFlexGrid1.TextMatrix(0, 7))
         MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(I, 8)) + Val(MSHFlexGrid1.TextMatrix(0, 8))
         MSHFlexGrid1.TextMatrix(0, 9) = Val(MSHFlexGrid1.TextMatrix(I, 9)) + Val(MSHFlexGrid1.TextMatrix(0, 9))
         MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(I, 11)) + Val(MSHFlexGrid1.TextMatrix(0, 11))
         MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(I, 12)) + Val(MSHFlexGrid1.TextMatrix(0, 12))
         MSHFlexGrid1.TextMatrix(0, 13) = Val(MSHFlexGrid1.TextMatrix(I, 13)) + Val(MSHFlexGrid1.TextMatrix(0, 13))
         MSHFlexGrid1.TextMatrix(0, 14) = Val(MSHFlexGrid1.TextMatrix(I, 14)) + Val(MSHFlexGrid1.TextMatrix(0, 14))
         MSHFlexGrid1.Row = I
         MSHFlexGrid1.Col = 0
         MSHFlexGrid1.CellFontName = "Wingdings"
         MSHFlexGrid1.CellFontSize = 12
         MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
         MSHFlexGrid1.Text = strUnChecked
   
         ProgressBar1.Value = ProgressBar1.Value + 1
         TmpRs.MoveNext
      Next
      ProgressBar1.Value = ProgressBar1.Max
   Else
      MsgBox "No Record found!!!"
   End If
End If
End Sub

Private Sub CmdSave_Click()
If ChkRemarks = False Then Exit Sub
If Option1.Value = True Then
   For I = 2 To MSHFlexGrid1.Rows - 1
      If MSHFlexGrid1.TextMatrix(I, 0) = strChecked And MSHFlexGrid1.TextMatrix(I, 10) <> "" Then
         tmp = " Update Mst_GSL set UpdatedBy = '" & Gen_UserLoginID & "', UpdatedDate = '" & Now & "', WHRBags = " & Val(MSHFlexGrid1.TextMatrix(I, 7)) & ", WHRWeight = " & Val(MSHFlexGrid1.TextMatrix(I, 9)) & ", Remarks = '" & MSHFlexGrid1.TextMatrix(I, 10) & "' Where GSLID = " & MSHFlexGrid1.TextMatrix(I, 4) & " And SM_Prefix = '" & MSHFlexGrid1.TextMatrix(I, 15) & "'"
         Call TmpTable
      End If
   Next

ElseIf Option2.Value = True Then
   For I = 2 To MSHFlexGrid1.Rows - 1
      If MSHFlexGrid1.TextMatrix(I, 0) = strChecked And MSHFlexGrid1.TextMatrix(I, 10) <> "" Then
         tmp = " Update Mst_GSL set UpdatedBy = '" & Gen_UserLoginID & "', UpdatedDate = '" & Now & "', CDTFBags = " & Val(MSHFlexGrid1.TextMatrix(I, 7)) & ", CDTFWeight = " & Val(MSHFlexGrid1.TextMatrix(I, 9)) & ", BalanceBags = " & Val(MSHFlexGrid1.TextMatrix(I, 11)) & ", BalanceWeight = " & Val(MSHFlexGrid1.TextMatrix(I, 12)) & ", Remarks = '" & MSHFlexGrid1.TextMatrix(I, 10) & "' Where GSLID = " & MSHFlexGrid1.TextMatrix(I, 4) & " And SM_Prefix = '" & MSHFlexGrid1.TextMatrix(I, 15) & "'"
         Call TmpTable
      End If
   Next
End If
Call CmdGetReport_Click
End Sub

Private Function ChkRemarks() As Boolean
ChkRemarks = True
For I = 2 To MSHFlexGrid1.Rows - 1
   If MSHFlexGrid1.TextMatrix(I, 0) = strChecked And MSHFlexGrid1.TextMatrix(I, 10) = "" Then
      MsgBox "Blank Remarks not allowed!!!"
      ChkRemarks = False
      Exit Function
   End If
Next
End Function

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame11.Enabled = False
   Exit Sub
End If

If LCase(Gen_UserRole) <> "power" Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame11.Enabled = False
   Exit Sub
End If

Call CmdClear_Click
Frame11.Enabled = True

End Sub
Private Sub Grid_Head()
Call Grid_Head1
Call Grid_Head2
Call Grid_Head3
Label22.Caption = ""
CmdExcel.Enabled = False
CmdSave.Enabled = False
ProgressBar1.Value = 0
Label2.Caption = "Difference between Deposited and WHR Bags is 0"
Label4.Caption = "Difference between Deposited and WHR Qty is 0"
Label7.Caption = "Difference between GSL-WHR and WHR Bags is 0"
Label9.Caption = "Difference between GSL-WHR and WHR Qty is 0"
With MSHFlexGrid1
   .Clear
   .WordWrap = True
   .Rows = 3
   .FixedRows = 2
   .Cols = 17
   
   .TextMatrix(1, 0) = "Select"
   .TextMatrix(1, 1) = "CMP Name" '"Location"
   .TextMatrix(1, 2) = "Godown No" '"Client Name"
   .TextMatrix(1, 3) = "CDTF Date" '"Commodity"
   .TextMatrix(1, 4) = "GSLID"
   .TextMatrix(1, 5) = "Flag"
   .TextMatrix(1, 6) = "Deposite Bags"
   If Option1.Value = True Then
      .TextMatrix(1, 7) = "WHR Bags"
   ElseIf Option2.Value = True Then
      .TextMatrix(1, 7) = "CDTF Bags"
   End If
   .TextMatrix(1, 8) = "Deposite Qty"
   If Option1.Value = True Then
      .TextMatrix(1, 9) = "WHR Qty"
   ElseIf Option2.Value = True Then
      .TextMatrix(1, 9) = "CDTF Qty"
   End If
   .TextMatrix(1, 10) = "Enter Remarks"
   .TextMatrix(1, 11) = "Balance Bags"
   .TextMatrix(1, 12) = "Balance Qty"
   .TextMatrix(1, 13) = "Difference in Bags"
   .TextMatrix(1, 14) = "Difference in Qty"
   .TextMatrix(1, 15) = "SM_Prefix"
   .TextMatrix(1, 16) = "Current Remarks"
      
   .ColWidth(0) = 700
   .ColWidth(1) = 1500 '0
   .ColWidth(2) = 700 '0
   .ColWidth(3) = 1500 '0
   .ColWidth(4) = 700
   .ColWidth(5) = 600
   .ColWidth(6) = 900
   .ColWidth(7) = 900
   .ColWidth(8) = 1000
   .ColWidth(9) = 1000
   .ColWidth(10) = 4500
   .ColWidth(11) = 900
   .ColWidth(12) = 1000
   If Option1.Value = True Then
      .ColWidth(13) = 900
      .ColWidth(14) = 1000
   ElseIf Option1.Value = True Then
      .ColWidth(13) = 0
      .ColWidth(14) = 0
   End If
   
   .ColWidth(15) = 0 '800
   .ColWidth(16) = 4500
   .RowHeight(1) = 800
End With
End Sub

Private Sub Grid_Head1()
With MSHFlexGrid2
   .Clear
   .WordWrap = True
   .Rows = 3
   .FixedRows = 2
   .Cols = 15
   
   .TextMatrix(1, 0) = "Location"
   .TextMatrix(1, 1) = "CMP Name"
   .TextMatrix(1, 2) = "Godown No"
   .TextMatrix(1, 3) = "Client Name"
   .TextMatrix(1, 4) = "Commodity"
   .TextMatrix(1, 5) = "WHR No"
   .TextMatrix(1, 6) = "WHR Date"
   .TextMatrix(1, 7) = "WHR No of Bags"
   .TextMatrix(1, 8) = "WHR Qty"
   .TextMatrix(1, 9) = "Delivered Bags"
   .TextMatrix(1, 10) = "Delivered Qty"
   .TextMatrix(1, 11) = "Pledge ID"
   .TextMatrix(1, 12) = "Pledge Bags"
   .TextMatrix(1, 13) = "Pledge Qty"
   .TextMatrix(1, 14) = "WHR No as per CM"
      
   .ColWidth(0) = 0 '1500
   .ColWidth(1) = 2000
   .ColWidth(2) = 1200
   .ColWidth(3) = 0 '2500
   .ColWidth(4) = 0 '1200
   .ColWidth(5) = 1200
   .ColWidth(6) = 1200
   .ColWidth(7) = 1200
   .ColWidth(8) = 1200
   .ColWidth(9) = 1200
   .ColWidth(10) = 1200
   .ColWidth(11) = 1200
   .ColWidth(12) = 1200
   .ColWidth(13) = 800
   .ColWidth(14) = 1200
   
   .RowHeight(1) = 800
End With
End Sub

Private Sub Grid_Head2()
With MSHFlexGrid3
   .Clear
   .Rows = 2
   .Cols = 5
   
   .TextMatrix(0, 0) = "GSLID"
   .TextMatrix(0, 1) = "Deposite Bags"
   .TextMatrix(0, 2) = "Deposite Weight"
   .TextMatrix(0, 3) = "Withdraw Bags"
   .TextMatrix(0, 4) = "Withdraw Weight"
      
   .ColWidth(0) = 1500
   .ColWidth(1) = 2000
   .ColWidth(2) = 2000
   .ColWidth(3) = 2000 '2500
   .ColWidth(4) = 2000 '1200
End With
End Sub

Private Sub Grid_Head3()
With MSHFlexGrid4
   .Clear
   .WordWrap = True
   .Rows = 3
   .FixedRows = 2
   .Cols = 8
   
   .TextMatrix(1, 0) = "CMP Name"
   .TextMatrix(1, 1) = "Godown No"
   .TextMatrix(1, 2) = "Deposite Bags"
   .TextMatrix(1, 3) = "Deposite Weight"
   .TextMatrix(1, 4) = "WHR Bags as per GSL"
   .TextMatrix(1, 5) = "WHR Qty as per GSL"
   .TextMatrix(1, 6) = "WHR Bags as per WHR Detail"
   .TextMatrix(1, 7) = "WHR Qty as per WHR Detail"
      
   .ColWidth(0) = 2200
   .ColWidth(1) = 1500
   .ColWidth(2) = 1500
   .ColWidth(3) = 1500
   .ColWidth(4) = 1500
   .ColWidth(5) = 1500
   .ColWidth(6) = 1500
   .ColWidth(7) = 1500
   
   .RowHeight(1) = 800
End With
End Sub

Private Sub SetScrollBars()
'    Frame1.Width = 10000
'    Frame1.Height = 20000
VScroll1.Top = 0
HScroll1.Left = 0
VScroll1.Left = Me.Width - (VScroll1.Width + 120)
HScroll1.Top = Me.Height - ((HScroll1.Height * 2) + 250)
VScroll1.Height = Me.Height - 500
HScroll1.Width = Me.Width - 500
VScroll1.Min = 0
HScroll1.Min = 0
VScroll1.Max = Me.ScaleHeight - Frame4.Height
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame4.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame4.Width >= Me.Width Then HScroll1.Enabled = True
If Frame4.Width < Me.Width Then HScroll1.Enabled = False
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub MSHFlexGrid1_Click()
mColsel = MSHFlexGrid1.ColSel
mRowsel = MSHFlexGrid1.RowSel
If Option1.Value = True Then
   If mColsel = 0 And mRowsel <> MSHFlexGrid1.Rows - 1 Then
      If MSHFlexGrid1.TextMatrix(mRowsel, mColsel) = strUnChecked Then
         MSHFlexGrid1.TextMatrix(mRowsel, mColsel) = strChecked
      Else
         MSHFlexGrid1.TextMatrix(mRowsel, mColsel) = strUnChecked
      End If
   ElseIf MSHFlexGrid1.TextMatrix(mRowsel, 0) = strChecked Then
      If (mColsel = 7 Or mColsel = 9 Or mColsel = 10) And mRowsel <> MSHFlexGrid1.Rows - 1 Then
         Call VisibleEnterTxt
      End If
   End If
ElseIf Option2.Value = True Then
   If mColsel = 0 And mRowsel <> MSHFlexGrid1.Rows - 1 Then
      If MSHFlexGrid1.TextMatrix(mRowsel, mColsel) = strUnChecked Then
         MSHFlexGrid1.TextMatrix(mRowsel, mColsel) = strChecked
         Call FillDepositeDetails
      Else
         MSHFlexGrid1.TextMatrix(mRowsel, mColsel) = strUnChecked
      End If
   ElseIf MSHFlexGrid1.TextMatrix(mRowsel, 0) = strChecked And MSHFlexGrid1.TextMatrix(mRowsel, 4) = MSHFlexGrid3.TextMatrix(1, 0) Then
      If (mColsel = 5 Or mColsel = 7 Or mColsel = 9 Or mColsel = 10) And mRowsel <> MSHFlexGrid1.Rows - 1 Then
         Call VisibleEnterTxt
      End If
   End If
End If
End Sub

Private Sub FillDepositeDetails()
Call Grid_Head2
tmp = " Select GSLID, sum(NoOfbags) Bags,sum(DematWeight) Weight ,TxnType from Txn_CXTF_GSL"
tmp = tmp & " Where GSLID = " & MSHFlexGrid1.TextMatrix(mRowsel, 4) & " And SM_Prefix = '" & MSHFlexGrid1.TextMatrix(mRowsel, 15) & "'"
tmp = tmp & " group by GSLID, TxnType"
tmp = tmp & " order by TxnType"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.TextMatrix(1, 0) = IIf(IsNull(TmpRs!GSLID), 0, TmpRs!GSLID)
   MSHFlexGrid3.TextMatrix(1, 1) = IIf(IsNull(TmpRs!Bags), 0, TmpRs!Bags)
   MSHFlexGrid3.TextMatrix(1, 2) = IIf(IsNull(TmpRs!Weight), 0, TmpRs!Weight)
   TmpRs.MoveNext
   If Not TmpRs.EOF Then
      MSHFlexGrid3.TextMatrix(1, 3) = IIf(IsNull(TmpRs!Bags), 0, TmpRs!Bags)
      MSHFlexGrid3.TextMatrix(1, 4) = IIf(IsNull(TmpRs!Weight), 0, TmpRs!Weight)
   Else
      MSHFlexGrid3.TextMatrix(1, 3) = 0
      MSHFlexGrid3.TextMatrix(1, 4) = 0
   End If
End If
End Sub

Private Sub VisibleEnterTxt()
With TxtEnterData
   .Text = MSHFlexGrid1.TextMatrix(mRowsel, mColsel)
   .Left = MSHFlexGrid1.Left + MSHFlexGrid1.CellLeft
   .Top = MSHFlexGrid1.Top + MSHFlexGrid1.CellTop
   .Width = MSHFlexGrid1.CellWidth - 10
   .Height = MSHFlexGrid1.CellHeight - 10
   .Visible = True
   .ZOrder
   .SetFocus
End With
End Sub

'Private Sub TxtEnterData_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyDown Then
'   If CheckTxtEnterData = False Then Exit Sub
'   MSHFlexGrid1.TextMatrix(mRowsel, mColsel) = TxtEnterData.Text
'   MSHFlexGrid1.SetFocus
'   DoEvents
'   If MSHFlexGrid1.Row < MSHFlexGrid1.Rows - 2 Then
'      MSHFlexGrid1.Row = MSHFlexGrid1.Row + 1
'      mRowsel = MSHFlexGrid1.Row
'      Call MSHFlexGrid1_Click
'   End If
'ElseIf KeyCode = vbKeyUp Then
'   If CheckTxtEnterData = False Then Exit Sub
'   MSHFlexGrid1.TextMatrix(mRowsel, mColsel) = TxtEnterData.Text
'   MSHFlexGrid1.SetFocus
'   DoEvents
'   If MSHFlexGrid1.Row > 1 Then
'      MSHFlexGrid1.Row = MSHFlexGrid1.Row - 1
'      mRowSelTxt = MSHFlexGrid1.Row
'      Call MSHFlexGrid1_Click
'   End If
'End If
'End Sub

Private Function CheckTxtEnterData() As Boolean
CheckTxtEnterData = True
If Trim(TxtEnterData.Text) = "" Then Exit Function
If mColsel = 10 Then Exit Function
If mColsel = 5 Then
   If TxtEnterData.Text <> "J" And TxtEnterData.Text <> "O" Then
      MsgBox "Enter 'J' or 'O' as Flag!!!"
      TxtEnterData.SetFocus
      CheckTxtEnterData = False
      Exit Function
   Else
      Exit Function
   End If
ElseIf IsNumeric(TxtEnterData.Text) = False Or Val(TxtEnterData.Text) < 0 Then
   MsgBox "Enter Numeric Value!!!"
   TxtEnterData.SetFocus
   CheckTxtEnterData = False
   Exit Function
ElseIf mColsel = 7 And Val(TxtEnterData.Text) > Val(MSHFlexGrid1.TextMatrix(mRowsel, mColsel - 1)) Then
   MsgBox "You can not enter more than deposite bags!!!"
   TxtEnterData.SetFocus
   CheckTxtEnterData = False
   Exit Function
ElseIf mColsel = 7 And InStr(1, TxtEnterData.Text, ".") Then
   MsgBox "Bags are not in Decimal Place!!!"
   TxtEnterData.SetFocus
   CheckTxtEnterData = False
   Exit Function
ElseIf mColsel = 9 And Val(TxtEnterData.Text) > Val(MSHFlexGrid1.TextMatrix(mRowsel, mColsel - 1)) Then
   MsgBox "You can not enter more than deposite qty!!!"
   TxtEnterData.SetFocus
   CheckTxtEnterData = False
   Exit Function
ElseIf Option2.Value = True And mColsel = 7 And Val(TxtEnterData.Text) > Val(MSHFlexGrid3.TextMatrix(1, 3)) Then
   MsgBox "You can not enter more than Withdrawal Bags!!!"
   TxtEnterData.SetFocus
   CheckTxtEnterData = False
   Exit Function
ElseIf Option2.Value = True And mColsel = 9 And Val(TxtEnterData.Text) > Val(MSHFlexGrid3.TextMatrix(1, 4)) Then
   MsgBox "You can not enter more than Withdrawal Qty!!!"
   TxtEnterData.SetFocus
   CheckTxtEnterData = False
   Exit Function
End If
End Function

Private Sub MSHFlexGrid1_Scroll()
TxtEnterData.Text = ""
TxtEnterData.Visible = False
End Sub

Private Sub Option1_Click()
MSHFlexGrid3.Visible = False
MSHFlexGrid1.Height = 6465
CmbFlag.Enabled = True
Label2.Visible = True
Label4.Visible = True
Label7.Visible = True
Label9.Visible = True
Call Grid_Head
End Sub

Private Sub Option2_Click()
MSHFlexGrid3.Visible = True
MSHFlexGrid1.Height = 5625
CmbFlag.Text = ""
CmbFlag.Enabled = False
Label2.Visible = False
Label4.Visible = False
Label7.Visible = False
Label9.Visible = False
Call Grid_Head
End Sub

Private Sub TxtEnterData_LostFocus()
If TxtEnterData.Text = "" Then
   TxtEnterData.Visible = False
   Exit Sub
End If
If CheckTxtEnterData = False Then
   TxtEnterData.SetFocus
   Exit Sub
Else
   MSHFlexGrid1.TextMatrix(mRowsel, mColsel) = TxtEnterData.Text
   TxtEnterData.Visible = False
'   If (mColsel = 7 Or mColsel = 9) And Val(TxtEnterData.Text) = 0 Then
'      MSHFlexGrid1.TextMatrix(mRowsel, 6) = 0
'      MSHFlexGrid1.TextMatrix(mRowsel, 8) = 0
'   End If
End If
End Sub

Private Sub VScroll1_Change()
   Dim L As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   L = (a * x) / 1000
   Frame4.Top = -L

'Frame3.Top = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Dim L As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   L = (a * x) / 1000
   Frame4.Top = -L

'Frame3.Top = VScroll1.Value
End Sub

Private Sub HScroll1_Change()
   Dim L As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   L = (a * x) / 1000
   Frame4.Left = -L
End Sub

Private Sub HScroll1_Scroll()
   Dim L As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   L = (a * x) / 1000
   Frame4.Left = -L
End Sub

