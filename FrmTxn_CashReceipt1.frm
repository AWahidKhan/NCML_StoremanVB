VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_CashReceipt 
   Caption         =   "Cash Bank Receipt"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13125
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   13125
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H008080FF&
      Caption         =   "Hologram No "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4800
      TabIndex        =   46
      Top             =   2520
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
         Height          =   400
         Left            =   2880
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton CmdSaveHologram 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save"
         Height          =   400
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   840
         Width           =   1440
      End
      Begin VB.TextBox TxtHologramNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   47
         Top             =   360
         Width           =   2880
      End
      Begin VB.Label Label33 
         BackColor       =   &H008080FF&
         Caption         =   "Hologram No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   50
         Top             =   390
         Width           =   1260
      End
   End
   Begin VB.Frame Frame0 
      Height          =   2205
      Left            =   13080
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3235
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   22
         Top             =   150
         Width           =   1000
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   31
      Top             =   5520
      Width           =   12765
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   825
         Left            =   10800
         TabIndex        =   45
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   8925
         TabIndex        =   13
         Top             =   240
         Width           =   1875
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   6810
         TabIndex        =   17
         Top             =   240
         Width           =   2115
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   4575
         TabIndex        =   16
         Top             =   240
         Width           =   2235
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   2340
         TabIndex        =   12
         Top             =   240
         Width           =   2235
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   15
         Top             =   240
         Width           =   2235
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   8925
         TabIndex        =   14
         Top             =   660
         Width           =   1875
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   6810
         TabIndex        =   21
         Top             =   660
         Width           =   2115
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   4575
         TabIndex        =   20
         Top             =   660
         Width           =   2235
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   2340
         TabIndex        =   19
         Top             =   660
         Width           =   2235
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   100
         TabIndex        =   18
         Top             =   660
         Width           =   2235
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5505
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   12765
      Begin VB.TextBox TxtPreFix 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   10620
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1440
      End
      Begin VB.ComboBox CmbLocationID 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   4920
         Width           =   10200
      End
      Begin VB.CommandButton CmdLocationID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   12120
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   4920
         Width           =   375
      End
      Begin MSComCtl2.DTPicker TxtInstrumentDate 
         Height          =   375
         Left            =   6960
         TabIndex        =   6
         Top             =   1680
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49283073
         CurrentDate     =   39884
      End
      Begin VB.TextBox InstrumentAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   6960
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1275
         Width           =   2040
      End
      Begin VB.ComboBox CmbInstrumentType 
         Height          =   315
         ItemData        =   "FrmTxn_CashReceipt.frx":0000
         Left            =   1800
         List            =   "FrmTxn_CashReceipt.frx":000D
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1215
         Width           =   3120
      End
      Begin VB.TextBox TxtInstrumentNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1635
         Width           =   3120
      End
      Begin VB.TextBox TxtCRID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   3120
      End
      Begin VB.TextBox TxtRemark 
         Appearance      =   0  'Flat
         Height          =   840
         Left            =   1800
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3900
         Width           =   10500
      End
      Begin VB.TextBox TxtInvoiceDetail 
         Appearance      =   0  'Flat
         Height          =   840
         Left            =   1815
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   2955
         Width           =   10500
      End
      Begin VB.TextBox TxtDrawnOn 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   150
         TabIndex        =   7
         Top             =   2085
         Width           =   10260
      End
      Begin VB.TextBox TxtBranchName 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   150
         TabIndex        =   8
         Top             =   2520
         Width           =   10260
      End
      Begin VB.ComboBox CmbNonNcdexClient 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   750
         Width           =   10275
      End
      Begin VB.CommandButton CmdClient 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12120
         TabIndex        =   25
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox TxtClient 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   26
         Top             =   750
         Visible         =   0   'False
         Width           =   10260
      End
      Begin MSComCtl2.DTPicker TxtCRDate 
         Height          =   375
         Left            =   6960
         TabIndex        =   1
         Top             =   240
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49283073
         CurrentDate     =   39884
      End
      Begin VB.TextBox TxtLocationID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   4920
         Width           =   10560
      End
      Begin VB.Label Label9 
         Caption         =   "Prefix"
         Height          =   180
         Left            =   9840
         TabIndex        =   44
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label Label8 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   42
         Top             =   4950
         Width           =   1440
      End
      Begin VB.Label Label7 
         Caption         =   "Instrument Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5400
         TabIndex        =   39
         Top             =   1710
         Width           =   1800
      End
      Begin VB.Label Label6 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5400
         TabIndex        =   38
         Top             =   1365
         Width           =   1800
      End
      Begin VB.Label Label5 
         Caption         =   "Instrument No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   36
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label Label4 
         Caption         =   "Remark"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   35
         Top             =   4230
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Invoice Detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   34
         Top             =   3270
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "Instrument Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   33
         Top             =   1215
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Branch Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   32
         Top             =   2550
         Width           =   1800
      End
      Begin VB.Label LblFlag 
         Caption         =   "Receipt Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5400
         TabIndex        =   30
         Top             =   270
         Width           =   1800
      End
      Begin VB.Label LblStateID 
         Caption         =   "Drawn On"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   29
         Top             =   2190
         Width           =   1800
      End
      Begin VB.Label LblLocationName 
         Caption         =   "Client Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   28
         Top             =   735
         Width           =   1800
      End
      Begin VB.Label LblLocationID 
         Caption         =   "Receipt ID"
         Height          =   300
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   1800
      End
   End
End
Attribute VB_Name = "FrmTxn_CashReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mClientIDRow, mLocationId As Double

Private Sub CmbInstrumentType_LostFocus()
If CmbInstrumentType.Text = "" Then Exit Sub
CmbInstrumentType.Text = Trim(CmbInstrumentType.Text)

If LCase(CmbInstrumentType.Text) = "cash" Then
   TxtDrawnOn.Enabled = False
   TxtBranchName.Enabled = False
   TxtInstrumentNo.Enabled = False
   TxtInstrumentDate.Enabled = False
   TxtDrawnOn = ""
   TxtBranchName = ""
   TxtInstrumentNo = ""
'   TxtInstrumentDate = ""
   Exit Sub
End If
If LCase(CmbInstrumentType.Text) = "cheque" Or LCase(CmbInstrumentType.Text) = "dd" Then
   TxtDrawnOn.Enabled = True
   TxtBranchName.Enabled = True
   TxtInstrumentNo.Enabled = True
   TxtInstrumentDate.Enabled = True
   Exit Sub
End If

MsgBox "Invalid Selection !!!!"
CmbInstrumentType.SetFocus

End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location(" Where LocationId > 0 ")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid  selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationId = RsMst_Location!LocationID
TxtPreFix = GetPreFixByLocation(RsMst_Location!LocationID, "P")
End Sub


Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = 1")
CmbNonNcdexClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Client.RecordCount
    CmbNonNcdexClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbNonNcdexClient.Text = tmp
End Sub

Private Sub CmbNonNcdexClient_LostFocus()
If CmbNonNcdexClient.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbNonNcdexClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid  selection "
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientIDRow
Dim ans As Variant
ans = GetClientNameByRow(RsMst_Client!ClientIDRow)
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)

End Sub

Private Sub CmdCancel_Click()
TxtHologramNo = ""
Frame3.Visible = False
End Sub

Private Sub CmdClient_Click()
FrmMst_Client.Show
End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdPrint_Click()

If RsTxn_CashReceipt!HologramNo = "" Or IsNull(RsTxn_CashReceipt!HologramNo) Then
   Frame3.Visible = True
   TxtHologramNo.SetFocus
   Call TableMst_Hologram
   Exit Sub
Else
   TxtHologramNo = ""
   Frame3.Visible = False
End If
RsTxn_CashReceipt!Flag = "P"
RsTxn_CashReceipt.Update
Gen_mTimeStamp = GetTimeStamp
Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")
Call Xls_Print_Rpt

End Sub
Private Sub TxtHologramNo_LostFocus()
If TxtHologramNo = "" Then Exit Sub
RsMst_Hologram.MoveFirst
RsMst_Hologram.Find "HologramNo = '" & TxtHologramNo & "' "
If Not RsMst_Hologram.EOF Then
   
   If RsMst_Hologram!DocType = "WHR" Then
      tmp = "Select CMPID from Txn_WHR_Detail Where WHRNo = " & RsMst_Hologram!DocNo & ""
      Call Tmp4Table
      If TmpRs4.RecordCount > 0 Then
         MsgBox "Hologram Already Used In " & RsMst_Hologram!DocType & " " & TmpRs4!CMPID & "-" & (Val(RsMst_Hologram!DocNo) - TmpRs4!CMPID * 1000000)
         TmpRs4.Close
      End If
   Else
      MsgBox "Hologram Already Used In " & RsMst_Hologram!DocType & "-" & RsMst_Hologram!DocNo
   End If
   TxtHologramNo.SetFocus
   CmdSaveHologram.Enabled = False
   Exit Sub
   
End If
CmdSaveHologram.Enabled = True
CmdSaveHologram.SetFocus
End Sub
Private Sub CmdSaveHologram_Click()
Dim ans As Variant
If TxtHologramNo = "" Then
   MsgBox "Blank Hologram No not allowed!!!"
   TxtHologramNo.SetFocus
   Exit Sub
End If
ans = MsgBox("Are you Sure you want to save this HologramNo", vbYesNo)
If ans = vbYes Then
   'RsMst_Hologram.MoveLast
   RsMst_Hologram.AddNew
   RsMst_Hologram!HologramNo = TxtHologramNo
   RsMst_Hologram!DocType = "Cash Receipt"
   RsMst_Hologram!DocNo = TxtCRID
   RsMst_Hologram.Update
   'RsTxn_CashReceipt.Filter = " CRID = " & TxtCRID & " And SM_Prefix = '" & TxtPreFix & "' "
   RsTxn_CashReceipt!HologramNo = TxtHologramNo
   RsTxn_CashReceipt!Flag = "P"
   RsTxn_CashReceipt.Update
   'RsTxn_CashReceipt.Filter = ""
   
   TxtHologramNo = ""
   Frame3.Visible = False
   Gen_mTimeStamp = GetTimeStamp
   
   'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
   Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")
   Call Xls_Print_Rpt
   RsMst_Hologram.Close
Else
   TxtHologramNo.SetFocus
End If

End Sub
Public Sub Create_Xls(FileName As String)
'Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
'Dim Pat As String
Pat = App.Path
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
Gen_FLName = FileName
oBook.SaveAs (Pat & "\Excel\" & FileName)

Set oSheet = oBook.Worksheets(1)
With oExcel
'.Application.Sheets.Delete
.SheetsInNewWorkbook = 15 'MSHFlexGrid2.Rows - 2
.Workbooks.Add
End With
'oBook.SaveAs (Pat & "\" & FileName)

oExcel.Quit
'oExcel.Close
End Sub
Public Sub Xls_Print_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mRow As Integer
'Dim mCol As Integer
Dim ans As Variant
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_CashReceipt-" & Trim(TxtPreFix) & "" & Trim(TxtCRID) & ".xls")
 
i = 2
Set oWS = oWB.Worksheets("Sheet" & 1)
'oWS.PageSetup.RightMargin = 0.05
'oWS.PageSetup.TopMargin = 0.05
oWS.PageSetup.RightMargin = 0.05
oWS.PageSetup.TopMargin = 4.6
oWS.Name = TxtPreFix & "-" & TxtCRID  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 8

oWS.Range("a1:g1").Merge
oWS.Range("a:a").ColumnWidth = 10.86
oWS.Range("b:b").ColumnWidth = 12.86
oWS.Range("c:c").ColumnWidth = 12.86
oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 12.86 '32.14
oWS.Range("f:f").ColumnWidth = 12.86
oWS.Range("g:g").ColumnWidth = 12.86

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10

'oWS.Cells(1, 1).Font.ColorIndex = 4
mRow = mRow + 1

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Location -" & TxtLocationID
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Receipt No"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = TxtPreFix & "-" & TxtCRID
oWS.Cells(mRow, c + 2).Font.Size = 7
oWS.Cells(mRow, c + 2).HorizontalAlignment = -4108

oWS.Cells(mRow, c + 4) = "Receipt Date"
oWS.Cells(mRow, c + 4).Font.Size = 7
oWS.Cells(mRow, c + 4).Font.Bold = True

tmp = "e" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 5) = Format(TxtCRDate, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).HorizontalAlignment = -4108


mRow = mRow + 1

Dim Rsp As String
Rsp = NumStrConv(Val(InstrumentAmount))
tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Received with thanks from Shri " & TxtClient & " a sum of Rs " & InstrumentAmount & " ( Rupees " & Rsp & " )" & _
                         " towards the following by " & CmbInstrumentType & " "
'No:................. (instrument no) dated ................. ( instrument date) drawn on ............................ (drawn on Bank) ........................ branch (drawn on branch)."
If CmbInstrumentType.Text <> "Cash" Then
   oWS.Cells(mRow, c + 1) = oWS.Cells(mRow, c + 1) & " No: " & TxtInstrumentNo & " dated " & TxtInstrumentDate & " drawn on " & TxtDrawnOn & " " & TxtBranchName & " branch (subject to realisation). "
End If
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).HorizontalAlignment = 2
oWS.Cells(mRow, c + 1).Font.Size = 10
oWS.Cells(mRow, c + 1).RowHeight = 60

mRow = mRow + 1
'mCol = 5

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Cash Receipt Details"
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 10

mRow = mRow + 1
'mCol = 5




tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = TxtInvoiceDetail
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
'oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).RowHeight = 35

mRow = mRow + 1


'tmp = "a" & mRow & ":" & "a" & mRow + 1
'oWS.Range(tmp).Merge


tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge


tmp = "e" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 5) = "Affix Hologram here"
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 5).RowHeight = 45

mRow = mRow + 1



tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge


oWS.Cells(mRow, c + 1) = "Name and Signature of authorised representative of NCMSL"
oWS.Cells(mRow, c + 1).Font.Size = 7
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True


tmp = "e" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 5) = "Hologram No:- " & RsTxn_CashReceipt!HologramNo
oWS.Cells(mRow, c + 5).Font.Size = 7
oWS.Cells(mRow, c + 5).Font.Bold = True
oWS.Cells(mRow, c + 5).RowHeight = 40

mRow = mRow + 1
tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge


k = mRow

For Cell = 8 To k
    tmp = "a" & Cell & ":" & "g" & Cell
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    oWS.Range(tmp).Font.Name = "Zurich BT"
    'oWS.Range(tmp).Font.Size = 7
    oWS.Range(tmp).WrapText = True
Next Cell
If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If
oWB.Save

Call Log_Print_History(TxtPreFix, "CR", TxtCRID, Gen_FLName)

MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Call TableTxn_CashReceipt(" Order by SM_Prefix,CRID")
If RsTxn_CashReceipt.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
Call InData
End Sub
Private Sub CmdExcel_Click()
Call Gen_Create_Xls("Mst_Location.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Mst_Location.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub InstrumentAmount_LostFocus()
If InstrumentAmount = "" Then Exit Sub


If Not IsNumeric(InstrumentAmount) Then
   MsgBox "Invalid Number Format !!! "
   InstrumentAmount.SetFocus
   Exit Sub
End If

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CashReceipt.MoveNext
If RsTxn_CashReceipt.EOF Then
   RsTxn_CashReceipt.MoveLast
   LF_Flag = "Y"
End If
Call InData
If LF_Flag = "Y" Then
    NextCmd.Enabled = False
    LastCmd.Enabled = False
End If
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CashReceipt.MovePrevious
If RsTxn_CashReceipt.BOF Then
   RsTxn_CashReceipt.MoveFirst
   LF_Flag = "Y"
End If
Call InData
If LF_Flag = "Y" Then
    PreviousCmd.Enabled = False
    FirstCmd.Enabled = False
End If
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsTxn_CashReceipt.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_CashReceipt.MoveLast
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_CashReceipt.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Me.Height - 1000
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 11
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsTxn_CashReceipt.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsTxn_CashReceipt.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For i = 1 To RsTxn_CashReceipt.RecordCount
    For c = 0 To MSHFlexGrid1.Cols - 1
       'MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(rsTxn_CashReceipt.Fields(C)), "", rsTxn_CashReceipt.Fields(C))
        MSHFlexGrid1.TextMatrix(i, 0) = RsTxn_CashReceipt!CRID
        MSHFlexGrid1.TextMatrix(i, 1) = RsTxn_CashReceipt!SM_Prefix
        MSHFlexGrid1.TextMatrix(i, 2) = RsTxn_CashReceipt!CRDate
        MSHFlexGrid1.TextMatrix(i, 3) = GetClientName(RsTxn_CashReceipt!ClientIDRow)  ' !LocationId
        MSHFlexGrid1.TextMatrix(i, 4) = RsTxn_CashReceipt!InstrumentType
        MSHFlexGrid1.TextMatrix(i, 5) = RsTxn_CashReceipt!DrawnOn
        MSHFlexGrid1.TextMatrix(i, 6) = RsTxn_CashReceipt!BranchName
        MSHFlexGrid1.TextMatrix(i, 7) = RsTxn_CashReceipt!InstrumentNo
        MSHFlexGrid1.TextMatrix(i, 8) = RsTxn_CashReceipt!InstrumentDate
        MSHFlexGrid1.TextMatrix(i, 9) = RsTxn_CashReceipt!InstrumentAmount
        MSHFlexGrid1.TextMatrix(i, 10) = RsTxn_CashReceipt!Remark
        MSHFlexGrid1.TextMatrix(i, 11) = RsTxn_CashReceipt!InvoiceDetail

    Next
    RsTxn_CashReceipt.MoveNext
Next
End Sub
Public Sub InData()
Frame1.Enabled = False

CmdClient.Visible = False
CmbNonNcdexClient.Visible = False
TxtClient.Visible = True


CmbLocationID.Visible = False
CmdLocationID.Visible = False
TxtLocationID.Visible = True
CmbLocationID.Text = ""

TxtCRID = RsTxn_CashReceipt!CRID
TxtCRDate = RsTxn_CashReceipt!CRDate
mClientIDRow = RsTxn_CashReceipt!ClientIDRow
TxtClient = GetClientName(RsTxn_CashReceipt!ClientIDRow)
CmbInstrumentType = RsTxn_CashReceipt!InstrumentType
TxtDrawnOn = RsTxn_CashReceipt!DrawnOn
TxtBranchName = RsTxn_CashReceipt!BranchName
TxtInstrumentNo = RsTxn_CashReceipt!InstrumentNo
TxtInstrumentDate = RsTxn_CashReceipt!InstrumentDate
InstrumentAmount = RsTxn_CashReceipt!InstrumentAmount
TxtRemark = RsTxn_CashReceipt!Remark
TxtInvoiceDetail = RsTxn_CashReceipt!InvoiceDetail
TxtFlag = IIf(IsNull(RsTxn_CashReceipt!Flag), "", RsTxn_CashReceipt!Flag)
mLocationId = IIf(IsNull(RsTxn_CashReceipt!LocationID), 0, RsTxn_CashReceipt!LocationID)
TxtLocationID = GetLocationName(mLocationId)
TxtPreFix = RsTxn_CashReceipt!SM_Prefix

If RsTxn_CashReceipt.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
CmdPrint.Enabled = True
End Sub

Private Sub TxtLocationID_LostFocus()
'If TxtLocationID = "" Then Exit Sub
'If IsNumeric(TxtLocationID) = False Then
'   MsgBox "Invalid number format !!!!"
'   TxtLocationID.SetFocus
'   Exit Sub
'End If
'TxtLocationID = Val(TxtLocationID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame1.Enabled = True

TxtCRDate.Enabled = True
TxtInstrumentDate.Enabled = True
CmbInstrumentType.Visible = True

CmbNonNcdexClient.Visible = True
TxtClient.Visible = False
CmdClient.Visible = True


CmbLocationID.Visible = True
CmdLocationID.Visible = True
TxtLocationID.Visible = False
CmbLocationID.Text = ""

TxtCRID.Locked = True
TxtInstrumentNo.Locked = False
InstrumentAmount.Locked = False
TxtDrawnOn.Locked = False
TxtBranchName.Locked = False
TxtInvoiceDetail.Locked = False
TxtRemark.Locked = False

TxtCRID = ""
TxtInstrumentNo = ""
TxtCRDate = Date
TxtInstrumentDate = Date
CmbInstrumentType.Text = ""
CmbNonNcdexClient.Text = ""
InstrumentAmount = ""
TxtDrawnOn = ""
TxtBranchName = ""
TxtInvoiceDetail = ""
TxtRemark = ""
CmdPrint.Enabled = False
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
Frame1.Enabled = True

CmdClient.Visible = False
CmbNonNcdexClient.Visible = False
TxtClient.Visible = True
CmbNonNcdexClient.Text = TxtClient
CmdPrint.Enabled = False

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If ans = vbYes Then
   RsTxn_CashReceipt.Delete
   RsTxn_CashReceipt.Update
   If RsTxn_CashReceipt.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CashReceipt.MoveNext
   If RsTxn_CashReceipt.EOF() Then
      RsTxn_CashReceipt.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_CashReceipt.CancelUpdate
End Sub
Private Sub SaveCmd_Click()
If CmbNonNcdexClient.Text = "" Then
   MsgBox "Blank Client Name Not Allowed !!! "
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If

If CmbInstrumentType.Text = "" Then
   MsgBox "Blank Instrument Type Not Allowed !!! "
   CmbInstrumentType.SetFocus
   Exit Sub
End If

If Val(InstrumentAmount) <= 0 Then
   MsgBox "Blank Instrument Amount Not Allowed !!! "
   InstrumentAmount.SetFocus
   Exit Sub
End If

If TxtRemark.Text = "" Then
   MsgBox "Blank Remark Not Allowed !!! "
   TxtRemark.SetFocus
   Exit Sub
End If
If TxtInvoiceDetail.Text = "" Then
   MsgBox "Blank Invoice Detail Not Allowed !!! "
   TxtInvoiceDetail.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" And CmbLocationID.Text = "" Then
   MsgBox "Blank Location Name Not Allowed !!! "
   CmbLocationID.SetFocus
   Exit Sub
End If

If LCase(CmbInstrumentType.Text) <> "cash" Then
   If TxtDrawnOn.Text = "" Then
      MsgBox "Blank  Drawn On Not Allowed !!! "
      TxtDrawnOn.SetFocus
      Exit Sub
   End If
   If TxtBranchName.Text = "" Then
      MsgBox "Blank Branch Name Not Allowed !!! "
      TxtBranchName.SetFocus
      Exit Sub
   End If
   If TxtInstrumentNo.Text = "" Then
      MsgBox "Blank Instrument No Not Allowed !!! "
      TxtInstrumentNo.SetFocus
      Exit Sub
   End If
End If
If Edit_Flg = "A" Then
   RsTxn_CashReceipt.AddNew
   RsTxn_CashReceipt!CRID = GetCRID
   RsTxn_CashReceipt!G_UID = GetGUID
   TxtCRID.Text = RsTxn_CashReceipt!CRID
Else
    RsTxn_CashReceipt.MoveFirst
    RsTxn_CashReceipt.Find " CRID= " & TxtCRID.Text
End If

RsTxn_CashReceipt!CRID = TxtCRID
RsTxn_CashReceipt!CRDate = TxtCRDate.Value
RsTxn_CashReceipt!ClientIDRow = mClientIDRow
RsTxn_CashReceipt!InstrumentType = CmbInstrumentType.Text
RsTxn_CashReceipt!DrawnOn = TxtDrawnOn
RsTxn_CashReceipt!BranchName = TxtBranchName
RsTxn_CashReceipt!InstrumentNo = TxtInstrumentNo
RsTxn_CashReceipt!InstrumentDate = TxtInstrumentDate.Value
RsTxn_CashReceipt!InstrumentAmount = Val(InstrumentAmount)
RsTxn_CashReceipt!Remark = TxtRemark
RsTxn_CashReceipt!InvoiceDetail = TxtInvoiceDetail
RsTxn_CashReceipt!Flag = "0" 'TxtFlag.Text
RsTxn_CashReceipt!SM_Prefix = TxtPreFix  'TxtFlag.Text
 RsTxn_CashReceipt!LocationID = mLocationId
RsTxn_CashReceipt.Update
Call InData
End Sub
Private Function GetCRID() As Double
Dim RetVal As Double
tmp = "Select max(CRID) from Txn_CashReceipt where sm_prefix = '" & TxtPreFix & "'"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetCRID = RetVal
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 12
MSHFlexGrid1.TextMatrix(0, 0) = "Receipt ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Prefix"
MSHFlexGrid1.TextMatrix(0, 2) = "Receipt Date"
MSHFlexGrid1.TextMatrix(0, 3) = "client Name"
MSHFlexGrid1.TextMatrix(0, 4) = "Type"
MSHFlexGrid1.TextMatrix(0, 5) = "Drawn On"
MSHFlexGrid1.TextMatrix(0, 6) = "Branch Name"
MSHFlexGrid1.TextMatrix(0, 7) = "Instrument No"
MSHFlexGrid1.TextMatrix(0, 8) = "Instrument Date"
MSHFlexGrid1.TextMatrix(0, 9) = "Instrument Amount"
MSHFlexGrid1.TextMatrix(0, 10) = "Remark"
MSHFlexGrid1.TextMatrix(0, 11) = "Invoice Detail"


End Sub

Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsTxn_CashReceipt.MoveFirst
RsTxn_CashReceipt.Find "SM_Prefix = '" & MSHFlexGrid1.TextMatrix(i, 1) & "'"
If Not RsTxn_CashReceipt.EOF Then
   RsTxn_CashReceipt.Find "CRID = " & MSHFlexGrid1.TextMatrix(i, 0)
   Call InData
   Frame0.Visible = False
End If
End Sub



Private Sub TxtLocationName_LostFocus()
TxtLocationName = Trim(TxtLocationName)
End Sub

