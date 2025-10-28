VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_GodownwiseNCDEXStock 
   Caption         =   "Godown wise NCDEX Stock"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   10125
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   8055
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7665
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13520
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
   Begin VB.Frame Frame2 
      Height          =   1605
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14985
      Begin VB.ComboBox CmbSExchangeType 
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
         Left            =   105
         TabIndex        =   7
         Top             =   1170
         Width           =   2835
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
         Left            =   13485
         TabIndex        =   10
         Top             =   1035
         Width           =   1215
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
         Left            =   11055
         TabIndex        =   8
         Top             =   1035
         Width           =   1215
      End
      Begin VB.ComboBox CmbSCmp 
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
         Left            =   8640
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   525
         Width           =   3885
      End
      Begin VB.ComboBox CmbSGodown 
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
         Left            =   12540
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   525
         Width           =   2205
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
         Left            =   12270
         TabIndex        =   9
         Top             =   1035
         Width           =   1215
      End
      Begin VB.ComboBox CmbPreFix 
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
         Left            =   5685
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   525
         Width           =   2925
      End
      Begin VB.ComboBox CmbStateID 
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
         Height          =   360
         ItemData        =   "FrmRpt_GodownwiseNCDEXStock.frx":0000
         Left            =   3015
         List            =   "FrmRpt_GodownwiseNCDEXStock.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   525
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   525
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   22478849
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   360
         Left            =   1590
         TabIndex        =   2
         Top             =   525
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   22478849
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.Label Label5 
         Caption         =   "Exchange Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   690
         TabIndex        =   21
         Top             =   915
         Width           =   1665
      End
      Begin VB.Label Label7 
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
         Left            =   5730
         TabIndex        =   18
         Top             =   1185
         Width           =   2415
      End
      Begin VB.Label Label1 
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
         Left            =   8160
         TabIndex        =   17
         Top             =   1200
         Width           =   90
      End
      Begin VB.Label Label2 
         Caption         =   "CMP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10335
         TabIndex        =   16
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Godown"
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
         Left            =   13230
         TabIndex        =   15
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label8 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6720
         TabIndex        =   14
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label21 
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4065
         TabIndex        =   13
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "From  Date"
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
         Left            =   315
         TabIndex        =   12
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "To  Date"
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
         Left            =   1875
         TabIndex        =   11
         Top             =   240
         Width           =   780
      End
   End
End
Attribute VB_Name = "FrmRpt_GodownwiseNCDEXStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID As Variant
Dim mCMPID, mCommodityID, mGodownID, mExchangeTypeID As Double
Dim msSM_prefix, mSLocationID As Variant
Dim mBGSLIDs As String
Private Sub CmbPreFix_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Plase Select State !!!"
   CmbStateID.SetFocus
   Exit Sub
End If

tmp = CmbPreFix.Text
'Call TableMst_StoreManLocations("where " & Gen_WcSMLocation)

If Gen_WcLocation = "" Then
   Call TableMst_Location(" where StateID=" & mStateID)
Else
   Call TableMst_Location("where StateID=" & mStateID & " and " & Gen_WcSMLocation)
End If


CmbPreFix.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbPreFix.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbPreFix.AddItem RsMst_Location!LocationName

    RsMst_Location.MoveNext
Next
CmbPreFix.Text = tmp

End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbPreFix.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid PreFix Selection !!!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
If tmp <> CmbPreFix.Text Then
   CmbSCmp.Text = ""
   CmbSGodown.Text = ""
End If
msSM_prefix = RsMst_Location!SM_Prefix
mSLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSCmp_GotFocus()
If CmbPreFix.Text = "" Then
   MsgBox "Plase Select Location !!!"
   CmbPreFix.SetFocus
   Exit Sub
End If

tmp1 = ""
If Gen_WcCMP <> "" Then
   tmp1 = "And " & Gen_WcCMP
End If

tmp = CmbSCmp.Text
Call TableMst_CMP(" Where LocationID = " & mSLocationID & "" & tmp1)
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
    CmbSCmp.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSCmp.Text = tmp

End Sub

Private Sub CmbSCmp_LostFocus()
If CmbSCmp.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCmp.SetFocus
   Exit Sub
End If
If tmp <> CmbSCmp.Text Then
   CmbSGodown.Text = ""
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbSExchangeType_GotFocus()
tmp = CmbSExchangeType.Text
Call TableMst_ExchangeType
CmbSExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_ExchangeType.RecordCount
    CmbSExchangeType.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbSExchangeType.Text = tmp
End Sub

Private Sub CmbSExchangeType_LostFocus()
If CmbSExchangeType.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbSExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbSExchangeType.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID

End Sub


Private Sub CmbSGodown_GotFocus()
If CmbSCmp.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbSCmp.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown("Where CMPID = " & mCMPID)
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Godown.RecordCount
    CmbSGodown.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbSGodown.Text = tmp
End Sub

Private Sub CmbSGodown_LostFocus()
If CmbSGodown.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbSGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbSGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_State.RecordCount
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
If tmp <> CmbStateID.Text Then
   CmbPreFix.Text = ""
   CmbSCmp.Text = ""
   CmbSGodown.Text = ""
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmdClear_Click()
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
CmbStateID.Text = ""
CmbPreFix.Text = ""
CmbSCmp.Text = ""
CmbSGodown.Text = ""
Label1.Caption = "-"
CmdExcel.Enabled = False
Call Grid_Head
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownwiseNCDEXStock.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GodownwiseNCDEXStock.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 2
tmp = "a" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 1) = "CMP Name : " & CmbSCmp.Text

mRow = mRow + 1
tmp = "a" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Godown No : " & CmbSGodown.Text

mRow = mRow + 1
tmp = "a" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "From Date :" & Format(TxtFromDate.Value, "dd-MMM-yyyy") & " To Date : " & Format(TxtToDate.Value, "dd-MMM-yyyy")

mRow = mRow + 2

For i = 0 To MSHFlexGrid1.Rows - 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MMM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
        
        
    Next
    mRow = mRow + 1
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True

End Sub


Private Sub CmdGetReport_Click()

Dim mTotDays As Double
Dim mDate As Date

If TxtFromDate > TxtToDate Then
   MsgBox "FromDate Should be Less Than Todate."
   TxtFromDate.SetFocus
   Exit Sub
End If
If CmbStateID.Text = "" Then
   MsgBox "Select State!!"
   CmbStateID.SetFocus
   Exit Sub
End If

If CmbPreFix.Text = "" Then
   MsgBox "Select Location!!"
   CmbPreFix.SetFocus
   Exit Sub
End If

If CmbSCmp.Text = "" Then
   MsgBox "Select CMP!!"
   CmbSCmp.SetFocus
   Exit Sub
End If

If CmbSGodown.Text = "" Then
   MsgBox "Select Godown !!"
   CmbSGodown.SetFocus
   Exit Sub
End If

If CmbSExchangeType.Text = "" Then
   MsgBox "Select exchange !!"
   CmbSExchangeType.SetFocus
   Exit Sub
End If

Call Grid_Head

MsgBox "Please wait till next Message!!"

tmp = "Select MS.StateID , MS.StateName, ML.LocationID, ML.LocationName,MC.CMPID,MC.CMPName,MG.GodownID,MG.GodownNo"
tmp = tmp & " ,MG.NCDEXAccreditedQty,MG.GodownAlias,MG.GodownCapacity,UT.UnitType,MG.TelephoneNo,MG.StartDate"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Inner join mst_UnitType UT on MG.UnitTypeID=UT.UnitTypeID"
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " where MG.GodownID=" & mGodownID

Call TmpTable

mTotDays = DateDiff("d", TxtFromDate.Value, TxtToDate.Value)
mDate = TxtFromDate.Value

tmp = "select sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.gslid = b.gslid"
'tmp = tmp & " where a.exchangetypeid = 2 And a.TxnType='D' And a.CxTFDate<'" & TxtFromDate & "' And C.GodownID=" & mGodownID & " And c.SM_Prefix='" & msSM_prefix & "'"
tmp = tmp & " where a.exchangetypeid = " & mExchangeTypeID & " And a.TxnType='D' And a.CxTFDate<'" & TxtFromDate & "' And C.GodownID=" & mGodownID & " And c.SM_Prefix='" & msSM_prefix & "'"

Call Tmp1Table

tmp = "select sum(b.NoOfBags) as NoOfBags ,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.gslid = b.gslid"
'tmp = tmp & " where a.exchangetypeid = 2 And a.TxnType='W' And a.CxTFDate <'" & TxtFromDate & "' And C.GodownID=" & mGodownID & " And c.SM_Prefix='" & msSM_prefix & "'"
tmp = tmp & " where a.exchangetypeid = " & mExchangeTypeID & " And a.TxnType='W' And a.CxTFDate <'" & TxtFromDate & "' And C.GodownID=" & mGodownID & " And c.SM_Prefix='" & msSM_prefix & "'"

Call Tmp2Table

tmp = "select a.TxnType,a.CxTFDate,sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.gslid = b.gslid"
'tmp = tmp & " where a.exchangetypeid =2 And a.CxTFDate>='" & TxtFromDate & "' And  a.CxTFDate<='" & TxtToDate & "'  And C.GodownID=" & mGodownID & "  And c.SM_Prefix='" & msSM_prefix & "'"

tmp = tmp & " where a.exchangetypeid = " & mExchangeTypeID & " And a.CxTFDate>='" & TxtFromDate & "' And  a.CxTFDate<='" & TxtToDate & "'  And C.GodownID=" & mGodownID & "  And c.SM_Prefix='" & msSM_prefix & "'"
tmp = tmp & " group by a.CxTFDate,a.TxnType"
tmp = tmp & " order by  a.CxTFDate,a.TxnType"

Call Tmp3Table

mOpDepWt = 0
mOpWidWt = 0

If TmpRs1.RecordCount > 0 Then
   mOpDepWt = IIf(IsNull(TmpRs1.Fields(1)), 0, TmpRs1.Fields(1))
End If
If TmpRs2.RecordCount > 0 Then
   mOpWidWt = IIf(IsNull(TmpRs2.Fields(1)), 0, TmpRs2.Fields(1))
End If


For i = 0 To mTotDays
    r = MSHFlexGrid1.Rows - 1
    MSHFlexGrid1.TextMatrix(r, 0) = Format(mDate, "dd-MMM-yyyy")
    MSHFlexGrid1.TextMatrix(r, 1) = TmpRs!StateName
    MSHFlexGrid1.TextMatrix(r, 2) = TmpRs!LocationName
    MSHFlexGrid1.TextMatrix(r, 3) = TmpRs!CMPName
    MSHFlexGrid1.TextMatrix(r, 4) = TmpRs!TelephoneNo
    MSHFlexGrid1.TextMatrix(r, 5) = TmpRs!GodownNo
    MSHFlexGrid1.TextMatrix(r, 6) = Format(TmpRs!StartDate, "dd-MMM-yyyy")
    MSHFlexGrid1.TextMatrix(r, 7) = TmpRs!GodownCapacity
    MSHFlexGrid1.TextMatrix(r, 8) = TmpRs!UnitType
    MSHFlexGrid1.TextMatrix(r, 9) = TmpRs!NCDEXAccreditedQty
    
    If TmpRs3.RecordCount > 0 Then TmpRs3.MoveFirst
    TmpRs3.Find "CxTFDate='" & mDate & "'"
    If TmpRs3.EOF = False Then
       If TmpRs3!TxnType = "D" Then
          MSHFlexGrid1.TextMatrix(r, 12) = IIf(IsNull(TmpRs3!GrossWeight), 0, TmpRs3!GrossWeight)
       Else
          MSHFlexGrid1.TextMatrix(r, 15) = IIf(IsNull(TmpRs3!GrossWeight), 0, TmpRs3!GrossWeight)
       End If
    End If
    
    MSHFlexGrid1.TextMatrix(r, 13) = TransferReceived(mDate)
    MSHFlexGrid1.TextMatrix(r, 14) = TransferIssued(mDate)
    
    If r = 2 Then
       MSHFlexGrid1.TextMatrix(r, 11) = Format((mOpDepWt + OpeningTransferReceived()) - (mOpWidWt + OpeningTransferIssued()), "#####0.000")
       'MSHFlexGrid1.TextMatrix(r, 13) = OpeningTransferReceived()
       'MSHFlexGrid1.TextMatrix(r, 14) = OpeningTransferIssued()
       MSHFlexGrid1.TextMatrix(r, 10) = GetMonthlyRent()
    Else
       MSHFlexGrid1.TextMatrix(r, 11) = MSHFlexGrid1.TextMatrix(r - 1, 16)
       MSHFlexGrid1.TextMatrix(r, 10) = MSHFlexGrid1.TextMatrix(r - 1, 10)
    End If
    
    MSHFlexGrid1.TextMatrix(r, 16) = Format(Val(MSHFlexGrid1.TextMatrix(r, 11)) + Val(MSHFlexGrid1.TextMatrix(r, 12)) + Val(MSHFlexGrid1.TextMatrix(r, 13)) - (Val(MSHFlexGrid1.TextMatrix(r, 14)) + Val(MSHFlexGrid1.TextMatrix(r, 15))), "#####0.000")
    
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    mDate = mDate + 1
Next
Label1.Caption = mTotDays
Label1.Refresh
CmdExcel.Enabled = True
MsgBox "Done!!"

End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.Cols = 17

MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(1, 0) = "Date"
MSHFlexGrid1.TextMatrix(1, 1) = "State"
MSHFlexGrid1.TextMatrix(1, 2) = "Location"
MSHFlexGrid1.TextMatrix(1, 3) = "CMP Name"
MSHFlexGrid1.TextMatrix(1, 4) = "NCDEX Godown ID"
MSHFlexGrid1.TextMatrix(1, 5) = "Godown No"
MSHFlexGrid1.TextMatrix(1, 6) = "Godown Start Date"
MSHFlexGrid1.TextMatrix(1, 7) = "Godown Capacity"
MSHFlexGrid1.TextMatrix(1, 8) = "Godown Type"
MSHFlexGrid1.TextMatrix(1, 9) = "NCDEX Accredited Qty"
MSHFlexGrid1.TextMatrix(1, 10) = "Monthly Godown Rent Amount"
MSHFlexGrid1.TextMatrix(1, 11) = "Opening "
MSHFlexGrid1.TextMatrix(1, 12) = "Deposited "
MSHFlexGrid1.TextMatrix(1, 13) = "Transfer Received"
MSHFlexGrid1.TextMatrix(1, 14) = "Transfer Issued"
MSHFlexGrid1.TextMatrix(1, 15) = "Withdrawn"
MSHFlexGrid1.TextMatrix(1, 16) = "Balance"

MSHFlexGrid1.ColWidth(0) = 1800
MSHFlexGrid1.ColWidth(1) = 1800
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 1200
MSHFlexGrid1.ColWidth(9) = 1600
MSHFlexGrid1.ColWidth(10) = 2000

MSHFlexGrid1.RowHeight(1) = 500

End Sub
Private Function GetMonthlyRent() As Double
Dim Retval As Double
Retval = 0

tmp = "Select TGL.GodownID,sum(TGL.RentAmount)RentAmount"
tmp = tmp & " From Txn_GodownLessorAgreeDetail TGL"
tmp = tmp & " inner join Txn_LessorAgreement TLA on TGL.LAID=TLA.LAID"
tmp = tmp & " Inner Join Mst_Godown MG on TGL.GodownID=MG.GodownID"
tmp = tmp & " Where TLA.StartDate <= '" & TxtToDate.Value & "' And TLA.ExpiryDate>='" & TxtFromDate.Value & "' And (MG.CloseDate >= '" & TxtFromDate.Value & "' or MG.CloseDate is null) And TLA.Flag<>'R'"
tmp = tmp & " And MG.GodownID=" & mGodownID
tmp = tmp & " Group By TGL.GodownID"
Call Tmp5Table
If TmpRs5.RecordCount > 0 Then
   Retval = IIf(IsNull(TmpRs5!RentAmount), 0, TmpRs5!RentAmount)
End If
GetMonthlyRent = Retval
End Function

Private Function FillGodownRent() As Double
Dim mRentDays, mTotRentAmt, mRentAmt, mTotRent, mTotMonths As Double
Dim mRentStartDate, mRentEndDate As Variant
Dim mSStartDate, mSEndDate As Variant
Dim mRentStart, mRentEnd As Variant
Dim mPaidRent  As Double
Dim mRow As Integer
'Dim mTotRent As Double

tmp = "Select TGL.GodownID,TGL.LAID,TGL.RentAmount,TLA.StartDate,TLA.ExpiryDate,MG.CloseDate,MG.StartDate as GodownStartDate,TLA.Flag"
tmp = tmp & " From Txn_GodownLessorAgreeDetail TGL"
tmp = tmp & " inner join Txn_LessorAgreement TLA on TGL.LAID=TLA.LAID"
tmp = tmp & " Inner Join Mst_Godown MG on TGL.GodownID=MG.GodownID"
tmp = tmp & " Where TLA.StartDate <= '" & TxtToDate.Value & "' And TLA.ExpiryDate>='" & TxtFromDate.Value & "' And (MG.CloseDate >= '" & TxtFromDate.Value & "' or MG.CloseDate is null) And TLA.Flag<>'R'"
tmp = tmp & " And MG.GodownID=" & mGodownID

Call Tmp5Table
mTotRent = 0
If TmpRs5.RecordCount > 0 Then
   mTotRentAmt = 0
   mRow = 2
   TmpRs5.MoveFirst
   For r = 1 To TmpRs5.RecordCount
        mRentAmt = 0
        mTotMonths = DateDiff("m", TxtFromDate.Value, TxtToDate.Value)
        For i = 0 To mTotMonths
            mCurrMonth = DateAdd("m", i, mStartDate)
            mYear = Year(mCurrMonth)
            mMonthNo = Month(mCurrMonth)
            mTotDays = Day(DateSerial(mYear, mMonthNo + 1, 0))
            mSStartDate = mYear & "-" & mMonthNo & "-" & "01"
            mSEndDate = mYear & "-" & mMonthNo & "-" & mTotDays
            mRentStartDate = IIf(TmpRs5!StartDate < TmpRs5!GodownStartDate, TmpRs5!GodownStartDate, TmpRs5!StartDate)
            
            If IsNull(TmpRs5!CloseDate) = False Then
               mRentEndDate = IIf(TmpRs5!ExpiryDate > TmpRs5!CloseDate, TmpRs5!CloseDate, TmpRs5!ExpiryDate)
            Else
               mRentEndDate = TmpRs5!ExpiryDate
            End If
            If mTotDays <> 0 Then
               
               mTotRent = TmpRs5!RentAmount / mTotDays
               
               If mRentStartDate < CDate(mSStartDate) And mRentEndDate > CDate(mSEndDate) Then
                  mRentDays = DateDiff("d", CDate(mSStartDate), CDate(mSEndDate))
                  mRentStart = mSStartDate
                  mRentEnd = mSEndDate
               
               ElseIf mRentStartDate < CDate(mSStartDate) And mRentEndDate <= CDate(mSEndDate) And mRentEndDate >= CDate(mSStartDate) Then
                  mRentDays = DateDiff("d", CDate(mSStartDate), mRentEndDate)
                  mRentStart = mSStartDate
                  mRentEnd = mRentEndDate
               ElseIf mRentStartDate < CDate(mSStartDate) And mRentEndDate <= CDate(mSEndDate) Then
                  mRentDays = 0
                  mRentStart = mSStartDate
                  mRentEnd = mSEndDate
                  
               ElseIf mRentStartDate >= CDate(mSEndDate) And mRentEndDate >= CDate(mSStartDate) Then  ' And mRentEndDate > CDate(mSEndDate) Then
                  mRentDays = 0
                  mRentStart = mSStartDate
                  mRentEnd = mSEndDate
               ElseIf mRentStartDate <= CDate(mSEndDate) And mRentEndDate > CDate(mSEndDate) Then
                  mRentDays = DateDiff("d", mRentStartDate, CDate(mSEndDate))
                  mRentStart = mRentStartDate
                  mRentEnd = mSEndDate
               Else
                  mRentDays = mTotDays 'DateDiff("d", mRentStartDate, mRentEndDate)
               End If
               If mRentDays <> 0 Then
                  mRentDays = mRentDays + 1
                  mRentAmt = mRentAmt + mTotRent * mRentDays
               End If
            End If
            mTotRent = mTotRent + Format(mTotRent * mRentDays, "#####0.00")
        Next
        TmpRs5.MoveNext
    Next
End If
FillGodownRent = mTotRent
End Function
Private Function OpeningTransferReceived() As Variant
Dim mDepStock, mWidStock, mStock As Variant

Dim mGSLIDs As String
mGSLIDs = GetBGSLIDs()


tmp = "Select GSLTransferDate,a.TxnType,sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.gslid = b.gslid"
tmp = tmp & " inner join Txn_GSL_Transfer d on  c.SM_Prefix =d.SM_Prefix and c.gslid = d.oldGSLID"
tmp = tmp & " where a.exchangetypeid =2 And d.GSLTransferDate<'" & TxtFromDate.Value & "' "
tmp = tmp & " And c.GSLID in (" & mGSLIDs & ")"
tmp = tmp & " And c.GodownID<>" & mGodownID & " And c.SM_Prefix='" & msSM_prefix & "'"
tmp = tmp & " group by d.GSLTransferDate,a.TxnType"

Call Tmp5Table

mStock = 0
mWidStock = 0
mDepStock = 0
If TmpRs5.RecordCount > 0 Then
   For i = 1 To TmpRs5.RecordCount
       If TmpRs5!TxnType = "D" Then
          mDepStock = TmpRs5!GrossWeight
       Else
          mWidStock = TmpRs5!GrossWeight
       End If
       TmpRs5.MoveNext
   Next
End If
mStock = Format(mDepStock - mWidStock, "#####0.000")
OpeningTransferReceived = mStock
End Function
Private Function OpeningTransferIssued() As Variant
Dim mDepStock, mWidStock, mStock As Variant

tmp = "Select GSLTransferDate,a.TxnType,sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.gslid = b.gslid"
tmp = tmp & " inner join Txn_GSL_Transfer d on  c.SM_Prefix =d.SM_Prefix and c.gslid = d.oldGSLID"
tmp = tmp & " where a.exchangetypeid =2 And d.GSLTransferDate<'" & TxtFromDate.Value & "' "
tmp = tmp & " And c.GSLID in ( select GSLID from  mst_gsl C where c.exchangetypeid =2 And C.GodownID=" & mGodownID & " And FGSLID is not null)"
tmp = tmp & " And c.GodownID<>" & mGodownID & " And c.SM_Prefix='" & msSM_prefix & "'"
tmp = tmp & " group by d.GSLTransferDate,a.TxnType"

Call Tmp5Table

mStock = 0
mWidStock = 0
mDepStock = 0
If TmpRs5.RecordCount > 0 Then
   For i = 1 To TmpRs5.RecordCount
       If TmpRs5!TxnType = "D" Then
          mDepStock = TmpRs5!GrossWeight
       Else
          mWidStock = TmpRs5!GrossWeight
       End If
       TmpRs5.MoveNext
   Next
End If
mStock = Format(mDepStock - mWidStock, "#####0.000")
OpeningTransferIssued = mStock
End Function
Private Function TransferReceived(mDate_ As Variant) As Variant
Dim mDepStock, mWidStock, mStock As Variant

Dim mGSLIDs As String
mGSLIDs = GetBGSLIDs()

tmp = "Select GSLTransferDate,a.TxnType,sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.gslid = b.gslid"
tmp = tmp & " inner join Txn_GSL_Transfer d on  c.SM_Prefix =d.SM_Prefix and c.gslid = d.oldGSLID"
'tmp = tmp & " where a.exchangetypeid =2 And d.GSLTransferDate>='" & mDate_ & "' And d.GSLTransferDate<='" & mDate_ & "'"
tmp = tmp & " where a.exchangetypeid = " & mExchangeTypeID & " And d.GSLTransferDate>='" & mDate_ & "' And d.GSLTransferDate<='" & mDate_ & "'"
tmp = tmp & " And c.GSLID in (" & mGSLIDs & ")"
tmp = tmp & " And c.GodownID<>" & mGodownID & " And c.SM_Prefix='" & msSM_prefix & "'"
tmp = tmp & " group by d.GSLTransferDate,a.TxnType"

Call Tmp5Table

mStock = 0
mWidStock = 0
mDepStock = 0
If TmpRs5.RecordCount > 0 Then
   For i = 1 To TmpRs5.RecordCount
       If TmpRs5!TxnType = "D" Then
          mDepStock = TmpRs5!GrossWeight
       Else
          mWidStock = TmpRs5!GrossWeight
       End If
       TmpRs5.MoveNext
   Next
End If
mStock = Format(mDepStock - mWidStock, "#####0.000")
TransferReceived = mStock
End Function

Private Function TransferIssued(mDate_ As Variant) As Variant
Dim mDepStock, mWidStock, mStock As Variant
If mDate_ = "11/16/2011" Then
  m = m
End If

tmp = "Select GSLTransferDate,a.TxnType,sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
tmp = tmp & " from txn_cxtf_details a"
tmp = tmp & " inner join Txn_CxTF_GSL b on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
tmp = tmp & " inner join mst_gsl c on  c.SM_Prefix = b.SM_Prefix and c.gslid = b.gslid"
tmp = tmp & " inner join Txn_GSL_Transfer d on  c.SM_Prefix =d.SM_Prefix and c.gslid = d.oldGSLID"
tmp = tmp & " inner join mst_gsl c1 on  c1.SM_Prefix = d.SM_Prefix and c1.gslid = d.newgslid"
'tmp = tmp & " where a.exchangetypeid =2 And d.GSLTransferDate>='" & mDate_ & "' And d.GSLTransferDate<='" & mDate_ & "'"
tmp = tmp & " where a.exchangetypeid = " & mExchangeTypeID & " And d.GSLTransferDate>='" & mDate_ & "' And d.GSLTransferDate<='" & mDate_ & "'"
tmp = tmp & " And c.GSLID in ( select GSLID from  mst_gsl C where c.exchangetypeid =2 And C.GodownID=" & mGodownID & " And FGSLID is not null)"
tmp = tmp & " And c1.GodownID<>" & mGodownID & " And c.SM_Prefix='" & msSM_prefix & "'"
tmp = tmp & " group by d.GSLTransferDate,a.TxnType"

Call Tmp5Table

mStock = 0
mWidStock = 0
mDepStock = 0
If TmpRs5.RecordCount > 0 Then
   For i = 1 To TmpRs5.RecordCount
       If TmpRs5!TxnType = "D" Then
          mDepStock = TmpRs5!GrossWeight
       Else
          mWidStock = TmpRs5!GrossWeight
       End If
       TmpRs5.MoveNext
   Next
End If
mStock = Format(mDepStock - mWidStock, "#####0.000")
TransferIssued = mStock

End Function
Private Function GetBGSLIDs() As String
mBGSLIDs = ""
tmp = "select distinct GSLID from  mst_gsl C where c.exchangetypeid =2 And C.GodownID=" & mGodownID & ""
Call GetBGSLID(tmp)
If mBGSLIDs = "" Then mBGSLIDs = 0
GetBGSLIDs = mBGSLIDs
End Function

Private Sub GetBGSLID(mGSLID_ As Variant)
Dim mTempBGSLIDs As String
tmp = "select distinct C.BGSLID from  mst_gsl C where C.GSLID in (" & mGSLID_ & ") And SM_Prefix='" & msSM_prefix & "'"
Call Tmp0Table
mTempBGSLIDs = ""
If TmpRs0.RecordCount > 0 Then
   For i = 1 To TmpRs0.RecordCount
       If IsNull(TmpRs0!BGSLID) = False Then
          If mTempBGSLIDs = "" Then
             mTempBGSLIDs = TmpRs0!BGSLID
          Else
             mTempBGSLIDs = mTempBGSLIDs & "," & TmpRs0!BGSLID
          End If
       End If
       TmpRs0.MoveNext
   Next
   If mTempBGSLIDs <> "" Then
      If mBGSLIDs = "" Then
         mBGSLIDs = mTempBGSLIDs
      Else
         mBGSLIDs = mBGSLIDs & "," & mTempBGSLIDs
      End If
      Call GetBGSLID(mTempBGSLIDs)
   End If
End If

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True

Call CmdClear_Click
End Sub
