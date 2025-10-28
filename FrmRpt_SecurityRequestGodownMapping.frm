VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_SecurityRequestGodownMapping 
   Caption         =   "Security Request Godown Mapping"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   60
      TabIndex        =   8
      Top             =   45
      Width           =   15135
      Begin VB.OptionButton OptFreeSecurity 
         Caption         =   "Security Not mapped to any live Godown"
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
         Left            =   180
         TabIndex        =   18
         Top             =   840
         Width           =   2325
      End
      Begin VB.ComboBox CmbSecurityDesignation 
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
         Left            =   10680
         Sorted          =   -1  'True
         TabIndex        =   16
         Top             =   525
         Width           =   2895
      End
      Begin VB.OptionButton OptGodMapped 
         Caption         =   "Security Details with Godown"
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
         Left            =   180
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   2955
      End
      Begin VB.OptionButton OptGodNotMapped 
         Caption         =   "Godown Not Mapped"
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
         Left            =   180
         TabIndex        =   14
         Top             =   600
         Width           =   2325
      End
      Begin VB.ComboBox CmbSecurityType 
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
         Left            =   6240
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   525
         Width           =   4335
      End
      Begin VB.ComboBox CmbSecurityAgency 
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
         Left            =   4620
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   7515
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
         Height          =   360
         Left            =   10800
         TabIndex        =   4
         Top             =   1440
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
         Height          =   360
         Left            =   12135
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
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
         Height          =   360
         Left            =   13470
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
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
         Left            =   3060
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   525
         Width           =   3045
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Security Designation"
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
         Left            =   11205
         TabIndex        =   17
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Security Type"
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
         Left            =   7785
         TabIndex        =   13
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Security Agency"
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
         Left            =   3060
         TabIndex        =   12
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label LblTotalRecord 
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
         Left            =   9600
         TabIndex        =   11
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Record Found :"
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
         Left            =   7200
         TabIndex        =   10
         Top             =   1560
         Width           =   2205
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   4200
         TabIndex        =   9
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8055
      Left            =   60
      TabIndex        =   0
      Top             =   1965
      Width           =   15135
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7695
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   13573
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
Attribute VB_Name = "FrmRpt_SecurityRequestGodownMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mSecurityTypeID, mSecurityDesignationID, mSecurityAgencyID As Variant

Private Sub CmbSecurityAgency_GotFocus()
tmp = CmbSecurityAgency.Text
Call TableMst_SecurityAgency("Where Flag = 'Y'")
CmbSecurityAgency.Clear
If RsMst_SecurityAgency.RecordCount = 0 Then
   MsgBox "No Security Agency Found!!!"
   CmbSecurityAgency.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityAgency.RecordCount
    CmbSecurityAgency.AddItem RsMst_SecurityAgency!AgencyName
    RsMst_SecurityAgency.MoveNext
Next
CmbSecurityAgency.Text = tmp
End Sub

Private Sub CmbSecurityAgency_LostFocus()
If CmbSecurityAgency.Text = "" Then Exit Sub

RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "AgencyName = '" & CmbSecurityAgency.Text & "'"

If RsMst_SecurityAgency.EOF Then
   MsgBox "Invalid Selection!!!"
   CmbSecurityAgency.SetFocus
   Exit Sub
End If
mSecurityAgencyID = RsMst_SecurityAgency!SecurityAgencyID

End Sub

Private Sub CmdClear_Click()
Call Grid_Head
CmdExcel.Enabled = False
CmbLocation.Text = ""
CmbSecurityAgency.Text = ""
CmbSecurityDesignation.Text = ""
CmbSecurityType.Text = ""
LblTotalRecord.Caption = "-"
End Sub

Private Sub CmdExcel_Click()
'If MSHFlexGrid1.Rows < 3 And MSHFlexGrid1.TextMatrix(1, 0) = "" Then MsgBox "No Record Found": Exit Sub
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_SecurityRequestGodwonMapping.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_SecurityRequestGodwonMapping.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
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
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGetReport_Click()

'If CmbLocation.Text = "" Then
'    MsgBox "Blank Location not Allowed!!!"
'    Exit Sub
'End If

Call Grid_Head

wc = ""
If CmbLocation.Text <> "" Then
   wc = wc & " Where ML.LocationID=" & mLocationID
Else
    If Gen_WcLocation <> "" Then
        wc = wc & " Where ML." & Gen_WcLocation
    End If
End If
Wc1 = wc
If CmbSecurityAgency.Text <> "" Then
    If wc <> "" Then
        wc = wc & " And MSA.SecurityAgencyID=" & mSecurityAgencyID
    Else
        wc = " Where MSA.SecurityAgencyID=" & mSecurityAgencyID
    End If
End If
If CmbSecurityDesignation.Text <> "" Then
    If wc <> "" Then
        wc = wc & " And MSD.SecurityDesignationID=" & mSecurityDesignationID
    Else
        wc = " Where MSD.SecurityDesignationID=" & mSecurityDesignationID
    End If
End If
If CmbSecurityType.Text <> "" Then
    If wc <> "" Then
        wc = wc & " And MSS.SecurityShiftsID=" & mSecurityTypeID
    Else
        wc = " Where MSS.SecurityShiftsID=" & mSecurityTypeID
    End If
End If


If OptGodMapped.Value = True Then
    tmp = "Select "
    tmp = tmp & " MSLM.TxnID, MSLM.TxnDate, ML.LocationName, MSA.AgencyName, MSS.SecurityShifts, MSD.SecurityDesignation,"
    tmp = tmp & " MSLMD.TxnDetailID SecurityID, MSLMD.StartDate, MSLMD.EndDate, MSLMD.NoofGodown, MSLMD.RatePerMonth,"
    tmp = tmp & " MSLGMD.BalanceBags , MSLGMD.BalanceQty, MSLGMD.AUM, MCMP.CMPName, MG.GodownNo, MG.Address,MG.CloseDate"
    tmp = tmp & " From Mst_SecLocMappingDetail MSLMD"
    tmp = tmp & " Inner Join Mst_SecLocMapping MSLM on MSLM.TxnID = MSLMD.TxnID"
    tmp = tmp & " Inner Join Mst_SecurityAgency MSA on MSA.SecurityAgencyID = MSLM.SecurityAgencyID"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MSLM.LocationID"
    tmp = tmp & " Inner Join Mst_SecurityShifts MSS on MSS.SecurityShiftsID = MSLMD.SecurityTypeID"
    tmp = tmp & " Inner Join Mst_SecurityDesignation MSD on MSD.SecurityDesignationID = MSLMD.SecurityDesignationID"
    tmp = tmp & " left Join Mst_SecLocGodMappingDetail MSLGMD on MSLGMD.TxnDetailID = MSLMD.TxnDetailID"
    tmp = tmp & " left Join Mst_CMP MCMP on MSLGMD.CMPID = MCMP.CMPID"
    tmp = tmp & " left Join Mst_Godown MG on MSLGMD.GodownID = MG.GodownID"
    tmp = tmp & wc & " Order by ML.LocationName, MSA.AgencyName, MSS.SecurityShifts"
    
    Call TmpTable

    If TmpRs.RecordCount > 0 Then
        With MSHFlexGrid1
            .Rows = .Rows + TmpRs.RecordCount
            For I = 1 To TmpRs.RecordCount
                .TextMatrix(I, 0) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName) '"Location"
                .TextMatrix(I, 1) = IIf(IsNull(TmpRs!AgencyName), "", TmpRs!AgencyName) '"Agency Name"
                .TextMatrix(I, 2) = IIf(IsNull(TmpRs!SecurityShifts), "", TmpRs!SecurityShifts) '"Security Type"
                .TextMatrix(I, 3) = IIf(IsNull(TmpRs!SecurityID), "", TmpRs!SecurityID) '"Security ID"
                .TextMatrix(I, 4) = IIf(IsNull(TmpRs!SecurityDesignation), "", TmpRs!SecurityDesignation) '"Security Designation"
                .TextMatrix(I, 5) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName) '"CMP Name" 'NCDEX / Non - NCDEX
                .TextMatrix(I, 6) = IIf(IsNull(TmpRs!GodownNo), "", "'" & TmpRs!GodownNo) '"Godown No"
                .TextMatrix(I, 7) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address) '"Address"
                .TextMatrix(I, 8) = IIf(IsNull(TmpRs!TxnID), "", TmpRs!TxnID) '"TxnID"
                .TextMatrix(I, 9) = IIf(IsNull(TmpRs!TxnDate), "", Format(TmpRs!TxnDate, "dd-MMM-yyyy")) '"TxnDate"
                .TextMatrix(I, 10) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy")) '"Start Date"
                .TextMatrix(I, 11) = IIf(IsNull(TmpRs!EndDate), "", Format(TmpRs!EndDate, "dd-MMM-yyyy")) '"End Date"
                .TextMatrix(I, 12) = IIf(IsNull(TmpRs!NoofGodown), "", TmpRs!NoofGodown) '"No of Godown"
                .TextMatrix(I, 13) = IIf(IsNull(TmpRs!RatePerMonth), "", TmpRs!RatePerMonth) '"Rate per Month"
                '.TextMatrix(0, 13) = Val(.TextMatrix(I, 13)) + Val(.TextMatrix(, 13))
                .TextMatrix(I, 14) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags) '"Balance Bags"    'in the GSL
                .TextMatrix(I, 15) = IIf(IsNull(TmpRs!BalanceQty), "", TmpRs!BalanceQty) '"Balance Quantity"
                .TextMatrix(I, 16) = IIf(IsNull(TmpRs!AUM), "", TmpRs!AUM) '"AUM"
                
                ''----- Ullhas Patil 28 Sep 2017
                .TextMatrix(I, 17) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-YYYY"))
                
                TmpRs.MoveNext
            Next
        End With
       CmdExcel.Enabled = True
       MsgBox "Done !!!", vbInformation
    Else
       MsgBox " No Record Found "
    End If
    LblTotalRecord.Caption = TmpRs.RecordCount
ElseIf OptFreeSecurity.Value = True Then
    Call Grid_Head
    tmp = " Select"
    tmp = tmp & " MSLM.TxnID, MSLM.TxnDate, ML.LocationName, MSA.AgencyName, MSS.SecurityShifts, MSD.SecurityDesignation,"
    tmp = tmp & " MSLMD.TxnDetailID SecurityID, MSLMD.StartDate, MSLMD.EndDate, MSLMD.RatePerMonth"
    tmp = tmp & " From Mst_SecLocMappingDetail MSLMD"
    tmp = tmp & " Inner Join Mst_SecLocMapping MSLM on MSLM.TxnID = MSLMD.TxnID"
    tmp = tmp & " Inner Join Mst_SecurityAgency MSA on MSA.SecurityAgencyID = MSLM.SecurityAgencyID"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MSLM.LocationID"
    tmp = tmp & " Inner Join Mst_SecurityShifts MSS on MSS.SecurityShiftsID = MSLMD.SecurityTypeID"
    tmp = tmp & " Inner Join Mst_SecurityDesignation MSD on MSD.SecurityDesignationID = MSLMD.SecurityDesignationID"
    tmp = tmp & " left Join Mst_SecLocGodMappingDetail MSLGMD on MSLGMD.TxnDetailID = MSLMD.TxnDetailID"
    tmp = tmp & " left Join Mst_CMP MCMP on MSLGMD.CMPID = MCMP.CMPID"
    tmp = tmp & " left Join Mst_Godown MG on MSLGMD.GodownID = MG.GodownID"
    If wc <> "" Then
       tmp = tmp & wc
       tmp = tmp & " And MSLMD.EndDate > '" & Date & "' And MG.CloseDate is not null"
    Else
       tmp = tmp & " Where MSLMD.EndDate > '" & Date & "' And MG.CloseDate is not null"
    End If
    tmp = tmp & " And MSLMD.TxnDetailID not in (Select Distinct TxnDetailID From Mst_SecLocGodMappingDetail  MSLGMD inner join  Mst_Godown MG on MSLGMD.GodownID = MG.GodownID Where MG.CloseDate is null )"
    tmp = tmp & " Group by MSLM.TxnID, MSLM.TxnDate, ML.LocationName, MSA.AgencyName, MSS.SecurityShifts, MSD.SecurityDesignation,"
    tmp = tmp & " MSLMD.TxnDetailID , MSLMD.StartDate, MSLMD.EndDate,  MSLMD.RatePerMonth"
    'tmp = tmp & " Having Count(MG.GodownID) = Count(MG.CloseDate)"
    tmp = tmp & " Union All"
    tmp = tmp & " Select"
    tmp = tmp & " MSLM.TxnID, MSLM.TxnDate, ML.LocationName, MSA.AgencyName, MSS.SecurityShifts, MSD.SecurityDesignation,"
    tmp = tmp & " MSLMD.TxnDetailID SecurityID, MSLMD.StartDate, MSLMD.EndDate, MSLMD.RatePerMonth"
    tmp = tmp & " From Mst_SecLocMappingDetail MSLMD"
    tmp = tmp & " Inner Join Mst_SecLocMapping MSLM on MSLM.TxnID = MSLMD.TxnID"
    tmp = tmp & " Inner Join Mst_SecurityAgency MSA on MSA.SecurityAgencyID = MSLM.SecurityAgencyID"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MSLM.LocationID"
    tmp = tmp & " Inner Join Mst_SecurityShifts MSS on MSS.SecurityShiftsID = MSLMD.SecurityTypeID"
    tmp = tmp & " Inner Join Mst_SecurityDesignation MSD on MSD.SecurityDesignationID = MSLMD.SecurityDesignationID"
    If wc <> "" Then
       tmp = tmp & wc
       tmp = tmp & " And MSLMD.EndDate > '" & Date & "' And MSLMD.TxnDetailID not in (Select Distinct TxnDetailID From Mst_SecLocGodMappingDetail)"
    Else
       tmp = tmp & " Where MSLMD.EndDate > '" & Date & "' And MSLMD.TxnDetailID not in (Select Distinct TxnDetailID From Mst_SecLocGodMappingDetail)"
    End If
    
    
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        With MSHFlexGrid1
            .Rows = .Rows + TmpRs.RecordCount
            For I = 1 To TmpRs.RecordCount
                .TextMatrix(I, 0) = IIf(IsNull(TmpRs!TxnID), "", TmpRs!TxnID)
                .TextMatrix(I, 1) = IIf(IsNull(TmpRs!TxnDate), "", Format(TmpRs!TxnDate, "dd-mmm-yyyy"))
                .TextMatrix(I, 2) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
                .TextMatrix(I, 3) = IIf(IsNull(TmpRs!AgencyName), "", TmpRs!AgencyName)
                .TextMatrix(I, 4) = IIf(IsNull(TmpRs!SecurityShifts), "", TmpRs!SecurityShifts)
                .TextMatrix(I, 5) = IIf(IsNull(TmpRs!SecurityDesignation), "", TmpRs!SecurityDesignation)
                .TextMatrix(I, 6) = IIf(IsNull(TmpRs!SecurityID), "", TmpRs!SecurityID)
                .TextMatrix(I, 7) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-mmm-yyyy"))
                .TextMatrix(I, 8) = IIf(IsNull(TmpRs!EndDate), "", Format(TmpRs!EndDate, "dd-mmm-yyyy"))
                .TextMatrix(I, 9) = IIf(IsNull(TmpRs!RatePerMonth), "", TmpRs!RatePerMonth)
                
                TmpRs.MoveNext
            Next
        End With
       CmdExcel.Enabled = True
       MsgBox "Done !!!", vbInformation
    Else
       MsgBox " No Record Found "
    End If
    LblTotalRecord.Caption = TmpRs.RecordCount
Else
    tmp = " select ms.statename, ml.locationname, mcmp.cmpname, mg.godownno, mg.address"
    '--,(select ma.agencyname from mst_securityagency ma inner join mst_seclocmapping msl on msl.securityagencyid = ma.agencyid where msl.locationid = ml.locationid )
    tmp = tmp & " from mst_godown mg"
    tmp = tmp & " inner join mst_cmp mcmp on mcmp.cmpid = mg.cmpid"
    tmp = tmp & " inner join mst_location ml on ml.locationid = mcmp.locationid"
    tmp = tmp & " inner join mst_state ms on ms.stateid = ml.stateid"
    If Wc1 <> "" Then
        tmp = tmp & Wc1 & " and "
    Else
        tmp = tmp & " where "
    End If
    tmp = tmp & " mg.closedate is null and mg.godownid not in (select godownid from Mst_SecLocGodMappingDetail)"
    tmp = tmp & " Order By ms.statename, ml.locationname, mcmp.cmpname"
    Call TmpTable

    If TmpRs.RecordCount > 0 Then
        With MSHFlexGrid1
            .Rows = .Rows + TmpRs.RecordCount
            For I = 1 To TmpRs.RecordCount
                .TextMatrix(I, 0) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
                .TextMatrix(I, 1) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
                .TextMatrix(I, 2) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
                .TextMatrix(I, 3) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
                '.TextMatrix(I, 4) = IIf(IsNull(TmpRs!AgencyName), "", TmpRs!AgencyName)
                TmpRs.MoveNext
            Next
        End With
       CmdExcel.Enabled = True
       MsgBox "Done !!!", vbInformation
    Else
       MsgBox " No Record Found "
    End If
    LblTotalRecord.Caption = TmpRs.RecordCount

End If
End Sub

Private Sub CmbSecurityType_GotFocus()
tmp = CmbSecurityType.Text
Call TableMst_SecurityShifts("")

CmbSecurityType.Clear
If RsMst_SecurityShifts.RecordCount = 0 Then
   MsgBox "No Security Type found!!!"
   CmbSecurityType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityShifts.RecordCount
    CmbSecurityType.AddItem RsMst_SecurityShifts!SecurityShifts
    RsMst_SecurityShifts.MoveNext
Next
CmbSecurityType.Text = tmp
End Sub

Private Sub CmbSecurityType_LostFocus()
If CmbSecurityType.Text = "" Then Exit Sub

RsMst_SecurityShifts.MoveFirst
RsMst_SecurityShifts.Find "SecurityShifts = '" & CmbSecurityType.Text & "'"

If RsMst_SecurityShifts.EOF Then
   MsgBox "Invalid Shift Selection "
   CmbSecurityType.SetFocus
   Exit Sub
End If
mSecurityTypeID = RsMst_SecurityShifts!SecurityShiftsID

End Sub
Private Sub CmbSecurityDesignation_GotFocus()
tmp = CmbSecurityDesignation.Text
Call TableMst_SecurityDesignation("")

CmbSecurityDesignation.Clear
If RsMst_SecurityDesignation.RecordCount = 0 Then
   MsgBox "No Security Designation found!!!"
   CmbSecurityType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_SecurityDesignation.RecordCount
    CmbSecurityDesignation.AddItem RsMst_SecurityDesignation!SecurityDesignation
    RsMst_SecurityDesignation.MoveNext
Next
CmbSecurityDesignation.Text = tmp
End Sub

Private Sub CmbSecurityDesignation_LostFocus()
If CmbSecurityDesignation.Text = "" Then Exit Sub

RsMst_SecurityDesignation.MoveFirst
RsMst_SecurityDesignation.Find "SecurityDesignation = '" & CmbSecurityDesignation.Text & "'"

If RsMst_SecurityDesignation.EOF Then
   MsgBox "Invalid Designation Selection "
   CmbSecurityDesignation.SetFocus
   Exit Sub
End If
mSecurityDesignationID = RsMst_SecurityDesignation!SecurityDesignationID
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
CmdExcel.Enabled = False
Call Grid_Head
If Gen_SecLocation <> "" Then
    CmbLocation.Text = Gen_SecLocation
    mLocationID = Gen_SecLocationID
End If

End Sub


Private Sub Grid_Head()
CmdExcel.Enabled = False
LblTotalRecord.Caption = ""
If OptGodMapped.Value = True Then
    With MSHFlexGrid1
        .Clear
        .Rows = 2
        .FixedRows = 1
        .Cols = 18  '17
        .Font.Size = 10
        .AllowUserResizing = flexResizeBoth
        
        .TextMatrix(0, 0) = "Location"
        .TextMatrix(0, 1) = "Agency Name"
        .TextMatrix(0, 2) = "Security Type"
        .TextMatrix(0, 3) = "Security ID"
        .TextMatrix(0, 4) = "Security Designation"
        .TextMatrix(0, 5) = "CMP Name" 'NCDEX / Non - NCDEX
        .TextMatrix(0, 6) = "Godown No"
        .TextMatrix(0, 7) = "Address"
        .TextMatrix(0, 8) = "TxnID"
        .TextMatrix(0, 9) = "TxnDate"
        .TextMatrix(0, 10) = "Start Date"
        .TextMatrix(0, 11) = "End Date"
        .TextMatrix(0, 12) = "No of Godown"
        .TextMatrix(0, 13) = "Rate per Month"
        .TextMatrix(0, 14) = "Balance Bags"    'in the GSL
        .TextMatrix(0, 15) = "Balance Quantity"
        .TextMatrix(0, 16) = "AUM"
        
        .ColWidth(0) = 1800
        .ColWidth(1) = 1800
        .ColWidth(2) = 1800
        .ColWidth(3) = 1800
        .ColWidth(4) = 1800
        .ColWidth(5) = 1800
        .ColWidth(6) = 1500
        .ColWidth(7) = 2500
        .ColWidth(8) = 1200
        .ColWidth(9) = 1500
        .ColWidth(10) = 1500
        .ColWidth(11) = 1500
        .ColWidth(12) = 1500
        .ColWidth(13) = 1500
        .ColWidth(14) = 1500
        .ColWidth(15) = 1500
        .ColWidth(16) = 1500
        
        ''--- Ullhas Patil 28 Sep 2017
        .TextMatrix(0, 17) = "Godown Close Date": .ColWidth(17) = 2000
        ''----------------
        
        .RowHeight(0) = 400
    End With
ElseIf OptGodNotMapped.Value = True Then
    With MSHFlexGrid1
        .Clear
        .Rows = 2
        .FixedRows = 1
        .Cols = 4
        .Font.Size = 10
        .AllowUserResizing = flexResizeBoth
        
        '.TextMatrix(0, 0) = "State Name"
        .TextMatrix(0, 0) = "Location Name"
        .TextMatrix(0, 1) = "CMP Name"
        .TextMatrix(0, 2) = "Godown No"
        .TextMatrix(0, 3) = "Godown Address"
        '.TextMatrix(0, 4) = "Availability of Security Agency"
            
        '.ColWidth(0) = 1900
        .ColWidth(0) = 2500
        .ColWidth(1) = 2500
        .ColWidth(2) = 2500
        .ColWidth(3) = 8000
        '.ColWidth(4) = 3500
        
        .RowHeight(0) = 400
    End With
Else

    With MSHFlexGrid1
        .Clear
        .Rows = 2
        .FixedRows = 1
        .Cols = 10
        .Font.Size = 11
        .AllowUserResizing = flexResizeBoth
        
        .TextMatrix(0, 0) = "Txn ID"
        .TextMatrix(0, 1) = "Txn Date"
        .TextMatrix(0, 2) = "Location Name"
        .TextMatrix(0, 3) = "Agency Name"
        .TextMatrix(0, 4) = "Security Shifts"
        .TextMatrix(0, 5) = "Security Designation"
        .TextMatrix(0, 6) = "Security ID"
        .TextMatrix(0, 7) = "Start Date"
        .TextMatrix(0, 8) = "End Date"
        .TextMatrix(0, 9) = "Rate Per Month"
            
        '.ColWidth(0) = 1900
        .ColWidth(0) = 800
        .ColWidth(1) = 1000
        .ColWidth(2) = 1400
        .ColWidth(3) = 2000
        '.ColWidth(4) = 3500
        
        .RowHeight(0) = 400
    End With

End If
End Sub

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
    MsgBox "No Record found !!!! "
    Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then Exit Sub

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub OptGodMapped_Click()
Call Grid_Head

CmbSecurityAgency.Enabled = True
CmbSecurityDesignation.Enabled = True
CmbSecurityType.Enabled = True
End Sub

Private Sub OptGodNotMapped_Click()
Call Grid_Head

CmbSecurityAgency.Enabled = False
CmbSecurityDesignation.Enabled = False
CmbSecurityType.Enabled = False

CmbSecurityAgency.Text = ""
CmbSecurityDesignation.Text = ""
CmbSecurityType.Text = ""
End Sub
