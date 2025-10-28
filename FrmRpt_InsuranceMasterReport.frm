VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_InsuranceMasterReport 
   Caption         =   "Insurance Policy Master Report"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   9135
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   16260
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   14760
         TabIndex        =   11
         Top             =   2200
         Width           =   1290
      End
      Begin VB.ComboBox CmbSLocationID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   12240
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1150
         Width           =   3840
      End
      Begin VB.ComboBox CmbSClientID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1800
         Width           =   7920
      End
      Begin VB.ComboBox CmbSCommodityID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1800
         Width           =   3960
      End
      Begin VB.ComboBox CmbSStateID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   8280
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1150
         Width           =   3840
      End
      Begin VB.ComboBox CmbSInsuranceByID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4200
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1150
         Width           =   3960
      End
      Begin VB.ComboBox CmbSInsurancePolicyID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   8010
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   13470
         TabIndex        =   10
         Top             =   2200
         Width           =   1260
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   375
         Left            =   12210
         TabIndex        =   9
         Top             =   2200
         Width           =   1260
      End
      Begin VB.ComboBox CmbSInsuranceCoID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   8265
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   7845
      End
      Begin VB.ComboBox CmbSInsuranceRiskTypeID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1150
         Width           =   3960
      End
      Begin VB.ComboBox CmbSStatus 
         Appearance      =   0  'Flat
         Height          =   360
         ItemData        =   "FrmRpt_InsuranceMasterReport.frx":0000
         Left            =   12225
         List            =   "FrmRpt_InsuranceMasterReport.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1800
         Width           =   3855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5655
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   9975
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   25
         Top             =   8400
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Client Name"
         Height          =   240
         Left            =   7598
         TabIndex        =   24
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Commodity"
         Height          =   240
         Left            =   1530
         TabIndex        =   23
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Location"
         Height          =   240
         Left            =   13763
         TabIndex        =   22
         Top             =   900
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "State"
         Height          =   240
         Left            =   9953
         TabIndex        =   21
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Insurance Risk Type"
         Height          =   240
         Left            =   1178
         TabIndex        =   20
         Top             =   900
         Width           =   1845
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Insurance By"
         Height          =   240
         Left            =   5588
         TabIndex        =   19
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label Label11 
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
         Height          =   255
         Left            =   8445
         TabIndex        =   18
         Top             =   2250
         Width           =   2295
      End
      Begin VB.Label LblTotalRecords 
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
         Left            =   10755
         TabIndex        =   17
         Top             =   2265
         Width           =   90
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Insurance Company"
         Height          =   240
         Left            =   11535
         TabIndex        =   16
         Top             =   225
         Width           =   1785
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Policy No. And Cover Note No."
         Height          =   240
         Left            =   2760
         TabIndex        =   15
         Top             =   225
         Width           =   2745
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Insurance Status"
         Height          =   240
         Left            =   13395
         TabIndex        =   14
         Top             =   1560
         Width           =   1485
      End
   End
End
Attribute VB_Name = "FrmRpt_InsuranceMasterReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSInsurancePolicyID, mSInsuranceCompanyID, mSInsuranceRiskTypeID, mSInsuranceByID, mSStateID, mSLocationID, mSCommodityID, mSClientID As Variant


Private Sub CmbSClientID_GotFocus()
tmp = CmbSClientID.Text
'Call TableMst_Client("Where ExchangeTypeID = 1")
Call TableMst_Client
CmbSClientID.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbSClientID.AddItem RsMst_Client!ClientName & "#" & RsMst_Client!ClientIDRow & "#" & RsMst_Client!ClientID
    RsMst_Client.MoveNext
Next
CmbSClientID.Text = tmp
End Sub
Private Sub CmbSClientId_LostFocus()
Dim mSplit As Variant
If CmbSClientID.Text = "" Then Exit Sub
mSplit = Split(CmbSClientID.Text, "#")
If UBound(mSplit) <> 2 Then
   MsgBox "Invalid selection."
   CmbSClientID.SetFocus
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientID = '" & mSplit(2) & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbSClientID.SetFocus
   Exit Sub
End If
If LCase(RsMst_Client!ClientName & "#" & RsMst_Client!ClientIDRow) <> LCase(mSplit(0) & "#" & mSplit(1)) Then
   MsgBox "Invalid Client."
   CmbSClientID.SetFocus
   Exit Sub
End If
mSClientID = RsMst_Client!ClientID
End Sub

Private Sub CmdGo_Click()
Dim mWC As Variant
Dim mR As Double, mRow As Double

MsgBox "Wait till next message."
ProgressBar1.Value = 0
Call Grid_Head
LblTotalRecords.Caption = ""
CmdExcel.Enabled = False

If CmbSInsurancePolicyID.Text <> "" Then
   mWC = " Where MID.InsurancePolicyID = " & mSInsurancePolicyID & " "
End If

If CmbSInsuranceCoID.Text <> "" Then
   If mWC = "" Then
    mWC = " Where MID.InsuranceCoID = " & mSInsuranceCompanyID & " "
   Else
    mWC = mWC & " And MID.InsuranceCoID = " & mSInsuranceCompanyID & " "
   End If
End If
If CmbSInsuranceRiskTypeID.Text <> "" Then
   If mWC = "" Then
    mWC = " Where MID.InsuranceRiskCoveredID = " & mSInsuranceRiskTypeID & " "
   Else
    mWC = mWC & " And MID.InsuranceRiskCoveredID = " & mSInsuranceRiskTypeID & " "
   End If
End If
If CmbSInsuranceByID.Text <> "" Then
   If mWC = "" Then
    mWC = " Where MID.InsuranceByID = " & mSInsuranceByID & " "
   Else
    mWC = mWC & " And MID.InsuranceByID = " & mSInsuranceByID & " "
   End If
End If
If CmbSStateID.Text <> "" Then
   If mWC = "" Then
    mWC = " Where MS.StateID = " & mSStateID & " "
   Else
    mWC = mWC & " And MS.StateID = " & mSStateID & " "
   End If
End If
If CmbSLocationID.Text <> "" Then
   If mWC = "" Then
    mWC = " Where ML.LocationID = " & mSLocationID & " "
   Else
    mWC = mWC & " And ML.LocationID = " & mSLocationID & " "
   End If
End If
If CmbSCommodityID.Text <> "" Then
   If mWC = "" Then
    mWC = " Where MComm.LocationID = " & mSCommodityID & " "
   Else
    mWC = mWC & " And MComm.LocationID = " & mSCommodityID & " "
   End If
End If
If CmbSClientID.Text <> "" Then
   If mWC = "" Then
    mWC = " Where MICDD.ClientID = " & mSClientID & " "
   Else
    mWC = mWC & " And MICDD.ClientID = " & mSClientID & " "
   End If
End If
If CmbSStatus.Text <> "" Then
   If mWC = "" Then
    mWC = " Where MID.Flag = '" & Left(CmbSStatus.Text, 1) & "' "
   Else
    mWC = mWC & " And MID.Flag = '" & Left(CmbSStatus.Text, 1) & "' "
   End If
End If

tmp = " Select MID.InsurancePolicyID,MID.PolicyNo,MID.CoverNoteNo,MID.InsuranceCoID,MICM.InsuranceCoName,MID.InsuranceRiskCoveredID,"
tmp = tmp & " MIRC.InsuranceRiskCovered , MID.InsuranceByID, MIB.InsuranceByName, MID.ValidFrom, MID.ValidTo, MID.Flag, MID.Amount, MID.StructureWiseLimit, MID.Remark"
tmp = tmp & " ,MIGM.GodownID,MG.GodownNo,Case MIGM.Connected WHen 'Y' Then 'Yes' When 'N' Then 'No' Else '' End Connected,MUT.UnitType,CMP.CMPName,ML.LocationName,MS.StateName,MR.Region"
tmp = tmp & " ,MICDI.CommodityID,MComm.Commodity,Case Isnull(MIGM.GodownID,0) When 0 Then '' Else Case MICDI.EntryType When 'I' Then 'Include' When 'E' Then 'Exclude' Else 'Include All Commodity' End End EntryType"
tmp = tmp & " ,MICDD.ClientID,MC.ClientName,MC.ClientIDRow,ME.ExchangeType"
tmp = tmp & " from Mst_InsuranceDetails MID"
tmp = tmp & " Left Join Mst_InsuranceBy MIB On MID.InsuranceByID = MIB.InsuranceByID"
tmp = tmp & " Left Join Mst_InsuranceRiskCovered MIRC On MID.InsuranceRiskCoveredID = MIRC.InsuranceRiskCoveredID"
tmp = tmp & " Left Join Mst_InsuranceCoMaster MICM On MID.InsuranceCoID = MICM.InsuranceCoID"
tmp = tmp & " Left Join Mst_InsuranceGodownMapping MIGM  On MID.InsurancePolicyID = MIGM.InsurancePolicyID"
tmp = tmp & " Left Join Mst_InsuranceCommodityDetail MICDI On MID.InsurancePolicyID = MICDI.InsurancePolicyID"
tmp = tmp & " Left Join Mst_InsuranceClientDetail MICDD On MID.InsurancePolicyID = MICDD.InsurancePolicyID"
tmp = tmp & " Left Join Mst_Godown MG On MIGM.GodownID = MG.GodownID"
tmp = tmp & " Left Join Mst_UnitType MUT On MG.UnitTypeID = MUT.UnitTypeID"
tmp = tmp & " Left Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
tmp = tmp & " Left Join Mst_Location ML On CMP.LocationID = ML.LocationID"
tmp = tmp & " Left Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Left Join Mst_Region MR On MS.RegionID = MR.RegionID"
tmp = tmp & " Left Join Mst_Commodity MComm On MICDI.COmmodityID = MComm.CommodityID"
tmp = tmp & " Left Join Mst_Client MC On MICDD.ClientID = MC.ClientID"
tmp = tmp & " Left Join Mst_ExchangeType ME On MC.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " " & mWC & ""
tmp = tmp & " Order By MID.InsurancePolicyID"
Call TmpTable

If TmpRs.RecordCount > 0 Then
With MSHFlexGrid1
ProgressBar1.Max = TmpRs.RecordCount
For mR = 1 To TmpRs.RecordCount
        
        .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!InsurancePolicyID), "", TmpRs!InsurancePolicyID)
        .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!PolicyNo), "", TmpRs!PolicyNo)
        .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!CoverNoteNo), "", TmpRs!CoverNoteNo)
        .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!InsuranceCoName), "", TmpRs!InsuranceCoName)
        .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!InsuranceRiskCovered), "", TmpRs!InsuranceRiskCovered)
        .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!InsuranceByName), "", TmpRs!InsuranceByName)
        .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!ValidFrom), "", Format(TmpRs!ValidFrom, "dd-MMM-yyyy"))
        .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!ValidTo), "", Format(TmpRs!ValidTo, "dd-MMM-yyyy"))
        .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!Flag), "", IIf(LCase(TmpRs!Flag) = "a", "Active", "Deactive"))
        .TextMatrix(mR, 9) = IIf(IsNull(TmpRs!Amount), "", Format(TmpRs!Amount, "#0.00"))
        .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!StructureWiseLimit), "", Format(TmpRs!StructureWiseLimit, "#0.00"))
        .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!Remark), "", TmpRs!Remark)
        .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
        .TextMatrix(mR, 13) = IIf(IsNull(TmpRs!Connected), "", TmpRs!Connected)
        .TextMatrix(mR, 14) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
        .TextMatrix(mR, 15) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
        .TextMatrix(mR, 16) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
        .TextMatrix(mR, 17) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
        .TextMatrix(mR, 18) = IIf(IsNull(TmpRs!Region), "", TmpRs!Region)
        .TextMatrix(mR, 19) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
        .TextMatrix(mR, 20) = IIf(IsNull(TmpRs!EntryType), "", TmpRs!EntryType)
        .TextMatrix(mR, 21) = IIf(IsNull(TmpRs!ClientID), "", IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName) & "#" & IIf(IsNull(TmpRs!ClientIDRow), "", TmpRs!ClientIDRow) & "#" & IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID))
        .TextMatrix(mR, 22) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
        
    ProgressBar1.Value = mR
    .Rows = .Rows + 1
    TmpRs.MoveNext
Next mR
End With
End If
LblTotalRecords.Caption = TmpRs.RecordCount
MsgBox "Done."
CmdExcel.Enabled = True
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Exit Sub
End If
Call CmdClear_Click
End Sub

Private Sub CmdClear_Click()
CmbSInsurancePolicyID = ""
CmbSInsuranceCoID = ""
CmbSInsuranceRiskTypeID = ""
CmbSInsuranceByID = "": CmbSClientID.Enabled = True
CmbSStateID = ""
CmbSLocationID = ""
CmbSCommodityID = ""
CmbSClientID = ""
CmbSStatus = ""
CmdExcel.Enabled = False
LblTotalRecords.Caption = "-"
Call Grid_Head
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .FixedRows = 1
    .FixedCols = 1
    .WordWrap = True
    .Cols = 23
    .TextMatrix(0, 0) = "Insurance Policy ID":        .ColWidth(0) = 1000
    .TextMatrix(0, 1) = "Policy No":                  .ColWidth(1) = 2500
    .TextMatrix(0, 2) = "Cover Note No":              .ColWidth(2) = 2500
    .TextMatrix(0, 3) = "Insurance Company":          .ColWidth(3) = 2500
    .TextMatrix(0, 4) = "Insurance Risk Covered":     .ColWidth(4) = 2500
    .TextMatrix(0, 5) = "Insurance By":               .ColWidth(5) = 1200
    .TextMatrix(0, 6) = "Valid From":                 .ColWidth(6) = 1500
    .TextMatrix(0, 7) = "Valid To":                   .ColWidth(7) = 1500
    .TextMatrix(0, 8) = "Status":                     .ColWidth(8) = 1000
    .TextMatrix(0, 9) = "Amount":                     .ColWidth(9) = 1800
    .TextMatrix(0, 10) = "Structure Wise Limit":      .ColWidth(10) = 1800
    .TextMatrix(0, 11) = "Remark":                    .ColWidth(11) = 3000
    .TextMatrix(0, 12) = "Godown No":                 .ColWidth(12) = 1200
    .TextMatrix(0, 13) = "Connected":                 .ColWidth(13) = 1200
    .TextMatrix(0, 14) = "Unit Type":                 .ColWidth(14) = 2000
    .TextMatrix(0, 15) = "CMP Name":                  .ColWidth(15) = 1500
    .TextMatrix(0, 16) = "Location Name":             .ColWidth(16) = 2000
    .TextMatrix(0, 17) = "State Name":                .ColWidth(17) = 2000
    .TextMatrix(0, 18) = "Region Name":               .ColWidth(18) = 2000
    .TextMatrix(0, 19) = "Commodity":                 .ColWidth(19) = 2000
    .TextMatrix(0, 20) = "Entry Type":                .ColWidth(20) = 1200
    .TextMatrix(0, 21) = "Client Name":               .ColWidth(21) = 3500
    .TextMatrix(0, 22) = "Exchange Type":             .ColWidth(22) = 1500
    .RowHeight(0) = 800
End With
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_InsuranceMasterReport.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_InsuranceMasterReport.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

With MSHFlexGrid1
ProgressBar1.Max = MSHFlexGrid1.Rows - 1
    For I = 0 To .Rows - 1
        mRow = I + 1 ' + 6
        For C = 0 To .Cols - 1
            If InStr(.TextMatrix(I, C), "/") > 0 Then
               If IsDate(.TextMatrix(I, C)) Then
                  oWS.Cells(mRow, C + 1) = Format(.TextMatrix(I, C), "MM/dd/yyyy")
               Else
                  oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
               End If
            Else
               If C = 1 Or C = 2 Or C = 12 Then
                oWS.Cells(mRow, C + 1) = "'" & .TextMatrix(I, C)
               Else
                oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
               End If
            End If
        Next
        ProgressBar1.Value = I
    Next
End With
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmbSInsurancePolicyID_GotFocus()
tmp = CmbSInsurancePolicyID.Text
Call TableMst_PolicyDetail(" Order By InsurancePolicyID ")
CmbSInsurancePolicyID.Clear
If RsMst_PolicyDetail.RecordCount = 0 Then
   MsgBox "No Insurance Policy found"
   CmbSInsurancePolicyID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_PolicyDetail.RecordCount
    CmbSInsurancePolicyID.AddItem RsMst_PolicyDetail!PolicyNo & "~" & RsMst_PolicyDetail!CoverNoteNo & "~" & RsMst_PolicyDetail!InsurancePolicyID
    RsMst_PolicyDetail.MoveNext
Next
CmbSInsurancePolicyID.Text = tmp
End Sub

Private Sub CmbSInsurancePolicyID_LostFocus()
Dim mSplit As Variant
If CmbSInsurancePolicyID.Text = "" Then Exit Sub
mSplit = Split(CmbSInsurancePolicyID.Text, "~")
If UBound(mSplit) <> 2 Then
   MsgBox "Invalid Insurance Policy."
   CmbSInsurancePolicyID.SetFocus
   Exit Sub
End If
RsMst_PolicyDetail.MoveFirst
RsMst_PolicyDetail.Find "InsurancePolicyID = '" & mSplit(2) & "'"
If RsMst_PolicyDetail.EOF Then
   MsgBox "Invalid Insurance Policy."
   CmbSInsurancePolicyID.SetFocus
   Exit Sub
End If

If LCase(RsMst_PolicyDetail!PolicyNo & "~" & RsMst_PolicyDetail!CoverNoteNo) <> LCase(mSplit(0) & "~" & mSplit(1)) Then
   MsgBox "Invalid Insurance Policy."
   CmbSInsurancePolicyID.SetFocus
   Exit Sub
End If

mSInsurancePolicyID = RsMst_PolicyDetail!InsurancePolicyID
End Sub

Private Sub CmbSInsuranceByID_GotFocus()
tmp = CmbSInsuranceByID.Text
Call TableMst_InsuranceBy(" Where Flag = 'A' ")
CmbSInsuranceByID.Clear
If RsMst_InsuranceBy.RecordCount = 0 Then
   MsgBox "No Insurance By found"
   CmbSInsuranceByID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_InsuranceBy.RecordCount
    CmbSInsuranceByID.AddItem RsMst_InsuranceBy!InsuranceByName
    RsMst_InsuranceBy.MoveNext
Next
CmbSInsuranceByID.Text = tmp
End Sub
Private Sub CmbSInsuranceByID_LostFocus()
If CmbSInsuranceByID.Text = "" Then CmbSClientID.Enabled = True: Exit Sub
RsMst_InsuranceBy.MoveFirst
RsMst_InsuranceBy.Find "InsuranceByName = '" & CmbSInsuranceByID.Text & "'"
If RsMst_InsuranceBy.EOF Then
   MsgBox "Invalid Insurance By selection "
   CmbSInsuranceByID.SetFocus
   Exit Sub
End If
mSInsuranceByID = RsMst_InsuranceBy!InsuranceByID

If mSInsuranceByID = 1 Then
  CmbSClientID.Text = "": CmbSClientID.Enabled = False
Else
  CmbSClientID.Text = "": CmbSClientID.Enabled = True
End If

End Sub

Private Sub CmbSInsuranceCoID_GotFocus()
tmp = CmbSInsuranceCoID.Text
Call TableMst_InsuranceCo(" Where Flag = 'A' ")
CmbSInsuranceCoID.Clear
If RsMst_InsuranceCo.RecordCount = 0 Then
   MsgBox "No InsuranceCo found"
   CmbSInsuranceCoID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_InsuranceCo.RecordCount
    CmbSInsuranceCoID.AddItem RsMst_InsuranceCo!InsuranceCoName
    RsMst_InsuranceCo.MoveNext
Next
CmbSInsuranceCoID.Text = tmp
End Sub
Private Sub CmbSInsuranceCoID_LostFocus()
If CmbSInsuranceCoID.Text = "" Then
   Exit Sub
End If
RsMst_InsuranceCo.MoveFirst
RsMst_InsuranceCo.Find "InsuranceCoName = '" & CmbSInsuranceCoID.Text & "'"

If RsMst_InsuranceCo.EOF Then
   MsgBox "Invalid Insurance Company selection "
   CmbSInsuranceCoID.SetFocus
   Exit Sub
End If
mSInsuranceCompanyID = RsMst_InsuranceCo.Fields(0)
End Sub

Private Sub CmbSInsuranceRiskTypeID_GotFocus()
tmp = CmbSInsuranceRiskTypeID.Text
Call TableMst_InsuranceRiskCovered(" Where Status = 'A' ")
CmbSInsuranceRiskTypeID.Clear
If RsMst_InsuranceRiskCovered.RecordCount = 0 Then
   MsgBox "No Insurance Risk Type found"
   CmbSInsuranceRiskTypeID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_InsuranceRiskCovered.RecordCount
    CmbSInsuranceRiskTypeID.AddItem RsMst_InsuranceRiskCovered!InsuranceRiskCovered
    RsMst_InsuranceRiskCovered.MoveNext
Next
CmbSInsuranceRiskTypeID.Text = tmp
End Sub
Private Sub CmbSInsuranceRiskTypeID_LostFocus()
If CmbSInsuranceRiskTypeID.Text = "" Then Exit Sub
RsMst_InsuranceRiskCovered.MoveFirst
RsMst_InsuranceRiskCovered.Find "InsuranceRiskCovered = '" & CmbSInsuranceRiskTypeID.Text & "'"
If RsMst_InsuranceRiskCovered.EOF Then
   MsgBox "Invalid Insurance Risk Type selection "
   CmbSInsuranceRiskTypeID.SetFocus
   Exit Sub
End If
mSInsuranceRiskTypeID = RsMst_InsuranceRiskCovered!InsuranceRiskCoveredID
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub
If LCase(CmbSStatus.Text) <> LCase("Active") And LCase(CmbSStatus.Text) <> LCase("Deactive") Then
   MsgBox "Invalid Selection"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSStateID_GotFocus()
tmp = CmbSStateID.Text
If Gen_WcState = "" Then
   Call TableMst_State
Else
   Call TableMst_State(" Where " & Gen_WcState)
End If
CmbSStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbSStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbSStateID.Text = tmp
End Sub
Private Sub CmbSStateID_LostFocus()
If CmbSStateID.Text = "" Then Exit Sub
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbSStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbSStateID.SetFocus
   Exit Sub
End If
If mSStateID <> RsMst_State!StateID Then
   CmbSLocationID.Clear
End If
mSStateID = RsMst_State!StateID
End Sub

Private Sub CmbSLocationID_GotFocus()
If CmbSStateID.Text = "" Then
   MsgBox "Please Select State Name !!"
   CmbSStateID.SetFocus
   Exit Sub
End If
tmp = CmbSLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location("Where StateID=" & mSStateID & "")
Else
   Call TableMst_Location("Where StateID=" & mStateID & " And " & Gen_WcLocation)
End If
CmbSLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbSLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbSLocationID.Text = tmp
End Sub
Private Sub CmbSLocationID_LostFocus()
If CmbSLocationID.Text = "" Then Exit Sub
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbSLocationID.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSCommodityID_GotFocus()
tmp = CmbSCommodityID.Text
Call TableMst_Commodity
CmbSCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbSCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbSCommodityID.Text = tmp
End Sub
Private Sub CmbSCommodityID_LostFocus()
If CmbSCommodityID.Text = "" Then Exit Sub
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbSCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection"
   CmbSCommodityID.SetFocus
   Exit Sub
End If
mSCommodityID = RsMst_Commodity!CommodityID
End Sub

