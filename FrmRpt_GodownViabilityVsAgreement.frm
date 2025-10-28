VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_GodownViabilityVsAgreement 
   Caption         =   "Godown Viability V/s Agreement Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8415
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   13996
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         Appearance      =   0
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   8160
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbSLocation 
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
         ItemData        =   "FrmRpt_GodownViabilityVsAgreement.frx":0000
         Left            =   120
         List            =   "FrmRpt_GodownViabilityVsAgreement.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   450
         Width           =   4365
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
         Left            =   8920
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   450
         Width           =   3045
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   12840
         TabIndex        =   4
         Top             =   880
         Width           =   1000
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
         Left            =   4515
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   450
         Width           =   4365
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
         Height          =   350
         Left            =   13845
         TabIndex        =   6
         Top             =   880
         Width           =   1000
      End
      Begin VB.TextBox TxtSViabilityID 
         Alignment       =   1  'Right Justify
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
         Left            =   12000
         TabIndex        =   3
         Top             =   450
         Width           =   2925
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
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
         Left            =   1920
         TabIndex        =   15
         Top             =   150
         Width           =   765
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
         Left            =   9120
         TabIndex        =   13
         Top             =   900
         Width           =   2295
      End
      Begin VB.Label Label2 
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
         Left            =   11520
         TabIndex        =   12
         Top             =   900
         Width           =   90
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "Godown"
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
         Left            =   10200
         TabIndex        =   10
         Top             =   150
         Width           =   750
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "CMP"
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
         Left            =   6480
         TabIndex        =   9
         Top             =   150
         Width           =   435
      End
      Begin VB.Label Label82 
         AutoSize        =   -1  'True
         Caption         =   "Viability ID"
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
         Left            =   13080
         TabIndex        =   8
         Top             =   150
         Width           =   945
      End
   End
End
Attribute VB_Name = "FrmRpt_GodownViabilityVsAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mR, mC As Variant
Dim mType, mLocationID, mCMPID, mGodownID, mGodownTypeID As Variant

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
End If
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
   CmbSLocation.AddItem RsMst_Location!Locationname
   RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp
End Sub

Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then
   CmbSCmp.Text = ""
   CmbSGodown.Text = ""
   Exit Sub
End If
If tmp <> CmbSLocation.Text Then
   CmbSCmp.Text = ""
   CmbSGodown.Text = ""
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbLocation.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSCmp_GotFocus()
If Trim(CmbSLocation.Text) = "" Then
   MsgBox "Select Location!!!"
   CmbSLocation.SetFocus
   Exit Sub
End If
tmp = CmbSCmp.Text
If Gen_WcCMP = "" Then
   Call TableMst_CMP(" where CloseDate is null and LocationID = " & mLocationID)
Else
   Call TableMst_CMP(" where " & Gen_WcCMP & " And CloseDate is null and LocationID = " & mLocationID)
End If
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
   CmbSCmp.AddItem RsMst_CMP!CMPName
   RsMst_CMP.MoveNext
Next
CmbSCmp.Text = tmp
End Sub

Private Sub CmbSCmp_LostFocus()
If CmbSCmp.Text = "" Then
   CmbSGodown.Text = ""
   Exit Sub
End If
If tmp <> CmbSCmp.Text Then
   CmbSGodown.Text = ""
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSCmp.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub


Private Sub CmbSGodown_GotFocus()
If Trim(CmbSCmp.Text) = "" Then
   MsgBox "Select CMP!!!"
   CmbSCmp.SetFocus
   Exit Sub
End If
tmp = CmbSGodown.Text
Call TableMst_Godown(" where CloseDate is null and CMPID = " & mCMPID)
CmbSGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
   CmbSGodown.AddItem RsMst_Godown!GodownNo
   RsMst_Godown.MoveNext
Next
CmbSGodown.Text = tmp
End Sub

Private Sub CmbSGodown_LostFocus()
If CmbSGodown.Text = "" Then
   Exit Sub
End If
If tmp <> CmbSGodown.Text Then
   ''write if clear date
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo= '" & CmbSGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub
If LCase(CmbSStatus.Text) <> "pending" And LCase(CmbSStatus.Text) <> "approved" And LCase(CmbSStatus.Text) <> "rejected" And LCase(CmbSStatus.Text) <> "godown creation" Then
   MsgBox "Invalid selection!!! "
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & "Rpt_GodownViabilityVsAgreement.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message!!!")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & "Rpt_GodownViabilityVsAgreement.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
   mRow = I + 1 ' + 6
   For C = 0 To MSHFlexGrid1.Cols - 1
      oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
   Next
Next
oWB.Save
MsgBox ("Excel file created!!!")
oXL.Visible = True
End Sub

Private Sub CmdSearch_Click()
Dim wc As Variant
CmdExcel.Enabled = False
wc = ""

Call Grid_Head
wc = "  Where MGI.StartDate <= '" & Format(Date, "YYYY-MM-DD") & "' AND MGI.CloseDate >= '" & Format(Date, "YYYY-MM-DD") & "' And MGI.Type = 'R' And MGI.UnitTypeID = 1 And MGI.Status = 'G' AND (MG.CloseDate >= '" & Format(Date, "YYYY-MM-DD") & "' Or MG.CloseDate is null) "

If CmbSLocation.Text <> "" Then
   If wc = "" Then
      wc = " Where MGI.LocationID = " & mLocationID
   Else
      wc = wc & " And MGI.LocationID = " & mLocationID
   End If
ElseIf Gen_WcLocation <> "" Then
   If wc = "" Then
      wc = " Where MGI." & Gen_WcLocation
   Else
      wc = wc & " And MGI." & Gen_WcLocation
   End If
End If

If CmbSCmp.Text <> "" Then
   If wc = "" Then
      wc = " Where MGI.CMPID = " & mCMPID
   Else
      wc = wc & " And MGI.CMPID = " & mCMPID
   End If
ElseIf Gen_WcCMP <> "" Then
   If wc = "" Then
      wc = " Where MGI." & Gen_WcCMP
   Else
      wc = wc & " And MGI." & Gen_WcCMP
   End If
End If

If CmbSGodown.Text <> "" Then
   If wc = "" Then
      wc = " Where MGI.GodownID = " & mGodownID
   Else
      wc = wc & " And MGI.GodownID = " & mGodownID
   End If
End If

If TxtSViabilityID <> "" Then
   If wc = "" Then
      wc = " Where MGI.InspectionID = " & Val(TxtSViabilityID) & ""
   Else
      wc = wc & " And MGI.InspectionID = " & Val(TxtSViabilityID) & ""
   End If
End If

tmp = " Select MGI.InspectionID, MGI.Type"
tmp = tmp & " ,ML.LocationName, MGI.LocationID ,MCMP.CMPName, MGI.CMPID , isnull(MGI.GodownNo, MG.GodownNo) GodownNo, MGI.GodownID, MUT.UnitType, MGI.UnitTypeID , MGI.GodownCategoryID , MGI.LicenseID,  MC.Category, MLC.LicenseNo, MGI.Capacity , MGI.Address , MGI.Area ,MGI.Status ,MGI.StartDate ,MGI.CloseDate ,MGI.CapacityUtilisation"
tmp = tmp & " ,MGI.NCDEXStock ,MGI.NNNRStock ,MGI.ReservationStock ,MGI.G_UID ,MGI.Concurrency ,MGI.CreatedDate ,MGI.CreatedBy ,MGI.UpdatedDate ,MGI.UpdatedBy"
tmp = tmp & " ,MGI.noofstaff ,MGI.staffsalary ,MGI.noofsec ,MGI.secsalary ,MGI.totalrevenue ,MGI.directexpense ,MGI.grossprofit, MGI.indirectexpense ,MGI.netprofit  ,MGI.CreatedDate,MGI.UpdatedDate "
tmp = tmp & " from Mst_GodownInspection MGI"
tmp = tmp & " Left Join Mst_Location ML on ML.LocationID = MGI.LocationID"
tmp = tmp & " Left Join Mst_CMP MCMP on MCMP.CMPID = MGI.CMPID"
tmp = tmp & " Left Join Mst_Godown MG on MG.GodownID = MGI.GodownID"
tmp = tmp & " Left Join Mst_UnitType MUT on MUT.UnitTypeID = MGI.UnitTypeID"
tmp = tmp & " Left Join Mst_Category MC on MC.CategoryID = MGI.GodownCategoryID"
tmp = tmp & " Left Join Mst_License MLC on MLC.LicenseID = MGI.LicenseID"
tmp = tmp & wc & " Order By MGI.InspectionID"
Call TmpTable

tmp = " SELECT TLA.LAID,TLA.StartDate,TLA.ExpiryDate,TLA.Flag,TGAD.GLADID,TGAD.GodownID,LessorID "
tmp = tmp & " FROM Txn_LessorAgreement TLA Inner Join Txn_GodownLessorAgreeDetail TGAD On TLA.LAID = TGAD.LAID "
tmp = tmp & " Where TLA.StartDate <= '" & Format(Date, "YYYY-MM-DD") & "' AND TLA.ExpiryDate >= '" & Format(Date, "YYYY-MM-DD") & "'"
Call Tmp1Table

If TmpRs.RecordCount > 0 Then
    With MSHFlexGrid1
       ProgressBar1.Max = TmpRs.RecordCount
       For iu = 1 To TmpRs.RecordCount
          If TmpRs1.RecordCount > 0 Then
            If TmpRs.RecordCount > 0 Then
                TmpRs1.MoveFirst
                TmpRs1.Find "GodownID =" & IIf(IsNull(TmpRs!GodownID), 0, TmpRs!GodownID)
                If TmpRs1.EOF = True Then
                    mR = .Rows - 1
                    .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!InspectionID), "", TmpRs!InspectionID)
                    .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
                    .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
                    .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
                    .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!Capacity), "", TmpRs!Capacity)
                    .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!Area), "", TmpRs!Area)
                    .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
                    .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy"))
                    .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-yyyy"))
                    .Rows = .Rows + 1
                End If
            End If
          End If
          TmpRs.MoveNext
          ProgressBar1.Value = iu
       Next
    End With
End If
Label2.Caption = MSHFlexGrid1.Rows - 2
CmdExcel.Enabled = True
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
   .Clear
   .Cols = 9
   .Rows = 2
   .RowHeight(0) = 600
   .AllowUserResizing = flexResizeColumns
   .Font.Size = 10
   .WordWrap = True
   .TextMatrix(0, 0) = "Viability ID":              .ColWidth(0) = 1200
   .TextMatrix(0, 1) = "Location Name":             .ColWidth(1) = 1500
   .TextMatrix(0, 2) = "CMP Name":                  .ColWidth(2) = 1800
   .TextMatrix(0, 3) = "Godown No":                 .ColWidth(3) = 1200
   .TextMatrix(0, 4) = "Capacity (in MT)":          .ColWidth(4) = 1200
   .TextMatrix(0, 5) = "Area (in Sq.Ft)":           .ColWidth(5) = 1200
   .TextMatrix(0, 6) = "Address":                   .ColWidth(6) = 3500
   .TextMatrix(0, 7) = "Start Date":                .ColWidth(7) = 1500
   .TextMatrix(0, 8) = "Close Date":                .ColWidth(8) = 1500
End With
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   CmdSearch.Enabled = False
   CmdExcel.Enabled = False
   Exit Sub
End If
Call Grid_Head
CmdSearch.Enabled = True
CmdExcel.Enabled = False
End Sub

