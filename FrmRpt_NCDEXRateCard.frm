VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_NCDEXRateCard 
   Caption         =   "NCDEX Rate Card"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   75
      TabIndex        =   11
      Top             =   0
      Width           =   15000
      Begin VB.OptionButton Option2 
         Caption         =   "Not in Rate Card"
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
         Left            =   1920
         TabIndex        =   2
         Top             =   390
         Width           =   1890
      End
      Begin VB.OptionButton Option1 
         Caption         =   "In Rate Card"
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
         TabIndex        =   1
         Top             =   390
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.ComboBox CmbCMPName 
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
         Left            =   7830
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   405
         Width           =   6855
      End
      Begin VB.ComboBox CmbSLocationID 
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
         Top             =   405
         Width           =   3825
      End
      Begin VB.ComboBox CmbGodown 
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
         TabIndex        =   5
         Top             =   1065
         Width           =   3660
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
         Left            =   12645
         TabIndex        =   8
         Top             =   1560
         Width           =   1000
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
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
         Left            =   11640
         TabIndex        =   7
         Top             =   1560
         Width           =   1000
      End
      Begin VB.ComboBox CmbSCommodityID 
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
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1065
         Width           =   10740
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
         Height          =   350
         Left            =   13650
         TabIndex        =   9
         Top             =   1560
         Width           =   1000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Godowns"
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
         Left            =   150
         TabIndex        =   17
         Top             =   165
         Width           =   855
      End
      Begin VB.Label LblGodown 
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
         Left            =   1508
         TabIndex        =   15
         Top             =   840
         Width           =   885
      End
      Begin VB.Label lblCMP 
         AutoSize        =   -1  'True
         Caption         =   "CMP Name"
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
         Left            =   10740
         TabIndex        =   14
         Top             =   165
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Commodity"
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
         Left            =   8723
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Height          =   210
         Left            =   5430
         TabIndex        =   12
         Top             =   165
         Width           =   870
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   90
      TabIndex        =   0
      Top             =   2025
      Width           =   14970
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   6945
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6615
         Left            =   120
         TabIndex        =   10
         Top             =   255
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   11668
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
End
Attribute VB_Name = "FrmRpt_NCDEXRateCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mStateID, mLOcationID, mCMPID, mExchangeTypeID, mGodownID, mCommodityId As Variant

Private Sub CmbCMPName_GotFocus()
If CmbSLocationID.Text = "" Then
   MsgBox "Select Location!!!"
   CmbSLocationID.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text
Call TableMst_CMP(" where LocationID=" & mLOcationID & " And CloseDate is null ")
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "First Select Location!!!! "
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub

Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   CmbGodown.Text = ""
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If

If mCMPID <> RsMst_CMP!CMPID Then
   CmbGodown.Text = ""
End If

mCMPID = RsMst_CMP!CMPID
End Sub



Private Sub CmbGodown_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodown.Text
Call TableMst_Godown("Where CMPID = " & mCMPID & " And CloseDate is null ")
CmbGodown.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Godown.RecordCount
    CmbGodown.AddItem RsMst_Godown!godownno
    RsMst_Godown.MoveNext
Next
CmbGodown.Text = tmp
End Sub

Private Sub CmbGodown_LostFocus()
If CmbGodown.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodown.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbGodown.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub

Private Sub CmbSLocationID_GotFocus()
tmp = CmbSLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
End If

CmbSLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
   CmbSLocationID.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbSLocationID.Text = tmp
    
End Sub

'----
Private Sub CmbSLocationID_LostFocus()
If CmbSLocationID.Text = "" Then
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "Locationname = '" & CmbSLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocationID.SetFocus
   Exit Sub
End If
mLOcationID = RsMst_Location!LocationID

End Sub

Private Sub CmbSCommodityID_GotFocus()

tmp = CmbSCommodityID.Text
Call TableMst_CommodityGroup
CmbSCommodityID.Clear
If RsMst_CommodityGroup.RecordCount = 0 Then
  MsgBox "No Record found !!!! "
  Exit Sub
End If
For i = 1 To RsMst_CommodityGroup.RecordCount
   CmbSCommodityID.AddItem RsMst_CommodityGroup!COmmodityGroup
   RsMst_CommodityGroup.MoveNext
Next
CmbSCommodityID.Text = tmp
    
End Sub
'----
Private Sub CmbSCommodityID_LostFocus()
If CmbSCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_CommodityGroup.MoveFirst
RsMst_CommodityGroup.Find "CommodityGroup = '" & CmbSCommodityID.Text & "'"
If RsMst_CommodityGroup.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodityID.SetFocus
   Exit Sub
End If
mSCommodityID = RsMst_CommodityGroup!CommodityGroupID
End Sub

Private Sub CmdClear_Click()

CmbSLocationID.Text = ""
CmbCMPName.Text = ""
CmbGodown.Text = ""
CmbSCommodityID.Text = ""
Option1.Value = True
Call Grid_Head
ProgressBar1.Value = 0

End Sub

Private Sub CmdGo_Click()

Dim wc As String
Dim tmp1 As String
If Option2.Value = True Then
   Call Grid_Head1
Else
   Call Grid_Head
End If

ProgressBar1.Value = 0
MsgBox "Wait till next message!!!"

wc = ""

If CmbSLocationID.Text <> "" Then
   wc = GenWc(wc, CmbSLocationID.Text, "ML.LocationName", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp2 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, tmp2, "", "S", "I")
   End If
End If

If CmbCMPName.Text <> "" Then
   wc = GenWc(wc, CmbCMPName.Text, "MC.CMPName", "S", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp2 = Replace(Gen_WcCMP, "CMPID", "MC.CMPID")
      wc = GenWc(wc, tmp2, "", "S", "I")
   End If
End If

If CmbGodown.Text <> "" Then
   wc = GenWc(wc, CmbGodown.Text, "MG.GodownNo", "S", "N")
End If

If Option1.Value = True Then
   If CmbSCommodityID.Text <> "" Then
      wc = GenWc(wc, CmbSCommodityID.Text, "CommodityGroup", "S", "N")
   End If


   tmp = " Select ML.LocationName,MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.StartDate,MG.CloseDate,MCG.CommodityGroup,NRCD.Rate"
   tmp = tmp & " From Mst_NCDEXStorageRateCard NRC"
   tmp = tmp & " Left Join Mst_NCDEXStorageRateCardDetail NRCD on NRC.NSRCID = NRCD.NSRCID"
   tmp = tmp & " Left Join Mst_Godown MG On NRCD.GodownID = MG.GodownID"
   tmp = tmp & " Left Join Mst_CMP MC On MG.CMPID = MC.CMPID"
   tmp = tmp & " Left Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Left Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
   tmp = tmp & " Left Join Mst_CommodityGroup MCG On NRC.CommodityGroupID = MCG.CommodityGroupID"
      
   tmp1 = " Order By ML.LocationName,MC.CMPName,MG.GodownNo"
   
   tmp = tmp & " " & wc & " " & tmp1
   
   Call TmpTable
      
   If TmpRs.RecordCount > 0 Then
      ProgressBar1.Max = TmpRs.RecordCount + 2
      MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
      For i = 1 To TmpRs.RecordCount
          For C = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
          Next
          MSHFlexGrid1.TextMatrix(i, 5) = Format(MSHFlexGrid1.TextMatrix(i, 5), "dd-mmm-yyyy")
          MSHFlexGrid1.TextMatrix(i, 6) = Format(MSHFlexGrid1.TextMatrix(i, 6), "dd-mmm-yyyy")
          TmpRs.MoveNext
          ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End If
   ProgressBar1.Value = ProgressBar1.Max
Else

   tmp = " Select ML.LocationName,MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.StartDate,MG.CloseDate ,MG.GodownID"
   tmp = tmp & " From Mst_Godown  MG"
   tmp = tmp & " Left Join Mst_CMP MC On MG.CMPID = MC.CMPID"
   tmp = tmp & " Left Join Mst_Location ML On MC.LocationID = ML.LocationID"
   tmp = tmp & " Left Join Mst_UnitType MU on MG.UnitTypeID = MU.UnitTypeID"
   
   
   If wc = "" Then
      tmp1 = " Where MG.GodownID not in (Select GodownID from Mst_NCDEXStorageRateCardDetail) And MG.CloseDate is null"
      tmp1 = tmp1 & " And MG.GodownID in (Select Distinct GodownID From Mst_GSL Where ExchangeTypeID = 2)"
   Else
      tmp1 = " And MG.GodownID not in (Select GodownID from Mst_NCDEXStorageRateCardDetail) And MG.CloseDate is null"
      tmp1 = tmp1 & " And MG.GodownID in (Select Distinct GodownID From Mst_GSL Where ExchangeTypeID = 2)"
   End If
   
   tmp1 = tmp1 & " Order By ML.LocationName,MC.CMPName,MG.GodownNo"
   
   tmp = tmp & " " & wc & " " & tmp1
   
   Call TmpTable
      
   If TmpRs.RecordCount > 0 Then
      ProgressBar1.Max = TmpRs.RecordCount + 2
      MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
      For i = 1 To TmpRs.RecordCount
          For C = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
          Next
          MSHFlexGrid1.TextMatrix(i, 5) = Format(MSHFlexGrid1.TextMatrix(i, 5), "dd-mmm-yyyy")
          MSHFlexGrid1.TextMatrix(i, 6) = Format(MSHFlexGrid1.TextMatrix(i, 6), "dd-mmm-yyyy")
          TmpRs.MoveNext
          ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End If
   ProgressBar1.Value = ProgressBar1.Max


End If



MsgBox "Done!!!"

End Sub

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Frame3.Enabled = False
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True


Call CmdClear_Click
  

End Sub
Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .WordWrap = True
    .Cols = 9
       
    .TextMatrix(0, 0) = "Location"
    .TextMatrix(0, 1) = "CMP Name"
    .TextMatrix(0, 2) = "Godown No"
    .TextMatrix(0, 3) = "Godown address"
    .TextMatrix(0, 4) = "Godown type"
    .TextMatrix(0, 5) = "Godown Start Date"
    .TextMatrix(0, 6) = "Godown Close date"
    .TextMatrix(0, 7) = "Commodity"
    .TextMatrix(0, 8) = "Rate"
    
    
    'Location    CMP name    Godown No   Godown address  Unit type   Start Date  Close date  Commodity   Rate
    
    .ColWidth(0) = 1600
    .ColWidth(1) = 1600
    .ColWidth(2) = 1400
    .ColWidth(3) = 2000
    .ColWidth(4) = 2200
    .ColWidth(5) = 1400
    .ColWidth(6) = 1400
    .ColWidth(7) = 1400
    .ColWidth(8) = 1000
       
    
    .RowHeight(0) = 700
End With
End Sub
Private Sub Grid_Head1()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .WordWrap = True
    .Cols = 7
       
    .TextMatrix(0, 0) = "Location"
    .TextMatrix(0, 1) = "CMP Name"
    .TextMatrix(0, 2) = "Godown No"
    .TextMatrix(0, 3) = "Godown address"
    .TextMatrix(0, 4) = "Godown type"
    .TextMatrix(0, 5) = "Godown Start Date"
    .TextMatrix(0, 6) = "Godown Close date"
   
    
    'Location    CMP name    Godown No   Godown address  Unit type   Start Date  Close date  Commodity   Rate
    
    .ColWidth(0) = 1600
    .ColWidth(1) = 1600
    .ColWidth(2) = 1400
    .ColWidth(3) = 2000
    .ColWidth(4) = 2200
    .ColWidth(5) = 1400
    .ColWidth(6) = 1400
   
    .RowHeight(0) = 700

End With
End Sub


Private Sub CmdExcel_Click()
'If MSHFlexGrid1.TextMatrix(1, 0) <> "" Then
   Gen_mTimeStamp = GetTimeStamp
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NCDEXRateCard.xls")
   Call Xls_Detail_Rpt
'Else
'   MsgBox "No Record To Export"
'End If
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NCDEXRateCard.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(i, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub


Private Sub Option1_Click()

If Option2.Value = True Then
   Call Grid_Head1
Else
   Call Grid_Head
End If

End Sub

Private Sub Option2_Click()
Call Option1_Click
End Sub
