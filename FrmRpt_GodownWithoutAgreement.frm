VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_GodownWithoutAgreement 
   Caption         =   "Godown Without Agreement"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   135
      TabIndex        =   2
      Top             =   0
      Width           =   15015
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
         Left            =   120
         TabIndex        =   4
         Top             =   285
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
         Height          =   375
         Left            =   1335
         TabIndex        =   3
         Top             =   285
         Width           =   1215
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
         Left            =   3960
         TabIndex        =   6
         Top             =   360
         Width           =   2415
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
         Left            =   6480
         TabIndex        =   5
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   855
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14631
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_GodownWithoutAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGetReport_Click()
Call Grid_Head

tmp = "Select ML.LocationName,MC.CMPName,MG.GodownNo,MG.Address,MU.UnitType,MG.BagsCapacity,MG.GodownCapacity,"
tmp = tmp & " MG.Area,MG.StartDate, MG.AgreementEndDate, MG.Amount, MG.NCDEXAccreditedQty,MG.GodownID"
tmp = tmp & " From Mst_Godown  MG"
tmp = tmp & " Left Join Txn_GodownLessorAgreeDetail TGLAD On MG.GodownID = TGLAD.GodownID"
tmp = tmp & " Inner join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_UnitType MU On MG.UnitTypeID = MU.UnitTypeID"
tmp = tmp & " Where MG.CloseDate is Null And MU.UnitType = 'leased' and TGLAD.Godownid is null"
tmp = tmp & " Order by ML.LocationName,MC.CMPName,MG.GodownNo"

Call TmpTable
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 1
   For mRow = 1 To TmpRs.RecordCount
       With MSHFlexGrid1
            .TextMatrix(mRow, 0) = TmpRs!LocationName
            .TextMatrix(mRow, 1) = TmpRs!CMPName
            .TextMatrix(mRow, 2) = TmpRs!GodownNo
            .TextMatrix(mRow, 3) = TmpRs!Address
            .TextMatrix(mRow, 4) = TmpRs!UnitType
            .TextMatrix(mRow, 5) = TmpRs!BagsCapacity
            .TextMatrix(mRow, 6) = Format(TmpRs!GodownCapacity, "########0.000")
            .TextMatrix(mRow, 7) = Format(TmpRs!Area, "########0.000")
            .TextMatrix(mRow, 8) = Format(TmpRs!StartDate, "dd-MMM-yyyy")
            '.TextMatrix(mRow, 9) = Format(TmpRs!AgreementEndDate, "dd-MMM-yyyy")  '--Commented on 24 Feb 12 as mail by sankar Sir
            .TextMatrix(mRow, 9) = Format(TmpRs!Amount, "########0.00")
            .TextMatrix(mRow, 10) = Format(TmpRs!NCDEXAccreditedQty, "########0.000")
            .TextMatrix(mRow, 11) = TmpRs!GodownID
       End With
       TmpRs.MoveNext
   Next
   LblTotalRecord.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
   MsgBox "Done !!!"
Else
   MsgBox "Record Not Found !!!"
   CmdExcel.Enabled = False
End If
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   CmdGetReport.Enabled = False
   CmdExcel.Enabled = False
   Exit Sub
End If
Call Grid_Head
CmdExcel.Enabled = False
End Sub
Private Sub Grid_Head()
With MSHFlexGrid1
     .Rows = 2
     .Cols = 12
     .AllowUserResizing = flexResizeBoth
     .TextMatrix(mRow, 0) = "Location Name"
     .TextMatrix(mRow, 1) = "CMP Name"
     .TextMatrix(mRow, 2) = "Godown No"
     .TextMatrix(mRow, 3) = "Address"
     .TextMatrix(mRow, 4) = "Unite Type"
     .TextMatrix(mRow, 5) = "Bags Capacity"
     .TextMatrix(mRow, 6) = "Godown Capacity"
     .TextMatrix(mRow, 7) = "Area"
     .TextMatrix(mRow, 8) = "Godown Start Date"
     '.TextMatrix(mRow, 9) = "Agreement End Date"
     .TextMatrix(mRow, 9) = "Amount"
     .TextMatrix(mRow, 10) = "NCDEX Accredited Qty"
     .TextMatrix(mRow, 11) = "GodownID"
     .ColWidth(0) = 1800
     .ColWidth(1) = 1800
     .ColWidth(2) = 1000
     .ColWidth(3) = 3500
     .ColWidth(4) = 1000
     .ColWidth(5) = 1400
     .ColWidth(6) = 1400
     .ColWidth(7) = 1400
     .ColWidth(8) = 1400
     .ColWidth(9) = 1400
     .ColWidth(10) = 1800
     .ColWidth(11) = 0
     .RowHeight(0) = 400
End With
End Sub
Private Sub CmdExcel_Click()
If MSHFlexGrid1.Rows = 2 Then Exit Sub
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownWithoutAgreement.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GodownWithoutAgreement.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

