VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_GSLISINLocationMisMatch 
   Caption         =   "Mis Match In GSL ISIN Location"
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
      Height          =   8535
      Left            =   120
      TabIndex        =   6
      Top             =   1230
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8175
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14420
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
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
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
         Height          =   495
         Left            =   1695
         TabIndex        =   3
         Top             =   360
         Width           =   1455
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
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1455
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
         Height          =   495
         Left            =   3150
         TabIndex        =   1
         Top             =   360
         Width           =   1455
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
         Left            =   8280
         TabIndex        =   5
         Top             =   480
         Width           =   90
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
         Left            =   5760
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmRpt_GSLISINLocationMisMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClear_Click()
Call Grid_Head1
CmdExcel.Enabled = False
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_MisMatchGSLISINLocation.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_MisMatchGSLISINLocation.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "dd/MMM/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()

Call Grid_Head1
Label2.Caption = ""

'tmp = "Select MG.GSLID,MG.SM_Prefix,MG.ISINID,ML.LocationName as 'GSL Location',ML1.LocationName as 'ISINLocation',MG.Flag " & _
'      " From Mst_ISIN MI  " & _
'      " Inner Join  Mst_GSL MG On MI.ISINID = MG.ISINID   " & _
'      " Inner join  Mst_CMP MC On MG.CMPID = MC.CMPID   " & _
'      " Inner Join  Mst_Location ML On MC.LOcationID = ML.LocationID   " & _
'      " Inner Join  Mst_Location ML1 On MI.LocationID = ML1.LocationID   " & _
'      " Where  ML.LocationID <> ML1.LocationID  And MG.ExchangeTypeID = 2 " & _
'      " And MG.ISINID Not like '%Reject%' And MG.ISINID <> '' "


tmp = " Select MG.GSLID,MG.SM_Prefix,MG.ISINID,ML.LocationName as 'GSL Location',ML1.LocationName as 'ISINLocation',"
tmp = tmp & " MG.Flag , com.Commodity 'ISIN Commodity',    com1.Commodity 'GSL Commodity'"
tmp = tmp & " From Mst_ISIN MI"
tmp = tmp & " Inner Join  Mst_GSL MG On MI.ISINID = MG.ISINID"
tmp = tmp & " Inner join  Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Inner Join  Mst_Location ML On MC.LOcationID = ML.LocationID"
tmp = tmp & " Inner Join  Mst_Location ML1 On MI.LocationID = ML1.LocationID"
tmp = tmp & " Inner Join  Mst_Commodity com On MI.CommodityID = com.CommodityID"
tmp = tmp & " Inner Join  Mst_Commodity com1 On MG.CommodityID = com1.CommodityID"
tmp = tmp & " Where (ML.LocationID <> ML1.LocationID Or com1.CommodityID <> com.CommodityID) And MG.ExchangeTypeID = 2"
tmp = tmp & " And MG.ISINID Not like '%Reject%' And MG.ISINID <> ''"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For I = 1 To TmpRs.RecordCount
       For c = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       TmpRs.MoveNext
   Next
   CmdExcel.Enabled = True
   Label2.Caption = TmpRs.RecordCount
   Label2.Refresh
   MsgBox "Done!!"
Else
   MsgBox "Record Not Found."
End If

End Sub

Private Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 8
MSHFlexGrid1.FixedRows = 1
MSHFlexGrid1.FixedCols = 1
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(0, 0) = " GSL ID"
MSHFlexGrid1.TextMatrix(0, 1) = " SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 2) = " ISIN"
MSHFlexGrid1.TextMatrix(0, 3) = " GSL Location"
MSHFlexGrid1.TextMatrix(0, 4) = " ISIN Location"
MSHFlexGrid1.TextMatrix(0, 5) = " Flag"
MSHFlexGrid1.TextMatrix(0, 6) = " ISIN Commodity"
MSHFlexGrid1.TextMatrix(0, 7) = " GSL Commodity"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 1800
MSHFlexGrid1.ColWidth(2) = 1800
MSHFlexGrid1.ColWidth(3) = 1800
MSHFlexGrid1.ColWidth(4) = 1800
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 2200
MSHFlexGrid1.ColWidth(7) = 2200



End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True

Call Grid_Head1
CmdExcel.Enabled = False
End Sub


