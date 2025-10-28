VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_LocationRentMIS 
   Caption         =   "Location Rent MIS"
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
      Height          =   8655
      Left            =   120
      TabIndex        =   5
      Top             =   855
      Width           =   15030
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   6
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
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         Left            =   1785
         TabIndex        =   2
         Top             =   225
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
         Height          =   495
         Left            =   585
         TabIndex        =   1
         Top             =   225
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Left            =   6120
         TabIndex        =   4
         Top             =   360
         Width           =   75
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
         Left            =   3600
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmRpt_LocationRentMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdGetReport_Click()

tmp = "Select c.LocationName, min(a.Amount/a.Area) 'MinRent',max(a.Amount/a.Area) 'MaxRent',sum(a.Amount)/sum(a.Area) 'AvgRent'" & _
      " From Mst_Location c" & _
      " Inner Join Mst_CMP b on c.LocationID  =  b.LocationID" & _
      " Inner Join Mst_Godown a  on a.cmpid =  b.cmpid" & _
      " Where UnitTypeID = 1" & _
      " Group by c.LocationName"

Call TmpTable

Call dispdata

MsgBox ("Done!!!")

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_LocationRentMIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_LocationRentMIS.xls")
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

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 5
MSHFlexGrid1.Font.Size = 10
MSHFlexGrid1.RowHeight(0) = 400

MSHFlexGrid1.TextMatrix(0, 0) = "SR NO"
MSHFlexGrid1.TextMatrix(0, 1) = "Location name"
MSHFlexGrid1.TextMatrix(0, 2) = "Min. Rent Per Sq.Ft"
MSHFlexGrid1.TextMatrix(0, 3) = "Max. Rent Per Sq.Ft"
MSHFlexGrid1.TextMatrix(0, 4) = "Avg. Rent Per Sq.Ft"


MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 2500
MSHFlexGrid1.ColWidth(2) = 2500
MSHFlexGrid1.ColWidth(3) = 2500
MSHFlexGrid1.ColWidth(4) = 2500



End Sub
Public Sub dispdata()
Dim x As Double
MsgBox ("Wait till Next Message")
x = 1



If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       x = MSHFlexGrid1.Rows
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       MSHFlexGrid1.TextMatrix(I, 0) = I
       MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!Locationname
       MSHFlexGrid1.TextMatrix(I, 2) = Format(TmpRs!MinRent, "#######0.00")
       MSHFlexGrid1.TextMatrix(I, 3) = Format(TmpRs!MaxRent, "#######0.00")
       MSHFlexGrid1.TextMatrix(I, 4) = Format(TmpRs!AvgRent, "#######0.00")
         
       
       
'       For c = 0 To TmpRs.Fields.Count - 1
'           MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
'       Next
       TmpRs.MoveNext
   Next
   Label2.Refresh
   Label2.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
Else
   Call Grid_Head
End If

End Sub



