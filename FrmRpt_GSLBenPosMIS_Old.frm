VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_GSLBenPosMIS 
   Caption         =   "GSL BenPos MIS"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   9000
      Left            =   120
      TabIndex        =   8
      Top             =   750
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8655
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   15266
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
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
      Height          =   750
      Left            =   120
      TabIndex        =   0
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
         Left            =   6735
         TabIndex        =   4
         Top             =   225
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
         Left            =   7950
         TabIndex        =   3
         Top             =   225
         Width           =   1215
      End
      Begin VB.ComboBox CmbLocationID 
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   5400
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
         Left            =   9165
         TabIndex        =   1
         Top             =   225
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
         Left            =   10800
         TabIndex        =   7
         Top             =   285
         Width           =   2415
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
         Left            =   13200
         TabIndex        =   6
         Top             =   285
         Width           =   90
      End
      Begin VB.Label LblLocationID 
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
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   765
      End
   End
End
Attribute VB_Name = "FrmRpt_GSLBenPosMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID As Integer
Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location("")

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub
Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID

End Sub
Private Sub CmdClear_Click()
Call Grid_Head
CmbLocationID = ""
CmdExcel.Enabled = False
Label2 = ""
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GSLBenPosMIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GSLBenPosMIS.xls")
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


Private Sub CmdGetReport_Click()
Dim x As Integer

Call Grid_Head

MsgBox "Wait till next message!!!"

Label2.Caption = ""

tmp = "Select * from Txn_BenposDateWise Where ImportDate = (Select Max(ImportDate) from Txn_BenposDateWise) "
Call Tmp1Table


tmp = "Select ML.LocationName,MI.ISINID,Mcm.Commodity, sum(BalanceWeight) as 'GSLBalance' from Mst_GSL "
tmp = tmp & " inner join Mst_ISIN MI on Mst_GSL.ISINID=MI.ISINID"
tmp = tmp & " inner join Mst_CMP MC on MI.CMPID=MC.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " inner join Mst_Commodity MCm on MI.CommodityID=MCm.CommodityID"
tmp = tmp & " Where Mst_GSL.ExchangeTypeID = 2 and Mst_GSL.Flag <> 'E' And Mst_GSL.Flag <> 'J' and Mst_GSL.Flag <> 'Z' And MI.ISINID <> '' "
If CmbLocationID.Text <> "" Then
   tmp = tmp & " And ML.LocationName='" & CmbLocationID.Text & "'"
End If
tmp = tmp & " Group by ML.LocationName,MI.ISINID,Mcm.Commodity"
tmp = tmp & " Order By ML.LocationName,Mcm.Commodity,MI.ISINID"

Call TmpTable

x = 2
MSHFlexGrid1.Rows = TmpRs.RecordCount + 3

For i = 1 To TmpRs.RecordCount
    MSHFlexGrid1.TextMatrix(x, 0) = TmpRs!LocationName
    MSHFlexGrid1.TextMatrix(x, 1) = TmpRs!ISINID
    MSHFlexGrid1.TextMatrix(x, 2) = TmpRs!Commodity
    MSHFlexGrid1.TextMatrix(x, 4) = TmpRs!GSLBalance
    
    If TmpRs1.RecordCount > 0 Then
       TmpRs1.MoveFirst
    End If
    TmpRs1.Find "ISINID = '" & MSHFlexGrid1.TextMatrix(x, 1) & "'"
    
    If Not TmpRs1.EOF = True Then
       MSHFlexGrid1.TextMatrix(x, 3) = TmpRs1!Qty
    Else
       MSHFlexGrid1.TextMatrix(x, 3) = 0
    End If
    
    MSHFlexGrid1.TextMatrix(x, 5) = Format(Val(MSHFlexGrid1.TextMatrix(x, 3)) - Val(MSHFlexGrid1.TextMatrix(x, 4)), "#######0.000")
    MSHFlexGrid1.TextMatrix(0, 3) = Format(Val(MSHFlexGrid1.TextMatrix(0, 3)) + Val(MSHFlexGrid1.TextMatrix(x, 3)), "#######0.000")
    MSHFlexGrid1.TextMatrix(0, 4) = Format(Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(x, 4)), "#######0.000")
    MSHFlexGrid1.TextMatrix(0, 5) = Format(Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(x, 5)), "#######0.000")
    TmpRs.MoveNext
    x = x + 1
Next

If CmbLocationID.Text = "" Then

   TmpRs.MoveFirst
   TmpRs1.MoveFirst
   For i = 1 To TmpRs1.RecordCount
       If TmpRs.RecordCount > 0 Then
          TmpRs.MoveFirst
       End If
   
       TmpRs.Find "ISINID = '" & TmpRs1!ISINID & "'"
       If TmpRs.EOF = True Then
          MSHFlexGrid1.TextMatrix(x, 0) = "" 'TmpRs!LocationName
          MSHFlexGrid1.TextMatrix(x, 1) = TmpRs1!ISINID
          MSHFlexGrid1.TextMatrix(x, 2) = "" 'TmpRs!Commodity
          MSHFlexGrid1.TextMatrix(x, 4) = 0
          MSHFlexGrid1.TextMatrix(x, 3) = TmpRs1!Qty
          MSHFlexGrid1.TextMatrix(x, 5) = Format(Val(MSHFlexGrid1.TextMatrix(x, 3)) - Val(MSHFlexGrid1.TextMatrix(x, 4)), "#######0.000")
          MSHFlexGrid1.TextMatrix(0, 3) = Format(Val(MSHFlexGrid1.TextMatrix(0, 3)) + Val(MSHFlexGrid1.TextMatrix(x, 3)), "#######0.000")
          MSHFlexGrid1.TextMatrix(0, 4) = Format(Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(x, 4)), "#######0.000")
          MSHFlexGrid1.TextMatrix(0, 5) = Format(Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(x, 5)), "#######0.000")
          MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
          x = x + 1
       
       End If
       TmpRs1.MoveNext
   Next
End If

CmdExcel.Enabled = True
Label2.Caption = MSHFlexGrid1.Rows - 3
Label2.Refresh

MsgBox "Done!!!"

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
Call Grid_Head
CmdExcel.Enabled = False
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.FixedRows = 2
MSHFlexGrid1.Cols = 6

MSHFlexGrid1.TextMatrix(1, 0) = "Location"
MSHFlexGrid1.TextMatrix(1, 1) = "ISINID"
MSHFlexGrid1.TextMatrix(1, 2) = "Commodity"
MSHFlexGrid1.TextMatrix(1, 3) = "BenPos Balance"
MSHFlexGrid1.TextMatrix(1, 4) = "GSL Balance"
MSHFlexGrid1.TextMatrix(1, 5) = "Difference"

MSHFlexGrid1.ColWidth(0) = 2800
MSHFlexGrid1.ColWidth(1) = 2400
MSHFlexGrid1.ColWidth(2) = 3400
MSHFlexGrid1.ColWidth(3) = 1800
MSHFlexGrid1.ColWidth(4) = 1800
MSHFlexGrid1.ColWidth(5) = 1800

End Sub

