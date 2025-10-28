VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_GSLCommodityMap 
   Caption         =   "GSL Commodity Map Report"
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
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   15075
      Begin VB.OptionButton Option1 
         Caption         =   "Commodity Mapped to GSL"
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
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1770
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All Commodities"
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
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.ComboBox CmbExchangeTypeID 
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
         ItemData        =   "FrmRpt_GSLCommodityMap.frx":0000
         Left            =   5400
         List            =   "FrmRpt_GSLCommodityMap.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2415
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
         Left            =   9135
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdGetData 
         Caption         =   "GetData"
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
         Left            =   7920
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
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
         Left            =   10350
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Exchange Type"
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
         Left            =   3840
         TabIndex        =   10
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Height          =   375
         Left            =   11640
         TabIndex        =   9
         Top             =   240
         Width           =   2205
      End
      Begin VB.Label LblTotalRecords 
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
         Left            =   14025
         TabIndex        =   8
         Top             =   255
         Width           =   75
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   885
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   15266
      _Version        =   393216
      WordWrap        =   -1  'True
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
Attribute VB_Name = "FrmRpt_GSLCommodityMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mExchangeTypeID As Variant
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdGetData_Click()
LblTotalRecords.Caption = ""

MsgBox ("Wait till Next Message")
Call GridHead

If Option1.Value = True Then

   tmp = "Select distinct b.commodity from mst_gsl a"
   tmp = tmp & " Inner join mst_commodity b on  a.CommodityID = b.CommodityID"
   tmp = tmp & " Order by b.Commodity"
   Call TmpTable
   
End If
tmp = "select a.Commodity,c.ExchangeType,a.Minimum,a.Maximum,b.GunnyWt,b.StdDeduction,b.LotSize,b.Variation,b.SCRate"
tmp = tmp & " From mst_commodity a"
tmp = tmp & " Inner join Mst_CommodityDetail b  on a.CommodityID = b.CommodityID"
tmp = tmp & " Inner join mst_exchangetype c on b.ExchangeTypeID = c.ExchangeTypeID"

If CmbExchangeTypeID.Text <> "" Then
   tmp = tmp & " Where c.ExchangeType='" & CmbExchangeTypeID.Text & "'"
End If

tmp = tmp & " Order by a.Commodity,"
tmp = tmp & " c.ExchangeType"

Call Tmp0Table

mRow = 1
If Option1.Value = True Then
   If TmpRs.EOF = False Then
      For I = 1 To TmpRs.RecordCount
          TmpRs0.Filter = "Commodity='" & TmpRs!Commodity & "' " 'and ExchangeType='" & cmbExchangeType.Text & "' "
          For j = 1 To TmpRs0.RecordCount
              For c = 0 To TmpRs0.Fields.Count - 1
                  MSHFlexGrid1.TextMatrix(mRow, c) = IIf(IsNull(TmpRs0.Fields(c)), "", TmpRs0.Fields(c))
              Next
              mRow = mRow + 1
              MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
              TmpRs0.MoveNext
          Next
          TmpRs.MoveNext
          TmpRs0.Filter = ""
      Next
   End If
Else
    For j = 1 To TmpRs0.RecordCount
        For c = 0 To TmpRs0.Fields.Count - 1
            MSHFlexGrid1.TextMatrix(mRow, c) = IIf(IsNull(TmpRs0.Fields(c)), "", TmpRs0.Fields(c))
        Next
        mRow = mRow + 1
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
        TmpRs0.MoveNext
    Next
End If

LblTotalRecords.Caption = MSHFlexGrid1.Rows - 2
MsgBox "Done!!!"
CmdExcel.Enabled = True
End Sub


Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   CmdGetData.Enabled = False
   CmdExcel.Enabled = False
   CmbExchangeTypeID.Enabled = False
   Exit Sub
End If

Call GridHead
CmdExcel.Enabled = False

End Sub
Private Function GridHead()
With MSHFlexGrid1
     .Clear
     .AllowUserResizing = flexResizeColumns
     .Cols = 9
     .Rows = 2
     .TextMatrix(0, 0) = "Commodity"
     .TextMatrix(0, 1) = "Exchange Type"
     .TextMatrix(0, 2) = "Minimum"
     .TextMatrix(0, 3) = "Maximum"
     .TextMatrix(0, 4) = "Gunny Wt"
     .TextMatrix(0, 5) = "STD Deduction"
     .TextMatrix(0, 6) = "Lot Size"
     .TextMatrix(0, 7) = "Variation"
     .TextMatrix(0, 8) = "SC Rate"

     .ColWidth(0) = 3400
     .ColWidth(1) = 2000
     .ColWidth(2) = 1000
     .ColWidth(3) = 1000
     .ColWidth(4) = 1000
     .ColWidth(5) = 1600
     .ColWidth(6) = 1500
     .ColWidth(7) = 1500
     .ColWidth(8) = 1500
End With
End Function
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GSLCommodityMap.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GSLCommodityMap.xls")
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
Private Sub CmbExchangeTypeID_GotFocus()
tmp = CmbExchangeTypeID.Text
Call TableMst_ExchangeType
CmbExchangeTypeID.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeTypeID.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeTypeID.Text = tmp
End Sub
'----
Private Sub CmbExchangeTypeID_LostFocus()
If CmbExchangeTypeID.Text = "" Then
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeTypeID.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid selection "
   CmbExchangeTypeID.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType!ExchangeTypeID
End Sub


