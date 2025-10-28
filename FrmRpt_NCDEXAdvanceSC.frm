VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRpt_NCDEXAdvanceSC 
   Caption         =   "NCDEX Advance Storage Charges"
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
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   14775
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
         Left            =   3240
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   3855
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
         Height          =   420
         Left            =   9720
         TabIndex        =   5
         Top             =   540
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
         Height          =   420
         Left            =   11055
         TabIndex        =   4
         Top             =   540
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
         Height          =   420
         Left            =   12390
         TabIndex        =   3
         Top             =   540
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
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
         CustomFormat    =   "MM/dd/yyyy hh:mm:ss"
         Format          =   61079555
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
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
         CustomFormat    =   "MM/dd/yyyy hh:mm:ss"
         Format          =   61079555
         CurrentDate     =   39850
         MinDate         =   38353
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
         Height          =   255
         Left            =   4740
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "To Date"
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
         Left            =   1883
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   255
         TabIndex        =   10
         Top             =   240
         Width           =   1095
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
         Left            =   12120
         TabIndex        =   9
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label5 
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
         Left            =   9720
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   14775
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8055
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   14208
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmRpt_NCDEXAdvanceSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mLocationName As Variant
Dim mTotalDepBags, mTotalDepWt, mFinalDepBags, mFinalDepWt As Variant

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
Call TableMst_Location("")

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocation.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp

End Sub


Private Sub CmbLocation_LostFocus()

If CmbLocation.Text = "" Then
   mLocationID = 0
   Exit Sub
End If


RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID

End Sub

Private Sub CmdClear_Click()
Call Grid_Head
CmdExcel.Enabled = False
CmbLocation.Text = ""
TxtFromDate.Value = Date
TxtToDate.Value = Date
End Sub

Private Sub CmdExcel_Click()
If MSHFlexGrid1.Rows < 3 And MSHFlexGrid1.TextMatrix(1, 0) = "" Then MsgBox "No Record Found": Exit Sub
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NCDEXAdvanceSC.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NCDEXAdvanceSC.xls")
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


Private Sub CmdGetReport_Click()
Call Grid_Head

''tmp = "SELECT  Mst_Location.LocationName,Mst_CMP.CMPName, Mst_Commodity.Commodity,Mst_Client.ClientName, "
''tmp = tmp & " Mst_GSL.CdtfNo , Mst_GSL.CDTFDate, Sum(Mst_GSL.DepositeBags) 'Deposite Bags',"
''tmp = tmp & " Sum (Mst_GSL.DepositeWeight) 'Deposite Weight',Sum(Mst_GSL.CRAmount)'Amount',Mst_GSL.CRDate"
''tmp = tmp & " From Mst_GSL"
''tmp = tmp & " Inner JOIN Mst_CMP on Mst_CMP.CMPID = Mst_GSL.CMPID"
''tmp = tmp & " Inner join Mst_Location on Mst_Location.LocationID=Mst_CMP.LocationID"
''tmp = tmp & " Inner JOIN Mst_Commodity on Mst_Commodity.CommodityID = Mst_GSL.CommodityID"
''tmp = tmp & " Inner JOIN Mst_Client on Mst_Client.ClientIDRow = Mst_GSL.ClientIDRow"
''tmp = tmp & " Where (Mst_GSL.CRAmount Is Null Or Mst_GSL.CRAmount = 0) And Mst_GSL.ExchangeTypeID = 2"



tmp = "SELECT  Mst_Location.LocationName, "
tmp = tmp & " Mst_GSL.CdtfNo , Mst_GSL.CDTFDate, Sum(Mst_GSL.DepositeBags) 'Deposite Bags',"
tmp = tmp & " Sum (Mst_GSL.DepositeWeight) 'Deposite Weight'"
tmp = tmp & " From Mst_GSL"
tmp = tmp & " Inner JOIN Mst_CMP on Mst_CMP.CMPID = Mst_GSL.CMPID"
tmp = tmp & " Inner join Mst_Location on Mst_Location.LocationID=Mst_CMP.LocationID"
tmp = tmp & " Inner JOIN Mst_Commodity on Mst_Commodity.CommodityID = Mst_GSL.CommodityID"
tmp = tmp & " Inner JOIN Mst_Client on Mst_Client.ClientIDRow = Mst_GSL.ClientIDRow"
tmp = tmp & " Where (Mst_GSL.CRAmount Is Null Or Mst_GSL.CRAmount = 0) And Mst_GSL.ExchangeTypeID = 2"



If CmbLocation.Text <> "" Then
   tmp = tmp & " And Mst_Location.LocationID=" & mLocationID & ""
End If
If Gen_DBType = "MDB" Then
   tmp = tmp & " And (CDTFDate between #" & Format(TxtFromDate.Value, Gen_DatFormat) & "# And #" & Format(TxtToDate.Value, Gen_DatFormat) & "#)"
Else
   tmp = tmp & " And (CDTFDate between '" & Format(TxtFromDate.Value, Gen_DatFormat) & "'  And '" & Format(TxtToDate.Value, Gen_DatFormat) & "')"
End If

tmp = tmp & "  Group by  Mst_Location.LocationName,"
tmp = tmp & " Mst_GSL.CDTFNo, Mst_GSL.CDTFDate"
Call TmpTable

mLocationName = ""
mRow = 1
mTotalDepBags = 0
mTotalDepWt = 0
mFinalDepBags = 0
mFinalDepWt = 0
mrec = TmpRs.RecordCount + 1

If TmpRs.RecordCount > 0 Then
   For m = 1 To mrec
       If m > TmpRs.RecordCount Then
          MSHFlexGrid1.TextMatrix(mRow, 0) = mLocationName & " Total"
          MSHFlexGrid1.TextMatrix(mRow, 3) = mTotalDepBags
          MSHFlexGrid1.TextMatrix(mRow, 4) = mTotalDepWt
          mFinalDepBags = mFinalDepBags + mTotalDepBags
          mFinalDepWt = mFinalDepWt + mTotalDepWt
          Exit For
       End If
   
       If mLocationName <> TmpRs!LocationName Then
          If mLocationName <> "" Then
             MSHFlexGrid1.TextMatrix(mRow, 0) = mLocationName & " Total"
             MSHFlexGrid1.TextMatrix(mRow, 3) = mTotalDepBags
             MSHFlexGrid1.TextMatrix(mRow, 4) = mTotalDepWt
             mFinalDepBags = mFinalDepBags + mTotalDepBags
             mFinalDepWt = mFinalDepWt + mTotalDepWt
             
             MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
             mRow = mRow + 1
                
             mLocationName = TmpRs!LocationName
             MSHFlexGrid1.TextMatrix(mRow, 0) = mLocationName
             mTotalDepBags = 0
             mTotalDepWt = 0
             
           Else
             mLocationName = TmpRs!LocationName
             MSHFlexGrid1.TextMatrix(mRow, 0) = mLocationName
          End If
       End If
              
       ''TO FILL COLUMNS
       For mCol = 1 To MSHFlexGrid1.Cols - 1
           If InStr(TmpRs.Fields(mCol), "/") > 0 Then
              If IsDate(TmpRs.Fields(mCol)) Then
                 MSHFlexGrid1.TextMatrix(mRow, mCol) = Format(TmpRs.Fields(mCol), "dd-MMM-yyyy")
              Else
                 MSHFlexGrid1.TextMatrix(mRow, mCol) = IIf(IsNull(TmpRs.Fields(mCol)), "", TmpRs.Fields(mCol))
              End If
           Else
              MSHFlexGrid1.TextMatrix(mRow, mCol) = IIf(IsNull(TmpRs.Fields(mCol)), "", TmpRs.Fields(mCol))
           End If
           
       Next
       mTotalDepBags = mTotalDepBags + Val(MSHFlexGrid1.TextMatrix(mRow, 3))
       mTotalDepWt = mTotalDepWt + Val(MSHFlexGrid1.TextMatrix(mRow, 4))
              
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       mRow = mRow + 1
       TmpRs.MoveNext
   Next
   LblTotalRecord.Caption = TmpRs.RecordCount
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
   mRow = mRow + 1
   MSHFlexGrid1.TextMatrix(mRow, 0) = "Grand Total"
   MSHFlexGrid1.TextMatrix(mRow, 3) = mFinalDepBags
   MSHFlexGrid1.TextMatrix(mRow, 4) = mFinalDepWt
   MsgBox "Done !!!", vbInformation
Else
   MsgBox " No Record Found "
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

Call Grid_Head
TxtFromDate.Value = Date
TxtToDate.Value = Date
End Sub


Private Sub Grid_Head()

With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 5
    .AllowUserResizing = flexResizeBoth
    .TextMatrix(0, 0) = "Location"
    .TextMatrix(0, 1) = "CDTF NO"
    .TextMatrix(0, 2) = "CDTF Date"
    .TextMatrix(0, 3) = "Deposited Bags"
    .TextMatrix(0, 4) = "Deposited Weight"
    .ColWidth(0) = 2500
    .ColWidth(1) = 2000
    .ColWidth(2) = 2000
    .ColWidth(3) = 2500
    .ColWidth(4) = 2500
    .RowHeight(0) = 400
End With

End Sub
