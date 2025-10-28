VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_WHRFundingMIS 
   Caption         =   "WHR Funding MIS"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   14985
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
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
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
         Left            =   1200
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   3165
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
         Left            =   7860
         TabIndex        =   7
         Top             =   240
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
         Height          =   375
         Left            =   6390
         TabIndex        =   2
         Top             =   240
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
         Height          =   375
         Left            =   4920
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         Height          =   300
         Left            =   180
         TabIndex        =   9
         Top             =   300
         Width           =   840
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
         Left            =   14040
         TabIndex        =   4
         Top             =   360
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
         Left            =   11520
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FrmRpt_WHRFundingMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID As Variant

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
If CmbLocation.Text = "" Then
   mLocationID = Null
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocation.SetFocus
   Exit Sub
End If


mLocationID = RsMst_Location!LocationID

End Sub

Private Function GridHead()
With MSHFlexGrid1
    
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 7
    .TextMatrix(1, 0) = "Location"
    .TextMatrix(1, 1) = "SM Balance"
    .TextMatrix(1, 2) = "WHR Balance"
    .TextMatrix(1, 3) = "Pledge Balance"
    .TextMatrix(1, 4) = "WHR as % of Stock"
    .TextMatrix(1, 5) = "Pledge as % of WHR"
    .TextMatrix(1, 6) = "Pledge as % of Stock"
    
    .ColWidth(0) = 2600
    .ColWidth(1) = 1800
    .ColWidth(2) = 1800
    .ColWidth(3) = 1800
    .ColWidth(4) = 1800
    .ColWidth(5) = 1800
    .ColWidth(6) = 1800
    
    .RowHeight(1) = 600
End With
End Function

Private Sub CmdClear_Click()
Call GridHead
Label2.Caption = ""
CmdExcel.Enabled = False
CmbLocation.Text = ""
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_WHRFundingMIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_WHRFundingMIS.xls")
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
Dim mRow As Variant
Dim wc As String

Dim tmp1 As String
tmp1 = ""

Call GridHead
Label2.Caption = ""
wc = ""

If CmbLocation.Text <> "" Then
   wc = GenWc(wc, mLocationID, "L.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "L.LocationID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If


tmp = "select GSL.SM_Prefix,L.LocationName,sum(BalanceBags) as BalanceBags,sum(BalanceWeight)as BalanceWeight "
tmp = tmp & " From Mst_GSL GSL Inner join Mst_Location L on L.SM_Prefix = GSL.SM_Prefix "
If wc <> "" Then
   tmp = tmp & wc
   tmp = tmp & " And GSL.Flag = 'O'  and ExchangeTypeID = 1 "
Else
   tmp = tmp & " where GSL.Flag = 'O'  and ExchangeTypeID = 1 "
End If
tmp = tmp & " Group By GSL.SM_Prefix,L.LocationName Order By L.LocationName"

Call TmpTable

tmp = "Select SM_Prefix,sum(WHRNoOfBags-DeliveredNoOfBags) As BalanceBags, sum(WHRQuantity - DeliveredQuantity) As BalanceWeight"
tmp = tmp & " From Txn_WHR_Detail"
'tmp = tmp & " Where WHRNoOfBags - DeliveredNoOfBags > 0 Or WHRQuantity - DeliveredQuantity > 0"
tmp = tmp & " Where WHRNo in (Select WHRNo from Txn_WHR_Detail Where WHRNoOfBags - DeliveredNoOfBags > 0 Or WHRQuantity - DeliveredQuantity > 0 )"
tmp = tmp & " Group by SM_Prefix Order by SM_Prefix"

Call Tmp1Table

tmp = "Select SM_Prefix,sum(WHRNoOfBags-DeliveredNoOfBags) As BalanceBags, sum(WHRQuantity - DeliveredQuantity) as BalanceWeight"
tmp = tmp & " From Txn_WHR_Detail"
'tmp = tmp & " Where (WHRNoOfBags - DeliveredNoOfBags > 0 Or WHRQuantity - DeliveredQuantity > 0) And PledgeID Is Not Null And PledgeID > 0"
tmp = tmp & " Where WHRNo in (Select WHRNo from Txn_WHR_Detail Where WHRNoOfBags - DeliveredNoOfBags > 0 Or WHRQuantity - DeliveredQuantity > 0 ) And PledgeID Is Not Null And PledgeID > 0"
tmp = tmp & " Group by SM_Prefix Order by SM_Prefix"

Call Tmp2Table

If TmpRs.RecordCount > 0 Then
   mRow = 2
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 3
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.TextMatrix(mRow, 0) = TmpRs!LocationName
       MSHFlexGrid1.TextMatrix(mRow, 1) = TmpRs!BalanceWeight
       If TmpRs1.RecordCount > 0 Then
          TmpRs1.MoveFirst
       End If
       TmpRs1.Find "SM_Prefix='" & TmpRs!SM_Prefix & "'"
       If Not TmpRs1.EOF Then
          MSHFlexGrid1.TextMatrix(mRow, 2) = TmpRs1!BalanceWeight
       Else
          MSHFlexGrid1.TextMatrix(mRow, 2) = 0
       End If
       If TmpRs2.RecordCount > 0 Then
          TmpRs2.MoveFirst
       End If
       TmpRs2.Find "SM_Prefix='" & TmpRs!SM_Prefix & "'"
       If Not TmpRs2.EOF Then
          MSHFlexGrid1.TextMatrix(mRow, 3) = TmpRs2!BalanceWeight
       Else
          MSHFlexGrid1.TextMatrix(mRow, 3) = 0
       End If
      
       
       If Val(MSHFlexGrid1.TextMatrix(mRow, 2)) <> 0 Then
          MSHFlexGrid1.TextMatrix(mRow, 5) = Round(Val(MSHFlexGrid1.TextMatrix(mRow, 3)) / Val(MSHFlexGrid1.TextMatrix(mRow, 2)) * 100, 2)
       Else
          MSHFlexGrid1.TextMatrix(mRow, 5) = 0
       End If
       
       If Val(MSHFlexGrid1.TextMatrix(mRow, 1)) <> 0 Then
          MSHFlexGrid1.TextMatrix(mRow, 6) = Round(Val(MSHFlexGrid1.TextMatrix(mRow, 3)) / Val(MSHFlexGrid1.TextMatrix(mRow, 1)) * 100, 2)
          MSHFlexGrid1.TextMatrix(mRow, 4) = Round(Val(MSHFlexGrid1.TextMatrix(mRow, 2)) / Val(MSHFlexGrid1.TextMatrix(mRow, 1)) * 100, 2)
       Else
          MSHFlexGrid1.TextMatrix(mRow, 6) = 0
          MSHFlexGrid1.TextMatrix(mRow, 4) = 0
       End If
       
       MSHFlexGrid1.TextMatrix(0, 1) = Val(MSHFlexGrid1.TextMatrix(0, 1)) + Val(MSHFlexGrid1.TextMatrix(mRow, 1))
       MSHFlexGrid1.TextMatrix(0, 2) = Val(MSHFlexGrid1.TextMatrix(0, 2)) + Val(MSHFlexGrid1.TextMatrix(mRow, 2))
       MSHFlexGrid1.TextMatrix(0, 3) = Val(MSHFlexGrid1.TextMatrix(0, 3)) + Val(MSHFlexGrid1.TextMatrix(mRow, 3))

       TmpRs.MoveNext
       mRow = mRow + 1
   Next
   MSHFlexGrid1.TextMatrix(0, 4) = Round(Val(MSHFlexGrid1.TextMatrix(0, 2)) / Val(MSHFlexGrid1.TextMatrix(0, 1)) * 100, 2)
   MSHFlexGrid1.TextMatrix(0, 5) = Round(Val(MSHFlexGrid1.TextMatrix(0, 3)) / Val(MSHFlexGrid1.TextMatrix(0, 2)) * 100, 2)
   MSHFlexGrid1.TextMatrix(0, 6) = Round(Val(MSHFlexGrid1.TextMatrix(0, 3)) / Val(MSHFlexGrid1.TextMatrix(0, 1)) * 100, 2)
     
   Label2.Caption = TmpRs.RecordCount
   Label2.Refresh
   CmdExcel.Enabled = True
   MsgBox "Done !!!"
Else
  MsgBox "No Record Found."
End If
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
Call GridHead
CmdExcel.Enabled = False
End Sub
