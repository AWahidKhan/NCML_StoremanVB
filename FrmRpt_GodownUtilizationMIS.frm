VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_GodownUtilizationMIS 
   Caption         =   "Godown wise Utilization MIS"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   13080
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   8295
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   15015
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   120
         TabIndex        =   24
         Top             =   7860
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7575
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13361
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   15000
      Begin VB.ComboBox CmbUnitTypeID 
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
         Left            =   11670
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   435
         Width           =   2445
      End
      Begin VB.ComboBox CmbGodownID 
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
         Left            =   9165
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   435
         Width           =   2445
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
         Left            =   5700
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   435
         Width           =   3405
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
         Left            =   2880
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   435
         Width           =   2760
      End
      Begin VB.ComboBox CmbStateID 
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
         TabIndex        =   0
         Top             =   435
         Width           =   2700
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
         Left            =   11700
         TabIndex        =   6
         Top             =   870
         Width           =   1215
      End
      Begin VB.CommandButton CmdGetReport 
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
         Height          =   375
         Left            =   10485
         TabIndex        =   5
         Top             =   870
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
         Left            =   12915
         TabIndex        =   7
         Top             =   870
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CommandButton CmdProcess 
            Caption         =   "Process"
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
            Left            =   2040
            TabIndex        =   18
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton optAll 
            Caption         =   "All"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.OptionButton optSpot 
            Caption         =   "Spot"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.OptionButton optNonNcdex 
            Caption         =   "Non NCDEX"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.OptionButton optNCDEX 
            Caption         =   "NCDEX"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Godown Type"
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
         Left            =   12232
         TabIndex        =   23
         Top             =   120
         Width           =   1320
      End
      Begin VB.Label LblGodownID 
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
         Height          =   300
         Left            =   9967
         TabIndex        =   22
         Top             =   120
         Width           =   840
      End
      Begin VB.Label LblCMPName 
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
         Height          =   300
         Left            =   7162
         TabIndex        =   21
         Top             =   120
         Width           =   480
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
         Left            =   3878
         TabIndex        =   20
         Top             =   120
         Width           =   765
      End
      Begin VB.Label LblStateID 
         Caption         =   "NCML Area"
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
         Left            =   870
         TabIndex        =   19
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label lblRec 
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
         Left            =   6240
         TabIndex        =   10
         Top             =   960
         Width           =   2325
      End
      Begin VB.Label lblRecordCount 
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
         Left            =   8760
         TabIndex        =   9
         Top             =   960
         Width           =   210
      End
   End
End
Attribute VB_Name = "FrmRpt_GodownUtilizationMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCMPID, mGodownID, mUnitTypeID As Variant

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState = "" Then
   Call TableMst_State
Else
   Call TableMst_State(" Where " & Gen_WcState)
End If
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub
'----
Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If

If mStateID <> RsMst_State!StateID Then
   CmbLocationID.Clear
   mLocationID = Null
End If

mStateID = RsMst_State!StateID

End Sub
Private Sub CmbLocationID_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Please Select NCML Area !!"
   CmbStateID.SetFocus
   Exit Sub
End If

tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("Where StateID=" & mStateID & "")
Else
   Call TableMst_Location("Where StateID=" & mStateID & " And " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
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
Private Sub CmbCMPName_GotFocus()
If CmbLocationID.Text = "" Then
   MsgBox "Please Select Location !!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text
tmp1 = ""
If Gen_WcCMP <> "" Then
   tmp1 = "And " & Gen_WcCMP & " And LocationID= " & mLocationID
Else
   tmp1 = " And LocationID = " & mLocationID
End If
Call TableMst_CMP(" where CloseDate is null And LocationID =" & mLocationID)

CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub
'----
Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub
'----
Private Sub CmbGodownID_GotFocus()
If CmbCMPName.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbCMPName.SetFocus
   Exit Sub
End If
tmp = CmbGodownID.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & mCMPID)
CmbGodownID.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownID.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownID.Text = tmp
End Sub
'----
Private Sub CmbGodownID_LostFocus()
If CmbGodownID.Text = "" Then
   Exit Sub
End If
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownID.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbGodownID.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
End Sub


Private Sub CmbUnitTypeID_GotFocus()
tmp = CmbUnitTypeID.Text
CmbUnitTypeID.Clear
Call TableMst_UnitType
If RsMst_UnitType.RecordCount = 0 Then
   MsgBox "No Record Found !!"
   CmbUnitTypeID.SetFocus
   Exit Sub
End If

For I = 1 To RsMst_UnitType.RecordCount
    CmbUnitTypeID.AddItem RsMst_UnitType!UnitType
    RsMst_UnitType.MoveNext
Next
CmbUnitTypeID.Text = tmp
End Sub

Private Sub CmbUnitTypeID_LostFocus()
If CmbUnitTypeID.Text = "" Then Exit Sub
RsMst_UnitType.MoveFirst
RsMst_UnitType.Find "UnitType='" & CmbUnitTypeID.Text & "'"
If RsMst_UnitType.EOF Then
   MsgBox "Invalid Selection !!"
   CmbUnitTypeID.SetFocus
   Exit Sub
End If
mUnitTypeID = RsMst_UnitType!UnitTypeID
End Sub

Private Sub CmdProcess_Click()
MsgBox "Please Wait!!!"
tmp = "update Mst_GSL set BagEquivalent = 1 where flag <> 'B' and (BagEquivalent=0 or BagEquivalent is null) "
Call TmpTable

tmp = "update Mst_GSL set BagEquivalent = 2.00 where flag <> 'B' and CommodityID = 43 "
Call TmpTable

tmp = "update Mst_GSL set BagSize=BalanceWeight/BalanceBags,WtQty = BagEquivalent*BalanceWeight where flag <> 'B' and BalanceBags<>0"
Call TmpTable
MsgBox "Done!!!"
CmdProcess.Enabled = False
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
lblRecordCount.Caption = ""
CmdExcel.Enabled = False
CmbStateID.Text = ""
CmbLocationID.Text = ""
CmbCMPName.Text = ""
CmbGodownID.Text = ""
CmbUnitTypeID.Text = ""
CmbStateID.SetFocus
ProgressBar1.Value = 0
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NewDailyStockMIS.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NewDailyStockMIS.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, C + 1).WrapText = True
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "dd/MMM/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           If C = 4 Then
              oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        End If
    Next
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGetReport_Click()
Dim mVal As Variant
Dim wc As String
wc = ""
Dim tmp1 As String
tmp1 = ""
ProgressBar1.Value = 0
CmdExcel.Enabled = True
Call Grid_Head

If CmbStateID.Text <> "" Then
   wc = GenWc(wc, mStateID, "MS.StateID", "N", "N")
Else
   If Gen_WcState <> "" Then
      tmp1 = Replace(Gen_WcState, "StateID", "MS.StateID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbLocationID.Text <> "" Then
   wc = GenWc(wc, mLocationID, "ML.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbCMPName.Text <> "" Then
   wc = GenWc(wc, mCMPID, "MC.CMPID", "N", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp1 = Replace(Gen_WcCMP, "CMPID", "MC.CMPID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbGodownID.Text <> "" Then
   wc = GenWc(wc, mGodownID, "MG.GodownID", "N", "N")
End If

If CmbUnitTypeID.Text <> "" Then
   wc = GenWc(wc, mUnitTypeID, "MG.UnitTypeID", "N", "N")
End If

'--- Location, Min Date And Start Date
'
'tmp = "Select distinct Mst_CMP.LocationID,StateName,LocationName,min(CurrentDate)as MinDate,max(StartDate) as StartDate"
'tmp = tmp & " From Mst_CMP left outer join Mst_Location  on Mst_Location.LocationID=Mst_CMP.LocationID"
'tmp = tmp & " inner  join Mst_State on Mst_State.StateID=Mst_Location.StateID"
'tmp = tmp & " where CloseDate is null group by Mst_CMP.LocationID,StateName,LocationName"
'tmp = tmp & " order by Mst_State.StateName,Mst_Location.LocationName"


'---Below Old Query Commented By Kishor 31 July 2010 11:40 PM
'
'tmp = "Select distinct GodownID,StateName,LocationName,"
'tmp = tmp & " Mst_CMP.CMPName,mst_godown.GodownNo,mst_godown.Address,(mst_godown.StartDate) as StartDate,"
'tmp = tmp & " (GodownCapacity) , mst_godown.amount From mst_godown"
'tmp = tmp & " inner join Mst_CMP   on mst_godown.CMPID = Mst_CMP.CMPID"
'tmp = tmp & " inner join Mst_Location  on Mst_Location.LocationID=Mst_CMP.LocationID"
'tmp = tmp & " inner join Mst_State on Mst_State.StateID=Mst_Location.StateID"
'tmp = tmp & " Where mst_godown.CloseDate Is Null and UnitTypeID = 1"
'tmp = tmp & " order by Mst_State.StateName,Mst_Location.LocationName"

tmp = "Select Distinct MG.GodownID,MS.StateName,ML.LocationName, "
tmp = tmp & " MC.CMPName,MG.GodownNo,MG.Address,(MG.StartDate) as StartDate, "
tmp = tmp & " (GodownCapacity),0,0,0,0,0, MG.Amount,MU.UnitType,MG.AgreementEndDate,MG.Area From Mst_Godown MG "
tmp = tmp & " inner join Mst_CMP   MC on MG.CMPID = MC.CMPID "
tmp = tmp & " inner join Mst_Location ML on ML.LocationID=MC.LocationID "
tmp = tmp & " inner join Mst_State MS on MS.StateID=ML.StateID "
tmp = tmp & " inner join Mst_UnitType MU on MG.UnitTypeID=MU.UnitTypeID "
If wc <> "" Then
   tmp = tmp & wc
   tmp = tmp & " And MG.CloseDate Is Null " 'and UnitTypeID = 1"
Else
   tmp = tmp & " Where MG.CloseDate Is Null " 'and UnitTypeID = 1"
End If
tmp = tmp & " order by MS.StateName,ML.LocationName"


Call TmpTable


tmp = " Select TGA.GodownID ,Max(TA.ExpiryDate) 'ExpDate',Sum(RentAmount) 'RentAmount' from Txn_GodownLessorAgreeDetail TGA "
tmp = tmp & " Inner Join Txn_LessorAgreement TA On TGA.LAID = TA.LAID "
tmp = tmp & " Where TA.Flag not in ('R') And  TA.ExpiryDate >= '" & Format(Date, Gen_DatFormat) & "' "
tmp = tmp & " Group BY TGA.GodownID "
      
Call Tmp3Table

ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 2

For I = 2 To TmpRs.RecordCount + 1
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To TmpRs.Fields.Count - 3
        MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid1.TextMatrix(I, 6) = Format(MSHFlexGrid1.TextMatrix(I, 6), "dd-MMM-yyyy")
    MSHFlexGrid1.TextMatrix(I, 13) = 0
    MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
    MSHFlexGrid1.TextMatrix(I, 15) = IIf(IsNull(TmpRs!AgreementEndDate), "", Format(TmpRs!AgreementEndDate, "dd-mmm-yyyy"))
    MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs!Area), "", TmpRs!Area)
    
    If LCase(MSHFlexGrid1.TextMatrix(I, 14)) = "leased" Then
       If TmpRs3.RecordCount > 0 Then
          TmpRs3.MoveFirst
          TmpRs3.Find "GodownID = " & Val(TmpRs!GodownID) & ""
          If TmpRs3.EOF = False Then
             MSHFlexGrid1.TextMatrix(I, 15) = Format(TmpRs3!ExpDate, "dd-mmm-yyyy")
             MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs3!RentAmount), "", TmpRs3!RentAmount)
          Else
             MSHFlexGrid1.TextMatrix(I, 15) = "" 'Format(TmpRs3!ExpDate, "dd-mmm-yyyy")
             MSHFlexGrid1.TextMatrix(I, 13) = 0
          End If
       End If
    End If
    ProgressBar1.Value = ProgressBar1.Value + 1
    TmpRs.MoveNext
Next

ProgressBar1.Value = ProgressBar1.Max
lblRecordCount.Caption = TmpRs.RecordCount

'--- For Capacity

'--- For Reservation Capacity

'*************/
'''tmp = "Select GodownID,Sum(ReservationQty) as 'ReservationQty' "
'''tmp = tmp & " From  Txn_SRGodownMapDetails a"
'''tmp = tmp & " inner join Txn_SRGodownMapping b on a.SRGodownMappingID = b.SRGodownMappingID"
'''tmp = tmp & " inner join Mst_StorageContract c on b.contractid = c.contractid and BookedTo>= getdate()"
'''tmp = tmp & " group by GodownID"
'************/


'--For Reservation Capacity

tmp = "Select a.GodownID,Sum(a.ReservationQty) as 'ReservationQty'"
tmp = tmp & " From  Txn_SRGodownMapDetails a"
tmp = tmp & " inner join Txn_SRGodownMapping b on a.SRGodownMappingID = b.SRGodownMappingID"
tmp = tmp & " inner join Mst_StorageContract c on b.contractid = c.contractid and ExtendedPeriod>= getdate()"
tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=a.godownID"
tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID=MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID=MC.LocationID"
tmp = tmp & " inner join Mst_State MS on MS.StateID=ML.StateID "
tmp = tmp & wc
tmp = tmp & " Group by a.GodownID Order By a.GodownID"


Call TmpTable
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
If TmpRs.RecordCount > 0 Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.MoveFirst
       TmpRs.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(I, 0)) & ""
       If Not TmpRs.EOF Then
          Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 9, Format(Val(IIf(IsNull(TmpRs!ReservationQty), "0", TmpRs!ReservationQty)), "0"))
       Else
          Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 9, "0")
       End If
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
End If
ProgressBar1.Value = ProgressBar1.Max
'--- Total Wt. Stock

'************/
'''tmp = "Select GodownID,Sum(WtQty) as 'Stock' From  mst_gsl  WHERE (Flag <> 'E') AND (Flag <> 'Z') group by GodownID "
'************/

tmp = "Select MGSL.GodownID,Sum(WtQty) as 'Stock'"
tmp = tmp & " From  Mst_GSL MGSL"
tmp = tmp & " Inner Join Mst_Godown  MG On MG.GodownID=MGSL.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID=MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID=MC.LocationID"
tmp = tmp & " inner join Mst_State MS on MS.StateID=ML.StateID "
'Parag
'If wc <> "" Then
'   tmp = tmp & wc
'   tmp = tmp & " And (MGSL.Flag <> 'E') AND (MGSL.Flag <> 'Z')"
'Else
'   tmp = tmp & " Where (MGSL.Flag <> 'E') AND (MGSL.Flag <> 'Z')"
'End If

If wc <> "" Then
   tmp = tmp & wc
   tmp = tmp & " And (MGSL.Status <> 'E') AND (MGSL.Status <> 'Z')"
Else
   tmp = tmp & " Where (MGSL.Status <> 'E') AND (MGSL.Status <> 'Z')"
End If

tmp = tmp & " Group by MGSL.GodownID"
tmp = tmp & " Order By MGSL.GodownID"

Call TmpTable
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   For I = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.MoveFirst
       TmpRs.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(I, 0)) & ""
       If Not TmpRs.EOF Then
          Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 8, Format(Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
       Else
          Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 8, "0")
       End If
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
Else
   ProgressBar1.Value = 0
   For I = 2 To MSHFlexGrid1.Rows - 2
       Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 8, "0")
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
End If
ProgressBar1.Value = ProgressBar1.Max
' Reserved Stock
'********/
'''tmp = "Select GodownID,Sum(WtQty) as 'Stock' From  mst_gsl where ((Flag <> 'E') AND (Flag <> 'Z')) And (ContractID <> 0 and ContractID is not null) group by GodownID"
'********/

'tmp = " Select MGSL.GodownID,Sum(MGSL.WtQty) as 'Stock'"
'tmp = tmp & " From  Mst_GSL  MGSL"
'tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=MGSL.GodownID"
'tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID=MG.CMPID"
'tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID=MC.LocationID"
'tmp = tmp & " Inner Join Mst_State MS On MS.StateID=ML.StateID "
'If wc <> "" Then
'   tmp = tmp & wc
'   tmp = tmp & " And ((MGSL.Flag <> 'E') AND (MGSL.Flag <> 'Z')) And (MGSL.ContractID <> 0 and MGSL.ContractID is not null)"
'Else
'   tmp = tmp & " Where ((MGSL.Flag <> 'E') AND (MGSL.Flag <> 'Z')) And (MGSL.ContractID <> 0 and MGSL.ContractID is not null)"
'End If
'tmp = tmp & " Group by MGSL.GodownID"
'tmp = tmp & " Order By MGSL.GodownID"

tmp = "Select a.GodownID,Sum(a.WtQty) as 'Stock'"
tmp = tmp & " From  Txn_SRGodownMapDetails a"
tmp = tmp & " inner join Txn_SRGodownMapping b on a.SRGodownMappingID = b.SRGodownMappingID"
tmp = tmp & " inner join Mst_StorageContract c on b.contractid = c.contractid and ExtendedPeriod>= getdate()"
tmp = tmp & " Inner Join Mst_Godown MG On MG.GodownID=a.godownID"
tmp = tmp & " Inner Join Mst_CMP MC On MC.CMPID=MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID=MC.LocationID"
tmp = tmp & " inner join Mst_State MS on MS.StateID=ML.StateID "
tmp = tmp & wc
tmp = tmp & " Group by a.GodownID Order By a.GodownID"


Call TmpTable

If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   For I = 2 To MSHFlexGrid1.Rows - 2
       TmpRs.MoveFirst
       TmpRs.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(I, 0)) & ""
       If Not TmpRs.EOF Then
          Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 10, Format(Val(IIf(IsNull(TmpRs!Stock), "0", TmpRs!Stock)), "0"))
       Else
          Call dispdata(MSHFlexGrid1.TextMatrix(I, 0), 10, "0")
       End If
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
End If

ProgressBar1.Value = 0
For I = 2 To MSHFlexGrid1.Rows - 2

    MSHFlexGrid1.TextMatrix(I, 11) = Val(MSHFlexGrid1.TextMatrix(I, 8)) - Val(MSHFlexGrid1.TextMatrix(I, 10))
    
    If Val(MSHFlexGrid1.TextMatrix(I, 10)) > Val(MSHFlexGrid1.TextMatrix(I, 9)) Then
       MSHFlexGrid1.TextMatrix(I, 11) = Val(MSHFlexGrid1.TextMatrix(I, 11)) + Val(MSHFlexGrid1.TextMatrix(I, 10))
    Else
       MSHFlexGrid1.TextMatrix(I, 11) = Val(MSHFlexGrid1.TextMatrix(I, 11)) + Val(MSHFlexGrid1.TextMatrix(I, 9))
    End If
    
    If Val(MSHFlexGrid1.TextMatrix(I, 11)) <> 0 And Val(MSHFlexGrid1.TextMatrix(I, 7)) <> 0 Then
       MSHFlexGrid1.TextMatrix(I, 12) = (Val(MSHFlexGrid1.TextMatrix(I, 11)) / Val(MSHFlexGrid1.TextMatrix(I, 7))) * 100
    End If
    
    MSHFlexGrid1.TextMatrix(I, 12) = Format(Val(MSHFlexGrid1.TextMatrix(I, 12)), "####")
    
    MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(I, 7))
    MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(I, 8))
    MSHFlexGrid1.TextMatrix(0, 9) = Val(MSHFlexGrid1.TextMatrix(0, 9)) + Val(MSHFlexGrid1.TextMatrix(I, 9))
    MSHFlexGrid1.TextMatrix(0, 10) = Val(MSHFlexGrid1.TextMatrix(0, 10)) + Val(MSHFlexGrid1.TextMatrix(I, 10))
    MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(I, 11))
    MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(I, 12))
    
    MSHFlexGrid1.TextMatrix(0, 7) = Val(Format(MSHFlexGrid1.TextMatrix(0, 7), "########"))
    MSHFlexGrid1.TextMatrix(0, 8) = Val(Format(MSHFlexGrid1.TextMatrix(0, 8), "########"))
    MSHFlexGrid1.TextMatrix(0, 9) = Val(Format(MSHFlexGrid1.TextMatrix(0, 9), "########"))
    MSHFlexGrid1.TextMatrix(0, 10) = Val(Format(MSHFlexGrid1.TextMatrix(0, 10), "########"))
    MSHFlexGrid1.TextMatrix(0, 11) = Val(Format(MSHFlexGrid1.TextMatrix(0, 11), "########"))
    MSHFlexGrid1.TextMatrix(0, 12) = Val(Format(MSHFlexGrid1.TextMatrix(0, 12), "########"))
    ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max

If Val(MSHFlexGrid1.TextMatrix(0, 7)) <> 0 And Val(MSHFlexGrid1.TextMatrix(0, 11)) <> 0 Then
   MSHFlexGrid1.TextMatrix(0, 12) = (Val(MSHFlexGrid1.TextMatrix(0, 11)) / Val(MSHFlexGrid1.TextMatrix(0, 7))) * 100
End If
MSHFlexGrid1.TextMatrix(0, 12) = Format(Val(MSHFlexGrid1.TextMatrix(0, 12)), "####")


'tmp = "select Mst_CMP.LocationID,g.UnitTypeID,sum(a.ReservationQty) as 'ResQty'"
'tmp = tmp & " From  Txn_SRGodownMapDetails a"
'tmp = tmp & " inner join Txn_SRGodownMapping  b on a.SRGodownMappingID = b.SRGodownMappingID"
'tmp = tmp & " inner join Mst_StorageContract c on b.ContractID = c.ContractID  and BookedTo>= getdate() "
'tmp = tmp & " inner join Mst_Godown g on a.GodownID = g.GodownID"
'tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = g.CMPID"
'tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
'tmp = tmp & " Group by Mst_CMP.LocationID,g.UnitTypeID"
'tmp = tmp & " Having g.UnitTypeID <> 1"
'Call TmpTable
'If TmpRs.RecordCount > 0 Then
'   For i = 2 To MSHFlexGrid1.Rows - 2
'       TmpRs.Filter = ""
'       TmpRs.Filter = "LocationID=" & MSHFlexGrid1.TextMatrix(i, 0) & ""
'       If TmpRs.RecordCount > 0 Then
'          Call DispData(MSHFlexGrid1.TextMatrix(i, 0), 10, Format(Val(IIf(IsNull(TmpRs!ResQty), "0", TmpRs!ResQty)), "0"))
'          mVal = (Val(IIf(MSHFlexGrid1.TextMatrix(i, 6) = "", "0", MSHFlexGrid1.TextMatrix(i, 6)) - Format(Val(IIf(IsNull(TmpRs!ResQty), "0", TmpRs!ResQty)), "0")))
'          Call DispData(MSHFlexGrid1.TextMatrix(i, 0), 6, mVal)
'       End If
'   Next
'End If
'
'
'Dim ct As Variant
'ct = 0
'For x = 2 To MSHFlexGrid1.Rows - 1
'    ' If (MSHFlexGrid1.TextMatrix(x, 5)) <> "" And (MSHFlexGrid1.TextMatrix(x, 4)) <> "" And (MSHFlexGrid1.TextMatrix(x, 8)) = "" Then
'
'    MSHFlexGrid1.TextMatrix(x, 8) = ""
'    MSHFlexGrid1.TextMatrix(x, 8) = IIf(Val(MSHFlexGrid1.TextMatrix(x, 6)) > Val(MSHFlexGrid1.TextMatrix(x, 7)), MSHFlexGrid1.TextMatrix(x, 6), MSHFlexGrid1.TextMatrix(x, 7))
'
'    ct = ct + 1
'    MSHFlexGrid1.TextMatrix(x, 0) = ct
'    If Val(MSHFlexGrid1.TextMatrix(x, 4)) <> 0 Then
'       MSHFlexGrid1.TextMatrix(x, 9) = ((Val(MSHFlexGrid1.TextMatrix(x, 5)) + Val(MSHFlexGrid1.TextMatrix(x, 8))) * 100) / Val(MSHFlexGrid1.TextMatrix(x, 4))
'       MSHFlexGrid1.TextMatrix(x, 9) = Format(MSHFlexGrid1.TextMatrix(x, 9), "######0")
'    End If
'
'    MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(x, 4))
'    MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(x, 5))
'    MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(x, 6))
'    MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(x, 7))
'    MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(x, 8))
'
'
'    MSHFlexGrid1.TextMatrix(0, 10) = Val(MSHFlexGrid1.TextMatrix(0, 10)) + Val(MSHFlexGrid1.TextMatrix(x, 10))
'    MSHFlexGrid1.TextMatrix(0, 11) = Val(MSHFlexGrid1.TextMatrix(0, 11)) + Val(MSHFlexGrid1.TextMatrix(x, 11))
'    MSHFlexGrid1.TextMatrix(0, 12) = Val(MSHFlexGrid1.TextMatrix(0, 12)) + Val(MSHFlexGrid1.TextMatrix(x, 12))
'    MSHFlexGrid1.TextMatrix(0, 13) = Val(MSHFlexGrid1.TextMatrix(0, 13)) + Val(MSHFlexGrid1.TextMatrix(x, 13))
'    MSHFlexGrid1.TextMatrix(0, 14) = Val(MSHFlexGrid1.TextMatrix(0, 14)) + Val(MSHFlexGrid1.TextMatrix(x, 14))
'
'    MSHFlexGrid1.TextMatrix(0, 17) = Val(MSHFlexGrid1.TextMatrix(0, 17)) + Val(MSHFlexGrid1.TextMatrix(x, 17))
'    MSHFlexGrid1.TextMatrix(0, 18) = Val(MSHFlexGrid1.TextMatrix(0, 18)) + Val(MSHFlexGrid1.TextMatrix(x, 18))
'    MSHFlexGrid1.TextMatrix(0, 19) = Val(MSHFlexGrid1.TextMatrix(0, 19)) + Val(MSHFlexGrid1.TextMatrix(x, 19))
'    MSHFlexGrid1.TextMatrix(0, 20) = Val(MSHFlexGrid1.TextMatrix(0, 20)) + Val(MSHFlexGrid1.TextMatrix(x, 20))
'
'    MSHFlexGrid1.TextMatrix(0, 9) = ((Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(0, 8))) * 100) / Val(MSHFlexGrid1.TextMatrix(0, 4))
'    MSHFlexGrid1.TextMatrix(0, 9) = Format(MSHFlexGrid1.TextMatrix(0, 9), "######0")
'Next
'tmp = ""

MsgBox " Done!!!"

End Sub

Private Sub dispdata(L_ID As Variant, mCol As Variant, Rec As Variant)
For k = 2 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(k, 0) = L_ID Then
       MSHFlexGrid1.TextMatrix(k, mCol) = Rec
       Exit For
    End If
Next
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
CmdExcel.Enabled = False
Call Grid_Head

End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .Cols = 17 '16 '14
    .FixedRows = 2
    .WordWrap = True
    .RowHeight(1) = 800
    .TextMatrix(1, 0) = "Sr No."  'Location ID
    .TextMatrix(1, 1) = "NCML Area"
    .TextMatrix(1, 2) = "Location"
    .TextMatrix(1, 3) = "CMP"
    .TextMatrix(1, 4) = "Godown No"
    .TextMatrix(1, 5) = "Godown Address"
    .TextMatrix(1, 6) = "Godown Start Date "
    .TextMatrix(1, 7) = "Godown Capacity (in MT)"
    .TextMatrix(1, 8) = "Weighted Stock (in MT)"
    .TextMatrix(1, 9) = "Reserved Capacity (in MT)"
    .TextMatrix(1, 10) = "Reserved Weighted Stock (in MT)"
    .TextMatrix(1, 11) = "Total Utilised Capacity"
    .TextMatrix(1, 12) = "Capacity Utilisation in % "
    .TextMatrix(1, 13) = "Lease Amount"
    .TextMatrix(1, 14) = "Godown Type"
    .TextMatrix(1, 15) = "Agreement End Date"
    .TextMatrix(1, 16) = "Area of Godown"
    
    .ColWidth(0) = 0
    .ColWidth(1) = 2200
    .ColWidth(2) = 2200
    .ColWidth(3) = 1600
    .ColWidth(4) = 1600
    .ColWidth(5) = 1600
    .ColWidth(6) = 1600
    .ColWidth(7) = 1600
    .ColWidth(8) = 1600
    .ColWidth(9) = 1600
    .ColWidth(10) = 1400
    .ColWidth(11) = 1400
    .ColWidth(12) = 1400
    .RowHeight(1) = 800
    
    
End With
End Sub



