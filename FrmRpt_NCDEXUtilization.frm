VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_NCDEXUtilization 
   Caption         =   "NCDEX Utilization"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12105
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   12105
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15045
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   555
         Width           =   2760
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
         Left            =   2940
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   555
         Width           =   3405
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
         Left            =   6405
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   555
         Width           =   2445
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
         Left            =   13710
         TabIndex        =   5
         Top             =   555
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
         Left            =   11280
         TabIndex        =   4
         Top             =   555
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
         Height          =   360
         Left            =   12495
         TabIndex        =   3
         Top             =   555
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All"
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
         Left            =   10560
         TabIndex        =   2
         Top             =   548
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Above 90 %"
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
         Left            =   9000
         TabIndex        =   1
         Top             =   548
         Value           =   -1  'True
         Width           =   1455
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
         Left            =   1125
         TabIndex        =   14
         Top             =   240
         Width           =   765
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
         Left            =   3720
         TabIndex        =   13
         Top             =   240
         Width           =   720
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
         Left            =   7200
         TabIndex        =   12
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblTotalRecords 
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
         Left            =   13605
         TabIndex        =   7
         Top             =   240
         Width           =   675
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
         Left            =   11280
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8655
      Left            =   120
      TabIndex        =   8
      Top             =   1125
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15266
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   1
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
Attribute VB_Name = "FrmRpt_NCDEXUtilization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mExchangeTypeID As Variant
Dim mLocationID, mCMPID, mGodownID As Variant

Private Sub CmbLocationID_GotFocus()

tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location("Where " & Gen_WcLocation)
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
   tmp1 = "And LocationID = " & mLocationID
End If
Call TableMst_CMP(" where CloseDate is null " & tmp1)

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

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdGetData_Click()
Dim Wc As String
Wc = ""
Dim tmp1 As String
tmp1 = ""

LblTotalRecords.Caption = ""
Call GridHead
If CmbLocationID.Text <> "" Then
   Wc = GenWc(Wc, mLocationID, "ML.LocationID", "N", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      Wc = GenWc(Wc, "", tmp1, "N", "I")
   End If
End If
If CmbCMPName.Text <> "" Then
   Wc = GenWc(Wc, mCMPID, "MC.CMPID", "N", "N")
Else
   If Gen_WcCMP <> "" Then
      tmp1 = Replace(Gen_WcCMP, "CMPID", "MC.CMPID")
      Wc = GenWc(Wc, "", tmp1, "N", "I")
   End If
End If

tmp = "Select ML.LocationName,MC.CMPName,MG.GodownNo,MG.Address,Isnull(MG.NCDEXAccreditedQty,0) as NCDEXAccreditedQty ,sum(GSL.BalanceWeight)As BalanceWeight , Sum(GSL.BalanceWeight/NULLIF(MG.NCDEXAccreditedQty,0)*100) As Utilization"
tmp = tmp & " From Mst_GSL GSL"
tmp = tmp & " Inner join Mst_Godown MG on GSL.GodownID=MG.GodownID"
tmp = tmp & " inner join Mst_CMP MC on GSL.CMPID=MC.CMPID"
tmp = tmp & " inner join Mst_Location ML on MC.LocationID=ML.LocationID "
If Wc <> "" Then
   tmp = tmp & Wc
   tmp = tmp & " And (GSL.Flag='O' or GSL.Flag='J') and GSL.ExchangeTypeID=2"
Else
   tmp = tmp & " Where (GSL.Flag='O' or GSL.Flag='J') and GSL.ExchangeTypeID=2"
End If
tmp = tmp & " Group By ML.LocationName,MC.CMPName,MG.GodownNo,MG.Address,MG.NCDEXAccreditedQty"
If Option1.Value = True Then
   tmp = tmp & " Having Sum(GSL.BalanceWeight/NULLIF(MG.NCDEXAccreditedQty,0)*100) > 90 "
End If
tmp = tmp & " order by ML.LocationName,MC.CMPName,MG.GodownNo"

Call TmpTable

mRow = 1
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       For c = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), 0, TmpRs.Fields(c))
       Next
       MSHFlexGrid1.TextMatrix(I, 6) = Round(MSHFlexGrid1.TextMatrix(I, 6), 2)
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       TmpRs.MoveNext
   Next
   LblTotalRecords.Caption = MSHFlexGrid1.Rows - 2
   CmdExcel.Enabled = True
Else
   MsgBox "No record Found."
End If

End Sub


Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Exit Sub
End If

Call GridHead
Frame1.Enabled = True
CmdExcel.Enabled = False
End Sub
Private Function GridHead()
With MSHFlexGrid1
     .Clear
     .Cols = 7
     .Rows = 2
     .TextMatrix(0, 0) = "Location"
     .TextMatrix(0, 1) = "CMP Name"
     .TextMatrix(0, 2) = "Godown No"
     .TextMatrix(0, 3) = "Address"
     .TextMatrix(0, 4) = "NCDEX Accre. Qty"
     .TextMatrix(0, 5) = "Balance Stock"
     .TextMatrix(0, 6) = "Utilization(%)"
     
     .ColWidth(0) = 2600
     .ColWidth(1) = 2400
     .ColWidth(2) = 1600
     .ColWidth(3) = 2800
     .ColWidth(4) = 1800
     .ColWidth(5) = 1600
     .ColWidth(6) = 1600
End With
End Function
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_NCDEXUtilization.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_NCDEXUtilization.xls")
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
