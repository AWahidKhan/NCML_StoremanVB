VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_CMPBOMDetails 
   Caption         =   "CMP BOM Details"
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
      Height          =   8415
      Left            =   105
      TabIndex        =   9
      Top             =   1155
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8055
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
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
   Begin VB.Frame Frame1 
      Height          =   990
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
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
         Left            =   4560
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   3135
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
         Left            =   9975
         TabIndex        =   5
         Top             =   300
         Width           =   990
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
         Left            =   8985
         TabIndex        =   4
         Top             =   300
         Width           =   990
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
         Height          =   420
         Left            =   7995
         TabIndex        =   3
         Top             =   300
         Width           =   990
      End
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblCMP 
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
         Height          =   360
         Left            =   3960
         TabIndex        =   11
         Top             =   360
         Width           =   435
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
         Left            =   11205
         TabIndex        =   8
         Top             =   390
         Width           =   2295
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
         Left            =   13485
         TabIndex        =   7
         Top             =   390
         Width           =   90
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
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   353
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmRpt_CMPBOMDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCMPID, mLocationID As Variant

Private Sub CmdClear_Click()
Call Grid_Head
CmbCMPName.Text = ""
CmbLocation.Text = ""
LblTotalRecord.Caption = "-"
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_CMPBOMDetails.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_CMPBOMDetails.xls")
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


Private Sub CmdGo_Click()

Call Grid_Head


MsgBox "Please wait till next message !!"
tmp = "Select ML.LocationName,MC.CMPName,MC.CMPAlias,MC.Address,MC.TelephoneNo,MC.CurrentDate,MC.CMPLockedDate,MC.CMPID "
tmp = tmp & " From Mst_CMP MC"
tmp = tmp & " Inner Join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where Month(MC.CMPLOckedDate) >= Month(Getdate()) And MC.CloseDate Is Null"
If CmbLocation.Text <> "" Then
   tmp = tmp & " And MC.LocationID=" & mLocationID & ""
Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " And MC." & Gen_WcLocation
   End If
End If

If CmbCMPName.Text <> "" Then
   tmp = tmp & " And MC.CMPID=" & mCMPID & ""
End If

tmp = tmp & " Order By ML.LocationName,MC.CMPName"
Call TmpTable


If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For mr = 1 To TmpRs.RecordCount
       MSHFlexGrid1.TextMatrix(mr, 0) = TmpRs!LocationName
       MSHFlexGrid1.TextMatrix(mr, 1) = TmpRs!CMPName
       MSHFlexGrid1.TextMatrix(mr, 2) = TmpRs!CMPAlias
       MSHFlexGrid1.TextMatrix(mr, 3) = TmpRs!Address
       MSHFlexGrid1.TextMatrix(mr, 4) = TmpRs!TelephoneNo
       MSHFlexGrid1.TextMatrix(mr, 5) = Format(TmpRs!CurrentDate, "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(mr, 6) = Format(TmpRs!CMPLockedDate, "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(mr, 7) = TmpRs!CMPID
       TmpRs.MoveNext
   Next
   MsgBox "Done !!"
Else
MsgBox " Record Not Found !!!"
End If
LblTotalRecord.Refresh
LblTotalRecord.Caption = TmpRs.RecordCount
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 8
     .WordWrap = True
     .AllowUserResizing = flexResizeBoth
     .RowHeight(0) = 800
     .TextMatrix(0, 0) = "Location Name"
     .TextMatrix(0, 1) = "CMP Name"
     .TextMatrix(0, 2) = "CMP Alise"
     .TextMatrix(0, 3) = "Address"
     .TextMatrix(0, 4) = "Telephone"
     .TextMatrix(0, 5) = "BOD (Beginning of Day)" '"Current Date"
     .TextMatrix(0, 6) = "BOM (Beginning of Month)" '"CMP Locked Date"
     .TextMatrix(0, 7) = "CMP ID"
     .ColWidth(0) = 2200
     .ColWidth(1) = 2200
     .ColWidth(2) = 2200
     .ColWidth(3) = 3200
     .ColWidth(4) = 1600
     .ColWidth(5) = 1500
     .ColWidth(6) = 1500
     .ColWidth(7) = 0
     .Font.Size = 10
End With
End Sub

Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location
End If

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

Private Sub CmbCMPName_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Please Select Location !!!"
   CmbLocation.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text

Call TableMst_CMP(" where LocationID=" & mLocationID & " ")
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "First Select Location!!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub

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
'CmdExcel.Enabled = False

End Sub
