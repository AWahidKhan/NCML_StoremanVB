VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_UpdateCMPBlockDays 
   Caption         =   "Update CMP Block Days"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   8400
      Width           =   12000
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
         Height          =   615
         Left            =   9855
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7920
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8325
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.CommandButton CmdGetDate 
         Caption         =   "Get Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3465
         TabIndex        =   6
         Top             =   1740
         Width           =   1935
      End
      Begin VB.ComboBox CmbRegion 
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
         Left            =   1665
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   165
         Width           =   3990
      End
      Begin VB.CheckBox ChkCheckAll 
         Caption         =   "Select All"
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
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ComboBox CmbState 
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
         Left            =   6720
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   165
         Width           =   5085
      End
      Begin VB.TextBox TxtPreFix 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   570
         Width           =   3975
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
         Left            =   1665
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   570
         Width           =   4005
      End
      Begin VB.TextBox TxtRemark 
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
         Height          =   720
         Left            =   1665
         MaxLength       =   3500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   960
         Width           =   10155
      End
      Begin VB.TextBox TxtNewBlockDays 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1725
         Width           =   1410
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   5700
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   11820
         _ExtentX        =   20849
         _ExtentY        =   10054
         _Version        =   393216
         FixedCols       =   0
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Region"
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
         Left            =   195
         TabIndex        =   18
         Top             =   225
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "State"
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
         Left            =   5760
         TabIndex        =   16
         Top             =   225
         Width           =   465
      End
      Begin VB.Label Label31 
         Caption         =   "SM_Prefix"
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
         Left            =   5760
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label LblLocationID 
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
         Left            =   195
         TabIndex        =   13
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label14 
         Caption         =   "Remark"
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
         Left            =   195
         TabIndex        =   12
         Top             =   1193
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "New Block No Of Days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   195
         TabIndex        =   11
         Top             =   1665
         Width           =   1395
      End
   End
End
Attribute VB_Name = "FrmTxn_UpdateCMPBlockDays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mStateID, mRegionID As Variant
Const strChecked = "þ"
Const strUnChecked = "q"


Private Sub ChkCheckAll_Click()

If MSHFlexGrid2.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For i = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.row = i
       MSHFlexGrid2.Col = 0
       MSHFlexGrid2.TextMatrix(i, 0) = strChecked
   Next
Else
   For i = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.row = i
       MSHFlexGrid2.TextMatrix(i, 0) = strUnChecked
   Next
End If

End Sub

Private Sub CmbState_GotFocus()
If CmbRegion.Text = "" Then
   MsgBox "Select Region!!!"
   CmbRegion.SetFocus
   Exit Sub
End If

tmp = CmbState.Text
Call TableMst_State(" Where RegionID = " & mRegionID & " And Flag = 'A'")

CmbState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No state found"
   CmbState.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_State.RecordCount
    CmbState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbState.Text = tmp
End Sub

Private Sub CmbState_LostFocus()
If CmbState.Text = "" Then
   CmbLocationID.Text = ""
   TxtPreFix = ""
   mLocationID = ""
   Call Grid_Head
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbState.Text & "'"

If RsMst_State.EOF Then
   MsgBox "Invalid State selection "
   CmbState.SetFocus
   Exit Sub
End If

If CmbState.Text <> tmp Then
   CmbLocationID.Text = ""
   TxtPreFix = ""
   mLocationID = ""
   Call Grid_Head
End If

mStateID = RsMst_State.Fields(0)
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub ClearData()
CmbState.Text = ""
CmbLocationID.Text = ""
TxtRemark.Text = ""
TxtNewBlockDays = ""
Call Grid_Head
ChkCheckAll.Value = 0

End Sub

Private Sub CmdGetDate_Click()

If CmbRegion.Text = "" Then
   MsgBox "Select Region!!!"
   CmbRegion.SetFocus
   Exit Sub
End If

Call FillGrid

End Sub

Private Sub CmdUpdate_Click()

If CmbRegion.Text = "" Then
   MsgBox "Select Region!!!"
   CmbRegion.SetFocus
   Exit Sub
End If

If TxtNewBlockDays = "" Then
   MsgBox "Blank Block Days not allowed!!!"
   TxtNewBlockDays.SetFocus
   Exit Sub
End If

If TxtRemark = "" Then
   MsgBox "Blank Remark not allowed!!!"
   TxtRemark.SetFocus
   Exit Sub
End If

For i = 1 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 0) = strChecked Then
       tmp = "Update Mst_CMP Set BlockDays = " & Val(TxtNewBlockDays) & " Where CMPID = " & MSHFlexGrid2.TextMatrix(i, 4) & ""
       Call TmpTable
       Call SaveLogHistory(Val(MSHFlexGrid2.TextMatrix(i, 4)), Val(MSHFlexGrid2.TextMatrix(i, 5)))
    End If
Next

MsgBox " Done!!!"

Call FillGrid

'Call ClearData

End Sub
Private Sub SaveLogHistory(mCMPID_, mOldDays_)
Call TableLog_UpdateCMPBlockDays
RsLog_UpdateCMPBlockDays.AddNew
RsLog_UpdateCMPBlockDays!CMPUpdateID = GetMaxID
RsLog_UpdateCMPBlockDays!CMPID = mCMPID_
RsLog_UpdateCMPBlockDays!OldBlockDays = Val(mOldDays_)
RsLog_UpdateCMPBlockDays!NewBlockDays = Val(TxtNewBlockDays)
RsLog_UpdateCMPBlockDays!Remark = Trim(TxtRemark.Text)
RsLog_UpdateCMPBlockDays!UpdatedBy = Gen_UserLoginID
RsLog_UpdateCMPBlockDays!UpdatedOn = Now
RsLog_UpdateCMPBlockDays.Update
End Sub
Private Function GetMaxID() As Double
Dim Retval As Double
tmp = "Select max(CMPUpdateID) from Log_UpdateCMPBlockDays"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxID = Retval
End Function

Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
Call ClearData

End Sub

Private Sub CmbLocationID_GotFocus()
If CmbState.Text = "" Then
   MsgBox "Select State!!!"
   CmbState.SetFocus
   Exit Sub
End If

tmp = CmbLocationID.Text
Call TableMst_Location(" Where StateID = " & mStateID & " And LocationID <> 0 ANd LocationID in (Select Distinct LOcationID from Mst_CMP Where CloseDate is null)")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   mLocationID = Null
   TxtPreFix = ""
   Call Grid_Head
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid  selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
If mLocationID <> RsMst_Location!LocationID Then
   Call Grid_Head
   TxtPreFix = ""
End If

mLocationID = RsMst_Location!LocationID
TxtPreFix = RsMst_Location!SM_Prefix

End Sub

Private Sub CmbCMPID_GotFocus()
tmp = CmbCMPID.Text
Call TableMst_CMP(" Where CloseDate is null and LocationID = " & mLocationID & " ")
CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
    CmbCMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp

End Sub
'----
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   mCMPID = Null
   TxtCMPAddress = ""
   TxtEHoPCMPID = ""
   TxStartDate = ""
   TxtCurrentDate = ""
   TxtOldBlockDays = ""
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid  selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
   TxtStartDate = ""
   TxtCurrentDate = ""
   TxtOldBlockDays = ""
   TxtEHoPCMPID = ""
End If

mCMPID = RsMst_CMP!CMPID
TxtCMPAddress = IIf(IsNull(RsMst_CMP!Address), "", RsMst_CMP!Address)
TxtStartDate = RsMst_CMP!StartDate 'Format(RsMst_CMP!StartDate, "dd/mm/yyyy")
TxtCurrentDate = RsMst_CMP!CurrentDate + 1 ' CDate(Format(RsMst_CMP!CurrentDate, "dd/mm/yyyy")) + 1
TxtEHoPCMPID = RsMst_CMP!EHOPCMPID
TxtOldBlockDays = IIf(IsNull(RsMst_CMP!BlockDays), 0, RsMst_CMP!BlockDays)


End Sub
Private Sub MSHFlexGrid2_Click()

With MSHFlexGrid2
     If .Col = 0 Then
        If .TextMatrix(.row, .Col) = "" Then Exit Sub
           .Col = 0
           If .TextMatrix(.row, .Col) = strUnChecked Then
              .TextMatrix(.row, .Col) = strChecked
           Else
              .TextMatrix(.row, .Col) = strUnChecked
           End If
     End If
End With

End Sub

Private Sub TxtNewBlockDays_LostFocus()

If TxtNewBlockDays = "" Then Exit Sub

If IsNumeric(TxtNewBlockDays) = False Or Val(TxtNewBlockDays) <= 0 Then
   MsgBox "Invalid Format!!!"
   TxtNewBlockDays.SetFocus
   Exit Sub
End If
TxtNewBlockDays = Format(TxtNewBlockDays, "########0")
End Sub

Private Sub Grid_Head()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 6

MSHFlexGrid2.TextMatrix(0, 0) = "Selection "
MSHFlexGrid2.TextMatrix(0, 1) = "Location Name"
MSHFlexGrid2.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid2.TextMatrix(0, 3) = "Date Upto which Data Entry Completed"
MSHFlexGrid2.TextMatrix(0, 4) = "CMPID"
MSHFlexGrid2.TextMatrix(0, 5) = "Old Block Days"


MSHFlexGrid2.ColWidth(0) = 800
MSHFlexGrid2.ColWidth(1) = 1400
MSHFlexGrid2.ColWidth(2) = 5500
MSHFlexGrid2.ColWidth(3) = 2000
MSHFlexGrid2.ColWidth(4) = 0
MSHFlexGrid2.ColWidth(5) = 1000


End Sub

Private Sub FillGrid()
Dim k As Variant

If CmbLocationID.Text = "" Then
   If CmbState.Text = "" Then
      mLocationID_ = "(0"
      
      tmp = "Select Distinct LocationID from Mst_Location WHere StateID in (Select Distinct STateID from Mst_State WHere RegionID = " & mRegionID & ") "
      Call TmpTable
      
      For i = 1 To TmpRs.RecordCount
          mLocationID_ = mLocationID_ & "," & TmpRs!LocationID
          TmpRs.MoveNext
      Next
      mLocationID_ = mLocationID_ & ")"
      Call TableMst_CMP(" Where CloseDate is null and LocationID in " & mLocationID_ & " Order By CMPName ")
   Else
      mLocationID_ = "(0"
      
      tmp = "Select Distinct LOcationID from Mst_Location WHere StateID = " & mStateID & " "
      Call TmpTable
      
      For i = 1 To TmpRs.RecordCount
          mLocationID_ = mLocationID_ & "," & TmpRs!LocationID
          TmpRs.MoveNext
      Next
      mLocationID_ = mLocationID_ & ")"
      Call TableMst_CMP(" Where CloseDate is null and LocationID in " & mLocationID_ & "  Order By CMPName ")
   End If
Else
   Call TableMst_CMP(" Where CloseDate is null and LocationID = " & mLocationID & " Order By CMPName ")
End If

Call Grid_Head

If RsMst_CMP.RecordCount > 0 Then
   RsMst_CMP.MoveFirst
   MSHFlexGrid2.Rows = RsMst_CMP.RecordCount + 2
   k = 1
   For i = 0 To RsMst_CMP.RecordCount - 1
       'MSHFlexGrid2.TextMatrix(k, 0) = k
       MSHFlexGrid2.TextMatrix(k, 1) = GetLocationName(RsMst_CMP!LocationID)
       MSHFlexGrid2.TextMatrix(k, 2) = RsMst_CMP!CMPName
       MSHFlexGrid2.TextMatrix(k, 3) = IIf(IsNull(RsMst_CMP!CurrentDate), RsMst_CMP!StartDate, Format(RsMst_CMP!CurrentDate, "dd-mmm-yyyy"))
       MSHFlexGrid2.TextMatrix(k, 4) = RsMst_CMP!CMPID
       MSHFlexGrid2.TextMatrix(k, 5) = IIf(IsNull(RsMst_CMP!BlockDays), "", RsMst_CMP!BlockDays)
       MSHFlexGrid2.row = k
       MSHFlexGrid2.Col = 0
       MSHFlexGrid2.CellFontName = "Wingdings"
       MSHFlexGrid2.CellFontSize = 12
       MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid2.Text = strUnChecked
       k = k + 1
       RsMst_CMP.MoveNext
   Next
Else
   MsgBox "No CMP found for selected location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbRegion_GotFocus()

tmp = CmbRegion.Text

Call TableMst_Region

CmbRegion.Clear

If RsMst_Region.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Region.RecordCount
    CmbRegion.AddItem RsMst_Region!Region
    RsMst_Region.MoveNext
Next
CmbRegion.Text = tmp

End Sub

Private Sub CmbRegion_LostFocus()

If CmbRegion.Text = "" Then
   CmbState.Text = ""
   CmbLocationID.Text = ""
   Call Grid_Head
   Exit Sub
End If
RsMst_Region.MoveFirst
RsMst_Region.Find "Region = '" & CmbRegion.Text & "'"
If RsMst_Region.EOF Then
   MsgBox "Invalid selection "
   CmbRegion.SetFocus
   Exit Sub
End If

If tmp <> CmbRegion.Text Then
   CmbState.Text = ""
   CmbLocationID.Text = ""
   Call Grid_Head
End If

mRegionID = RsMst_Region!RegionID

End Sub
