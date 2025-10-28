VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_GodownWiseApprovalMargin 
   Caption         =   "Godown Wise Approval Margin"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   MDIChild        =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   10500
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   15015
      Begin VB.TextBox TxtMinimumRent 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   960
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.CheckBox CmdSelectAll 
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
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   13
         Top             =   9280
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.TextBox TxtLimit 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
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
         Left            =   10680
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmbGet 
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
         Height          =   375
         Left            =   9480
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox CmbLocation 
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
         Left            =   5490
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   270
         Width           =   3810
      End
      Begin VB.ComboBox CmbState 
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
         ItemData        =   "FrmTxn_GodownWiseApprovalMargin.frx":0000
         Left            =   960
         List            =   "FrmTxn_GodownWiseApprovalMargin.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   3150
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   14631
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox TxtState 
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   272
         Width           =   2895
      End
      Begin VB.TextBox TxtLocation 
         Appearance      =   0  'Flat
         Height          =   310
         Left            =   5490
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   272
         Width           =   3780
      End
      Begin VB.Label lblTotalRecords 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   14055
         TabIndex        =   15
         Top             =   330
         Width           =   75
      End
      Begin VB.Label Label23 
         Caption         =   "Total Records Found :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   12000
         TabIndex        =   14
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Location"
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
         Left            =   4230
         TabIndex        =   10
         Top             =   300
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "State"
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
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmTxn_GodownWiseApprovalMargin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mStateID, mLocationID As Variant
Dim tmp1 As Variant
Dim mSelRow, mSelCol As Variant

Public Sub Grid_Head()
With MSHFlexGrid1
     .Font.Size = 10
     .Clear
     .Rows = 2
     .Cols = 10
     .FixedCols = 0
     .WordWrap = True
     .TextMatrix(0, 0) = "Select":                                              .ColWidth(0) = 800
     .TextMatrix(0, 1) = "CMP Name":                                            .ColWidth(1) = 2000
     .TextMatrix(0, 2) = "Godown No":                                           .ColWidth(2) = 1200
     .TextMatrix(0, 3) = "Address":                                             .ColWidth(3) = 4000
     .TextMatrix(0, 4) = "Godown Type":                                         .ColWidth(4) = 1600
     .TextMatrix(0, 5) = "Area in Sq. Ft.":                                     .ColWidth(5) = 1500
     .TextMatrix(0, 6) = "AM/RM Approval Limit (%)":                            .ColWidth(6) = 1700
     .TextMatrix(0, 7) = "Minimum Rent Value in Sq. Ft. (Franchisee)":          .ColWidth(7) = 1700
     .TextMatrix(0, 8) = "GodownID":                                            .ColWidth(8) = 0
     .TextMatrix(0, 9) = "UnitTypeID":                                          .ColWidth(9) = 0
     .RowHeight(0) = 800
End With
End Sub

Private Sub CmbGet_Click()
Dim mR As Variant

If Trim(CmbState.Text) = "" Then
   MsgBox "Select State!!!"
   CmbState.SetFocus
   Exit Sub
End If
If Trim(CmbLocation.Text) = "" Then
   MsgBox "Select Location!!!"
   CmbLocation.SetFocus
   Exit Sub
End If

Call Grid_Head
tmp = " SELECT MG.CMPID,CMP.CMPName,MG.GodownID,MG.GodownNo,MG.Address,MG.Area,MG.ApprovalLimit,MG.RentMinimumValue,MG.UnitTypeID,MUT.UnitType "
tmp = tmp & " FROM Mst_Godown MG "
tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID "
tmp = tmp & " Inner Join Mst_UnitType MUT On MG.UnitTypeID = MUT.UnitTypeID "
tmp = tmp & " Where MG.AgreementEndDate >= '" & Format(Date, "YYYY-MM-DD") & "' And CMP.LocationID = " & mLocationID & " Order By MG.GodownNo "
Call TmpTable
If TmpRs.RecordCount > 0 Then
With MSHFlexGrid1

    ProgressBar1.Max = TmpRs.RecordCount
    For mR = 1 To TmpRs.RecordCount
        .row = mR
        .Col = 0
        .CellFontName = "Wingdings"
        .CellFontSize = 12
        .CellAlignment = flexAlignCenterCenter
        .Text = strUnChecked
        .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
        .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
        .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
        .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
        .TextMatrix(mR, 5) = Format(IIf(IsNull(TmpRs!Area), "", TmpRs!Area), "#0.000")
        .TextMatrix(mR, 6) = Format(IIf(IsNull(TmpRs!ApprovalLimit), "", TmpRs!ApprovalLimit), "#0.00")
        .TextMatrix(mR, 7) = Format(IIf(IsNull(TmpRs!RentMinimumValue), "", TmpRs!RentMinimumValue), "#0.00")
        .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
        .TextMatrix(mR, 9) = IIf(IsNull(TmpRs!UnitTypeID), "", TmpRs!UnitTypeID)
        .Rows = .Rows + 1
        TmpRs.MoveNext
        ProgressBar1.Value = mR
    Next mR
End With
End If

lblTotalRecords.Caption = TmpRs.RecordCount

End Sub

Private Sub CmbLocation_GotFocus()
If Trim(CmbState.Text) = "" Then
   MsgBox "Select State!!!"
   CmbState.SetFocus
   Exit Sub
End If
tmp1 = CmbLocation.Text
tmp = CmbLocation.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location(" Where StateID =" & mStateID & "")
Else
   Call TableMst_Location(" Where " & Gen_WcLocation & " And StateID = " & mStateID & " ")
End If
CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
   CmbLocation.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub

Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then Call ClearGrid: Exit Sub
If RsMst_Location.RecordCount > 0 Then RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"

If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocation.SetFocus
   Exit Sub
End If
If RsMst_Location!LocationName <> CmbLocation Then Call ClearGrid
mLocationID = RsMst_Location.Fields("LocationID")
mSM_Prefix = RsMst_Location.Fields("SM_Prefix")
If LCase(tmp1) <> LCase(CmbLocation.Text) Then Call ClearGrid
End Sub

Private Sub ClearGrid()
Call Grid_Head
TxtLimit.Visible = False
CmdSave.Enabled = False
End Sub


Private Sub CmbState_GotFocus()
tmp = CmbState.Text
CmbLocation.Text = ""
Call Grid_Head
CmdSave.Enabled = False
tmp1 = ""
If Gen_WcState <> "" Then
   tmp1 = "And " & Gen_WcState
End If

Call TableMst_State(" Where StateID <> 0 " & tmp1)
CmbState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record Found!"
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
   CmbLocation.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbState.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbState.SetFocus
   Exit Sub
End If

If RsMst_State!StateName <> CmbState.Text Then
   CmbLocation.Text = ""
   Call ClearGrid
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmdSave_Click()
Dim mR As Double
With MSHFlexGrid1
    If CmbState.Text <> "" And CmbLocation.Text <> "" Then
        ''-- Save the Approval Limit godownID Wise
        tmp = " SELECT * FROM Mst_Godown Where CMPID IN (SELECT CMPID From Mst_CMP WHere LocationID = " & mLocationID & " ) And AgreementEndDate >= '" & Format(Date, "YYYY-MM-DD") & "'"
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
            ProgressBar1.Max = .Rows - 2
            For mR = 1 To .Rows - 2
                TmpRs.MoveFirst
                
                ''--- Update Values
                If .TextMatrix(mR, 0) = strChecked Then
                    If .TextMatrix(mR, 8) <> "" Then
                       TmpRs.Find "GodownID =" & Val(.TextMatrix(mR, 8))
                       If TmpRs.EOF = False Then
                           TmpRs!ApprovalLimit = Format(Val(.TextMatrix(mR, 6)), "#0.00")
                           If Val(.TextMatrix(mR, 9)) = 2 Then
                            TmpRs!RentMinimumValue = Format(Val(.TextMatrix(mR, 7)), "#0.00")
                           End If
                           TmpRs!UpdatedDate = Now
                           TmpRs!UpdatedBy = Gen_UserLoginID
                           TmpRs.Update
                       End If
                    End If
                End If
                
                ProgressBar1.Value = mR
            Next mR
        End If
        MsgBox "Done."
    Else
        MsgBox "Select Location"
        Exit Sub
    End If
End With
End Sub

Private Sub CmdSelectAll_Click()
With MSHFlexGrid1
    If CmdSelectAll.Value = 1 Then
        For i = 1 To .Rows - 1
           If .TextMatrix(i, 1) <> "" Then
              .row = i
              .Col = 0
              .CellFontName = "Wingdings"
              .CellFontSize = 12
              .CellAlignment = flexAlignCenterCenter
              .Text = strChecked
          End If
        Next
    Else
        For i = 1 To .Rows - 1
           If .TextMatrix(i, 1) <> "" Then
              .row = i
              .Col = 0
              .CellFontName = "Wingdings"
              .CellFontSize = 12
              .CellAlignment = flexAlignCenterCenter
              .Text = strUnChecked
          End If
        Next
    End If
End With
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call ClearGrid
End Sub

Private Sub MSHFlexGrid1_Click()
With MSHFlexGrid1
    mSelCol = .Col
    mSelRow = .row
    TxtLimit.Visible = False
    TxtMinimumRent.Visible = False
    
    If .Col = 0 Then
       .Col = 0
       If .TextMatrix(.row, .Col) = strUnChecked Then
          .TextMatrix(.row, .Col) = strChecked
       Else
          .TextMatrix(.row, .Col) = strUnChecked
       End If
    End If
    
    If .TextMatrix(mSelRow, 0) = strChecked Then
        If .TextMatrix(mSelRow, 1) = "" Then Exit Sub
        If .TextMatrix(mSelRow, 8) <> "" Then
            If .Col = 6 Then
                TxtLimit.Visible = True
                TxtLimit.Left = .Left + .CellLeft + 5
                TxtLimit.Top = .Top + .CellTop + 5
                TxtLimit.Width = .CellWidth - 10
                TxtLimit.Height = .CellHeight - 10
                TxtLimit.Text = .Text
                TxtLimit.SelStart = 0
                TxtLimit.SelLength = Len(.Text)
                TxtLimit.SetFocus
            ElseIf .Col = 7 Then
                If Val(.TextMatrix(mSelRow, 9)) = 2 Then
                    TxtMinimumRent.Visible = True
                    TxtMinimumRent.Left = .Left + .CellLeft + 5
                    TxtMinimumRent.Top = .Top + .CellTop + 5
                    TxtMinimumRent.Width = .CellWidth - 10
                    TxtMinimumRent.Height = .CellHeight - 10
                    TxtMinimumRent.Text = .Text
                    TxtMinimumRent.SelStart = 0
                    TxtMinimumRent.SelLength = Len(.Text)
                    TxtMinimumRent.SetFocus
                Else
                    TxtMinimumRent.Visible = False
                End If
            Else
                TxtLimit.Visible = False
                TxtMinimumRent.Visible = False
            End If
        End If
    Else
          .TextMatrix(mSelRow, 6) = ""
          .TextMatrix(mSelRow, 7) = ""
    End If
    
End With
End Sub

Private Sub MSHFlexGrid1_Scroll()
TxtLimit.Visible = False
TxtMinimumRent.Visible = False
End Sub

Private Sub TxtLimit_LostFocus()
With MSHFlexGrid1

If .TextMatrix(mSelRow, 0) = strChecked Then
    If TxtLimit.Text = "" Then
        TxtLimit.Visible = False
        .TextMatrix(mSelRow, mSelCol) = IIf(Trim(TxtLimit) = "", "", Format(0, "#0.00"))
        Exit Sub
    End If
    If Val(TxtLimit) < 0 Or Not IsNumeric(TxtLimit) Then
       MsgBox "Invalid Numeric Format!!!"
       TxtLimit.SelStart = 0
       TxtLimit.SelLength = Len(TxtLimit)
       TxtLimit.SetFocus
       Exit Sub
    End If
    If Val(TxtLimit.Text) > 100 Then
         MsgBox "Approval Limit can not more then 100%."
         TxtLimit.Visible = True
         TxtLimit.SetFocus
         Exit Sub
    End If
    .TextMatrix(mSelRow, mSelCol) = IIf(Trim(TxtLimit) = "", "", Format(TxtLimit, "#0.00"))
    TxtLimit.Visible = False
    CmdSave.Enabled = True
    .SetFocus
End If

End With
End Sub

Private Sub TxtMinimumRent_LostFocus()
With MSHFlexGrid1

If .TextMatrix(mSelRow, 0) = strChecked Then
    If TxtMinimumRent.Text = "" Then
        TxtMinimumRent.Visible = False
        .TextMatrix(mSelRow, mSelCol) = IIf(Trim(TxtMinimumRent) = "", "", Format(0, "#0.00"))
        Exit Sub
    End If
    If Val(TxtMinimumRent) < 0 Or Not IsNumeric(TxtMinimumRent) Then
       MsgBox "Invalid Numeric Format!!!"
       TxtMinimumRent.SelStart = 0
       TxtMinimumRent.SelLength = Len(TxtMinimumRent)
       TxtMinimumRent.SetFocus
       Exit Sub
    End If
    .TextMatrix(mSelRow, mSelCol) = IIf(Trim(TxtMinimumRent) = "", "", Format(TxtMinimumRent, "#0.00"))
    TxtMinimumRent.Visible = False
    CmdSave.Enabled = True
    .SetFocus
End If

End With
End Sub
