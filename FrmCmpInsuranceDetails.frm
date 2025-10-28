VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmCmpInsuranceDetails 
   Caption         =   "Cmp Insurance Details"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame0 
      Height          =   5175
      Left            =   6720
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   4695
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   4455
         Begin VB.ComboBox CmbSCMPID 
            Height          =   315
            Left            =   1080
            TabIndex        =   32
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton CmdGO 
            Caption         =   "GO"
            Height          =   375
            Left            =   3360
            TabIndex        =   31
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblCmp 
            Caption         =   "CMP : -"
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
            TabIndex        =   30
            Top             =   240
            Width           =   735
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3975
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   7011
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   6360
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5085
         TabIndex        =   23
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   3870
         TabIndex        =   22
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   2535
         TabIndex        =   21
         Top             =   615
         Width           =   1335
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1320
         TabIndex        =   20
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5085
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3870
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2535
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.ComboBox CmbCMP 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox TxtPolicyType 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   12
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox TxtPolicyAmount 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   11
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox TxtPolicyUpto 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox TxtPolicy 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox TxtInsuranceCompany 
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
         Height          =   375
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox TxtCmpInsuranceUID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label7 
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
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Policy Type"
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
         TabIndex        =   6
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Policy Amount"
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
         TabIndex        =   5
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Policy Validity UpTo"
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
         TabIndex        =   4
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Policy No"
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
         TabIndex        =   3
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Insurance Company"
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
         TabIndex        =   2
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Cmp Insurance UID"
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
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   5265
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmCmpInsuranceDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Edit_Flg, vSCMPID As Variant
'Dim sCMPUid As String
'Private Sub CmbCMP_GotFocus()
'        tmp = CmbCMP.Text
'        Call CMPTable("Order By Name")
'        CmbCMP.Clear
'
'        If CMPRs.RecordCount = 0 Then
'           MsgBox "No CMP found"
'           Exit Sub
'        End If
'
'        For i = 1 To CMPRs.RecordCount
'            CmbCMP.AddItem CMPRs!Name
'            CMPRs.MoveNext
'        Next
'        CmbCMP.Text = tmp
'
'End Sub
'
'
'Private Sub CmbCMP_LostFocus()
'
'    Dim sCMPID As String
'    sCMPID = CmbCMP.Text
'    If sCMPID <> "" Then
'        CMPRs.MoveFirst
'        CMPRs.Find "Name = '" & sCMPID & "'"
'        If CMPRs.RecordCount <> 0 Then
'           sCMPUid = CMPRs!CMPUID
'        End If
'    End If
'
'
'
'End Sub
'
'Private Sub CmbSCMPID_GotFocus()
'      tmp = CmbSCMPID.Text
'    Call CMPTable("Order By Name")
'    CmbSCMPID.Clear
'
'    If CMPRs.RecordCount = 0 Then
'       MsgBox "No CMP found"
'       Exit Sub
'    End If
'
'    For i = 1 To CMPRs.RecordCount
'        CmbSCMPID.AddItem CMPRs!Name
'        CMPRs.MoveNext
'    Next
'    CmbCMP.Text = tmp
'End Sub
'
'Private Sub CmbSCMPID_LostFocus()
'    Dim sCMPID As String
'    sCMPID = CmbSCMPID.Text
'    If sCMPID = "" Then
'      Exit Sub
'    End If
'        CMPRs.MoveFirst
'        CMPRs.Find "Name = '" & sCMPID & "'"
'       If Not CMPRs.EOF Then
'          vSCMPID = CMPRs!CMPUID
'       Else
'        MsgBox "Invalid Location Selection.!!!"
'        CmbSCMPID.Text = ""
'        CmbSCMPID.SetFocus
'    End If
'
'End Sub
'
'
'Private Sub CmdGo_Click()
'
'    Dim mCaId As Variant
'
'    If vSCMPID = "" Then
'       MsgBox "Please Enter CMP. "
'       CmbSCMPID.SetFocus
'       Exit Sub
'    End If
'
'    Call CmpInsuranceDetailsTable("where FK_CMPUID like '" & vSCMPID & "'  Order By FK_CMPUID ")
'
'    If CmpInsuranceDetailsRs.EOF Then
'            'AssayerRs.MoveLast
'            MsgBox "No Record Found"
'            CmbSCMPID.Text = ""
'            CmbSCMPID.SetFocus
'            MSHFlexGrid1.Clear
'            MSHFlexGrid1.Visible = False
'            Exit Sub
'    Else
'        If CmpInsuranceDetailsRs.RecordCount > 1 Then
'            Call setGridColumn
'            Call bindgrid
'          Else
'            Frame0.Visible = False
'            Call InData
'        End If
'    End If
'
'
'
'
'End Sub
'
'''===== Delete Previous Record ===
'Private Sub DeleteCmd_Click()
'    Dim ans As Variant
'    On Error GoTo msgError
'    ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
'    If ans = vbYes Then
'       CmpInsuranceDetailsRs.Delete
'       CmpInsuranceDetailsRs.Update
'       'MsgBox "Record Successfully Removed"
'       If CmpInsuranceDetailsRs.RecordCount = 0 Then
'          Call AddCmd_Click
'          Exit Sub
'       End If
'       CmpInsuranceDetailsRs.MoveNext
'       If CmpInsuranceDetailsRs.EOF() Then
'          CmpInsuranceDetailsRs.MoveLast
'       End If
'       Call InData
'    End If
'
'    Exit Sub
'msgError:
'    MsgBox "Child Record Present "
'    CmpInsuranceDetailsRs.CancelUpdate
'End Sub
'''==== End
'
'''== Edit Previous Records ====
'Private Sub EditCmd_Click()
'Edit_Flg = "E"
'
'Call GButtonLock(Me, False)
'TxtInsuranceCompany.Locked = False
'TxtCmpInsuranceUID.Locked = True
'TxtPolicy.Locked = False
'TxtPolicyAmount.Locked = False
'TxtPolicyType.Locked = False
'TxtPolicyUpto.Locked = False
'CmbCMP.Locked = False
'End Sub
''' End ======
'
'''' === Close Active Form ===
'Private Sub ExitCmd_Click()
'Unload Me
'End Sub
''' End ===
'
''' ==== Show First Record of Table on Controls
'Private Sub FirstCmd_Click()
'CmpInsuranceDetailsRs.MoveFirst
'Call InData
'PreviousCmd.Enabled = False
'FirstCmd.Enabled = False
'NextCmd.Enabled = True
'LastCmd.Enabled = True
'End Sub
'''= End ===
'
'''====To Get The Valide Date from user
''Public Function ValidateDate(ByVal thisDate As Variant) As Boolean
''On Error GoTo errTrap
''
''ValidateDate = True
''
''If IsDate(thisDate) = False Then
''MsgBox "Please enter a valid date!", vbOKOnly
''ValidateDate = False
''Exit Function
''End Function
''''=======End ===========
'
'''=== Form Load Event =====
'
'Private Sub Form_Load()
'Call StatusBarData(Me)
'
'Call CmpInsuranceDetailsTable
'If CmpInsuranceDetailsRs.RecordCount = 0 Then
'  Call AddCmd_Click
'   Exit Sub
'End If
'Call InData
'Call GButtonLock(Me, True)
'End Sub
''' End ======
'
'' == Enable Control for adding New Record ====='
'Private Sub AddCmd_Click()
'Edit_Flg = "A"
'Call GButtonLock(Me, False)
'TxtInsuranceCompany.Locked = False
'TxtCmpInsuranceUID.Locked = True
'TxtPolicy.Locked = False
'TxtPolicyAmount.Locked = False
'TxtPolicyType.Locked = False
'TxtPolicyUpto.Locked = False
'CmbCMP.Locked = False
'
'TxtUnitType = ""
'TxtInsuranceCompany = ""
'TxtCmpInsuranceUID = ""
'TxtPolicy = ""
'TxtPolicyAmount = ""
'TxtPolicyType = ""
'TxtPolicyUpto = ""
'CmbCMP.Text = ""
'If CmpInsuranceDetailsRs.RecordCount > 0 Then
'   CmbCMP.SetFocus
'Else
'   SearchCmd.Enabled = False
'End If
'End Sub
''' End ======
'
''==== Control Binding
'Public Sub InData()
'TxtInsuranceCompany.Locked = True
'TxtCmpInsuranceUID.Locked = True
'TxtPolicy.Locked = True
'TxtPolicyAmount.Locked = True
'TxtPolicyType.Locked = True
'TxtPolicyUpto.Locked = True
'
' Dim sCMPUid As String
'TxtInsuranceCompany = IIf(IsNull(CmpInsuranceDetailsRs!Insurance_Company), "", CmpInsuranceDetailsRs!Insurance_Company)
'TxtCmpInsuranceUID = IIf(IsNull(CmpInsuranceDetailsRs!UID), "", CmpInsuranceDetailsRs!UID)
'TxtPolicy = IIf(IsNull(CmpInsuranceDetailsRs!Policy_No), "", CmpInsuranceDetailsRs!Policy_No)
'TxtPolicyAmount = IIf(IsNull(CmpInsuranceDetailsRs!Policy_Amount), "", CmpInsuranceDetailsRs!Policy_Amount)
'TxtPolicyType = IIf(IsNull(CmpInsuranceDetailsRs!Policy_type), "", CmpInsuranceDetailsRs!Policy_type)
'TxtPolicyUpto = IIf(IsNull(CmpInsuranceDetailsRs!Policy_Validy_upto), "", CmpInsuranceDetailsRs!Policy_Validy_upto)
'sCMPUid = IIf(IsNull(CmpInsuranceDetailsRs!FK_CMPUID), "", CmpInsuranceDetailsRs!FK_CMPUID)
'
'''======== To Get Commodity Uid from Commodity table
'Dim sQuery As String
'sQuery = "WHERE CMP.CMPUID = '" & sCMPUid & "' "
'Call CMPTable(sQuery)
'If CMPRs.RecordCount <> 0 Then
'   CmbCMP.Text = CMPRs!Name
'End If
'''===== End
'CmbCMP.Locked = True
'Call GButtonLock(Me, True)
'End Sub
''' End ======
'
'''=== Show last record of Table on Controls
'Private Sub LastCmd_Click()
'CmpInsuranceDetailsRs.MoveLast
'Call InData
'End Sub
''' === End ===
'
''' === Show Next record on Controls
'Private Sub NextCmd_Click()
'Dim LF_Flag As Variant
'LF_Flag = "N"
'CmpInsuranceDetailsRs.MoveNext
'If CmpInsuranceDetailsRs.EOF Then
'   CmpInsuranceDetailsRs.MoveLast
'   LF_Flag = "Y"
'End If
'Call InData
'If LF_Flag = "Y" Then
'   NextCmd.Enabled = False
'   LastCmd.Enabled = False
'End If
'PreviousCmd.Enabled = True
'FirstCmd.Enabled = True
'End Sub
''' End
'
''' === Show Previous record on Controls
'Private Sub PreviousCmd_Click()
'Dim LF_Flag As Variant
'LF_Flag = "N"
'CmpInsuranceDetailsRs.MovePrevious
'If CmpInsuranceDetailsRs.BOF Then
'   CmpInsuranceDetailsRs.MoveFirst
'   LF_Flag = "Y"
'End If
'Call InData
'If LF_Flag = "Y" Then
'   PreviousCmd.Enabled = False
'   FirstCmd.Enabled = False
'End If
'NextCmd.Enabled = True
'LastCmd.Enabled = True
'
'End Sub
'''=== End ===
'
'''' =====Add New And Edit Record in Table ===
'Private Sub SaveCmd_Click()
'If SaveCmd.Enabled = False Then
'   Exit Sub
'End If
'
'If TxtInsuranceCompany = "" Then
'   MsgBox "Blank Insurance Company not Allowed"
'    TxtInsuranceCompany.SetFocus
'   Exit Sub
'
'End If
'
'If CmbCMP.Text = "" Then
'  MsgBox "Please Select One CMP from List"
'    CmbCMP.SetFocus
'   Exit Sub
'End If
'
'
'If Edit_Flg = "A" Then
'   If CmpInsuranceDetailsRs.RecordCount = 0 Then
'      tmp = GetGUID
'   Else
''        CmpInsuranceDetailsRs.MoveFirst
''        CmpInsuranceDetailsRs.Find "FK_CMPUID = '" & sCMPUid & "'"
''        If Not CmpInsuranceDetailsRs.EOF Then
''           MsgBox "Duplicate CMP not Allowed"
''           CmbCMP.SetFocus
''           Exit Sub
''        End If
'        CmpInsuranceDetailsRs.MoveLast
'        tmp = GetGUID
'    End If
'    CmpInsuranceDetailsRs.AddNew
'    CmpInsuranceDetailsRs!UID = tmp
'    CmpInsuranceDetailsRs!FK_CMPUID = sCMPUid
'    CmpInsuranceDetailsRs!Insurance_Company = TxtInsuranceCompany
'    CmpInsuranceDetailsRs!Policy_No = TxtPolicy
'    If TxtPolicyUpto <> "" Then
'     CmpInsuranceDetailsRs!Policy_Validy_upto = TxtPolicyUpto
'    End If
'
'    CmpInsuranceDetailsRs!Policy_Amount = TxtPolicyAmount
'    CmpInsuranceDetailsRs!Policy_type = TxtPolicyType
'  Else
'   CmpInsuranceDetailsRs.MoveFirst
'   CmpInsuranceDetailsRs.Find "FK_CMPUID = '" & sCMPUid & "'"
'    If Not CmpInsuranceDetailsRs.EOF Then
'       If CmpInsuranceDetailsRs!UID <> TxtCmpInsuranceUID Then
'          MsgBox "Duplicate Client Name not Allowed"
'          Exit Sub
'       End If
'    End If
'   CmpInsuranceDetailsRs.MoveFirst
'   CmpInsuranceDetailsRs.Find "FK_CMPUID = '" & sCMPUid & "'"
'End If
'    CmpInsuranceDetailsRs!FK_CMPUID = sCMPUid
'    CmpInsuranceDetailsRs!Insurance_Company = TxtInsuranceCompany
'    CmpInsuranceDetailsRs!Policy_No = TxtPolicy
'    If TxtPolicyUpto <> "" Then
'     CmpInsuranceDetailsRs!Policy_Validy_upto = TxtPolicyUpto
'    End If
'    CmpInsuranceDetailsRs!Policy_Amount = TxtPolicyAmount
'    CmpInsuranceDetailsRs!Policy_type = TxtPolicyType
'CmpInsuranceDetailsRs.Update
'Call GButtonLock(Me, True)
'TxtCmpInsuranceUID.Locked = True
'TxtInsuranceCompany.Locked = True
'TxtPolicy.Locked = True
'TxtPolicyAmount.Locked = True
'TxtPolicyType.Locked = True
'TxtPolicyUpto.Locked = True
'Call InData
'
'End Sub
'''=====End ====
'
'
'''==== Search Record from Data base and show on grid ===''
'Private Sub SearchCmd_Click()
'If SearchCmd.Caption = "Cancel" Then
'   Call GButtonLock(Me, True)
'   Call FirstCmd_Click
'   Exit Sub
'End If
'
'Call CmpInsuranceDetailsTable("Order BY FK_CMPUID")
'
' CmpInsuranceDetailsRs.MoveFirst
'CmbSCMPID.Text = ""
'Call setGridColumn
'Call bindgrid
'
'
'End Sub
''' End ==
'
'Public Function setGridColumn()
'            MSHFlexGrid1.Clear
'            MSHFlexGrid1.Width = Frame0.Width - 200
'            MSHFlexGrid1.Height = Frame0.Height - 1300
'
'
'            MSHFlexGrid1.Cols = 7
'            MSHFlexGrid1.Rows = CmpInsuranceDetailsRs.RecordCount + 1
'            MSHFlexGrid1.TextMatrix(0, 0) = "Insurance UID"
'            MSHFlexGrid1.TextMatrix(0, 1) = " CMP "
'            MSHFlexGrid1.TextMatrix(0, 2) = "Insurance Company"
'            MSHFlexGrid1.TextMatrix(0, 3) = "Policy No"
'            MSHFlexGrid1.TextMatrix(0, 4) = "Policy Validy upto"
'            MSHFlexGrid1.TextMatrix(0, 5) = "Policy Amount"
'            MSHFlexGrid1.TextMatrix(0, 6) = "Policy Type"
'
'            ''width declaration
'
'            MSHFlexGrid1.ColWidth(0) = 0
'            MSHFlexGrid1.ColWidth(1) = 1500
'            MSHFlexGrid1.ColWidth(2) = 3000
'            MSHFlexGrid1.ColWidth(3) = 1500
'            MSHFlexGrid1.ColWidth(4) = 1500
'            MSHFlexGrid1.ColWidth(5) = 1500
'            MSHFlexGrid1.ColWidth(6) = 1500
'End Function
'
'Public Function bindgrid()
'            Frame0.Height = 5100
'            Frame0.Left = 120
'            Frame0.Top = 120
'            Frame0.Width = 6375
'            Frame0.Visible = True
'            MSHFlexGrid1.Clear
'            Call setGridColumn
'            MSHFlexGrid1.Width = Frame0.Width - 200
'            MSHFlexGrid1.Height = Frame0.Height - 1300
'
'            CmpInsuranceDetailsRs.MoveFirst
'
'            Dim mCmpName As Variant
'            Dim sCMPUid As String
'            For i = 1 To CmpInsuranceDetailsRs.RecordCount
'                 sCMPUid = IIf(IsNull(CmpInsuranceDetailsRs.Fields(1)), "", CmpInsuranceDetailsRs.Fields(1))
'                Call CMPTable("where CMPUID = '" & sCMPUid & "' ")
'                If CMPRs.RecordCount <> 0 Then
'                   mCmpName = CMPRs!Name
'                End If
'
'                MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(CmpInsuranceDetailsRs.Fields(0)), "", CmpInsuranceDetailsRs.Fields(0))
'                MSHFlexGrid1.TextMatrix(i, 1) = mCmpName
'                MSHFlexGrid1.TextMatrix(i, 2) = IIf(IsNull(CmpInsuranceDetailsRs.Fields(2)), "", CmpInsuranceDetailsRs.Fields(2))
'                MSHFlexGrid1.TextMatrix(i, 3) = IIf(IsNull(CmpInsuranceDetailsRs.Fields(3)), "", CmpInsuranceDetailsRs.Fields(3))
'                MSHFlexGrid1.TextMatrix(i, 4) = IIf(IsNull(CmpInsuranceDetailsRs.Fields(4)), "", CmpInsuranceDetailsRs.Fields(4))
'                MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(CmpInsuranceDetailsRs.Fields(5)), "", CmpInsuranceDetailsRs.Fields(5))
'                MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(CmpInsuranceDetailsRs.Fields(6)), "", CmpInsuranceDetailsRs.Fields(6))
'
'
'
'             CmpInsuranceDetailsRs.MoveNext
'          Next
'End Function
'
''''=====  Get the Grid Event=====
'
'Private Sub MSHFlexGrid1_Click()
'Dim x As Variant
'Dim mCaId As Variant
'
'x = MSHFlexGrid1.RowSel
'
'mCaId = MSHFlexGrid1.TextMatrix(x, 0)
'
''CmpInsuranceDetailsRs.MoveFirst
'
'If CmpInsuranceDetailsRs.RecordCount = 0 Then
'   Call AddCmd_Click
'   Exit Sub
'End If
'CmpInsuranceDetailsRs.MoveFirst
'CmpInsuranceDetailsRs.Find "UID = '" & mCaId & " '"
'If CmpInsuranceDetailsRs.EOF Then
'   CmpInsuranceDetailsRs.MoveLast
'End If
'Frame0.Visible = False
'Call InData
'
'End Sub
'
''''== == End
'
'Private Sub TxtPolicyUpto_LostFocus()
'     If (TxtPolicyUpto) <> "" Then
'       If IsDate(TxtPolicyUpto) = False Then
'          MsgBox "Invalid Date Format !!!!!"
'          TxtPolicyUpto.SetFocus
'       End If
'    End If
'End Sub
