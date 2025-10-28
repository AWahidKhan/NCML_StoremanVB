VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_QualityParameter 
   Caption         =   "Quality Parameter"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox TxtSQualityParameter 
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
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   4575
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
         Height          =   350
         Left            =   6960
         TabIndex        =   13
         Top             =   240
         Width           =   1000
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
         Height          =   350
         Left            =   7965
         TabIndex        =   14
         Top             =   240
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   120
         TabIndex        =   16
         Top             =   795
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3916
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label23 
         Caption         =   "Total Records Found :"
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
         Left            =   9600
         TabIndex        =   30
         Top             =   288
         Width           =   2370
      End
      Begin VB.Label Label22 
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
         TabIndex        =   29
         Top             =   300
         Width           =   90
      End
      Begin VB.Label Label3 
         Caption         =   "Quality Parameter"
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
         TabIndex        =   28
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   17
      Top             =   165
      Width           =   7395
      Begin VB.ComboBox CmbFumigationApplicable 
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
         ItemData        =   "FrmMst_QualityParameter.frx":0000
         Left            =   2625
         List            =   "FrmMst_QualityParameter.frx":000A
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox TxtQualityParameter 
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
         Left            =   2640
         TabIndex        =   1
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox TxtQualityParameterID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtCreated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox TxtUpdated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2160
         Width           =   4575
      End
      Begin VB.TextBox TxtFumigationApplicable 
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
         Left            =   2625
         MaxLength       =   25
         TabIndex        =   26
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Fumigation Applicable"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1260
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Quality Parameter"
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
         Left            =   240
         TabIndex        =   24
         Top             =   780
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Quality Parameter ID"
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
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Created  :"
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
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Updated :"
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
         Left            =   240
         TabIndex        =   21
         Top             =   2160
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1140
      Left            =   120
      TabIndex        =   25
      Top             =   2805
      Width           =   7410
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
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
         Left            =   5775
         TabIndex        =   7
         Top             =   615
         Width           =   1395
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
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
         Left            =   4380
         TabIndex        =   8
         Top             =   615
         Width           =   1395
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
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
         Left            =   2985
         TabIndex        =   9
         Top             =   615
         Width           =   1395
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
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
         Left            =   1590
         TabIndex        =   10
         Top             =   615
         Width           =   1395
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
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
         Left            =   195
         TabIndex        =   11
         Top             =   615
         Width           =   1395
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
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
         Left            =   5775
         TabIndex        =   6
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
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
         Left            =   4380
         TabIndex        =   5
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
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
         Left            =   2985
         TabIndex        =   4
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton SaveCmd 
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
         Left            =   1590
         TabIndex        =   3
         Top             =   240
         Width           =   1395
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
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
         Left            =   195
         TabIndex        =   0
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Commodity Quality Parameter Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9780
      Left            =   7680
      TabIndex        =   31
      Top             =   165
      Width           =   7455
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   9210
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   16245
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmMst_QualityParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtQualityParameter.Locked = False
TxtQualityParameter = ""
TxtQualityParameterID = ""
CmbFumigationApplicable.Visible = True
TxtFumigationApplicable.Visible = False
CmbFumigationApplicable.Text = ""

TxtCreated = ""
TxtUpdated = ""
Call Grid_HeadCommodityDetails

If RsMst_QualityParameter.RecordCount > 0 Then
   TxtQualityParameter.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmbFumigationApplicable_LostFocus()
If CmbFumigationApplicable.Text = "" Then Exit Sub

If LCase(CmbFumigationApplicable.Text) <> "yes" And LCase(CmbFumigationApplicable.Text) <> "no" Then
   MsgBox "Invalid Selection."
   CmbFumigationApplicable.SetFocus
End If
End Sub

Private Sub CmdGo_Click()
Call Grid_Head

Label22.Caption = ""

tmp = "Select QualityParameterID,QualityParameter,FumigationApplicable from Mst_QualityParameter"

If TxtSQualityParameter.Text <> "" Then
   tmp = tmp & "  Where QualityParameter Like '%" & TxtSQualityParameter.Text & "%'"
End If

tmp = tmp & " Order By QualityParameter"
Call TmpTable
For I = 1 To TmpRs.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!QualityParameterID
    MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!QualityParameter
    If LCase(TmpRs!FumigationApplicable) = "y" Then
       MSHFlexGrid1.TextMatrix(I, 2) = "Yes"
    ElseIf LCase(TmpRs!FumigationApplicable) = "n" Then
       MSHFlexGrid1.TextMatrix(I, 2) = "No"
    End If
    TmpRs.MoveNext
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
Next
Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   
   RsMst_QualityParameter.Delete
   RsMst_QualityParameter.Update
   If RsMst_QualityParameter.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_QualityParameter.MoveNext
   If RsMst_QualityParameter.EOF() Then
      RsMst_QualityParameter.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_QualityParameter.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtQualityParameter.Locked = False
CmbFumigationApplicable.Visible = True
TxtFumigationApplicable.Visible = False
CmbFumigationApplicable.Text = TxtFumigationApplicable.Text

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_QualityParameter.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
TxtSQualityParameter.Text = ""
Label22.Caption = ""
Call Grid_Head

Call TableMst_QualityParameter(" Order by QualityParameterID ")

If RsMst_QualityParameter.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata

Call GButtonLock(Me, True)


End Sub

Public Sub Indata()
TxtQualityParameter.Locked = True
TxtFumigationApplicable.Locked = True

CmbFumigationApplicable.Visible = False
TxtFumigationApplicable.Visible = True
TxtFumigationApplicable.Text = ""

TxtQualityParameterID = RsMst_QualityParameter!QualityParameterID
TxtQualityParameter = RsMst_QualityParameter!QualityParameter
If LCase(RsMst_QualityParameter!FumigationApplicable) = "y" Then
   TxtFumigationApplicable.Text = "Yes"
ElseIf LCase(RsMst_QualityParameter!FumigationApplicable) = "n" Then
   TxtFumigationApplicable.Text = "No"
Else
   TxtFumigationApplicable.Text = ""
End If
TxtCreated = IIf(IsNull(RsMst_QualityParameter!CreatedBy), "", RsMst_QualityParameter!CreatedBy) & " :- " & IIf(IsNull(RsMst_QualityParameter!CreatedDate), "", RsMst_QualityParameter!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_QualityParameter!UpdatedBy), "", RsMst_QualityParameter!UpdatedBy) & " :- " & IIf(IsNull(RsMst_QualityParameter!UpdatedDate), "", RsMst_QualityParameter!UpdatedDate)
Call FillCommodityDetails
Call GButtonLock(Me, True)
End Sub

Private Sub FillCommodityDetails()
Call Grid_HeadCommodityDetails

tmp = "Select A.CommodityQualityParameterID,MC.Commodity,A.MinValue,A.MaxValue"
tmp = tmp & " From Mst_CommodityQualityParameter A"
tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=A.CommodityID"
tmp = tmp & " Where A.QualityParameterID=" & Val(TxtQualityParameterID.Text) & " "
tmp = tmp & " Order by MC.Commodity"
Call TmpTable
For I = 1 To TmpRs.RecordCount
    For c = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid2.TextMatrix(I, c) = TmpRs.Fields(c)
    Next
    TmpRs.MoveNext
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
Next
End Sub

Private Sub Grid_HeadCommodityDetails()
With MSHFlexGrid2
     .Clear
     .Font.Size = 10
     .Rows = 2
     .Cols = 4
     .AllowUserResizing = flexResizeBoth
     .WordWrap = True
     .RowHeight(0) = 750
     .TextMatrix(0, 0) = "Commodity Quality Parameter ID"
     .TextMatrix(0, 1) = "Commodity"
     .TextMatrix(0, 2) = "Minimum Value"
     .TextMatrix(0, 3) = "Maximum Value"
     .ColWidth(0) = 1300
     .ColWidth(1) = 3500
     .ColWidth(2) = 1000
     .ColWidth(3) = 1000
End With
End Sub
Private Sub LastCmd_Click()
RsMst_QualityParameter.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
If mCaId = "" Then Exit Sub
RsMst_QualityParameter.MoveFirst

If RsMst_QualityParameter.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_QualityParameter.MoveFirst
RsMst_QualityParameter.Find "QualityParameterID = " & Val(mCaId) & ""
If RsMst_QualityParameter.EOF Then
   RsMst_QualityParameter.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_QualityParameter.MoveNext
If RsMst_QualityParameter.EOF Then
   RsMst_QualityParameter.MoveLast
   LF_Flag = "Y"
End If
Call Indata
If LF_Flag = "Y" Then
   NextCmd.Enabled = False
   LastCmd.Enabled = False
End If
PreviousCmd.Enabled = True
FirstCmd.Enabled = True

End Sub

Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_QualityParameter.MovePrevious
If RsMst_QualityParameter.BOF Then
   RsMst_QualityParameter.MoveFirst
   LF_Flag = "Y"
End If
Call Indata
If LF_Flag = "Y" Then
   PreviousCmd.Enabled = False
   FirstCmd.Enabled = False
End If
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub SaveCmd_Click()
If SaveCmd.Enabled = False Then
   Exit Sub
End If

If TxtQualityParameter = "" Then
   MsgBox "Blank QualityParameter  not Allowed"
   TxtQualityParameter.SetFocus
   Exit Sub
End If

If CmbFumigationApplicable.Text = "" Then
   MsgBox "Blank Fumigation Applicable Flag Not Allowed !!"
   CmbFumigationApplicable.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_QualityParameter.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_QualityParameter.MoveFirst
        RsMst_QualityParameter.Find "QualityParameter = '" & TxtQualityParameter & "'"
        If Not RsMst_QualityParameter.EOF Then
           MsgBox "Duplicate Quality Parameter not Allowed"
           Exit Sub
        End If
        RsMst_QualityParameter.MoveLast
        
        tmp = "select max(QualityParameterID) from Mst_QualityParameter"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_QualityParameter.AddNew
    RsMst_QualityParameter!QualityParameterID = tmp
    TxtQualityParameterID = tmp
Else
    RsMst_QualityParameter.MoveFirst
    RsMst_QualityParameter.Find "QualityParameter = '" & TxtQualityParameter & "'"
    If Not RsMst_QualityParameter.EOF Then
       If RsMst_QualityParameter!QualityParameterID <> TxtQualityParameterID Then
          MsgBox "Duplicate QualityParameter not Allowed"
          Exit Sub
       End If
    End If
   RsMst_QualityParameter.MoveFirst
   RsMst_QualityParameter.Find "QualityParameterID = " & TxtQualityParameterID & ""
End If
RsMst_QualityParameter!QualityParameter = TxtQualityParameter
RsMst_QualityParameter!FumigationApplicable = UCase(Left(CmbFumigationApplicable.Text, 1))

If IsNull(RsMst_QualityParameter!CreatedBy) = True Or RsMst_QualityParameter!CreatedBy = "" Then
   RsMst_QualityParameter!CreatedBy = Gen_UserLoginID
   RsMst_QualityParameter!CreatedDate = Now
Else
   RsMst_QualityParameter!UpdatedBy = Gen_UserLoginID
   RsMst_QualityParameter!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If
RsMst_QualityParameter.Update
RsMst_QualityParameter.Requery
RsMst_QualityParameter.MoveFirst
RsMst_QualityParameter.Find "QualityParameterID = " & TxtQualityParameterID & ""

Call GButtonLock(Me, True)
TxtQualityParameter.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_QualityParameter.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_QualityParameter.EOF Then
      RsMst_QualityParameter.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 600
Frame0.Width = Me.Width - 300
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top

Frame0.Visible = True


MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 900
RsMst_QualityParameter.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'TxtSQualityParameter.Text = ""
'Label22.Caption = ""
'Call Grid_Head
''===========comment on 7 july 2012 for not clearing search grid=============


End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Font.Size = 10
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 500
MSHFlexGrid1.TextMatrix(0, 0) = "Quality Parameter ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Quality Parameter"
MSHFlexGrid1.TextMatrix(0, 2) = "Fumigation Applicable"
MSHFlexGrid1.ColWidth(0) = 1600
MSHFlexGrid1.ColWidth(1) = 9000
MSHFlexGrid1.ColWidth(2) = 2300
End Sub
Private Sub TxtQualityParameter_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtQualityParameter")
End If
End Sub

Private Sub TxtQualityParameter_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtQualityParameter)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtQualityParameter.SetFocus
End If
End Sub



Private Sub TxtQualityParameterID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtQualityParameterID")
End If
End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & "Mst_QualityParameter.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & "Mst_QualityParameter.xls")
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
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

