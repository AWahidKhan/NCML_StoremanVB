VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_DocumentType 
   Caption         =   "Document Type"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1305
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   14055
      Begin VB.TextBox TxtSearchDocumentType 
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
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   465
         Width           =   5415
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
         Left            =   8385
         TabIndex        =   9
         Top             =   450
         Width           =   1020
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
         Left            =   7350
         TabIndex        =   8
         Top             =   450
         Width           =   1020
      End
      Begin VB.ComboBox CmbSFlag 
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
         ItemData        =   "FrmMst_DocumentType.frx":0000
         Left            =   5640
         List            =   "FrmMst_DocumentType.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   7
         Text            =   "Active"
         Top             =   480
         Width           =   1620
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   11
         Top             =   885
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
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
      Begin VB.Label Label22 
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
         Left            =   11985
         TabIndex        =   15
         Top             =   465
         Width           =   75
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
         Left            =   9585
         TabIndex        =   14
         Top             =   465
         Width           =   2355
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Document Type"
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
         Left            =   2107
         TabIndex        =   13
         Top             =   210
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Status"
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
         Left            =   6120
         TabIndex        =   12
         Top             =   210
         Width           =   555
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   26
      Top             =   2640
      Width           =   12330
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
         Height          =   400
         Left            =   9645
         TabIndex        =   5
         Top             =   660
         Width           =   2385
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   9645
         TabIndex        =   4
         Top             =   240
         Width           =   2385
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
         Height          =   400
         Left            =   7260
         TabIndex        =   32
         Top             =   660
         Width           =   2385
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
         Height          =   400
         Left            =   7260
         TabIndex        =   31
         Top             =   240
         Width           =   2385
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
         Height          =   400
         Left            =   4875
         TabIndex        =   30
         Top             =   660
         Width           =   2385
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
         Height          =   400
         Left            =   4875
         TabIndex        =   29
         Top             =   240
         Width           =   2385
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
         Height          =   400
         Left            =   2490
         TabIndex        =   28
         Top             =   660
         Width           =   2385
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
         Height          =   400
         Left            =   2490
         TabIndex        =   3
         Top             =   240
         Width           =   2385
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
         Height          =   400
         Left            =   100
         TabIndex        =   27
         Top             =   660
         Width           =   2385
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   100
         TabIndex        =   0
         Top             =   240
         Width           =   2385
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   12330
      Begin VB.TextBox TxtDocumentType 
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
         Left            =   2310
         MaxLength       =   500
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   9930
      End
      Begin VB.TextBox TxtDocumentTypeID 
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
         Height          =   360
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1560
      End
      Begin VB.ComboBox CmbStatusFlag 
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
         ItemData        =   "FrmMst_DocumentType.frx":0020
         Left            =   2310
         List            =   "FrmMst_DocumentType.frx":002A
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   1905
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
         Height          =   360
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1665
         Width           =   9930
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
         Height          =   360
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2145
         Width           =   9930
      End
      Begin VB.TextBox TxtStatusFlag 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2310
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label LblFlag 
         Caption         =   "Document Type"
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
         Left            =   240
         TabIndex        =   25
         Top             =   754
         Width           =   2280
      End
      Begin VB.Label Lable1 
         Caption         =   "Document Type ID"
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
         Left            =   240
         TabIndex        =   24
         Top             =   300
         Width           =   1920
      End
      Begin VB.Label Label1 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   23
         Top             =   1275
         Width           =   1800
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
         Height          =   300
         Left            =   240
         TabIndex        =   22
         Top             =   1755
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
         Height          =   225
         Left            =   240
         TabIndex        =   21
         Top             =   2220
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmMst_DocumentType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant


Private Sub CmdGo_Click()

tmp = "Select DocumentTypeID,DocumentType,Flag from Mst_DocumentType "

Dim wc As Variant
wc = ""
If TxtSearchDocumentType <> "" Then
   wc = " Where DocumentType Like '%" & TxtSearchDocumentType & "%' "
End If

If CmbSFlag.Text <> "" Then
   If wc = "" Then
      wc = " Where Flag='" & Left$(CmbSFlag.Text, 1) & "' " 'Left$(CmbStatusFlag.Text, 1)
   Else
      wc = wc & " And Flag='" & Left$(CmbSFlag.Text, 1) & "' " 'Left$(CmbStatusFlag.Text, 1)
   End If
End If

tmp = tmp & wc & " Order By DocumentType "

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    MSHFlexGrid1.TextMatrix(iu, 2) = IIf(MSHFlexGrid1.TextMatrix(iu, 2) = "A", "Active", "Deactive")
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Call Grid_Head
TxtSearchDocumentType = ""
Label22.Caption = ""

Call TableMst_DocumentType
If RsMst_DocumentType.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
Call InData
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_DocumentType.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_DocumentType.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
Private Sub NextCmd_Click()
Dim LF_PerLimit As Variant
LF_PerLimit = "N"
RsMst_DocumentType.MoveNext
If RsMst_DocumentType.EOF Then
   RsMst_DocumentType.MoveLast
   LF_PerLimit = "Y"
End If
Call InData
If LF_PerLimit = "Y" Then
    NextCmd.Enabled = False
    LastCmd.Enabled = False
End If
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_PerLimit As Variant
LF_PerLimit = "N"
RsMst_DocumentType.MovePrevious
If RsMst_DocumentType.BOF Then
   RsMst_DocumentType.MoveFirst
   LF_PerLimit = "Y"
End If
Call InData
If LF_PerLimit = "Y" Then
    PreviousCmd.Enabled = False
    FirstCmd.Enabled = False
End If
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsMst_DocumentType.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_DocumentType.MoveLast
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_DocumentType.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call InData
   Exit Sub
End If

Frame0.Height = Me.Height - 250
Frame0.Left = Frame1.Left ' - 100
Frame0.Top = Frame1.Top ' - 100
Frame0.Width = Me.Width - 600
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_DocumentType.MoveFirst
'--Commented By sagar on 7-7-2012

''Call Grid_Head
''TxtSearchDocumentType = ""
''Label22.Caption = ""

End Sub
Public Sub InData()

TxtStatusFlag.Locked = True
TxtStatusFlag.Visible = True
CmbStatusFlag.Visible = False
TxtDocumentTypeID.Locked = True
TxtDocumentType.Locked = True
TxtDocumentTypeID = IIf(IsNull(RsMst_DocumentType!DocumentTypeID), "", RsMst_DocumentType!DocumentTypeID)
TxtDocumentType = IIf(IsNull(RsMst_DocumentType!DocumentType), "", RsMst_DocumentType!DocumentType)
If LCase(RsMst_DocumentType!Flag) = "a" Then
   TxtStatusFlag.Text = "Active"
ElseIf LCase(RsMst_DocumentType!Flag) = "d" Then
   TxtStatusFlag.Text = "DeActive"
Else
   TxtStatusFlag.Text = ""
End If

TxtCreated = IIf(IsNull(RsMst_DocumentType!CreatedBy), "", RsMst_DocumentType!CreatedBy) & " :- " & IIf(IsNull(RsMst_DocumentType!CreatedDate), "", RsMst_DocumentType!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_DocumentType!UpdatedBy), "", RsMst_DocumentType!UpdatedBy) & " :- " & IIf(IsNull(RsMst_DocumentType!UpdatedDate), "", RsMst_DocumentType!UpdatedDate)


If RsMst_DocumentType.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtDocumentTypeID.Locked = True
TxtDocumentType.Locked = False
TxtDocumentTypeID = ""
TxtDocumentType = ""

TxtCreated = ""
TxtUpdated = ""

CmbStatusFlag.Visible = True
TxtStatusFlag.Visible = False
CmbStatusFlag.Text = ""

If RsMst_DocumentType.RecordCount > 0 Then
   TxtDocumentType.SetFocus
End If
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtDocumentTypeID.Locked = False
CmbStatusFlag.Visible = True
TxtStatusFlag.Visible = False
CmbStatusFlag.Text = TxtStatusFlag.Text

End Sub
Private Sub DeleteCmd_Click()
Dim Ans As Variant
On Error GoTo msgError
Ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If Ans = vbYes Then
   RsMst_DocumentType.Delete
   RsMst_DocumentType.Update
   If RsMst_DocumentType.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_DocumentType.MoveNext
   If RsMst_DocumentType.EOF() Then
      RsMst_DocumentType.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_DocumentType.CancelUpdate
End Sub
Private Sub SaveCmd_Click()

If TxtDocumentType.Text = "" Then
   MsgBox "Blank DocumentType Not Allowed !!! "
   TxtDocumentType.SetFocus
   Exit Sub
End If

If CmbStatusFlag.Text = "" Then
   MsgBox "Please Select Status !!!"
   CmbStatusFlag.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_DocumentType.RecordCount > 0 Then
      RsMst_DocumentType.MoveFirst
      RsMst_DocumentType.Find "DocumentType= '" & TxtDocumentType.Text & "'"
      If Not RsMst_DocumentType.EOF Then
        MsgBox "Duplicate Document Type Not Allowed !!! "
         TxtDocumentType.SetFocus
         Exit Sub
      End If
   End If
      RsMst_DocumentType.AddNew
      RsMst_DocumentType!DocumentTypeID = GetDocumentTypeID
      RsMst_DocumentType!G_UID = GetGUID
      TxtDocumentTypeID.Text = RsMst_DocumentType!DocumentTypeID
Else
    RsMst_DocumentType.MoveFirst
    RsMst_DocumentType.Find "DocumentType= '" & TxtDocumentType.Text & "'"
    If Not RsMst_DocumentType.EOF Then
       If RsMst_DocumentType!DocumentTypeID <> TxtDocumentTypeID.Text Then
          MsgBox "Duplicate Document Type Not Allowed !!! "
          TxtDocumentType.SetFocus
          Exit Sub
       End If
    End If
    RsMst_DocumentType.MoveFirst
    RsMst_DocumentType.Find " DocumentTypeID= " & TxtDocumentTypeID.Text
End If

RsMst_DocumentType!DocumentTypeID = TxtDocumentTypeID.Text
RsMst_DocumentType!DocumentType = Trim(TxtDocumentType.Text)
RsMst_DocumentType!Flag = Left$(CmbStatusFlag.Text, 1)

If IsNull(RsMst_DocumentType!CreatedBy) = True Or RsMst_DocumentType!CreatedBy = "" Then
   RsMst_DocumentType!CreatedBy = Gen_UserLoginID
   RsMst_DocumentType!CreatedDate = Now
Else
   RsMst_DocumentType!UpdatedBy = Gen_UserLoginID
   RsMst_DocumentType!UpdatedDate = Now
End If


RsMst_DocumentType.Update
Call InData
End Sub
Private Function GetDocumentTypeID() As Double
Dim Retval As Double
tmp = "Select max(DocumentTypeID) from Mst_DocumentType"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetDocumentTypeID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.TextMatrix(0, 0) = "Document Type ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Document Type"
MSHFlexGrid1.TextMatrix(0, 2) = "Status"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 4000
MSHFlexGrid1.ColWidth(2) = 2000


End Sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_DocumentType.MoveFirst
RsMst_DocumentType.Find "DocumentTypeID = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsMst_DocumentType.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub

Private Sub CmbStatusFlag_LostFocus()
If CmbStatusFlag.Text = "" Then Exit Sub

If LCase(CmbStatusFlag.Text) = "active" Or LCase(CmbStatusFlag.Text) = "deactive" Then
Else
   MsgBox "Invalid selection!!!"
   CmbStatusFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmbSFlag_LostFocus()
If CmbSFlag.Text = "" Then Exit Sub

If LCase(CmbSFlag.Text) = "active" Or LCase(CmbSFlag.Text) = "deactive" Then
Else
   MsgBox "Invalid selection!!!"
   CmbSFlag.SetFocus
   Exit Sub
End If

End Sub




