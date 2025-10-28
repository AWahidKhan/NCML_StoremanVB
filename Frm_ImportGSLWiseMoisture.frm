VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frm_ImportGSLWiseMoisture 
   Caption         =   "Import GSL Wise Moisture"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1890
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbGodown 
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
         Left            =   12833
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   380
         Width           =   2085
      End
      Begin VB.CommandButton CmdGetData 
         Caption         =   "Get Data"
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
         Left            =   12833
         TabIndex        =   28
         Top             =   1050
         Width           =   1140
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
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
         Left            =   14000
         TabIndex        =   27
         Top             =   1050
         Width           =   920
      End
      Begin VB.TextBox TxtClientName 
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
         Left            =   8080
         TabIndex        =   25
         Top             =   1050
         Width           =   4680
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
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
         Left            =   10515
         TabIndex        =   22
         Top             =   1450
         Width           =   1500
      End
      Begin VB.CommandButton CmdValidate 
         Caption         =   "Validation"
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
         Left            =   6015
         TabIndex        =   7
         Top             =   1450
         Width           =   1500
      End
      Begin VB.ComboBox CmbCommodity 
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
         Left            =   4050
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1050
         Width           =   4000
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
         Left            =   4050
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   380
         Width           =   4000
      End
      Begin VB.ComboBox CmbStackNo 
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
         ItemData        =   "Frm_ImportGSLWiseMoisture.frx":0000
         Left            =   100
         List            =   "Frm_ImportGSLWiseMoisture.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1050
         Width           =   3915
      End
      Begin VB.ComboBox CmbCMP 
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
         Left            =   8080
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   380
         Width           =   4725
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
         Left            =   105
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   375
         Width           =   3915
      End
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "Get File"
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
         Left            =   3000
         TabIndex        =   5
         Top             =   1450
         Width           =   1500
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
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
         Left            =   7515
         TabIndex        =   8
         Top             =   1450
         Width           =   1500
      End
      Begin VB.CommandButton CmdReset 
         Caption         =   "Reset"
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
         Left            =   9015
         TabIndex        =   9
         Top             =   1450
         Width           =   1500
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
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
         Left            =   4500
         TabIndex        =   6
         Top             =   1450
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Client Name"
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
         Left            =   9873
         TabIndex        =   26
         Top             =   795
         Width           =   1095
      End
      Begin VB.Label lblRecords 
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
         Left            =   13920
         TabIndex        =   24
         Top             =   1500
         Width           =   90
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Records : "
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
         Left            =   12240
         TabIndex        =   23
         Top             =   1500
         Width           =   1680
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   13507
         TabIndex        =   21
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   10185
         TabIndex        =   20
         Top             =   120
         Width           =   435
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Stack No"
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
         Left            =   1645
         TabIndex        =   19
         Top             =   800
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Commodity"
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
         Left            =   5540
         TabIndex        =   18
         Top             =   795
         Width           =   1020
      End
      Begin VB.Label Label3 
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
         Left            =   5588
         TabIndex        =   17
         Top             =   120
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
         Left            =   1770
         TabIndex        =   16
         Top             =   120
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   120
      TabIndex        =   12
      Top             =   1875
      Width           =   15015
      Begin VB.Label lblFilePath 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         Top             =   150
         Width           =   13695
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1020
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6735
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   11880
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   270
      Left            =   120
      TabIndex        =   11
      Top             =   9255
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "Frm_ImportGSLWiseMoisture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mCMPID, mGodownID, mCommodityID, mClientID As Double
Dim mStackID As String, mFileName As String, mFileNamePath As String
Dim mClientName As Variant
Dim I, j, k As Double
Dim mImportYN, mValidation As Boolean

Private Sub CmbState_GotFocus()
Me.MousePointer = vbHourglass

    tmp = CmbState.Text
    Call TableMst_State
    CmbState.Clear
    If RsMst_State.RecordCount = 0 Then
        Me.MousePointer = vbDefault
        'MsgBox "State not found", vbOKOnly + vbCritical, App.Title
        CmbState.SetFocus
        Exit Sub
    End If
    For I = 1 To RsMst_State.RecordCount
        CmbState.AddItem RsMst_State!StateName
        RsMst_State.MoveNext
    Next
    CmbState.Text = tmp


Me.MousePointer = vbDefault
End Sub
Private Sub CmbState_LostFocus()
Me.MousePointer = vbHourglass
    If CmbState.Text = "" Then
       Me.MousePointer = vbDefault
       CmbLocation.Text = ""
       CmbCMP.Text = ""
       CmbGodown.Text = ""
       CmbStackNo.Text = ""
       Exit Sub
    End If
    RsMst_State.MoveFirst
    RsMst_State.Find "StateName = '" & CmbState.Text & "'"
    If RsMst_State.EOF Then
       Me.MousePointer = vbDefault
       MsgBox "Invalid Selection !!!!"
       CmbState.SetFocus
       Exit Sub
    End If
    mStateID = RsMst_State!StateID
Me.MousePointer = vbDefault
End Sub

Private Sub CmbLocation_GotFocus()
Me.MousePointer = vbHourglass

    If CmbState.Text = "" Then
        Me.MousePointer = vbDefault
        MsgBox "Select State.", vbOKOnly + vbInformation, App.Title
        CmbState.SetFocus
        Exit Sub
    End If

    tmp = CmbLocation.Text
    Call TableMst_Location("Where StateID = " & mStateID & " ")
    CmbLocation.Clear
    If RsMst_Location.RecordCount = 0 Then
        Me.MousePointer = vbDefault
        'MsgBox "Location not found", vbOKOnly + vbCritical, App.Title
        CmbLocation.SetFocus
        Exit Sub
    End If
    For I = 1 To RsMst_Location.RecordCount
        CmbLocation.AddItem RsMst_Location!LocationName
        RsMst_Location.MoveNext
    Next
    CmbLocation.Text = tmp


Me.MousePointer = vbDefault
End Sub
Private Sub CmbLocation_LostFocus()
Me.MousePointer = vbHourglass
    If CmbLocation.Text = "" Then
       Me.MousePointer = vbDefault
       CmbCMP.Text = ""
       CmbGodown.Text = ""
       CmbStackNo.Text = ""
       Exit Sub
    End If
    RsMst_Location.MoveFirst
    RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
    If RsMst_Location.EOF Then
       Me.MousePointer = vbDefault
       MsgBox "Invalid Selection !!!!"
       CmbLocation.SetFocus
       Exit Sub
    End If
    mLocationID = RsMst_Location!LocationID
Me.MousePointer = vbDefault
End Sub

Private Sub CmbCMP_GotFocus()
Me.MousePointer = vbHourglass

    If CmbLocation.Text = "" Then
        Me.MousePointer = vbDefault
        MsgBox "Select Location.", vbOKOnly + vbInformation, App.Title
        CmbLocation.SetFocus
        Exit Sub
    End If

    tmp = CmbCMP.Text
    Call TableMst_CMP("Where CloseDate IS NULL AND LocationID = " & mLocationID & " ")
    CmbCMP.Clear
    If RsMst_CMP.RecordCount = 0 Then
        Me.MousePointer = vbDefault
        ''MsgBox "CMP not found", vbOKOnly + vbCritical, App.Title
        CmbCMP.SetFocus
        Exit Sub
    End If
    For I = 1 To RsMst_CMP.RecordCount
        CmbCMP.AddItem RsMst_CMP!CMPName
        RsMst_CMP.MoveNext
    Next
    CmbCMP.Text = tmp


Me.MousePointer = vbDefault
End Sub
Private Sub CmbCMP_LostFocus()
Me.MousePointer = vbHourglass
    If CmbCMP.Text = "" Then
       Me.MousePointer = vbDefault
       CmbGodown.Text = ""
       CmbStackNo.Text = ""
       Exit Sub
    End If
    RsMst_CMP.MoveFirst
    RsMst_CMP.Find "CMPName = '" & CmbCMP.Text & "'"
    If RsMst_CMP.EOF Then
       Me.MousePointer = vbDefault
       MsgBox "Invalid Selection !!!!"
       CmbCMP.SetFocus
       Exit Sub
    End If
    mCMPID = RsMst_CMP!CMPID
Me.MousePointer = vbDefault
End Sub

Private Sub CmbGodown_GotFocus()
Me.MousePointer = vbHourglass

    If CmbCMP.Text = "" Then
        Me.MousePointer = vbDefault
        MsgBox "Select CMP.", vbOKOnly + vbInformation, App.Title
        CmbCMP.SetFocus
        Exit Sub
    End If

    tmp = CmbGodown.Text
    Call TableMst_Godown("Where CloseDate IS NULL AND CMPID = " & mCMPID & " ")
    CmbGodown.Clear
    If RsMst_Godown.RecordCount = 0 Then
        Me.MousePointer = vbDefault
        ''MsgBox "Godown not found", vbOKOnly + vbCritical, App.Title
        CmbGodown.SetFocus
        Exit Sub
    End If
    For I = 1 To RsMst_Godown.RecordCount
        CmbGodown.AddItem RsMst_Godown!GodownNo
        RsMst_Godown.MoveNext
    Next
    CmbGodown.Text = tmp


Me.MousePointer = vbDefault
End Sub
Private Sub CmbGodown_LostFocus()
Me.MousePointer = vbHourglass
    If CmbGodown.Text = "" Then
       Me.MousePointer = vbDefault
       CmbStackNo.Text = ""
       Exit Sub
    End If
    RsMst_Godown.MoveFirst
    RsMst_Godown.Find "GodownNo = '" & CmbGodown.Text & "'"
    If RsMst_Godown.EOF Then
       Me.MousePointer = vbDefault
       MsgBox "Invalid Selection !!!!"
       CmbGodown.SetFocus
       Exit Sub
    End If
    mGodownID = RsMst_Godown!GodownID
Me.MousePointer = vbDefault
End Sub

Private Sub CmbStackNo_GotFocus()
Me.MousePointer = vbHourglass

    If CmbGodown.Text = "" Then
        Me.MousePointer = vbDefault
        MsgBox "Select Godown.", vbOKOnly + vbInformation, App.Title
        CmbGodown.SetFocus
        Exit Sub
    End If

    tmp = CmbStackNo.Text
    Call TableMst_GSL("Where (KRNo Is Null OR KRNo = 0) AND MadeupFlag = 0 AND SpillageFlag = 0 AND CMPID = " & mCMPID & " AND GodownID = " & mGodownID & "", "D")
    CmbStackNo.Clear
    If RsMst_GSL.RecordCount = 0 Then
        Me.MousePointer = vbDefault
        ''MsgBox "Stack not found", vbOKOnly + vbCritical, App.Title
        CmbStackNo.SetFocus
        Exit Sub
    Else
    For I = 1 To RsMst_GSL.RecordCount
        CmbStackNo.AddItem RsMst_GSL!StackNo
        RsMst_GSL.MoveNext
    Next
    CmbStackNo.Text = tmp
    End If

Me.MousePointer = vbDefault
End Sub
Private Sub CmbStackNo_LostFocus()
Me.MousePointer = vbHourglass
    If CmbStackNo.Text = "" Then
       Me.MousePointer = vbDefault
       Exit Sub
    End If
    RsMst_GSL.MoveFirst
    RsMst_GSL.Find "StackNo = '" & CmbStackNo.Text & "'"
    If RsMst_GSL.EOF Then
       Me.MousePointer = vbDefault
       MsgBox "Invalid Selection !!!!"
       CmbStackNo.SetFocus
       Exit Sub
    End If
    mStackID = RsMst_GSL!StackNo
Me.MousePointer = vbDefault
End Sub

Private Sub CmbCommodity_GotFocus()
Me.MousePointer = vbHourglass
    tmp = CmbCommodity.Text
    Call TableMst_Commodity(" Where MAWFlag = 1 ")
    CmbCommodity.Clear
    If RsMst_Commodity.RecordCount = 0 Then
       Me.MousePointer = vbDefault
       ''MsgBox "Commodity not found", vbOKOnly + vbCritical, App.Title
       Exit Sub
    End If
    For I = 1 To RsMst_Commodity.RecordCount
     CmbCommodity.AddItem RsMst_Commodity!Commodity
     RsMst_Commodity.MoveNext
    Next
    CmbCommodity.Text = tmp
Me.MousePointer = vbDefault
End Sub

Private Sub CmbCommodity_LostFocus()
Me.MousePointer = vbHourglass
    If CmbCommodity.Text = "" Then
       Me.MousePointer = vbDefault
       Exit Sub
    End If
    RsMst_Commodity.MoveFirst
    RsMst_Commodity.Find "Commodity = '" & CmbCommodity.Text & "'"
    If RsMst_Commodity.EOF Then
       Me.MousePointer = vbDefault
       MsgBox "Invalid selection.", vbOKOnly + vbInformation, App.Title
       CmbCommodity.SetFocus
       Exit Sub
    End If
    mCommodityID = RsMst_Commodity!CommodityID
Me.MousePointer = vbDefault
End Sub

Private Sub CmdExcel_Click()
    Gen_mTimeStamp = GetTimeStamp
    Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_ImportGSLWiseMoisture.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_ImportGSLWiseMoisture.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

    With MSHFlexGrid1
    ProgressBar1.Value = 0
    ProgressBar1.Max = .Rows - 1
        For I = 0 To .Rows - 1
            mRow = I + 1 ' + 6
            For C = 0 To .Cols - 1
                If InStr(.TextMatrix(I, C), "/") > 0 Then
                   If IsDate(.TextMatrix(I, C)) Then
                      oWS.Cells(mRow, C + 1) = Format(.TextMatrix(I, C), "MM/dd/yyyy")
                   Else
                      oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
                   End If
                Else
                   oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
                End If
                .row = I
                .Col = C
                
                If .CellBackColor = vbRed Then
                   oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbRed
                ElseIf .CellBackColor = vbBlue Then
                   oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbBlue
                ElseIf .CellBackColor = vbYellow Then
                   oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbYellow
                End If
            Next
            ProgressBar1.Value = I
        Next
    End With

oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdGetData_Click()
Dim wc As String
wc = ""

Call Grid_Head
lblRecords.Caption = "-"

If CmbState <> "" Then
   wc = " Where MS.StateID = " & mStateID & " "
End If

If CmbLocation <> "" Then
    If wc = "" Then
        wc = " Where ML.LocationID = " & mLocationID & " "
    Else
        wc = wc & " AND ML.LocationID = " & mLocationID & " "
    End If
End If

If CmbCMP <> "" Then
    If wc = "" Then
        wc = " Where Mst_CMP.CMPID = " & mCMPID & " "
    Else
        wc = wc & " AND Mst_CMP.CMPID = " & mCMPID & " "
    End If
End If

If CmbGodown <> "" Then
   If wc = "" Then
    wc = " Where Mst_Godown.GodownID = " & mGodownID & " "
   Else
    wc = wc & " AND Mst_Godown.GodownID = " & mGodownID & " "
   End If
End If

If CmbStackNo <> "" Then
   If wc = "" Then
    wc = " Where Mst_GSL.StackNo = '" & mStackID & "'"
   Else
    wc = wc & " AND Mst_GSL.StackNo = '" & mStackID & "' "
   End If
End If

If CmbCommodity <> "" Then
   If wc = "" Then
    wc = " Where Mst_Commodity.CommodityID = " & mCommodityID & " "
   Else
    wc = wc & " AND Mst_Commodity.CommodityID = " & mCommodityID & " "
   End If
End If

If Trim(TxtClientName) <> "" Then
   If wc = "" Then
    wc = " Where Mst_Client.ClientName like '%" & Replace(TxtClientName, "'", "`") & "%'"
   Else
    wc = wc & " AND Mst_Client.ClientName like '%" & Replace(TxtClientName, "'", "`") & "%'"
   End If
End If

'If wc = "" Then
'   wc = " Where Mst_Commodity.MAWFlag = 1 AND (Mst_GSL.KRNo Is Null OR Mst_GSL.KRNo = 0) AND Mst_GSL.ExchangeTypeID = 2 AND Mst_GSL.Flag not in ('E','Z','B','J') And Mst_GSL.MadeupFlag = '0' "
'Else
'   wc = wc & " AND Mst_Commodity.MAWFlag = 1 AND (Mst_GSL.KRNo Is Null OR Mst_GSL.KRNo = 0) AND Mst_GSL.ExchangeTypeID = 2 AND Mst_GSL.Flag not in ('E','Z','B','J') And Mst_GSL.MadeupFlag = '0'"
'End If

If wc = "" Then
   wc = " Where Mst_Commodity.MAWFlag = 1 AND (Mst_GSL.KRNo Is Null OR Mst_GSL.KRNo = 0) AND Mst_GSL.ExchangeTypeID = 2 AND Mst_GSL.Status not in ('E','Z','B','J') And Mst_GSL.Status Not in ('B','E','Z') And Mst_GSL.State Not in ('J') And Mst_GSL.MadeupFlag = '0' "
Else
   wc = wc & " AND Mst_Commodity.MAWFlag = 1 AND (Mst_GSL.KRNo Is Null OR Mst_GSL.KRNo = 0) AND Mst_GSL.ExchangeTypeID = 2 AND Mst_GSL.Status not in ('E','Z','B','J') And Mst_GSL.Status Not in ('B','E','Z') And Mst_GSL.State Not in ('J') And Mst_GSL.MadeupFlag = '0'"
End If


tmp = "SELECT MS.StateName,ML.LocationName,Mst_CMP.CMPName,Mst_Godown.GodownNo,Mst_GSL.StackNo,Mst_GSL.LotNo,Mst_Commodity.Commodity,Mst_GSL.Mst_GSL_ID," '' Mst_GSL.GSLID
tmp = tmp & " ML.SM_Prefix,Mst_Client.ClientName,Mst_GSL.GrossWt,Mst_GSL.DepositeWeight,Mst_GSL.MoistureValue,''[MAW%],Mst_GSL.MAWWt From Mst_GSL"
tmp = tmp & " left JOIN Mst_CMP on Mst_CMP.CMPID = Mst_GSL.CMPID"
tmp = tmp & " left join Mst_Location ML on Mst_CMP.LocationID=ML.LocationID"
tmp = tmp & " left join Mst_State MS on MS.StateID=ML.StateID"
tmp = tmp & " left JOIN Mst_Godown on Mst_Godown.GodownID = Mst_GSL.GodownID"
tmp = tmp & " left JOIN Mst_Commodity on Mst_Commodity.CommodityID = Mst_GSL.CommodityID"
''tmp = tmp & " left JOIN Mst_Client on Mst_Client.ClientIDRow = Mst_GSL.ClientIDRow"
tmp = tmp & " left JOIN Mst_Client on Mst_Client.ClientID = Mst_GSL.ClientID"
tmp = tmp & wc & " ORDER BY MS.StateName, ML.LocationName,Mst_CMP.CMPName,Mst_GSL.GSLID "

Call TmpTable
If TmpRs.RecordCount > 0 Then
    ProgressBar1.Value = 0
    ProgressBar1.Max = TmpRs.RecordCount
    With MSHFlexGrid1
        For I = 1 To TmpRs.RecordCount
            For j = 0 To TmpRs.Fields.Count - 1
                .TextMatrix(I, j) = IIf(IsNull(TmpRs(j).Value), "", TmpRs(j).Value)
            Next j
            TmpRs.MoveNext
            .Refresh
            .Rows = .Rows + 1
            ProgressBar1.Value = I
        Next I
    End With
End If
lblRecords.Caption = TmpRs.RecordCount
MsgBox "Done.", vbOKOnly + vbInformation, App.Title
End Sub

Private Sub CmdGetFile_Click()
On Error GoTo Error

With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "Excel files|*.xls;*.xlsx"
     .CancelError = False
     .ShowOpen
End With
mFileNamePath = CommonDialog1.FileName
lblFilePath.Caption = mFileNamePath

If FileExists(lblFilePath.Caption) = True Then
    CmdGetData.Enabled = False
    CmdGetFile.Enabled = False
    CmdImport.Enabled = True
    CmdValidate.Enabled = False
    CmdSave.Enabled = False
    CmdReset.Enabled = True
    CmdExit.Enabled = True
Else
    MsgBox "File not found.", vbOKOnly + vbInformation, App.Title
    Exit Sub
End If
  
Error:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation, "Error"
    End If
End Sub

Private Function FileExists(FileName As String) As Boolean
On Error GoTo ErrorHandler
' get the attributes and ensure that it isn't a directory
FileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
' if an error occurs, this function returns False
End Function

Private Sub CmdImport_Click()
   mImportYN = FillData("Sheet1", MSHFlexGrid1)
   
   lblRecords.Caption = MSHFlexGrid1.Rows - 2
   
   If mImportYN = True Then
    CmdGetData.Enabled = False
    CmdGetFile.Enabled = False
    CmdImport.Enabled = False
    CmdValidate.Enabled = True
    CmdSave.Enabled = False
    CmdExit.Enabled = True
   End If
   
End Sub

Private Sub CmdReset_Click()
    Call Grid_Head
    CmbState = ""
    CmbLocation = ""
    CmbCMP = ""
    CmbGodown = ""
    CmbStackNo = ""
    CmbCommodity = ""
    TxtClientName = ""
    lblFilePath = ""
    lblRecords = "-"
    
    CmdGetData.Enabled = True
    CmdExcel.Enabled = True
    CmdGetFile.Enabled = True
    CmdImport.Enabled = False
    CmdValidate.Enabled = False
    CmdSave.Enabled = False
    CmdExit.Enabled = True
    
End Sub

Private Sub CmdSave_Click()
    With MSHFlexGrid1
    
        ''Call TableMst_GSL(" Where (KRNo Is Null OR KRNo = 0) AND (ExchangeTypeID = 2) AND (Flag<>'E' And Flag<>'B'And Flag<>'Z') AND CommodityID in (SELECT CommodityID From Mst_Commodity Where MAWFlag = 1) ")
        ''tmp = " SELECT GSLID,SM_Prefix,Lower(Convert(VARCHAR(10),GSLID)+'~'+SM_Prefix)[GSLIDSM_Prefix],CommodityID,MoistureValue,MAWWt,CreatedBy,CreatedDate,UpdatedDate,UpdatedBy From Mst_GSL"
        ''tmp = tmp & " Where (KRNo Is Null OR KRNo = 0) AND (ExchangeTypeID = 2) AND (Flag<>'E' And Flag<>'B'And Flag<>'Z')"
        
        tmp = " SELECT GSLID,Mst_GSL_ID,SM_Prefix,Lower(Convert(VARCHAR(10),Mst_GSL_ID)+'~'+SM_Prefix)[GSLIDSM_Prefix],CommodityID,MoistureValue,MAWWt,CreatedBy,CreatedDate,UpdatedDate,UpdatedBy From Mst_GSL"
        tmp = tmp & " Where (KRNo Is Null OR KRNo = 0) AND (ExchangeTypeID = 2) AND (Status<>'E' And Status<>'B'And Status<>'Z') AND (Status<>'E' And Status<>'B'And Status<>'Z')"
        tmp = tmp & " AND CommodityID in (SELECT CommodityID From Mst_Commodity Where MAWFlag = 1) Order By GSLID"
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
            ProgressBar1.Value = 0
            For I = 1 To .Rows - 2
                
                ProgressBar1.Max = .Rows - 2
                ProgressBar1.Value = I
                
                ''-- Value + MAWD + Final Wt
                If Trim(.TextMatrix(I, 12)) <> "" And Trim(.TextMatrix(I, 13)) <> "" And Trim(.TextMatrix(I, 14)) <> "" Then
                
                    '''---------------
                    TmpRs.MoveFirst
                    TmpRs.Find ("GSLIDSM_Prefix = '" & LCase(Trim(.TextMatrix(I, 7)) & "~" & Trim(.TextMatrix(I, 8))) & "'")
                    If TmpRs.EOF = False Then
                        TmpRs!MoistureValue = Val(.TextMatrix(I, 12))
                        TmpRs!MAWWt = Val(.TextMatrix(I, 14))
                        If IsNull(TmpRs!CreatedBy) = True Or TmpRs!CreatedBy = "" Then
                           TmpRs!CreatedBy = Gen_UserLoginID
                           TmpRs!CreatedDate = Now
                        Else
                           TmpRs!UpdatedBy = Gen_UserLoginID
                           TmpRs!UpdatedDate = Now
                        End If
                        TmpRs.Update
                        TmpRs.Requery
                    End If
                    '''---------------
                
                End If
                
            Next I
                MsgBox "Done"
                CmdSave.Enabled = False
        End If
    End With
End Sub

Private Sub CmdValidate_Click()
Dim mGSLID, mSMPrefix As Variant
Dim mMax, mMin, mValue As Variant
Dim mNetW, mMAWD, mMAWW As Double

    ''Call TableMst_Commodity(" Where MAWFlag = 1 ")
    Call TableMst_MAW(" Where Flag = 'A' ")

    If mImportYN = True Then
       With MSHFlexGrid1
            
           .Cols = .Cols + 1
           .TextMatrix(0, .Cols - 1) = "CommodityID"
           .ColWidth(.Cols - 1) = 0
           
           .Cols = .Cols + 1
           .TextMatrix(0, .Cols - 1) = "Error"
           .ColWidth(.Cols - 1) = 3000
            
           ProgressBar1.Value = 0
           ProgressBar1.Max = .Rows - 2
            
            ''tmp = " SELECT Mst_GSL.GSLID,Mst_GSL.SM_Prefix,Lower(Convert(VARCHAR(10),GSLID)+'~'+SM_Prefix)[GSLIDSM_Prefix],Mst_GSL.CommodityID,Mst_GSL.GrossWt,Mst_GSL.DepositeWeight "
            ''tmp = tmp & " FROM Mst_GSL Where Mst_GSL.GSLID = " & Val(.TextMatrix(I, 7)) & " AND Mst_GSL.SM_Prefix = '" & Trim(.TextMatrix(I, 8)) & "' "
            
            tmp = " SELECT Mst_GSL.GSLID,Mst_GSL.Mst_GSL_ID,Mst_GSL.SM_Prefix,Lower(Convert(VARCHAR(10),Mst_GSL_ID)+'~'+SM_Prefix)[GSLIDSM_Prefix],Mst_GSL.CommodityID,Mst_GSL.GrossWt,Mst_GSL.DepositeWeight "
            tmp = tmp & " FROM Mst_GSL Where (Mst_GSL.KRNo Is Null OR Mst_GSL.KRNo = 0) "
            tmp = tmp & " AND (Mst_GSL.ExchangeTypeID = 2) AND Mst_GSL.MadeupFlag = 0 AND Mst_GSL.SpillageFlag = 0 "
            tmp = tmp & " AND (Mst_GSL.Status<>'E' And Mst_GSL.Status<>'B'And Mst_GSL.Status<>'Z') "
            tmp = tmp & " AND (Mst_GSL.Status<>'E' And Mst_GSL.Status<>'B'And Mst_GSL.Status<>'Z') "
            tmp = tmp & " AND CommodityID in (SELECT CommodityID From Mst_Commodity Where MAWFlag = 1) Order By GSLID "
            Call Tmp2Table
           
           For I = 1 To .Rows - 2
            ProgressBar1.Value = I
           
            If TmpRs2.RecordCount > 0 Then
                TmpRs2.MoveFirst
                TmpRs2.Find ("GSLIDSM_Prefix = '" & LCase(Trim(.TextMatrix(I, 7)) & "~" & Trim(.TextMatrix(I, 8))) & "'")
                If TmpRs2.EOF = False Then
                
                    .TextMatrix(I, 10) = IIf(IsNull(TmpRs2!GrossWt), "", TmpRs2!GrossWt)
                    .TextMatrix(I, 11) = IIf(IsNull(TmpRs2!DepositeWeight), "", TmpRs2!DepositeWeight)
                    mNetW = .TextMatrix(I, 11)  'Val(.TextMatrix(I, 11)) 'Commeneted on 16 Nov 2015 after updating excel of change request of s&P
                    .TextMatrix(I, 15) = IIf(IsNull(TmpRs2!CommodityID), "", TmpRs2!CommodityID)
                
                Else
                    Call FlexCellColor(MSHFlexGrid1, CDbl(I), 7, "Invalid record GSLID - " & .TextMatrix(I, 7) & " And SM_Prefix - " & .TextMatrix(I, 8) & "")
                End If
            Else
                Call FlexCellColor(MSHFlexGrid1, CDbl(I), 7, "Invalid record GSLID - " & .TextMatrix(I, 7) & " And SM_Prefix - " & .TextMatrix(I, 8) & "")
            End If
           
            ''-------
              If IsNumeric(Trim(.TextMatrix(I, 11))) = True And Trim(.TextMatrix(I, 15)) <> "" Then
                 If RsMst_MAW.RecordCount > 0 Then
                    RsMst_MAW.MoveFirst
                    RsMst_MAW.Find "CommodityID = " & .TextMatrix(I, 15)
                    If RsMst_MAW.EOF = False Then
                      mMax = IIf(IsNull(RsMst_MAW!MaximumRange), 0, RsMst_MAW!MaximumRange)
                      mMin = IIf(IsNull(RsMst_MAW!MinimumRange), 0, RsMst_MAW!MinimumRange)
                      mValue = IIf(Trim(.TextMatrix(I, 12)) = "", "", Val(.TextMatrix(I, 12)))
                      
                      If mValue <> "" Then
                        If mValue >= mMin And mValue <= mMax Then
                            
                            tmp = " SELECT MAWD FROM Mst_MAWDetails Where MAWID = " & Val(RsMst_MAW!MAWID) & " AND (MinimumMoisture<= " & mValue & " AND MaximumMoisture >= " & mValue & ")"
                            Call TmpTable
                            If TmpRs.RecordCount > 0 Then
                               .TextMatrix(I, 13) = IIf(IsNull(TmpRs!MAWD), 0, TmpRs!MAWD)
                               mMAWD = Val(.TextMatrix(I, 13))
                               .TextMatrix(I, 14) = Round(mNetW - ((mNetW * mMAWD) / 100), 3)
                               
                            Else
                               Call FlexCellColor(MSHFlexGrid1, CDbl(I), 12, "Moisture Value out of range.")
                            End If
                        Else
                            Call FlexCellColor(MSHFlexGrid1, CDbl(I), 12, "Moisture Value out of range.")
                        End If
                      Else
                        Call FlexCellColor(MSHFlexGrid1, CDbl(I), 12, "Moisture Value not given.")
                      End If

                    End If
                 End If
              Else
                If IsNumeric(Trim(.TextMatrix(I, 11))) = False Then Call FlexCellColor(MSHFlexGrid1, CDbl(I), 11, "Net Weight is not available.")
                If Trim(.TextMatrix(I, 15)) = "" = False Then Call FlexCellColor(MSHFlexGrid1, CDbl(I), 15, "Invalid Commodity.")
              End If
            ''-------
           
           Next I
           
           k = 0
           For I = 1 To .Rows - 2
               If Trim(.TextMatrix(I, 14)) <> "" Then
                  k = k + 1
               End If
           Next I
           
           If k = .Rows Then
                MsgBox "All data is not valid, try again with correct data.", vbOKOnly + vbInformation, App.Title
                CmdValidate.Enabled = False
                CmdSave.Enabled = False
           ElseIf k < .Rows Then
                MsgBox "Validation Done.", vbOKOnly + vbInformation, App.Title
                CmdValidate.Enabled = False
                CmdSave.Enabled = True
           End If
           
            CmdGetData.Enabled = False
            CmdExcel.Enabled = True
            CmdGetFile.Enabled = False
            CmdImport.Enabled = False
            CmdExit.Enabled = True
           
       End With
        
    End If
    
End Sub

Private Sub FlexCellColor(Grid As MSHFlexGrid, row As Double, Col As Double, ErrorMsg As Variant)
On Error GoTo Error

    With Grid
        .row = row
        .Col = Col
        .CellBackColor = vbRed
        .row = row
        .Col = Grid.Cols - 1
        .CellBackColor = vbRed
        .TextMatrix(row, .Cols - 1) = Trim(.TextMatrix(row, .Cols - 1)) & ";" & ErrorMsg
        ''TxtValidationError.Text = TxtValidationError.Text & "Row:" & row & " Column:" & Col + 1 & " -->>" & ErrorMsg & vbNewLine
        ''TxtValidationError.Refresh
        mValidation = False
        Me.Refresh
    End With
    
Error:
    If Err.Number <> 0 Then
        MsgBox "Error Number :" & Err.Number & vbNewLine & "Error Description :" & Err.Description, vbOKOnly + vbExclamation, App.Title
    End If
End Sub



Private Sub Form_Load()
    Dim FrmLoadAccess() As Boolean
    FrmLoadAccess = GetFrmLoadAccess(Me.Name)
    If FrmLoadAccess(0) = False Then
       MsgBox "Access denied !!!!!!!!!"
       Frame1.Enabled = False
       Frame2.Enabled = False
       Exit Sub
    End If
    
    Frame1.Enabled = True
    Frame2.Enabled = True
    
    Call CmdReset_Click
End Sub


Private Sub Grid_Head()
    With MSHFlexGrid1
        .Clear
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 15
        .Rows = 2
        .AllowUserResizing = flexResizeColumns
        .WordWrap = True
        
        .RowHeight(0) = 700
        .TextMatrix(0, 0) = "State":                .ColWidth(0) = 2000
        .TextMatrix(0, 1) = "Location":             .ColWidth(1) = 2000
        .TextMatrix(0, 2) = "CMP Name":             .ColWidth(2) = 2000
        .TextMatrix(0, 3) = "Godown No":            .ColWidth(3) = 1500
        .TextMatrix(0, 4) = "Stack No":             .ColWidth(4) = 1500
        .TextMatrix(0, 5) = "Lot No":               .ColWidth(5) = 1500
        .TextMatrix(0, 6) = "Commodity":            .ColWidth(6) = 2000
        .TextMatrix(0, 7) = "GSL ID":               .ColWidth(7) = 1500
        .TextMatrix(0, 8) = "SM_Prefix":            .ColWidth(8) = 1800
        .TextMatrix(0, 9) = "Client Name":          .ColWidth(9) = 2000
        .TextMatrix(0, 10) = "Gross Wt.":           .ColWidth(10) = 1500
        .TextMatrix(0, 11) = "Net Wt.":             .ColWidth(11) = 1500
        .TextMatrix(0, 12) = "Moisture Value":      .ColWidth(12) = 1500
        .TextMatrix(0, 13) = "Final MAWD in %":     .ColWidth(13) = 1500
        .TextMatrix(0, 14) = "MAW Wt.":             .ColWidth(14) = 1500
    End With
End Sub


Private Function FillData(SheetNumber As Variant, Grid As MSHFlexGrid) As Boolean

On Error GoTo ExcelError
Dim intNumberRows As Double
Dim WorkSheetName As Variant
    With Grid
        Set oXL = New Excel.Application
        Set oWB = oXL.Workbooks.Open(mFileNamePath)
        Set oWS = oWB.Worksheets(SheetNumber)
        oWS.ClearArrows
        
        Call Grid_Head
        
        ''-- Sample Data Header Checking
        For I = 1 To .Cols
            If Trim(LCase(oWS.Cells(1, I))) <> Trim(LCase(.TextMatrix(0, I - 1))) Then
                MsgBox "Excel Caption mismatch!" & vbNewLine & "Invalid Caption Name for column - " & .TextMatrix(0, I - 1) & vbNewLine & "Column No - " & I, vbOKOnly + vbExclamation, App.Title
                oXL.Quit
                Me.MousePointer = vbDefault
                FillData = False
                Exit Function
            End If
        Next I

        mRow = 1
        RowCount = oWB.Worksheets(1).UsedRange.Rows.Count
        ProgressBar1.Value = 0
        ProgressBar1.Max = RowCount
        
        Do
           If oWS.Cells(mRow + 1, 1) = "" Then
                If mRow = 1 Then
                    oXL.Quit
                    Me.MousePointer = vbDefault
                    FillData = False
                    Exit Function
                End If
              Exit Do
           End If
           For C = 0 To .Cols - 1
               .TextMatrix(mRow, C) = Trim(oWS.Cells(mRow + 1, C + 1).Value)
           Next
           .Rows = .Rows + 1
           .Refresh
           ProgressBar1.Value = ProgressBar1.Value + 1
           mRow = mRow + 1
        Loop Until False
        
        ProgressBar1.Value = ProgressBar1.Max
        oXL.Quit
        MsgBox "Data Imported Successfully ", vbInformation, App.Title
        
        FillData = True
        
    End With
ExcelError:
    If Err.Number > 0 Then
        MsgBox "Error Number : " & Err.Number & vbNewLine & "Error Description : " & Err.Description, vbOKOnly + vbCritical, App.Title
        oXL.Quit
        FillData = False
        Me.MousePointer = Default
    End If
End Function

