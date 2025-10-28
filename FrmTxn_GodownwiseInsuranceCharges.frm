VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_GodownwiseInsuranceCharges 
   Caption         =   "Monthly Insurance Charges Godown wise for P&L MIS"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   10095
      Left            =   15000
      TabIndex        =   42
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   9885
      Width           =   10935
   End
   Begin VB.Frame Frame4 
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14940
      Begin VB.Frame Frame0 
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   9600
         Visible         =   0   'False
         Width           =   13905
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
            Height          =   345
            Left            =   11595
            TabIndex        =   4
            Top             =   540
            Width           =   1125
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
            Left            =   12720
            TabIndex        =   3
            Top             =   540
            Width           =   1125
         End
         Begin VB.TextBox TxtSClaimID 
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
            Left            =   3120
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   480
            Width           =   2835
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1455
            Left            =   90
            TabIndex        =   5
            Top             =   1020
            Width           =   13755
            _ExtentX        =   24262
            _ExtentY        =   2566
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker STo 
            Height          =   360
            Left            =   1620
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   480
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   635
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
            Format          =   118358017
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker SFrom 
            Height          =   360
            Left            =   60
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   480
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   118358017
            CurrentDate     =   36494
         End
         Begin VB.Label Label44 
            Caption         =   "From Date"
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
            Left            =   240
            TabIndex        =   12
            Top             =   225
            Width           =   990
         End
         Begin VB.Label Label45 
            Caption         =   "To Date"
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
            Left            =   1913
            TabIndex        =   11
            Top             =   225
            Width           =   795
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
            Left            =   10200
            TabIndex        =   10
            Top             =   600
            Width           =   90
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
            Left            =   7800
            TabIndex        =   9
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Label19 
            Caption         =   "Txn ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4237
            TabIndex        =   8
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2205
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   13710
         Begin VB.TextBox TxtTxnID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   195
            Width           =   1455
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
            Height          =   915
            Left            =   1845
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   1170
            Width           =   9585
         End
         Begin VB.TextBox TxtG_UID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Enabled         =   0   'False
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
            Left            =   12660
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   2040
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.TextBox TxtInsuranceAmount 
            Alignment       =   1  'Right Justify
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
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   735
            Width           =   4185
         End
         Begin VB.TextBox TxtActualAmount 
            Alignment       =   1  'Right Justify
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
            Left            =   7245
            MaxLength       =   18
            TabIndex        =   25
            Top             =   720
            Width           =   4155
         End
         Begin MSComCtl2.DTPicker TxtDate 
            Height          =   375
            Left            =   4560
            TabIndex        =   30
            Top             =   195
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   118358017
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtExpenseFrom 
            Height          =   375
            Left            =   7245
            TabIndex        =   31
            Top             =   195
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   118358017
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker TxtExpenseTo 
            Height          =   375
            Left            =   9960
            TabIndex        =   32
            Top             =   195
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   118358017
            CurrentDate     =   39884
         End
         Begin VB.Label LbClaimID 
            Caption         =   "Txn ID"
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
            Left            =   120
            TabIndex        =   40
            Top             =   277
            Width           =   900
         End
         Begin VB.Label Label41 
            Caption         =   "Txn Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   3480
            TabIndex        =   39
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Remarks"
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
            Left            =   120
            TabIndex        =   38
            Top             =   1500
            Width           =   1080
         End
         Begin VB.Label Label12 
            Caption         =   "G_UID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   11640
            TabIndex        =   37
            Top             =   2085
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label15 
            Caption         =   "Insurance Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   36
            Top             =   780
            Width           =   1740
         End
         Begin VB.Label Label17 
            Caption         =   "From Date "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6240
            TabIndex        =   35
            Top             =   247
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "To Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   9000
            TabIndex        =   34
            Top             =   240
            Width           =   765
         End
         Begin VB.Label Label9 
            Caption         =   "Actual Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6240
            TabIndex        =   33
            Top             =   660
            Width           =   990
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1125
         Left            =   120
         TabIndex        =   13
         Top             =   8550
         Width           =   13770
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   10905
            TabIndex        =   23
            Top             =   600
            Width           =   2700
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   10905
            TabIndex        =   22
            Top             =   195
            Width           =   2700
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   8205
            TabIndex        =   21
            Top             =   600
            Width           =   2700
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   8205
            TabIndex        =   20
            Top             =   195
            Width           =   2700
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   5505
            TabIndex        =   19
            Top             =   600
            Width           =   2700
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   5505
            TabIndex        =   18
            Top             =   195
            Width           =   2700
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   2805
            TabIndex        =   17
            Top             =   600
            Width           =   2700
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   2805
            TabIndex        =   16
            Top             =   195
            Width           =   2700
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   105
            TabIndex        =   15
            Top             =   600
            Width           =   2700
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   105
            TabIndex        =   14
            Top             =   195
            Width           =   2700
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6135
         Left            =   120
         TabIndex        =   43
         Top             =   2400
         Width           =   13755
         _ExtentX        =   24262
         _ExtentY        =   10821
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Godown Details"
         TabPicture(0)   =   "FrmTxn_GodownwiseInsuranceCharges.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "CmdExcelDetail"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Log History"
         TabPicture(1)   =   "FrmTxn_GodownwiseInsuranceCharges.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label25"
         Tab(1).Control(1)=   "Label24"
         Tab(1).Control(2)=   "TxtUpdated"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "TxtCreated"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         Begin VB.CommandButton CmdExcelDetail 
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
            Height          =   345
            Left            =   12600
            TabIndex        =   49
            Top             =   405
            Width           =   1000
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
            Left            =   -73080
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   720
            Width           =   5535
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
            Left            =   -73080
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1215
            Width           =   5535
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   5175
            Left            =   60
            TabIndex        =   44
            Top             =   810
            Width           =   13590
            _ExtentX        =   23971
            _ExtentY        =   9128
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
         Begin VB.Label Label24 
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
            Height          =   240
            Left            =   -74880
            TabIndex        =   48
            Top             =   780
            Width           =   1695
         End
         Begin VB.Label Label25 
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
            Height          =   255
            Left            =   -74880
            TabIndex        =   47
            Top             =   1275
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_GodownwiseInsuranceCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Private Sub CmdExcelDetail_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_MonthlyInsuranceChargesDetail.xls")
Call Xls_Detail_Rpt(MSHFlexGrid2)
End Sub


Private Sub Xls_Detail_Rpt(MSH_Grid As MSHFlexGrid)
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_MonthlyInsuranceChargesDetail.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSH_Grid.Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To MSH_Grid.Cols - 1
        If InStr(MSH_Grid.TextMatrix(i, C), "/") > 0 Then
           If IsDate(MSH_Grid.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSH_Grid.TextMatrix(i, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSH_Grid.TextMatrix(i, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSH_Grid.TextMatrix(i, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub



Private Sub EditCmd_Click()

Frame1.Enabled = True
TxtActualAmount.Locked = False
TxtRemark.Locked = False
TxtActualAmount.SetFocus
Call GButtonLock(Me, False)


End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim wc As Variant

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head

Frame1.Enabled = True

TxtDate.Value = Date
TxtExpenseFrom.Value = Year(Date)
mTotalDays = Day(DateSerial(Year(TxtDate.Value), Month(TxtDate.Value) + 1, 0))
TxtExpenseFrom.Value = Format(Year(TxtDate.Value) & "-" & Month(TxtDate.Value) & "-" & "01", "yyyy-mm-dd")
TxtExpenseTo.Value = Format(Year(TxtDate.Value) & "-" & Month(TxtDate.Value) & "-" & mTotalDays, "yyyy-mm-dd")

Call TableTxn_MonthlyInsuranceCharges(" Order by TxnID")

If RsTxn_MonthlyInsuranceCharges.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Frame1.Enabled = True
   Call GButtonLockAD(Me)
   Exit Sub
End If
Call GButtonLock(Me, True)
Call Indata

End Sub
'Private Sub TxtExpenseFrom_LostFocus()
'TxtExpenseFrom.Value = GetLastDateOfMonth(TxtExpenseFrom.Value)
'End Sub
Private Function GetLastDateOfMonth(mDate_ As Date)
Dim mSMonth, mSYear, mSTotalDays As Variant
mSMonth = Month(mDate_)
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
End Function

Public Sub Indata()

Frame1.Enabled = False

TxtTxnID = RsTxn_MonthlyInsuranceCharges!TxnID
TxtDate.Value = RsTxn_MonthlyInsuranceCharges!TxnDate
TxtInsuranceAmount = RsTxn_MonthlyInsuranceCharges!InsuranceAmount
TxtActualAmount = IIf(IsNull(RsTxn_MonthlyInsuranceCharges!ActualAmount), 0, RsTxn_MonthlyInsuranceCharges!ActualAmount)
TxtRemark.Text = IIf(IsNull(RsTxn_MonthlyInsuranceCharges!Remarks), "", RsTxn_MonthlyInsuranceCharges!Remarks)
TxtExpenseFrom.Value = RsTxn_MonthlyInsuranceCharges!YearMonth
TxtExpenseTo.Value = GetLastDateOfMonth(TxtExpenseFrom.Value)

TxtCreated = IIf(IsNull(RsTxn_MonthlyInsuranceCharges!CreatedBy), "", RsTxn_MonthlyInsuranceCharges!CreatedBy) & " :- " & IIf(IsNull(RsTxn_MonthlyInsuranceCharges!CreatedDate), "", RsTxn_MonthlyInsuranceCharges!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_MonthlyInsuranceCharges!UpdatedBy), "", RsTxn_MonthlyInsuranceCharges!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_MonthlyInsuranceCharges!UpdatedDate), "", RsTxn_MonthlyInsuranceCharges!UpdatedDate)

Call Grid_Head


tmp = " Select TMICD.DetailID,ML.LocationName,MC.CMPName,MG.GodownNo,MG.GodownID,MU.UnitType,MCom.Commodity,"
tmp = tmp & " TMICD.InsuranceAmount , TMICD.BalanceBags, TMICD.BalanceWeight, TMICD.CommodityRate, TMICD.AUM, TMICD.ActualAmount"
tmp = tmp & " From Txn_MonthlyInsuranceChargesDetail TMICD"
tmp = tmp & " Inner Join Mst_Godown MG On TMICD.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_UnitType MU On MG.UnitTypeID = MU.UnitTypeID"
tmp = tmp & " Inner Join Mst_Commodity MCom On TMICD.CommodityID = MCom.CommodityID"
tmp = tmp & " Where TMICD.TxnID = " & Val(TxtTxnID) & ""
tmp = tmp & " Order By TMICD.DetailID,ML.LocationName,MC.CMPName,MG.GodownNo"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid2.Rows = TmpRs.RecordCount + 3
   x = 2
   For i = 1 To TmpRs.RecordCount
       For C = 0 To MSHFlexGrid2.Cols - 3
           MSHFlexGrid2.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid2.TextMatrix(x, 13) = TxtTxnID
       MSHFlexGrid2.TextMatrix(x, 14) = Format(TxtExpenseFrom.Value, "dd-mmm-yyyy")
       MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(0, 7)) + Val(MSHFlexGrid2.TextMatrix(x, 7)) '"Insurance Amount"
       MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(0, 8)) + Val(MSHFlexGrid2.TextMatrix(x, 8)) '"Insurance Amount"
       MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) + Val(MSHFlexGrid2.TextMatrix(x, 9)) '"Insurance Amount"
       MSHFlexGrid2.TextMatrix(0, 11) = Val(MSHFlexGrid2.TextMatrix(0, 11)) + Val(MSHFlexGrid2.TextMatrix(x, 11)) '"Insurance Amount"
       MSHFlexGrid2.TextMatrix(0, 12) = Val(MSHFlexGrid2.TextMatrix(0, 12)) + Val(MSHFlexGrid2.TextMatrix(x, 12)) '"Insurance Amount"
       x = x + 1
       TmpRs.MoveNext
   Next
End If

If RsTxn_MonthlyInsuranceCharges.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

End Sub

Private Sub Grid_Head()

MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 15
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.WordWrap = True

MSHFlexGrid2.TextMatrix(1, 0) = "Detail ID"
MSHFlexGrid2.TextMatrix(1, 1) = "Location Name"
MSHFlexGrid2.TextMatrix(1, 2) = "CMP Name"
MSHFlexGrid2.TextMatrix(1, 3) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 4) = "Godown ID"
MSHFlexGrid2.TextMatrix(1, 5) = "Godown Type"
MSHFlexGrid2.TextMatrix(1, 6) = "Commodity"
MSHFlexGrid2.TextMatrix(1, 7) = "Insurance Amount"
MSHFlexGrid2.TextMatrix(1, 8) = "Balance Bags"
MSHFlexGrid2.TextMatrix(1, 9) = "Balance Weight"
MSHFlexGrid2.TextMatrix(1, 10) = "Market Price"
MSHFlexGrid2.TextMatrix(1, 11) = "AUM"
MSHFlexGrid2.TextMatrix(1, 12) = "Actual Amount"
MSHFlexGrid2.TextMatrix(1, 13) = "Txn ID"
MSHFlexGrid2.TextMatrix(1, 14) = "Year Month"

MSHFlexGrid2.ColWidth(0) = 1000
MSHFlexGrid2.ColWidth(1) = 1400
MSHFlexGrid2.ColWidth(2) = 1400
MSHFlexGrid2.ColWidth(3) = 1000
MSHFlexGrid2.ColWidth(4) = 1000
MSHFlexGrid2.ColWidth(5) = 1000
MSHFlexGrid2.ColWidth(6) = 1400
MSHFlexGrid2.ColWidth(7) = 1400
MSHFlexGrid2.ColWidth(8) = 1000
MSHFlexGrid2.ColWidth(9) = 1000
MSHFlexGrid2.ColWidth(10) = 1000
MSHFlexGrid2.ColWidth(11) = 1000
MSHFlexGrid2.ColWidth(12) = 1000
MSHFlexGrid2.ColWidth(13) = 1000
MSHFlexGrid2.ColWidth(14) = 1000

MSHFlexGrid2.RowHeight(1) = 500


End Sub

Private Sub SaveCmd_Click()

SaveCmd.Enabled = False

If Val(TxtActualAmount) = 0 Then
   For i = 2 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.TextMatrix(i, 12) = 0
   Next
Else
   MSHFlexGrid2.TextMatrix(0, 12) = 0
'   For i = 2 To MSHFlexGrid2.Rows - 2
'       'MSHFlexGrid2.TextMatrix(i, 11) = IIf(IsNull(TxtActualAmount), 0, Format(TxtActualAmount * Val(MSHFlexGrid2.TextMatrix(i, 6)) / Val(MSHFlexGrid2.TextMatrix(0, 6)), "#############.00"))
'       If Val(MSHFlexGrid2.TextMatrix(i, 8)) > 0 Then
'          MSHFlexGrid2.TextMatrix(i, 12) = IIf(IsNull(TxtActualAmount), 0, Format(TxtActualAmount * Val(MSHFlexGrid2.TextMatrix(i, 11)) / Val(MSHFlexGrid2.TextMatrix(0, 11)), "#############.00"))
'          MSHFlexGrid2.TextMatrix(0, 12) = Val(MSHFlexGrid2.TextMatrix(0, 12)) + Val(MSHFlexGrid2.TextMatrix(i, 12))
'       End If
'   Next

   For i = 2 To MSHFlexGrid2.Rows - 2
       'MSHFlexGrid2.TextMatrix(i, 11) = IIf(IsNull(TxtActualAmount), 0, Format(TxtActualAmount * Val(MSHFlexGrid2.TextMatrix(i, 6)) / Val(MSHFlexGrid2.TextMatrix(0, 6)), "#############.00"))
       If Val(MSHFlexGrid2.TextMatrix(i, 8)) > 0 Then
          MSHFlexGrid2.TextMatrix(i, 12) = IIf(IsNull(TxtActualAmount), 0, Format(TxtActualAmount * Val(MSHFlexGrid2.TextMatrix(i, 7)) / Val(MSHFlexGrid2.TextMatrix(0, 7)), "#############.00"))
          MSHFlexGrid2.TextMatrix(0, 12) = Val(MSHFlexGrid2.TextMatrix(0, 12)) + Val(MSHFlexGrid2.TextMatrix(i, 12))
       End If
   Next

End If

If Val(TxtActualAmount) <> Val(MSHFlexGrid2.TextMatrix(0, 12)) Then
   mDiff = Val(TxtActualAmount) - Val(MSHFlexGrid2.TextMatrix(0, 12))
   mDiff = Format(mDiff, "#######0.00")
   MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 2, 12) = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.Rows - 2, 12)) + Val(mDiff)
   MSHFlexGrid2.TextMatrix(0, 12) = Val(MSHFlexGrid2.TextMatrix(0, 12)) + Val(mDiff)
'   MsgBox "Mismatch in Actual Amount!!!"
'   TxtActualAmount.SetFocus
'   Exit Sub
End If

RsTxn_MonthlyInsuranceCharges.MoveFirst
RsTxn_MonthlyInsuranceCharges.Find "TxnID = " & Val(TxtTxnID) & ""
If RsTxn_MonthlyInsuranceCharges.EOF = False Then
   RsTxn_MonthlyInsuranceCharges!ActualAmount = Val(TxtActualAmount)
   RsTxn_MonthlyInsuranceCharges!Remarks = Trim(TxtRemark)
   RsTxn_MonthlyInsuranceCharges!UpdatedBy = Gen_UserLoginID
   RsTxn_MonthlyInsuranceCharges!UpdatedDate = Now
   RsTxn_MonthlyInsuranceCharges.Update
   RsTxn_MonthlyInsuranceCharges.Requery
End If

Call TableTxn_MonthlyInsuranceChargesDetail(" Where TxnID = " & Val(TxtTxnID) & " ")

For i = 2 To MSHFlexGrid2.Rows - 2
    If RsTxn_MonthlyInsuranceChargesDetail.RecordCount > 0 Then
       If Val(MSHFlexGrid2.TextMatrix(i, 12)) <> 0 Then
          RsTxn_MonthlyInsuranceChargesDetail.MoveFirst
          RsTxn_MonthlyInsuranceChargesDetail.Find "DetailID = " & Val(MSHFlexGrid2.TextMatrix(i, 0)) & ""
          If RsTxn_MonthlyInsuranceChargesDetail.EOF = False Then
             RsTxn_MonthlyInsuranceChargesDetail!ActualAmount = Val(MSHFlexGrid2.TextMatrix(i, 12))
             RsTxn_MonthlyInsuranceChargesDetail!UpdatedBy = Gen_UserLoginID
             RsTxn_MonthlyInsuranceChargesDetail!UpdatedDate = Now
             RsTxn_MonthlyInsuranceChargesDetail.Update
             'RsTxn_MonthlyInsuranceChargesDetail.Requery
          End If
       End If
    End If
Next

RsTxn_MonthlyInsuranceCharges.MoveFirst
RsTxn_MonthlyInsuranceCharges.Find "TxnID = " & Val(TxtTxnID) & ""

Call Indata

End Sub

Private Sub TxtActualAmount_LostFocus()

If TxtActualAmount = "" Then Exit Sub

If IsNumeric(TxtActualAmount) = False Or Val(TxtActualAmount) < 0 Then
   MsgBox "Invalid Format!!!"
   TxtActualAmount.SetFocus
   Exit Sub
End If

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_MonthlyInsuranceCharges.MoveNext
If RsTxn_MonthlyInsuranceCharges.EOF Then
   RsTxn_MonthlyInsuranceCharges.MoveLast
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
RsTxn_MonthlyInsuranceCharges.MovePrevious
If RsTxn_MonthlyInsuranceCharges.BOF Then
   RsTxn_MonthlyInsuranceCharges.MoveFirst
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


Private Sub LastCmd_Click()
RsTxn_MonthlyInsuranceCharges.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub FirstCmd_Click()
RsTxn_MonthlyInsuranceCharges.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

