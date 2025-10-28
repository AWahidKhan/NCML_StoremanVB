VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_CRDepositeDetails 
   Caption         =   "Cash/Cheque Receipt Deposit Details"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1125
      Left            =   3135
      TabIndex        =   29
      Top             =   4440
      Visible         =   0   'False
      Width           =   7440
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
         Height          =   350
         Left            =   13425
         TabIndex        =   50
         Top             =   855
         Width           =   1000
      End
      Begin VB.ComboBox CmbPreFix 
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
         Left            =   3330
         Sorted          =   -1  'True
         TabIndex        =   46
         Top             =   405
         Width           =   3570
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
         Height          =   345
         Left            =   11415
         TabIndex        =   32
         Top             =   855
         Width           =   1005
      End
      Begin VB.ComboBox CmbSBank 
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
         Left            =   7020
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   405
         Width           =   7845
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
         Left            =   12420
         TabIndex        =   30
         Top             =   855
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1215
         Left            =   75
         TabIndex        =   33
         Top             =   1365
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2143
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   345
         Left            =   1680
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   405
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   609
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
         Format          =   51707905
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   345
         Left            =   90
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   405
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
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
         Format          =   51707905
         CurrentDate     =   36494
      End
      Begin VB.Label Label36 
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
         Left            =   4718
         TabIndex        =   49
         Top             =   120
         Width           =   795
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
         Left            =   270
         TabIndex        =   48
         Top             =   120
         Width           =   1095
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
         Left            =   2085
         TabIndex        =   47
         Top             =   120
         Width           =   855
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
         Left            =   5520
         TabIndex        =   36
         Top             =   930
         Width           =   2355
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
         Left            =   7920
         TabIndex        =   35
         Top             =   930
         Width           =   90
      End
      Begin VB.Label Label11 
         Caption         =   "Bank"
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
         Left            =   10687
         TabIndex        =   34
         Top             =   165
         Width           =   510
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cash/Cheque Receipts not yet deposited"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5685
      Left            =   120
      TabIndex        =   28
      Top             =   2895
      Width           =   14970
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4020
         TabIndex        =   42
         Top             =   300
         Width           =   495
      End
      Begin VB.CheckBox ChkCheckAll 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4845
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   14715
         _ExtentX        =   25956
         _ExtentY        =   8546
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label6 
         Caption         =   "Cash/Cheque Receipt not yet printed"
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
         Left            =   4590
         TabIndex        =   43
         Top             =   315
         Width           =   3600
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   11
      Top             =   8520
      Width           =   14970
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
         Left            =   255
         TabIndex        =   21
         Top             =   615
         Width           =   2865
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
         Left            =   3120
         TabIndex        =   20
         Top             =   615
         Width           =   2865
      End
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
         Left            =   11715
         TabIndex        =   19
         Top             =   615
         Width           =   2865
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
         Left            =   11715
         TabIndex        =   18
         Top             =   195
         Width           =   2865
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
         Left            =   8850
         TabIndex        =   17
         Top             =   615
         Width           =   2865
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
         Left            =   8850
         TabIndex        =   16
         Top             =   195
         Width           =   2865
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
         Left            =   5985
         TabIndex        =   15
         Top             =   615
         Width           =   2865
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
         Left            =   5985
         TabIndex        =   14
         Top             =   195
         Width           =   2865
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
         Left            =   3120
         TabIndex        =   13
         Top             =   195
         Width           =   2865
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
         Left            =   240
         TabIndex        =   12
         Top             =   195
         Width           =   2865
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   14985
      Begin VB.CommandButton CmdFAGBank 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   14205
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1155
         Width           =   375
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
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   12030
      End
      Begin VB.TextBox TxtPreFix 
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
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   240
         Width           =   1635
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2355
         Width           =   5205
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
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2355
         Width           =   5430
      End
      Begin VB.TextBox TxtCRDepositeID 
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
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   255
         Width           =   1455
      End
      Begin VB.ComboBox CmbBankName 
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
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1125
         Width           =   12030
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
         Height          =   615
         Left            =   2040
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1605
         Width           =   12030
      End
      Begin VB.TextBox TxtBankName 
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
         Left            =   2040
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1125
         Width           =   12030
      End
      Begin MSComCtl2.DTPicker TxtDepositeDate 
         Height          =   375
         Left            =   6030
         TabIndex        =   0
         Top             =   255
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         Format          =   51707907
         CurrentDate     =   40254
      End
      Begin VB.TextBox TxtLocationID 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   720
         Width           =   12030
      End
      Begin VB.Label Label8 
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
         Height          =   300
         Left            =   120
         TabIndex        =   40
         Top             =   750
         Width           =   1440
      End
      Begin VB.Label Label5 
         Caption         =   "Prefix"
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
         Left            =   12090
         TabIndex        =   39
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2415
         Width           =   1815
      End
      Begin VB.Label Label2 
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
         Left            =   7680
         TabIndex        =   26
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "CR Deposit ID"
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
         TabIndex        =   22
         Top             =   255
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Bank Name"
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
         TabIndex        =   10
         Top             =   1155
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   1755
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Date of Deposit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4245
         TabIndex        =   7
         Top             =   300
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FrmTxn_CRDepositeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mFAGBankID, mSFAGBankID As Variant
Dim mLocationID As Double
Dim Edit_Flg As Variant
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Frame1.Enabled = True
TxtDepositeDate.Value = Date
TxtCRDepositeID = ""


CmbLocationID.Text = ""
TxtLocationID.Visible = False
CmbLocationID.Visible = True

TxtPreFix = ""

TxtBankName.Text = ""
CmbBankName.Text = ""
TxtBankName.Visible = False
CmbBankName.Visible = True
CmdFAGBank.Visible = True

TxtRemark = ""
TxtCreated = ""
TxtUpdated = ""
ChkCheckAll.Value = 0
Call Grid_Head
If RsTxn_CRDepositDetails.RecordCount > 0 Then
   TxtDepositeDate.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub ChkCheckAll_Click()
If SaveCmd.Enabled = False Then
   ChkCheckAll.Value = 0
   Exit Sub
End If
If MSHFlexGrid1.Rows = 3 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 2 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       If MSHFlexGrid1.CellBackColor <> vbRed Then
          If MSHFlexGrid1.TextMatrix(I, 0) <> strChecked Then
             MSHFlexGrid1.TextMatrix(I, 0) = strChecked
             MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 8))
          End If
       End If
   Next
Else
   For I = 2 To MSHFlexGrid1.Rows - 2
       MSHFlexGrid1.row = I
       MSHFlexGrid1.Col = 0
       If MSHFlexGrid1.CellBackColor <> vbRed Then
          MSHFlexGrid1.TextMatrix(I, 0) = strUnChecked
          MSHFlexGrid1.TextMatrix(0, 8) = 0 'Val(MSHFlexGrid1.TextMatrix(0, 8)) - Val(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 8))
       End If
   Next
End If
End Sub
Private Sub CmbSBank_GotFocus()
tmp = CmbSBank.Text
Call TableMst_FAGBank("")
CmbSBank.Clear
If RsMst_FAGBank.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbSBank.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_FAGBank.RecordCount
    CmbSBank.AddItem RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
    RsMst_FAGBank.MoveNext
Next
CmbSBank.Text = tmp

End Sub

Private Sub CmbSBank_LostFocus()

Dim strSplitString() As String
Dim AccountNo_, BankName_ As String

If CmbSBank.Text = "" Then
   mSFAGBankID = 0
   Exit Sub
End If



strSplitString = Split(CmbSBank, "~")

If UBound(strSplitString) < 1 Then
   MsgBox "Invalid Selection!!!"
   CmbSBank.SetFocus
   Exit Sub
End If

BankName_ = Trim(strSplitString(0))
AccountNo_ = Trim(strSplitString(1))

RsMst_FAGBank.MoveFirst
RsMst_FAGBank.Find "AccountNo = '" & AccountNo_ & "'"

If RsMst_FAGBank.EOF Then
   MsgBox "Invalid FAG Bank selection "
   CmbSBank.SetFocus
   Exit Sub
End If

mSFAGBankID = RsMst_FAGBank!FAGBankID
End Sub

Private Sub CmdClear_Click()
SFrom.Value = Date - 365
STo.Value = Date
CmbPreFix.Text = ""
CmbSBank.Text = ""
Label22.Caption = ""
Call Grid_Head1
End Sub

Private Sub CmdFAGBank_Click()
FrmMst_FAGBank.Show
End Sub

Private Sub CmdGo_Click()
Call Grid_Head1
If SFrom.Value > STo.Value Then
   MsgBox "'From Date' must be less than 'To Date' always!!!"
   SFrom.SetFocus
   Exit Sub
End If


tmp = "Select TCRD.CRDepositID,TCRD.DepositDate,ML.LocationName,MFAGBank.BankName + '~' + MFAGBank.AccountNo as 'Bank',TCRD.Remark " & _
      " From Txn_CashReceiptDepositDetails TCRD  " & _
      " Inner Join Mst_FAGBank MFAGBank On TCRD.FAGBankID = MFAGBank.FAGBankID " & _
      " Inner Join Mst_Location ML On TCRD.LocationID = ML.LocationID "

Dim wc As Variant

wc = ""

If Gen_DBType = "MDB" Then
   wc = " Where TCRD.DepositDate Between #" & SFrom.Value & "# And #" & STo.Value & "#"
Else
   wc = " Where TCRD.DepositDate Between '" & SFrom.Value & "' And '" & STo.Value & "'"
End If

If CmbPreFix.Text <> "" Then
   wc = wc & " And ML.LocationName = '" & CmbPreFix.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc = " And ML." & Gen_WcLocation
   End If
End If

If CmbSBank.Text <> "" Then
   If wc <> "" Then
      wc = wc & " And MFAGBank.FAGBankID = '" & mSFAGBankID & "'"
   Else
      wc = " Where MFAGBank.FAGBankID = '" & mSFAGBankID & "'"
   End If
End If
tmp = tmp & wc & " Order by CRDepositID"
Call TmpTable
Call Grid_Head1


For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
    For C = 0 To MSHFlexGrid2.Cols - 1
        MSHFlexGrid2.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    MSHFlexGrid2.TextMatrix(iu, 1) = Format(MSHFlexGrid2.TextMatrix(iu, 1), "dd-mmm-yyyy")
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then

   Call DeleteCRDetails
   
   RsTxn_CRDepositDetails.Delete
   RsTxn_CRDepositDetails.Update
   
   If RsTxn_CRDepositDetails.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_CRDepositDetails.MoveNext
   If RsTxn_CRDepositDetails.EOF() Then
      RsTxn_CRDepositDetails.MoveLast
   End If
   Call Indata
   
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_ContractType.CancelUpdate

End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
Frame1.Enabled = True

CmbLocationID.Text = TxtLocationID
CmbLocationID.Visible = True
TxtLocationID.Visible = False

CmbBankName.Text = TxtBankName.Text
CmbBankName.Visible = True
TxtBankName.Visible = False
CmdFAGBank.Visible = True
Call FillCRDetails(mLocationID, Val(TxtCRDepositeID))

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True

If Gen_WcLocation <> "" Then
   Call TableTxn_CRDepositDetails(" Where  " & Gen_WcLocation & " Order by CRDepositID ")
Else
   Call TableTxn_CRDepositDetails(" Order by CRDepositID ")
End If
If RsTxn_CRDepositDetails.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)
End Sub
Public Sub Indata()
Frame1.Enabled = False

CmbLocationID.Text = ""
TxtLocationID.Text = ""
CmbLocationID.Visible = False
TxtLocationID.Visible = True


CmbBankName.Text = ""
TxtBankName.Text = ""
CmbBankName.Visible = False
TxtBankName.Visible = True



TxtCRDepositeID = RsTxn_CRDepositDetails!CRDepositID
TxtDepositeDate.Value = RsTxn_CRDepositDetails!DepositDate

mLocationID = RsTxn_CRDepositDetails!LocationID
TxtLocationID = GetLocationName(mLocationID)
TxtPreFix = GetPreFixByLocation(mLocationID, "P")

TxtRemark.Text = RsTxn_CRDepositDetails!Remark
mFAGBankID = RsTxn_CRDepositDetails!FAGBankID
CmdFAGBank.Visible = False

tmp = "Select MFB.BankName + '~' + MFB.AccountNo as 'Bank' From Txn_CashReceiptDepositDetails TCRD"
tmp = tmp & " Inner Join Mst_FAGBank MFB On TCRD.FAGBankID = MFB.FAGBankID "
tmp = tmp & " Where MFB.FAGBankID = " & mFAGBankID & ""

Call TmpTable

'TxtBankName.Text = GetFAGBankName(RsTxn_CRDepositDetails!FAGBankId, "N")

If TmpRs.RecordCount > 0 Then
   TxtBankName.Text = TmpRs!Bank
Else
   TxtBankName.Text = ""
End If

TxtCreated = IIf(IsNull(RsTxn_CRDepositDetails!CreatedBy), "", RsTxn_CRDepositDetails!CreatedBy) & " :- " & IIf(IsNull(RsTxn_CRDepositDetails!CreatedDate), "", RsTxn_CRDepositDetails!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_CRDepositDetails!UpdatedBy), "", RsTxn_CRDepositDetails!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_CRDepositDetails!UpdatedDate), "", RsTxn_CRDepositDetails!UpdatedDate)

Call InDataCRDetails(Val(TxtCRDepositeID))

Call GButtonLock(Me, True)
End Sub

Private Sub CmbBankName_GotFocus()
tmp = CmbBankName.Text
Call TableMst_FAGBank("")
CmbBankName.Clear
If RsMst_FAGBank.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbBankName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_FAGBank.RecordCount
    CmbBankName.AddItem RsMst_FAGBank!BankName & "~" & RsMst_FAGBank!AccountNo
    RsMst_FAGBank.MoveNext
Next
CmbBankName.Text = tmp
End Sub
Private Sub CmbBankName_LostFocus()
Dim strSplitString() As String
Dim AccountNo_, BankName_ As String

If CmbBankName.Text = "" Then
   Exit Sub
End If

strSplitString = Split(CmbBankName, "~")

If UBound(strSplitString) <> 1 Then
   MsgBox "Invalid Selection!!!"
   CmbBankName.SetFocus
   Exit Sub
End If

BankName_ = Trim(strSplitString(0))
AccountNo_ = Trim(strSplitString(1))

RsMst_FAGBank.MoveFirst
RsMst_FAGBank.Find "AccountNo = '" & AccountNo_ & "'"

If RsMst_FAGBank.EOF Then
   MsgBox "Invalid FAG Bank selection "
   CmbBankName.SetFocus
   Exit Sub
End If
mFAGBankID = RsMst_FAGBank!FAGBankID

End Sub

Private Sub InDataCRDetails(mLocationID, Optional CRDepositeID_ As Double)

tmp = "Select TCR.CRID,TCR.SM_Prefix,ML.LocationName,TCR.CRDate,TCR.InstrumentType,TCR.InstrumentNo, TCR.DrawnOn,TCR.InstrumentAmount, "
tmp = tmp & " MA.AgentName ,TCR.Flag ,TCR.CRDepositID "
tmp = tmp & " From Txn_CashReceipt TCR "
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID "
tmp = tmp & " Inner Join Mst_Agent MA On TCR.AgentID = MA.AgentID "
tmp = tmp & " Where ML.LocationName = '" & TxtLocationID & "' And CRDepositID = " & Val(TxtCRDepositeID) & " "


Call TmpTable

Call Grid_Head
If TmpRs.RecordCount > 0 Then
   x = 2
   For I = 1 To TmpRs.RecordCount
       For C = 0 To TmpRs.Fields.Count - 3
           MSHFlexGrid1.TextMatrix(x, C + 1) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid1.TextMatrix(x, 4) = Format(MSHFlexGrid1.TextMatrix(x, 4), "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(x, 8) = Format(Val(MSHFlexGrid1.TextMatrix(x, 8)), "#####0.00")
       MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(x, 8))
       MSHFlexGrid1.row = x
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid1.TextMatrix(x, 0) = strChecked
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       x = x + 1
       TmpRs.MoveNext
   Next
End If
End Sub
Private Sub FillCRDetails(mLocationID_ As Double, Optional CRDepositeID_ As Double)

tmp = " Select TCR.CRID,TCR.SM_Prefix,ML.LocationName,TCR.CRDate,TCR.InstrumentType,TCR.InstrumentNo, TCR.DrawnOn,TCR.InstrumentAmount,"
tmp = tmp & " MA.AgentName , TCR.Flag ,TCR.CRDepositID"
tmp = tmp & " From Txn_CashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Agent MA On TCR.AgentID = MA.AgentID "
tmp = tmp & " Where TCR.SM_prefix = '" & TxtPreFix & "' And (CRDepositID is Null Or CRDepositID = 0 "

If CRDepositeID_ <> 0 Then
   tmp = tmp & " Or CRDepositID= " & Val(TxtCRDepositeID) & ")"
Else
   tmp = tmp & ") "
End If

tmp = tmp & " Order By TCR.SM_Prefix,TCR.CRID"
Call TmpTable

Call Grid_Head
ChkCheckAll.Value = 0
If TmpRs.RecordCount > 0 Then
   x = 2
   For I = 1 To TmpRs.RecordCount
       For C = 0 To TmpRs.Fields.Count - 3
           MSHFlexGrid1.TextMatrix(x, C + 1) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid1.TextMatrix(x, 4) = Format(MSHFlexGrid1.TextMatrix(x, 4), "dd-MMM-yyyy")
       MSHFlexGrid1.TextMatrix(x, 8) = Format(Val(MSHFlexGrid1.TextMatrix(x, 8)), "#####0.00")
       
       MSHFlexGrid1.row = x
       MSHFlexGrid1.Col = 0
       MSHFlexGrid1.CellFontName = "Wingdings"
       MSHFlexGrid1.CellFontSize = 12
       MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
       MSHFlexGrid1.TextMatrix(x, 0) = strUnChecked
       If CRDepositeID_ <> 0 Then
          If CRDepositeID_ = IIf(IsNull(TmpRs!CRDepositID), 0, TmpRs!CRDepositID) Then
             MSHFlexGrid1.TextMatrix(x, 0) = strChecked
             MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(x, 8))
          End If
       End If
       If LCase(TmpRs!Flag) <> "p" Or IsNull(TmpRs!Flag) = True Then
          'MSHFlexGrid1.Row = i
          For o = 0 To MSHFlexGrid1.Cols - 1
              MSHFlexGrid1.Col = o
              MSHFlexGrid1.CellBackColor = vbRed
          Next
       End If
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       x = x + 1
       TmpRs.MoveNext
   Next
Else
   MsgBox "No Cash Receipt record found!!!"
End If
End Sub
Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .Cols = 10
    .Font.Size = 10
    .FixedCols = 0
    .FixedRows = 2
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(1, 0) = "Select"
    .TextMatrix(1, 1) = "CR ID"
    .TextMatrix(1, 2) = "SM_Prefix"
    .TextMatrix(1, 3) = "Location"
    .TextMatrix(1, 4) = "CR Date"
    .TextMatrix(1, 5) = "Type"
    .TextMatrix(1, 6) = "Cheque No"
    .TextMatrix(1, 7) = "Drawn On"
    .TextMatrix(1, 8) = "Amount"
    .TextMatrix(1, 9) = "Agent Name"

    .ColAlignment(5) = vbLeftJustify
    .ColWidth(0) = 800
    .ColWidth(1) = 800
    .ColWidth(2) = 0
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 3200
    .ColWidth(8) = 1200
    .ColWidth(9) = 3400
End With
End Sub
Public Sub Grid_Head1()
With MSHFlexGrid2
    .Clear
    .Rows = 2
    .Cols = 5
    .Font.Size = 10
    .FixedCols = 0
    .FixedRows = 1
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Cash Receipt Deposit ID"
    .TextMatrix(0, 1) = "Deposit Date"
    .TextMatrix(0, 2) = "Location Name"
    .TextMatrix(0, 3) = "Bank Name"
    .TextMatrix(0, 4) = "Remark"
    .ColWidth(0) = 1600
    .ColWidth(1) = 1400
    .ColWidth(2) = 3400
    .ColWidth(3) = 4300
    .ColWidth(4) = 3000
    
    .WordWrap = True
    .RowHeight(0) = 500
    
End With
End Sub
Private Sub MSHFlexGrid1_Click()
If SaveCmd.Enabled = False Then Exit Sub
With MSHFlexGrid1
    If .TextMatrix(.row, 2) = "" Then Exit Sub
    If .Col = 0 Then
        If (.CellBackColor) <> vbRed Then
           If .TextMatrix(.row, 0) = strUnChecked Then
              .TextMatrix(.row, 0) = strChecked
              .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) + Val(.TextMatrix(.row, 8))
           Else
              .TextMatrix(.row, 0) = strUnChecked
              .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) - Val(.TextMatrix(.row, 8))
           End If
        End If
    End If
End With
End Sub

Private Sub MSHFlexGrid2_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid2.RowSel

If MSHFlexGrid2.TextMatrix(x, 0) = "" Then Exit Sub
mCaId = MSHFlexGrid2.TextMatrix(x, 0)

RsTxn_CRDepositDetails.MoveFirst

If RsTxn_CRDepositDetails.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsTxn_CRDepositDetails.MoveFirst
RsTxn_CRDepositDetails.Find "CRDepositID = " & mCaId & ""
If RsTxn_CRDepositDetails.EOF Then
   RsTxn_CRDepositDetails.MoveLast
End If
Frame0.Visible = False
Call Indata
End Sub

Private Sub SaveCmd_Click()
If SaveCmd.Enabled = False Then
   Exit Sub
End If

If CmbBankName = "" Then
   MsgBox "Blank Bank Name not Allowed"
   CmbBankName.SetFocus
   Exit Sub
End If

If TxtRemark.Text = "" Then
   MsgBox "Blank Remark not Allowed"
   TxtRemark.SetFocus
   Exit Sub
End If

If MSHFlexGrid1.Rows < 4 Then
   MsgBox "Blank Cash Receipt Details not Allowed"
   Exit Sub
End If

If CheckCRSelect = False Then
   MsgBox "Select atleast one Cash Receipt ID from CR Details!!!"
   Exit Sub
End If

If Edit_Flg = "A" Then
    RsTxn_CRDepositDetails.AddNew
    RsTxn_CRDepositDetails!CRDepositID = GetMaxCRDepositID
    RsTxn_CRDepositDetails!G_UID = GetGUID
    TxtCRDepositeID = RsTxn_CRDepositDetails!CRDepositID
Else
   RsTxn_CRDepositDetails.MoveFirst
   RsTxn_CRDepositDetails.Find "CRDepositID = " & TxtCRDepositeID & ""
   Call DeleteCRDetails
End If
RsTxn_CRDepositDetails!LocationID = mLocationID
RsTxn_CRDepositDetails!FAGBankID = mFAGBankID
RsTxn_CRDepositDetails!DepositDate = TxtDepositeDate.Value
RsTxn_CRDepositDetails!Remark = Trim(TxtRemark.Text)

If IsNull(RsTxn_CRDepositDetails!CreatedBy) = True Or RsTxn_CRDepositDetails!CreatedBy = "" Then
   RsTxn_CRDepositDetails!CreatedBy = Gen_UserLoginID
   RsTxn_CRDepositDetails!CreatedDate = Now
Else
   RsTxn_CRDepositDetails!UpdatedBy = Gen_UserLoginID
   RsTxn_CRDepositDetails!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If

RsTxn_CRDepositDetails.Update
RsTxn_CRDepositDetails.Requery

Call UpdateCRDetails
RsTxn_CRDepositDetails.MoveFirst
RsTxn_CRDepositDetails.Find "CRDepositID = " & TxtCRDepositeID & ""

Call GButtonLock(Me, True)
Call Indata
End Sub
Private Sub DeleteCRDetails()
tmp = " Update Txn_CashReceipt Set CRDepositID = 0 Where CRDepositID=" & Val(TxtCRDepositeID)
Call TmpTable
End Sub

Private Sub UpdateCRDetails()
For I = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       tmp = " Update Txn_CashReceipt Set CRDepositID = " & Val(TxtCRDepositeID) & " Where CRID = " & Val(MSHFlexGrid1.TextMatrix(I, 1)) & " And SM_Prefix= '" & Trim(MSHFlexGrid1.TextMatrix(I, 2)) & "'"
       Call TmpTable
    End If
Next
End Sub
Private Function GetMaxCRDepositID() As Double
Dim Retval As Double
tmp = "Select max(CRDepositID) from Txn_CashReceiptDepositDetails"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxCRDepositID = Retval
End Function
Private Function CheckCRSelect() As Boolean
Dim Retval As Boolean
Retval = False
For I = 1 To MSHFlexGrid1.Rows - 2
    If MSHFlexGrid1.TextMatrix(I, 0) = strChecked Then
       Retval = True
       Exit For
    End If
Next
CheckCRSelect = Retval
End Function

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CRDepositDetails.MoveNext
If RsTxn_CRDepositDetails.EOF Then
   RsTxn_CRDepositDetails.MoveLast
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
RsTxn_CRDepositDetails.MovePrevious
If RsTxn_CRDepositDetails.BOF Then
   RsTxn_CRDepositDetails.MoveFirst
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
Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_CRDepositDetails.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True

End Sub
Private Sub LastCmd_Click()
RsTxn_CRDepositDetails.MoveLast
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsTxn_CRDepositDetails.EOF Then
      RsTxn_CRDepositDetails.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top '120
Frame0.Width = Me.Width - 350
Frame0.Visible = True

MSHFlexGrid2.Width = Frame0.Width - 200
MSHFlexGrid2.Height = Frame0.Height - 1400

RsTxn_CRDepositDetails.MoveFirst
Call CmdClear_Click

End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

'Call TableMst_Location(" Where LocationId > 0 ")

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

If tmp <> CmbLocationID.Text Then
   Call Grid_Head
End If

mLocationID = RsMst_Location!LocationID
TxtPreFix = GetPreFixByLocation(RsMst_Location!LocationID, "P")
Call FillCRDetails(mLocationID)

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_CashReceiptDepositDetails.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_CashReceiptDepositDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 1 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If


CmbPreFix.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbPreFix.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbPreFix.Text = tmp
End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbPreFix.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbPreFix.SetFocus
   Exit Sub
End If

End Sub

