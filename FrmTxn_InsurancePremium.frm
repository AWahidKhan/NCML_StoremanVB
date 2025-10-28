VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_InsurancePremium 
   Caption         =   "Insurance Premium"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   40
      Top             =   9480
      Visible         =   0   'False
      Width           =   16260
      Begin VB.ComboBox CmbSPolicyAndCoverNoteNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   41
         Top             =   480
         Width           =   10245
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   14790
         TabIndex        =   44
         Top             =   493
         Width           =   1260
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   375
         Left            =   13530
         TabIndex        =   43
         Top             =   493
         Width           =   1260
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   495
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   16095
         _ExtentX        =   28390
         _ExtentY        =   873
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label11 
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
         Left            =   10485
         TabIndex        =   47
         Top             =   553
         Width           =   2295
      End
      Begin VB.Label Label12 
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
         Left            =   12795
         TabIndex        =   46
         Top             =   560
         Width           =   90
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Policy And Cover Note No"
         Height          =   240
         Left            =   4035
         TabIndex        =   45
         Top             =   225
         Width           =   2340
      End
   End
   Begin VB.Frame ButtonFrm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   29
      Top             =   8400
      Width           =   16260
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   12915
         TabIndex        =   39
         Top             =   580
         Width           =   3150
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   9885
         TabIndex        =   38
         Top             =   580
         Width           =   3030
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   6615
         TabIndex        =   37
         Top             =   580
         Width           =   3270
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   3345
         TabIndex        =   36
         Top             =   580
         Width           =   3270
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   75
         TabIndex        =   35
         Top             =   580
         Width           =   3270
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   12915
         TabIndex        =   34
         Top             =   195
         Width           =   3150
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   9885
         TabIndex        =   33
         Top             =   195
         Width           =   3030
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   6615
         TabIndex        =   32
         Top             =   195
         Width           =   3270
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   3345
         TabIndex        =   31
         Top             =   195
         Width           =   3270
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   75
         TabIndex        =   30
         Top             =   195
         Width           =   3270
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8280
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   16260
      Begin TabDlg.SSTab SSTab1 
         Height          =   5775
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   16005
         _ExtentX        =   28231
         _ExtentY        =   10186
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Additional Premium Deatils"
         TabPicture(0)   =   "FrmTxn_InsurancePremium.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Log History"
         TabPicture(1)   =   "FrmTxn_InsurancePremium.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TxtCreated"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "TxtUpdated"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "MSHFlexGrid3"
         Tab(1).Control(3)=   "Label17"
         Tab(1).Control(4)=   "Label16"
         Tab(1).ControlCount=   5
         Begin VB.Frame Frame2 
            Height          =   855
            Left            =   120
            TabIndex        =   58
            Top             =   310
            Width           =   15735
            Begin VB.CommandButton CmdUpdateToList 
               Caption         =   "Update to List"
               Height          =   375
               Left            =   14040
               TabIndex        =   19
               Top             =   420
               Width           =   1620
            End
            Begin VB.TextBox TxtAdditionalPremium 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   120
               TabIndex        =   14
               Top             =   435
               Width           =   2055
            End
            Begin VB.TextBox TxtAdditionalPremiumRemark 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   7800
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   18
               Top             =   435
               Width           =   6135
            End
            Begin VB.TextBox TxtAdditionalPremiumMonthly 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   360
               Left            =   5640
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   435
               Width           =   2040
            End
            Begin MSComCtl2.DTPicker TxtAdditionalPremiumFromDate 
               Height          =   375
               Left            =   2280
               TabIndex        =   15
               Top             =   420
               Width           =   1575
               _ExtentX        =   2778
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
               Format          =   107675649
               CurrentDate     =   39884
            End
            Begin MSComCtl2.DTPicker TxtAdditionalPremiumToDate 
               Height          =   375
               Left            =   3960
               TabIndex        =   16
               Top             =   420
               Width           =   1575
               _ExtentX        =   2778
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
               Format          =   107675649
               CurrentDate     =   39884
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               Caption         =   "Additional Premium"
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
               Left            =   135
               TabIndex        =   63
               Top             =   180
               Width           =   2025
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "From Date"
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
               Left            =   2280
               TabIndex        =   62
               Top             =   180
               Width           =   1575
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "To Date"
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
               Left            =   3960
               TabIndex        =   61
               Top             =   180
               Width           =   1575
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Premium (Monthly)"
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
               Left            =   5640
               TabIndex        =   60
               Top             =   180
               Width           =   2040
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
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
               Left            =   10440
               TabIndex        =   59
               Top             =   180
               Width           =   840
            End
         End
         Begin VB.TextBox TxtCreated 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   -73950
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   360
            Width           =   6750
         End
         Begin VB.TextBox TxtUpdated 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   -65910
            Locked          =   -1  'True
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   360
            Width           =   6750
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   4815
            Left            =   -74880
            TabIndex        =   21
            Top             =   840
            Width           =   15735
            _ExtentX        =   27755
            _ExtentY        =   8493
            _Version        =   393216
            AllowBigSelection=   0   'False
            SelectionMode   =   1
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   4455
            Left            =   120
            TabIndex        =   20
            Top             =   1200
            Width           =   15735
            _ExtentX        =   27755
            _ExtentY        =   7858
            _Version        =   393216
            AllowBigSelection=   0   'False
            SelectionMode   =   1
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Created  :"
            Height          =   240
            Left            =   -74880
            TabIndex        =   57
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Updated :"
            Height          =   240
            Left            =   -66960
            TabIndex        =   56
            Top             =   420
            Width           =   885
         End
      End
      Begin VB.TextBox TxtStructureWiseLimit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   13080
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1960
         Width           =   3000
      End
      Begin VB.TextBox TxtInsuranceAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   7920
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1960
         Width           =   3000
      End
      Begin VB.TextBox TxtInsuranceCompany 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   7920
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1520
         Width           =   8160
      End
      Begin VB.TextBox TxtInsuranceRiskType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   2880
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1520
         Width           =   3000
      End
      Begin VB.TextBox TxtValidFromDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   7920
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3000
      End
      Begin VB.TextBox TxtValidToDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   13080
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3000
      End
      Begin VB.TextBox TxtInsuranceBy 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   2880
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3000
      End
      Begin VB.TextBox TxtInsurancePolicyID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   13080
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   210
         Width           =   3000
      End
      Begin VB.TextBox TxtInsurancePremiumDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   7920
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   3000
      End
      Begin VB.TextBox TxtPolicyPremiumAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2880
         TabIndex        =   10
         Top             =   1960
         Width           =   3015
      End
      Begin VB.ComboBox CmbPolicyAndCoverNoteNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2880
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   650
         Width           =   13245
      End
      Begin VB.TextBox TxtInsurancePremiumID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   2880
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   210
         Width           =   3000
      End
      Begin VB.TextBox TxtPolicyAndCoverNoteNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   645
         Width           =   13215
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Structure Wise Limit"
         Height          =   240
         Left            =   11040
         TabIndex        =   53
         Top             =   2020
         Width           =   1740
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Insurance Amount"
         Height          =   240
         Left            =   6000
         TabIndex        =   52
         Top             =   2020
         Width           =   1590
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Valid To Date"
         Height          =   240
         Left            =   11040
         TabIndex        =   51
         Top             =   1140
         Width           =   1245
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Valid From Date"
         Height          =   240
         Left            =   6000
         TabIndex        =   50
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Insurance Policy ID"
         Height          =   240
         Left            =   11010
         TabIndex        =   49
         Top             =   270
         Width           =   1725
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   240
         Left            =   6030
         TabIndex        =   48
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Policy And Cover Note No"
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
         Left            =   105
         TabIndex        =   28
         Top             =   705
         Width           =   2715
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Insurance Company"
         Height          =   240
         Left            =   6000
         TabIndex        =   27
         Top             =   1580
         Width           =   1785
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Premium (Annually)"
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
         Left            =   150
         TabIndex        =   26
         Top             =   2020
         Width           =   2010
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Insurance Risk Type"
         Height          =   240
         Left            =   120
         TabIndex        =   25
         Top             =   1580
         Width           =   1875
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Insurance By"
         Height          =   240
         Left            =   105
         TabIndex        =   24
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Insurance Premium ID"
         Height          =   240
         Left            =   105
         TabIndex        =   23
         Top             =   270
         Width           =   1965
      End
   End
End
Attribute VB_Name = "FrmTxn_InsurancePremium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mInsurancePolicyID As Double, mSInsurancePolicyID As Double

Private Sub CmbPolicyAndCoverNoteNo_GotFocus()
tmp = CmbPolicyAndCoverNoteNo.Text
Call TableMst_PolicyDetail(" Where Flag = 'A' And InsuranceByID = 1 ")
CmbPolicyAndCoverNoteNo.Clear
If RsMst_PolicyDetail.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_PolicyDetail.RecordCount
    CmbPolicyAndCoverNoteNo.AddItem RsMst_PolicyDetail!PolicyNo & "~" & RsMst_PolicyDetail!CoverNoteNo
    RsMst_PolicyDetail.MoveNext
Next
CmbPolicyAndCoverNoteNo.Text = tmp
End Sub
Private Sub CmbPolicyAndCoverNoteNo_LostFocus()
Dim mSplit As Variant
If CmbPolicyAndCoverNoteNo.Text = "" Then
    Call InsurancePolicyDetails(0)
    Exit Sub
End If
mSplit = Split(CmbPolicyAndCoverNoteNo.Text, "~")
If UBound(mSplit) <> 1 Then
   MsgBox "Invalid selection "
   CmbPolicyAndCoverNoteNo.SetFocus
   Exit Sub
End If
RsMst_PolicyDetail.Filter = ""
RsMst_PolicyDetail.MoveFirst
RsMst_PolicyDetail.Filter = "PolicyNo = '" & mSplit(0) & "' And CoverNoteNo = '" & mSplit(1) & "'  "
If RsMst_PolicyDetail.EOF Then
   MsgBox "Invalid selection "
   CmbPolicyAndCoverNoteNo.SetFocus
   Exit Sub
End If
mInsurancePolicyID = RsMst_PolicyDetail!InsurancePolicyID
Call InsurancePolicyDetails(mInsurancePolicyID)

''-- Check the Premium Plan is already genrate for select policy
tmp = " Select * from Txn_InsurancePremium Where InsurancePolicyID = " & mInsurancePolicyID & ""
Call TmpTable
If TmpRs.RecordCount > 0 Then
    Select Case MsgBox("Insurance Premium is already entered for select policy." & vbNewLine & "Do you want to see the entry?", vbYesNo)
        Case vbYes
            TxtInsurancePremiumID = TmpRs!InsurancePremiumID
            Call Indata
        Case vbNo
            CmbPolicyAndCoverNoteNo.SetFocus
            Exit Sub
    End Select
End If
''--------------------------

If DateDiff("d", Date, CDate(TxtValidToDate)) < 0 Then
   MsgBox "Insurance Policy is Expired."
   Call Grid_Head2
   Exit Sub
End If
If CmbPolicyAndCoverNoteNo.Text = "" Then
   Call Grid_Head2
   Exit Sub
End If
If Val(TxtPolicyPremiumAmt) <= 0 Then
   Call Grid_Head2
   Exit Sub
End If
If DateDiff("d", Date, CDate(TxtValidToDate)) >= 0 Then
   Call MonthlyPremiumPlan
Else
   MsgBox "Insurance Policy is Expired."
   Exit Sub
End If
End Sub

Private Sub CmbSPolicyAndCoverNoteNo_GotFocus()
tmp = CmbSPolicyAndCoverNoteNo.Text
Call TableMst_PolicyDetail(" Where Flag = 'A' And InsuranceByID = 1 ")
CmbSPolicyAndCoverNoteNo.Clear
If RsMst_PolicyDetail.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_PolicyDetail.RecordCount
    CmbSPolicyAndCoverNoteNo.AddItem RsMst_PolicyDetail!PolicyNo & "~" & RsMst_PolicyDetail!CoverNoteNo
    RsMst_PolicyDetail.MoveNext
Next
CmbSPolicyAndCoverNoteNo.Text = tmp
End Sub
Private Sub CmbSPolicyAndCoverNoteNo_LostFocus()
Dim mSplit As Variant
If CmbSPolicyAndCoverNoteNo.Text = "" Then Exit Sub
mSplit = Split(CmbSPolicyAndCoverNoteNo.Text, "~")
If UBound(mSplit) <> 1 Then
   MsgBox "Invalid selection "
   CmbSPolicyAndCoverNoteNo.SetFocus
   Exit Sub
End If
RsMst_PolicyDetail.Filter = ""
RsMst_PolicyDetail.MoveFirst
RsMst_PolicyDetail.Filter = "PolicyNo = '" & mSplit(0) & "' And CoverNoteNo = '" & mSplit(1) & "'  "
If RsMst_PolicyDetail.EOF Then
   MsgBox "Invalid selection "
   CmbSPolicyAndCoverNoteNo.SetFocus
   Exit Sub
End If
mSInsurancePolicyID = RsMst_PolicyDetail!InsurancePolicyID
End Sub

Private Sub CmdUpdateToList_Click()
Dim mRow As Double
Dim mR As Double
Dim mPremium As Double

If Val(mInsurancePolicyID) = 0 Then
   MsgBox "Select Insurance Policy."
   CmbPolicyAndCoverNoteNo.SetFocus
   Exit Sub
End If
If TxtPolicyPremiumAmt = "" Then
   MsgBox "Enter Policy Premium Amount."
   TxtPolicyPremiumAmt.SetFocus
   Exit Sub
End If

If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" Then
    If Val(TxtAdditionalPremium) = 0 Then
       MsgBox "Enter Additional Premium Amount."
       TxtAdditionalPremium.SetFocus
       Exit Sub
    End If
    If Trim(TxtAdditionalPremiumRemark.Text) = "" Then
       MsgBox "Enter Remark."
       TxtAdditionalPremiumRemark.SetFocus
       Exit Sub
    End If
Else
    If Val(TxtAdditionalPremium) > 0 Then
        If Trim(TxtAdditionalPremiumRemark.Text) = "" Then
           MsgBox "Enter Remark."
           TxtAdditionalPremiumRemark.SetFocus
           Exit Sub
        End If
    End If
End If

If CDate(TxtAdditionalPremiumFromDate.Value) < CDate(TxtValidFromDate) Then
   MsgBox "Additional Premium From Date can not less then Policy From Date."
   TxtAdditionalPremiumFromDate.SetFocus
   Exit Sub
End If
If CDate(TxtAdditionalPremiumFromDate.Value) > CDate(TxtValidToDate) Then
   MsgBox "Additional Premium From Date can not bigger then Policy To Date."
   TxtAdditionalPremiumFromDate.SetFocus
   Exit Sub
End If
If CDate(TxtAdditionalPremiumToDate.Value) < CDate(TxtValidFromDate) Then
   MsgBox "Additional Premium To Date can not less then Policy From Date."
   TxtAdditionalPremiumFromDate.SetFocus
   Exit Sub
End If
If CDate(TxtAdditionalPremiumToDate.Value) > CDate(TxtValidToDate) Then
   MsgBox "Additional Premium to Date can not bigger then Policy To Date."
   TxtAdditionalPremiumFromDate.SetFocus
   Exit Sub
End If
If CDate(TxtAdditionalPremiumFromDate.Value) > CDate(TxtAdditionalPremiumToDate.Value) Then
   MsgBox "Additional Premium From Date can not bigger then Additional Premium to Date."
   TxtAdditionalPremiumFromDate.SetFocus
   Exit Sub
End If
If Format(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) & "-" & "01", "mmm-yyyy") <> Format(TxtAdditionalPremiumFromDate.Value, "mmm-yyyy") Then
   MsgBox "Additional Premium From Date's 'Month-Year' can not more than and less then the selected premium month."
   TxtAdditionalPremiumFromDate.SetFocus
   Exit Sub
End If

''-- Validation to Check Back Date Month not updated
mRow = MSHFlexGrid2.RowSel
    If CheckForPrevoiusMonthPremiumDetail(mRow) = True Then
        If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" Then
            MsgBox "Previous Month-Year Premium details are not allow to add and update, if future Month-Year Premium details are already updated."
            Exit Sub
        Else
            If MSHFlexGrid2.TextMatrix(mRow, 6) <> "" And MSHFlexGrid2.TextMatrix(mRow, 7) <> "" Then
                ''-- if the rentered Back date month-year are updating
                Select Case MsgBox("Previous Month-Year Premium are not allow to add and update." & vbNewLine & "Do you want to update the Premium Details? It will change the mothly premium as per calculation.", vbYesNo + vbQuestion)
                    Case vbNo
                        Exit Sub
                End Select
            Else
                ''-- if the Blank Back date month-year are updating
                MsgBox "Previous Month-Year Premium details are not allow to add and update, if future Month-Year Premium details are already updated."
                Exit Sub
            End If
        End If
    End If
''---

''-- Check for PnL
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "headcmg" Then
    If MSHFlexGrid2.TextMatrix(mRow, 6) <> "" And MSHFlexGrid2.TextMatrix(mRow, 7) <> "" Then
        ''-- If Pnl is completed for same month then not allow to update.
        If CheckForPLforSameMonth(Format(MSHFlexGrid2.TextMatrix(mRow, 1), "YYYY-MM-01")) = True Then
            MsgBox "PnL Process is completed for " & MSHFlexGrid2.TextMatrix(mRow, 1) & ",therefore same Month-Year Premium are not allow to update."
            Exit Sub
        End If
    End If
Else
    If MSHFlexGrid2.TextMatrix(mRow, 6) <> "" And MSHFlexGrid2.TextMatrix(mRow, 7) <> "" Then
        ''-- If Pnl is completed for same month then not allow to update.
        If CheckForPLforSameMonth(Format(MSHFlexGrid2.TextMatrix(mRow, 1), "YYYY-MM-01")) = True Then
            Select Case MsgBox("PnL Process is completed for " & MSHFlexGrid2.TextMatrix(mRow, 1) & ",therefore same Month-Year Premium are not allow to update." & vbNewLine & "Do you want to update the Premium?", vbYesNo + vbQuestion)
                Case vbNo
                    Exit Sub
            End Select
        End If
    End If
End If
''---

With MSHFlexGrid2
    mNoofMonths = Format(DateDiff("d", TxtAdditionalPremiumFromDate.Value, TxtAdditionalPremiumToDate.Value) / 30, "#0")
    mRow = .RowSel
    If Val(TxtAdditionalPremium.Text) > 0 Then
        .TextMatrix(mRow, 3) = Format(TxtAdditionalPremium.Text, "#0.00")
        .TextMatrix(mRow, 6) = Format(TxtAdditionalPremiumFromDate.Value, "DD-MMM-YYYY")
        .TextMatrix(mRow, 7) = Format(TxtAdditionalPremiumToDate.Value, "DD-MMM-YYYY")
        .TextMatrix(mRow, 8) = TxtAdditionalPremiumRemark.Text
        .TextMatrix(mRow, 4) = Format(TxtAdditionalPremiumMonthly, "#0.00")
        .TextMatrix(mRow, 5) = Format(Val(.TextMatrix(mRow, 2)) + Val(.TextMatrix(mRow, 4)), "#0.00")
    Else
        mPremium = IIf(Val(.TextMatrix(mRow, 3)) = 0, 0, Val(.TextMatrix(mRow, 3)) / mNoofMonths)
        .TextMatrix(mRow, 3) = Format(0, "#0.00")
        .TextMatrix(mRow, 6) = ""
        .TextMatrix(mRow, 7) = ""
        .TextMatrix(mRow, 8) = TxtAdditionalPremiumRemark.Text
        Call RemovePremiumPlan(mRow, mPremium)
    End If
    Call UpdatePremiumPlan
    
    ''-- Clear Frame
    TxtAdditionalPremium.Text = ""
    TxtAdditionalPremiumFromDate.Value = Date
    TxtAdditionalPremiumToDate.Value = Date
    TxtAdditionalPremiumMonthly.Text = ""
    TxtAdditionalPremiumRemark = ""
    Frame2.Enabled = False
    
End With
End Sub

Private Sub RemovePremiumPlan(mROw_ As Double, mPremium_ As Double)
Dim mR As Double, mRow As Double
With MSHFlexGrid2
    For mR = mROw_ To .Rows - 2
        .TextMatrix(mR, 4) = Format(Val(.TextMatrix(mR, 4)) - Val(mPremium_), "#0.00")
        .TextMatrix(mR, 5) = Format(Val(.TextMatrix(mR, 2)) + Val(.TextMatrix(mR, 4)), "#0.00")
        .TextMatrix(mR, 10) = "1"
    Next mR
End With
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
Call TableTxn_InsurancePremium
If RsTxn_InsurancePremium.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsTxn_InsurancePremium.MoveLast
Call Indata
Call GButtonLock(Me, True)
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Cols = 11
    .Rows = 2
    .WordWrap = True
    .TextMatrix(0, 0) = "Insurance Premium ID":             .ColWidth(0) = 1200
    .TextMatrix(0, 1) = "Date":                             .ColWidth(1) = 1500
    .TextMatrix(0, 2) = "Policy No.":                       .ColWidth(2) = 2500
    .TextMatrix(0, 3) = "Cover Note No.":                   .ColWidth(3) = 2500
    .TextMatrix(0, 4) = "Insurance By":                     .ColWidth(4) = 1500
    .TextMatrix(0, 5) = "Valid From":                       .ColWidth(5) = 1500
    .TextMatrix(0, 6) = "Valid To":                         .ColWidth(6) = 1500
    .TextMatrix(0, 7) = "Insurance Risk Type":              .ColWidth(7) = 2000
    .TextMatrix(0, 8) = "Insurance Company Name":           .ColWidth(8) = 2500
    .TextMatrix(0, 9) = "Insurance Amount":                 .ColWidth(9) = 2000
    .TextMatrix(0, 10) = "Insurance Premium (Annually)":    .ColWidth(10) = 2000
    .RowHeight(0) = 600
End With
End Sub


Private Sub Grid_Head2()
With MSHFlexGrid2
    .Clear
    .Cols = 11
    .Rows = 2
    .WordWrap = True
    .TextMatrix(0, 0) = "Sr.No.":                                              .ColWidth(0) = 600
    .TextMatrix(0, 1) = "Premium Month":                                       .ColWidth(1) = 1100
    .TextMatrix(0, 2) = "Policy Premium (Monthly) As Per Annually Premium":    .ColWidth(2) = 2500
    .TextMatrix(0, 3) = "Additional Premium":                                  .ColWidth(3) = 1800
    .TextMatrix(0, 4) = "Additional Premium (Monthly)":                        .ColWidth(4) = 1800
    .TextMatrix(0, 5) = "Total Monthly Premium":                               .ColWidth(5) = 1500
    .TextMatrix(0, 6) = "Additional Premium From Date":                        .ColWidth(6) = 1900
    .TextMatrix(0, 7) = "Additional Premium To Date":                          .ColWidth(7) = 1900
    .TextMatrix(0, 8) = "Remark":                                              .ColWidth(8) = 2500
    .TextMatrix(0, 9) = "DetailID":                                            .ColWidth(9) = 0
    .TextMatrix(0, 10) = "Flag":                                               .ColWidth(10) = 0
    .RowHeight(0) = 600
End With
End Sub

Private Sub Grid_Head3()
With MSHFlexGrid3
    .Clear
    .Cols = 9
    .Rows = 2
    .WordWrap = True
    .TextMatrix(0, 0) = "Premium Month":                                       .ColWidth(0) = 1100
    .TextMatrix(0, 1) = "Policy Premium (Monthly) As Per Annually Premium":    .ColWidth(1) = 1800
    .TextMatrix(0, 2) = "Additional Premium":                                  .ColWidth(2) = 1500
    .TextMatrix(0, 3) = "Additional Premium (Monthly)":                        .ColWidth(3) = 1500
    .TextMatrix(0, 4) = "Total Monthly Premium":                               .ColWidth(4) = 1500
    .TextMatrix(0, 5) = "Additional Premium From Date":                        .ColWidth(5) = 1500
    .TextMatrix(0, 6) = "Additional Premium To Date":                          .ColWidth(6) = 1500
    .TextMatrix(0, 7) = "Remark":                                              .ColWidth(7) = 2300
    .TextMatrix(0, 8) = "UpdatedBy":                                           .ColWidth(8) = 2800
    .RowHeight(0) = 800
End With
End Sub

Private Sub ClearForm()
    TxtInsurancePremiumDate.Text = ""
    TxtInsurancePolicyID = ""
    TxtPolicyAndCoverNoteNo = ""
    CmbPolicyAndCoverNoteNo.Text = ""
    TxtInsuranceBy.Text = ""
    TxtValidFromDate.Text = ""
    TxtValidToDate.Text = ""
    TxtInsuranceRiskType.Text = ""
    TxtInsuranceCompany.Text = ""
    TxtPolicyPremiumAmt.Text = "": TxtPolicyPremiumAmt.Locked = False
    TxtInsuranceAmount.Text = ""
    TxtStructureWiseLimit.Text = ""
    
    ''-- Additional Premium
    TxtAdditionalPremium.Text = ""
    TxtAdditionalPremiumFromDate.Value = Date
    TxtAdditionalPremiumToDate.Value = Date
    TxtAdditionalPremiumMonthly.Text = ""
    TxtAdditionalPremiumRemark = ""
    ''-- Premium Detail Grid
    Call Grid_Head2
    ''-- Premium Detail Log Grid
    Call Grid_Head3
    TxtCreated = ""
    TxtUpdated = ""
    SSTab1.Tab = 0
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
Call ClearForm
TxtInsurancePremiumID.Text = ""
TxtPolicyAndCoverNoteNo.Visible = False
CmbPolicyAndCoverNoteNo.Visible = True
TxtInsurancePremiumDate.Text = Format(Date, "DD-MMM-YYYY")
Frame2.Enabled = False
If RsTxn_InsurancePremium.RecordCount > 0 Then
   CmbPolicyAndCoverNoteNo.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmdGo_Click()
Dim mR As Double
Dim wc As Variant
wc = ""
Call Grid_Head
Label12.Caption = "-"

If CmbSPolicyAndCoverNoteNo <> "" Then
   wc = " Where TIP.InsurancePolicyID =  " & Val(mSInsurancePolicyID) & " "
End If
tmp = " Select TIP.InsurancePremiumID,TIP.InsurancePremiumDate,TIP.InsurancePolicyID,TIP.PolicyPremium,A.InsurancePolicyID,A.PolicyNo,A.CoverNoteNo,A.InsuranceByID,B.InsuranceByName,A.Amount,A.StructureWiseLimit,A.ValidFrom,A.ValidTo,A.InsuranceCoID,C.InsuranceCoName,A.InsuranceRiskCoveredID,D.InsuranceRiskCovered"
tmp = tmp & " From Txn_InsurancePremium TIP"
tmp = tmp & " Inner Join Mst_InsuranceDetails As A On TIP.InsurancePolicyID = A.InsurancePolicyID"
tmp = tmp & " Left Join Mst_InsuranceBy B On A.InsuranceByID = B.InsuranceByID"
tmp = tmp & " Left Join Mst_InsuranceCoMaster C On A.InsuranceCoID = C.InsuranceCoID"
tmp = tmp & " Left Join Mst_InsuranceRiskCovered D On A.InsuranceRiskCoveredID = D.InsuranceRiskCoveredID"
tmp = tmp & wc
Call TmpTable
For mR = 1 To TmpRs.RecordCount
With MSHFlexGrid1
    
    .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!InsurancePremiumID), "", TmpRs!InsurancePremiumID)
    .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!InsurancePremiumDate), "", Format(TmpRs!InsurancePremiumDate, "DD-MMM-YYYY"))
    .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!PolicyNo), "", TmpRs!PolicyNo)
    .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!CoverNoteNo), "", TmpRs!CoverNoteNo)
    .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!InsuranceByName), "", TmpRs!InsuranceByName)
    .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!ValidFrom), "", Format(TmpRs!ValidFrom, "DD-MMM-YYYY"))
    .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!ValidTo), "", Format(TmpRs!ValidTo, "DD-MMM-YYYY"))
    .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!InsuranceRiskCovered), "", TmpRs!InsuranceRiskCovered)
    .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!InsuranceCoName), "", TmpRs!InsuranceCoName)
    .TextMatrix(mR, 9) = IIf(IsNull(TmpRs!Amount), "", Format(TmpRs!Amount, "#0.00"))
    .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!PolicyPremium), "", Format(TmpRs!PolicyPremium, "#0.00"))
    
    .Rows = .Rows + 1
    TmpRs.MoveNext
End With
Next
Label12.Refresh
Label12.Caption = TmpRs.RecordCount

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   
   tmp = " Delete From Txn_InsurancePremiumDetail Where InsurancePremiumID = " & Val(TxtInsurancePremiumID) & " "
   Call TmpTable
   
   tmp = " Delete From Log_Txn_InsurancePremiumDetail Where InsurancePremiumID = " & Val(TxtInsurancePremiumID) & " "
   Call TmpTable
   
   RsTxn_InsurancePremium.Delete
   RsTxn_InsurancePremium.Update
   MsgBox "Record Successfully Removed"
   If RsTxn_InsurancePremium.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_InsurancePremium.MoveNext
   If RsTxn_InsurancePremium.EOF() Then
      RsTxn_InsurancePremium.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_InsurancePremium.CancelUpdate
End Sub

Private Sub EditCmd_Click()
''-- If Policy is expired then not allow to edit the premium plan
tmp = " Select * from Mst_InsuranceDetails Where InsurancePolicyID = " & mInsurancePolicyID & ""
Call TmpTable
If TmpRs.RecordCount > 0 Then
    If DateDiff("d", Date, TmpRs!ValidTo) < 0 Then
       MsgBox "Insurance Policy is Expired, So not allow to Edit the Insurance Premium."
       Exit Sub
    End If
End If
''--------------------------
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtPolicyAndCoverNoteNo.Visible = True
CmbPolicyAndCoverNoteNo.Visible = False
TxtPolicyPremiumAmt.Locked = True
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_InsurancePremium.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Public Sub Indata()
Call ClearForm
TxtPolicyAndCoverNoteNo.Visible = True
CmbPolicyAndCoverNoteNo.Visible = False
Frame2.Enabled = False

TxtInsurancePremiumID = RsTxn_InsurancePremium!InsurancePremiumID
If RsTxn_InsurancePremium.EOF = False Then
    mInsurancePolicyID = RsTxn_InsurancePremium!InsurancePolicyID
    Call InsurancePolicyDetails(mInsurancePolicyID)
    TxtInsurancePremiumDate = Format(RsTxn_InsurancePremium!InsurancePremiumDate, "DD-MMM-YYYY")
    TxtPolicyPremiumAmt = Format(RsTxn_InsurancePremium!PolicyPremium, "#0.00")
    TxtPolicyPremiumAmt.Locked = True
    TxtCreated = IIf(IsNull(RsTxn_InsurancePremium!CreatedBy), "", RsTxn_InsurancePremium!CreatedBy) & " :- " & IIf(IsNull(RsTxn_InsurancePremium!CreatedDate), "", RsTxn_InsurancePremium!CreatedDate)
    TxtUpdated = IIf(IsNull(RsTxn_InsurancePremium!UpdatedBy), "", RsTxn_InsurancePremium!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_InsurancePremium!UpdatedDate), "", RsTxn_InsurancePremium!UpdatedDate)
    Call InDataDetails
    Call IndataDetailsLog
End If
Call GButtonLock(Me, True)
End Sub

Public Sub InDataDetails()
Dim mR As Double
Call Grid_Head2
With MSHFlexGrid2
    Call TableTxn_InsurancePremiumDetail(" Where InsurancePremiumID = " & Val(TxtInsurancePremiumID) & " ")
    If RsTxn_InsurancePremiumDetail.EOF = False Then
        For mR = 1 To RsTxn_InsurancePremiumDetail.RecordCount
            .TextMatrix(mR, 0) = mR
            .TextMatrix(mR, 1) = Format(RsTxn_InsurancePremiumDetail!PremiumMonth, "MMM-YYYY")
            .TextMatrix(mR, 2) = Format(RsTxn_InsurancePremiumDetail!PolicyPremiumMonthly, "#0.00")
            .TextMatrix(mR, 3) = Format(RsTxn_InsurancePremiumDetail!AdditionalPremium, "#0.00")
            .TextMatrix(mR, 4) = Format(RsTxn_InsurancePremiumDetail!AdditionalPremiumMonthly, "#0.00")
            .TextMatrix(mR, 5) = Format(RsTxn_InsurancePremiumDetail!TotalMonthlyPremium, "#0.00")
            .TextMatrix(mR, 6) = IIf(IsNull(RsTxn_InsurancePremiumDetail!FromDate), "", Format(RsTxn_InsurancePremiumDetail!FromDate, "DD-MMM-YYYY"))
            .TextMatrix(mR, 7) = IIf(IsNull(RsTxn_InsurancePremiumDetail!ToDate), "", Format(RsTxn_InsurancePremiumDetail!ToDate, "DD-MMM-YYYY"))
            .TextMatrix(mR, 8) = IIf(IsNull(RsTxn_InsurancePremiumDetail!Remark), "", RsTxn_InsurancePremiumDetail!Remark)
            .TextMatrix(mR, 9) = RsTxn_InsurancePremiumDetail!InsurancePremiumDetailID
            .TextMatrix(mR, 10) = "0"
            .Rows = .Rows + 1
            RsTxn_InsurancePremiumDetail.MoveNext
        Next mR
    End If
End With
End Sub

Public Sub IndataDetailsLog()
Dim mR As Double
Call Grid_Head3
With MSHFlexGrid3
    Call TableLog_Txn_InsurancePremiumDetail(" Where InsurancePremiumID = " & Val(TxtInsurancePremiumID) & " Order By InsurancePremiumDetailID Asc,LogConcurrency Desc ")
    If RsLog_Txn_InsurancePremiumDetail.EOF = False Then
        For mR = 1 To RsLog_Txn_InsurancePremiumDetail.RecordCount
            .TextMatrix(mR, 0) = Format(RsLog_Txn_InsurancePremiumDetail!PremiumMonth, "MMM-YYYY")
            .TextMatrix(mR, 1) = Format(RsLog_Txn_InsurancePremiumDetail!PolicyPremiumMonthly, "#0.00")
            .TextMatrix(mR, 2) = Format(RsLog_Txn_InsurancePremiumDetail!AdditionalPremium, "#0.00")
            .TextMatrix(mR, 3) = Format(RsLog_Txn_InsurancePremiumDetail!AdditionalPremiumMonthly, "#0.00")
            .TextMatrix(mR, 4) = Format(RsLog_Txn_InsurancePremiumDetail!TotalMonthlyPremium, "#0.00")
            .TextMatrix(mR, 5) = IIf(IsNull(RsLog_Txn_InsurancePremiumDetail!FromDate), "", Format(RsLog_Txn_InsurancePremiumDetail!FromDate, "DD-MMM-YYYY"))
            .TextMatrix(mR, 6) = IIf(IsNull(RsLog_Txn_InsurancePremiumDetail!ToDate), "", Format(RsLog_Txn_InsurancePremiumDetail!ToDate, "DD-MMM-YYYY"))
            .TextMatrix(mR, 7) = IIf(IsNull(RsLog_Txn_InsurancePremiumDetail!Remark), "", RsLog_Txn_InsurancePremiumDetail!Remark)
            .TextMatrix(mR, 8) = IIf(IsNull(RsLog_Txn_InsurancePremiumDetail!UpdatedBy), "", RsLog_Txn_InsurancePremiumDetail!UpdatedBy & "~" & Format(RsLog_Txn_InsurancePremiumDetail!UpdatedDate, "DD-MMM-YYYY hh:mm:ss"))
            .Rows = .Rows + 1
            RsLog_Txn_InsurancePremiumDetail.MoveNext
        Next mR
    End If
End With
End Sub

Private Sub LastCmd_Click()
RsTxn_InsurancePremium.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
With MSHFlexGrid1
RsTxn_InsurancePremium.MoveFirst
If RsTxn_InsurancePremium.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsTxn_InsurancePremium.MoveFirst
RsTxn_InsurancePremium.Find "InsurancePremiumID = " & Val(.TextMatrix(.RowSel, 0)) & ""
If RsTxn_InsurancePremium.EOF Then
   RsTxn_InsurancePremium.MoveLast
End If
Frame0.Visible = False
Call Indata
End With
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
Dim mR As Double
With MSHFlexGrid2
    If .Rows > 2 Then
        mR = .RowSel
        If Trim(.TextMatrix(mR, 1)) = "" Then
            TxtAdditionalPremium.Text = ""
            TxtAdditionalPremiumMonthly.Text = ""
            TxtAdditionalPremiumFromDate.Value = Date
            TxtAdditionalPremiumToDate.Value = Date
            TxtAdditionalPremiumRemark.Text = ""
            Frame2.Enabled = False
            Exit Sub
        End If
        Frame2.Enabled = True
        If Val(.TextMatrix(.RowSel, 3)) <= 0 Then
            TxtAdditionalPremium.Text = ""
            TxtAdditionalPremiumMonthly.Text = ""
            TxtAdditionalPremiumFromDate.Value = CDate(Format(.TextMatrix(mR, 1) & "-" & Format(CDate(TxtValidFromDate), "dd"), "yyyy-mm-dd"))
            TxtAdditionalPremiumToDate.Value = CDate(TxtValidToDate)
            TxtAdditionalPremiumRemark.Text = ""
        Else
            TxtAdditionalPremium.Text = Val(.TextMatrix(.RowSel, 3))
            TxtAdditionalPremiumFromDate.Value = CDate(.TextMatrix(.RowSel, 6))
            TxtAdditionalPremiumToDate.Value = CDate(.TextMatrix(.RowSel, 7))
            TxtAdditionalPremiumRemark.Text = .TextMatrix(.RowSel, 8)
            TxtAdditionalPremiumMonthly.Text = Format(Val(.TextMatrix(.RowSel, 3)) / Format(DateDiff("d", TxtAdditionalPremiumFromDate.Value, TxtAdditionalPremiumToDate.Value) / 30, "#0"), "#0.00")
        End If
    Else
        TxtAdditionalPremium.Text = ""
        TxtAdditionalPremiumMonthly.Text = ""
        TxtAdditionalPremiumFromDate.Value = Date
        TxtAdditionalPremiumToDate.Value = Date
        TxtAdditionalPremiumRemark.Text = ""
        Frame2.Enabled = False
    End If
End With
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_InsurancePremium.MoveNext
If RsTxn_InsurancePremium.EOF Then
   RsTxn_InsurancePremium.MoveLast
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
RsTxn_InsurancePremium.MovePrevious
If RsTxn_InsurancePremium.BOF Then
   RsTxn_InsurancePremium.MoveFirst
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
If Edit_Flg = "A" Then
    If Trim(CmbPolicyAndCoverNoteNo) = "" Then
       MsgBox "Blank Insurance Policy And Cover Note No not allow."
       CmbPolicyAndCoverNoteNo.SetFocus
       Exit Sub
    End If
    If Trim(TxtPolicyPremiumAmt) = "" Then
       MsgBox "Blank Policy Premium Amount not Allowed !!!"
       TxtPolicyPremiumAmt.SetFocus
       Exit Sub
    End If
End If

If Edit_Flg = "A" Then
   If RsTxn_InsurancePremium.RecordCount > 0 Then
        RsTxn_InsurancePremium.MoveFirst
        RsTxn_InsurancePremium.Find "InsurancePolicyID = " & Val(InsurancePolicyID) & ""
        If Not RsTxn_InsurancePremium.EOF Then
           MsgBox "Duplicate Insurance Premium with same Policy No not Allowed."
           Exit Sub
        End If
   End If
   RsTxn_InsurancePremium.AddNew
   RsTxn_InsurancePremium!InsurancePremiumID = GetMaxID("InsurancePremiumID", "Txn_InsurancePremium")
   TxtInsurancePremiumID = RsTxn_InsurancePremium!InsurancePremiumID
   RsTxn_InsurancePremium!G_UID = GetGUID
   RsTxn_InsurancePremium!InsurancePremiumDate = Date
ElseIf Edit_Flg = "E" Then
   RsTxn_InsurancePremium.MoveFirst
   RsTxn_InsurancePremium.Find "InsurancePremiumID = " & Val(TxtInsurancePremiumID)
   If Not RsTxn_InsurancePremium.EOF Then
      If RsTxn_InsurancePremium!InsurancePremiumID <> Val(TxtInsurancePremiumID) Then
         MsgBox "Duplicate Insurance Premium with same Policy No not Allowed."
         Exit Sub
      End If
   End If
   RsTxn_InsurancePremium.MoveFirst
   RsTxn_InsurancePremium.Find "InsurancePremiumID = " & Val(TxtInsurancePremiumID)
End If

RsTxn_InsurancePremium!InsurancePolicyID = mInsurancePolicyID
RsTxn_InsurancePremium!PolicyPremium = TxtPolicyPremiumAmt
If IsNull(RsTxn_InsurancePremium!CreatedBy) = True Or RsTxn_InsurancePremium!CreatedBy = "" Then
   RsTxn_InsurancePremium!CreatedBy = Gen_UserLoginID
   RsTxn_InsurancePremium!CreatedDate = Now
Else
   RsTxn_InsurancePremium!UpdatedBy = Gen_UserLoginID
   RsTxn_InsurancePremium!UpdatedDate = Now
End If
RsTxn_InsurancePremium.Update
Call SaveInsurancePremiumDetail

RsTxn_InsurancePremium.Requery
RsTxn_InsurancePremium.MoveFirst
RsTxn_InsurancePremium.Find "InsurancePremiumID = " & Val(TxtInsurancePremiumID)
Call GButtonLock(Me, True)
Call Indata
End Sub

Private Sub SaveInsurancePremiumDetail()
Dim mR As Double, mC As Double
With MSHFlexGrid2
    
    Call TableTxn_InsurancePremiumDetail(" Where InsurancePremiumID = " & Val(TxtInsurancePremiumID) & " ")
    For mR = 1 To .Rows - 2
        If .TextMatrix(mR, 1) <> "" Then
            If .TextMatrix(mR, 10) = "1" Then
                If RsTxn_InsurancePremiumDetail.RecordCount > 0 Then RsTxn_InsurancePremiumDetail.MoveFirst
                RsTxn_InsurancePremiumDetail.Find "InsurancePremiumDetailID = " & Val(.TextMatrix(mR, 9))
                If RsTxn_InsurancePremiumDetail.EOF = True Then
                    RsTxn_InsurancePremiumDetail.AddNew
                    RsTxn_InsurancePremiumDetail!InsurancePremiumDetailID = GetMaxID("InsurancePremiumDetailID", "Txn_InsurancePremiumDetail")
                    .TextMatrix(mR, 9) = RsTxn_InsurancePremiumDetail!InsurancePremiumDetailID
                    RsTxn_InsurancePremiumDetail!InsurancePremiumID = Val(TxtInsurancePremiumID)
                    RsTxn_InsurancePremiumDetail!G_UID = GetGUID
                End If
            
                RsTxn_InsurancePremiumDetail!PremiumMonth = Format("01-" & .TextMatrix(mR, 1), "YYYY-MM-01")
                RsTxn_InsurancePremiumDetail!PolicyPremiumMonthly = Val(.TextMatrix(mR, 2))
                RsTxn_InsurancePremiumDetail!AdditionalPremium = Val(.TextMatrix(mR, 3))
                RsTxn_InsurancePremiumDetail!AdditionalPremiumMonthly = Val(.TextMatrix(mR, 4))
                RsTxn_InsurancePremiumDetail!TotalMonthlyPremium = Val(.TextMatrix(mR, 5))
                RsTxn_InsurancePremiumDetail!FromDate = IIf(.TextMatrix(mR, 6) = "", Null, Format(.TextMatrix(mR, 6), "YYYY-MM-DD"))
                RsTxn_InsurancePremiumDetail!ToDate = IIf(.TextMatrix(mR, 7) = "", Null, Format(.TextMatrix(mR, 7), "YYYY-MM-DD"))
                RsTxn_InsurancePremiumDetail!Remark = IIf(.TextMatrix(mR, 8) = "", Null, Trim(.TextMatrix(mR, 8)))
                
                If IsNull(RsTxn_InsurancePremiumDetail!CreatedBy) = True Or RsTxn_InsurancePremiumDetail!CreatedBy = "" Then
                   RsTxn_InsurancePremiumDetail!CreatedBy = Gen_UserLoginID
                   RsTxn_InsurancePremiumDetail!CreatedDate = Now
                Else
                   RsTxn_InsurancePremiumDetail!UpdatedBy = Gen_UserLoginID
                   RsTxn_InsurancePremiumDetail!UpdatedDate = Now
                End If
                RsTxn_InsurancePremiumDetail.Update
                RsTxn_InsurancePremiumDetail.Requery
            End If
        End If
    Next mR

End With
End Sub


Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsTxn_InsurancePremium.EOF Then
      RsTxn_InsurancePremium.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Frame1.Width
Frame0.Height = Me.Height - 500
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1500
Frame0.Visible = True

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_InsurancePremium.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_InsurancePremium.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
With MSHFlexGrid1
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
           If C = 2 And C = 3 Then
            oWS.Cells(mRow, C + 1) = "`" & .TextMatrix(I, C)
           Else
            oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
           End If
        End If
    Next
Next
End With
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Sub InsurancePolicyDetails(Id_ As Double)
Dim mR As Double
    tmp = " Select A.InsurancePolicyID,A.PolicyNo,A.CoverNoteNo,A.InsuranceByID,B.InsuranceByName,A.Amount,A.StructureWiseLimit,A.ValidFrom,A.ValidTo,A.InsuranceCoID,C.InsuranceCoName,A.InsuranceRiskCoveredID,D.InsuranceRiskCovered"
    tmp = tmp & " From Mst_InsuranceDetails As A"
    tmp = tmp & " Left Join Mst_InsuranceBy B On A.InsuranceByID = B.InsuranceByID"
    tmp = tmp & " Left Join Mst_InsuranceCoMaster C On A.InsuranceCoID = C.InsuranceCoID"
    tmp = tmp & " Left Join Mst_InsuranceRiskCovered D On A.InsuranceRiskCoveredID = D.InsuranceRiskCoveredID"
    tmp = tmp & " Where A.InsurancePolicyID = " & Id_ & ""
    Call Tmp1Table
    If TmpRs1.RecordCount > 0 Then
        TxtPolicyAndCoverNoteNo.Text = IIf(IsNull(TmpRs1!PolicyNo), "", TmpRs1!PolicyNo) & "~" & IIf(IsNull(TmpRs1!CoverNoteNo), "", TmpRs1!CoverNoteNo)
        TxtInsurancePolicyID.Text = IIf(IsNull(TmpRs1!InsurancePolicyID), "", TmpRs1!InsurancePolicyID)
        TxtInsuranceBy.Text = IIf(IsNull(TmpRs1!InsuranceByName), "", TmpRs1!InsuranceByName)
        TxtInsuranceCompany.Text = IIf(IsNull(TmpRs1!InsuranceCoName), "", TmpRs1!InsuranceCoName)
        TxtValidFromDate.Text = Format(TmpRs1!ValidFrom, "DD-MMM-YYYY")
        TxtValidToDate.Text = Format(TmpRs1!ValidTo, "DD-MMM-YYYY")
        TxtInsuranceAmount.Text = Format(IIf(IsNull(TmpRs1!Amount), "", TmpRs1!Amount), "#0.00")
        TxtStructureWiseLimit.Text = IIf(IsNull(TmpRs1!StructureWiseLimit), "", Format(TmpRs1!StructureWiseLimit, "#0.00"))
    Else
        TxtPolicyAndCoverNoteNo.Text = ""
        TxtInsurancePolicyID.Text = ""
        TxtInsuranceBy.Text = ""
        TxtInsuranceCompany.Text = ""
        TxtValidFromDate.Text = ""
        TxtValidToDate.Text = ""
        TxtInsuranceAmount.Text = ""
        TxtStructureWiseLimit.Text = ""
    End If
End Sub

Private Sub TxtAdditionalPremium_LostFocus()
Dim mNoofMonths As Double
If TxtAdditionalPremium.Text = "" Then
   TxtAdditionalPremiumMonthly.Text = ""
   Exit Sub
End If
If IsNumeric(TxtAdditionalPremium) = False Then
   MsgBox "Invalid Input."
   TxtAdditionalPremium.SetFocus
   Exit Sub
End If
If TxtAdditionalPremium <= 0 Then
   MsgBox "Additional Premium Amount can not be Zero."
   TxtAdditionalPremium.SetFocus
   Exit Sub
End If
If Trim(TxtAdditionalPremium.Text) <> "" Then
    TxtAdditionalPremium = Format(TxtAdditionalPremium, "#0.00")
    mNoofMonths = Format(DateDiff("d", TxtAdditionalPremiumFromDate.Value, TxtAdditionalPremiumToDate.Value) / 30, "#0")
    mNoofMonths = IIf(mNoofMonths = 0, 1, mNoofMonths)
    TxtAdditionalPremiumMonthly = Format(TxtAdditionalPremium.Text / mNoofMonths, "#0.00")
Else
    TxtAdditionalPremiumMonthly = ""
End If
End Sub

Private Sub TxtAdditionalPremiumFromDate_LostFocus()
Dim mNoofMonths As Double
Dim mR As Double

With MSHFlexGrid2
    If CDate(TxtAdditionalPremiumFromDate.Value) < CDate(TxtValidFromDate) Then
       MsgBox "Additional Premium From Date can not less then Policy From Date."
       Exit Sub
       ''TxtAdditionalPremiumFromDate.SetFocus
    End If
    If CDate(TxtAdditionalPremiumFromDate.Value) > CDate(TxtValidToDate) Then
       MsgBox "Additional Premium From Date can not bigger then Policy To Date."
       Exit Sub
       ''TxtAdditionalPremiumFromDate.SetFocus
    End If
    If CDate(TxtAdditionalPremiumFromDate.Value) > CDate(TxtAdditionalPremiumToDate.Value) Then
       MsgBox "Additional Premium From Date can not bigger then Additional Premium to Date."
       Exit Sub
       ''TxtAdditionalPremiumFromDate.SetFocus
    End If
    If Format(.TextMatrix(.RowSel, 1) & "-" & "01", "mmm-yyyy") <> Format(TxtAdditionalPremiumFromDate.Value, "mmm-yyyy") Then
       MsgBox "Additional Premium From Date's 'Month-Year' can not more than and less then the selected premium month."
       Exit Sub
       ''TxtAdditionalPremiumFromDate.SetFocus
    End If
    If Trim(TxtAdditionalPremium.Text) <> "" Then
        mNoofMonths = Format(DateDiff("d", TxtAdditionalPremiumFromDate.Value, TxtAdditionalPremiumToDate.Value) / 30, "#0")
        mNoofMonths = IIf(mNoofMonths = 0, 1, mNoofMonths)
        TxtAdditionalPremiumMonthly = Format(TxtAdditionalPremium.Text / mNoofMonths, "#0.00")
    Else
        TxtAdditionalPremiumMonthly = ""
    End If
End With
End Sub

Private Sub TxtAdditionalPremiumToDate_LostFocus()
With MSHFlexGrid2
    If CDate(TxtAdditionalPremiumToDate.Value) < CDate(TxtValidFromDate) Then
       MsgBox "Additional Premium To Date can not less then Policy From Date."
       Exit Sub
       ''TxtAdditionalPremiumToDate.SetFocus
    End If
    If CDate(TxtAdditionalPremiumToDate.Value) > CDate(TxtValidToDate) Then
       MsgBox "Additional Premium To Date can not more then Policy to Date."
       Exit Sub
       ''TxtAdditionalPremiumToDate.SetFocus
    End If
    If CDate(TxtAdditionalPremiumFromDate.Value) > CDate(TxtAdditionalPremiumToDate.Value) Then
       MsgBox "Additional Premium From Date can not bigger then Additional Premium to Date."
       Exit Sub
       ''TxtAdditionalPremiumToDate.SetFocus
    End If
    
    If Trim(TxtAdditionalPremium.Text) <> "" Then
        mNoofMonths = Format(DateDiff("d", TxtAdditionalPremiumFromDate.Value, TxtAdditionalPremiumToDate.Value) / 30, "#0")
        mNoofMonths = IIf(mNoofMonths = 0, 1, mNoofMonths)
        TxtAdditionalPremiumMonthly = Format(TxtAdditionalPremium.Text / mNoofMonths, "#0.00")
    Else
        TxtAdditionalPremiumMonthly = ""
    End If
    
End With
End Sub

Private Sub TxtPolicyPremiumAmt_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
If TxtInsurancePremiumID.Text <> "" Then Exit Sub
If TxtPolicyPremiumAmt.Text = "" Then
   Call Grid_Head2
   Exit Sub
End If
If IsNumeric(TxtPolicyPremiumAmt) = False Then
   MsgBox "Invalid Input."
   TxtPolicyPremiumAmt.SetFocus
   Exit Sub
End If
If TxtPolicyPremiumAmt <= 0 Then
   MsgBox "Premium Amount can not be Zero or More then Insurance Amount."
   TxtPolicyPremiumAmt.SetFocus
   Exit Sub
End If
If TxtInsuranceAmount <> "" Then
    If TxtPolicyPremiumAmt > Val(TxtInsuranceAmount) Then
       MsgBox "Premium Amount can not be Zero or More then Insurance Amount."
       TxtPolicyPremiumAmt.SetFocus
       Exit Sub
    End If
End If

TxtPolicyPremiumAmt = Format(TxtPolicyPremiumAmt, "#0.00")
If CmbPolicyAndCoverNoteNo.Text = "" Then
   Call Grid_Head2
   Exit Sub
End If
If Val(TxtPolicyPremiumAmt) <= 0 Then
   Call Grid_Head2
   Exit Sub
End If
If Val(TxtInsurancePremiumID.Text) <> 0 Then
   Exit Sub
End If
If DateDiff("d", Date, CDate(TxtValidToDate)) >= 0 Then
   Call MonthlyPremiumPlan
Else
   MsgBox "Insurance Policy is Expired."
   Call Grid_Head2
   Exit Sub
End If
End Sub

Private Sub MonthlyPremiumPlan()
Dim mR As Double, mC As Double
Dim mNoofMonths As Double
Dim mNoofDays As Double
Call Grid_Head2
With MSHFlexGrid2
    mNoofMonths = Format(DateDiff("d", TxtValidFromDate, TxtValidToDate) / 30, "#0")
    mNoofDays = Format(DateDiff("d", TxtValidFromDate, TxtValidToDate), "#0")
    For mR = 1 To mNoofMonths
        .TextMatrix(mR, 0) = mR
        .TextMatrix(mR, 1) = Format(DateAdd("m", mR - 1, CDate(TxtValidFromDate)), "MMM-YYYY")
        .TextMatrix(mR, 2) = Format(TxtPolicyPremiumAmt / mNoofMonths, "#0.00")
        .TextMatrix(mR, 3) = "0.00"
        .TextMatrix(mR, 4) = "0.00"
        .TextMatrix(mR, 5) = Format(Val(.TextMatrix(mR, 2)) + Val(.TextMatrix(mR, 4)), "#0.00")
        .TextMatrix(mR, 6) = ""
        .TextMatrix(mR, 7) = ""
        .TextMatrix(mR, 8) = ""
        .TextMatrix(mR, 10) = "1"
        .Rows = .Rows + 1
    Next mR
    ''-- Additional Premium
    TxtAdditionalPremium.Text = ""
    TxtAdditionalPremiumFromDate.Value = Date
    TxtAdditionalPremiumToDate.Value = Date
    TxtAdditionalPremiumMonthly.Text = ""
    TxtAdditionalPremiumRemark = ""
    Frame2.Enabled = False
End With
End Sub

Private Sub UpdatePremiumPlan()
Dim mRow As Double, mR As Double, mRW As Double
Dim mCurrentPosition As Double, mNoofMonths As Double, mOldNoofMonths As Double
Dim mCarryForwardMontlyAmount As Double, mMontlyPremiunAmount As Double
With MSHFlexGrid2
    If .Rows > 2 Then
        mCarryForwardMontlyAmount = 0#
        For mRow = 1 To .Rows - 2
            If Trim(.TextMatrix(mRow, 1)) <> "" Then
                If Val(.TextMatrix(mRow, 3)) > 0 And Trim(.TextMatrix(mRow, 6)) <> "" And Trim(.TextMatrix(mRow, 7)) <> "" Then
                   mCurrentPosition = mRow
                   mNoofMonths = Format(DateDiff("d", CDate(.TextMatrix(mRow, 6)), CDate(.TextMatrix(mRow, 7))) / 30, "#0")
                   mNoofMonths = IIf(mNoofMonths = 0, 1, mNoofMonths)
                   mMontlyPremiunAmount = Format(Val(.TextMatrix(mRow, 3)) / mNoofMonths, "#0.00")
                   If Val(mOldNoofMonths) < Val(mNoofMonths) Then
                      mCarryForwardMontlyAmount = 0
                   End If
                   mMontlyPremiunAmount = mMontlyPremiunAmount + mCarryForwardMontlyAmount
                   ''--- Update Premium to the Rows --
                    For mR = 0 To mNoofMonths - 1
                        If mR = 0 Then
                            mRW = mCurrentPosition
                        Else
                            mRW = mR + mCurrentPosition
                        End If
                        
                       .TextMatrix(mRW, 4) = Format(0, "#0.00")
                       .TextMatrix(mRW, 5) = Format(Val(.TextMatrix(mRW, 2)), "#0.00")
                        
                       .TextMatrix(mRW, 4) = Format(mMontlyPremiunAmount, "#0.00")
                       .TextMatrix(mRW, 5) = Format(Val(.TextMatrix(mRW, 2)) + mMontlyPremiunAmount, "#0.00")
                       .TextMatrix(mRW, 10) = "1"
                    Next mR
                    mCarryForwardMontlyAmount = mMontlyPremiunAmount 'Format(Val(.TextMatrix(mRow, 3)) / mNoofMonths, "#0.00")
                    mOldNoofMonths = mNoofMonths
                   ''-----------------
                End If
            End If
        Next mRow
    End If
End With
End Sub

Function CheckForPrevoiusMonthPremiumDetail(mR_ As Double) As Boolean
Dim mR As Double
Dim mResult As Boolean
With MSHFlexGrid2
    If .Rows > 2 Then
       For mR = 1 To .Rows - 2
           If Val(.TextMatrix(mR, 3)) > 0 And (mR_ < mR) Then
            mResult = True
           End If
       Next mR
    End If
End With
CheckForPrevoiusMonthPremiumDetail = mResult
End Function

Function CheckForPLforSameMonth(mMonthYear_ As Variant) As Boolean
Dim mResult As Boolean
tmp = " Select Distinct MonthYear from Txn_MonGodownWiseProfitability Where MonthYear  = '" & mMonthYear_ & "' Order By MonthYear"
Call TmpTable
If TmpRs.RecordCount > 0 Then
   mResult = True
End If
CheckForPLforSameMonth = mResult
End Function
