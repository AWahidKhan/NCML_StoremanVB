VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_SecurityAgency 
   Caption         =   "Security Agency"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   8475
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2775
      Left            =   120
      TabIndex        =   41
      Top             =   8040
      Visible         =   0   'False
      Width           =   14985
      Begin VB.ComboBox CmbSStatus 
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
         ItemData        =   "FrmMst_SecurityAgency.frx":0000
         Left            =   7860
         List            =   "FrmMst_SecurityAgency.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   232
         Width           =   1725
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   390
         Left            =   10800
         TabIndex        =   47
         Top             =   217
         Width           =   1020
      End
      Begin VB.TextBox TxtSearchAgency 
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
         Left            =   1485
         TabIndex        =   46
         Top             =   225
         Width           =   5535
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   390
         Left            =   9750
         TabIndex        =   45
         Top             =   217
         Width           =   1020
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1260
         Left            =   120
         TabIndex        =   42
         Top             =   720
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   2223
         _Version        =   393216
         Rows            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label12 
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
         Height          =   255
         Left            =   7185
         TabIndex        =   52
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Label8 
         Caption         =   "Agency Name"
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
         TabIndex        =   50
         Top             =   285
         Width           =   1350
      End
      Begin VB.Label Label13 
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
         Left            =   11940
         TabIndex        =   49
         Top             =   285
         Width           =   2295
      End
      Begin VB.Label LblTotalRec 
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
         Left            =   14250
         TabIndex        =   48
         Top             =   285
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6780
      Left            =   120
      TabIndex        =   30
      Top             =   90
      Width           =   8865
      Begin VB.TextBox TxtTANNo 
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
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   13
         Top             =   4950
         Width           =   2415
      End
      Begin VB.TextBox TxtRTGSNo 
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
         Left            =   5880
         MaxLength       =   100
         TabIndex        =   11
         Top             =   3700
         Width           =   2895
      End
      Begin VB.CommandButton CmdTDSID 
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
         Height          =   315
         Left            =   8400
         TabIndex        =   17
         Top             =   5423
         Width           =   375
      End
      Begin VB.ComboBox CmbTDSID 
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
         Left            =   1920
         TabIndex        =   15
         Top             =   5400
         Width           =   6405
      End
      Begin VB.CommandButton CmdBank 
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
         Height          =   330
         Left            =   8385
         TabIndex        =   7
         Top             =   2760
         Width           =   375
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
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   2760
         Width           =   6405
      End
      Begin VB.CommandButton CmbBranchName 
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
         Height          =   330
         Left            =   8385
         TabIndex        =   9
         Top             =   3240
         Width           =   375
      End
      Begin VB.ComboBox CmbBranch 
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
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   3240
         Width           =   6405
      End
      Begin VB.TextBox TxtAccNo 
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
         MaxLength       =   100
         TabIndex        =   10
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox TxtTDSID 
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
         Left            =   1920
         TabIndex        =   16
         Top             =   5400
         Width           =   6855
      End
      Begin VB.TextBox TxtRemarks 
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
         Height          =   735
         Left            =   1920
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   4140
         Width           =   6855
      End
      Begin VB.ComboBox CmbFlag 
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
         ItemData        =   "FrmMst_SecurityAgency.frx":0021
         Left            =   5880
         List            =   "FrmMst_SecurityAgency.frx":002B
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   4950
         Width           =   2895
      End
      Begin VB.TextBox TxtTelephone2 
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
         Left            =   5880
         TabIndex        =   5
         Top             =   2325
         Width           =   2895
      End
      Begin VB.TextBox TxtMobile 
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
         TabIndex        =   4
         Top             =   2325
         Width           =   2415
      End
      Begin VB.TextBox TxtAddress 
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
         Height          =   780
         Left            =   1920
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1470
         Width           =   6855
      End
      Begin VB.TextBox TxtContactPerson 
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
         TabIndex        =   2
         Top             =   1035
         Width           =   6855
      End
      Begin VB.TextBox TxtAgencyName 
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
         TabIndex        =   1
         Top             =   585
         Width           =   6855
      End
      Begin VB.TextBox TxtAgencyID 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   150
         Width           =   2415
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   5880
         Width           =   6855
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   6315
         Width           =   6855
      End
      Begin VB.TextBox TxtFlag 
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
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   4950
         Width           =   2865
      End
      Begin VB.TextBox TxtBranchName 
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
         Left            =   1920
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   3240
         Width           =   6870
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
         Height          =   360
         Left            =   1920
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   2760
         Width           =   6870
      End
      Begin VB.Label Label33 
         Caption         =   "TAN Number"
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
         TabIndex        =   61
         Top             =   5010
         Width           =   1395
      End
      Begin VB.Label Label16 
         Caption         =   "RTGS No"
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
         Left            =   4440
         TabIndex        =   60
         Top             =   3780
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Bank Name"
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
         TabIndex        =   57
         Top             =   2813
         Width           =   1695
      End
      Begin VB.Label LblBranchName 
         Caption         =   "Branch Name"
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
         TabIndex        =   56
         Top             =   3293
         Width           =   1695
      End
      Begin VB.Label LblAccNo 
         Caption         =   "Account No"
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
         TabIndex        =   55
         Top             =   3780
         Width           =   1695
      End
      Begin VB.Label LblRTGSNo 
         Caption         =   "RTGS No"
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
         Left            =   4350
         TabIndex        =   54
         Top             =   1005
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "TDS Description"
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
         TabIndex        =   53
         Top             =   5460
         Width           =   1710
      End
      Begin VB.Label Label7 
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
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   4380
         Width           =   1695
      End
      Begin VB.Label Label11 
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
         Height          =   255
         Left            =   4560
         TabIndex        =   40
         Top             =   5010
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Telephone 2"
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
         Left            =   4440
         TabIndex        =   39
         Top             =   2385
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Mobile"
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
         TabIndex        =   38
         Top             =   2385
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
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
         TabIndex        =   37
         Top             =   1733
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Contact Person"
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
         TabIndex        =   36
         Top             =   1095
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Agency ID"
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
         TabIndex        =   35
         Top             =   210
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Agency Name"
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
         TabIndex        =   34
         Top             =   645
         Width           =   1695
      End
      Begin VB.Label Label10 
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
         TabIndex        =   32
         Top             =   5940
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         Left            =   120
         TabIndex        =   31
         Top             =   6375
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Width           =   8865
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   6900
         TabIndex        =   21
         Top             =   615
         Width           =   1695
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   5205
         TabIndex        =   29
         Top             =   615
         Width           =   1695
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   3510
         TabIndex        =   28
         Top             =   615
         Width           =   1695
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1815
         TabIndex        =   27
         Top             =   615
         Width           =   1695
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   615
         Width           =   1695
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   6900
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5205
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3510
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1815
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmMst_SecurityAgency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mFlag, mSFlag, mTDSID, mBankID, mBranchID As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtCreated = ""
TxtUpdated = ""
CmbFlag.Visible = True
TxtFlag.Visible = False
CmbFlag.Text = ""
TxtAgencyName.Locked = False
TxtContactPerson.Locked = False
TxtAddress.Locked = False
TxtMobile.Locked = False
TxtTelephone2.Locked = False
TxtRemarks.Locked = False
TxtAgencyID = ""
TxtAgencyName = ""
TxtContactPerson = ""
TxtAddress = ""
CmbBankName = ""
TxtBankName = ""
CmbBranch = ""
TxtBranchName = ""
TxtAccNo = ""
TxtRTGSNo = ""
TxtMobile = ""
TxtTelephone2 = ""
TxtRemarks = ""

TxtBankName.Visible = False
TxtBranchName.Visible = False
CmbBankName.Visible = True
CmbBranch.Visible = True
CmbTDSID.Visible = True
CmdTDSID.Visible = True
TxtTDSID.Visible = False
CmbTDSID.Text = ""
''--- 10 Aug 2017 Ullhas Patil
TxtTANNO = ""
TxtTANNO.Locked = False
''----------------------------
Frame1.Enabled = True
If RsMst_SecurityAgency.RecordCount > 0 Then
   TxtAgencyName.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmbBankName_GotFocus()
tmp = CmbBankName.Text
Call TableMst_Bank("")
CmbBranch.Clear
CmbBankName.Clear
If RsMst_Bank.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbBankName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Bank.RecordCount
    CmbBankName.AddItem RsMst_Bank!BankName
    RsMst_Bank.MoveNext
Next
CmbBankName.Text = tmp
End Sub


Private Sub CmbBankName_LostFocus()
If CmbBankName.Text = "" Then
   mBankID = ""
   Exit Sub
End If
RsMst_Bank.MoveFirst
RsMst_Bank.Find "BankName = '" & CmbBankName.Text & "'"
If RsMst_Bank.EOF Then
   MsgBox "Invalid Bank selection "
   CmbBankName.SetFocus
   Exit Sub
End If
mBankID = RsMst_Bank!BankID
End Sub


Private Sub CmbBranch_GotFocus()
tmp = CmbBranch.Text
If CmbBankName = "" Then
'   MsgBox "Select Bank Name First."
'   CmbBankName.SetFocus
   Exit Sub
End If
Call TableMst_Branch(" Where BankID=" & mBankID & "")

CmbBranch.Clear
If RsMst_Branch.RecordCount = 0 Then
   MsgBox "No Branch found"
   CmbBranch.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Branch.RecordCount
    CmbBranch.AddItem RsMst_Branch!BranchName
    RsMst_Branch.MoveNext
Next
CmbBranch.Text = tmp
End Sub


Private Sub CmbBranch_LostFocus()
If CmbBranch.Text = "" Then
   mBranchID = ""
   Exit Sub
End If
RsMst_Branch.MoveFirst
RsMst_Branch.Find "BranchName = '" & CmbBranch.Text & "'"
If RsMst_Branch.EOF Then
   MsgBox "Invalid Branch selection "
   CmbBranch.SetFocus
   Exit Sub
End If
mBranchID = RsMst_Branch!BranchID
End Sub


Private Sub CmbBranchName_Click()
FrmMst_Branch.Show
End Sub

Private Sub CmbFlag_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbFlag")
End If
End Sub

Private Sub CmbFlag_LostFocus()
If CmbFlag.Text = "" Then Exit Sub

If LCase(CmbFlag.Text) = "active" Then
   mFlag = "Y"
ElseIf LCase(CmbFlag.Text) = "de-active" Then
   mFlag = "N"
Else
   MsgBox "Invalid Selection!!!"
   CmbFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) = "active" Then
   mFlag = "Y"
ElseIf LCase(CmbSStatus.Text) = "de-active" Then
   mFlag = "N"
Else
   MsgBox "Invalid Selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdBank_Click()
FrmMst_Bank.Show
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_SecurityAgency.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_SecurityAgency.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()
Dim wc As Variant
wc = ""
Call Grid_Head
tmp = "Select MSA.SecurityAgencyID,MSA.AgencyName,MTD.Description,MB.BankName,Mbr.BranchName,MSA.BankAcNo,MSA.IFSCCode,MSA.TANNO "
tmp = tmp & " from Mst_SecurityAgency MSA"
tmp = tmp & " Left Join Mst_TDS MTD On MSA.TDSID = MTD.TDSID "
tmp = tmp & " Left Join Mst_Bank MB On MSA.BankId = MB.BankId "
tmp = tmp & " Left Join Mst_Branch MBr On MSA.BranchID = MBr.BranchID"
If TxtSearchAgency.Text <> "" Then
   wc = " Where AgencyName Like '%" & TxtSearchAgency.Text & "%'"
End If

If LCase(CmbSStatus.Text) = "active" Then
   mSFlag = "Y"
ElseIf LCase(CmbSStatus.Text) = "de-active" Then
   mSFlag = "N"
Else
   mSFlag = ""
End If

If CmbSStatus.Text <> "" Then
   If wc = "" Then
      wc = " Where MSA.Flag = '" & mSFlag & "'"
   Else
      wc = wc & " AND MSA.Flag = '" & mSFlag & "'"
   End If
End If

tmp = tmp & wc
Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!SecurityAgencyID
       MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!AgencyName
       MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!Description), "", TmpRs!Description)
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!BankName), "", TmpRs!BankName)
       MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!BranchName), "", TmpRs!BranchName)
       MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs!BankAcNo), "", "'" & TmpRs!BankAcNo)
       MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!IFSCCode), "", TmpRs!IFSCCode)
       MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!TANNO), "", TmpRs!TANNO)
       TmpRs.MoveNext
   Next
End If
LblTotalRec.Refresh
LblTotalRec.Caption = TmpRs.RecordCount
End Sub

Private Sub CmdTDSID_Click()
FrmMst_TDS.Show
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   tmp = "Select * from Mst_SecurityDetails Where SecurityAgencyID = " & TxtAgencyID & ""
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      MsgBox "Child Record Present!!!"
      Exit Sub
   End If
   RsMst_SecurityAgency.Delete
   RsMst_SecurityAgency.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_SecurityAgency.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_SecurityAgency.MoveNext
   If RsMst_SecurityAgency.EOF() Then
      RsMst_SecurityAgency.MoveLast
   End If
   Call Indata
   TmpRs.Close
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_SecurityAgency.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame1.Enabled = True
Call GButtonLock(Me, False)
TxtAgencyName.Locked = False
TxtContactPerson.Locked = False
TxtAddress.Locked = False
TxtMobile.Locked = False
TxtTelephone2.Locked = False
TxtRemarks.Locked = False
TxtFlag.Visible = False
CmbFlag.Visible = True
CmbFlag.Text = TxtFlag

CmbTDSID.Visible = True
CmdTDSID.Visible = True
TxtTDSID.Visible = False
CmbTDSID.Text = TxtTDSID
CmbBankName.Visible = True
CmbBranch.Visible = True
''--- 10 Aug 2017 Ullhas Patil
TxtTANNO.Locked = False
''----------------------------
CmbBankName.Text = TxtBankName
CmbBranch.Text = TxtBranchName
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_SecurityAgency.MoveFirst
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
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If

Grid_Head
TxtSearchAgency.Text = ""
LblTotalRec.Caption = "-"

Call TableMst_SecurityAgency("  Order by SecurityAgencyID ")
If RsMst_SecurityAgency.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If

Call Indata
Call GButtonLock(Me, True)

End Sub

Public Sub Indata()
tmp = ""
CmbFlag.Visible = False
TxtFlag.Visible = True
TxtAgencyName.Locked = True
TxtContactPerson.Locked = True
TxtAddress.Locked = True
TxtMobile.Locked = True
TxtTelephone2.Locked = True
TxtRemarks.Locked = True

TxtTDSID.Locked = True
TxtBankName.Locked = True
TxtBranchName.Locked = True
CmbBankName.Visible = False
CmbBranch.Visible = False
CmbTDSID.Visible = False
CmdTDSID.Visible = False
TxtTDSID.Visible = True
TxtBankName.Visible = True
TxtBranchName.Visible = True
''--- 10 Aug 2017 Ullhas Patil
TxtTANNO = ""
TxtTANNO.Locked = True
''----------------------------
TxtAgencyID = RsMst_SecurityAgency!SecurityAgencyID

tmp = " Select MSA.SecurityAgencyID,MSA.AgencyName,MSA.ContactPerson,MSA.Address,MSA.Mobile,MSA.Telephone,"
tmp = tmp & " MB.BankId,MB.BankName,MBR.BranchID,MBR.BranchName,MSA.BankAcNo,MSA.Remarks,MSA.IFSCCode,MSA.Flag,MTD.TDSID,MTD.Description,"
tmp = tmp & " MSA.CreatedBy, MSA.UpdatedBy, MSA.CreatedDate, MSA.UpdatedDate, MSA.TANNO"
tmp = tmp & " from Mst_SecurityAgency as MSA"
tmp = tmp & " Left join Mst_TDS as MTD on MSA.TDSID=MTD.TDSID"
tmp = tmp & " Left join Mst_Bank as MB on MSA.BankId=MB.BankId"
tmp = tmp & " Left join Mst_Branch as MBR on MSA.BranchId=MBR.BranchId"
tmp = tmp & " where MSA.SecurityAgencyID=" & TxtAgencyID & ""

Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
    TxtAgencyID = TmpRs1!SecurityAgencyID
    TxtAgencyName = TmpRs1!AgencyName
    TxtContactPerson = IIf(IsNull(TmpRs1!ContactPerson), "", TmpRs1!ContactPerson)
    TxtAddress = IIf(IsNull(TmpRs1!Address), "", TmpRs1!Address)
    TxtMobile = IIf(IsNull(TmpRs1!Mobile), "", TmpRs1!Mobile)
    TxtTelephone2 = IIf(IsNull(TmpRs1!Telephone), "", TmpRs1!Telephone)
    mBankID = IIf(IsNull(TmpRs1!BankID), "", TmpRs1!BankID)
    mBranchID = IIf(IsNull(TmpRs1!BranchID), "", TmpRs1!BranchID)
    TxtBankName = IIf(IsNull(TmpRs1!BankName), "", TmpRs1!BankName)
    TxtBranchName = IIf(IsNull(TmpRs1!BranchName), "", TmpRs1!BranchName)
    TxtAccNo = IIf(IsNull(TmpRs1!BankAcNo), "", TmpRs1!BankAcNo)
    TxtRTGSNo = IIf(IsNull(TmpRs1!IFSCCode), "", TmpRs1!IFSCCode)
    TxtRemarks = IIf(IsNull(TmpRs1!Remarks), "", TmpRs1!Remarks)
    mFlag = TmpRs1!Flag
    
    If mFlag = "Y" Then
       TxtFlag = "Active"
    Else
       TxtFlag = "De-Active"
    End If
    
    mTDSID = IIf(IsNull(TmpRs1!TDSID), 0, TmpRs1!TDSID)
    TxtTDSID = IIf(IsNull(TmpRs1!Description), 0, TmpRs1!Description)
    ''--- 10 Aug 2017 Ullhas Patil
    TxtTANNO = IIf(IsNull(TmpRs1!TANNO), "", TmpRs1!TANNO)
    ''----------------------------
    'If mTDSID <> 0 Then
    '   TxtTDSID.Text = GetTDSDescription(RsMst_SecurityAgency!TDSID)
    'Else
    '   TxtTDSID.Text = ""
    'End If
    
    TxtCreated = IIf(IsNull(TmpRs1!CreatedBy), "", TmpRs1!CreatedBy) & " :- " & IIf(IsNull(TmpRs1!CreatedDate), "", TmpRs1!CreatedDate)
    TxtUpdated = IIf(IsNull(TmpRs1!UpdatedBy), "", TmpRs1!UpdatedBy) & " :- " & IIf(IsNull(TmpRs1!UpdatedDate), "", TmpRs1!UpdatedDate)
End If

Frame1.Enabled = False
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_SecurityAgency.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_SecurityAgency.MoveFirst

If RsMst_SecurityAgency.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "SecurityAgencyID = " & Val(mCaId) & ""
If RsMst_SecurityAgency.EOF Then
   RsMst_SecurityAgency.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_SecurityAgency.MoveNext
If RsMst_SecurityAgency.EOF Then
   RsMst_SecurityAgency.MoveLast
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
RsMst_SecurityAgency.MovePrevious
If RsMst_SecurityAgency.BOF Then
   RsMst_SecurityAgency.MoveFirst
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
If TxtAgencyName = "" Then
   MsgBox "Blank Agency Name not Allowed"
   TxtAgencyName.SetFocus
   Exit Sub
End If
If CmbFlag.Text = "" Then
   MsgBox "Blank Status Not Allowed"
   CmbFlag.SetFocus
   Exit Sub
End If
If CmbTDSID.Text = "" Then
   MsgBox "Blank TDS Description Not Allowed!!!"
   CmbTDSID.SetFocus
   Exit Sub
End If
''--- 10 Aug 2017 Ullhas Patil
''If TxtTANNO.Text = "" Then
''   MsgBox "Blank TAN No Allowed !!! "
''   TxtTANNO.SetFocus
''   Exit Sub
''Else
If Trim(TxtTANNO.Text) <> "" Then
   If ValidateTANNo(UCase(TxtTANNO.Text)) = False Then
      MsgBox "Invalid TAN No !!!"
      TxtTANNO.SetFocus
      Exit Sub
   End If
End If
''----------------------------
If Edit_Flg = "A" Then
   If RsMst_SecurityAgency.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_SecurityAgency.MoveFirst
        RsMst_SecurityAgency.Find "AgencyName = '" & TxtAgencyName & "'"
        If Not RsMst_SecurityAgency.EOF Then
           MsgBox "Duplicate Agency Name not Allowed"
           Exit Sub
        End If
        ''-- Alter for duplicate TAN No 10 Aug 2017 Ullhas Patil
        RsMst_SecurityAgency.Filter = ""
        RsMst_SecurityAgency.MoveFirst
        If Trim(TxtTANNO.Text) <> "" Then
        RsMst_SecurityAgency.Find ("TANNO ='" & Trim(TxtTANNO.Text) & "'")
        If Not RsMst_SecurityAgency.EOF Then
          Select Case MsgBox("Duplicate TAN No Found with another Security Agency, Do you want to save the Security Agency with same TAN No?", vbYesNo)
              Case vbNo
                  TxtTANNO.SetFocus
                  Exit Sub
          End Select
        End If
        End If
        ''----------------------------
        RsMst_SecurityAgency.MoveLast
        tmp = "select max(SecurityAgencyID) from Mst_SecurityAgency"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
    End If
    RsMst_SecurityAgency.AddNew
    RsMst_SecurityAgency!SecurityAgencyID = tmp
    TxtAgencyID = tmp
Else
    RsMst_SecurityAgency.MoveFirst
    RsMst_SecurityAgency.Find "AgencyName = '" & TxtAgencyName & "'"
    If Not RsMst_SecurityAgency.EOF Then
       If RsMst_SecurityAgency!SecurityAgencyID <> TxtAgencyID Then
          MsgBox "Duplicate Agency Name not Allowed"
          Exit Sub
       End If
    End If
   ''-- Alter for duplicate TAN No 09 Aug 2017 Ullhas Patil
   RsMst_SecurityAgency.Filter = ""
   RsMst_SecurityAgency.MoveFirst
   If Trim(TxtTANNO.Text) <> "" Then
   RsMst_SecurityAgency.Find ("TANNO ='" & Trim(TxtTANNO.Text) & "'")
   If Not RsMst_SecurityAgency.EOF Then
    If RsMst_SecurityAgency!SecurityAgencyID <> TxtAgencyID.Text Then
      Select Case MsgBox("Duplicate TAN No Found with another Security Agency, Do you want to save the Security Agency with same TAN No?", vbYesNo)
          Case vbNo
              TxtTANNO.SetFocus
              Exit Sub
      End Select
    End If
   End If
   End If
   ''----------------------------
   RsMst_SecurityAgency.MoveFirst
   RsMst_SecurityAgency.Find "SecurityAgencyID = " & TxtAgencyID & ""
End If
RsMst_SecurityAgency!AgencyName = TxtAgencyName
RsMst_SecurityAgency!ContactPerson = TxtContactPerson
RsMst_SecurityAgency!Address = TxtAddress
RsMst_SecurityAgency!Mobile = TxtMobile
RsMst_SecurityAgency!Telephone = TxtTelephone2
RsMst_SecurityAgency!Flag = mFlag
RsMst_SecurityAgency!Remarks = TxtRemarks
RsMst_SecurityAgency!TDSID = mTDSID
''--- 10 Aug 2017 Ullhas Patil
RsMst_SecurityAgency!TANNO = TxtTANNO
''----------------------------
If mBankID <> "" And mBranchID <> "" Then
    RsMst_SecurityAgency!BankID = mBankID
    RsMst_SecurityAgency!BranchID = mBranchID
    RsMst_SecurityAgency!BankAcNo = TxtAccNo
    RsMst_SecurityAgency!IFSCCode = TxtRTGSNo
Else
    RsMst_SecurityAgency!BankID = Null
    RsMst_SecurityAgency!BranchID = Null
    RsMst_SecurityAgency!BankAcNo = Null
    RsMst_SecurityAgency!IFSCCode = Null
End If

If IsNull(RsMst_SecurityAgency!CreatedBy) = True Or RsMst_SecurityAgency!CreatedBy = "" Then
   RsMst_SecurityAgency!CreatedBy = Gen_UserLoginID
   RsMst_SecurityAgency!CreatedDate = Now
Else
   RsMst_SecurityAgency!UpdatedBy = Gen_UserLoginID
   RsMst_SecurityAgency!UpdatedDate = Now
End If

RsMst_SecurityAgency.Update
RsMst_SecurityAgency.Requery
RsMst_SecurityAgency.MoveFirst
RsMst_SecurityAgency.Find "SecurityAgencyID = " & TxtAgencyID & ""
If RsMst_SecurityAgency.EOF Then
   RsMst_SecurityAgency.MoveFirst
End If
Call GButtonLock(Me, True)
TxtAgencyID.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_SecurityAgency.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_SecurityAgency.EOF Then
      RsMst_SecurityAgency.MoveFirst
   End If
   Call Indata

   Exit Sub
End If


Frame0.Height = Me.Height - 700
Frame0.Width = Me.Width - 500
Frame0.Left = 120
Frame0.Top = 120

Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000

''===========comment on 7 july 2012 for not clearing search grid=============
'Grid_Head
'TxtSearchAgency.Text = ""
'LblTotalRec.Caption = "-"
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAddress")
End If
End Sub

Private Sub TxtAgencyID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAgencyID")
End If
End Sub

Private Sub TxtAgencyID_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtAgencyID)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtAgencyID.SetFocus
End If
End Sub

Private Sub TxtAgencyName_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtAgencyName")
End If
End Sub

Private Sub TxtAgencyName_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtAgencyName)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtAgencyName.SetFocus
End If
End Sub

Private Sub TxtContactPerson_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtContactPerson")
End If
End Sub
Private Sub TxtMobile_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtMobile")
End If
End Sub
Private Sub TxtRemarks_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtRemarks")
End If
End Sub

Private Sub TxtTANNo_LostFocus()
If TxtTANNO.Locked = False Then
    If Trim(TxtTANNO.Text) <> "" Then
       TxtTANNO = UCase(TxtTANNO)
       If ValidateTANNo(TxtTANNO) = False Then
        MsgBox "Invalid TAN No"
        TxtTANNO.SetFocus
        Exit Sub
       End If
    End If
End If
End Sub

Private Sub TxtTelephone2_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtTelephone2")
End If
End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 8
MSHFlexGrid1.RowHeight(0) = 500
MSHFlexGrid1.AllowUserResizing = flexResizeBoth

MSHFlexGrid1.TextMatrix(0, 0) = "Agency ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Agency Name"
MSHFlexGrid1.TextMatrix(0, 2) = "TDS Description"
MSHFlexGrid1.TextMatrix(0, 3) = "Bank Name"
MSHFlexGrid1.TextMatrix(0, 4) = "Branch Name"
MSHFlexGrid1.TextMatrix(0, 5) = "Account No"
MSHFlexGrid1.TextMatrix(0, 6) = "RTGS No"
MSHFlexGrid1.TextMatrix(0, 7) = "TAN No"
MSHFlexGrid1.ColWidth(0) = 900
MSHFlexGrid1.ColWidth(1) = 4000
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1500
MSHFlexGrid1.ColWidth(6) = 1500
MSHFlexGrid1.ColWidth(7) = 1800

MSHFlexGrid1.WordWrap = True
End Sub

Private Sub CmbTDSID_GotFocus()
tmp = CmbTDSID.Text
Call TableMst_TDS
CmbTDSID.Clear
If RsMst_TDS.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_TDS.RecordCount
    CmbTDSID.AddItem RsMst_TDS!Description
    RsMst_TDS.MoveNext
Next
CmbTDSID.Text = tmp
End Sub

Private Sub CmbTDSID_LostFocus()
If CmbTDSID.Text = "" Then
   Exit Sub
End If
RsMst_TDS.MoveFirst
RsMst_TDS.Find "Description = '" & CmbTDSID.Text & "'"
If RsMst_TDS.EOF Then
   MsgBox "Invalid selection "
   CmbTDSID.SetFocus
   Exit Sub
End If
mTDSID = RsMst_TDS!TDSID
End Sub
