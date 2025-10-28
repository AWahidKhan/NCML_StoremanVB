VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTxn_WHRRevalidation 
   Caption         =   "WHR Revalidation"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.Frame ButtonFrm 
      Height          =   1100
      Left            =   120
      TabIndex        =   108
      Top             =   8760
      Width           =   15015
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
         Height          =   765
         Left            =   9000
         TabIndex        =   110
         Top             =   240
         Width           =   2595
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
         Height          =   765
         Left            =   12195
         TabIndex        =   109
         Top             =   240
         Width           =   2595
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.TextBox TxtReasonForRevalidation 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1380
         Left            =   5160
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   3960
         Width           =   9735
      End
      Begin VB.TextBox TxtWHRNoNew 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtWHRNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1680
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TxtLocation 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox TxtCMPName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox TxtGodownName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox TxtGodownAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   975
         Width           =   13215
      End
      Begin VB.TextBox TxtClient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1350
         Width           =   13215
      End
      Begin VB.TextBox TxtCommodity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1730
         Width           =   2895
      End
      Begin VB.TextBox TxtCommodityCondition 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1730
         Width           =   3735
      End
      Begin VB.TextBox TxtPackagingMarks 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1730
         Width           =   3735
      End
      Begin VB.TextBox TxtInsuranceCo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2100
         Width           =   7695
      End
      Begin VB.TextBox TxtAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2460
         Width           =   3720
      End
      Begin VB.CheckBox ChkSaidToContain 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   12840
         TabIndex        =   11
         Top             =   3570
         Width           =   225
      End
      Begin VB.CheckBox ChkEstimatedWgt 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   14640
         TabIndex        =   10
         Top             =   3570
         Width           =   210
      End
      Begin VB.TextBox TxtRisks 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2820
         Width           =   13215
      End
      Begin VB.TextBox TxtStorageCharges 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3180
         Width           =   2040
      End
      Begin VB.TextBox TxtStorageAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   300
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3180
         Width           =   1680
      End
      Begin VB.TextBox TxtTotalWHRQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3555
         Width           =   1815
      End
      Begin VB.TextBox TxtValueofCommodity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Height          =   300
         Left            =   6240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3555
         Width           =   1680
      End
      Begin VB.TextBox TxtAgreeDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   13680
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3960
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox TxtRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   3555
         Width           =   2040
      End
      Begin VB.TextBox TxtAgreementEndDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   13200
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3180
         Width           =   1650
      End
      Begin VB.TextBox TxtPolicyNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2460
         Width           =   7695
      End
      Begin MSComCtl2.DTPicker TxtWHRDate 
         Height          =   300
         Left            =   8160
         TabIndex        =   24
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   109117441
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtStorageFrom 
         Height          =   300
         Left            =   10680
         TabIndex        =   25
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   109117441
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtStorageTo 
         Height          =   300
         Left            =   13560
         TabIndex        =   26
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   109117441
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtValidFrom 
         Height          =   300
         Left            =   11160
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2100
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   109117441
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtValidTo 
         Height          =   300
         Left            =   13560
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2100
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   109117441
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtGradeValidUpto 
         Height          =   300
         Left            =   9360
         TabIndex        =   29
         Top             =   3180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   109117441
         CurrentDate     =   39884
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3255
         Left            =   120
         TabIndex        =   30
         Top             =   5400
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   5741
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "GSL Detail"
         TabPicture(0)   =   "FrmTxn_WHRRevalidation.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Pledge / Other Details"
         TabPicture(1)   =   "FrmTxn_WHRRevalidation.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame6"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "CITF Details"
         TabPicture(2)   =   "FrmTxn_WHRRevalidation.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "MSHFlexGrid2"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "New WHR Details / Log Details"
         TabPicture(3)   =   "FrmTxn_WHRRevalidation.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label60"
         Tab(3).Control(1)=   "Label59"
         Tab(3).Control(2)=   "Label58"
         Tab(3).Control(3)=   "Label57"
         Tab(3).Control(4)=   "Label56"
         Tab(3).Control(5)=   "Label52"
         Tab(3).Control(6)=   "Label48"
         Tab(3).Control(7)=   "Label33"
         Tab(3).Control(8)=   "Label17"
         Tab(3).Control(9)=   "Label16"
         Tab(3).Control(10)=   "TxtNewWHRNoDate"
         Tab(3).Control(11)=   "TxtNewWHRBags"
         Tab(3).Control(11).Enabled=   0   'False
         Tab(3).Control(12)=   "TxtNewWHRStorageCharges"
         Tab(3).Control(12).Enabled=   0   'False
         Tab(3).Control(13)=   "TxtNewWHRQty"
         Tab(3).Control(13).Enabled=   0   'False
         Tab(3).Control(14)=   "TxtNewWHRValueofCommodity"
         Tab(3).Control(14).Enabled=   0   'False
         Tab(3).Control(15)=   "TxtNewWHRNo"
         Tab(3).Control(15).Enabled=   0   'False
         Tab(3).Control(16)=   "TxtNewCommodityRate"
         Tab(3).Control(16).Enabled=   0   'False
         Tab(3).Control(17)=   "TxtNewStorageRate"
         Tab(3).Control(17).Enabled=   0   'False
         Tab(3).Control(18)=   "TxtUpdated"
         Tab(3).Control(18).Enabled=   0   'False
         Tab(3).Control(19)=   "TxtCreated"
         Tab(3).Control(19).Enabled=   0   'False
         Tab(3).Control(20)=   "TxtNewWHRID"
         Tab(3).Control(20).Enabled=   0   'False
         Tab(3).ControlCount=   21
         Begin VB.TextBox TxtNewWHRID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   -72360
            Locked          =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtCreated 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   -73920
            Locked          =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1920
            Width           =   6120
         End
         Begin VB.TextBox TxtUpdated 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   -66480
            Locked          =   -1  'True
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   1920
            Width           =   6120
         End
         Begin VB.Frame Frame6 
            Height          =   2505
            Left            =   -74880
            TabIndex        =   38
            Top             =   360
            Width           =   14535
            Begin VB.TextBox TxtCARDate 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   300
               Left            =   12360
               Locked          =   -1  'True
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   2100
               Width           =   2040
            End
            Begin VB.TextBox TxtCARNo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   300
               Left            =   9240
               Locked          =   -1  'True
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   2100
               Width           =   1680
            End
            Begin VB.TextBox TxtPledgeBranchName 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   300
               Left            =   9240
               Locked          =   -1  'True
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   1350
               Width           =   5160
            End
            Begin VB.TextBox TxtPledgeBankName 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   300
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   1350
               Width           =   4800
            End
            Begin VB.TextBox TxtPledgeBags 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   300
               Left            =   5640
               Locked          =   -1  'True
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   2100
               Width           =   1800
            End
            Begin VB.TextBox TxtPledgeQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   300
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   2100
               Width           =   1920
            End
            Begin VB.TextBox TxtPledgeDate 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   300
               Left            =   9240
               Locked          =   -1  'True
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   1725
               Width           =   1680
            End
            Begin VB.TextBox TxtPledgeNo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   300
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   1725
               Width           =   4800
            End
            Begin VB.TextBox TxtAssayerCerti 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   300
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   955
               Width           =   11760
            End
            Begin VB.TextBox TxtAssayer 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   300
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   600
               Width           =   11760
            End
            Begin VB.TextBox TxtWeigher 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   300
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   240
               Width           =   11760
            End
            Begin VB.Label Label49 
               Caption         =   "CAR Date"
               Height          =   240
               Left            =   11160
               TabIndex        =   60
               Top             =   2130
               Width           =   855
            End
            Begin VB.Label Label45 
               Caption         =   "CAR No"
               Height          =   240
               Left            =   7560
               TabIndex        =   59
               Top             =   2130
               Width           =   1215
            End
            Begin VB.Label Label44 
               Caption         =   "Pledge Branch Name"
               Height          =   270
               Left            =   7560
               TabIndex        =   58
               Top             =   1410
               Width           =   2055
            End
            Begin VB.Label Label39 
               Caption         =   "Pledge Bank Name"
               Height          =   240
               Left            =   120
               TabIndex        =   57
               Top             =   1425
               Width           =   1695
            End
            Begin VB.Label Label37 
               Caption         =   "Pledge Bags"
               Height          =   240
               Left            =   4680
               TabIndex        =   56
               Top             =   2130
               Width           =   975
            End
            Begin VB.Label Label35 
               Caption         =   "Pledge Quantity"
               Height          =   240
               Left            =   120
               TabIndex        =   55
               Top             =   2130
               Width           =   1695
            End
            Begin VB.Label Label34 
               Caption         =   "Pledge Date"
               Height          =   240
               Left            =   7560
               TabIndex        =   54
               Top             =   1755
               Width           =   1215
            End
            Begin VB.Label Label32 
               Caption         =   "Pledge ID"
               Height          =   240
               Left            =   120
               TabIndex        =   53
               Top             =   1755
               Width           =   1695
            End
            Begin VB.Label Label13 
               Caption         =   "Grade as per Assayer’s Certificate"
               Height          =   180
               Left            =   120
               TabIndex        =   52
               Top             =   1015
               Width           =   2700
            End
            Begin VB.Label Label14 
               Caption         =   "Name and License No. of Weigher"
               Height          =   180
               Left            =   120
               TabIndex        =   51
               Top             =   270
               Width           =   2700
            End
            Begin VB.Label Label15 
               Caption         =   "Name and License No. of Assayer"
               Height          =   180
               Left            =   120
               TabIndex        =   50
               Top             =   660
               Width           =   2700
            End
         End
         Begin VB.TextBox TxtNewStorageRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   -72360
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   840
            Width           =   2040
         End
         Begin VB.TextBox TxtNewCommodityRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   -72360
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2040
         End
         Begin VB.TextBox TxtNewWHRNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   -72360
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   480
            Width           =   2040
         End
         Begin VB.TextBox TxtNewWHRValueofCommodity 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   300
            Left            =   -68400
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1215
            Width           =   1680
         End
         Begin VB.TextBox TxtNewWHRQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   -65280
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   495
            Width           =   1815
         End
         Begin VB.TextBox TxtNewWHRStorageCharges 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   300
            Left            =   -68400
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   840
            Width           =   1680
         End
         Begin VB.TextBox TxtNewWHRBags 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   300
            Left            =   -65280
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   840
            Width           =   1815
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   2880
            Left            =   -74880
            TabIndex        =   63
            Top             =   480
            Width           =   14340
            _ExtentX        =   25294
            _ExtentY        =   5080
            _Version        =   393216
            Rows            =   3
            FixedRows       =   2
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   2745
            Left            =   120
            TabIndex        =   64
            Top             =   405
            Width           =   14580
            _ExtentX        =   25718
            _ExtentY        =   4842
            _Version        =   393216
            Rows            =   3
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3105
            Left            =   -74880
            TabIndex        =   65
            Top             =   405
            Width           =   14580
            _ExtentX        =   25718
            _ExtentY        =   5477
            _Version        =   393216
            Rows            =   3
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
         Begin MSComCtl2.DTPicker TxtNewWHRNoDate 
            Height          =   300
            Left            =   -68400
            TabIndex        =   66
            Top             =   480
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   109117441
            CurrentDate     =   39884
         End
         Begin VB.Label Label46 
            Caption         =   "Created  :"
            Height          =   240
            Left            =   -74760
            TabIndex        =   79
            Top             =   900
            Width           =   1695
         End
         Begin VB.Label Label47 
            Caption         =   "Updated :"
            Height          =   255
            Left            =   -67560
            TabIndex        =   78
            Top             =   915
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "Created  :"
            Height          =   240
            Left            =   -74880
            TabIndex        =   77
            Top             =   1980
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "Updated :"
            Height          =   255
            Left            =   -67440
            TabIndex        =   76
            Top             =   1980
            Width           =   855
         End
         Begin VB.Label Label33 
            Caption         =   "Storage Rate Per Month Per Bag"
            Height          =   225
            Left            =   -74880
            TabIndex        =   75
            Top             =   885
            Width           =   2940
         End
         Begin VB.Label Label48 
            Caption         =   "WHR Date"
            Height          =   255
            Left            =   -70200
            TabIndex        =   74
            Top             =   510
            Width           =   1380
         End
         Begin VB.Label Label52 
            Caption         =   "Commodity Price (Rate Per MT)"
            Height          =   210
            Left            =   -74880
            TabIndex        =   73
            Top             =   1245
            Width           =   2820
         End
         Begin VB.Label Label56 
            Caption         =   "New WHR No"
            Height          =   225
            Left            =   -74880
            TabIndex        =   72
            Top             =   525
            Width           =   2940
         End
         Begin VB.Label Label57 
            Caption         =   "Value of Commodity"
            Height          =   255
            Left            =   -70200
            TabIndex        =   71
            Top             =   1245
            Width           =   1680
         End
         Begin VB.Label Label58 
            Caption         =   "Total WHR Qty"
            Height          =   255
            Left            =   -66600
            TabIndex        =   70
            Top             =   525
            Width           =   1500
         End
         Begin VB.Label Label59 
            Caption         =   "Storage Charge Amt"
            Height          =   255
            Left            =   -70200
            TabIndex        =   69
            Top             =   870
            Width           =   1560
         End
         Begin VB.Label Label60 
            Caption         =   "Total WHR Bags"
            Height          =   255
            Left            =   -66600
            TabIndex        =   68
            Top             =   870
            Width           =   1500
         End
      End
      Begin MSComCtl2.DTPicker TxtNewStorageTo 
         Height          =   300
         Left            =   2520
         TabIndex        =   111
         Top             =   4440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Format          =   109117441
         CurrentDate     =   39884
      End
      Begin VB.Label Label23 
         Caption         =   "Reason for Revalidation (Max 5000 Characters)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4080
         TabIndex        =   113
         Top             =   4223
         Width           =   1455
      End
      Begin VB.Label LblWHRNo 
         Caption         =   "WHR No"
         Height          =   250
         Left            =   120
         TabIndex        =   107
         Top             =   265
         Width           =   900
      End
      Begin VB.Label LblWHRDate 
         Caption         =   "WHR Date"
         Height          =   250
         Left            =   7320
         TabIndex        =   106
         Top             =   265
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Condition of Commodities"
         Height          =   420
         Left            =   4680
         TabIndex        =   105
         Top             =   1670
         Width           =   1260
      End
      Begin VB.Label Label10 
         Caption         =   "Storage upto Date"
         Height          =   250
         Left            =   12135
         TabIndex        =   104
         Top             =   265
         Width           =   1320
      End
      Begin VB.Label Label9 
         Caption         =   "Date of Deposit "
         Height          =   250
         Left            =   9480
         TabIndex        =   103
         Top             =   265
         Width           =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "Location"
         Height          =   250
         Left            =   120
         TabIndex        =   102
         Top             =   625
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "CMP Name"
         Height          =   250
         Left            =   4680
         TabIndex        =   101
         Top             =   625
         Width           =   1140
      End
      Begin VB.Label Label5 
         Caption         =   "Godown No"
         Height          =   250
         Left            =   9480
         TabIndex        =   100
         Top             =   625
         Width           =   1620
      End
      Begin VB.Label Label6 
         Caption         =   "Godown Address"
         Height          =   250
         Left            =   120
         TabIndex        =   99
         Top             =   1000
         Width           =   1500
      End
      Begin VB.Label Label7 
         Caption         =   "Client Name"
         Height          =   250
         Left            =   120
         TabIndex        =   98
         Top             =   1375
         Width           =   1500
      End
      Begin VB.Label LblWHRCommodity 
         Caption         =   "Commodity"
         Height          =   250
         Left            =   120
         TabIndex        =   97
         Top             =   1755
         Width           =   1140
      End
      Begin VB.Label Label8 
         Caption         =   "Private Marks on Packaging"
         Height          =   420
         Left            =   9480
         TabIndex        =   96
         Top             =   1670
         Width           =   1620
      End
      Begin VB.Label Label38 
         Caption         =   "Insurance Co."
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   2130
         Width           =   1575
      End
      Begin VB.Label Label40 
         Caption         =   "Valid To"
         Height          =   255
         Left            =   12600
         TabIndex        =   94
         Top             =   2130
         Width           =   735
      End
      Begin VB.Label Label41 
         Caption         =   "Valid From"
         Height          =   255
         Left            =   9480
         TabIndex        =   93
         Top             =   2130
         Width           =   1620
      End
      Begin VB.Label Label42 
         Caption         =   "Amount"
         Height          =   255
         Left            =   9480
         TabIndex        =   92
         Top             =   2490
         Width           =   780
      End
      Begin VB.Label Label21 
         Caption         =   "Said To Contain"
         Height          =   255
         Left            =   11520
         TabIndex        =   91
         Top             =   3570
         Width           =   1260
      End
      Begin VB.Label Label22 
         Caption         =   "Estimated Weight"
         Height          =   255
         Left            =   13200
         TabIndex        =   90
         Top             =   3570
         Width           =   1335
      End
      Begin VB.Label Label43 
         Caption         =   "Risks"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   2850
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Storage Rate Per Month Per Bag"
         Height          =   225
         Left            =   120
         TabIndex        =   88
         Top             =   3225
         Width           =   2340
      End
      Begin VB.Label Label11 
         Caption         =   "Storage Charge Amt"
         Height          =   255
         Left            =   4680
         TabIndex        =   87
         Top             =   3210
         Width           =   1560
      End
      Begin VB.Label Label12 
         Caption         =   "Grade Valid Till"
         Height          =   255
         Left            =   8040
         TabIndex        =   86
         Top             =   3210
         Width           =   1470
      End
      Begin VB.Label LblTotalWHRQty 
         Caption         =   "Total WHR Qty"
         Height          =   255
         Left            =   8040
         TabIndex        =   85
         Top             =   3585
         Width           =   1500
      End
      Begin VB.Label Label20 
         Caption         =   "Value of Commodity"
         Height          =   255
         Left            =   4680
         TabIndex        =   84
         Top             =   3585
         Width           =   1680
      End
      Begin VB.Label Label19 
         Caption         =   "Commodity Price (Rate Per MT)"
         Height          =   210
         Left            =   120
         TabIndex        =   83
         Top             =   3600
         Width           =   2340
      End
      Begin VB.Label Label50 
         Caption         =   "Agreement End Date"
         Height          =   255
         Left            =   11520
         TabIndex        =   82
         Top             =   3210
         Width           =   1620
      End
      Begin VB.Label LblClientName 
         Caption         =   "New Storage Upto Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   81
         Top             =   4440
         Width           =   2340
      End
      Begin VB.Label Label24 
         Caption         =   "Policy No"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   2490
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmTxn_WHRRevalidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Dim mCMPID, mWHRNo, mLocationID, mGodownID, mSCMFlag, mClientIDRow, mNewClientIDRow, mExchangeTypeID, mCommodityCondition, mCommodityID, mPolicyNo, mSM_Prefix, mNewWHRDate, mWHRTransferDate As Variant
Dim mCoverNoteNo, mLotWiseWHR, mConvertedClient, mStoregeRatePerMonthPerBag, mPledgeID As Variant

Dim mGodownYearlyBillingFlg, mCommodityYearlyBillingFlg, mRate, mRejInvIDs As Variant

Dim mSFromDate, mSToDate, mSLocationID, msSM_prefix, mSCMPID, mSClientRowID, mSNewClientRowID, mSWHRNo, mSNewWHRNo As Variant
Dim mWHRCheck, mAddNew As Boolean
Dim mSCMPledgeID As Variant

Dim mNewStorageDate As Variant
Dim mSMonth, mSYear, mSTotalDays As Variant
Dim mNCMLFinance As Variant

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub TxtNewStorageTo_LostFocus()
mSMonth = Month(TxtNewStorageTo.Value)
mSYear = Year(TxtNewStorageTo.Value)
mSTotalDays = 20
TxtNewStorageTo.Value = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
If TxtAgreementEndDate <> "" Then
   If CDate(TxtAgreementEndDate) < CDate(TxtNewStorageTo.Value) Then
      TxtNewStorageTo.Value = CDate(TxtAgreementEndDate)
   End If
End If

End Sub

Private Sub TxtWHRNo_LostFocus()

If TxtWHRNo.Text <> "" Then
Dim strSplitString() As String
   strSplitString = Split(TxtWHRNo.Text, "-")
   tmp1 = UBound(strSplitString)
   If tmp1 <> 1 Then
      MsgBox "Invalid WHR No selection"
      TxtWHRNo.SetFocus
      Exit Sub
   End If
   mWHRNo = Val(strSplitString(0)) * 1000000 + Val(strSplitString(1))
   mCMPID = Val(strSplitString(0))
   TxtWHRNoNew = Val(mWHRNo)
   Call WHRDetailDisplay
Else
   Call ClearForm
End If
End Sub


''-- Clear all fields and related with WHR No
Private Sub ClearForm()


TxtReasonForRevalidation = ""
TxtNewWHRID = ""
TxtWHRDate.Value = Date
TxtStorageFrom.Value = Date
TxtStorageTo.Value = Date
TxtNewStorageTo.Value = Date
TxtLocation = ""
TxtCMPName = ""
TxtGodownName = ""
TxtGodownAddress = ""
TxtClient = ""
TxtCommodity = ""
TxtCommodityCondition = ""
TxtPackagingMarks = ""
TxtInsuranceCo = ""
TxtPolicyNo = ""
TxtValidFrom.Value = Date
TxtValidTo.Value = Date
TxtRisks = ""
TxtAmount = ""
TxtStorageCharges = ""
TxtStorageAmount = ""
TxtGradeValidUpto.Value = Date
ChkSaidToContain.Value = vbUnchecked
ChkEstimatedWgt.Value = vbUnchecked
TxtRate = ""
TxtValueofCommodity = ""
TxtTotalWHRQty = ""
TxtAgreementEndDate = ""
TxtAgreeDate = ""

''-- GSL Details

Call Grid_Head1

''-- Pledge / Other Details
TxtWeigher = ""
TxtAssayer = ""
TxtAssayerCerti = ""
TxtPledgeBankName = ""
TxtPledgeBranchName = ""
TxtPledgeNo = ""
TxtPledgeDate = ""
TxtPledgeQty = ""
TxtPledgeBags = ""
TxtCarNo = ""
TxtCARDate = ""

''-- CITF Details
Call Grid_Head2

''-- Log History Details
TxtCreated = ""
TxtUpdated = ""

''-- New WHR Details
TxtNewWHRID = ""
TxtNewWHRNo = ""
TxtNewStorageRate = ""
TxtNewCommodityRate = ""
TxtNewWHRNoDate.Value = Date
TxtNewWHRStorageCharges = ""
TxtNewWHRValueofCommodity = ""
TxtNewWHRQty = ""
TxtNewWHRBags = ""
Label33.Caption = "Storage Rate Per Month Per Bag"

SSTab1.Tab = 0
End Sub


Private Sub WHRDetailDisplay()


tmp = " SELECT WHRD.WHRNo,WHRD.SM_Prefix,WHRD.WHRDate,WHRD.ConditionID,WHRD.StorageFrom,WHRD.StorageTo,WHRD.Grade,"
tmp = tmp & " WHRD.Weigher,WHRD.Assayer,WHRD.GradeByAssayer,WHRD.GradeValidTill,WHRD.StorageCharges,WHRD.StorageAmount,"
tmp = tmp & " WHRD.ClientID,WHRD.CommodityID,WHRD.CMPID,WHRD.LicenseID,WHRD.WHRNoOfBags,WHRD.WHRQuantity,WHRD.PrintFlag,"
tmp = tmp & " WHRD.Remark,WHRD.Flag,WHRD.Txn_PledgeCreation_ID,WHRD.OldLicenseID,WHRD.GodownID,WHRD.DeliveredQuantity,"
tmp = tmp & " WHRD.DeliveredNoOfBags,WHRD.Rate,WHRD.G_UID,WHRD.HologramNo,WHRD.PolicyNo,WHRD.CoverNoteNo,"
tmp = tmp & " WHRD.SaidToContain , WHRD.EstimatedWeight, WHRD.SRAsPerBag, WHRD.CARNo, WHRD.ConvertedClient, WHRD.SCMFlag,WHRD.CARNo"
tmp = tmp & " ,CMP.CMPName,CMP.LocationID,ML.LocationName,ML.SM_Prefix,MG.GodownNo,MG.AgreementEndDate,MG.Address,MC.ClientName,Mcom.Commodity,isnull(WHRD.LotWiseWHR,0) 'LotWiseWHR',MLc.LicenseNo,MLc.LicenseExpiryDate,MG.YearlyBillingFlg[GodownYearlyBillingFlg],MCom.YearlyBillingFlg[CommodityYearlyBillingFlg] "
tmp = tmp & " FROM Txn_WHR_Detail WHRD"
tmp = tmp & " Inner Join Mst_CMP CMP On WHRD.CMPID = CMP.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Godown MG On WHRD.GodownID = MG.GodownID"
tmp = tmp & " Left Join Mst_License MLc ON MLc.LicenseID = MG.LicenseID"
tmp = tmp & " Inner Join Mst_Client MC On WHRD.ClientID = MC.ClientID"
tmp = tmp & " Inner Join Mst_Commodity MCom On WHRD.CommodityID = MCom.CommodityID"
tmp = tmp & " Where WHRD.WHRNo = " & Val(TxtWHRNoNew) & " And WHRD.WHRNoOfBags - WHRD.DeliveredNoOfBags > 0 And isnull(WHRD.Txn_PledgeCreation_ID,0) <> 0 And isnull(WHRD.CARNo,0) = 0  "

Call TmpTable

Call ClearForm

If TmpRs.RecordCount > 0 Then
        
   TxtWHRDate.Value = IIf(IsNull(TmpRs!WHRDate), "", TmpRs!WHRDate)
   TxtStorageFrom.Value = IIf(IsNull(TmpRs!StorageFrom), "", TmpRs!StorageFrom)
   TxtStorageTo.Value = IIf(IsNull(TmpRs!StorageTo), "", TmpRs!StorageTo)
   mSM_Prefix = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
   mCMPID = IIf(IsNull(TmpRs!CMPID), "", TmpRs!CMPID)
   mLocationID = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
   mGodownID = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
   mSCMFlag = IIf(IsNull(TmpRs!SCMFlag), 0, TmpRs!SCMFlag)
   mClientIDRow = IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID)
   mExchangeTypeID = GetExchangeID("NON NCDEX")
   mLotWiseWHR = IIf(IsNull(TmpRs!LotWiseWHR), 0, TmpRs!LotWiseWHR)
   mCommodityCondition = IIf(IsNull(TmpRs!ConditionID), "", TmpRs!ConditionID)
   mCommodityID = IIf(IsNull(TmpRs!CommodityID), "", TmpRs!CommodityID)
   mPolicyNo = IIf(IsNull(TmpRs!PolicyNo), "", TmpRs!PolicyNo)
   mCoverNoteNo = IIf(IsNull(TmpRs!CoverNoteNo), "", TmpRs!CoverNoteNo)
   mConvertedClient = IIf(IsNull(TmpRs!ConvertedClient), 0, TmpRs!ConvertedClient)
   mPledgeID = IIf(IsNull(TmpRs!Txn_PledgeCreation_ID), 0, TmpRs!Txn_PledgeCreation_ID)
   ''-- Bags = True / MT = False
   mStoregeRatePerMonthPerBag = IIf(IsNull(TmpRs!SRAsPerBag), False, IIf(TmpRs!SRAsPerBag = 0, False, True))
   TxtLocation = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
   TxtCMPName = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
   TxtGodownName = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
   TxtGodownAddress = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
   TxtAgreementEndDate = Format(IIf(IsNull(TmpRs!AgreementEndDate), "", TmpRs!AgreementEndDate), "dd-mmm-yyyy")
   TxtTotalWHRQty = IIf(IsNull(TmpRs!WHRQuantity), 0, TmpRs!WHRQuantity)
   TxtWeigher = IIf(IsNull(TmpRs!Weigher), "", TmpRs!Weigher)
   TxtAssayer = IIf(IsNull(TmpRs!Assayer), "", TmpRs!Assayer)
   TxtAssayerCerti = IIf(IsNull(TmpRs!GradeByAssayer), "", TmpRs!GradeByAssayer)
   TxtGradeValidUpto.Value = IIf(IsNull(TmpRs!GradeValidTill), "", TmpRs!GradeValidTill)
   TxtStorageCharges = IIf(IsNull(TmpRs!StorageCharges), "", TmpRs!StorageCharges)
   TxtStorageAmount = IIf(IsNull(TmpRs!StorageAmount), "", TmpRs!StorageAmount)
   TxtRate = IIf(IsNull(TmpRs!Rate), "", TmpRs!Rate)
   TxtPolicyNo = TmpRs!PolicyNo & " ~ " & TmpRs!CoverNoteNo
   ChkSaidToContain.Value = IIf(IsNull(TmpRs!SaidToContain), 0, TmpRs!SaidToContain)
   ChkEstimatedWgt.Value = IIf(IsNull(TmpRs!EstimatedWeight), 0, TmpRs!EstimatedWeight)
   mLicenseID = IIf(IsNull(TmpRs!LicenseID), "", TmpRs!LicenseID)
   mLicenseNo = IIf(IsNull(TmpRs!LicenseNo), "", TmpRs!LicenseNo)
   mLicenseExpiryDate = IIf(IsNull(TmpRs!LicenseExpiryDate), "", TmpRs!LicenseExpiryDate)
   
   mGodownYearlyBillingFlg = IIf(IsNull(TmpRs!GodownYearlyBillingFlg), 0, TmpRs!GodownYearlyBillingFlg)
   mCommodityYearlyBillingFlg = IIf(IsNull(TmpRs!CommodityYearlyBillingFlg), 0, TmpRs!CommodityYearlyBillingFlg)
   
   If mStoregeRatePerMonthPerBag = True Then
      Label3.Caption = "Storage Rate Per Month Per Bag"
   Else
      Label3.Caption = "Storage Rate Per Month Per MT"
   End If
    
   If mCommodityCondition = 0 Then
      TxtCommodityCondition.Text = "Average"
   ElseIf mCommodityCondition = 1 Then
      TxtCommodityCondition.Text = "Fair"
   ElseIf mCommodityCondition = 2 Then
      TxtCommodityCondition.Text = "Good"
   End If
    
   TxtGrade = IIf(IsNull(TmpRs!Grade), "", TmpRs!Grade)
   TxtClient = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
   TxtCommodity.Text = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
   


   TxtNewStorageTo.Value = DateAdd("M", 3, TxtStorageTo.Value)
   mSMonth = Month(TxtNewStorageTo.Value)
   mSYear = Year(TxtNewStorageTo.Value)
   mSTotalDays = 20
   TxtNewStorageTo.Value = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
   If TxtAgreementEndDate <> "" Then
      If CDate(TxtAgreementEndDate) < CDate(TxtNewStorageTo.Value) Then
         TxtNewStorageTo.Value = CDate(TxtAgreementEndDate)
      End If
   End If
   
   Call TableMst_PolicyDetail("where PolicyNo = '" & mPolicyNo & "' And CoverNoteNo = '" & mCoverNoteNo & "'")
   If RsMst_PolicyDetail.RecordCount > 0 Then
      TxtInsuranceCo = GetInsuranceCoName(RsMst_PolicyDetail!InsuranceCoID)
      TxtValidFrom.Value = RsMst_PolicyDetail!ValidFrom
      TxtValidTo.Value = RsMst_PolicyDetail!ValidTo
      TxtAmount = RsMst_PolicyDetail!Amount
      TxtRisks = RsMst_PolicyDetail!Risks
      RsMst_PolicyDetail.Close
   End If
   Call Grid_Head1
   Call GetGSLData
   
   ''-- Commodity Value ------------
   With MSHFlexGrid1
    If mStoregeRatePerMonthPerBag = True Then
       TxtStorageAmount = Val(.TextMatrix(0, 5)) * Val(TxtStorageCharges)
    Else
       TxtStorageAmount = Format(Val(.TextMatrix(0, 6)) * Val(TxtStorageCharges), "#########0.00")
    End If
    TxtValueofCommodity = Format(Val(.TextMatrix(0, 6)) * Val(TxtRate), "#########0.00")
   End With
   ''------------------------------
    Call CITFDetail
    Call PledgeDetail
    mWHRCheck = True
    SaveCmd.Enabled = True
Else
   mWHRCheck = False
   MsgBox "Either wrong WHR provided or WHR is fully withdrawn. Please check !!!"
   Exit Sub
End If
End Sub


Private Sub GetGSLData()
k = 2
If mLotWiseWHR = 0 Then
   tmp = " Select * from Mst_GSL Where SM_Prefix = '" & mSM_Prefix & "' And ClientID = " & mClientIDRow & " And GodownID = " & mGodownID & " And CommodityID = " & mCommodityID & " And SCMFlag = " & mSCMFlag & " And WHRNO is null And BalanceBags > 0 And WHRBags > 0 And Status not in ('E','Z') And (BalanceBags - PledgeBags) > 0  Order By StackNo,LotNo,Mst_GSL_ID "
Else
   tmp = " Select * from Mst_GSL Where WHRNO = " & Val(TxtWHRNoNew) & " And BalanceBags > 0 And Status not in ('E','Z') And (BalanceBags - PledgeBags) > 0 Order By StackNo,LotNo,Mst_GSL_ID "
End If

Call Tmp2Table
With MSHFlexGrid1
For I = 1 To TmpRs2.RecordCount
    .Rows = .Rows + 1
    .TextMatrix(k, 1) = TmpRs2!StackNo
    .TextMatrix(k, 2) = TmpRs2!LotNo
    .TextMatrix(k, 3) = TmpRs2!BalanceBags - TmpRs2!PledgeBags
    .TextMatrix(k, 4) = TmpRs2!BalanceWeight - TmpRs2!PledgeWeight
    .TextMatrix(k, 5) = TmpRs2!WHRBags
    .TextMatrix(k, 6) = TmpRs2!WHRWeight
    .TextMatrix(k, 7) = IIf(IsNull(TmpRs2!CDTFWeight), 0, TmpRs2!CDTFWeight)
    .TextMatrix(k, 8) = TmpRs2!Mst_GSL_ID
    .TextMatrix(k, 9) = TmpRs2!SM_Prefix
    .TextMatrix(0, 3) = Val(.TextMatrix(0, 3)) + Val(.TextMatrix(k, 3))
    .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) + Val(.TextMatrix(k, 4))
    .TextMatrix(0, 5) = Val(.TextMatrix(0, 5)) + Val(.TextMatrix(k, 5))
    .TextMatrix(0, 6) = Val(.TextMatrix(0, 6)) + Val(.TextMatrix(k, 6))
    .row = k
    .Col = 0
    .CellFontName = "Wingdings"
    .CellFontSize = 12
    .CellAlignment = flexAlignCenterCenter
    .Text = strUnChecked
    k = k + 1
    TmpRs2.MoveNext
Next
End With
End Sub

Private Sub CITFDetail()
    Call Grid_Head2
    
    tmp = "Select 'C'[TransType],ML.LocationName,a.CxTF_DetailsID,b.CxTFDate,a.AdjRRNQty,a.AdjRRNNoofBags,b.CreatedDate "
    tmp = tmp & " From Txn_CxTF_ClientDetail a "
    tmp = tmp & " Inner Join Txn_CxTF_Details b on a.SM_Prefix = b.SM_Prefix And a.CxTF_DetailsID = b.CxTF_DetailsID  And a.txntype = b.txntype "
    tmp = tmp & " Inner Join Mst_CMP MC On b.CMPID = MC.CMPID "
    tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID "
    tmp = tmp & " Where b.BookedSpaceFlag <> 1 And b.exchangetypeid= 1 and b.txntype = 'W' "
    tmp = tmp & " And a.RRNID =  " & TxtWHRNoNew & " "
'    tmp = tmp & " Union all"
'    tmp = tmp & " SELECT 'T',ML.LocationName,WHRTD.WHRTransferID,WHRT.WHRTransferDate,SUM(WHRTD.Qty)Qty,Sum(WHRTD.NoofBags)NoofBags,WHRT.CreatedDate"
'    tmp = tmp & " FROM Txn_WHRTransferDetail WHRTD"
'    tmp = tmp & " Inner Join Txn_WHRTransfer WHRT On WHRTD.WHRTransferID = WHRT.WHRTransferID"
'    tmp = tmp & " Inner Join Txn_WHR_Detail WHRD On WHRT.WHRNO = WHRD.WHRNo And WHRT.SM_Prefix = WHRD.SM_Prefix"
'    tmp = tmp & " Inner Join Mst_CMP CMP On WHRD.CMPID = CMP.CMPID"
'    tmp = tmp & " Inner Join Mst_Location ML On CMP.LocationID = ML.LocationID"
'    tmp = tmp & " Where WHRT.WHRNO = " & TxtWHRNoNew & ""
'    tmp = tmp & " Group By ML.LocationName,WHRTD.WHRTransferID,WHRT.WHRTransferDate,WHRT.CreatedDate"
    Call Tmp2Table

    With MSHFlexGrid2
    If TmpRs2.RecordCount > 0 Then
       .Rows = TmpRs2.RecordCount + 3
       mRow = 2
       For x = 1 To TmpRs2.RecordCount
           
           ''--- C (CITF) / T (WHR Transfer)
           .TextMatrix(mRow, 0) = IIf(IsNull(TmpRs2!TransType), "", IIf(LCase(TmpRs2!TransType) = "c", "CITF", IIf(LCase(TmpRs2!TransType) = "t", "WHR Transfer", "")))
           
           .TextMatrix(mRow, 1) = IIf(IsNull(TmpRs2!Locationname), "", TmpRs2!Locationname)
           .TextMatrix(mRow, 2) = IIf(IsNull(TmpRs2!CxTF_DetailsID), "", TmpRs2!CxTF_DetailsID)
           .TextMatrix(mRow, 3) = IIf(IsNull(TmpRs2!CxTFDate), "", Format(TmpRs2!CxTFDate, "DD-MMM-YYYY"))
           .TextMatrix(mRow, 4) = IIf(IsNull(TmpRs2!AdjRRNQty), "", TmpRs2!AdjRRNQty)
           .TextMatrix(mRow, 5) = IIf(IsNull(TmpRs2!AdjRRNNoofBags), "", TmpRs2!AdjRRNNoofBags)
           .TextMatrix(mRow, 6) = IIf(IsNull(TmpRs2!CreatedDate), "", Format(TmpRs2!CreatedDate, "dd-mmm-yyyy hh:mm AM/PM"))
           
           .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) + Val(.TextMatrix(mRow, 4))
           .TextMatrix(0, 5) = Val(.TextMatrix(0, 5)) + Val(.TextMatrix(mRow, 5))

           mRow = mRow + 1
           TmpRs2.MoveNext
       Next
    End If
    End With
End Sub

Private Sub PledgeDetail()

tmp = "SELECT  TPC.Txn_PledgeCreation_ID,TPC.PledgeDate,isnull(sum(PledgeQuantity),0) PledgeQuantity,isnull(sum(PledgeNoOfBags),0) PledgeBags,MB.BankName,MBr.BranchName,isnull(TPC.NCMLFinance,0) 'NCMLFinance'"
tmp = tmp & " FROM  Txn_WHR_Detail TWD"
tmp = tmp & " left join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
tmp = tmp & " left Join Txn_PledgeCreation_GSL TPRG on TPRG.Txn_PledgeCreation_ID = TPC.Txn_PledgeCreation_ID And TPRG.SM_Prefix = TPC.SM_Prefix"
tmp = tmp & " left Join Mst_Bank MB on MB.BankID = TPC.BankID"
tmp = tmp & " left Join Mst_Branch MBr on MBr.BranchID = TPC.BranchID"
tmp = tmp & " Where TWD.WHRNo = " & TxtWHRNoNew & " And TPC.Txn_PledgeCreation_ID = " & mPledgeID & ""
tmp = tmp & " Group By TWD.Txn_PledgeCreation_ID,TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity,TPC.Txn_PledgeCreation_ID,TPC.PledgeReqNo,TPC.PledgeDate,MB.BankName,MBr.BranchName,TPC.NCMLFinance"
tmp = tmp & " Having IsNull(Sum(PledgeQuantity), 0) > 0"
Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
   TxtPledgeNo = TmpRs2!Txn_PledgeCreation_ID
   TxtPledgeDate = Format(TmpRs2!PledgeDate, "dd-MMM-yyyy")
   TxtPledgeQty = TmpRs2!PledgeQuantity
   TxtPledgeBags = TmpRs2!PledgeBags
   TxtPledgeBankName = TmpRs2!BankName
   TxtPledgeBranchName = TmpRs2!BranchName
   mNCMLFinance = IIf(IsNull(TmpRs2!NCMLFinance), "0", TmpRs2!NCMLFinance)
Else
   TxtPledgeNo = ""
   TxtPledgeDate = ""
   TxtPledgeQty = ""
   TxtPledgeBags = ""
   TxtPledgeBankName = ""
   TxtPledgeBranchName = ""
   'CmdPledge.Enabled = True
   mNCMLFinance = 0
End If

End Sub

Private Sub SaveCmd_Click()
Dim mIsWHRToBeGenerated As Boolean

If Trim(TxtWHRNo) = "" Then
    MsgBox "WHR No not Define to be Trasnfer."
    TxtWHRNo.SetFocus
    Exit Sub
End If

If Not mWHRCheck Then
    MsgBox "Wrong WHR No or Fully withdrawn WHR not allowed."
    Exit Sub
End If

If Trim(TxtReasonForRevalidation) = "" Then
   MsgBox "Blank Reason for revalidation not allowed!!!"
   TxtReasonForRevalidation.SetFocus
   Exit Sub
End If

If TxtStorageTo.Value >= TxtNewStorageTo.Value Then
   MsgBox "Storage upto Date should be less than New Storage upto Date!!!"
   TxtNewStorageTo.SetFocus
   Exit Sub
End If

'mNewStorageDate = ""
'mSMonth = Month(StorageTo.Value)
'mSYear = Year(StorageTo.Value)
'mSTotalDays = 20
'StorageTo.Value = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
'TxtNewStorageTo.Value = DateAdd("M", 3, StorageTo.Value)
'mNewStorageDate = TxtNewStorageTo.Value
'If TxtAgreementEndDate <> "" Then
'   If CDate(TxtAgreementEndDate) < CDate(TxtNewStorageTo.Value) Then
'      mNewStorageDate = CDate(TxtAgreementEndDate)
'   End If
'End If
'If mNewStorageDate <> "" Then
'   If CDate(mNewStorageDate) < TxtNewStorageTo.Value Then
'      MsgBox "New Storage Date must be greater than " & Format(StorageTo.Value, "DD-MMM-YYYY") & " And Less than " & Format(CDate(mNewStorageDate), "DD-MMM-YYYY") & " "
'   End If
'End If

If TxtAgreementEndDate <> "" Then
   If TxtNewStorageTo.Value > CDate(TxtAgreementEndDate) Then
      MsgBox "'New Storage To Date' cannot be greater than 'Agreement End Date' !!!"
      TxtNewStorageTo.SetFocus
      Exit Sub
   End If
End If


Dim ans As Variant
If mNCMLFinance = "1" Then
   ans = MsgBox("This WHR is pledged marked by NFIN , please check if given WHR is funded, still want to process ?", vbYesNo)
   If ans = vbNo Then
      Exit Sub
   End If
End If

SaveCmd.Enabled = False

tmp = "Select * from Log_WHR_RevalidationDetail Where WHRNO is null "
Call Tmp4Table

TmpRs4.AddNew
TmpRs4!WHRNo = TxtWHRNoNew
TmpRs4!SM_Prefix = mSM_Prefix
TmpRs4!WHRDate = CDate(TxtWHRDate)
TmpRs4!OldStorageTo = CDate(TxtStorageTo.Value)
TmpRs4!OldGradeValidTill = CDate(TxtGradeValidUpto.Value)
TmpRs4!NewStorageTo = TxtNewStorageTo.Value
TmpRs4!NewGradeValidTill = TxtNewStorageTo.Value
TmpRs4!ReasonForRevalidation = Trim(TxtReasonForRevalidation)
TmpRs4!CreatedBy = Gen_UserLoginID
TmpRs4!CreatedDate = Now()
TmpRs4!G_UID = GetGUID
TmpRs4.Update
TmpRs4.Close

tmp = "Update Txn_WHR_Detail Set StorageTo = '" & TxtNewStorageTo.Value & "', GradeValidTill = '" & TxtNewStorageTo.Value & "' Where CMPID = " & mCMPID & " And WHRNO = " & Val(TxtWHRNoNew) & " "
Call Tmp4Table

Call Gen_Email


Call ClearForm
TxtWHRNo = ""
TxtWHRNoNew = ""

End Sub
Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg, mGridMsg As Variant
Dim strSplitString() As String
Dim EmpNo_, EmpName_ As String

mMsg = ""
TOADD = ""
TOCC = ""

mMsg = "<tr><td colspan=2>WHR Details (From Store-Man)</td></tr>"
mMsg = mMsg & "<tr><td>WHR No</td><td>" & TxtWHRNo.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>WHR Date</td><td>" & Format(TxtWHRDate.Value, "dd-MMM-yyyy") & "</td></tr>"
mMsg = mMsg & "<tr><td>Client Name</td><td>" & TxtClient & "</td> </tr>"
mMsg = mMsg & "<tr><td>Commodity</td><td>" & TxtCommodity & "</td></tr>"
mMsg = mMsg & "<tr><td>Godown No</td><td>" & TxtCMPName & "</td></tr>"
mMsg = mMsg & "<tr><td>Godown No</td><td>" & TxtGodownName & "</td></tr>"
mMsg = mMsg & "<tr><td>Godown Address</td><td>" & TxtGodownAddress & "</td></tr>"
mMsg = mMsg & "<tr><td>Old Storage Upto</td><td>" & Format(TxtStorageTo.Value, "DD-MMM-YYYY") & "</td></tr>"
mMsg = mMsg & "<tr><td>New Storage Upto</td><td>" & Format(TxtNewStorageTo.Value, "DD-MMM-YYYY") & "</td></tr>"
mMsg = mMsg & "<tr><td>Reason for Revalidation</td><td>" & Trim(TxtReasonForRevalidation) & "</td></tr>"


If mNCMLFinance = 0 Then
   tmp = " Select EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,EMP.EmployeeEmailID,EMP1.EmployeeID as 'AMEmpID',EMP1.EmployeeNo as 'AMEmpNo',EMP1.EmployeeName as 'AMName',EMP1.EmployeeEmailID as 'AMEmailID'" & _
         " From Mst_EmpCMPMapping EGM" & _
         " Inner Join Mst_Employee EMP on EGM.EmployeeID = EMP.EmployeeID" & _
         " Inner Join Mst_Employee EMP1 on EMP.DirectReportingTo = EMP1.EmployeeID" & _
         " Where EMP.Flag = 'Y' And EGM.CMPID = " & mCMPID
   '      " Group By EMP.EmployeeID,EMP.EmployeeNo,EMP.EmployeeName,EMP.EmployeeEmailID"

   Call Tmp3Table
   If TmpRs3.RecordCount > 0 Then
      TOADD = TmpRs3!EmployeeEmailID
      TOCC = IIf(IsNull(TmpRs3!AMEmailID), "", TmpRs3!AMEmailID) '""
   End If

   If TOADD = "" Then
      TOADD = "manu.ag@ncml.com"
   End If

    '"cmg@ncml.com" & "," & Gen_UserEmailID  'Gen_UserEmailID
   If TOCC <> "" Then
      TOCC = TOCC & "," & Gen_UserEmailID
   Else
      TOCC = Gen_UserEmailID
   End If
Else
   TOADD = Gen_ReadNFinMailIDINI("NFinTo")
   TOCC = Gen_ReadNFinMailIDINI("NFinCC")
   If TOCC <> "" Then
      TOCC = TOCC & "," & Gen_UserEmailID
   Else
      TOCC = Gen_UserEmailID
   End If
End If


'TOADD = "parag.r@ncml.com"
mMsgSubject = "Revalidation of WHR - " & TxtWHRNo & " of Location - " & TxtLocation & " "

Call SendMail(TOADD, TOCC, mMsgSubject, mMsg)


End Sub

Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face= Candara color = Blue><table border = 1>"
htmlStringtail = "</table></Font></body></html>"
   
'With oSMTP
'  'connection
'  .Server = servertxt
'  'message
'  .MailFrom = "info@ncml.com" 'Gen_UserEmailID 'TOCC_ '"mayur.d@ncml.com"
'  .SendTo = ToADD_
''  TOCC_ = TOCC_ '& "," & mVendorEmailID
'  If ToADD_ <> TOCC_ Then
'     .CC = TOCC_
'  End If
''  .BCC = "mayur.d@ncml.com"
'  .Attachments.Add (mPDFFileName)
'  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
'  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
'  .TimeStamp = Now()
'  If LCase(Gen_EmailConfigFlag) = "y" Then
'     .SendEmail
'  End If
'End With


Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
If mNCMLFinance = 1 Then
   objSMTP.ServerName = "webmail.ncml.com" ' = "mail.ncml.com" '"mail.ncml.com"
Else
   objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
End If
objSMTP.FromAddr = "info@ncml.com" 'Gen_UserEmailID
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If
If mPDFFileName <> "" Then
   objSMTP.AddAttachment mPDFFileName
End If
objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If

End Function


Public Sub Grid_Head1()
With MSHFlexGrid1
    .Clear
    .Cols = 14
    .Rows = 3
    .FixedRows = 2
    .Font.Size = 9
    .Cols = 14
    .TextMatrix(1, 0) = "Select":                       .ColWidth(0) = 0
    .TextMatrix(1, 1) = "Stack No ":                    .ColWidth(1) = 1000
    .TextMatrix(1, 2) = "Lot No":                       .ColWidth(2) = 1000
    .TextMatrix(1, 3) = "Available No Of Bags":         .ColWidth(3) = 1000
    .TextMatrix(1, 4) = "Available Weight":             .ColWidth(4) = 1500
    .TextMatrix(1, 5) = "WHR No Of Bags":               .ColWidth(5) = 1200
    .TextMatrix(1, 6) = "WHR Quantity":                 .ColWidth(6) = 1200
    .TextMatrix(1, 7) = "Delivered Qty":                .ColWidth(7) = 1200
    .TextMatrix(1, 8) = "GSL ID":                       .ColWidth(8) = 1000
    .TextMatrix(1, 9) = "Prefix":                       .ColWidth(9) = 1000
    .TextMatrix(1, 10) = "New GSL ID":                  .ColWidth(10) = 1000
    .TextMatrix(1, 11) = "New Stack No":                .ColWidth(11) = 1000
    .TextMatrix(1, 12) = "New Lot No":                  .ColWidth(12) = 1000
    .TextMatrix(1, 13) = "WHRTransferDetailID":         .ColWidth(13) = 0
    .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head2()
With MSHFlexGrid2
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 7
    .Font.Size = 9
    
    .TextMatrix(1, 0) = "Type Of Transaction":         .ColWidth(0) = 2400
    .TextMatrix(1, 1) = "Location":                    .ColWidth(1) = 2400
    .TextMatrix(1, 2) = "CITF No":                     .ColWidth(2) = 2000
    .TextMatrix(1, 3) = "CITF Date":                   .ColWidth(3) = 1500
    .TextMatrix(1, 4) = "CITF Qty":                    .ColWidth(4) = 2000
    .TextMatrix(1, 5) = "CITF No of Bags":             .ColWidth(5) = 2000
    .TextMatrix(1, 6) = "Created Date":                .ColWidth(6) = 2000
End With
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   SaveCmd.Enabled = False
   MsgBox "Access denied!!!"
   Exit Sub
End If
TxtWHRNo = ""
TxtWHRNoNew = ""
Call ClearForm

End Sub
