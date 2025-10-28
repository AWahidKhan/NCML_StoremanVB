VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_WHRDetail 
   Caption         =   "WHR Detail"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   15075
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   123
      Top             =   10320
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9015
      Left            =   15240
      TabIndex        =   122
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9855
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   15015
      Begin VB.Frame Frame0 
         Height          =   2535
         Left            =   240
         TabIndex        =   35
         Top             =   6240
         Visible         =   0   'False
         Width           =   14790
         Begin VB.ComboBox CmbSSCMFlag 
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
            ItemData        =   "Frm_WHRDetail.frx":0000
            Left            =   8880
            List            =   "Frm_WHRDetail.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   143
            Top             =   1680
            Width           =   1125
         End
         Begin VB.CheckBox ChkSWHRPledgePrint 
            Appearance      =   0  'Flat
            Caption         =   "WHR Pledged but not Printed"
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
            Left            =   4395
            TabIndex        =   142
            Top             =   1725
            Width           =   3165
         End
         Begin VB.CheckBox ChkSCARGenerated 
            Appearance      =   0  'Flat
            Caption         =   "WHR Pledged but CAR Not generated"
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
            Left            =   150
            TabIndex        =   141
            Top             =   1725
            Width           =   4005
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   12210
            TabIndex        =   44
            Top             =   1290
            Width           =   1245
         End
         Begin VB.ComboBox CmbSearchStatus 
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
            ItemData        =   "Frm_WHRDetail.frx":0017
            Left            =   8880
            List            =   "Frm_WHRDetail.frx":0021
            Sorted          =   -1  'True
            TabIndex        =   43
            Top             =   1297
            Width           =   1890
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   350
            Left            =   10950
            TabIndex        =   42
            Top             =   1290
            Width           =   1245
         End
         Begin VB.ComboBox CmbSClientName 
            Height          =   315
            Left            =   3345
            Sorted          =   -1  'True
            TabIndex        =   40
            Top             =   1320
            Width           =   3705
         End
         Begin VB.ComboBox CmbPreFix 
            Height          =   315
            Left            =   4230
            Sorted          =   -1  'True
            TabIndex        =   39
            Top             =   600
            Width           =   2565
         End
         Begin VB.ComboBox CmbSCmp 
            Height          =   315
            Left            =   6900
            Sorted          =   -1  'True
            TabIndex        =   38
            Top             =   600
            Width           =   3840
         End
         Begin VB.ComboBox CmbSCommodity 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   1320
            Width           =   3150
         End
         Begin VB.TextBox TxtSWHRNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7125
            MaxLength       =   50
            TabIndex        =   36
            Top             =   1320
            Width           =   1680
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1335
            Left            =   135
            TabIndex        =   41
            Top             =   2190
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   2355
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker DTPWHRFromDate 
            Height          =   345
            Left            =   150
            TabIndex        =   45
            Top             =   585
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            _Version        =   393216
            Format          =   109379585
            CurrentDate     =   40012
         End
         Begin MSComCtl2.DTPicker DTPWHRToDate 
            Height          =   345
            Left            =   2190
            TabIndex        =   46
            Top             =   585
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   609
            _Version        =   393216
            Format          =   109379585
            CurrentDate     =   40012
         End
         Begin VB.Label Label54 
            Caption         =   "SCM WHR"
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
            Left            =   7680
            TabIndex        =   144
            Top             =   1763
            Width           =   975
         End
         Begin VB.Label LblRecordCount 
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
            Left            =   13455
            TabIndex        =   57
            Top             =   645
            Width           =   90
         End
         Begin VB.Label Label25 
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
            Left            =   14160
            TabIndex        =   56
            Top             =   240
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
            Left            =   10965
            TabIndex        =   55
            Top             =   630
            Width           =   2415
         End
         Begin VB.Label Label24 
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
            Height          =   375
            Left            =   9495
            TabIndex        =   54
            Top             =   953
            Width           =   735
         End
         Begin VB.Label lblWHRDateFrom 
            Caption         =   "From WHR Date"
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
            Left            =   285
            TabIndex        =   53
            Top             =   225
            Width           =   1665
         End
         Begin VB.Label LblWHRToDate 
            Caption         =   "To WHR Date"
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
            Left            =   2415
            TabIndex        =   52
            Top             =   225
            Width           =   1425
         End
         Begin VB.Label Label28 
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
            Height          =   300
            Left            =   4627
            TabIndex        =   51
            Top             =   990
            Width           =   1140
         End
         Begin VB.Label Label36 
            Caption         =   "Location Name"
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
            Left            =   4800
            TabIndex        =   50
            Top             =   225
            Width           =   1680
         End
         Begin VB.Label Label29 
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
            Height          =   300
            Left            =   8700
            TabIndex        =   49
            Top             =   225
            Width           =   480
         End
         Begin VB.Label Label30 
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
            Height          =   300
            Left            =   1140
            TabIndex        =   48
            Top             =   990
            Width           =   1110
         End
         Begin VB.Label Label31 
            Caption         =   "WHR No"
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
            Left            =   7440
            TabIndex        =   47
            Top             =   990
            Width           =   960
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H008080FF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   6840
         TabIndex        =   21
         Top             =   4560
         Visible         =   0   'False
         Width           =   3375
         Begin VB.OptionButton OptLetter 
            BackColor       =   &H008080FF&
            Caption         =   "Letter Head"
            Height          =   255
            Left            =   1560
            TabIndex        =   25
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptReg 
            BackColor       =   &H008080FF&
            Caption         =   "Regular"
            Height          =   255
            Left            =   480
            TabIndex        =   24
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.CommandButton CmdPCancel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cancel"
            Height          =   400
            Left            =   1680
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton CmdPOk 
            BackColor       =   &H00E0E0E0&
            Caption         =   "OK"
            Height          =   400
            Left            =   240
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   840
            Width           =   1440
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H008080FF&
         Caption         =   "Hologram No "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   5880
         TabIndex        =   26
         Top             =   3120
         Visible         =   0   'False
         Width           =   4935
         Begin VB.TextBox TxtHologramNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1440
            MaxLength       =   25
            TabIndex        =   32
            Top             =   360
            Width           =   2880
         End
         Begin VB.CommandButton CmdSaveHologram 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Save"
            Height          =   400
            Left            =   1440
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   840
            Width           =   1440
         End
         Begin VB.CommandButton CmdCancel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cancel"
            Height          =   400
            Left            =   2880
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   840
            Width           =   1455
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H008080FF&
            Caption         =   "Generate PDF"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1440
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H008080FF&
            Caption         =   "Regular"
            Height          =   255
            Left            =   2160
            TabIndex        =   28
            Top             =   1440
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H008080FF&
            Caption         =   "Letter Head"
            Height          =   255
            Left            =   3240
            TabIndex        =   27
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label33 
            BackColor       =   &H008080FF&
            Caption         =   "Hologram No"
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
            TabIndex        =   34
            Top             =   390
            Width           =   1260
         End
         Begin VB.Label Label48 
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            Caption         =   $"Frm_WHRDetail.frx":003D
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1260
            Left            =   120
            TabIndex        =   33
            Top             =   1800
            Width           =   4380
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8475
         Left            =   120
         TabIndex        =   68
         Top             =   120
         Width           =   14805
         Begin VB.TextBox TxtAgreementEndDate 
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
            Left            =   11640
            Locked          =   -1  'True
            TabIndex        =   171
            TabStop         =   0   'False
            Top             =   4485
            Width           =   2970
         End
         Begin VB.ComboBox CmbGodownID 
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
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   17
            Top             =   4440
            Width           =   6285
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Converted Client"
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
            Height          =   495
            Left            =   240
            TabIndex        =   7
            Top             =   1680
            Width           =   1650
         End
         Begin VB.TextBox TxtRisks 
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
            Left            =   3600
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   3270
            Width           =   10815
         End
         Begin VB.TextBox TxtWHRNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox CmbWHRCommodity 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   2100
            Width           =   10815
         End
         Begin VB.TextBox TxtTotalWHRQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   11640
            Locked          =   -1  'True
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   4080
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.ComboBox CmbClientId 
            Height          =   315
            Left            =   3615
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   1680
            Width           =   10815
         End
         Begin VB.ComboBox CmbCommodityCondition 
            Height          =   315
            ItemData        =   "Frm_WHRDetail.frx":00CE
            Left            =   7800
            List            =   "Frm_WHRDetail.frx":00DB
            TabIndex        =   5
            Top             =   1200
            Width           =   3375
         End
         Begin VB.TextBox TxtGrade 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   11280
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1200
            Width           =   3120
         End
         Begin VB.TextBox TxtStorageCharges 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   14
            Top             =   3660
            Width           =   1800
         End
         Begin VB.TextBox TxtStorageAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   8160
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   3645
            Width           =   1680
         End
         Begin VB.TextBox TxtValueofCommodity 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Height          =   360
            Left            =   8160
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   4035
            Width           =   1680
         End
         Begin VB.TextBox TxtAgreeDate 
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
            Left            =   13680
            Locked          =   -1  'True
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   4080
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox TxtWHRNoNew 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   600
            Width           =   1575
         End
         Begin VB.ComboBox CmbInsuranceCo 
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
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   2460
            Width           =   5205
         End
         Begin VB.CommandButton CmdInsuranceCo 
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
            Left            =   8805
            TabIndex        =   75
            Top             =   2460
            Width           =   375
         End
         Begin VB.ComboBox CmbPolicyNo 
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
            Left            =   9240
            Sorted          =   -1  'True
            TabIndex        =   11
            Top             =   2460
            Width           =   4965
         End
         Begin VB.CommandButton CmdInsurance 
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
            Left            =   14220
            TabIndex        =   74
            Top             =   2460
            Width           =   360
         End
         Begin VB.TextBox TxtAmount 
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
            Height          =   360
            Left            =   8160
            Locked          =   -1  'True
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1680
         End
         Begin VB.CheckBox ChkSaidToContain 
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   11640
            TabIndex        =   12
            Top             =   2925
            Width           =   225
         End
         Begin VB.CheckBox ChkEstimatedWgt 
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   14160
            TabIndex        =   13
            Top             =   2925
            Width           =   210
         End
         Begin VB.ComboBox CmbCMPID 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            TabIndex        =   1
            Top             =   360
            Width           =   10815
         End
         Begin VB.TextBox TxtBODDate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   375
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1200
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox TxtRate 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   16
            Top             =   4065
            Width           =   1800
         End
         Begin VB.OptionButton OptSRPerMT 
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3000
            TabIndex        =   70
            Top             =   3885
            Width           =   195
         End
         Begin VB.OptionButton OptSRPerBag 
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3000
            TabIndex        =   69
            Top             =   3682
            Value           =   -1  'True
            Width           =   195
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   3495
            Left            =   120
            TabIndex        =   76
            Top             =   4920
            Width           =   14565
            _ExtentX        =   25691
            _ExtentY        =   6165
            _Version        =   393216
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "GSL Detail"
            TabPicture(0)   =   "Frm_WHRDetail.frx":00F4
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "ChkSelectAll"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "MSHFlexGrid4"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Frame2"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "MSHFlexGrid2"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "Pledge / Other Details"
            TabPicture(1)   =   "Frm_WHRDetail.frx":0110
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame6"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "CITF Details"
            TabPicture(2)   =   "Frm_WHRDetail.frx":012C
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "MSHFlexGrid3"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Log Details"
            TabPicture(3)   =   "Frm_WHRDetail.frx":0148
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "TxtCreated"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).Control(1)=   "TxtUpdated"
            Tab(3).Control(1).Enabled=   0   'False
            Tab(3).Control(2)=   "Label46"
            Tab(3).Control(3)=   "Label47"
            Tab(3).ControlCount=   4
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
               Left            =   -72840
               Locked          =   -1  'True
               TabIndex        =   146
               TabStop         =   0   'False
               Top             =   840
               Width           =   5160
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
               Left            =   -65520
               Locked          =   -1  'True
               TabIndex        =   145
               TabStop         =   0   'False
               Top             =   840
               Width           =   4440
            End
            Begin VB.Frame Frame6 
               Height          =   2985
               Left            =   -74880
               TabIndex        =   77
               Top             =   360
               Width           =   14175
               Begin VB.TextBox TxtCARDate 
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
                  Left            =   12360
                  Locked          =   -1  'True
                  TabIndex        =   139
                  TabStop         =   0   'False
                  Top             =   2520
                  Width           =   1680
               End
               Begin VB.TextBox TxtCARNo 
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
                  Left            =   9600
                  Locked          =   -1  'True
                  TabIndex        =   137
                  TabStop         =   0   'False
                  Top             =   2520
                  Width           =   1680
               End
               Begin VB.TextBox TxtPledgeBranchName 
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
                  Left            =   9600
                  Locked          =   -1  'True
                  TabIndex        =   134
                  TabStop         =   0   'False
                  Top             =   1635
                  Width           =   4440
               End
               Begin VB.TextBox TxtPledgeBankName 
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
                  Left            =   2295
                  Locked          =   -1  'True
                  TabIndex        =   132
                  TabStop         =   0   'False
                  Top             =   1635
                  Width           =   5160
               End
               Begin VB.TextBox TxtPledgeBags 
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
                  Left            =   5640
                  Locked          =   -1  'True
                  TabIndex        =   130
                  TabStop         =   0   'False
                  Top             =   2535
                  Width           =   1800
               End
               Begin VB.TextBox TxtPledgeQty 
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
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   128
                  TabStop         =   0   'False
                  Top             =   2535
                  Width           =   1920
               End
               Begin VB.TextBox TxtPledgeDate 
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
                  Left            =   9600
                  Locked          =   -1  'True
                  TabIndex        =   126
                  TabStop         =   0   'False
                  Top             =   2085
                  Width           =   1680
               End
               Begin VB.TextBox TxtPledgeNo 
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
                  Left            =   2280
                  Locked          =   -1  'True
                  TabIndex        =   124
                  TabStop         =   0   'False
                  Top             =   2085
                  Width           =   5160
               End
               Begin VB.TextBox TxtAssayerCerti 
                  Appearance      =   0  'Flat
                  Height          =   360
                  Left            =   3600
                  MaxLength       =   50
                  TabIndex        =   80
                  Top             =   1185
                  Width           =   10440
               End
               Begin VB.TextBox TxtAssayer 
                  Appearance      =   0  'Flat
                  Height          =   360
                  Left            =   3600
                  MaxLength       =   50
                  TabIndex        =   79
                  Top             =   720
                  Width           =   10440
               End
               Begin VB.TextBox TxtWeigher 
                  Appearance      =   0  'Flat
                  Height          =   360
                  Left            =   3600
                  MaxLength       =   50
                  TabIndex        =   78
                  Top             =   240
                  Width           =   10440
               End
               Begin VB.Label Label49 
                  Caption         =   "CAR Date"
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
                  Left            =   11400
                  TabIndex        =   140
                  Top             =   2580
                  Width           =   1215
               End
               Begin VB.Label Label45 
                  Caption         =   "CAR No"
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
                  Left            =   7560
                  TabIndex        =   138
                  Top             =   2580
                  Width           =   1215
               End
               Begin VB.Label Label44 
                  Caption         =   "Pledge Branch Name"
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
                  Left            =   7560
                  TabIndex        =   135
                  Top             =   1687
                  Width           =   2055
               End
               Begin VB.Label Label39 
                  Caption         =   "Pledge Bank Name"
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
                  Left            =   345
                  TabIndex        =   133
                  Top             =   1702
                  Width           =   2415
               End
               Begin VB.Label Label37 
                  Caption         =   "Pledge Bags"
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
                  Left            =   4320
                  TabIndex        =   131
                  Top             =   2595
                  Width           =   1215
               End
               Begin VB.Label Label35 
                  Caption         =   "Pledge Quantity"
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
                  Left            =   360
                  TabIndex        =   129
                  Top             =   2602
                  Width           =   1455
               End
               Begin VB.Label Label34 
                  Caption         =   "Pledge Date"
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
                  Left            =   7560
                  TabIndex        =   127
                  Top             =   2152
                  Width           =   1215
               End
               Begin VB.Label Label32 
                  Caption         =   "Pledge ID"
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
                  Left            =   360
                  TabIndex        =   125
                  Top             =   2152
                  Width           =   1695
               End
               Begin VB.Label Label5 
                  Caption         =   "Grade as per Assayer’s Certificate"
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
                  Left            =   360
                  TabIndex        =   83
                  Top             =   1215
                  Width           =   3060
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Name and License No. of Weigher"
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
                  TabIndex        =   82
                  Top             =   270
                  Width           =   3300
               End
               Begin VB.Label Label4 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Name and License No. of Assayer"
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
                  TabIndex        =   81
                  Top             =   750
                  Width           =   3180
               End
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
               Height          =   2025
               Left            =   120
               TabIndex        =   84
               Top             =   1320
               Width           =   14340
               _ExtentX        =   25294
               _ExtentY        =   3572
               _Version        =   393216
               WordWrap        =   -1  'True
               AllowUserResizing=   3
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
               Height          =   2880
               Left            =   -74880
               TabIndex        =   85
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
            Begin VB.Frame Frame2 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   870
               Left            =   120
               TabIndex        =   149
               Top             =   360
               Visible         =   0   'False
               Width           =   14340
               Begin VB.TextBox TxtGodownNo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  Height          =   360
                  Left            =   5220
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   158
                  TabStop         =   0   'False
                  Top             =   405
                  Width           =   3600
               End
               Begin VB.TextBox TxtWHRNoOfBags 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   360
                  Left            =   8880
                  MaxLength       =   50
                  TabIndex        =   157
                  Top             =   405
                  Width           =   1710
               End
               Begin VB.CommandButton CmdUpdate 
                  Caption         =   "Update List"
                  Height          =   360
                  Left            =   12435
                  TabIndex        =   156
                  Top             =   405
                  Width           =   1680
               End
               Begin VB.TextBox TxtWHRQty 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  Height          =   360
                  Left            =   10605
                  MaxLength       =   50
                  TabIndex        =   155
                  Top             =   405
                  Width           =   1740
               End
               Begin VB.TextBox TxtEleWeight 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  Height          =   360
                  Left            =   3825
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   154
                  TabStop         =   0   'False
                  Top             =   885
                  Visible         =   0   'False
                  Width           =   1680
               End
               Begin VB.TextBox TxtLicenseDescription 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  Height          =   360
                  Left            =   7365
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   153
                  TabStop         =   0   'False
                  Top             =   885
                  Visible         =   0   'False
                  Width           =   3630
               End
               Begin VB.TextBox TxtLicenseNo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  Height          =   360
                  Left            =   300
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   152
                  TabStop         =   0   'False
                  Top             =   885
                  Visible         =   0   'False
                  Width           =   1575
               End
               Begin VB.TextBox TxtNoOfBags 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  Height          =   360
                  Left            =   1905
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   151
                  TabStop         =   0   'False
                  Top             =   885
                  Visible         =   0   'False
                  Width           =   1905
               End
               Begin VB.TextBox TxtCMPID 
                  Appearance      =   0  'Flat
                  BackColor       =   &H8000000A&
                  Height          =   360
                  Left            =   60
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   150
                  TabStop         =   0   'False
                  Top             =   405
                  Width           =   5130
               End
               Begin VB.Label Label18 
                  Caption         =   "Godown No"
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
                  Left            =   6450
                  TabIndex        =   166
                  Top             =   120
                  Width           =   1140
               End
               Begin VB.Label Label16 
                  Caption         =   "WHR No Of Bags"
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
                  Left            =   9000
                  TabIndex        =   165
                  Top             =   120
                  Width           =   1620
               End
               Begin VB.Label Label14 
                  Caption         =   "WHR Qty"
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
                  Left            =   11055
                  TabIndex        =   164
                  Top             =   120
                  Width           =   900
               End
               Begin VB.Label Label13 
                  Caption         =   "Electronic Weight"
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
                  Left            =   3915
                  TabIndex        =   163
                  Top             =   465
                  Visible         =   0   'False
                  Width           =   1620
               End
               Begin VB.Label Label12 
                  Caption         =   "License Description"
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
                  Left            =   8235
                  TabIndex        =   162
                  Top             =   495
                  Visible         =   0   'False
                  Width           =   1860
               End
               Begin VB.Label Label11 
                  Caption         =   "License No"
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
                  Left            =   540
                  TabIndex        =   161
                  Top             =   510
                  Visible         =   0   'False
                  Width           =   1020
               End
               Begin VB.Label Label15 
                  Caption         =   "Available No Of Bags"
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
                  Left            =   1860
                  TabIndex        =   160
                  Top             =   465
                  Visible         =   0   'False
                  Width           =   1965
               End
               Begin VB.Label Label17 
                  Caption         =   "CMP Name"
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
                  Left            =   2055
                  TabIndex        =   159
                  Top             =   150
                  Width           =   1140
               End
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
               Height          =   2745
               Left            =   120
               TabIndex        =   168
               Top             =   600
               Visible         =   0   'False
               Width           =   14340
               _ExtentX        =   25294
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
            Begin VB.CheckBox ChkSelectAll 
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
               Height          =   240
               Left            =   120
               TabIndex        =   170
               Top             =   360
               Width           =   2370
            End
            Begin VB.Label Label46 
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
               Left            =   -74760
               TabIndex        =   148
               Top             =   900
               Width           =   1695
            End
            Begin VB.Label Label47 
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
               Left            =   -67560
               TabIndex        =   147
               Top             =   915
               Width           =   975
            End
         End
         Begin MSComCtl2.DTPicker WHRDate 
            Height          =   330
            Left            =   1920
            TabIndex        =   2
            Top             =   1200
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   582
            _Version        =   393216
            Format          =   109379585
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker StorageFrom 
            Height          =   315
            Left            =   3600
            TabIndex        =   3
            Top             =   1200
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Format          =   109379585
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker StorageTo 
            Height          =   315
            Left            =   5760
            TabIndex        =   4
            Top             =   1200
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Format          =   109379585
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker GradeValidUpto 
            Height          =   375
            Left            =   11640
            TabIndex        =   15
            Top             =   3630
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   109379585
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker ValidFrom 
            Height          =   375
            Left            =   3600
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   109379585
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker ValidTo 
            Height          =   375
            Left            =   6120
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   109379585
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtCMPName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   330
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   98
            Top             =   360
            Width           =   10800
         End
         Begin VB.TextBox TxtInsuranceCo 
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
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   97
            Top             =   2460
            Width           =   5175
         End
         Begin VB.TextBox TxtPolicyNo 
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
            Left            =   9240
            Locked          =   -1  'True
            TabIndex        =   96
            Top             =   2460
            Width           =   4815
         End
         Begin VB.TextBox TxtClientID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   92
            Top             =   1680
            Width           =   10815
         End
         Begin VB.TextBox TxtWHRCommodityId 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   315
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   93
            Top             =   2100
            Width           =   10815
         End
         Begin VB.TextBox TxtGodownID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Height          =   330
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   169
            Top             =   4440
            Width           =   6285
         End
         Begin VB.Label Label50 
            Caption         =   "Agreement End Date"
            Height          =   225
            Left            =   10080
            TabIndex        =   172
            Top             =   4553
            Width           =   1695
         End
         Begin VB.Label LblGodownID 
            Caption         =   "Godown No"
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
            Left            =   720
            TabIndex        =   167
            Top             =   4500
            Width           =   1800
         End
         Begin VB.Label LblWHRNo 
            Caption         =   "WHR No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   600
            TabIndex        =   120
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "CMP Name"
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
            TabIndex        =   121
            Top             =   360
            Width           =   3300
         End
         Begin VB.Label LblTotalWHRQty 
            Caption         =   "Total WHR Qty"
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
            Left            =   10110
            TabIndex        =   119
            Top             =   4080
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label LblWHRDate 
            Caption         =   "WHR Date"
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
            Left            =   2160
            TabIndex        =   118
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label Label1 
            Caption         =   "Condition of Commodities"
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
            Left            =   7830
            TabIndex        =   117
            Top             =   855
            Width           =   3300
         End
         Begin VB.Label Label3 
            Caption         =   "Private Marks on Packaging"
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
            Left            =   11280
            TabIndex        =   116
            Top             =   840
            Width           =   3180
         End
         Begin VB.Label Label8 
            Caption         =   "Storage Charge Amount"
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
            Left            =   5760
            TabIndex        =   115
            Top             =   3660
            Width           =   2160
         End
         Begin VB.Label Label10 
            Caption         =   "Storage upto Date"
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
            Left            =   5775
            TabIndex        =   114
            Top             =   840
            Width           =   2040
         End
         Begin VB.Label Label9 
            Caption         =   "Date of Deposit "
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
            Left            =   3720
            TabIndex        =   113
            Top             =   840
            Width           =   1680
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Storage Rate Per Month Per Bag"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   112
            Top             =   3660
            Width           =   2820
         End
         Begin VB.Label LblClientName 
            Alignment       =   1  'Right Justify
            Caption         =   "Client Name"
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
            Left            =   1800
            TabIndex        =   111
            Top             =   1687
            Width           =   1500
         End
         Begin VB.Label LblWHRCommodity 
            Alignment       =   1  'Right Justify
            Caption         =   "Commodity"
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
            Left            =   1920
            TabIndex        =   110
            Top             =   2107
            Width           =   1500
         End
         Begin VB.Label Label20 
            Caption         =   "Value of Commodity"
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
            Left            =   5760
            TabIndex        =   109
            Top             =   4050
            Width           =   2160
         End
         Begin VB.Label Label6 
            Caption         =   "Grade Valid Till"
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
            Left            =   9960
            TabIndex        =   108
            Top             =   3667
            Width           =   1470
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            Caption         =   "Insurance Co. And Policy No."
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
            TabIndex        =   107
            Top             =   2513
            Width           =   3255
         End
         Begin VB.Label Label40 
            Caption         =   "Valid To"
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
            Left            =   5160
            TabIndex        =   106
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            Caption         =   "Valid From"
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
            TabIndex        =   105
            Top             =   2880
            Width           =   3375
         End
         Begin VB.Label Label42 
            Caption         =   "Amount"
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
            Left            =   7440
            TabIndex        =   104
            Top             =   2880
            Width           =   1695
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "Risks"
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
            TabIndex        =   103
            Top             =   3315
            Width           =   3375
         End
         Begin VB.Label Label21 
            Caption         =   "Said To Contain"
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
            Left            =   10080
            TabIndex        =   102
            Top             =   2925
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "Estimated Weight"
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
            Left            =   12360
            TabIndex        =   101
            Top             =   2925
            Width           =   1695
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Commodity Price (Rate Per MT)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   100
            Top             =   4155
            Width           =   3300
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Storage Rate Per Month Per MT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   99
            Top             =   3870
            Width           =   2820
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1200
         Left            =   120
         TabIndex        =   58
         Top             =   8520
         Width           =   14820
         Begin VB.CommandButton CmdPledge 
            Caption         =   "Create Pledge"
            Height          =   810
            Left            =   13425
            TabIndex        =   136
            Top             =   195
            Width           =   1215
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   105
            TabIndex        =   67
            Top             =   195
            Width           =   2220
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   105
            TabIndex        =   66
            Top             =   600
            Width           =   2220
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   2325
            TabIndex        =   18
            Top             =   195
            Width           =   2220
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   2325
            TabIndex        =   65
            Top             =   600
            Width           =   2220
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   4545
            TabIndex        =   64
            Top             =   195
            Width           =   2220
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   4545
            TabIndex        =   63
            Top             =   600
            Width           =   2220
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   6765
            TabIndex        =   62
            Top             =   195
            Width           =   2220
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   6765
            TabIndex        =   61
            Top             =   600
            Width           =   2220
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   8985
            TabIndex        =   20
            Top             =   195
            Width           =   2220
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   8985
            TabIndex        =   19
            Top             =   600
            Width           =   2220
         End
         Begin VB.CommandButton CmdPrint 
            Caption         =   "&Print"
            Height          =   400
            Left            =   11205
            TabIndex        =   60
            Top             =   600
            Width           =   2220
         End
         Begin VB.CommandButton CmdPrintPreview 
            Caption         =   "Print Preview"
            Height          =   400
            Left            =   11205
            TabIndex        =   59
            Top             =   195
            Width           =   2220
         End
      End
   End
End
Attribute VB_Name = "Frm_WHRDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mExchangeTypeID, mCommodityCondition, mWHRNoOfBags, mWHRQty As Double
Dim pq, temp, GSLQty, ChkSelection As Double
Dim mClientIDRow As Double
Dim mCommodityID As Double
Dim mGodownID As Variant
Dim mCMPID As Double
Dim mWHRNo As Double
Dim PolicyNo_, CNo As String
Dim Edit_Flg, mLicenseNo, mInsuranceCoID, mSM_Prefix As Variant
Dim PntPrv  As Variant
Dim mSLocationID, mSPrefix, mSCMPID, mSCommodityID, mSClientID, mLocationID As Variant

'Declaration for PDF report
Dim pgval As Double
Dim objPDF As New mjwPDF
Dim mLeft As Double
Dim mTop As Double
Dim mHeight As Double
Dim mRowHeight As Double
Dim mRow As Double
Dim mGlobalLeft As Double
Dim mTxtTop As Double
Dim mTxtLeft As Double
Dim mFontSize As Integer
Dim mTxtFromTop As Double
Dim mPDFFileName As Variant
Dim mPrintFlag As Variant
Dim mWHRVersion As Variant
Dim mPledgeID As Variant

Dim mLicenseID As Variant
Dim mSCMFlag As Variant
Dim mTypeFlag As Variant
Dim mSCMPledgeID As Variant
Dim mGSLForPledge As Variant
Dim mNCMLFinance As Variant
Dim mLotWiseWHR As Variant
Const strChecked = "þ"
Const strUnChecked = "q"
Dim mGSLIDinGrid As Variant
Dim mLicenseExpiryDate As Variant

Private Sub AddCmd_Click()
TxtAgreeDate = ""
Edit_Flg = "A"
mLotWiseWHR = Null ' "1"
MSHFlexGrid2.Visible = False
Frame2.Visible = False
MSHFlexGrid4.Visible = True
ChkSelectAll.Enabled = True
WHRDate.Enabled = True
CmbCommodityCondition.Locked = False
StorageFrom.Enabled = True
StorageTo.Enabled = True
TxtGrade.Locked = False
TxtInsuranceCo.Visible = False
TxtPolicyNo.Visible = False
CmbInsuranceCo.Visible = True
CmbPolicyNo.Visible = True
CmbInsuranceCo.Text = ""
CmbPolicyNo.Text = ""
CmdInsuranceCo.Visible = True
CmdInsurance.Visible = True
TxtClientID.Visible = False
CmbClientID.Visible = True
CmbClientID.Locked = False
TxtWHRCommodityId.Visible = False
CmbWHRCommodity.Visible = True
CmbWHRCommodity.Locked = False





TxtCMPName.Visible = False
CmbCMPID.Visible = True
CmbCMPID.Text = ""

TxtGodownID.Visible = False
CmbGodownID.Visible = True
CmbGodownID.Text = ""

TxtWeigher.Locked = False
TxtAssayer.Locked = False
TxtAssayerCerti.Locked = False
GradeValidUpto.Enabled = True
TxtStorageCharges.Locked = False
TxtRate.Locked = False
TxtWHRNoOfBags.Locked = True
TxtWHRQty.Locked = True
TxtWHRNo = ""
TxtPreFix = ""
TxtWHRNoNew = ""
mCMPID = 0
mGodownID = 0
CmbCommodityCondition.Text = ""
mCommodityID = 0
mClientIDRow = 0
WHRDate.Value = Date
StorageFrom.Value = Date
StorageTo.Value = Date
TxtGrade = ""
CmbClientID.Text = ""
CmbWHRCommodity.Text = ""
TxtWeigher = ""
TxtAssayer = ""
TxtAssayerCerti = ""
GradeValidUpto = Date
TxtStorageCharges = ""
TxtRate = ""
TxtValueofCommodity = ""
TxtAmount = ""
TxtRisks = ""

TxtCreated = ""
TxtUpdated = ""
TxtCARNo = ""
TxtCARDate = ""


TxtStorageAmount = ""
TxtTotalWHRQty = ""
ChkSaidToContain.Enabled = True
ChkEstimatedWgt.Enabled = True
ChkSaidToContain.Value = 1
ChkEstimatedWgt.Value = 1
ChkSelection = 0
OptSRPerBag.Enabled = True
OptSRPerMT.Enabled = True
OptSRPerBag.Value = True
TxtStorageCharges.Locked = False

TxtPledgeNo = ""
TxtPledgeDate = ""
TxtPledgeQty = ""
TxtPledgeBags = ""
TxtPledgeBankName = ""
TxtPledgeBranchName = ""


Check1.Value = vbUnchecked
If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg" Or LCase(Gen_UserRole) = "admin" Then
   Check1.Enabled = True
Else
   Check1.Enabled = False
End If


Call Grid_HeadNew

Call Grid_Head2

CmdUpdate.Enabled = False
Call GButtonLock(Me, False)
CmdPrint.Enabled = False
CmdPrintPreview.Enabled = False
CmdPledge.Enabled = False
ChkSelectAll.Visible = True
SSTab1.Tab = 0
  
  
End Sub

Private Sub Check1_Click()
CmbClientID.Text = ""
Call Grid_Head
End Sub

Private Sub ChkSelectAll_Click()
If mSCMFlag = 1 Then Exit Sub
If MSHFlexGrid4.Rows = 3 Then
   ChkSelectAll.Value = 0
   Exit Sub
End If

If ChkSelectAll.Value = vbChecked Then
   For I = 2 To MSHFlexGrid4.Rows - 2
       MSHFlexGrid4.row = I
       MSHFlexGrid4.Col = 0
       MSHFlexGrid4.TextMatrix(I, 5) = Val(MSHFlexGrid4.TextMatrix(I, 3))
       MSHFlexGrid4.TextMatrix(I, 6) = Val(MSHFlexGrid4.TextMatrix(I, 4))
       If MSHFlexGrid4.TextMatrix(I, 0) = strUnChecked Then
          MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(I, 5))
          MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(I, 6))
          MSHFlexGrid4.TextMatrix(I, 0) = strChecked
       End If
   Next
Else
   For I = 2 To MSHFlexGrid4.Rows - 2
       MSHFlexGrid4.row = I
       MSHFlexGrid4.TextMatrix(I, 0) = strUnChecked
       MSHFlexGrid4.TextMatrix(I, 5) = 0 'Val(MSHFlexGrid2.TextMatrix(I, 13))
       MSHFlexGrid4.TextMatrix(I, 6) = 0 'Val(MSHFlexGrid2.TextMatrix(I, 14))
       MSHFlexGrid4.TextMatrix(0, 5) = 0 'Val(MSHFlexGrid2.TextMatrix(0, 16)) - Val(MSHFlexGrid2.TextMatrix(I, 16))
       MSHFlexGrid4.TextMatrix(0, 6) = 0 'Val(MSHFlexGrid2.TextMatrix(0, 17)) - Val(MSHFlexGrid2.TextMatrix(I, 17))
   Next
End If

End Sub

Private Sub CmbClientId_GotFocus()
tmp = CmbClientID.Text
mExchangeTypeID = GetExchangeID("NON NCDEX")


'If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "head cmg" Then
'   wc = ""
'ElseIf LCase(Gen_UserRole) = "scm admin" Then
'   wc = " And SCMFlag = 1 "
'Else
'   wc = " And SCMFlag = 0 "
'End If

If Check1.Value = vbUnchecked Then
   Call TableMst_Client(" Where Converted = '0' And ExchangeTypeId = " & mExchangeTypeID & " " & wc & " ")
Else
   Call TableMst_Client(" Where Converted = '1' And ExchangeTypeId = " & mExchangeTypeID & " " & wc & " ")
End If
CmbClientID.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbClientID.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbClientID.Text = tmp
End Sub
Private Sub CmbClientId_LostFocus()
If CmbClientID.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbClientID.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection!!! "
   CmbClientID.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientID
mSCMFlag = IIf(IsNull(RsMst_Client!SCMFlag), 0, RsMst_Client!SCMFlag)
Call GetGridData(mExchangeTypeID, mCommodityID, mClientIDRow)


End Sub

Private Sub CmbCMPID_GotFocus()
tmp = CmbCMPID.Text

tmp1 = ""
If Gen_WcCMP <> "" Then
   tmp1 = "And " & Gen_WcCMP
End If

Call TableMst_CMP(" where CloseDate is null " & tmp1)
CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp

End Sub
'----
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection!!! "
   CmbCMPID.SetFocus
   Exit Sub
End If

If mCMPID <> RsMst_CMP!CMPID Then
   CmbWHRCommodity.Text = ""
End If

mCMPID = RsMst_CMP!CMPID
mLocationID = RsMst_CMP!LocationID
mSM_Prefix = GetPreFixByLocation(RsMst_CMP!LocationID, "P")
TxtBODDate = IIf(IsNull(RsMst_CMP!CurrentDate), RsMst_CMP!StartDate, RsMst_CMP!CurrentDate) + 1
WHRDate.Value = IIf(IsNull(RsMst_CMP!CurrentDate), RsMst_CMP!StartDate, RsMst_CMP!CurrentDate) + 1
StorageFrom.Value = WHRDate.Value
Call GetGridData(mExchangeTypeID, mCommodityID, mClientIDRow)

End Sub

Private Sub CmbCommodityCondition_LostFocus()
mCommodityCondition = CmbCommodityCondition.ListIndex
If CmbCommodityCondition.Text = "" Then Exit Sub
If CmbCommodityCondition.Text = "Average" Or CmbCommodityCondition.Text = "Fair" Or _
   CmbCommodityCondition.Text = "Good" Then
Else
   MsgBox "Invalid selection!!! "
   CmbCommodityCondition.SetFocus
   Exit Sub
End If
     
End Sub



Private Sub CmbFilterBy_GotFocus()
LblSearchTitle.Visible = True
TxtSearchWHR.Visible = True
LblSearchTitle.Caption = ""
TxtSearchWHR = ""

End Sub

Private Sub CmbFilterBy_LostFocus()
If CmbFilterBy.Text = "" Then
   Exit Sub
End If

If CmbFilterBy.Text <> "All" And CmbFilterBy.Text <> "WHR Date" And CmbFilterBy.Text <> "WHR No" And CmbFilterBy.Text <> "Location Name" And CmbFilterBy.Text <> "CMP Name" And CmbFilterBy.Text <> "Client Name" And CmbFilterBy.Text <> "Commodity Name" Then
   MsgBox "Invalid selection!!! "
   CmbFilterBy.SetFocus
   Exit Sub
End If

CmbSearchStatus.Text = ""
CmdGo.Visible = True

If CmbFilterBy.Text = "WHR Date" Then
    DTPWHRFromDate.Enabled = True
    DTPWHRToDate.Enabled = True
    LblSearchTitle.Visible = False
    TxtSearchWHR.Visible = False
Else
   DTPWHRFromDate.Enabled = False
   DTPWHRToDate.Enabled = False
End If
LblSearchTitle.Caption = CmbFilterBy.Text
End Sub

Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text
CmbSCmp.Clear
If Gen_WcLocation <> "" Then
   Qry = " Where  " & Gen_WcLocation & " "
Else
   Qry = ""
End If
   Call TableMst_Location(Qry)


CmbPreFix.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbPreFix.AddItem RsMst_Location!Locationname
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

If tmp <> CmbPreFix.Text Then
   CmbSCmp.Text = ""
End If

mSLocationID = RsMst_Location!LocationID
mSPrefix = RsMst_Location!SM_Prefix
End Sub




Private Sub CmbSClientName_GotFocus()
tmp = CmbSClientName.Text

Call TableMst_Client("Where ExchangeTypeId = 1 ") '" & mExchangeTypeID)
CmbSClientName.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbSClientName.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbSClientName.Text = tmp

End Sub

Private Sub CmbSClientName_LostFocus()
If CmbSClientName.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbSClientName.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbSClientName.SetFocus
   Exit Sub
End If
mSClientID = RsMst_Client!ClientID

End Sub

Private Sub CmbSCmp_GotFocus()
If CmbPreFix.Text = "" Then
   MsgBox "Select Location !!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
tmp = CmbSCmp.Text
Call TableMst_CMP(" Where LocationID = " & mSLocationID & "")
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbSCmp.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSCmp.Text = tmp
End Sub

Private Sub CmbSCmp_LostFocus()
If CmbSCmp.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCmp.SetFocus
   Exit Sub
End If
mSCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbSCommodity_GotFocus()
tmp = CmbSCommodity.Text
Call TableMst_Commodity
CmbSCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
 CmbSCommodity.AddItem RsMst_Commodity!Commodity
 RsMst_Commodity.MoveNext
Next
CmbSCommodity.Text = tmp
End Sub

Private Sub CmbSCommodity_LostFocus()
If CmbSCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbSCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodity.SetFocus
   Exit Sub
End If
mSCommodityID = RsMst_Commodity!CommodityID
End Sub

'Private Sub CmbWHRCommodity_GotFocus()
'tmp = CmbWHRCommodity.Text
'Call TableMst_Commodity
'CmbWHRCommodity.Clear
'If RsMst_Commodity.RecordCount = 0 Then
'   MsgBox "No record found!!! "
'   Exit Sub
'End If
'For I = 1 To RsMst_Commodity.RecordCount
'    CmbWHRCommodity.AddItem RsMst_Commodity!Commodity
'    RsMst_Commodity.MoveNext
'Next
'CmbWHRCommodity.Text = tmp
'End Sub
''----
'Private Sub CmbWHRCommodity_LostFocus()
'If CmbWHRCommodity.Text = "" Then
'   Exit Sub
'End If
'RsMst_Commodity.MoveFirst
'RsMst_Commodity.Find "Commodity = '" & CmbWHRCommodity.Text & "'"
'If RsMst_Commodity.EOF Then
'   MsgBox "Invalid selection!!! "
'   CmbWHRCommodity.SetFocus
'   Exit Sub
'End If
'mCommodityID = RsMst_Commodity!CommodityID
'
'Call GetGridData(mExchangeTypeID, mCommodityID, mClientIDRow)
'
'End Sub

'------------

Private Sub CmbWHRCommodity_GotFocus()
If CmbCMPID.Text = "" Then
    MsgBox "Please Select CMP!!!!!"
    CmbCMPID.SetFocus
    Exit Sub
End If

tmp1 = CmbWHRCommodity.Text
'Call TableMst_Commodity
If Edit_Flg = "" Then
    tmp = "Select MCom.CommodityID, MCom.Commodity From Mst_Commodity MCom Where MCom.CommodityID = " & mCommodityID
Else
    tmp = "Select MCom.CommodityID, MCom.Commodity From Mst_Commodity MCom Inner Join Mst_LocationCOmmodityMapping MLCM on MLCM.CommodityID = MCom.CommodityID "
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MLCM.LocationID "
    tmp = tmp & " Where ML.SM_Prefix = '" & mSM_Prefix & "'"
End If

Call Tmp7Table

CmbWHRCommodity.Clear
If TmpRs7.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To TmpRs7.RecordCount
    CmbWHRCommodity.AddItem TmpRs7!Commodity
    TmpRs7.MoveNext
Next
CmbWHRCommodity.Text = tmp1
End Sub


Private Sub CmbWHRCommodity_LostFocus()
If CmbWHRCommodity.Text = "" Then
   mCommodityID = 0
   Exit Sub
End If
TmpRs7.MoveFirst
TmpRs7.Find "Commodity = '" & CmbWHRCommodity.Text & "'"
If TmpRs7.EOF Then
   MsgBox "Invalid selection!!! "
   CmbWHRCommodity.SetFocus
   Exit Sub
End If
mCommodityID = TmpRs7!CommodityID

Call GetGridData(mExchangeTypeID, mCommodityID, mClientIDRow)

End Sub

Private Sub GetGridData(mExchangeTypeID_, mCommodityID_, mClientIDRow_)
Dim mStorageFrom As Date
Dim mWHRDate As Date
If mLotWiseWHR = "0" Then
   Call Grid_Head
Else
   Call Grid_HeadNew
End If
If mExchangeTypeID_ = "" Then Exit Sub
If mCommodityID_ = "" Or mCommodityID_ = 0 Then Exit Sub
If mClientIDRow_ = "" Then Exit Sub
If IsNull(mCMPID) Then Exit Sub
If IsNull(mGodownID) Then Exit Sub
If mGodownID = 0 Then Exit Sub

mStorageFrom = StorageFrom.Value
mWHRDate = WHRDate.Value
ChkSelection = 0

If mLotWiseWHR = "0" Then
   TxtCMPID = ""
   TxtGodownNo = ""
   TxtLicenseNo = ""
   TxtLicenseDescription = ""
   TxtNoofBags = ""
   TxtEleWeight = ""
   TxtWHRNoOfBags = ""
   TxtWHRQty = ""
   mGSLForPledge = ""
End If
'Commeneted on 19th June for removal of date condition mail by Sankar sir
'If Gen_DBType = "MDB" Then
'    tmp = "SELECT Count(Txn_CxTF_Details.CxTFNo) AS CountOfCxTFNo"
'    tmp = tmp & " From Txn_CxTF_Details"
'    tmp = tmp & " WHERE (Txn_CxTF_Details.CxTFDate Between #" & _
'    Format(mStorageFrom, Gen_DatFormat) & "# And #" & Format(mWHRDate, Gen_DatFormat) & "#) AND (Txn_CxTF_Details.Flag Is Null OR Txn_CxTF_Details.Flag='') "
'    tmp = tmp & " And (Txn_CxTF_Details.ExchangeTypeID = " & mExchangeTypeID_ & " And Txn_CxTF_Details.CommodityID = " & mCommodityID_ & ")"
'    tmp = tmp & " And (Txn_CxTF_Details.TxnType = 'D') And Txn_CxTF_Details.CMPID = " & mCMPID
'
'Else
'    tmp = "SELECT Count(Txn_CxTF_Details.CxTFNo) AS CountOfCxTFNo"
'    tmp = tmp & " From Txn_CxTF_Details"
'    tmp = tmp & " WHERE (Txn_CxTF_Details.CxTFDate Between '" & _
'    Format(mStorageFrom, Gen_DatFormat) & "' And '" & Format(mWHRDate, Gen_DatFormat) & "') AND (Txn_CxTF_Details.Flag Is Null OR Txn_CxTF_Details.Flag='') "
'    tmp = tmp & " And (Txn_CxTF_Details.ExchangeTypeID = " & mExchangeTypeID_ & " And Txn_CxTF_Details.CommodityID = " & mCommodityID_ & ")"
'    tmp = tmp & " And (Txn_CxTF_Details.TxnType = 'D') And Txn_CxTF_Details.CMPID = " & mCMPID
'End If


If Gen_DBType = "MDB" Then
    tmp = "SELECT Count(a.CxTF_DetailsID) AS CountOfCxTFNo"
    tmp = tmp & " From Txn_CxTF_Details a "
    tmp = tmp & " Inner Join Txn_CxTF_ClientDetail b on a.CxTF_DetailsID = b.CxTF_DetailsID and a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType"
    tmp = tmp & " WHERE  (a.Flag Is Null OR a.Flag='') "
    tmp = tmp & " And (a.ExchangeTypeID = " & mExchangeTypeID_ & " And a.CommodityID = " & mCommodityID_ & ")"
    tmp = tmp & " And (a.TxnType = 'D') And a.CMPID = " & mCMPID & " And b.ClientID = " & mClientIDRow & " "

Else
    tmp = "SELECT Count(a.CxTF_DetailsID) AS CountOfCxTFNo"
    tmp = tmp & " From Txn_CxTF_Details a "
    tmp = tmp & " Inner Join Txn_CxTF_ClientDetail b on a.CxTF_DetailsID = b.CxTF_DetailsID and a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType"
    tmp = tmp & " WHERE  (a.Flag Is Null OR a.Flag='') "
    tmp = tmp & " And (a.ExchangeTypeID = " & mExchangeTypeID_ & " And a.CommodityID = " & mCommodityID_ & ")"
    tmp = tmp & " And (a.TxnType = 'D') And a.CMPID = " & mCMPID & " And b.ClientID = " & mClientIDRow & " "
End If



Call TmpTable

If Not IsNull(TmpRs.Fields(0)) Then
   If TmpRs.Fields(0) > 0 Then
      MsgBox "WHR can't be generated since all related CDTFs are not printed!!! "
      Exit Sub
   End If
   
End If


If Gen_DBType = "MDB" Then
    If mLotWiseWHR = "0" Then
        tmp = " SELECT sum(BalanceBags) BalanceBags, sum(BalanceWeight) BalanceWeight, LicenseNo, LicenseDesc, CMPName, CommodityID, CMPID, ClientID,  LicenseID, GodownNo, GodownID, AgreementEndDate from ("
        tmp = tmp & " SELECT Mst_GSL.DepositeBags AS BalanceBags, Mst_GSL.DepositeWeight AS BalanceWeight, "
        tmp = tmp & " Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName, "
        tmp = tmp & " Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientID,  Mst_License.LicenseID,Mst_Godown.GodownNo, "
        tmp = tmp & " Mst_Godown.GodownID,Mst_Godown.AgreementEndDate "
        tmp = tmp & " FROM Mst_CMP INNER JOIN (Mst_License INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID) ON Mst_License.LicenseID = Mst_Godown.LicenseID) ON Mst_CMP.CMPID = Mst_GSL.CMPID "
        tmp = tmp & " WHERE (((Mst_GSL.ClientID)=" & mClientIDRow & ") AND ((Mst_GSL.ExchangeTypeID)=" & mExchangeTypeID_ & ") AND ((Mst_GSL.CommodityID)=" & mCommodityID_ & ") AND "
        tmp = tmp & " ((Mst_GSL.CDTFDate) <=  #" & Format(mWHRDate, Gen_DatFormat) & "#)) "
        tmp = tmp & " And Mst_GSL.CMPID = " & mCMPID & " And MST_GSL.Status <> 'Z' And MST_GSL.MadeUpFlag = '0' And GSL.CxTF_DetailsID > 0 "
        'tmp = tmp & " GROUP BY Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName , "
        'tmp = tmp & " Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientIDRow,  Mst_License.LicenseID,Mst_Godown.GodownNo,Mst_Godown.GodownID,Mst_Godown.AgreementEndDate "
        If mSCMFlag = 1 Then
           tmp = tmp & " And Mst_GSL.SCMFlag = 1 "
        End If
        tmp = tmp & " Union All "
        tmp = tmp & " SELECT Mst_GSL.DepositeBags AS BalanceBags, Mst_GSL.DepositeWeight AS BalanceWeight, "
        tmp = tmp & " Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName, "
        tmp = tmp & " Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientID,  Mst_License.LicenseID,Mst_Godown.GodownNo, "
        tmp = tmp & " Mst_Godown.GodownID,Mst_Godown.AgreementEndDate "
        tmp = tmp & " FROM Mst_CMP INNER JOIN (Mst_License INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID) ON Mst_License.LicenseID = Mst_Godown.LicenseID) ON Mst_CMP.CMPID = Mst_GSL.CMPID "
        tmp = tmp & " Inner Join Txn_RRNPWRGSL TRPG on (TRPG.NewMst_GSL_ID  = GSL.Mst_GSL_ID Or TRPG.NewTransferMst_GSL_ID  = GSL.Mst_GSL_ID ) and GSL.SM_prefix =  TRPG.SM_Prefix And TRPG.ConvertedID in (Select ConvertedID from Txn_RRNPWR Where Flag = 'A') "
        tmp = tmp & " WHERE (((Mst_GSL.ClientID)=" & mClientIDRow & ") AND ((Mst_GSL.ExchangeTypeID)=" & mExchangeTypeID_ & ") AND ((Mst_GSL.CommodityID)=" & mCommodityID_ & ") AND "
        tmp = tmp & " ((Mst_GSL.CDTFDate) <=  #" & Format(mWHRDate, Gen_DatFormat) & "#)) "
        tmp = tmp & " And Mst_GSL.CMPID = " & mCMPID & " And MST_GSL.Status <> 'Z' And MST_GSL.MadeUpFlag = '0' And GSL.CxTF_DetailsID < 0 ) a "
        tmp = tmp & " GROUP BY LicenseNo, LicenseDesc, CMPName , "
        tmp = tmp & " CommodityID,  CMPID, ClientID,  LicenseID, GodownNo, GodownID, AgreementEndDate "
    End If
Else
    If mLotWiseWHR = "0" Then
        tmp = " SELECT sum(BalanceBags) BalanceBags, sum(BalanceWeight) BalanceWeight, sum(WHRBags) WHRBags, sum(WHRWeight) WHRWeight,  LicenseNo, LicenseDesc, CMPName, CommodityID, CMPID, ClientID,  LicenseID, GodownNo, GodownID, AgreementEndDate from ("
        tmp = tmp & " SELECT Mst_GSL.DepositeBags AS BalanceBags, Mst_GSL.DepositeWeight AS BalanceWeight, "
        tmp = tmp & " Mst_GSL.WHRBags AS WHRBags, Mst_GSL.WHRWeight AS WHRWeight, "
        tmp = tmp & " Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName, "
        tmp = tmp & " Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientID,  Mst_License.LicenseID,Mst_Godown.GodownNo, "
        tmp = tmp & " Mst_Godown.GodownID,Mst_Godown.AgreementEndDate "
        tmp = tmp & " FROM Mst_CMP INNER JOIN (Mst_License INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID) ON Mst_License.LicenseID = Mst_Godown.LicenseID) ON Mst_CMP.CMPID = Mst_GSL.CMPID "
        tmp = tmp & " WHERE (((Mst_GSL.ClientID)=" & mClientIDRow & ") AND ((Mst_GSL.ExchangeTypeID)=" & mExchangeTypeID_ & ") AND ((Mst_GSL.CommodityID)=" & mCommodityID_ & ") AND "
        tmp = tmp & " ((Mst_GSL.CDTFDate) Between '" & Format(mStorageFrom, Gen_DatFormat) & "' And '" & Format(mWHRDate, Gen_DatFormat) & "')) "
        tmp = tmp & " And Mst_GSL.CMPID = " & mCMPID
        tmp = tmp & " And MST_GSL.Status = 'E' "
        tmp = tmp & " And MST_GSL.CxTF_DetailsID > 0 "
        tmp = tmp & " And MST_GSL.MadeUpFlag = '0' And Mst_Godown.GodownID = " & mGodownID & " "
        'tmp = tmp & " GROUP BY Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName , "
        'tmp = tmp & " Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientIDRow,  Mst_License.LicenseID,Mst_Godown.GodownNo,Mst_Godown.GodownID,Mst_Godown.AgreementEndDate "
        If mSCMFlag = 1 Then
           tmp = tmp & " And Mst_GSL.SCMFlag = 1 "
        End If
        tmp = tmp & " Union All"
        tmp = tmp & " SELECT Mst_GSL.DepositeBags AS BalanceBags, Mst_GSL.DepositeWeight AS BalanceWeight, "
        tmp = tmp & " Mst_GSL.WHRBags AS WHRBags, Mst_GSL.WHRWeight AS WHRWeight, "
        tmp = tmp & " Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName, "
        tmp = tmp & " Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientID,  Mst_License.LicenseID,Mst_Godown.GodownNo, "
        tmp = tmp & " Mst_Godown.GodownID,Mst_Godown.AgreementEndDate "
        tmp = tmp & " FROM Mst_CMP INNER JOIN (Mst_License INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID) ON Mst_License.LicenseID = Mst_Godown.LicenseID) ON Mst_CMP.CMPID = Mst_GSL.CMPID "
        tmp = tmp & " WHERE (((Mst_GSL.ClientID)=" & mClientIDRow & ") AND ((Mst_GSL.ExchangeTypeID)=" & mExchangeTypeID_ & ") AND ((Mst_GSL.CommodityID)=" & mCommodityID_ & ") AND "
        tmp = tmp & " ((Mst_GSL.CDTFDate) Between '" & Format(mStorageFrom, Gen_DatFormat) & "' And '" & Format(mWHRDate, Gen_DatFormat) & "')) "
        tmp = tmp & " And Mst_GSL.CMPID = " & mCMPID
        tmp = tmp & " And MST_GSL.Status = 'O' "
        tmp = tmp & " And MST_GSL.CxTF_DetailsID > 0 "
        tmp = tmp & " And MST_GSL.MadeUpFlag = '0' And Mst_Godown.GodownID = " & mGodownID & " And MST_GSL.WHRBags > 0 And MST_GSL.WHRBags < MST_GSL.DepositeBags  "
        'tmp = tmp & " GROUP BY Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName , "
        'tmp = tmp & " Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientIDRow,  Mst_License.LicenseID,Mst_Godown.GodownNo,Mst_Godown.GodownID,Mst_Godown.AgreementEndDate "
        If mSCMFlag = 1 Then
           tmp = tmp & " And Mst_GSL.SCMFlag = 1 "
        End If
        tmp = tmp & " Union All"
        tmp = tmp & " Select GSL.DepositeBags AS BalanceBags, GSL.DepositeWeight AS BalanceWeight, GSL.WHRBags , "
        tmp = tmp & " GSL.WHRWeight , MLC.LicenseNo, MLC.LicenseDesc, CMP.CMPName, GSL.CommodityID,  GSL.CMPID, GSL.ClientID,  MLC.LicenseID,MG.GodownNo, MG.GodownID,MG.AgreementEndDate "
        tmp = tmp & " From Mst_GSL  GSL "
        tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = GSL.GodownID "
        tmp = tmp & " Inner Join Mst_CMP CMP ON CMP.CMPID = GSL.CMPID "
        tmp = tmp & " Inner Join Mst_License MLC ON MLC.LicenseID = MG.LicenseID"
        tmp = tmp & " Inner Join Txn_RRNPWRGSL TRPG on (TRPG.NewMst_GSL_ID  = GSL.Mst_GSL_ID Or TRPG.NewTransferMst_GSL_ID  = GSL.Mst_GSL_ID ) and GSL.SM_prefix =  TRPG.SM_Prefix And TRPG.ConvertedID in (Select ConvertedID from Txn_RRNPWR Where Flag = 'A') "
        tmp = tmp & " Where GSL.ClientID = " & mClientIDRow & " And GSL.ExchangeTypeID = " & mExchangeTypeID_ & " And GSL.CommodityID = " & mCommodityID_
        tmp = tmp & " AND GSL.CDTFDate Between '" & Format(mStorageFrom, Gen_DatFormat) & "' And '" & Format(mWHRDate, Gen_DatFormat) & "' "
        tmp = tmp & " And GSL.CMPID = " & mCMPID & "  And GSL.Status = 'E'  And GSL.MadeUpFlag = '0'"
        tmp = tmp & " And GSL.CxTF_DetailsID < 0 And GSL.GodownID = " & mGodownID & "   "
        tmp = tmp & " Union All"
        tmp = tmp & " Select GSL.DepositeBags AS BalanceBags, GSL.DepositeWeight AS BalanceWeight, GSL.WHRBags , "
        tmp = tmp & " GSL.WHRWeight , MLC.LicenseNo, MLC.LicenseDesc, CMP.CMPName, GSL.CommodityID,  GSL.CMPID, GSL.ClientID,  MLC.LicenseID,MG.GodownNo, MG.GodownID,MG.AgreementEndDate "
        tmp = tmp & " From Mst_GSL  GSL "
        tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = GSL.GodownID "
        tmp = tmp & " Inner Join Mst_CMP CMP ON CMP.CMPID = GSL.CMPID "
        tmp = tmp & " Inner Join Mst_License MLC ON MLC.LicenseID = MG.LicenseID"
        tmp = tmp & " Inner Join Txn_RRNPWRGSL TRPG on (TRPG.NewMst_GSL_ID  = GSL.Mst_GSL_ID Or TRPG.NewTransferMst_GSL_ID  = GSL.Mst_GSL_ID ) and GSL.SM_prefix =  TRPG.SM_Prefix And TRPG.ConvertedID in (Select ConvertedID from Txn_RRNPWR Where Flag = 'A') "
        tmp = tmp & " Where GSL.ClientID = " & mClientIDRow & " And GSL.ExchangeTypeID = " & mExchangeTypeID_ & " And GSL.CommodityID = " & mCommodityID_
        tmp = tmp & " AND GSL.CDTFDate Between '" & Format(mStorageFrom, Gen_DatFormat) & "' And '" & Format(mWHRDate, Gen_DatFormat) & "' "
        tmp = tmp & " And GSL.CMPID = " & mCMPID & "  And GSL.Status <> 'O'  And GSL.MadeUpFlag = '0'"
        tmp = tmp & " And GSL.CxTF_DetailsID < 0 And GSL.GodownID = " & mGodownID & " And GSL.WHRBags > 0 And GSL.WHRBags < GSL.DepositeBags) A "
        tmp = tmp & " GROUP BY LicenseNo, LicenseDesc, CMPName , CommodityID,  CMPID, ClientID,  LicenseID, GodownNo, GodownID, AgreementEndDate "
        'tmp = tmp & " GROUP BY MLC.LicenseNo, MLC.LicenseDesc, CMP.CMPName , GSL.CommodityID, GSL.CMPID, GSL.ClientIDRow,  MLC.LicenseID, MG.GodownNo,MG.GodownID,MG.AgreementEndDate"

    Else
        tmp = "Select Mst_GSL_ID,SM_Prefix, StackNo,LotNo,WHRBags,CDTFBags,BalanceBags,WHRWeight,CDTFWeight,BalanceWeight From ("
        tmp = tmp & "Select Mst_GSL_ID,SM_Prefix, StackNo,LotNo,WHRBags,CDTFBags,BalanceBags,WHRWeight,CDTFWeight,BalanceWeight "
        tmp = tmp & " from Mst_GSL Where CxTF_DetailsID > 0 And MadeUpFlag = '0' And GodownID = " & mGodownID & " And ClientID = " & mClientIDRow & " "
        tmp = tmp & " And CommodityID = " & mCommodityID_ & " And Status = 'O' ANd ExchangeTypeID = " & mExchangeTypeID_ & " And WHRBags = 0 "
        tmp = tmp & " And CDTFDate Between '" & Format(mStorageFrom, Gen_DatFormat) & "' And '" & Format(mWHRDate, Gen_DatFormat) & "' And WHRNo is null"
        If mSCMFlag = 1 Then
           tmp = tmp & " And SCMFlag = '1' "
        End If
        tmp = tmp & " Union All"
        tmp = tmp & " Select a.Mst_GSL_ID,a.SM_Prefix, a.StackNo,a.LotNo,a.WHRBags,a.CDTFBags,a.BalanceBags,a.WHRWeight,a.CDTFWeight,a.BalanceWeight from Mst_GSL a"
        tmp = tmp & " Inner Join Txn_RRNPWRGSL TRPG on (TRPG.NewMst_GSL_ID  = a.Mst_GSL_ID Or TRPG.NewTransferMst_GSL_ID  = a.Mst_GSL_ID ) and a.SM_prefix =  TRPG.SM_Prefix   And TRPG.ConvertedID in (Select ConvertedID from Txn_RRNPWR Where Flag = 'A')"
        tmp = tmp & " Where CxTF_DetailsID < 0 And MadeUpFlag = '0' And GodownID = " & mGodownID & " And ClientID = " & mClientIDRow & " And CommodityID = " & mCommodityID_ & " And Status = 'O' "
        tmp = tmp & " ANd ExchangeTypeID = " & mExchangeTypeID_ & " And WHRBags = 0 And CDTFDate Between '" & Format(mStorageFrom, Gen_DatFormat) & "' And '" & Format(mWHRDate, Gen_DatFormat) & "' And WHRNo is null "
        If mSCMFlag = 1 Then
          tmp = tmp & " And a.SCMFlag = '1' "
        End If
        tmp = tmp & ") A Order By StackNo,LotNo "
    End If

End If

Call TmpTable


If TmpRs.RecordCount = 0 Then
   MsgBox "No deposits available for generating WHR!!!"
   Exit Sub
End If

k = 1
Dim BalBag, BalWt As Variant
For I = 0 To TmpRs.RecordCount - 1
    
'    tmp = " SELECT  Sum(Txn_WHR_Detail.WHRNoOfBags) AS WHRNoOfBags,Sum(Txn_WHR_Detail.WHRQuantity) As WHRQuantity "
'    tmp = tmp & " From Txn_WHR_Detail "
'    tmp = tmp & " where Txn_WHR_Detail.CommodityID = " & TmpRs!CommodityID
'    tmp = tmp & " And Txn_WHR_Detail.CMPID =" & TmpRs!CMPID
'    tmp = tmp & " And Txn_WHR_Detail.ClientIDRow =" & TmpRs!ClientIDRow
'    tmp = tmp & " And Txn_WHR_Detail.LicenseID =" & TmpRs!LicenseID
'    tmp = tmp & " And Txn_WHR_Detail.GodownID =" & TmpRs!GodownID
'    tmp = tmp & " And Txn_WHR_Detail.WHRDate <= '" & Format(mWHRDate, Gen_DatFormat) & "'"
'    Call Tmp1Table
    
    If mLotWiseWHR = "0" Then
       BalBag = TmpRs!BalanceBags - IIf(IsNull(TmpRs!WHRBags), 0, TmpRs!WHRBags)
       BalWt = TmpRs!BalanceWeight - IIf(IsNull(TmpRs!WHRWeight), 0, TmpRs!WHRWeight)
    
       If BalBag <> 0 And BalWt <> 0 Then
          k = MSHFlexGrid2.Rows - 1
          MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
          MSHFlexGrid2.TextMatrix(k, 1) = TmpRs!CMPName
          MSHFlexGrid2.TextMatrix(k, 2) = TmpRs!GodownNo
          MSHFlexGrid2.TextMatrix(k, 3) = TmpRs!LicenseNo
          MSHFlexGrid2.TextMatrix(k, 4) = TmpRs!LicenseDesc
          MSHFlexGrid2.TextMatrix(k, 5) = BalBag
          MSHFlexGrid2.TextMatrix(k, 6) = BalWt
          MSHFlexGrid2.TextMatrix(k, 7) = ""
          MSHFlexGrid2.TextMatrix(k, 11) = IIf(IsNull(TmpRs!AgreementEndDate), "", TmpRs!AgreementEndDate)
          MSHFlexGrid2.TextMatrix(k, 12) = GetPreFix(IIf(IsNull(TmpRs!CMPID), "", TmpRs!CMPID), "P")
      End If
  Else
      k = MSHFlexGrid4.Rows - 1
      MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
      MSHFlexGrid4.TextMatrix(k, 1) = TmpRs!StackNo
      MSHFlexGrid4.TextMatrix(k, 2) = TmpRs!LotNo
      MSHFlexGrid4.TextMatrix(k, 3) = TmpRs!BalanceBags
      MSHFlexGrid4.TextMatrix(k, 4) = TmpRs!BalanceWeight
      MSHFlexGrid4.TextMatrix(0, 3) = Val(MSHFlexGrid4.TextMatrix(0, 3)) + Val(MSHFlexGrid4.TextMatrix(k, 3))
      MSHFlexGrid4.TextMatrix(0, 4) = Val(MSHFlexGrid4.TextMatrix(0, 4)) + Val(MSHFlexGrid4.TextMatrix(k, 4))
      If mSCMFlag = 0 Then
         MSHFlexGrid4.TextMatrix(k, 5) = 0
         MSHFlexGrid4.TextMatrix(k, 6) = 0
      Else
         MSHFlexGrid4.TextMatrix(k, 5) = Val(MSHFlexGrid4.TextMatrix(k, 3))
         MSHFlexGrid4.TextMatrix(k, 6) = Val(MSHFlexGrid4.TextMatrix(k, 4))
         MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(k, 5))
         MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(k, 6))
      End If
      MSHFlexGrid4.TextMatrix(k, 7) = TmpRs!CDTFWeight
      MSHFlexGrid4.TextMatrix(k, 8) = IIf(IsNull(TmpRs!Mst_GSL_ID), "", TmpRs!Mst_GSL_ID)
      MSHFlexGrid4.TextMatrix(k, 9) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
      
      MSHFlexGrid4.row = k
      MSHFlexGrid4.Col = 0
      MSHFlexGrid4.CellFontName = "Wingdings"
      MSHFlexGrid4.CellFontSize = 12
      MSHFlexGrid4.CellAlignment = flexAlignCenterCenter
      If mSCMFlag = 0 Then
         MSHFlexGrid4.Text = strUnChecked
      Else
         MSHFlexGrid4.Text = strChecked
      End If
  End If
  TmpRs.MoveNext
Next

End Sub

Private Sub CmdCancel_Click()
TxtHologramNo = ""
Frame3.Visible = False
End Sub

Private Sub CmdInsurance_Click()
FrmMstInsurance.Show
End Sub

Private Sub CmdInsuranceCo_Click()
FrmMstInsuranceCo.Show
End Sub

Private Sub CmdPCancel_Click()
Frame5.Visible = False
End Sub

Private Sub CmdPledge_Click()

Gen_WHRDetail(0) = False
Gen_WHRDetail(1) = TxtWHRNo
Gen_WHRDetail(2) = "1"
Gen_WHRDetail(3) = TxtWHRNoNew
Gen_WHRDetail(4) = "Non NCDEX"

Frm_Pledge_Creation.Show

End Sub

Private Sub CmdPOk_Click()
If OptLetter.Value = True Then
   Call PDF_Print_RptLetterHead
Else
   Call PDF_Print_Rpt
End If
If PntPrv <> "Y" Then
   If mNCMLFinance = 1 Then
   
      Call TableLog_WHRPrint
   
      RsLog_WHRPrint.AddNew
      RsLog_WHRPrint!WHRNo = RsTxn_WHRDetail!WHRNo
      RsLog_WHRPrint!SM_Prefix = RsTxn_WHRDetail!SM_Prefix
      RsLog_WHRPrint!PrintedDate = Date
      RsLog_WHRPrint!HologramNo = "0" ' TxtHologramNo
      RsLog_WHRPrint!Version = "Pledgable"
      RsLog_WHRPrint!Txn_PledgeCreation_ID = IIf(IsNull(RsTxn_WHRDetail!Txn_PledgeCreation_ID), 0, RsTxn_WHRDetail!Txn_PledgeCreation_ID)
      RsLog_WHRPrint!UpdatedDate = Now
      RsLog_WHRPrint!UpdatedBy = Gen_UserLoginID
      RsLog_WHRPrint.Update
      
      Call TableTxn_WHRToNFin(" Where WHRNo = " & RsTxn_WHRDetail!WHRNo & " And Txn_PledgeCreation_ID = " & RsTxn_WHRDetail!Txn_PledgeCreation_ID & " ")
      
      If RsTxn_WHRToNFin.RecordCount = 0 Then
         RsTxn_WHRToNFin.AddNew
         RsTxn_WHRToNFin!SM_Prefix = RsTxn_WHRDetail!SM_Prefix
         RsTxn_WHRToNFin!WHRNo = RsTxn_WHRDetail!WHRNo
         RsTxn_WHRToNFin!Txn_PledgeCreation_ID = RsTxn_WHRDetail!Txn_PledgeCreation_ID
         RsTxn_WHRToNFin!Txn_WHR_Detail_ID = RsTxn_WHRDetail!Txn_WHR_Detail_ID
         RsTxn_WHRToNFin!StatusFlag = "P"
         RsTxn_WHRToNFin!CreatedDate = Now
         RsTxn_WHRToNFin!CreatedBy = Gen_UserLoginID
         RsTxn_WHRToNFin.Update
      End If
   End If
   Call Gen_Email
End If
Frame5.Visible = False

MsgBox "The WHR in PDF format is sent to your mailbox. You may download the PDF file and print"
End Sub

Private Sub CmdPrint_Click()
Dim mDRCName As String
'Check2.Visible = False
'If LCase(UserType) = "power" Then
'   Check2.Visible = True
'End If
'
'
'If RsTxn_WHRDetail!HologramNo = "" Or IsNull(RsTxn_WHRDetail!HologramNo) Then
'   If PntPrv = "N" Then
'      ans = MsgBox("We strongly recommend that you take a look at Print-Preview first before going ahead with this Print option. Do you still want to continue?", vbYesNo)
'      If ans = vbNo Then
'         Exit Sub
'      End If
'   End If
'
'   Frame3.Visible = True
'   TxtHologramNo.SetFocus
'   Call TableMst_Hologram
'   Exit Sub
'Else
'   TxtHologramNo = ""
'   Frame3.Visible = False
'End If
'
'RsTxn_WHRDetail!Flag = "P"
'RsTxn_WHRDetail.Update
'
'Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_WHR" & Trim(TxtWHRNoNew) & ".xls")
'Call Xls_Print_Rpt_New

Check2.Visible = False
If LCase(UserType) = "power" Then
   Check2.Visible = True
End If

PntPrv = "N"

If mNCMLFinance = 0 Then
   If RsTxn_WHRDetail!HologramNo = "" Or IsNull(RsTxn_WHRDetail!HologramNo) Then
      If PntPrv = "N" Then
         ans = MsgBox("We strongly recommend that you take a look at Print-Preview first before going ahead with this Print option. Do you still want to continue?", vbYesNo)
         If ans = vbNo Then
            Exit Sub
         End If
      End If
   End If
End If

''If RsTxn_WHRDetail!HologramNo = "" Or IsNull(RsTxn_WHRDetail!HologramNo) Then
''   If PntPrv = "N" Then
''      ans = MsgBox("We strongly recommend that you take a look at Print-Preview first before going ahead with this Print option. Do you still want to continue?", vbYesNo)
''      If ans = vbNo Then
''         Exit Sub
''      End If
''   End If
''
''   Frame3.Visible = True
''   TxtHologramNo.SetFocus
''   Call TableMst_Hologram
''   Exit Sub
''Else
''   TxtHologramNo = ""
''   Frame3.Visible = False
''   Frame5.Visible = True
''End If
''
''RsTxn_WHRDetail!Flag = "P"
''RsTxn_WHRDetail.Update

mPrintFlag = ""
mWHRVersion = ""
''tmp = "SELECT  TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity, sum(isnull(ReleasedQuantity,0))ReleasedQuantity"
''tmp = tmp & " FROM  Txn_WHR_Detail TWD"
''tmp = tmp & " inner join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
''tmp = tmp & " left Join Txn_PledgeRelease TPR on TPR.PledgeID = TPC.PledgeID And TPC.SM_Prefix = TPR.SM_Prefix"
''tmp = tmp & " left Join Txn_PledgeRelease_GSL TPRG on TPRG.PledgeReleaseID = TPR.PledgeReleaseID And TPRG.SM_Prefix = TPR.SM_Prefix"
''tmp = tmp & " Where TWD.WHRNo = " & RsTxn_WHRDetail!WHRNO
''tmp = tmp & " Group By TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity"
''tmp = tmp & " Having TWD.WHRQuantity > Sum(isnull(ReleasedQuantity,0))"

tmp = " SELECT  TWD.Txn_PledgeCreation_ID,TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity,isnull(sum(PledgeQuantity),0) PledgeQuantity,TPC.Flag,isnull(TPC.NCMLFinance,0)'NCMLFinance'"
tmp = tmp & " FROM  Txn_WHR_Detail TWD"
tmp = tmp & " left join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
tmp = tmp & " left Join Txn_PledgeCreation_GSL TPRG on TPRG.Txn_PledgeCreation_ID = TPC.Txn_PledgeCreation_ID And TPRG.SM_Prefix = TPC.SM_Prefix"
tmp = tmp & " Where TWD.WHRNo =" & RsTxn_WHRDetail!WHRNo & " And TPC.Txn_PledgeCreation_ID =" & IIf(IsNull(RsTxn_WHRDetail!Txn_PledgeCreation_ID), 0, RsTxn_WHRDetail!Txn_PledgeCreation_ID)
tmp = tmp & " Group By TWD.Txn_PledgeCreation_ID,TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity,TPC.Flag,TPC.NCMLFinance"
tmp = tmp & " Having IsNull(Sum(PledgeQuantity), 0) > 0"

Call Tmp2Table

If mSCMFlag = 0 Then
   If mNCMLFinance = 1 Then GoTo l
   If TmpRs2.RecordCount > 0 And Val(TxtCARNo) = 0 Then  '(IsNull(RsTxn_WHRDetail!CARNo) = True Or RsTxn_WHRDetail!CARNo <= 0) Then
      If IsNull(TmpRs2!Flag) = True Or TmpRs2!Flag = "" Then
         MsgBox "Pledge letter print pending. Can't print WHR!!!"
         Exit Sub
      End If
   
      tmp = "Select MSD.EmployeeID From Mst_StateDRCMapping MSD"
      tmp = tmp & " Where StateID = (Select StateID from Mst_Location Where SM_Prefix='" & RsTxn_WHRDetail!SM_Prefix & "')"
   
      Call Tmp3Table
      mDRCName = ""
      For I = 1 To TmpRs3.RecordCount
          tmp = " Select ME.EmployeeName from Mst_Employee ME"
          tmp = tmp & " Where EmployeeID = " & TmpRs3!EmployeeID & " "
          
          Call TmpTableCMPro
       
          If mDRCName = "" Then
             mDRCName = TmpRsCMPro!EmployeeName 'TmpRs3!EmployeeName
          Else
             mDRCName = mDRCName & "," & TmpRsCMPro!EmployeeName ' TmpRs3!EmployeeName
          End If
          TmpRs3.MoveNext
      Next
      If mDRCName <> "" Then
         MsgBox "No SR present for this WHR in CM-Pro, Contact to DRC " & mDRCName & ". you can't print!!"
      Else
         MsgBox "No SR present for this WHR in CM-Pro, Contact to DRC. Also map CM-DRC in Store-Man. You can't print!!"
      End If
      Exit Sub
   End If
End If
l:
If TmpRs2.RecordCount > 0 And IsNull(RsTxn_WHRDetail!Flag) = True Then
   mPrintFlag = "N"
   mWHRVersion = "Pledgable"
   If mNCMLFinance = 0 Then
      Frame3.Visible = True
      TxtHologramNo.SetFocus
      Call TableMst_Hologram
      Exit Sub
   Else
      RsTxn_WHRDetail!Flag = "N"
      RsTxn_WHRDetail.Update
   End If
ElseIf TmpRs2.RecordCount > 0 And RsTxn_WHRDetail!Flag = "N" Then
   TxtHologramNo = ""
   Frame3.Visible = False
   Frame5.Visible = True
   'RsTxn_WHRDetail!Flag = "P"
   'RsTxn_WHRDetail.Update
ElseIf TmpRs2.RecordCount > 0 And RsTxn_WHRDetail!Flag = "P" Then
   mPrintFlag = "N"
   mWHRVersion = "Pledgable"
   If mNCMLFinance = 0 Then
      Frame3.Visible = True
      TxtHologramNo.SetFocus
      Call TableMst_Hologram
      Exit Sub
   Else
      RsTxn_WHRDetail!Flag = "N"
      RsTxn_WHRDetail.Update
   End If
'   TxtHologramNo = ""
'   Frame3.Visible = False
'   Frame5.Visible = True
   
   
ElseIf TmpRs2.RecordCount = 0 And (IsNull(RsTxn_WHRDetail!Flag) = True Or RsTxn_WHRDetail!Flag = "N") Then
   mPrintFlag = "P"
   mWHRVersion = "Non Pledgable"
   If mNCMLFinance = 0 Then
      Frame3.Visible = True
      TxtHologramNo.SetFocus
      Call TableMst_Hologram
      Exit Sub
   Else
      RsTxn_WHRDetail!Flag = "P"
      RsTxn_WHRDetail.Update
   End If
ElseIf TmpRs2.RecordCount = 0 And RsTxn_WHRDetail!Flag = "P" Then
   TxtHologramNo = ""
   Frame3.Visible = False
   Frame5.Visible = True
   RsTxn_WHRDetail!Flag = "P"
   RsTxn_WHRDetail.Update
End If

If mNCMLFinance = 1 Then
   'Frame5.Visible = True
   If Frame5.Visible = False Then Frame5.Visible = True
End If

End Sub

Public Sub Create_Xls(FileName As String)
'Dim oExcel As Object
'Dim oBook As Object
'Dim oSheet As Object
'Dim Pat As String
Pat = App.Path
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add
Gen_FLName = FileName
oBook.SaveAs (Pat & "\excel\" & FileName)

Set oSheet = oBook.Worksheets(1)
With oExcel
'.Application.Sheets.Delete
.SheetsInNewWorkbook = 15 'MSHFlexGrid2.Rows - 2
.Workbooks.Add
End With
'oBook.SaveAs (Pat & "\" & FileName)

oExcel.Quit
'oExcel.Close
End Sub


'-----WHR Function-----

Public Sub Xls_Print_Rpt()
'Dim oXL As Excel.Application
'Dim oWB As Excel.Workbook
'Dim oWS As Excel.Worksheet
'Dim mRow As Integer
'Dim mCol As Integer
Dim ans As Variant
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_UserName & "_WHR" & Trim(TxtWHRNoNew) & ".xls")
 
I = 2
Set oWS = oWB.Worksheets("Sheet" & 1)
'oWS.PageSetup.RightMargin = 0.05
oWS.Name = TxtWHRNoNew  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 4

oWS.Range("a1:f1").Merge
oWS.Range("a:a").ColumnWidth = 10.86
oWS.Range("b:b").ColumnWidth = 17.71
oWS.Range("c:c").ColumnWidth = 17.14
oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 12.86 '32.14
oWS.Range("f:f").ColumnWidth = 17.14

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = GetCMPName(RsTxn_WHRDetail!CMPID) & "---" & GetCMPAddress(RsTxn_WHRDetail!CMPID, "A")
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).RowHeight = 24
mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Warehouse Receipt (WHR)"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).RowHeight = 24

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Warehouse License No"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = MSHFlexGrid2.TextMatrix(I, 3)
oWS.Cells(mRow, C + 2).HorizontalAlignment = 2
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 3) = "License Expiry Date"
oWS.Cells(mRow, C + 3).Font.Bold = True
oWS.Cells(mRow, C + 3).WrapText = True
oWS.Cells(mRow, C + 3).Font.Size = 7
oWS.Cells(mRow, C + 4) = Format(RsMst_License!LicenseExpiryDate, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 4).HorizontalAlignment = 2
oWS.Cells(mRow, C + 4).Font.Size = 7
'j = 23
'xlsdata (mRow)
'oWS.Range("d5:e5").Merge
oWS.Cells(mRow, C + 5) = "Receipt No"
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True
'oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = TxtWHRNoNew
oWS.Cells(mRow, C + 6).HorizontalAlignment = 2
oWS.Cells(mRow, C + 6).Font.Size = 7
'xlsdata (mRow)

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Location & Address of Warehouse" '& " " & Trim(TxtBankName) & " - " & TxtBranchName
'oWS.Range("b8:e8").Merge
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Bold = True
tmp = "b" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = GetLocationFromCMP(RsMst_Godown!CMPID) & " - " & GetGodownName(RsTxn_WHRDetail!GodownID, "A")
oWS.Cells(mRow, C + 2).Font.Size = 7
'oWS.Cells(mRow, c + 2).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 5) = "Godown No"
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True
'oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = MSHFlexGrid2.TextMatrix(I, 2)
oWS.Cells(mRow, C + 6).HorizontalAlignment = 2
oWS.Cells(mRow, C + 6).Font.Size = 7


mRow = mRow + 1
' tmp = "a" & mRow & ":" & "E" & mRow
' oWS.Range(tmp).Merge
' oWS.Cells(mRow, c + 1).RowHeight = 8
' mRow = mRow + 1

'oWS.Range("a9:e9").Merge
oWS.Cells(mRow, C + 1) = "Name and Address of the Depositor"
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).WrapText = True
tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
ans = GetClientNameByRow(mClientIDRow)

If ans(4) <> "" Then
   oWS.Cells(mRow, C + 2) = TxtClientID & "-" & Trim(ans(4))
Else
   oWS.Cells(mRow, C + 2) = TxtClientID
End If
oWS.Cells(mRow, C + 2).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Commodity"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = "No of Bags"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 3) = "Net Quantity(MT)"
oWS.Cells(mRow, C + 3).Font.Bold = True
oWS.Cells(mRow, C + 3).Font.Size = 7
oWS.Cells(mRow, C + 4) = "Name and License No of Weigher"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 5) = "Private Marks on Packaging"
oWS.Cells(mRow, C + 5).WrapText = True
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 6) = "Condition of Commodity"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).WrapText = True
oWS.Cells(mRow, C + 6).Font.Size = 7
 
mRow = mRow + 1
oWS.Cells(mRow, C + 1) = TxtWHRCommodityId
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = Format(MSHFlexGrid2.TextMatrix(I, 7), "#######0.000") 'MSHFlexGrid2.TextMatrix(i, 7)
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 3) = Format(MSHFlexGrid2.TextMatrix(I, 8), "#######0.000") 'MSHFlexGrid2.TextMatrix(i, 8)
oWS.Cells(mRow, C + 3).Font.Size = 7
oWS.Cells(mRow, C + 4) = TxtWeigher
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 5) = TxtGrade
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 6) = CmbCommodityCondition
oWS.Cells(mRow, C + 6).Font.Size = 7
mRow = mRow + 1

tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Name & License No of Assayer"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = ""
oWS.Cells(mRow, C + 5) = "Quality As per Grader"
oWS.Cells(mRow, C + 5).WrapText = True
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 6) = "Grade Valid Till"
oWS.Cells(mRow, C + 6).Font.Size = 7
oWS.Cells(mRow, C + 6).Font.Bold = True
mRow = mRow + 1

tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = TxtAssayer
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 5) = TxtAssayerCerti
oWS.Cells(mRow, C + 5).WrapText = True
oWS.Cells(mRow, C + 5).Font.Size = 7
If TxtAssayer <> "" Then
   oWS.Cells(mRow, C + 6) = Format(GradeValidUpto, "dd-mmm-yyyy")
   oWS.Cells(mRow, C + 6).Font.Size = 7
End If
mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Storage Charge(Rate)-excluding Taxes(Rupees per bag/month)"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).RowHeight = 24

oWS.Cells(mRow, C + 3) = Format(TxtStorageCharges, "#######0.000") 'TxtStorageCharges
oWS.Cells(mRow, C + 3).Font.Size = 7
oWS.Cells(mRow, C + 3).RowHeight = 24

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Storage Charge - excluding Taxes (Total Amount) per month"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).RowHeight = 24

oWS.Cells(mRow, C + 6) = Format(TxtStorageAmount, "#######0.000")  'MSHFlexGrid2.TextMatrix(i, 7) * TxtStorageCharges
oWS.Cells(mRow, C + 6).Font.Size = 7
oWS.Cells(mRow, C + 6).RowHeight = 24
mRow = mRow + 1


tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Risk Coverd"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).RowHeight = 24
tmp = "c" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = "Insurance Co."
oWS.Cells(mRow, C + 3).Font.Bold = True
oWS.Cells(mRow, C + 3).WrapText = True
oWS.Cells(mRow, C + 3).Font.Size = 7
'oWS.Cells(mRow, c + 1).RowHeight = 8
tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "Amount Insured"
oWS.Cells(mRow, C + 5).WrapText = True
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Fire,Burglary,Spontaneous Combustion,EarthQuakes,Terrorism and allied causes"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
'oWS.Cells(mRow, c + 1).RowHeight = 24
tmp = "c" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = ""
oWS.Cells(mRow, C + 3).WrapText = True
oWS.Cells(mRow, C + 3).Font.Size = 7
'oWS.Cells(mRow, c + 1).RowHeight = 8
tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = ""
oWS.Cells(mRow, C + 5).WrapText = True
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).RowHeight = 30

mRow = mRow + 1


tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Goods are Accepted for storage From"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7

oWS.Cells(mRow, C + 3) = Format(StorageFrom, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "To"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7

oWS.Cells(mRow, C + 6) = Format(StorageTo, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Market Rate at the time of deposit(Rs/MT)"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).RowHeight = 10

oWS.Cells(mRow, C + 3) = Format(TxtRate, "#######0.000") 'TxtRate
oWS.Cells(mRow, C + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Value of Commodities at time of deposit"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).RowHeight = 20

oWS.Cells(mRow, C + 6) = Format(TxtValueofCommodity, "#######0.000") 'TxtValueofCommodity
oWS.Cells(mRow, C + 6).Font.Size = 7

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Date"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7

oWS.Cells(mRow, C + 2) = Format(WHRDate, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 2).Font.Size = 7

tmp = "c" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = "Signature & Stamp of Warehouse Manager"
oWS.Cells(mRow, C + 3).Font.Bold = True
oWS.Cells(mRow, C + 3).WrapText = True
oWS.Cells(mRow, C + 3).Font.Size = 7
oWS.Cells(mRow, C + 4).RowHeight = 30

tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = ""
oWS.Cells(mRow, C + 5).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).RowHeight = 8

mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Date"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "Name and Address Of Lender"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).WrapText = True
oWS.Cells(mRow, C + 2).Font.Size = 7

oWS.Cells(mRow, C + 4) = "Pledge Number"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7

oWS.Cells(mRow, C + 5) = "Pledge Released on"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).WrapText = True
oWS.Cells(mRow, C + 5).Font.Size = 7

oWS.Cells(mRow, C + 6) = "Sign of Lender/Warehouse Manager"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).WrapText = True
oWS.Cells(mRow, C + 6).Font.Size = 7

mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Withdrawal History"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).RowHeight = 24

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "The good mentioned below are hereby released for delivery from the" & _
                         " Warehouse. The unreleased balance is subject to lien for" & _
                         " unpaid charges and borrowings from lending institutions."
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).WrapText = True
oWS.Cells(mRow, C + 2).Font.Size = 6

mRow = mRow + 1

tmp = "a" & mRow & ":" & "a" & (mRow + 1)
oWS.Range(tmp).Merge


oWS.Cells(mRow, C + 1) = "Date"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "Quantity Released"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).WrapText = True
oWS.Cells(mRow, C + 2).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Quantity Balance"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7

tmp = "f" & mRow & ":" & "f" & (mRow + 1)
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Sign & Stamp of the Warehouse Manager"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).WrapText = True
oWS.Cells(mRow, C + 6).Font.Size = 7


mRow = mRow + 1

oWS.Cells(mRow, C + 2) = "No. of Bags"
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 2).Font.Bold = True

oWS.Cells(mRow, C + 3) = "Weight in MT"
oWS.Cells(mRow, C + 3).Font.Size = 7
oWS.Cells(mRow, C + 3).Font.Bold = True

oWS.Cells(mRow, C + 4) = "No. Of Bags"
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).Font.Bold = True

oWS.Cells(mRow, C + 5) = "Weight in  MT"
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True

mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1

mRow = mRow + 1
tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).RowHeight = 10

mRow = mRow + 1
tmp = "d" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
'oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, C + 4) = "HOLOGRAM NO. :- " & "" & RsTxn_WHRDetail!HologramNo
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).RowHeight = 30



mRow = mRow + 1
tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).RowHeight = 10

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "The Warehouse Receipt is valid for transfer by endorsement, and/or for " & _
                         " marking of plegde, subject to confirmation by 'Warehouse Manager' "
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 6
oWS.Cells(mRow, C + 1).RowHeight = 18

k = mRow

For Cell = 4 To k
    tmp = "a" & Cell & ":" & "f" & Cell
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    oWS.Range(tmp).Font.Name = "Zurich BT"
    'oWS.Range(tmp).Font.Size = 7
    oWS.Range(tmp).WrapText = True
Next Cell
    
oWB.Save
'oWB.Close
'oXL.Quit
MsgBox ("Excel file created!!! ")
oXL.Visible = True

End Sub
Public Sub Xls_Print_Rpt_New()

Dim ans As Variant
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
'Dim Pat As String
Pat = App.Path
'Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_UserName & "_WHR" & Trim(TxtWHRNoNew) & ".xls")

Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "_WHR" & Trim(TxtWHRNoNew) & ".xls")
 

I = 2
Set oWS = oWB.Worksheets("Sheet" & 1)
'oWS.PageSetup.RightMargin = 0.05
oWS.Name = TxtWHRNoNew  'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 4

oWS.Range("a1:f1").Merge
oWS.Range("a:a").ColumnWidth = 10.86
oWS.Range("b:b").ColumnWidth = 17.71
oWS.Range("c:c").ColumnWidth = 17.14
oWS.Range("d:d").ColumnWidth = 12.86
oWS.Range("e:e").ColumnWidth = 12.86 '32.14
oWS.Range("f:f").ColumnWidth = 17.14

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = GetCMPAddress(RsTxn_WHRDetail!CMPID, "D") '"NCMSL" & "-" & GetCMPName(RsTxn_WHRDetail!CMPID)
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
'oWS.Cells(mRow, c + 1).RowHeight = 24
mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Warehouse Receipt"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 10
oWS.Cells(mRow, C + 1).RowHeight = 24

mRow = mRow + 1


tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Receipt No"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).Font.Bold = True
'oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = TxtWHRNoNew
oWS.Cells(mRow, C + 3).HorizontalAlignment = 2
oWS.Cells(mRow, C + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Receipt Date"
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).Font.Bold = True

oWS.Cells(mRow, C + 6).HorizontalAlignment = 2
oWS.Cells(mRow, C + 6) = Format(WHRDate, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Warehouse License No"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 3) = MSHFlexGrid2.TextMatrix(I, 3)
oWS.Cells(mRow, C + 3).HorizontalAlignment = 2
oWS.Cells(mRow, C + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "License Expiry Date"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7

oWS.Cells(mRow, C + 6) = Format(RsMst_License!LicenseExpiryDate, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 6).HorizontalAlignment = 2
oWS.Cells(mRow, C + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Location & Address of Warehouse" '& " " & Trim(TxtBankName) & " - " & TxtBranchName
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7
tmp = "c" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 3) = GetLocationFromCMP(RsMst_Godown!CMPID) '& " - " & GetGodownName(RsTxn_WHRDetail!GodownID, "A")
oWS.Cells(mRow, C + 3).Font.Size = 7
'oWS.Cells(mRow, c + 2).HorizontalAlignment = -4108

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Godown No" '& " " & Trim(TxtBankName) & " - " & TxtBranchName
'oWS.Range("b8:e8").Merge
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7
tmp = "c" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = GetGodownName(RsTxn_WHRDetail!GodownID, "N")
oWS.Cells(mRow, C + 3).Font.Size = 7
oWS.Cells(mRow, C + 3).HorizontalAlignment = 2

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 1) = "Godown Address"
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).Font.Bold = True
'oWS.Range(tmp).Merge
tmp = "c" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 3) = GetGodownName(RsTxn_WHRDetail!GodownID, "A")  'MSHFlexGrid2.TextMatrix(i, 2)
oWS.Cells(mRow, C + 3).HorizontalAlignment = 2
oWS.Cells(mRow, C + 3).Font.Size = 7


mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 1) = "Name and Address of the Depositor"
'oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).WrapText = True
tmp = "c" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
ans = GetClientNameByRow(mClientIDRow)

If ans(4) <> "" Then
   oWS.Cells(mRow, C + 3) = TxtClientID & "-" & Trim(ans(4))
Else
   oWS.Cells(mRow, C + 3) = TxtClientID
End If
oWS.Cells(mRow, C + 3).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 1) = "Commodity"
If ChkSaidToContain.Value = 1 Then
   oWS.Cells(mRow, C + 1) = "Commodity (Said to contain)"
End If
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7

oWS.Cells(mRow, C + 3) = TxtWHRCommodityId
oWS.Cells(mRow, C + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 4) = "Condition of Commodity"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 6) = CmbCommodityCondition
oWS.Cells(mRow, C + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "No of Bags"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 3) = Format(MSHFlexGrid2.TextMatrix(I, 7), "#######0.000") 'MSHFlexGrid2.TextMatrix(i, 7)
oWS.Cells(mRow, C + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Net Quantity(MT)"
If ChkEstimatedWgt.Value = 1 Then
   oWS.Cells(mRow, C + 4) = "Net Quantity(MT) (Estimated Weight)"
End If
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).RowHeight = 24

oWS.Cells(mRow, C + 6) = Format(MSHFlexGrid2.TextMatrix(I, 8), "#######0.000") 'MSHFlexGrid2.TextMatrix(i, 8)
oWS.Cells(mRow, C + 6).Font.Size = 7


mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 1) = "Name and License No of Weigher"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 3) = TxtWeigher
oWS.Cells(mRow, C + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Private Marks on Packaging"
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 6) = TxtGrade
oWS.Cells(mRow, C + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Name & License No of Assayer"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
tmp = "c" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 3) = TxtAssayer
oWS.Cells(mRow, C + 3).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Quality As per Grader"
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 3) = TxtAssayerCerti
oWS.Cells(mRow, C + 3).WrapText = True
oWS.Cells(mRow, C + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, C + 4) = "Grade Valid Till"
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).Font.Bold = True

If TxtAssayer <> "" Then
   oWS.Cells(mRow, C + 6) = Format(GradeValidUpto, "dd-mmm-yyyy")
   oWS.Cells(mRow, C + 6).Font.Size = 7
End If


mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
If OptSRPerBag.Value = True Then
   oWS.Cells(mRow, C + 1) = "Storage Charge(Rate)-excluding Taxes(Rupees per bag/month)"
Else
   oWS.Cells(mRow, C + 1) = "Storage Charge(Rate)-excluding Taxes(Rupees per MT/month)"
End If
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).RowHeight = 24

oWS.Cells(mRow, C + 3) = Format(TxtStorageCharges, "#########0.00") 'TxtStorageCharges
oWS.Cells(mRow, C + 3).Font.Size = 7
oWS.Cells(mRow, C + 3).RowHeight = 24

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Storage Charge - excluding Taxes (Total Amount) per month"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).RowHeight = 24

oWS.Cells(mRow, C + 6) = Format(TxtStorageAmount, "#########0.00")  'MSHFlexGrid2.TextMatrix(i, 7) * TxtStorageCharges
oWS.Cells(mRow, C + 6).Font.Size = 7
oWS.Cells(mRow, C + 6).RowHeight = 24
mRow = mRow + 1


tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Goods are Accepted for storage From"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7

oWS.Cells(mRow, C + 3) = Format(StorageFrom, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "To"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7

oWS.Cells(mRow, C + 6) = Format(StorageTo, "dd-mmm-yyyy")
oWS.Cells(mRow, C + 6).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Market Rate at the time of deposit(Rs/MT)"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).RowHeight = 20

oWS.Cells(mRow, C + 3) = Format(TxtRate, "#######0.000") 'TxtRate
oWS.Cells(mRow, C + 3).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Value of Commodities at time of deposit (Amount in Rs)"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).RowHeight = 20

oWS.Cells(mRow, C + 6) = Format(TxtValueofCommodity, "#######0.000") 'TxtValueofCommodity
oWS.Cells(mRow, C + 6).Font.Size = 7

mRow = mRow + 1


tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Signature & Stamp of Warehouse Manager"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).RowHeight = 50


tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Hologram No:- " & RsTxn_WHRDetail!HologramNo
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 6) = "Affix Hologram Here"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).Font.Size = 7
oWS.Cells(mRow, C + 6).RowHeight = 50


mRow = mRow + 1
tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).RowHeight = 8

mRow = mRow + 1


tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Insurance Details"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7
mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Valid From - Valid to"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).RowHeight = 20
oWS.Cells(mRow, C + 2) = "Policy & Cover Note No"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 3) = "Insurance Co."
oWS.Cells(mRow, C + 3).Font.Bold = True
oWS.Cells(mRow, C + 3).Font.Size = 7
oWS.Cells(mRow, C + 4) = "Policy Amount (in Rs.)"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).Font.Size = 7
tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = "Risk Covered"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).Font.Size = 7

mRow = mRow + 1
If ValidFrom.Value <> "1/1/1900" Or ValidTo <> "1/1/1900" Then
   oWS.Cells(mRow, C + 1) = Format(ValidFrom, "DD-MMM-YYYY") & " To " & Format(ValidTo, "DD-MMM-YYYY")   '"Valid From - Valid To"
   
Else
   oWS.Cells(mRow, C + 1) = ""
End If
oWS.Cells(mRow, C + 1).RowHeight = 20
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 2) = TxtPolicyNo
oWS.Cells(mRow, C + 2).WrapText = True
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 3) = TxtInsuranceCo
oWS.Cells(mRow, C + 3).WrapText = True
oWS.Cells(mRow, C + 3).Font.Size = 7
If TxtAmount <> "0" Then
   oWS.Cells(mRow, C + 4) = TxtAmount
Else
   oWS.Cells(mRow, C + 4) = ""
End If
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7
tmp = "e" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 5) = Trim(TxtRisks)
oWS.Cells(mRow, C + 5).WrapText = True
oWS.Cells(mRow, C + 5).Font.Size = 7

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "Pledge Details"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 7
mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Date"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "Name and Address Of Lender"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).WrapText = True
oWS.Cells(mRow, C + 2).Font.Size = 7

oWS.Cells(mRow, C + 4) = "Pledge Number"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7

oWS.Cells(mRow, C + 5) = "Pledge Released on"
oWS.Cells(mRow, C + 5).Font.Bold = True
oWS.Cells(mRow, C + 5).WrapText = True
oWS.Cells(mRow, C + 5).Font.Size = 7

oWS.Cells(mRow, C + 6) = "Sign of Lender/Warehouse Manager"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).WrapText = True
oWS.Cells(mRow, C + 6).Font.Size = 7

mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1
mRow = mRow + 1

oWS.Cells(mRow, C + 1) = "Withdrawal History"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7
oWS.Cells(mRow, C + 1).RowHeight = 24

tmp = "b" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "The good mentioned below are hereby released for delivery from the" & _
                         " Warehouse. The unreleased balance is subject to lien for" & _
                         " unpaid charges and borrowings from lending institutions."
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).WrapText = True
oWS.Cells(mRow, C + 2).Font.Size = 6

mRow = mRow + 1

tmp = "a" & mRow & ":" & "a" & (mRow + 1)
oWS.Range(tmp).Merge


oWS.Cells(mRow, C + 1) = "Date"
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 7

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 2) = "Quantity Released"
oWS.Cells(mRow, C + 2).Font.Bold = True
oWS.Cells(mRow, C + 2).WrapText = True
oWS.Cells(mRow, C + 2).Font.Size = 7

tmp = "d" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 4) = "Quantity Balance"
oWS.Cells(mRow, C + 4).Font.Bold = True
oWS.Cells(mRow, C + 4).WrapText = True
oWS.Cells(mRow, C + 4).Font.Size = 7

tmp = "f" & mRow & ":" & "f" & (mRow + 1)
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 6) = "Sign & Stamp of the Warehouse Manager"
oWS.Cells(mRow, C + 6).Font.Bold = True
oWS.Cells(mRow, C + 6).WrapText = True
oWS.Cells(mRow, C + 6).Font.Size = 7


mRow = mRow + 1

oWS.Cells(mRow, C + 2) = "No. of Bags"
oWS.Cells(mRow, C + 2).Font.Size = 7
oWS.Cells(mRow, C + 2).Font.Bold = True

oWS.Cells(mRow, C + 3) = "Weight in MT"
oWS.Cells(mRow, C + 3).Font.Size = 7
oWS.Cells(mRow, C + 3).Font.Bold = True

oWS.Cells(mRow, C + 4) = "No. Of Bags"
oWS.Cells(mRow, C + 4).Font.Size = 7
oWS.Cells(mRow, C + 4).Font.Bold = True

oWS.Cells(mRow, C + 5) = "Weight in  MT"
oWS.Cells(mRow, C + 5).Font.Size = 7
oWS.Cells(mRow, C + 5).Font.Bold = True

mRow = mRow + 1
mRow = mRow + 1
'mRow = mRow + 1
'mRow = mRow + 1
'mRow = mRow + 1

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1).RowHeight = 8

mRow = mRow + 1

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, C + 1) = "The Warehouse Receipt is valid for transfer by endorsement, and/or for " & _
                         " marking of plegde, subject to confirmation by Warehouse Manager "
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).WrapText = True
oWS.Cells(mRow, C + 1).Font.Size = 6
oWS.Cells(mRow, C + 1).RowHeight = 18

k = mRow

For Cell = 4 To k
    tmp = "a" & Cell & ":" & "f" & Cell
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    oWS.Range(tmp).Font.Name = "Zurich BT"
    'oWS.Range(tmp).Font.Size = 7
    oWS.Range(tmp).WrapText = True
Next Cell
    
    
If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If
oWB.Save
'oWB.Close
'oXL.Quit

mPrefix = GetPreFix(RsTxn_WHRDetail!CMPID, "P")
Call Log_Print_History(mPrefix, "WHR", TxtWHRNo, Gen_FLName)
    
    
MsgBox ("Excel file created!!! ")
oXL.Visible = True






End Sub


Private Sub CmdPrintPreview_Click()
'Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_WHR" & Trim(TxtWHRNoNew) & ".xls")
'Call Xls_Print_Rpt_New
'PntPrv = "Y"



'Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_WHR" & Trim(TxtWHRNoNew) & ".xls")
'Call Xls_Print_Rpt_New


tmp = " SELECT  TWD.Txn_PledgeCreation_ID,TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity,isnull(sum(PledgeQuantity),0) PledgeQuantity,TPC.Flag,isnull(TPC.NCMLFinance,0) 'NCMLFinance'"
tmp = tmp & " FROM  Txn_WHR_Detail TWD"
tmp = tmp & " left join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
tmp = tmp & " left Join Txn_PledgeCreation_GSL TPRG on TPRG.Txn_PledgeCreation_ID = TPC.Txn_PledgeCreation_ID And TPRG.SM_Prefix = TPC.SM_Prefix"
tmp = tmp & " Where TWD.WHRNo =" & RsTxn_WHRDetail!WHRNo & " And TPC.Txn_PledgeCreation_ID =" & IIf(IsNull(RsTxn_WHRDetail!Txn_PledgeCreation_ID), 0, RsTxn_WHRDetail!Txn_PledgeCreation_ID)
tmp = tmp & " Group By TWD.Txn_PledgeCreation_ID,TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity,TPC.Flag,TPC.NCMLFinance"
tmp = tmp & " Having IsNull(Sum(PledgeQuantity), 0) > 0"

Call Tmp2Table

If mSCMFlag = 0 Then
   If mNCMLFinance = 1 Then GoTo N
   If TmpRs2.RecordCount > 0 And (IsNull(RsTxn_WHRDetail!CARNo) = True Or RsTxn_WHRDetail!CARNo <= 0) Then
   
'   tmp = "Select ME.EmployeeNo +'~' +ME.EmployeeName as EmployeeName"
'   tmp = tmp & " From Mst_StateDRCMapping MSD"
'   tmp = tmp & " Inner join Mst_Employee ME on MSD.EmployeeID=ME.EmployeeID"
'   tmp = tmp & " where StateID = (select StateID from Mst_Location Where SM_Prefix='" & RsTxn_WHRDetail!SM_Prefix & "')"
   
      If IsNull(TmpRs2!Flag) = True Or TmpRs2!Flag = "" Then
         MsgBox "Pledge letter print pending. Can't preview WHR!!!"
         Exit Sub
      End If
   
      tmp = "Select MSD.EmployeeID From Mst_StateDRCMapping MSD"
      tmp = tmp & " Where StateID = (Select StateID from Mst_Location Where SM_Prefix='" & RsTxn_WHRDetail!SM_Prefix & "')"
   
   
      Call Tmp3Table
      mDRCName = ""
      For I = 1 To TmpRs3.RecordCount
       
          tmp = " Select ME.EmployeeName from Mst_Employee ME"
          tmp = tmp & " Where EmployeeID = " & TmpRs3!EmployeeID & " "
          
          Call TmpTableCMPro
       
          If mDRCName = "" Then
             mDRCName = TmpRsCMPro!EmployeeName 'TmpRs3!EmployeeName 'TmpRs3!EmployeeName
          Else
            mDRCName = mDRCName & "," & TmpRsCMPro!EmployeeName 'TmpRs3!EmployeeName
          End If
          TmpRs3.MoveNext
      Next
      If mDRCName <> "" Then
         MsgBox "No SR present for this WHR in CM-Pro, Contact to DRC " & mDRCName & ". you can't print!!"
      Else
         MsgBox "No SR present for this WHR in CM-Pro, Contact to DRC. Also map CM-DRC in Store-Man. You can't print!!"
      End If
      Exit Sub
   End If
N:
End If

PntPrv = "Y"
Frame5.Visible = True
''Call PDF_Print_Rpt
'Call Gen_Email



End Sub

Private Sub CmdUpdate_Click()
If TxtStorageCharges = "" Then
  MsgBox "Enter Storage Charges Rate!!!"
  TxtStorageCharges.SetFocus
  Exit Sub
End If
If TxtRate = "" Then
  MsgBox "Enter Commodity Price!!!"
  TxtRate.SetFocus
  Exit Sub
End If
If TxtWHRNoOfBags = "" Then
  MsgBox "Please provide WHR No of Bags!!!"
  TxtWHRNoOfBags.SetFocus
  Exit Sub
End If
If TxtWHRQty = "" Then
  MsgBox "Please provide WHR Quantity!!!"
  TxtWHRQty.SetFocus
  Exit Sub
End If

Dim a, b As Variant
a = Val(TxtEleWeight) - Val(TxtWHRQty)
b = Val(TxtNoofBags) - Val(TxtWHRNoOfBags)

If a = 0 Then
   If b <> 0 Then
      MsgBox "WHR should be issued for all available bags since it is issued for all available quantity!!!"
      TxtWHRQty.SetFocus
      Exit Sub
    End If
End If
If b = 0 Then
   If a <> 0 Then
      MsgBox "WHR should be issued for all available quantity since it is issued for all available no of bags!!!"
      TxtWHRQty.SetFocus
      Exit Sub
   End If
End If

'Don't check the weight per bag, if WHR is issued for all available bags and qty
If Val(TxtWHRQty) <> 0 And Val(TxtWHRQty) <> Val(TxtEleWeight) Then
   If Val(TxtNoofBags) <> Val(TxtWHRNoOfBags) Or Val(TxtEleWeight) <> Val(TxtWHRQty) Then
      tmp = Validate_ComBags(mCommodityID, Val(TxtWHRNoOfBags), Val(TxtWHRQty))
      If Len(tmp) > 1 Then
         MsgBox tmp
         TxtWHRQty.SetFocus
         Exit Sub
      End If
   End If
End If


If Val(TxtWHRNoOfBags.Text) = 0 And Val(TxtWHRQty) = 0 Then
    ChkSelection = 0
    MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = ""
Else
    MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "*"
    ChkSelection = 1
End If

MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = TxtCMPID
mCMPID = GetCMPID(TxtCMPID)
Dim m As Double
m = mCMPID
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 2) = TxtGodownNo
mGodownID = GetGodownID(TxtGodownNo, m)

MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 3) = TxtLicenseNo
mLicenseNo = TxtLicenseNo
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 4) = TxtLicenseDescription
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5) = Val(TxtNoofBags) - Val(TxtWHRNoOfBags)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6) = Val(TxtEleWeight) - Val(TxtWHRQty)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7) = Val(TxtWHRNoOfBags)
MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8) = Val(TxtWHRQty)

MSHFlexGrid2.TextMatrix(0, 7) = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
MSHFlexGrid2.TextMatrix(0, 8) = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))


If OptSRPerBag.Value = True Then
   TxtStorageAmount = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7)) * Val(TxtStorageCharges)
ElseIf OptSRPerMT.Value = True Then
   TxtStorageAmount = Format(Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8)) * Val(TxtStorageCharges), "#########0.00")
End If

TxtValueofCommodity = Format(Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8)) * Val(TxtRate), "#########0.00")
TxtAgreeDate = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 11)
TxtCMPID = ""
TxtGodownNo = ""
TxtLicenseNo = ""
TxtLicenseDescription = ""
TxtNoofBags = ""
TxtEleWeight = ""
TxtWHRNoOfBags = ""
TxtWHRQty = ""
Frame2.Enabled = False
CmdUpdate.Enabled = False
TxtStorageCharges.Locked = True
OptSRPerBag.Enabled = False
OptSRPerMT.Enabled = False
'SSTab1.Tab = 1
End Sub

Private Sub DeleteCmd_Click()

tmp = "Select * from Txn_PledgeCreation Where ExchangeTypeID = 1 And WHRNo = " & TxtWHRNo
Call TmpTable

If TmpRs.RecordCount > 0 Then
   MsgBox "Pledge is created for this WHR, you can't delete!!"
   Exit Sub
End If

If RsTxn_WHRDetail!Flag = "P" Then
   MsgBox "WHR is already printed.You can't delete this WHR!!! "
   Exit Sub
End If
If mLotWiseWHR = "0" Then
   If Val(MSHFlexGrid2.TextMatrix(2, 10)) > 0 Then
      MsgBox "You can't delete this WHR since there are withdrawals against this WHR!!! "
      Exit Sub
   End If
Else
   If MSHFlexGrid3.Rows > 3 Then
      MsgBox "You can't delete this WHR since there are withdrawals against this WHR!!! "
      Exit Sub
   End If
End If
Dim ans As Variant
'On Error GoTo msgError
If mLotWiseWHR = "0" Then
   mLicenseNo = GetLicenseNo(RsTxn_WHRDetail!LicenseID)
End If
ans = MsgBox("Do you really want to DELETE this record?", vbYesNo)
If ans = vbYes Then
   
   Call TableLog_WHR_Detail

   RsLog_WHR_Detail.AddNew
   RsLog_WHR_Detail!WHRNo = RsTxn_WHRDetail!WHRNo
   RsLog_WHR_Detail!SM_Prefix = RsTxn_WHRDetail!SM_Prefix
   RsLog_WHR_Detail!WHRDate = RsTxn_WHRDetail!WHRDate
   RsLog_WHR_Detail!ConditionID = RsTxn_WHRDetail!ConditionID
   RsLog_WHR_Detail!StorageFrom = RsTxn_WHRDetail!StorageFrom
   RsLog_WHR_Detail!StorageTo = RsTxn_WHRDetail!StorageTo
   RsLog_WHR_Detail!Grade = RsTxn_WHRDetail!Grade
   RsLog_WHR_Detail!Weigher = RsTxn_WHRDetail!Weigher
   RsLog_WHR_Detail!Assayer = RsTxn_WHRDetail!Assayer
   RsLog_WHR_Detail!GradeByAssayer = RsTxn_WHRDetail!GradeByAssayer
   RsLog_WHR_Detail!GradeValidTill = RsTxn_WHRDetail!GradeValidTill
   RsLog_WHR_Detail!StorageCharges = RsTxn_WHRDetail!StorageCharges
   RsLog_WHR_Detail!StorageAmount = RsTxn_WHRDetail!StorageAmount
   RsLog_WHR_Detail!ClientID = RsTxn_WHRDetail!ClientID
   RsLog_WHR_Detail!CommodityID = RsTxn_WHRDetail!CommodityID
   RsLog_WHR_Detail!CMPID = RsTxn_WHRDetail!CMPID
   RsLog_WHR_Detail!LicenseID = RsTxn_WHRDetail!LicenseID
   RsLog_WHR_Detail!WHRNoOfBags = RsTxn_WHRDetail!WHRNoOfBags
   RsLog_WHR_Detail!WHRQuantity = RsTxn_WHRDetail!WHRQuantity
   RsLog_WHR_Detail!PrintFlag = RsTxn_WHRDetail!PrintFlag
   RsLog_WHR_Detail!Remark = RsTxn_WHRDetail!Remark
   RsLog_WHR_Detail!Flag = RsTxn_WHRDetail!Flag
   RsLog_WHR_Detail!Txn_PledgeCreation_ID = RsTxn_WHRDetail!Txn_PledgeCreation_ID
   RsLog_WHR_Detail!oldLicenseID = RsTxn_WHRDetail!oldLicenseID
   RsLog_WHR_Detail!GodownID = RsTxn_WHRDetail!GodownID
   RsLog_WHR_Detail!DeliveredQuantity = RsTxn_WHRDetail!DeliveredQuantity
   RsLog_WHR_Detail!DeliveredNoOfBags = RsTxn_WHRDetail!DeliveredNoOfBags
   RsLog_WHR_Detail!Rate = RsTxn_WHRDetail!Rate
   RsLog_WHR_Detail!G_UID = RsTxn_WHRDetail!G_UID
   RsLog_WHR_Detail!HologramNo = RsTxn_WHRDetail!HologramNo
   RsLog_WHR_Detail!PolicyNo = RsTxn_WHRDetail!PolicyNo
   RsLog_WHR_Detail!CoverNoteNo = RsTxn_WHRDetail!CoverNoteNo
   RsLog_WHR_Detail!SaidToContain = RsTxn_WHRDetail!SaidToContain
   RsLog_WHR_Detail!EstimatedWeight = RsTxn_WHRDetail!EstimatedWeight
   RsLog_WHR_Detail!CreatedBy = RsTxn_WHRDetail!CreatedBy
   RsLog_WHR_Detail!CreatedDate = RsTxn_WHRDetail!CreatedDate
   RsLog_WHR_Detail!UpdatedDate = RsTxn_WHRDetail!UpdatedDate
   RsLog_WHR_Detail!UpdatedBy = RsTxn_WHRDetail!UpdatedBy
   RsLog_WHR_Detail!SRAsPerBag = RsTxn_WHRDetail!SRAsPerBag
   RsLog_WHR_Detail!DeletedBy = Gen_UserLoginID
   RsLog_WHR_Detail.Update
   RsLog_WHR_Detail.Close
   
   If mLotWiseWHR = 0 Then
      Call PostDeleteGSLData(RsTxn_WHRDetail!CMPID, RsTxn_WHRDetail!GodownID, RsTxn_WHRDetail!LicenseID, RsTxn_WHRDetail!WHRNoOfBags, RsTxn_WHRDetail!WHRQuantity)
   Else
      tmp = "Update Mst_GSL Set WHRBags = 0,WHRWeight = 0 ,WHRNO = Null Where WHRNo = " & TxtWHRNo & ""
      Call TmpTable
   End If
    RsTxn_WHRDetail.Delete
    RsTxn_WHRDetail.Update
    
   tmp = "Update Mst_GSL Set WHRBags = 0,WHRWeight = 0 ,WHRNO = Null Where WHRNo = " & TxtWHRNo & ""
   Call TmpTableAzureStoreMan
    
   tmp = "Delete from Txn_WHR_Detail where WHRNo = " & TxtWHRNo & ""
   Call TmpTableAzureStoreMan
    
    If RsTxn_WHRDetail.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
    End If
    
    RsTxn_WHRDetail.MoveNext
    If RsTxn_WHRDetail.EOF() Then
      RsTxn_WHRDetail.MoveLast
    End If
    Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "

End Sub
Function PostDeleteGSLData(CMPID_ As Double, GodownID_ As Double, LicenseID_ As Double, WHRNoBags_ As Double, WHRQty_ As Double)
Dim TBags, TQty, PBags, PQty As Variant

Call TableMst_GSL(" Where WHRNO is null ")
    '"((Mst_GSL.CDTFDate) <= '" & Format(WHRDate, Gen_DatFormat) & "')) " & _

    tmp = "SELECT ((Mst_GSL.DepositeBags)) AS BalanceBags, ((Mst_GSL.DepositeWeight)) AS BalanceWeight, " & _
          " ((Mst_GSL.WHRBags)) AS WHRBags, ((Mst_GSL.WHRWeight)) AS WHRWeight, " & _
    "Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName, " & _
    "Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientID,  Mst_License.LicenseID,Mst_Godown.GodownNo, " & _
    "Mst_Godown.GodownID,Mst_Godown.AgreementEndDate " & _
    "FROM Mst_CMP INNER JOIN (Mst_License INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = " & _
    "Mst_GSL.GodownID) ON Mst_License.LicenseID = Mst_Godown.LicenseID) ON Mst_CMP.CMPID = Mst_GSL.CMPID " & _
    "WHERE ( ((Mst_GSL.CMPID)=" & CMPID_ & ") AND ((Mst_GSL.GodownID)=" & GodownID_ & ") AND " & _
    "((Mst_GSL.ClientID)=" & mClientIDRow & ") AND ((Mst_GSL.ExchangeTypeID)=" & mExchangeTypeID & ") AND " & _
    "((Mst_GSL.CommodityID)=" & mCommodityID & ") AND " & _
    "((Mst_GSL.CDTFDate) Between '" & Format(StorageFrom, Gen_DatFormat) & "' And '" & Format(WHRDate, Gen_DatFormat) & "')) " & _
    " And Mst_GSL.Status <> 'Z' And Mst_GSL.MadeupFLag = '0' And WHRNo is null order by Mst_GSL.CDTFDate "
    
    Call Tmp1Table

TBags = WHRNoBags_
TQty = WHRQty_
For I = 1 To TmpRs1.RecordCount
    If TmpRs1!WHRBags > 0 Then
       If TBags >= TmpRs1!WHRBags Then
          PBags = TmpRs1!WHRBags
          TBags = TBags - PBags
       Else
          PBags = TBags
          TBags = 0
       End If
       TmpRs1!WHRBags = IIf(IsNull(TmpRs1!WHRBags), 0, TmpRs1!WHRBags) - PBags
       'TmpRs1!BalanceBags = TmpRs1!BalanceBags - PBags
       TmpRs1.Update
       If TBags = 0 Then Exit For
    End If
    TmpRs1.MoveNext
Next
If TmpRs1.RecordCount > 0 Then
    TmpRs1.MoveFirst
    TQty = WHRQty_
    For I = 1 To TmpRs1.RecordCount
        If TmpRs1!WHRWeight > 0 Then
           If TQty >= TmpRs1!WHRWeight Then
              PQty = TmpRs1!WHRWeight
              TQty = TQty - PQty
           Else
              PQty = TQty
              TQty = 0
           End If
           TmpRs1!WHRWeight = IIf(IsNull(TmpRs1!WHRWeight), 0, TmpRs1!WHRWeight) - PQty
           'TmpRs1!BalanceWeight = TmpRs1!BalanceWeight - PQty
           TmpRs1.Update
           If TQty = 0 Then Exit For
        End If
        TmpRs1.MoveNext
    Next
    TmpRs1.Close
End If
End Function

Private Sub EditCmd_Click()
  Edit_Flg = "E"
'  CmbCommodityCondition.Locked = False
'  TxtGrade.Locked = False
'  TxtClientID.Visible = True
'  CmbClientID.Visible = False
'  CmbClientID.Locked = True
'  TxtWHRCommodityId.Visible = True
'  CmbWHRCommodity.Visible = False
'  CmbWHRCommodity.Locked = True
'  TxtWeigher.Locked = False
'  TxtAssayer.Locked = False
'  TxtAssayerCerti.Locked = False
'  GradeValidUpto.Enabled = False
'  TxtStorageCharges.Locked = False
'  TxtWHRNoOfBags.Locked = True
'  TxtWHRQty.Locked = True
'  TxtInsuranceCo.Visible = False
'  TxtPolicyNo.Visible = False
'  CmbInsuranceCo.Visible = True
'  CmbPolicyNo.Visible = True
'  CmbInsuranceCo.Text = TxtInsuranceCo
'  CmbPolicyNo.Text = TxtPolicyNo
'  ChkSaidToContain.Enabled = True
'  ChkEstimatedWgt.Enabled = True
'  ChkSelection = 0

CmdPledge.Enabled = False

tmp = " SELECT  TWD.Txn_PledgeCreation_ID,TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity,isnull(sum(PledgeQuantity),0) PledgeQuantity"
tmp = tmp & " FROM  Txn_WHR_Detail TWD"
tmp = tmp & " left join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
tmp = tmp & " left Join Txn_PledgeCreation_GSL TPRG on TPRG.Txn_PledgeCreation_ID = TPC.Txn_PledgeCreation_ID And TPRG.SM_Prefix = TPC.SM_Prefix"
tmp = tmp & " Where TWD.WHRNo =  " & TxtWHRNo & " "
tmp = tmp & " Group By TWD.Txn_PledgeCreation_ID,TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity"
tmp = tmp & " Having IsNull(Sum(PledgeQuantity), 0) > 0"

Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
   MsgBox "Pledge is created for this WHR, you can't Edit!!!"
   Exit Sub
End If

SSTab1.Enabled = False
CmbCMPID.Text = TxtCMPID
CmbWHRCommodity.Text = TxtWHRCommodityId
TxtStorageCharges.Locked = False
TxtRate.Locked = False
CmbCommodityCondition.Locked = False
TxtGrade.Locked = False
CmbInsuranceCo.Visible = True
CmbPolicyNo.Visible = True
OptSRPerBag.Enabled = False
OptSRPerMT.Enabled = False

CmbInsuranceCo.Text = TxtInsuranceCo
CmbPolicyNo.Text = TxtPolicyNo

Call GButtonLock(Me, False)

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied!!!"
   Exit Sub
End If

Call Grid_Head
Call Grid_Head1
CmbPreFix.Text = ""
CmbSCmp.Text = ""
CmbSClientName.Text = ""
CmbSCommodity.Text = ""
TxtSWHRNo = ""
DTPWHRFromDate.Value = Date - 365
DTPWHRToDate.Value = Date

If Gen_WcCMP <> "" Then
   Call TableTxn_WHRDetail(" Where " & Gen_WcCMP)
Else
   Call TableTxn_WHRDetail("")
End If



If RsTxn_WHRDetail.RecordCount = 0 Then
  If UserType = "Auditor" Then
     Call GButtonLock(Me, False)
     CmdPrint.Enabled = True
     SaveCmd.Enabled = False
     SearchCmd.Enabled = False
     Exit Sub
  End If
  Call AddCmd_Click
  Exit Sub
End If

Call Indata
End Sub
Private Sub Indata()
SSTab1.Enabled = True
CmdPledge.Enabled = False
PntPrv = "N"
Edit_Flg = ""
ChkSelectAll.Enabled = False
WHRDate.Enabled = False
CmbCommodityCondition.Locked = True
StorageFrom.Enabled = False
StorageTo.Enabled = False
TxtGrade.Locked = True
TxtClientID.Visible = True
CmbClientID.Visible = False
CmbClientID.Locked = True
TxtWHRCommodityId.Visible = True
CmbWHRCommodity.Visible = False
CmbWHRCommodity.Locked = True
TxtWeigher.Locked = True
TxtAssayer.Locked = True
TxtAssayerCerti.Locked = True
GradeValidUpto.Enabled = False
TxtStorageCharges.Locked = True
TxtRate.Locked = True
TxtWHRNoOfBags.Locked = True
TxtWHRQty.Locked = True
TxtInsuranceCo.Visible = True
TxtPolicyNo.Visible = True
CmbInsuranceCo.Visible = False
CmbPolicyNo.Visible = False
CmdInsuranceCo.Visible = False
CmdInsurance.Visible = False
ChkSaidToContain.Enabled = False
ChkEstimatedWgt.Enabled = False
OptSRPerBag.Enabled = False
OptSRPerMT.Enabled = False
CmdPrintPreview.Enabled = True

TxtCMPName.Visible = True
CmbCMPID.Visible = False
Check1.Enabled = False

TxtGodownID.Visible = True
CmbGodownID.Visible = False


CmdUpdate.Enabled = False
Call GButtonLock(Me, True)
CmdPrint.Enabled = True
mExchangeTypeID = GetExchangeID("NON NCDEX")
mLotWiseWHR = IIf(IsNull(RsTxn_WHRDetail!LotWiseWHR), 0, RsTxn_WHRDetail!LotWiseWHR)
MSHFlexGrid2.Visible = False
MSHFlexGrid4.Visible = False
If mLotWiseWHR = 0 Then
   MSHFlexGrid2.Visible = True
   Frame2.Visible = True
   Call Grid_Head
   ChkSelectAll.Value = 0
   ChkSelectAll.Visible = False
Else
   MSHFlexGrid4.Visible = True
   Frame2.Visible = False
   Call Grid_HeadNew
   ChkSelectAll.Visible = True
   ChkSelectAll.Value = 0
End If

TxtWHRNo = IIf(IsNull(RsTxn_WHRDetail!WHRNo), "", RsTxn_WHRDetail!WHRNo)
TxtWHRNoNew = IIf(IsNull(RsTxn_WHRDetail!CMPID), 0, RsTxn_WHRDetail!CMPID) & "-" & Val(TxtWHRNo) - (IIf(IsNull(RsTxn_WHRDetail!CMPID), 0, RsTxn_WHRDetail!CMPID) * 1000000)




WHRDate.Value = IIf(IsNull(RsTxn_WHRDetail!WHRDate), "", RsTxn_WHRDetail!WHRDate)
mCommodityCondition = IIf(IsNull(RsTxn_WHRDetail!ConditionID), "", RsTxn_WHRDetail!ConditionID)
If mCommodityCondition = 0 Then
   CmbCommodityCondition.Text = "Average"
ElseIf mCommodityCondition = 1 Then
   CmbCommodityCondition.Text = "Fair"
ElseIf mCommodityCondition = 2 Then
   CmbCommodityCondition.Text = "Good"
End If
StorageFrom.Value = IIf(IsNull(RsTxn_WHRDetail!StorageFrom), "", RsTxn_WHRDetail!StorageFrom)
StorageTo.Value = IIf(IsNull(RsTxn_WHRDetail!StorageTo), "", RsTxn_WHRDetail!StorageTo)
TxtGrade = IIf(IsNull(RsTxn_WHRDetail!Grade), "", RsTxn_WHRDetail!Grade)
mClientIDRow = IIf(IsNull(RsTxn_WHRDetail!ClientID), "", RsTxn_WHRDetail!ClientID)
mCMPID = IIf(IsNull(RsTxn_WHRDetail!CMPID), "", RsTxn_WHRDetail!CMPID)
mGodownID = IIf(IsNull(RsTxn_WHRDetail!GodownID), "", RsTxn_WHRDetail!GodownID)
TxtCMPName = GetCMPName(IIf(IsNull(RsTxn_WHRDetail!CMPID), "", RsTxn_WHRDetail!CMPID))
TxtGodownID = GetGodownName(CDbl(mGodownID), "N")
TxtAgreementEndDate = Format(GetGodownName(CDbl(mGodownID), "ED"), "dd-mmm-yyyy")
mLocationID = GetLocationIDFromCMP(mCMPID)
mSCMFlag = IIf(IsNull(RsTxn_WHRDetail!SCMFlag), 0, RsTxn_WHRDetail!SCMFlag)
ans = GetClientNameByRow(mClientIDRow)

If ans(0) = "" Then
   MsgBox "Not found!!!"
   Exit Sub
End If
TxtClientID.Text = ans(1)
mCommodityID = IIf(IsNull(RsTxn_WHRDetail!CommodityID), "", RsTxn_WHRDetail!CommodityID)
TxtWHRCommodityId.Text = GetCommodityName(mCommodityID)
TxtWeigher = IIf(IsNull(RsTxn_WHRDetail!Weigher), "", RsTxn_WHRDetail!Weigher)
TxtAssayer = IIf(IsNull(RsTxn_WHRDetail!Assayer), "", RsTxn_WHRDetail!Assayer)
TxtAssayerCerti = IIf(IsNull(RsTxn_WHRDetail!GradeByAssayer), "", RsTxn_WHRDetail!GradeByAssayer)
GradeValidUpto.Value = IIf(IsNull(RsTxn_WHRDetail!GradeValidTill), "", RsTxn_WHRDetail!GradeValidTill)
TxtStorageCharges = IIf(IsNull(RsTxn_WHRDetail!StorageCharges), "", RsTxn_WHRDetail!StorageCharges)
TxtStorageAmount = IIf(IsNull(RsTxn_WHRDetail!StorageAmount), "", RsTxn_WHRDetail!StorageAmount)
TxtRate = IIf(IsNull(RsTxn_WHRDetail!Rate), "", RsTxn_WHRDetail!Rate)
TxtPolicyNo = RsTxn_WHRDetail!PolicyNo & " ~ " & RsTxn_WHRDetail!CoverNoteNo
PolicyNo_ = IIf(IsNull(RsTxn_WHRDetail!PolicyNo), "", RsTxn_WHRDetail!PolicyNo)
CNo = IIf(IsNull(RsTxn_WHRDetail!CoverNoteNo), "", RsTxn_WHRDetail!CoverNoteNo)
ChkSaidToContain.Value = IIf(IsNull(RsTxn_WHRDetail!SaidToContain), 0, RsTxn_WHRDetail!SaidToContain)
ChkEstimatedWgt.Value = IIf(IsNull(RsTxn_WHRDetail!EstimatedWeight), 0, RsTxn_WHRDetail!EstimatedWeight)
Check1.Value = IIf(IsNull(RsTxn_WHRDetail!ConvertedClient), 0, RsTxn_WHRDetail!ConvertedClient)
If RsTxn_WHRDetail!SRAsPerBag = 1 Then
   OptSRPerBag.Value = True
Else
   OptSRPerMT.Value = True
End If
TxtCreated = IIf(IsNull(RsTxn_WHRDetail!CreatedBy), "", RsTxn_WHRDetail!CreatedBy) & " :- " & IIf(IsNull(RsTxn_WHRDetail!CreatedDate), "", RsTxn_WHRDetail!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_WHRDetail!UpdatedBy), "", RsTxn_WHRDetail!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_WHRDetail!UpdatedDate), "", RsTxn_WHRDetail!UpdatedDate)
TxtCARNo.Text = ""
Call TableMst_PolicyDetail("where PolicyNo = '" & RsTxn_WHRDetail!PolicyNo & "' And CoverNoteNo = '" & RsTxn_WHRDetail!CoverNoteNo & "'")

If RsMst_PolicyDetail.RecordCount > 0 Then
   TxtInsuranceCo = GetInsuranceCoName(RsMst_PolicyDetail!InsuranceCoID)
   ValidFrom.Value = RsMst_PolicyDetail!ValidFrom
   ValidTo.Value = RsMst_PolicyDetail!ValidTo
   TxtAmount = RsMst_PolicyDetail!Amount
   TxtRisks = RsMst_PolicyDetail!Risks
   CmdInsuranceCo.Visible = False
   CmdInsurance.Visible = False
   RsMst_PolicyDetail.Close
End If
  
If mLotWiseWHR = "0" Then
   Call GetGridInData
Else
   mLicenseID = IIf(IsNull(RsTxn_WHRDetail!LicenseID), "", RsTxn_WHRDetail!LicenseID)
   Call TableMst_License(" Where LicenseID = " & mLicenseID)
   mLicenseNo = RsMst_License!LicenseNo
   mLicenseExpiryDate = RsMst_License!LicenseExpiryDate
   Call GetGridInDataLotWise
End If
If mLotWiseWHR = "0" Then
   If OptSRPerBag.Value = True Then
      TxtStorageAmount = Val(MSHFlexGrid2.TextMatrix(2, 7)) * Val(TxtStorageCharges)
   Else
      TxtStorageAmount = Format(Val(MSHFlexGrid2.TextMatrix(2, 8)) * Val(TxtStorageCharges), "#########0.00")
   End If

   TxtValueofCommodity = Format(Val(MSHFlexGrid2.TextMatrix(2, 8)) * Val(TxtRate), "#########0.00")
Else
   If OptSRPerBag.Value = True Then
      TxtStorageAmount = Val(MSHFlexGrid4.TextMatrix(0, 5)) * Val(TxtStorageCharges)
   Else
      TxtStorageAmount = Format(Val(MSHFlexGrid4.TextMatrix(0, 6)) * Val(TxtStorageCharges), "#########0.00")
   End If

   TxtValueofCommodity = Format(Val(MSHFlexGrid4.TextMatrix(0, 6)) * Val(TxtRate), "#########0.00")
End If

'EditCmd.Enabled = False
Frame3.Visible = False
Frame5.Visible = False
Call Indata_CITF
Call InDataPledgeDetail
TxtCARNo.Text = ""
If mNCMLFinance = 0 Then
   If Val(TxtPledgeNo) > 0 Then
'      If IsNull(RsTxn_WHRDetail!CARNo) = True Or RsTxn_WHRDetail!CARNo = 0 Then
         tmp = "Select WHRNo,CARNO from Txn_WHR_Detail Where WHRNO = " & TxtWHRNo & " " 'And CARNo is not null"
         Call TmpTableAzureStoreMan
         If TmpRsAzureStoreMan.RecordCount > 0 Then
            If IsNull(TmpRsAzureStoreMan!CARNo) = False Then
               tmp = "Update Txn_WHR_Detail Set CARNO = " & TmpRsAzureStoreMan!CARNo & " Where WHRNo = " & TxtWHRNo & " "
               TxtCARNo.Text = IIf(IsNull(TmpRsAzureStoreMan!CARNo), "", TmpRsAzureStoreMan!CARNo)
            Else
               tmp = "Update Txn_WHR_Detail Set CARNO = Null Where WHRNo = " & TxtWHRNo & " "
            End If
            Call TmpTable
            
         End If
'      End If
   End If
End If
Call IndataCARDetails
CmdPrint.Caption = "&Print"
If mNCMLFinance = 1 Then
   CmdPrint.Caption = "Send Mail to NFin"
End If



SSTab1.Tab = 0
  'Call GridInData
End Sub
Private Sub GetGridInData()
mLicenseID = IIf(IsNull(RsTxn_WHRDetail!LicenseID), "", RsTxn_WHRDetail!LicenseID)
k = 1
k = k + 1
Call TableMst_License(" Where LicenseID = " & mLicenseID)
MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
MSHFlexGrid2.TextMatrix(k, 1) = TxtCMPName  'TxtCMPID ' GetCMPName(RsTxn_WHRDetail!CMPID)
MSHFlexGrid2.TextMatrix(k, 2) = TxtGodownID  'GetGodownName(RsTxn_WHRDetail!GodownID, "N")
MSHFlexGrid2.TextMatrix(k, 3) = RsMst_License!LicenseNo
MSHFlexGrid2.TextMatrix(k, 4) = RsMst_License!LicenseDesc
MSHFlexGrid2.TextMatrix(k, 5) = 0
MSHFlexGrid2.TextMatrix(k, 6) = 0
MSHFlexGrid2.TextMatrix(k, 7) = RsTxn_WHRDetail!WHRNoOfBags
MSHFlexGrid2.TextMatrix(k, 8) = RsTxn_WHRDetail!WHRQuantity
MSHFlexGrid2.TextMatrix(k, 10) = IIf(IsNull(RsTxn_WHRDetail!DeliveredQuantity), 0, RsTxn_WHRDetail!DeliveredQuantity)
'MSHFlexGrid2.TextMatrix(k, 11) = RsTxn_WHRDetail!DeliveredQuantity
MSHFlexGrid2.TextMatrix(k, 12) = RsTxn_WHRDetail!SM_Prefix
'mGodownID = RsTxn_WHRDetail!GodownID

End Sub
Public Sub Grid_Head()

MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 13
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.TextMatrix(1, 0) = "   "
MSHFlexGrid2.TextMatrix(1, 1) = "CMP Name "
MSHFlexGrid2.TextMatrix(1, 2) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 3) = "License No"
MSHFlexGrid2.TextMatrix(1, 4) = "License Description "
MSHFlexGrid2.TextMatrix(1, 5) = "Available No Of Bags"
MSHFlexGrid2.TextMatrix(1, 6) = "Available Electronic Weight"
MSHFlexGrid2.TextMatrix(1, 7) = "WHR No Of Bags"
MSHFlexGrid2.TextMatrix(1, 8) = "WHR Quantity"
MSHFlexGrid2.TextMatrix(1, 9) = "WHR LicenseDetailID"
MSHFlexGrid2.TextMatrix(1, 10) = "Delivered Qty"
MSHFlexGrid2.TextMatrix(1, 11) = "Agreement End Date"
MSHFlexGrid2.TextMatrix(1, 12) = "Prefix"

MSHFlexGrid2.ColWidth(0) = 400
MSHFlexGrid2.ColWidth(1) = 1400
MSHFlexGrid2.ColWidth(2) = 1400
MSHFlexGrid2.ColWidth(3) = 1200
MSHFlexGrid2.ColWidth(4) = 0
MSHFlexGrid2.ColWidth(5) = 1600
MSHFlexGrid2.ColWidth(6) = 1800
MSHFlexGrid2.ColWidth(7) = 1800
MSHFlexGrid2.ColWidth(8) = 1600
MSHFlexGrid2.ColWidth(9) = 0
MSHFlexGrid2.ColWidth(12) = 0

MSHFlexGrid2.RowHeight(1) = 500
End Sub
Public Sub Grid_Head1()
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 32 '31 '26 '25 '24
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.FixedRows = 1

MSHFlexGrid1.TextMatrix(0, 0) = "WHR No"
MSHFlexGrid1.TextMatrix(0, 1) = "Location ID"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Condition Of Commodity"
MSHFlexGrid1.TextMatrix(0, 4) = "WHR Date "
MSHFlexGrid1.TextMatrix(0, 5) = "Deposit Date"
MSHFlexGrid1.TextMatrix(0, 6) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 7) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 8) = "Private Marks on Packaging"
MSHFlexGrid1.TextMatrix(0, 9) = "GUID"
MSHFlexGrid1.TextMatrix(0, 10) = "Bags"
MSHFlexGrid1.TextMatrix(0, 11) = "Quantity"
MSHFlexGrid1.TextMatrix(0, 12) = "Flag"
MSHFlexGrid1.TextMatrix(0, 13) = "Delivered Bags (excluding Spillage and Madeup)"
MSHFlexGrid1.TextMatrix(0, 14) = "Delivered Quantity(including Spillage and Madeup)"
MSHFlexGrid1.TextMatrix(0, 15) = "CARNo"
MSHFlexGrid1.TextMatrix(0, 16) = "Converted Client"
MSHFlexGrid1.TextMatrix(0, 17) = "SCM WHR"
MSHFlexGrid1.TextMatrix(0, 18) = "Rate"
MSHFlexGrid1.TextMatrix(0, 19) = "Value of Commodity"
MSHFlexGrid1.TextMatrix(0, 20) = "Pledge ID"
MSHFlexGrid1.TextMatrix(0, 21) = "Pledge Date"
MSHFlexGrid1.TextMatrix(0, 22) = "Bank Name"
MSHFlexGrid1.TextMatrix(0, 23) = "Branch Name"
MSHFlexGrid1.TextMatrix(0, 24) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 25) = "Lot wise WHR"
MSHFlexGrid1.TextMatrix(0, 26) = "PolicyNo~CoverNoteNo"
MSHFlexGrid1.TextMatrix(0, 27) = "Risk Covered in Insurance"
MSHFlexGrid1.TextMatrix(0, 28) = "Amount in Insurance Policy"
MSHFlexGrid1.TextMatrix(0, 29) = "Insurance Valid From"
MSHFlexGrid1.TextMatrix(0, 30) = "Insurance Valid To"
MSHFlexGrid1.TextMatrix(0, 31) = "Insurance Company"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 1000
MSHFlexGrid1.ColWidth(2) = 1800
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 1200
MSHFlexGrid1.ColWidth(6) = 1200
MSHFlexGrid1.ColWidth(7) = 1200
MSHFlexGrid1.ColWidth(8) = 1400
MSHFlexGrid1.ColWidth(9) = 0
MSHFlexGrid1.ColWidth(10) = 800
MSHFlexGrid1.ColWidth(11) = 600
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 800
'MSHFlexGrid1.ColWidth(13) = 800

MSHFlexGrid1.RowHeight(0) = 700
End Sub





Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_WHRDetail.MoveFirst
RsTxn_WHRDetail.Find "G_Uid = '" & MSHFlexGrid1.TextMatrix(I, 9) & "'"
If Not RsTxn_WHRDetail.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()

If SaveCmd.Enabled = False Then Exit Sub

If MSHFlexGrid2.Rows > 3 Then
    If MSHFlexGrid2.RowSel <> "0" Then
       If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = "" Then
          TxtCMPID = ""
          TxtGodownNo = ""
          TxtLicenseNo = ""
          TxtLicenseDescription = ""
          TxtNoofBags = ""
          TxtEleWeight = ""
          TxtWHRNoOfBags = ""
          TxtWHRQty = ""
          CmdUpdate.Enabled = False
          Frame2.Enabled = False
          Exit Sub
       End If
       Frame2.Enabled = True
       If ChkSelection = 1 Then
          If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) <> "*" Then
             MsgBox "You can't select more than one godown. To discard the existing selection, click on commodity name!!!"
             Exit Sub
          End If
       End If
      TxtWHRNoOfBags.Locked = False
      TxtWHRQty.Locked = False
      TxtCMPID = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1)
      TxtGodownNo = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 2)
      
      TxtLicenseNo = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 3)
      TxtLicenseDescription = MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 4)
      TxtNoofBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
      TxtEleWeight = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))
      TxtWHRNoOfBags = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 7))
      TxtWHRQty = Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 6)) + Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8))
      'TxtWHRNoOfBags.SetFocus
      CmdUpdate.Enabled = True
      TxtStorageCharges.Locked = False
      OptSRPerBag.Enabled = True
      OptSRPerMT.Enabled = True
      
      If mSCMFlag = 1 Then
         TxtWHRNoOfBags.Locked = True
         TxtWHRQty.Locked = True
      Else
         TxtWHRNoOfBags.Locked = False
         TxtWHRQty.Locked = False
      End If
      
      CmdUpdate.SetFocus
    End If
End If
End Sub

Private Sub OptAll_Click()
    DTPWHRFromDate.Enabled = False
   DTPWHRToDate.Enabled = False
End Sub


Private Sub MSHFlexGrid4_Click()

If SaveCmd.Enabled = False Then Exit Sub
If mSCMFlag = 1 Then Exit Sub
With MSHFlexGrid4
    If .TextMatrix(.row, 2) = "" Then Exit Sub
    If .Col = 0 Then
        If .TextMatrix(.row, 0) = strUnChecked Then
           .TextMatrix(.row, 0) = strChecked
           .TextMatrix(.row, 5) = Val(.TextMatrix(.row, 3))
           .TextMatrix(.row, 6) = Val(.TextMatrix(.row, 4))
           .TextMatrix(0, 5) = Val(.TextMatrix(0, 5)) + Val(.TextMatrix(.row, 5))
           .TextMatrix(0, 6) = Val(.TextMatrix(0, 6)) + Val(.TextMatrix(.row, 6))
        Else
           .TextMatrix(.row, 0) = strUnChecked
           .TextMatrix(0, 5) = Val(.TextMatrix(0, 5)) - Val(.TextMatrix(.row, 5))
           .TextMatrix(0, 6) = Val(.TextMatrix(0, 6)) - Val(.TextMatrix(.row, 6))
           .TextMatrix(.row, 5) = 0
           .TextMatrix(.row, 6) = 0
        End If
    End If
End With

End Sub

Private Sub SaveCmd_Click()

If CmbCommodityCondition.Text = "" Then
   MsgBox "Commodity Condition cannot be blank!!!"
   CmbCommodityCondition.SetFocus
   Exit Sub
End If

If CmbCMPID.Text = "" And Edit_Flg = "A" Then
   MsgBox "CMP Name cannot be blank!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If

If CmbGodownID.Text = "" And Edit_Flg = "A" Then
   MsgBox "Godown cannot be blank!!!"
   CmbGodownID.SetFocus
   Exit Sub
End If

If WHRDate.Value > Date Then
   MsgBox "WHR Date cannot be Future Date!!!"
   WHRDate.SetFocus
   Exit Sub
End If
If WHRDate.Value < StorageFrom.Value Then
   MsgBox "WHR Date cannot be less than Storage From Date!!!"
   WHRDate.SetFocus
   Exit Sub
End If


'Commeneted on 27th June for removal of date condition mail by Sankar sir
'If LCase(CmbCMPID.Text) <> "unjha" Or LCase(CmbCMPID.Text) <> "rajkot" Then
'   If Edit_Flg = "A" Then
'      If WHRDate.Value > CDate(TxtBODDate) Then
'         MsgBox "WHR Date can't be greater than 'BOD Date + 1' of the selected CMP i.e. " & TxtBODDate & "!!!"
'         WHRDate.SetFocus
'         Exit Sub
'      End If
'   End If
'End If


If Day(StorageTo.Value) <> 20 Then
   MsgBox "Storage To Date must be 20th of calender month !!!"
   StorageTo.SetFocus
   Exit Sub
End If

If TxtAgreementEndDate <> "" Then
   If StorageTo.Value > CDate(TxtAgreementEndDate) Then
      MsgBox "'Storage To Date' cannot be greater than 'Agreement End Date' !!!"
      StorageTo.SetFocus
      Exit Sub
   End If
End If

If WHRDate.Value > StorageTo.Value Then
   MsgBox "Storage To Date cannot be less than WHR Date!!!"
   WHRDate.SetFocus
   Exit Sub
End If
If StorageFrom.Value > StorageTo.Value Then
   MsgBox "Storage To Date must be greater than Storage From Date!!!"
   StorageTo.SetFocus
   Exit Sub
End If

If CmbInsuranceCo.Text = "" Then
   MsgBox "Insurance Company Name can't be blank!!!"
   CmbInsuranceCo.SetFocus
   Exit Sub
End If
      
If CmbPolicyNo.Text = "" Then
   MsgBox "'Insurance policy' can't be blank!!!"
   CmbPolicyNo.SetFocus
   Exit Sub
End If

If TxtGrade = "" Then
   MsgBox "'Private Marks on Packaging' can't be blank!!!"
   TxtGrade.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
    If CmbClientID.Text = "" Then
       MsgBox "'Client Name' can't be blank!!!"
       CmbClientID.SetFocus
       Exit Sub
    End If
    
    If CmbWHRCommodity.Text = "" Then
       MsgBox "'Commodity' can't be blank!!!"
       CmbWHRCommodity.SetFocus
       Exit Sub
    End If
End If
If TxtStorageCharges.Text = "" Then
   MsgBox "Storage Charges can't be blank!!!"
   TxtStorageCharges.SetFocus
   Exit Sub
End If
If TxtRate.Text = "" Then
   MsgBox "Commodity price can't be blank!!!"
   TxtRate.SetFocus
   Exit Sub
End If
If GradeValidUpto.Value < StorageTo.Value Then
   MsgBox "'Storage To' date can't be greater than 'Grade Valid till' Date!!!"
   GradeValidUpto.SetFocus
   Exit Sub
End If

If mLotWiseWHR = "0" Then
   If MSHFlexGrid2.Rows < 4 Then
      MsgBox "No deposits available for generating WHR!!!"
      Exit Sub
   End If
Else
   If MSHFlexGrid4.Rows < 4 Then
      MsgBox "No deposits available for generating WHR!!!"
      Exit Sub
   End If
   If Val(MSHFlexGrid4.TextMatrix(0, 5)) <= 0 Then
      MsgBox "No Lot Selected for generating WHR!!!"
      Exit Sub
   End If


End If
If Edit_Flg = "A" Then
   If mLotWiseWHR = "0" Then
      If ChkSelection = 0 Then
         MsgBox "Select atleast one godown from 'GSL Detail' tab for which WHR is to be generated!!!"
         Exit Sub
      End If
   Else
      If Val(MSHFlexGrid4.TextMatrix(0, 5)) = 0 Then
         MsgBox "Select atleast one lot from 'GSL Detail' tab for which WHR is to be generated!!!"
         Exit Sub
      End If
   End If
End If

If mLotWiseWHR = "1" Then
   If OptSRPerBag.Value = True Then
      TxtStorageAmount = Val(MSHFlexGrid4.TextMatrix(0, 5)) * Val(TxtStorageCharges)
   ElseIf OptSRPerMT.Value = True Then
      TxtStorageAmount = Format(Val(MSHFlexGrid4.TextMatrix(0, 6)) * Val(TxtStorageCharges), "#########0.00")
   End If
   TxtValueofCommodity = Format(Val(MSHFlexGrid4.TextMatrix(0, 6)) * Val(TxtRate), "#########0.00")
End If


If Edit_Flg = "A" Then
   RsTxn_WHRDetail.AddNew
   RsTxn_WHRDetail!WHRNo = GetWHRNo
   RsTxn_WHRDetail!G_UID = GetGUID
   TxtWHRNo.Text = RsTxn_WHRDetail!WHRNo
Else
  RsTxn_WHRDetail.MoveFirst
  RsTxn_WHRDetail.Find "WHRNo= " & TxtWHRNo.Text & ""
  If Not RsTxn_WHRDetail.EOF Then
    If RsTxn_WHRDetail!WHRNo <> TxtWHRNo.Text Then
      MsgBox "Duplicate WHR No not allowed!!! "
     'TxtISINRevalReqNo.SetFocus
      Exit Sub
    End If
  End If
'  RsTxn_WHRDetail.MoveFirst
'  RsTxn_WHRDetail.Find " WHRNo= " & TxtWHRNo.Text
  
  RsTxn_WHRDetail!ConditionID = mCommodityCondition
  RsTxn_WHRDetail!Grade = TxtGrade
  RsTxn_WHRDetail!StorageCharges = TxtStorageCharges
  RsTxn_WHRDetail!Rate = TxtRate
  
  RsTxn_WHRDetail!PolicyNo = PolicyNo_
  RsTxn_WHRDetail!CoverNoteNo = CNo
  RsTxn_WHRDetail!UpdatedBy = Gen_UserLoginID
  RsTxn_WHRDetail!UpdatedDate = Now
  RsTxn_WHRDetail.Update
  Call Indata
  Exit Sub
End If

RsTxn_WHRDetail!WHRDate = WHRDate.Value
RsTxn_WHRDetail!ConditionID = mCommodityCondition
RsTxn_WHRDetail!StorageFrom = StorageFrom.Value
RsTxn_WHRDetail!StorageTo = StorageTo.Value
RsTxn_WHRDetail!Grade = TxtGrade
RsTxn_WHRDetail!Weigher = TxtWeigher
RsTxn_WHRDetail!Assayer = TxtAssayer
RsTxn_WHRDetail!GradeByAssayer = TxtAssayerCerti
RsTxn_WHRDetail!GradeValidTill = GradeValidUpto.Value
RsTxn_WHRDetail!StorageCharges = TxtStorageCharges
RsTxn_WHRDetail!PolicyNo = PolicyNo_
RsTxn_WHRDetail!CoverNoteNo = CNo
RsTxn_WHRDetail!Rate = TxtRate
RsTxn_WHRDetail!SaidToContain = ChkSaidToContain.Value
RsTxn_WHRDetail!EstimatedWeight = ChkEstimatedWgt.Value
RsTxn_WHRDetail!StorageAmount = Val(TxtStorageAmount)
RsTxn_WHRDetail!ClientID = mClientIDRow
RsTxn_WHRDetail!CommodityID = mCommodityID
RsTxn_WHRDetail!CMPID = mCMPID
RsTxn_WHRDetail!GodownID = mGodownID
RsTxn_WHRDetail!LotWiseWHR = mLotWiseWHR
'RsTxn_WHRDetail!SM_Prefix = mSM_Prefix
If mLotWiseWHR = "1" Then
   RsTxn_WHRDetail!WHRNoOfBags = MSHFlexGrid4.TextMatrix(0, 5)
   RsTxn_WHRDetail!WHRQuantity = MSHFlexGrid4.TextMatrix(0, 6)
   RsTxn_WHRDetail!SM_Prefix = mSM_Prefix
   RsTxn_WHRDetail!LicenseID = mLicenseID
End If

If OptSRPerBag.Value = True Then
   RsTxn_WHRDetail!SRAsPerBag = 1
Else
   RsTxn_WHRDetail!SRAsPerBag = 0
End If
RsTxn_WHRDetail!ConvertedClient = Check1.Value
RsTxn_WHRDetail!SCMFlag = mSCMFlag
'RsTxn_WHRDetail.Update
'Call InData
    

If IsNull(RsTxn_WHRDetail!CreatedBy) = True Or RsTxn_WHRDetail!CreatedBy = "" Then
   RsTxn_WHRDetail!CreatedBy = Gen_UserLoginID
   RsTxn_WHRDetail!CreatedDate = Now
Else
   RsTxn_WHRDetail!UpdatedBy = Gen_UserLoginID
   RsTxn_WHRDetail!UpdatedDate = Now
End If

    
    
If mLotWiseWHR = "0" Then
   For I = 2 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(I, 0) = "*" Then
          RsTxn_WHRDetail!CMPID = GetCMPID(MSHFlexGrid2.TextMatrix(I, 1))
          RsTxn_WHRDetail!GodownID = GetGodownID(MSHFlexGrid2.TextMatrix(I, 2), RsTxn_WHRDetail!CMPID)
          RsTxn_WHRDetail!LicenseID = GetLicenseID(MSHFlexGrid2.TextMatrix(I, 3))
          If Edit_Flg = "A" Then
             RsTxn_WHRDetail!oldLicenseID = RsTxn_WHRDetail!LicenseID 'GetLicenseID(MSHFlexGrid2.TextMatrix(i, 2))
          End If
          RsTxn_WHRDetail!WHRNoOfBags = MSHFlexGrid2.TextMatrix(I, 7)
          RsTxn_WHRDetail!WHRQuantity = MSHFlexGrid2.TextMatrix(I, 8)
          RsTxn_WHRDetail!SM_Prefix = MSHFlexGrid2.TextMatrix(I, 12)
          RsTxn_WHRDetail.Update
          Call PostGSLData(RsTxn_WHRDetail!CMPID, RsTxn_WHRDetail!GodownID, RsTxn_WHRDetail!LicenseID, RsTxn_WHRDetail!WHRNoOfBags, RsTxn_WHRDetail!WHRQuantity)
          Exit For
        End If
   Next
Else
   RsTxn_WHRDetail.Update
   mGSLIDinGrid = "(0"
   For I = 2 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(I, 0) = strChecked Then
          mGSLIDinGrid = mGSLIDinGrid & "," & Val(MSHFlexGrid4.TextMatrix(I, 8))
       End If
   Next
   mGSLIDinGrid = mGSLIDinGrid & ")"
 
   tmp = "Select * from Mst_GSL Where SM_Prefix = '" & mSM_Prefix & "' And Mst_GSL_ID in " & mGSLIDinGrid & " And WHRNo is null"
   Call TmpTable
   For I = 2 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(I, 0) = strChecked Then
          TmpRs.MoveFirst
          TmpRs.Find "Mst_GSL_ID = " & Val(MSHFlexGrid4.TextMatrix(I, 8)) & ""
          If TmpRs.EOF = False Then
             TmpRs!WHRBags = Val(MSHFlexGrid4.TextMatrix(I, 5))
             TmpRs!WHRWeight = Val(MSHFlexGrid4.TextMatrix(I, 6))
             TmpRs!WHRNo = TxtWHRNo
             TmpRs.Update
          End If
       End If
   Next
End If

If mSCMFlag = 1 And Edit_Flg = "A" Then
   Call TableTxn_PledgeCreation
   If mLotWiseWHR = "0" Then
      For I = 2 To MSHFlexGrid2.Rows - 1
          If MSHFlexGrid2.TextMatrix(I, 0) = "*" Then
     
             RsTxn_PledgeCreation.AddNew
             RsTxn_PledgeCreation!PledgeID = GetPledgeID
             mSCMPledgeID = RsTxn_PledgeCreation!PledgeID
             RsTxn_PledgeCreation!G_UID = GetGUID
             mPledgeGUID = RsTxn_PledgeCreation!G_UID
             RsTxn_PledgeCreation!PledgeDate = WHRDate.Value
             RsTxn_PledgeCreation!WHRNo = TxtWHRNo.Text
             RsTxn_PledgeCreation!ExchangeTypeID = 1 'mExchangeTypeID
             RsTxn_PledgeCreation!ClientID = mClientIDRow
             RsTxn_PledgeCreation!SM_Prefix = mSM_Prefix

             RsTxn_PledgeCreation!BankID = GetBankID("SCM Bank")
             RsTxn_PledgeCreation!BranchID = GetBranchID("SCM Branch")
             RsTxn_PledgeCreation!PledgeReqNo = "SCM WHR " & TxtWHRNo & " " 'TxtPledgeReqNo
             RsTxn_PledgeCreation!PledgeFees = "500" 'Val(TxtPledgeFee)
             RsTxn_PledgeCreation!PledgeReqDate = WHRDate.Value ' PledgeRequestDate.Value
             RsTxn_PledgeCreation!Discount = "500" 'Val(TxtDiscount)
             RsTxn_PledgeCreation!SCMFlag = 1
             RsTxn_PledgeCreation!CreatedBy = Gen_UserLoginID
             RsTxn_PledgeCreation!CreatedDate = Now
             RsTxn_PledgeCreation.Update
             tmp = "Select * from  Txn_PledgeCreation Where G_UID =  '" & mPledgeGUID & "' "
             Call Tmp1Table
             If TmpRs1.RecordCount > 0 Then
                mSCMPledgeID = TmpRs1!Txn_PledgeCreation_ID
             End If
             tmp = "Update Txn_WHR_Detail Set Txn_PledgeCreation_ID = " & Val(mSCMPledgeID) & " Where WHRNo =  " & TxtWHRNo.Text & " "
             Call Tmp1Table
             Call SavePledgeCreation_GSL
          End If
      Next
   Else
         RsTxn_PledgeCreation.AddNew
         RsTxn_PledgeCreation!PledgeID = GetPledgeID
         mSCMPledgeID = RsTxn_PledgeCreation!PledgeID
         RsTxn_PledgeCreation!G_UID = GetGUID
         mPledgeGUID = RsTxn_PledgeCreation!G_UID
         RsTxn_PledgeCreation!PledgeDate = WHRDate.Value
         RsTxn_PledgeCreation!WHRNo = TxtWHRNo.Text
         RsTxn_PledgeCreation!ExchangeTypeID = 1 'mExchangeTypeID
         RsTxn_PledgeCreation!ClientID = mClientIDRow
         RsTxn_PledgeCreation!SM_Prefix = mSM_Prefix
         RsTxn_PledgeCreation!BankID = GetBankID("SCM Bank")
         RsTxn_PledgeCreation!BranchID = GetBranchID("SCM Branch")
         RsTxn_PledgeCreation!PledgeReqNo = "SCM WHR " & TxtWHRNo & " " 'TxtPledgeReqNo
         RsTxn_PledgeCreation!PledgeFees = "500" 'Val(TxtPledgeFee)
         RsTxn_PledgeCreation!PledgeReqDate = WHRDate.Value ' PledgeRequestDate.Value
         RsTxn_PledgeCreation!Discount = "500" 'Val(TxtDiscount)
         RsTxn_PledgeCreation!SCMFlag = 1
         RsTxn_PledgeCreation!CreatedBy = Gen_UserLoginID
         RsTxn_PledgeCreation!CreatedDate = Now
         RsTxn_PledgeCreation.Update

         tmp = "Select * from  Txn_PledgeCreation Where G_UID =  '" & mPledgeGUID & "' "
         Call Tmp1Table
         If TmpRs1.RecordCount > 0 Then
            mSCMPledgeID = TmpRs1!Txn_PledgeCreation_ID
         End If

         tmp = "Update Txn_WHR_Detail Set Txn_PledgeCreation_ID = " & Val(mSCMPledgeID) & " Where WHRNo =  " & TxtWHRNo.Text & " "
         Call Tmp1Table
         
         Call SavePledgeCreation_GSL
   
   End If
End If

Call Indata

End Sub
Private Function GetPledgeID() As Double

Dim Retval As Double
tmp = "Select max(PledgeID) from Txn_PledgeCreation where SM_Prefix = '" & mSM_Prefix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetPledgeID = Retval

End Function
Private Function GetPledgeNo() As Double
Dim Retval As Double
tmp = "Select max(PledgeNo) from Txn_PledgeCreation_GSL where SM_Prefix = '" & mSM_Prefix & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetPledgeNo = Retval
End Function

Private Sub SavePledgeCreation_GSL()

Call TableTxn_PledgeCreation_GSL("Where SM_Prefix = '" & mSM_Prefix & "' And Txn_PledgeCreation_ID = " & mSCMPledgeID)


If mLotWiseWHR = "1" Then
   If mGSLIDinGrid <> "" Then
      Call TableMst_GSL(" Where ExchangeTypeID = 1 And SCMFlag = '1' And Mst_GSL.Status = 'O' And Mst_GSL.MadeupFLag = '0' And SM_Prefix = '" & mSM_Prefix & "' And Mst_GSL_ID in " & mGSLIDinGrid & " ")
   End If

   For I = 2 To MSHFlexGrid4.Rows - 1
       If MSHFlexGrid4.TextMatrix(I, 0) = strChecked Then
          RsTxn_PledgeCreation_GSL.AddNew
          RsTxn_PledgeCreation_GSL!G_UID = GetGUID
          RsTxn_PledgeCreation_GSL!PledgeNo = GetPledgeNo
          RsTxn_PledgeCreation_GSL!Txn_PledgeCreation_ID = mSCMPledgeID
          RsTxn_PledgeCreation_GSL!Mst_GSL_ID = Val(MSHFlexGrid4.TextMatrix(I, 8))
          RsTxn_PledgeCreation_GSL!PledgeNoOfBags = Val(MSHFlexGrid4.TextMatrix(I, 5))
          RsTxn_PledgeCreation_GSL!PledgeQuantity = Val(MSHFlexGrid4.TextMatrix(I, 6))
          RsTxn_PledgeCreation_GSL!SM_Prefix = mSM_Prefix
          RsTxn_PledgeCreation_GSL.Update
          RsMst_GSL.MoveFirst
          RsMst_GSL.Find "Mst_GSL_ID = " & Val(MSHFlexGrid4.TextMatrix(I, 8)) & " "
          If Not RsMst_GSL.EOF Then
             If IsNull(RsMst_GSL!PledgeNo) Or RsMst_GSL!PledgeNo = 0 Or RsMst_GSL!PledgeNo = "" Then
                RsMst_GSL!PledgeNo = mSCMPledgeID
             Else
                RsMst_GSL!PledgeNo = Trim(RsMst_GSL!PledgeNo) & "~" & mSCMPledgeID
             End If
             RsMst_GSL!PledgeBags = Val(IIf(IsNull(RsMst_GSL!PledgeBags), "0", RsMst_GSL!PledgeBags)) + Val(MSHFlexGrid4.TextMatrix(I, 5))
             RsMst_GSL!PledgeWeight = Val(IIf(IsNull(RsMst_GSL!PledgeWeight), "0", RsMst_GSL!PledgeWeight)) + Val(MSHFlexGrid4.TextMatrix(I, 6)) 'MSHFlexGrid2.TextMatrix(i, 8)
             RsMst_GSL.Update
          End If
          RsTxn_PledgeCreation_GSL.Requery
       End If
   Next
Else
    tmp = "SELECT ((Mst_GSL.Mst_GSL_ID)) as Mst_GSL_ID, ((Mst_GSL.DepositeBags)) AS BalanceBags, ((Mst_GSL.DepositeWeight)) AS BalanceWeight, "
    tmp = tmp & " ((Mst_GSL.WHRBags)) AS WHRBags, ((Mst_GSL.WHRWeight)) AS WHRWeight, "
    tmp = tmp & " Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName, "
    tmp = tmp & " Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientID,  Mst_License.LicenseID,Mst_Godown.GodownNo, "
    tmp = tmp & " Mst_Godown.GodownID,Mst_Godown.AgreementEndDate "
    tmp = tmp & " FROM Mst_CMP INNER JOIN (Mst_License INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = "
    tmp = tmp & " Mst_GSL.GodownID) ON Mst_License.LicenseID = Mst_Godown.LicenseID) ON Mst_CMP.CMPID = Mst_GSL.CMPID "
    tmp = tmp & " WHERE ( ((Mst_GSL.CMPID)=" & mCMPID & ") AND ((Mst_GSL.GodownID)=" & mGodownID & ") AND "
    tmp = tmp & " ((Mst_GSL.ClientID)=" & mClientIDRow & ") AND ((Mst_GSL.ExchangeTypeID)= 1 ) AND "
    tmp = tmp & " ((Mst_GSL.CommodityID)=" & mCommodityID & ") AND "
    tmp = tmp & " ((Mst_GSL.CDTFDate) Between '" & Format(StorageFrom, Gen_DatFormat) & "' And '" & Format(WHRDate, Gen_DatFormat) & "')) "
    If mGSLForPledge <> "" Then
       tmp = tmp & " And Mst_GSL.Status <> 'Z' And Mst_GSL.MadeupFLag = '0' and mst_gsl.CxTF_DetailsID > 0  And Mst_GSL.SCMFlag = '1' And Mst_GSL.Mst_GSL_ID in " & mGSLForPledge & ""
    Else
       tmp = tmp & " And Mst_GSL.Status <> 'Z' And Mst_GSL.MadeupFLag = '0' and mst_gsl.CxTF_DetailsID > 0  And Mst_GSL.SCMFlag = '1' "
    End If
    tmp = tmp & " Union "
    tmp = tmp & " Select GSL.Mst_GSL_ID, GSL.DepositeBags AS BalanceBags, GSL.DepositeWeight AS BalanceWeight, GSL.WHRBags AS WHRBags, "
    tmp = tmp & " GSL.WHRWeight AS WHRWeight, MLC.LicenseNo, MLC.LicenseDesc, CMP.CMPName, GSL.CommodityID,  GSL.CMPID, GSL.ClientID,  MLC.LicenseID,MG.GodownNo, MG.GodownID,MG.AgreementEndDate "
    tmp = tmp & " From Mst_GSL  GSL "
    tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = GSL.GodownID "
    tmp = tmp & " Inner Join Mst_CMP CMP ON CMP.CMPID = GSL.CMPID "
    tmp = tmp & " Inner Join Mst_License MLC ON MLC.LicenseID = MG.LicenseID"
    tmp = tmp & " Inner Join Txn_RRNPWRGSL TRPG on (TRPG.NewMst_GSL_ID  = GSL.Mst_GSL_ID Or TRPG.NewTransferMst_GSL_ID  = GSL.Mst_GSL_ID )  and GSL.SM_prefix =  TRPG.SM_Prefix And TRPG.ConvertedID in (Select ConvertedID from Txn_RRNPWR Where Flag = 'A') "
    tmp = tmp & " WHERE GSL.CMPID=" & mCMPID & " AND GSL.GodownID=" & mGodownID & " AND "
    tmp = tmp & " GSL.ClientID=" & mClientID & " AND GSL.ExchangeTypeID=1 AND "
    tmp = tmp & " GSL.CommodityID=" & mCommodityID & " AND "
    tmp = tmp & " GSL.CDTFDate Between '" & Format(StorageFrom, Gen_DatFormat) & "' And '" & Format(WHRDate, Gen_DatFormat) & "'"
    tmp = tmp & " And GSL.Status <> 'Z' And GSL.MadeupFLag = '0' And GSL.SCMFlag = '1' "
    If mGSLForPledge <> "" Then
       tmp = tmp & " And GSL.CxTF_DetailsID < 0 ANd GSL.Mst_GSL_ID in " & mGSLForPledge & ""
    Else
       tmp = tmp & " And GSL.CxTF_DetailsID < 0 "
    End If
    
    Call Tmp1Table
    
    
    For I = 1 To TmpRs1.RecordCount
        RsTxn_PledgeCreation_GSL.AddNew
        RsTxn_PledgeCreation_GSL!G_UID = GetGUID
        RsTxn_PledgeCreation_GSL!PledgeNo = GetPledgeNo
        RsTxn_PledgeCreation_GSL!Txn_PledgeCreation_ID = mSCMPledgeID
        RsTxn_PledgeCreation_GSL!Mst_GSL_ID = TmpRs1!Mst_GSL_ID
        RsTxn_PledgeCreation_GSL!PledgeNoOfBags = TmpRs1!BalanceBags
        RsTxn_PledgeCreation_GSL!PledgeQuantity = TmpRs1!BalanceWeight
        RsTxn_PledgeCreation_GSL!SM_Prefix = mSM_Prefix
        RsTxn_PledgeCreation_GSL.Update
        RsMst_GSL.MoveFirst
        RsMst_GSL.Find "Mst_GSL_ID = " & TmpRs1!Mst_GSL_ID
        If Not RsMst_GSL.EOF Then
           If IsNull(RsMst_GSL!PledgeNo) Or RsMst_GSL!PledgeNo = 0 Or RsMst_GSL!PledgeNo = "" Then
              RsMst_GSL!PledgeNo = mSCMPledgeID
           Else
              RsMst_GSL!PledgeNo = Trim(RsMst_GSL!PledgeNo) & "~" & mSCMPledgeID
           End If
           RsMst_GSL!PledgeBags = Val(IIf(IsNull(RsMst_GSL!PledgeBags), "0", RsMst_GSL!PledgeBags)) + Val(TmpRs1!BalanceBags)
           RsMst_GSL!PledgeWeight = Val(IIf(IsNull(RsMst_GSL!PledgeWeight), "0", RsMst_GSL!PledgeWeight)) + Val(TmpRs1!BalanceWeight) 'MSHFlexGrid2.TextMatrix(i, 8)
           RsMst_GSL.Update
        End If
        RsTxn_PledgeCreation_GSL.Requery
        TmpRs1.MoveNext
    Next

End If



End Sub
  
  
'  End If
'Next


'Call InData
'End Sub
Private Sub InsertPledgeData()




End Sub


Function PostGSLData(CMPID_ As Double, GodownID_ As Double, LicenseID_ As Double, WHRNoBags_ As Double, WHRQty_ As Double)
Dim TBags, TQty, PBags, PQty As Variant
mGSLForPledge = ""
Call TableMst_GSL(" Where SM_Prefix = '" & mSM_Prefix & "' And  CMPID = " & CMPID_)
    '"((Mst_GSL.CDTFDate) <= '" & Format(WHRDate, Gen_DatFormat) & "')) " & _

tmp = "SELECT ((Mst_GSL.Mst_GSL_ID)) as Mst_GSL_ID, ((Mst_GSL.DepositeBags)) AS BalanceBags, ((Mst_GSL.DepositeWeight)) AS BalanceWeight, "
tmp = tmp & " ((Mst_GSL.WHRBags)) AS WHRBags, ((Mst_GSL.WHRWeight)) AS WHRWeight, "
tmp = tmp & " Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName, "
tmp = tmp & " Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientID,  Mst_License.LicenseID,Mst_Godown.GodownNo, "
tmp = tmp & " Mst_Godown.GodownID,Mst_Godown.AgreementEndDate "
tmp = tmp & " FROM Mst_CMP INNER JOIN (Mst_License INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = "
tmp = tmp & " Mst_GSL.GodownID) ON Mst_License.LicenseID = Mst_Godown.LicenseID) ON Mst_CMP.CMPID = Mst_GSL.CMPID "
tmp = tmp & " WHERE ( ((Mst_GSL.CMPID)=" & CMPID_ & ") AND ((Mst_GSL.GodownID)=" & GodownID_ & ") AND "
tmp = tmp & " ((Mst_GSL.ClientID)=" & mClientIDRow & ") AND ((Mst_GSL.ExchangeTypeID)=" & mExchangeTypeID & ") AND "
tmp = tmp & " ((Mst_GSL.CommodityID)=" & mCommodityID & ") AND "
tmp = tmp & " ((Mst_GSL.CDTFDate) Between '" & Format(StorageFrom, Gen_DatFormat) & "' And '" & Format(WHRDate, Gen_DatFormat) & "')) "
tmp = tmp & " And Mst_GSL.Status = 'E' And Mst_GSL.MadeupFLag = '0' and mst_gsl.CxTF_DetailsID > 0  And mst_gsl.WHRBags >= 0 And mst_Gsl.WHRBags < mst_GSL.DepositeBags "
tmp = tmp & " Union "
tmp = tmp & " SELECT ((Mst_GSL.Mst_GSL_ID)) as Mst_GSL_ID, ((Mst_GSL.DepositeBags)) AS BalanceBags, ((Mst_GSL.DepositeWeight)) AS BalanceWeight, "
tmp = tmp & " ((Mst_GSL.WHRBags)) AS WHRBags, ((Mst_GSL.WHRWeight)) AS WHRWeight, "
tmp = tmp & " Mst_License.LicenseNo, Mst_License.LicenseDesc, Mst_CMP.CMPName, "
tmp = tmp & " Mst_GSL.CommodityID,  Mst_GSL.CMPID, Mst_GSL.ClientID,  Mst_License.LicenseID,Mst_Godown.GodownNo, "
tmp = tmp & " Mst_Godown.GodownID,Mst_Godown.AgreementEndDate "
tmp = tmp & " FROM Mst_CMP INNER JOIN (Mst_License INNER JOIN (Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = "
tmp = tmp & " Mst_GSL.GodownID) ON Mst_License.LicenseID = Mst_Godown.LicenseID) ON Mst_CMP.CMPID = Mst_GSL.CMPID "
tmp = tmp & " WHERE ( ((Mst_GSL.CMPID)=" & CMPID_ & ") AND ((Mst_GSL.GodownID)=" & GodownID_ & ") AND "
tmp = tmp & " ((Mst_GSL.ClientID)=" & mClientIDRow & ") AND ((Mst_GSL.ExchangeTypeID)=" & mExchangeTypeID & ") AND "
tmp = tmp & " ((Mst_GSL.CommodityID)=" & mCommodityID & ") AND "
tmp = tmp & " ((Mst_GSL.CDTFDate) Between '" & Format(StorageFrom, Gen_DatFormat) & "' And '" & Format(WHRDate, Gen_DatFormat) & "')) "
tmp = tmp & " And Mst_GSL.Status = 'O' And Mst_GSL.MadeupFLag = '0' and mst_gsl.CxTF_DetailsID > 0 And mst_gsl.WHRBags > 0 And mst_Gsl.WHRBags < mst_GSL.DepositeBags "
tmp = tmp & " Union "
tmp = tmp & " Select GSL.Mst_GSL_ID, GSL.DepositeBags AS BalanceBags, GSL.DepositeWeight AS BalanceWeight, GSL.WHRBags AS WHRBags, "
tmp = tmp & " GSL.WHRWeight AS WHRWeight, MLC.LicenseNo, MLC.LicenseDesc, CMP.CMPName, GSL.CommodityID,  GSL.CMPID, GSL.ClientID,  MLC.LicenseID,MG.GodownNo, MG.GodownID,MG.AgreementEndDate "
tmp = tmp & " From Mst_GSL  GSL "
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = GSL.GodownID "
tmp = tmp & " Inner Join Mst_CMP CMP ON CMP.CMPID = GSL.CMPID "
tmp = tmp & " Inner Join Mst_License MLC ON MLC.LicenseID = MG.LicenseID"
tmp = tmp & " Inner Join Txn_RRNPWRGSL TRPG on (TRPG.NewMst_GSL_ID  = GSL.Mst_GSL_ID Or TRPG.NewTransferMst_GSL_ID  = GSL.Mst_GSL_ID )  and GSL.SM_prefix =  TRPG.SM_Prefix And TRPG.ConvertedID in (Select ConvertedID from Txn_RRNPWR Where Flag = 'A') "
tmp = tmp & " WHERE GSL.CMPID=" & CMPID_ & " AND GSL.GodownID=" & GodownID_ & " AND "
tmp = tmp & " GSL.ClientID=" & mClientIDRow & " AND GSL.ExchangeTypeID=" & mExchangeTypeID & " AND "
tmp = tmp & " GSL.CommodityID=" & mCommodityID & " AND "
tmp = tmp & " GSL.CDTFDate Between '" & Format(StorageFrom, Gen_DatFormat) & "' And '" & Format(WHRDate, Gen_DatFormat) & "'"
tmp = tmp & " And GSL.Status = 'E' And GSL.MadeupFLag = '0' And GSL.WHRBags >= 0 And GSL.WHRBags < GSL.DepositeBags"
tmp = tmp & " And GSL.CxTF_DetailsID < 0 "
tmp = tmp & " Union "
tmp = tmp & " Select GSL.Mst_GSL_ID, GSL.DepositeBags AS BalanceBags, GSL.DepositeWeight AS BalanceWeight, GSL.WHRBags AS WHRBags, "
tmp = tmp & " GSL.WHRWeight AS WHRWeight, MLC.LicenseNo, MLC.LicenseDesc, CMP.CMPName, GSL.CommodityID,  GSL.CMPID, GSL.ClientID,  MLC.LicenseID,MG.GodownNo, MG.GodownID,MG.AgreementEndDate "
tmp = tmp & " From Mst_GSL  GSL "
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = GSL.GodownID "
tmp = tmp & " Inner Join Mst_CMP CMP ON CMP.CMPID = GSL.CMPID "
tmp = tmp & " Inner Join Mst_License MLC ON MLC.LicenseID = MG.LicenseID"
tmp = tmp & " Inner Join Txn_RRNPWRGSL TRPG on (TRPG.NewMst_GSL_ID  = GSL.Mst_GSL_ID Or TRPG.NewTransferMst_GSL_ID  = GSL.Mst_GSL_ID )  and GSL.SM_prefix =  TRPG.SM_Prefix And TRPG.ConvertedID in (Select ConvertedID from Txn_RRNPWR Where Flag = 'A') "
tmp = tmp & " WHERE GSL.CMPID=" & CMPID_ & " AND GSL.GodownID=" & GodownID_ & " AND "
tmp = tmp & " GSL.ClientID=" & mClientIDRow & " AND GSL.ExchangeTypeID=" & mExchangeTypeID & " AND "
tmp = tmp & " GSL.CommodityID=" & mCommodityID & " AND "
tmp = tmp & " GSL.CDTFDate Between '" & Format(StorageFrom, Gen_DatFormat) & "' And '" & Format(WHRDate, Gen_DatFormat) & "'"
tmp = tmp & " And GSL.Status = 'O' And GSL.MadeupFLag = '0' "
tmp = tmp & " And GSL.CxTF_DetailsID < 0  And GSL.WHRBags > 0 And GSL.WHRBags < GSL.DepositeBags "
'tmp = tmp & " GROUP BY MLC.LicenseNo, MLC.LicenseDesc, CMP.CMPName , GSL.CommodityID, GSL.CMPID, GSL.ClientIDRow,  MLC.LicenseID, MG.GodownNo,MG.GodownID,MG.AgreementEndDate"
Call Tmp1Table

TBags = WHRNoBags_
TQty = WHRQty_
For I = 1 To TmpRs1.RecordCount
    If TmpRs1!BalanceBags - TmpRs1!WHRBags > 0 Then
        If TBags >= (TmpRs1!BalanceBags - TmpRs1!WHRBags) Then
           PBags = TmpRs1!BalanceBags - TmpRs1!WHRBags
           TBags = TBags - PBags
        Else
           PBags = TBags
           TBags = 0
        End If
        RsMst_GSL.MoveFirst
        RsMst_GSL.Find "Mst_GSL_ID = " & TmpRs1!Mst_GSL_ID
        If Not RsMst_GSL.EOF Then
            RsMst_GSL!WHRBags = IIf(IsNull(TmpRs1!WHRBags), 0, TmpRs1!WHRBags) + PBags
            RsMst_GSL.Update
        End If
        If mSCMFlag = "1" Then
           If mGSLForPledge = "" Then
              mGSLForPledge = "(" & TmpRs1!Mst_GSL_ID
           Else
              mGSLForPledge = mGSLForPledge & "," & TmpRs1!Mst_GSL_ID
           End If
        End If
        'TmpRs1!WHRBags = IIf(IsNull(TmpRs1!WHRBags), 0, TmpRs1!WHRBags) + PBags
        ''TmpRs1!BalanceBags = TmpRs1!BalanceBags - PBags
        'TmpRs1.Update
        If TBags = 0 Then Exit For
    End If
    TmpRs1.MoveNext
Next

If mSCMFlag = "1" Then
   If mGSLForPledge <> "" Then
      mGSLForPledge = mGSLForPledge & ")"
   End If
End If

TmpRs1.MoveFirst
TQty = WHRQty_
For I = 1 To TmpRs1.RecordCount
    If TmpRs1!BalanceWeight - TmpRs1!WHRWeight > 0 Then
       If TQty >= (TmpRs1!BalanceWeight - TmpRs1!WHRWeight) Then
          PQty = TmpRs1!BalanceWeight - TmpRs1!WHRWeight
          TQty = TQty - PQty
       Else
          PQty = TQty
          TQty = 0
       End If
       RsMst_GSL.MoveFirst
       RsMst_GSL.Find "Mst_GSL_ID = " & TmpRs1!Mst_GSL_ID
       If Not RsMst_GSL.EOF Then
           RsMst_GSL!WHRWeight = IIf(IsNull(TmpRs1!WHRWeight), 0, TmpRs1!WHRWeight) + PQty
           RsMst_GSL.Update
       End If
       'TmpRs1!WHRWeight = IIf(IsNull(TmpRs1!WHRWeight), 0, TmpRs1!WHRWeight) + PQty
       ''TmpRs1!BalanceWeight = TmpRs1!BalanceWeight - PQty
       'TmpRs1.Update
       If TQty = 0 Then Exit For
    End If
    TmpRs1.MoveNext
Next
TmpRs1.Close

End Function
'Private Function GetWHRNo() As Double
'Dim Retval As Double
'tmp = "Select WHRNo from Txn_WHR_Detail Where CMPID = " & mCMPID & " Order by WHRNo"
'Call TmpTable
'
'If Not TmpRs.EOF Then
'    TmpRs.MoveLast
'    mWHRNo = TmpRs.Fields(0)
'    mWHRNo = mWHRNo + 1
'Else
'    mWHRNo = mCMPID * 1000000
'    mWHRNo = mWHRNo + 1
'End If
'
''TmpRs.MoveLast
'Retval = mWHRNo 'IIf(IsNull(TmpRs!WHRNo), 0, TmpRs!WHRNo)
'GetWHRNo = Retval
'End Function
Private Function GetWHRNo() As Double

Dim Retval As Double
Dim LogRetVal As Double
tmp = "Select WHRNo from Txn_WHR_Detail Where CMPID = " & mCMPID & " Order by WHRNo"
Call TmpTable

tmp = "Select WHRNo from Log_WHR_Detail Where CMPID = " & mCMPID & " Order by WHRNo"
Call Tmp1Table

If Not TmpRs.EOF Then
    TmpRs.MoveLast
    mWHRNo = TmpRs.Fields(0)
    mWHRNo = mWHRNo + 1
Else
    mWHRNo = mCMPID * 1000000
    mWHRNo = mWHRNo + 1
End If

If Not TmpRs1.EOF Then
    TmpRs1.MoveLast
    mLogWHRNo = TmpRs1.Fields(0)
    mLogWHRNo = mLogWHRNo + 1
Else
    mLogWHRNo = mCMPID * 1000000
    mLogWHRNo = mLogWHRNo + 1
End If

If (mWHRNo - 1) > (mLogWHRNo - 1) Then
   Retval = mWHRNo 'IIf(IsNull(TmpRs!WHRNo), 0, TmpRs!WHRNo)
Else
   Retval = mLogWHRNo
End If

GetWHRNo = Retval
End Function

Private Function GetWHRLicenseDetailID() As Double
Dim Retval As Double
tmp = "Select max(WHRLicenseDetailID) from Txn_WHRLicense_Detail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetWHRLicenseDetailID = Retval
End Function

Private Sub StorageFrom_LostFocus()
Call GetGridData(mExchangeTypeID, mCommodityID, mClientIDRow)
StorageTo.Value = DateAdd("M", 9, StorageFrom.Value)


End Sub
Private Sub StorageTo_LostFocus()
Dim mSMonth, mSYear, mSTotalDays As Variant
mSMonth = Month(StorageTo.Value)
mSYear = Year(StorageTo.Value)
mSTotalDays = 20
StorageTo.Value = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
GradeValidUpto.Value = StorageTo.Value

End Sub

Private Sub TxtRate_Lostfocus()
If TxtRate = "" Then Exit Sub
If IsNumeric(TxtRate) = False Then
  MsgBox "Invalid numeric format!!!"
  TxtRate.SetFocus
  Exit Sub
End If
End Sub

Private Sub TxtStorageCharges_LostFocus()
If TxtStorageCharges = "" Then Exit Sub
If IsNumeric(TxtStorageCharges) = False Then
  MsgBox "Invalid numeric format!!!"
  TxtStorageCharges.SetFocus
  Exit Sub
End If
End Sub

Private Sub TxtWHRNoOfBags_LostFocus()
If TxtWHRNoOfBags = "" Then Exit Sub
If IsNumeric(TxtWHRNoOfBags) = False Then
  MsgBox "Invalid numeric format!!!"
  TxtWHRNoOfBags.SetFocus
  Exit Sub
End If
If Val(TxtWHRNoOfBags) < 0 Then
  MsgBox "WHR No oF Bags Can't be negetive"
  TxtWHRNoOfBags.SetFocus
  Exit Sub
End If
If Val(TxtWHRNoOfBags) > Val(TxtNoofBags) Then
   MsgBox "WHR No Of Bags can't be more than available number of bags!!!"
   TxtWHRNoOfBags.SetFocus
   Exit Sub
End If
TxtWHRQty.SetFocus
End Sub

Private Sub TxtWHRQty_LostFocus()
If TxtWHRQty = "" Then Exit Sub
If IsNumeric(TxtWHRQty) = False Then
  MsgBox "Invalid numeric format!!!"
  TxtWHRQty.SetFocus
  Exit Sub
End If
If Val(TxtWHRQty) < 0 Then
  MsgBox "WHR Quantity Can't be negetive!!!"
  TxtWHRQty.SetFocus
  Exit Sub
End If
If Val(TxtWHRQty) > Val(TxtEleWeight) Then
   MsgBox "WHR Qty can't be more than available quantity!!!"
   TxtWHRQty.SetFocus
   Exit Sub
End If
CmdUpdate.Enabled = True
'CmdUpdate.SetFocus
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_WHRDetail.MoveNext
If RsTxn_WHRDetail.EOF Then
   RsTxn_WHRDetail.MoveLast
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
RsTxn_WHRDetail.MovePrevious
If RsTxn_WHRDetail.BOF Then
   RsTxn_WHRDetail.MoveFirst
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
Private Sub FirstCmd_Click()
RsTxn_WHRDetail.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsTxn_WHRDetail.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub SearchCmd_Click()
Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_WHRDetail.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsTxn_WHRDetail.EOF Then
      RsTxn_WHRDetail.MoveFirst
   End If
   Call Indata
   Exit Sub
End If
RsTxn_WHRDetail.Filter = ""

'Frame0.Height = Me.Height - 500
'Frame0.Left = Frame1.Left
'Frame0.Top = Frame1.Top
'Frame0.Width = Me.Width - 400
'Frame0.Visible = True


Frame0.Height = Frame4.Height - 200 'Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width - 200 'Me.Width - 200
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 300
MSHFlexGrid1.Height = Frame0.Height - 2500

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'Call Grid_Head1
'CmbPreFix.Text = ""
'CmbSCmp.Text = ""
'CmbSClientName.Text = ""
'CmbSCommodity.Text = ""
'TxtSWHRNo = ""
'DTPWHRFromDate.Value = Date - 365
'DTPWHRToDate.Value = Date
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub CmdGo_Click()
Dim wc As String
Dim Wc1 As Variant
wc = ""
tmp = ""
'
'If CmbPreFix.Text = "" Then
'   MsgBox "Select Location."
'   Exit Sub
'End If

If DTPWHRFromDate.Value > DTPWHRToDate.Value Then
   MsgBox "'From Date' Must be less than 'To Date' !!!"
   DTPWHRFromDate.SetFocus
   Exit Sub
End If
'
tmp = "SELECT  TWD.WHRNo,TWD.SM_Prefix,CMP.CMPName,CM.ClientName,ComM.Commodity,TWD.WHRDate,TWD.ConditionID,TWD.StorageFrom,TWD.StorageTo,TWD.Grade,TWD.Weigher,TWD.Assayer,TWD.GradeByAssayer"
tmp = tmp & " ,TWD.GradeValidTill,TWD.StorageCharges,TWD.StorageAmount,TWD.ClientID,TWD.CommodityID,TWD.CMPID,TWD.LicenseID,TWD.WHRNoofBags,TWD.WHRQuantity, TWD.PrintFlag,TWD.Remark,TWD.Flag,TWD.Txn_PledgeCreation_ID,TWD.OldLicenseID,TWD.GodownID,TWD.DeliveredQuantity,TWD.DeliveredNoOfBags,TWD.Rate,TWD.G_UID, TWD.HologramNo"
tmp = tmp & " ,TWD.PolicyNo, TWD.CoverNoteNo, TWD.SaidToContain, TWD.EstimatedWeight,TWD.CARNo,TWD.ConvertedClient,TWD.SCMFlag,TWD.WHRQuantity*TWD.Rate as 'ValueofCommodity',TWD.Txn_PledgeCreation_ID,TPC.PledgeDate,MB.BankName,MBr.BranchName,GM.GodownNo,isnull(TWD.LotwiseWHR,0) 'LotwiseWHR',MI.Risks,MI.Amount,MI.ValidFrom,MI.ValidTo,MICo.InsuranceCoName "
tmp = tmp & " FROM  Txn_WHR_Detail TWD"
tmp = tmp & " Inner Join Mst_Location LM ON TWD.SM_Prefix=LM.SM_Prefix"
tmp = tmp & " Inner Join Mst_CMP CMP ON TWD.CMPID=CMP.CMPID"
tmp = tmp & " Inner Join Mst_Godown GM ON TWD.GodownID=GM.GodownID"
tmp = tmp & " Inner Join Mst_Client CM ON TWD.ClientID=CM.ClientID"
tmp = tmp & " Inner Join Mst_Commodity ComM ON TWD.CommodityID=ComM.CommodityID"
tmp = tmp & " left join Txn_PledgeCreation TPC on  TWD.Txn_PledgeCreation_ID = TPC.Txn_PledgeCreation_ID And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
tmp = tmp & " left Join Mst_Bank MB on MB.BankID = TPC.BankID"
tmp = tmp & " left Join Mst_Branch MBr on MBr.BranchID = TPC.BranchID"
tmp = tmp & " Left Join Mst_InsuranceDetails MI on  TWD.PolicyNo = MI.PolicyNo ANd TWD.CoverNoteNo = MI.CoverNoteNo"
tmp = tmp & " Left Join Mst_InsuranceCoMaster MICo On MI.InsuranceCoID = MICO.InsuranceCoID"
If Gen_DBType = "MDB" Then
   wc = " Where TWD.WHRDate Between #" & Format(DTPWHRFromDate.Value, Gen_DatFormat) & "# And #" & Format(DTPWHRToDate.Value, Gen_DatFormat) & "# "
Else
   wc = " Where TWD.WHRDate Between '" & Format(DTPWHRFromDate.Value, Gen_DatFormat) & "' And '" & Format(DTPWHRToDate.Value, Gen_DatFormat) & "' "
End If
'
Dim tmp1 As String
If CmbPreFix.Text <> "" Then
   wc = GenWc(wc, CmbPreFix.Text, "LM.LocationName", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp1 = Replace(Gen_WcLocation, "LocationID", "LM.LocationID")
      wc = GenWc(wc, "", tmp1, "N", "I")
   End If
End If

If CmbSCmp.Text <> "" Then
   wc = GenWc(wc, CmbSCmp.Text, "CMP.CmpName", "S", "N")
End If

If CmbSCommodity.Text <> "" Then
   wc = GenWc(wc, CmbSCommodity.Text, "ComM.Commodity", "S", "N")
End If

If CmbSClientName.Text <> "" Then
   wc = GenWc(wc, CmbSClientName.Text, "CM.ClientName", "S", "N")
End If

If TxtSWHRNo.Text <> "" Then
   Dim strSplitString() As String
   Dim mSWHRNo As Variant
   strSplitString = Split(TxtSWHRNo.Text, "-")
   tmp1 = UBound(strSplitString)
   If tmp1 <> 1 Then
      MsgBox "Invalid WHR No selection"
      TxtSWHRNo.SetFocus
      Exit Sub
   End If
   mSWHRNo = Val(strSplitString(0)) * 1000000 + Val(strSplitString(1))
   wc = GenWc(wc, mSWHRNo, "TWD.WHRNo", "", "")
End If

If CmbSearchStatus.Text <> "" Then
   If CmbSearchStatus.Text = "Printed" Then
        wc = GenWc(wc, "P", "TWD.Flag", "S", "N")
   Else
        wc = GenWc(wc, " is null", "TWD.Flag", "S", "I")
   End If
End If

If CmbSSCMFlag.Text <> "" Then
   If LCase(CmbSSCMFlag.Text) = "yes" Then
      wc = wc & " And isnull(TWD.SCMFlag,0) = '1' "
   ElseIf LCase(CmbSSCMFlag.Text) = "no" Then
      wc = wc & " And isnull(TWD.SCMFlag,0) = '0' "
   End If
End If


If ChkSCARGenerated.Value = 1 Then
   wc = wc & " And (TWD.Txn_PledgeCreation_ID is not null and TWD.Txn_PledgeCreation_ID <> 0)  And TWD.CARNo is Null "
End If

If ChkSWHRPledgePrint = 1 Then
   wc = wc & " And (TWD.Txn_PledgeCreation_ID is not null and TWD.Txn_PledgeCreation_ID <> 0) And TWD.Flag <> 'N' "
End If

tmp1 = "Order By LM.LocationName"
tmp = tmp & " " & wc & " " & tmp1

Call Tmp3Table
Call Grid_Head1

MSHFlexGrid1.Rows = TmpRs3.RecordCount + 2
For I = 1 To TmpRs3.RecordCount
   

    MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs3!CMPID), 0, TmpRs3!CMPID) & "-" & TmpRs3!WHRNo - (IIf(IsNull(TmpRs3!CMPID), 0, TmpRs3!CMPID) * 1000000)
    MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs3!SM_Prefix), "", TmpRs3!SM_Prefix)
    MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs3!CMPName), "", TmpRs3!CMPName) 'GetCMPName(TmpRs3!CMPID)
    If TmpRs3!ConditionID = 0 Then
       MSHFlexGrid1.TextMatrix(I, 3) = "Average"
    ElseIf TmpRs3!ConditionID = 1 Then
       MSHFlexGrid1.TextMatrix(I, 3) = "Fair"
    ElseIf TmpRs3!ConditionID = 2 Then
       MSHFlexGrid1.TextMatrix(I, 3) = "Good"
    End If
    MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs3!WHRDate), "", TmpRs3!WHRDate)
    MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs3!StorageFrom), "", TmpRs3!StorageFrom)
    MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs3!ClientName), "", TmpRs3!ClientName) 'GetClientName(TmpRs3!ClientIDRow)
    MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs3!Commodity), "", TmpRs3!Commodity) 'GetCommodityName(TmpRs3!CommodityID)
    MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs3!Grade), "", TmpRs3!Grade)
    MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs3!G_UID), "", TmpRs3!G_UID)
    MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs3!WHRNoOfBags), "", TmpRs3!WHRNoOfBags) 'WHRNoOfBags
    MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs3!WHRQuantity), "", TmpRs3!WHRQuantity)
    MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs3!Flag), "", TmpRs3!Flag)
    MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs3!DeliveredNoOfBags), "", TmpRs3!DeliveredNoOfBags) ' "Delivered Bags"
    MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs3!DeliveredQuantity), "", TmpRs3!DeliveredQuantity) ' "Delivered Quantity"
    
    MSHFlexGrid1.TextMatrix(I, 15) = IIf(IsNull(TmpRs3!CARNo), "", TmpRs3!CARNo) ' "Delivered Quantity"
    MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs3!ConvertedClient), 0, TmpRs3!ConvertedClient) ' "Delivered Quantity"
    MSHFlexGrid1.TextMatrix(I, 17) = IIf(IsNull(TmpRs3!SCMFlag), 0, TmpRs3!SCMFlag) ' "Delivered Quantity"
    
    MSHFlexGrid1.TextMatrix(I, 18) = IIf(IsNull(TmpRs3!Rate), 0, TmpRs3!Rate) ' "Delivered Quantity"
    MSHFlexGrid1.TextMatrix(I, 19) = IIf(IsNull(TmpRs3!ValueofCommodity), 0, TmpRs3!ValueofCommodity) ' "Delivered Quantity"
    MSHFlexGrid1.TextMatrix(I, 20) = IIf(IsNull(TmpRs3!Txn_PledgeCreation_ID), 0, TmpRs3!Txn_PledgeCreation_ID)
    MSHFlexGrid1.TextMatrix(I, 21) = IIf(IsNull(TmpRs3!PledgeDate), "", Format(TmpRs3!PledgeDate, "dd-mmm-yyyy"))
    MSHFlexGrid1.TextMatrix(I, 22) = IIf(IsNull(TmpRs3!BankName), "", TmpRs3!BankName)
    MSHFlexGrid1.TextMatrix(I, 23) = IIf(IsNull(TmpRs3!BranchName), "", TmpRs3!BranchName)
    MSHFlexGrid1.TextMatrix(I, 24) = IIf(IsNull(TmpRs3!GodownNo), "", TmpRs3!GodownNo)
    MSHFlexGrid1.TextMatrix(I, 25) = IIf(IsNull(TmpRs3!LotWiseWHR), "", TmpRs3!LotWiseWHR)
    MSHFlexGrid1.TextMatrix(I, 26) = IIf(IsNull(TmpRs3!PolicyNo), "", TmpRs3!PolicyNo) & "~" & IIf(IsNull(TmpRs3!CoverNoteNo), "", TmpRs3!CoverNoteNo)
    MSHFlexGrid1.TextMatrix(I, 27) = IIf(IsNull(TmpRs3!Risks), "", TmpRs3!Risks)
    MSHFlexGrid1.TextMatrix(I, 28) = IIf(IsNull(TmpRs3!Amount), "", TmpRs3!Amount)
    MSHFlexGrid1.TextMatrix(I, 29) = IIf(IsNull(TmpRs3!ValidFrom), "", Format(TmpRs3!ValidFrom, "dd-mmm-yyyy"))
    MSHFlexGrid1.TextMatrix(I, 30) = IIf(IsNull(TmpRs3!ValidTo), "", Format(TmpRs3!ValidTo, "dd-mmm-yyyy"))
    MSHFlexGrid1.TextMatrix(I, 31) = IIf(IsNull(TmpRs3!InsuranceCoName), "", TmpRs3!InsuranceCoName)
    
    If MSHFlexGrid1.TextMatrix(I, 16) = 1 Then
       MSHFlexGrid1.TextMatrix(I, 16) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(I, 16) = "No"
    End If
    
    If MSHFlexGrid1.TextMatrix(I, 17) = 1 Then
       MSHFlexGrid1.TextMatrix(I, 17) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(I, 17) = "No"
    End If
    If MSHFlexGrid1.TextMatrix(I, 25) = 1 Then
       MSHFlexGrid1.TextMatrix(I, 25) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(I, 25) = "No"
    End If
  
  TmpRs3.MoveNext
Next

Wc1 = "WHRDate >= '" & Format(DTPWHRFromDate.Value, Gen_DatFormat) & "' And WHRDate <= '" & Format(DTPWHRToDate.Value, Gen_DatFormat) & "'"

If CmbPreFix.Text <> "" Then
   Wc1 = Wc1 & " And SM_Prefix='" & mSPrefix & "'"
End If
If CmbSCmp.Text <> "" Then
   Wc1 = Wc1 & " And CMPID=" & mSCMPID & ""
End If
If CmbSClientName.Text <> "" Then
   Wc1 = Wc1 & " And ClientID=" & mSClientID & ""
End If

If CmbSCommodity.Text <> "" Then
   Wc1 = Wc1 & " And CommodityID=" & mSCommodityID & ""
End If

If TxtSWHRNo.Text <> "" Then
   Wc1 = Wc1 & " And WHRNo=" & mSWHRNo & ""
End If



RsTxn_WHRDetail.Filter = Wc1

LblRecordCount.Refresh
LblRecordCount.Caption = TmpRs3.RecordCount
End Sub




Private Sub SearchCmd_Click0ld()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_WHRDetail.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Frame1.Height
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Frame1.Width
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsTxn_WHRDetail.MoveFirst
Call Grid_Head1
MSHFlexGrid1.Rows = RsTxn_WHRDetail.RecordCount + 1
For I = 1 To RsTxn_WHRDetail.RecordCount
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(I, C) = IIf(IsNull(RsTxn_WHRDetail.Fields(C)), "", RsTxn_WHRDetail.Fields(C))
    Next
    If RsTxn_WHRDetail!ConditionID = 0 Then
       MSHFlexGrid1.TextMatrix(I, 2) = "Average"
    ElseIf RsTxn_WHRDetail!ConditionID = 1 Then
       MSHFlexGrid1.TextMatrix(I, 2) = "Fair"
    ElseIf RsTxn_WHRDetail!ConditionID = 2 Then
       MSHFlexGrid1.TextMatrix(I, 2) = "Good"
    End If
    
    MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(RsTxn_WHRDetail!CMPID), 0, RsTxn_WHRDetail!CMPID) & "-" & RsTxn_WHRDetail!WHRNo - (IIf(IsNull(RsTxn_WHRDetail!CMPID), 0, RsTxn_WHRDetail!CMPID) * 1000000)
    MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(RsTxn_WHRDetail!StorageFrom), "", RsTxn_WHRDetail!StorageFrom)
    MSHFlexGrid1.TextMatrix(I, 5) = GetClientName(RsTxn_WHRDetail!ClientID)
    MSHFlexGrid1.TextMatrix(I, 6) = GetCommodityName(RsTxn_WHRDetail!CommodityID)
    MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(RsTxn_WHRDetail!Grade), "", RsTxn_WHRDetail!Grade)
    MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(RsTxn_WHRDetail!G_UID), "", RsTxn_WHRDetail!G_UID)
    MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(RsTxn_WHRDetail!WHRNoOfBags), "", RsTxn_WHRDetail!WHRNoOfBags) 'WHRNoOfBags
    MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(RsTxn_WHRDetail!WHRQuantity), "", RsTxn_WHRDetail!WHRQuantity)
    MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(RsTxn_WHRDetail!Flag), "", RsTxn_WHRDetail!Flag)

    
  RsTxn_WHRDetail.MoveNext
Next


End Sub

Private Sub CmdExcel_Click()
Dim Gen_mTimeStamp As String

'Call ExportToCSV

Gen_mTimeStamp = GetTimeStamp & Gen_UserName & "Txn_WHRDetail" & Trim(TxtWHRNoNew) & ".xls"
Call Gen_Create_Xls(Gen_mTimeStamp)
Call Xls_Detail_Rpt(Gen_mTimeStamp)

End Sub
Public Sub Xls_Detail_Rpt(Fnam_ As String)
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Fnam_)
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
'              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
              If C <> 24 Then
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
              Else
                 oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
              End If
           End If
        Else
           If C <> 24 Then
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           Else
              oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
           End If
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
Private Sub WHRDate_LostFocus()
Call GetGridData(mExchangeTypeID, mCommodityID, mClientIDRow)
End Sub
Private Sub TxtHologramNo_LostFocus()
'If TxtHologramNo = "" Then Exit Sub
'If Check_HologramNo(TxtHologramNo) = False Then
'   TxtHologramNo.SetFocus
'   CmdSaveHologram.Enabled = False
'   Exit Sub
'End If
'CmdSaveHologram.Enabled = True
'CmdSaveHologram.SetFocus


If TxtHologramNo = "" Then Exit Sub
'If Check_HologramNo(TxtHologramNo) = False Then
'   TxtHologramNo.SetFocus
'   CmdSaveHologram.Enabled = False
'   Exit Sub
'End If


If TxtHologramNo <> "0" Then

   If Check_HologramNo(TxtHologramNo) = False Then
      TxtHologramNo.SetFocus
      CmdSaveHologram.Enabled = False
      Exit Sub
   End If
    
   If CheckHologramInRange(TxtHologramNo, mLocationID) = False Then
      MsgBox "Hologram No not found in any range."
      TxtHologramNo.SetFocus
      CmdSaveHologram.Enabled = False
      Exit Sub
   End If
End If

CmdSaveHologram.Enabled = True
CmdSaveHologram.SetFocus


End Sub
Private Sub CmdSaveHologram_Click()
On Error GoTo msgError
'Dim ans As Variant
'If TxtHologramNo = "" Then
'   MsgBox "Blank Hologram No not allowed!!!"
'   TxtHologramNo.SetFocus
'   Exit Sub
'End If
'ans = MsgBox("Are you Sure you want to save this Hologram No?", vbYesNo)
'If ans = vbYes Then
'   Call TableMst_Hologram
'   RsMst_Hologram.AddNew
'   RsMst_Hologram!HologramNo = TxtHologramNo
'   RsMst_Hologram!DocType = "WHR"
'   RsMst_Hologram!DocNo = TxtWHRNo
'   RsMst_Hologram!SM_Prefix = RsTxn_WHRDetail!SM_Prefix
'   RsMst_Hologram.Update
'   RsTxn_WHRDetail!HologramNo = TxtHologramNo
'   RsTxn_WHRDetail.Update
'   TxtHologramNo = ""
'   Frame3.Visible = False
'   RsTxn_WHRDetail!Flag = "P"
'   RsTxn_WHRDetail.Update
'   Gen_mTimeStamp = GetTimeStamp
'   If Check2.Value = 1 Then
'      Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_WHR" & Trim(TxtWHRNoNew) & ".xls")
'      'Call Xls_Print_Rpt
'      Call Xls_Print_Rpt_New
'   End If
'Else
'   TxtHologramNo.SetFocus
'End If

Dim ans As Variant
If TxtHologramNo = "" Then
   MsgBox "Blank Hologram No not allowed!!!"
   TxtHologramNo.SetFocus
   Exit Sub
End If
ans = MsgBox("Are you Sure you want to save this Hologram No?", vbYesNo)
If ans = vbYes Then
   If TxtHologramNo <> "0" Then
      Call TableMst_Hologram
      RsMst_Hologram.AddNew
      RsMst_Hologram!NewHologramNo = TxtHologramNo
      RsMst_Hologram!DocType = "WHR"
      RsMst_Hologram!DocNo = TxtWHRNo
      RsMst_Hologram!SM_Prefix = RsTxn_WHRDetail!SM_Prefix
      RsMst_Hologram!CreatedBy = Gen_UserLoginID
      RsMst_Hologram.Update
   End If
   
   Call TableLog_WHRPrint
   RsTxn_WHRDetail!HologramNo = TxtHologramNo
   RsTxn_WHRDetail.Update
   
   Frame3.Visible = False
   
   RsTxn_WHRDetail!Flag = mPrintFlag
   RsTxn_WHRDetail.Update
   
   RsLog_WHRPrint.AddNew
   RsLog_WHRPrint!WHRNo = RsTxn_WHRDetail!WHRNo
   RsLog_WHRPrint!SM_Prefix = RsTxn_WHRDetail!SM_Prefix
   RsLog_WHRPrint!PrintedDate = Date
   RsLog_WHRPrint!HologramNo = TxtHologramNo
   RsLog_WHRPrint!Version = mWHRVersion
   If mWHRVersion = "Pledgable" Then
      RsLog_WHRPrint!Txn_PledgeCreation_ID = IIf(IsNull(RsTxn_WHRDetail!Txn_PledgeCreation_ID), 0, RsTxn_WHRDetail!Txn_PledgeCreation_ID)
   End If
   RsLog_WHRPrint!UpdatedDate = Now
   RsLog_WHRPrint!UpdatedBy = Gen_UserLoginID
   RsLog_WHRPrint.Update
   TxtHologramNo = ""
   Gen_mTimeStamp = GetTimeStamp
   If Check2.Value = 1 Then
      'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_WHR" & Trim(TxtWHRNoNew) & ".xls")
      'Call Xls_Print_Rpt
      'Call Xls_Print_Rpt_New
      If Option2.Value = True Then
         Call PDF_Print_RptLetterHead
      Else
         Call PDF_Print_Rpt
      End If
      Call Gen_Email
   End If
   
   tmp = "Select LicenseID From Mst_Godown Where GodownID = " & mGodownID & ""
   Call TmpTable
   
   If Not TmpRs.EOF Then
      mNewLicenseID = TmpRs!LicenseID
   End If
   
   tmp = "Update Txn_WHR_Detail set LicenseID = " & mNewLicenseID & " Where LicenseID = " & mLicenseID & " And GodownID = " & mGodownID & " "
   Call TmpTable
   
   MsgBox "The WHR in PDF format is sent to your mailbox. You may download the PDF file and print"
Else
   TxtHologramNo.SetFocus
End If
Exit Sub

msgError:

MsgBox Err.Description, vbExclamation, "Error" '"Some error occurred during process!!!"

End Sub

Private Sub CmbInsuranceCo_GotFocus()
tmp = CmbInsuranceCo.Text
Call TableMst_InsuranceCo(" Where Flag='A'")

CmbInsuranceCo.Clear
If RsMst_InsuranceCo.RecordCount = 0 Then
   MsgBox "No Insurance Co found!!!"
   'CmdInsuranceCo.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_InsuranceCo.RecordCount
    CmbInsuranceCo.AddItem RsMst_InsuranceCo!InsuranceCoName
    RsMst_InsuranceCo.MoveNext
Next
CmbInsuranceCo.Text = tmp
   
End Sub
Private Sub CmbInsuranceCo_LostFocus()

If CmbInsuranceCo.Text = "" Then
   Exit Sub
End If
RsMst_InsuranceCo.MoveFirst
RsMst_InsuranceCo.Find "InsuranceCoName = '" & CmbInsuranceCo.Text & "'"

If RsMst_InsuranceCo.EOF Then
   MsgBox "Invalid Insurance Co selection!!! "
   CmbInsuranceCo.SetFocus
   Exit Sub
End If
mInsuranceCoID = RsMst_InsuranceCo.Fields(0)
If CmbInsuranceCo.Text <> tmp Then
   CmbPolicyNo.Text = ""
   TxtRisks = ""
   TxtAmount = ""
   ValidFrom.Value = Date
   ValidTo.Value = Date
End If

End Sub

Private Sub CmbPolicyNo_GotFocus()
If CmbInsuranceCo.Text = "" Or mInsuranceCoID = "" Then Exit Sub

tmp = CmbPolicyNo.Text
Call TableMst_PolicyDetail(" Where InsuranceCoID =  " & mInsuranceCoID & " And Flag='A' And ValidTo >= '" & Date & "' Order by  PolicyNo,CoverNoteNo")

CmbPolicyNo.Clear
If RsMst_PolicyDetail.RecordCount = 0 Then
   MsgBox "No Insurance Detail found!!!"
   CmbPolicyNo.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_PolicyDetail.RecordCount
    CmbPolicyNo.AddItem RsMst_PolicyDetail!PolicyNo & " ~ " & RsMst_PolicyDetail!CoverNoteNo
    RsMst_PolicyDetail.MoveNext
Next
CmbPolicyNo.Text = tmp
End Sub

Private Sub CmbPolicyNo_LostFocus()
Dim strSplitString() As String
If CmbPolicyNo.Text = "" Then
   Exit Sub
End If
strSplitString = Split(CmbPolicyNo.Text, "~")
'On Error GoTo M_lbl:
tmp = UBound(strSplitString)
If tmp <> 1 Then
   MsgBox "Invalid selection!!!"
   CmbPolicyNo.SetFocus
   Exit Sub
End If
PolicyNo_ = Trim(strSplitString(0))
CNo = Trim(strSplitString(1))
RsMst_PolicyDetail.MoveFirst
For I = 1 To RsMst_PolicyDetail.RecordCount
  If PolicyNo_ = Trim(RsMst_PolicyDetail!PolicyNo) And CNo = Trim(RsMst_PolicyDetail!CoverNoteNo) Then
     TxtRisks = RsMst_PolicyDetail!Risks
     TxtAmount = RsMst_PolicyDetail!Amount
     ValidFrom.Value = RsMst_PolicyDetail!ValidFrom
     ValidTo.Value = RsMst_PolicyDetail!ValidTo
     Exit Sub
  End If
  RsMst_PolicyDetail.MoveNext
Next
If RsMst_PolicyDetail.EOF Then
  MsgBox "Invalid Insurance Co selection!!! "
  CmbPolicyNo.SetFocus
  Exit Sub
End If
mInsuranceCoID = RsMst_PolicyDetail.Fields(0)

End Sub

Private Sub Grid_Head2()
MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.FixedCols = 0
MSHFlexGrid3.Cols = 6

MSHFlexGrid3.TextMatrix(1, 0) = "Location"
MSHFlexGrid3.TextMatrix(1, 1) = "CITF No"
MSHFlexGrid3.TextMatrix(1, 2) = "CITF Date"
MSHFlexGrid3.TextMatrix(1, 3) = "CITF Qty"
MSHFlexGrid3.TextMatrix(1, 4) = "CITF No of Bags"
MSHFlexGrid3.TextMatrix(1, 5) = "Created Date"

MSHFlexGrid3.ColWidth(0) = 2400
MSHFlexGrid3.ColWidth(1) = 2400
MSHFlexGrid3.ColWidth(2) = 2400
MSHFlexGrid3.ColWidth(3) = 2400
MSHFlexGrid3.ColWidth(4) = 2000
MSHFlexGrid3.ColWidth(5) = 2400


End Sub

Private Sub Indata_CITF()
Call Grid_Head2

tmp = "Select ML.LocationName,a.CxTF_DetailsID,b.CxTFDate,a.AdjRRNQty,a.AdjRRNNoofBags,b.CreatedDate " & _
      " From Txn_CxTF_ClientDetail a " & _
      " Inner Join Txn_CxTF_Details b on a.SM_Prefix = b.SM_Prefix And a.CxTF_DetailsID = b.CxTF_DetailsID  And a.txntype = b.txntype  " & _
      " Inner Join Mst_CMP MC On b.CMPID = MC.CMPID  " & _
      " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID  " & _
      " Where b.BookedSpaceFlag <> 1 And b.exchangetypeid= 1 and b.txntype = 'W'  " & _
      " And a.RRNID =  " & TxtWHRNo & " "

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs.RecordCount + 3
   mRow = 2
   For x = 1 To TmpRs.RecordCount
       For C = 0 To MSHFlexGrid3.Cols - 1
           MSHFlexGrid3.TextMatrix(mRow, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       MSHFlexGrid3.TextMatrix(mRow, 2) = Format(MSHFlexGrid3.TextMatrix(mRow, 2), "dd-mmm-yyyy")
       MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + Val(MSHFlexGrid3.TextMatrix(mRow, 3))
       MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(mRow, 4))
       MSHFlexGrid3.TextMatrix(mRow, 5) = Format(MSHFlexGrid3.TextMatrix(mRow, 5), "dd-mmm-yyyy hh:mm AM/PM")
       mRow = mRow + 1
       TmpRs.MoveNext
   Next
End If

End Sub

Private Sub DrawTable(Optional mROw_ As Double, Optional mCol_ As Double, Optional mTop_ As Double, Optional mHeight_ As Double, Optional mLeft_ As Double, Optional mRowHeight_ As Double)

If mROw_ = 0 Then mROw_ = 1
If mCol_ = 0 Then mCol_ = 1
If mLeft_ = 0 Then mLeft_ = mmasterLeft
If mTop_ = 0 Then mTop_ = 0
If mHeight_ = 0 Then mHeight_ = mTop_ + mRowHeight_

Dim pgval_ As Double

For ir = 1 To mROw_
    pgval_ = pgval
    pgval_ = pgval_ / mCol_
    mWidth_ = pgval_
    For ic = 1 To mCol_
        objPDF.PDFDrawPolygon Array(mLeft_, mTop_, mWidth_, mTop_, mWidth_, mHeight_, mLeft_, mHeight_)
        mLeft_ = pgval_ * ic
        mWidth_ = pgval_ * (ic + 1)
    Next
    mTop_ = mHeight_
    mHeight_ = mHeight_ + mRowHeight_
    'mLeft_ = mmasterLeft

    mLeft_ = mGlobalLeft

Next

mTop = mTop_
mRow = mROw_ + 1

End Sub
Private Sub PrintText(mFontSize_ As Integer, mFontFormat_ As PDFFontStl, mAlign_ As Variant, mTxtLeft_ As Double, mTxtTop_ As Double, mTexttoPrint As String)

'        objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
'        objPDF.PDFSetAlignement = ALIGN_CENTER
        
objPDF.PDFSetFont FONT_ARIAL, mFontSize_, mFontFormat_
objPDF.PDFSetAlignement = mAlign_
objPDF.PDFTextOut mTexttoPrint, mTxtLeft_, mTxtTop_


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
mMsg = mMsg & "<tr><td>WHR No</td><td>" & TxtWHRNoNew.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>WHR Date</td><td>" & Format(WHRDate.Value, "dd-MMM-yyyy") & "</td></tr>"
mMsg = mMsg & "<tr><td>Client Name</td><td>" & TxtClientID.Text & "</td> </tr>"
mMsg = mMsg & "<tr><td>Commodity</td><td>" & TxtWHRCommodityId.Text & "</td></tr>"
mMsg = mMsg & "<tr><td>Godown No</td><td>" & TxtCMPName & "</td></tr>"
mMsg = mMsg & "<tr><td>Godown No</td><td>" & GetGodownName(RsTxn_WHRDetail!GodownID, "N") & "</td></tr>"
mMsg = mMsg & "<tr><td>Godown Address</td><td>" & GetGodownName(RsTxn_WHRDetail!GodownID, "A") & "</td></tr>"
If mLotWiseWHR = 0 Then
   mMsg = mMsg & "<tr><td>No of Bags</td><td>" & Format(MSHFlexGrid2.TextMatrix(2, 7), "#######0") & "</td></tr>"
   mMsg = mMsg & "<tr><td>Net Qty (MT)</td><td>" & Format(MSHFlexGrid2.TextMatrix(2, 8), "#######0.000") & "</td></tr>"
Else
   mMsg = mMsg & "<tr><td>No of Bags</td><td>" & Format(MSHFlexGrid4.TextMatrix(0, 5), "#######0") & "</td></tr>"
   mMsg = mMsg & "<tr><td>Net Qty (MT)</td><td>" & Format(MSHFlexGrid4.TextMatrix(0, 6), "#######0.000") & "</td></tr>"
End If
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

mMsgSubject = "NCMSL " & TxtCMPName & " WHR Details"

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

objSMTP.AddAttachment mPDFFileName
objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If

End Function

Private Sub PDF_Print_Rpt()
Dim mHeadingTop  As Double
Dim mHeadingLeft  As Double
Dim ans As Variant
Set objPDF = New mjwPDF

I = 2
' Create a simple PDF file using the mjwPDF class

Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
    
    
' Set the PDF title and filename
objPDF.PDFTitle = "WHR PDF Document"
objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_WHR-" & Trim(TxtWHRNoNew) & ".pdf"
mPDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_WHR-" & Trim(TxtWHRNoNew) & ".pdf"
' We must tell the class where the PDF fonts are located
objPDF.PDFLoadAfm = App.Path & "\Fonts"

' Set the file properties
objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
objPDF.PDFFormatPage = FORMAT_A4
objPDF.PDFOrientation = ORIENT_PORTRAIT
objPDF.PDFSetUnit = UNIT_PT
objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU


' Lets us set see the bookmark pane when we view the PDF
'objPDF.PDFUseOutlines = True

' View the PDF file after we create it
objPDF.PDFView = True
'    objPDF.PDFWatermark = "NCMSL"
' Begin our PDF document

objPDF.PDFBeginDoc
    ' Lets add a heading
objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True

objPDF.PDFCell " ", 15, 15, _
objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20

If PntPrv = "Y" Then
   objPDF.PDFWatermark = "Preview"
Else
   objPDF.PDFWatermark = "NCML" '"NCMSL"
End If

objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True
   
pgval = ((objPDF.PDFGetPageWidth - 18))
  
mRowHeight = 15
mGlobalLeft = 12
'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
mLeft = mGlobalLeft
mTop = 15
mHeight = mTop + mRowHeight
mTxtFromTop = mRowHeight + 2
 
 'Call DrawTable(10, 2, 15, 30, 12)

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len(GetCMPAddress(RsTxn_WHRDetail!CMPID, "D"))
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 12

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, GetCMPAddress(RsTxn_WHRDetail!CMPID, "D"))

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Warehouse Receipt")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 10

'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mTxtLeft - 26, mTxtTop - 3, "Non Negotiable Warehouse Receipt")
mHeadingTop = mTxtTop - 3
mHeadingLeft = mTxtLeft - 26

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Depositor Details")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 7
Dim mDispTextSize  As Integer
mDispTextSize = 7

mTxtTop = mTxtTop + mRowHeight

Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Receipt No")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, TxtWHRNoNew)
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Receipt Date")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, Format(WHRDate, "dd-MMM-yyyy"))
'--
mTxtTop = mTxtTop + mRowHeight

Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)

If mLotWiseWHR = 0 Then
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Warehouse License No")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, MSHFlexGrid2.TextMatrix(2, 3))
'--
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "License Expiry Date")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, Format(RsMst_License!LicenseExpiryDate, "dd-MMM-yyyy"))
'--
Else
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Warehouse License No")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, CStr(mLicenseNo))
   '--
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "License Expiry Date")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, Format(CDate(mLicenseExpiryDate), "dd-MMM-yyyy"))
End If

mTxtTop = mTxtTop + mRowHeight

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Location & Address of Warehouse")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, GetLocationFromCMP(RsMst_Godown!CMPID))
'--

objPDF.PDFDrawLine mGlobalLeft + 132.5, mTop - mRowHeight, mGlobalLeft + 132.5, mTop 'mRowHeight  'mRowHeight 'mHeight - 5 'mRowHeight + 100

mTxtTop = mTxtTop + mRowHeight

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Godown No")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, GetGodownName(RsTxn_WHRDetail!GodownID, "N"))

objPDF.PDFDrawLine mGlobalLeft + 132.5, mTop - mRowHeight, mGlobalLeft + 132.5, mTop

mTxtTop = mTxtTop + mRowHeight
'    Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Godown Address")

Dim mGodownAddr As String
mGodownAddr = GetGodownName(RsTxn_WHRDetail!GodownID, "A")
Dim mLen As Integer, mMax As Integer, mPara As Integer
Dim t1 As Double
mMax = 120
mPara = 0
t1 = 1
mLen = Len(mGodownAddr)
mLen = RoundUp(mLen / mMax) '+ 1
For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       'mMax '= mMax + 35
       mPara = mPara + 7
    End If
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara - 5, mID$(mGodownAddr, t1, mMax))
Next

mTxtTop = mTxtTop + (mRowHeight)

objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 132.5, mTop - (mRowHeight * (T - 1)), mGlobalLeft + 132.5, mTxtTop 'mRowHeight

mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Name and Address of the Depositor")
    
Dim mDepName As String
ans = GetClientNameByRow(mClientIDRow)

If ans(4) <> "" Then
   mDepName = TxtClientID & "-" & Trim(ans(4))
Else
   mDepName = TxtClientID
End If

mMax = 145
mPara = 0
t1 = 1
mLen = Len(mDepName)
mLen = RoundUp(mLen / mMax) '+ 1
For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       'mMax '= mMax + 35
       mPara = mPara + 7
    End If
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara - 5, mID$(mDepName, t1, mMax))
Next

mTxtTop = mTxtTop + (mRowHeight)
 
objPDF.PDFDrawLine mGlobalLeft + 132.5, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 132.5, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)

'Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)

If ChkSaidToContain.Value = 1 Then
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Commodity (Said to contain)")
Else
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Commodity")
End If
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, TxtWHRCommodityId.Text)
 
 
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Condition of Commodity")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, CmbCommodityCondition)

objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)


Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "No of Bags")
If mLotWiseWHR = 0 Then
   Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, "[" & Format(MSHFlexGrid2.TextMatrix(I, 7), "#######0") & "]")
Else
   Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, "[" & Format(MSHFlexGrid4.TextMatrix(0, 5), "#######0") & "]")
End If
If ChkEstimatedWgt.Value = 1 Then
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Gross Quantity(MT) (Said to Contain)")
Else
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Gross Quantity(MT)")
End If
If mLotWiseWHR = 0 Then
   Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "[" & Format(MSHFlexGrid2.TextMatrix(I, 8), "#######0.000") & "]")
Else
   Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "[" & Format(MSHFlexGrid4.TextMatrix(0, 6), "#######0.000") & "]")
End If
objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Name and License No of Weigher")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, IIf(TxtWeigher = "", "---", TxtWeigher))
 
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Private Marks on Packaging")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, TxtGrade)
 
objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Name & License No of Assayer")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, IIf(TxtAssayer = "", "---", TxtAssayer))

objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   
mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Quality As per Grader")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, IIf(TxtAssayerCerti = "", "---", TxtAssayerCerti))

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Grade Valid Till")

If TxtAssayer <> "" Then
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, Format(GradeValidUpto, "dd-mmm-yyyy"))
Else
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "---")
End If

objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)
If OptSRPerBag.Value = True Then
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Storage Charge(Rate)-excluding Taxes")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10 + mRowHeight, "(Rupees per bag/month)")
Else
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Storage Charge(Rate)-excluding Taxes ")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10 + mRowHeight, "(Rupees per MT/month)")
End If
Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, Format(TxtStorageCharges, "#########0.00"))
     
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Storage Charge - excluding Taxes ")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 10 + mRowHeight, "(Total Amount) per month")

Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "[" & Format(TxtStorageAmount, "###,###,###,###,##0") & "]")
 
objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop + mRowHeight - 5
objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop + mRowHeight - 5
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop + mRowHeight - 5
objPDF.PDFDrawLine mGlobalLeft, mTxtTop + mRowHeight - 5, objPDF.PDFGetPageWidth - 18, mTxtTop + mRowHeight - 5

mTxtTop = mTxtTop + (mRowHeight) + (mRowHeight) - 5

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Goods are Accepted for storage From")

Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, Format(StorageFrom, "dd-mmm-yyyy"))

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Goods are Accepted for storage UpTo")
 
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, Format(StorageTo, "dd-mmm-yyyy"))
 
objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Market Rate at the time of deposit (Rs/MT)")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10 + mRowHeight, "(Rs/MT)")

Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, "[" & Format(TxtRate, "###,###,###,##0") & "]")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Value of Commodities at time of deposit(Amount in Rs) ")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 10 + mRowHeight, "")
 
Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "[" & Format(TxtValueofCommodity, "###,###,###,##0") & "]")
 
objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Gross Quantity in Words")

Dim mqtyStr, str1, str2, str3 As String
str1 = ""
str2 = ""
str3 = ""



If mLotWiseWHR = 0 Then
   mqtyStr = Format(MSHFlexGrid2.TextMatrix(I, 8), "#####0.000") 'Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "[" & Format(MSHFlexGrid2.TextMatrix(I, 8), "#######0.000") & "]")
Else
   mqtyStr = Format(MSHFlexGrid4.TextMatrix(0, 6), "#####0.000") 'Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "[" & Format(MSHFlexGrid4.TextMatrix(0, 6), "#######0.000") & "]")
End If

str1 = NumStrConv(mID(mqtyStr, 1, InStr(mqtyStr, ".") - 1))
str2 = NumStrConv(mID(mqtyStr, InStr(mqtyStr, ".") + 1, 3))
str1 = Replace(str1, "Only", "MT")
str2 = Replace(str2, "Only", "KG")
If str2 = "" Then
   str3 = str1 & " Only"
ElseIf str1 = "" Then
   str3 = str2 & " Only"
Else
   str3 = str1 & " And " & str2 & " Only"
End If

Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, str3)

objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   
mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Value of Commodity in Words")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, NumStrConv(Val(Format(TxtValueofCommodity.Text, "######0"))))

objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

''mTxtTop = mTxtTop + (mRowHeight) + (mRowHeight)
''
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop + 10, "Signature & Stamp of Warehouse")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop + mRowHeight + 5, "Manager")
''
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, "")
''
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop + mRowHeight + 5, "Hologram No:- " & RsTxn_WHRDetail!HologramNo)
''
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop + mRowHeight + 5, "Affix Hologram Here")

''objPDF.PDFDrawLine mGlobalLeft + 132.5, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 132.5, mTxtTop + (mRowHeight) + 8
''
''objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 276, mTxtTop + (mRowHeight) + 8
''objPDF.PDFDrawLine mGlobalLeft + 422, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 422, mTxtTop + (mRowHeight) + 8
''objPDF.PDFDrawLine mGlobalLeft, mTxtTop + (mRowHeight) + 8, objPDF.PDFGetPageWidth - 18, mTxtTop + (mRowHeight) + 8
''
''objPDF.PDFDrawLine mGlobalLeft, mTxtTop + (mRowHeight) + 10, objPDF.PDFGetPageWidth - 18, mTxtTop + (mRowHeight) + 10

''mTxtTop = mTxtTop + (mRowHeight) + (mRowHeight) + 10

mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Insurance Details")
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Valid From - Valid to")
' Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5 + mRowHeight, "Manager")
    
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "Policy & Cover Note No")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop - 5, "Insurance Co.")
 
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, "Policy Amount (in Rs.)")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, "Risk Covered")
 
objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
 
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)
 
If ValidFrom.Value <> "1/1/1900" Or ValidTo <> "1/1/1900" Then
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, Format(ValidFrom, "DD-MMM-YYYY") & " To ")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10 + mRowHeight, Format(ValidTo, "DD-MMM-YYYY"))
End If
Dim a1, a2, a3 As Variant
mMax = 26
mPara = 0
t1 = 1
mLen = Len(TxtPolicyNo)
mLen = RoundUp(mLen / mMax) '+ 1
For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       mPara = mPara + 7
    End If
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop + mPara - 5, mID$(TxtPolicyNo, t1, mMax))
Next
a1 = mTxtTop + mPara - 5
mMax = 35
mPara = 0
t1 = 1
mLen = Len(TxtInsuranceCo)
mLen = RoundUp(mLen / mMax) '+ 1
For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       mPara = mPara + 7
     End If
     Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop + mPara - 5, mID$(TxtInsuranceCo, t1, mMax))
Next
'a2 = mTxtTop + mPara - 5
If TxtAmount <> "0" Then
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, TxtAmount)
End If

mMax = 50
mPara = 0
t1 = 1
mLen = Len(TxtRisks)
mLen = RoundUp(mLen / mMax) '+ 1
For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       mPara = mPara + 7
    End If
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop + mPara - 5, mID$(TxtRisks, t1, mMax))
Next

a3 = mTxtTop + mPara - 5
 
If a1 > a3 Then
   objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, Val(a1) + 10 'mTxtTop + (mRowHeight) + a1

   objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, Val(a1) + 10 'mTxtTop + (mRowHeight) + a1
   objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, Val(a1) + 10 'mTxtTop + (mRowHeight) + a1
   objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, Val(a1) + 10 'mTxtTop + (mRowHeight) + a1
   objPDF.PDFDrawLine mGlobalLeft, Val(a1) + 10, objPDF.PDFGetPageWidth - 18, Val(a1) + 10
   mTxtTop = Val(a1) + 10
Else
   objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, Val(a3) + 10
   objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, Val(a3) + 10
   objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, Val(a3) + 10
   objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, Val(a3) + 10
   objPDF.PDFDrawLine mGlobalLeft, Val(a3) + 10, objPDF.PDFGetPageWidth - 18, Val(a3) + 10
   mTxtTop = Val(a3) + 10
End If

'''tmp = "SELECT  TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity, isnull(sum(ReleasedQuantity),0) ReleasedQuantity " 'sum(ReleasedQuantity)ReleasedQuantity"
'''tmp = tmp & " FROM  Txn_WHR_Detail TWD"
'''tmp = tmp & " inner join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
'''tmp = tmp & " left Join Txn_PledgeRelease TPR on TPR.PledgeID = TPC.PledgeID And TPC.SM_Prefix = TPR.SM_Prefix"
'''tmp = tmp & " left Join Txn_PledgeRelease_GSL TPRG on TPRG.PledgeReleaseID = TPR.PledgeReleaseID And TPRG.SM_Prefix = TPR.SM_Prefix"
'''tmp = tmp & " Where TWD.WHRNo = " & RsTxn_WHRDetail!WHRNO
'''tmp = tmp & " Group By TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity"
'''tmp = tmp & " Having TWD.WHRQuantity > isnull(sum(ReleasedQuantity),0) " 'Sum(ReleasedQuantity)"

tmp = " SELECT  TWD.Txn_PledgeCreation_ID,TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity,isnull(sum(PledgeQuantity),0) PledgeQuantity"
tmp = tmp & " FROM  Txn_WHR_Detail TWD"
tmp = tmp & " left join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
tmp = tmp & " left Join Txn_PledgeCreation_GSL TPRG on TPRG.Txn_PledgeCreation_ID = TPC.Txn_PledgeCreation_ID And TPRG.SM_Prefix = TPC.SM_Prefix"
tmp = tmp & " Where TWD.WHRNo =  " & RsTxn_WHRDetail!WHRNo & " And TPC.Txn_PledgeCreation_ID = " & IIf(IsNull(RsTxn_WHRDetail!Txn_PledgeCreation_ID), 0, RsTxn_WHRDetail!Txn_PledgeCreation_ID)
tmp = tmp & " Group By TWD.Txn_PledgeCreation_ID,TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity"
tmp = tmp & " Having IsNull(Sum(PledgeQuantity), 0) > 0"

Call Tmp4Table

If TmpRs4.RecordCount = 0 Then
   
   'Commented on 12 April 2012 Mail by Mayur Sir for removal of Non Negotiable
   'Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mHeadingLeft, mHeadingTop, "Non Pledgable Non Negotiable Warehouse Receipt")
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mHeadingLeft, mHeadingTop, "            Non Pledgable Warehouse Receipt")
   
   mTxtTop = mTxtTop + (mRowHeight)
   Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "This WHR can not be pledged to Banks/Financial Institutions.") ', subject to confirmation by Warehouse Manager")
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   mTxtTop = mTxtTop + 5
   mTypeFlag = "NP"
Else

'''    mTxtTop = mTxtTop + (mRowHeight)
'''
'''    Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "The Warehouse Receipt is valid for transfer by endorsement, and/or for  marking of plegde.") ', subject to confirmation by Warehouse Manager")
'''    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

   'Commented on 12 April 2012 Mail by Mayur Sir for removal of Non Negotiable
   'Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mHeadingLeft, mHeadingTop, " Non Negotiable Warehouse Receipt")
   mTypeFlag = "P"
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mHeadingLeft, mHeadingTop, "            Warehouse Receipt")
   
   mTxtTop = mTxtTop + (mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Pledge details: Any lien marking of this WHR is invalid without an authorised Pledge number under 'Lien/Pledge number as per StoreMan' column below")
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   mTxtTop = mTxtTop + (mRowHeight)
    
    'mTxtTop = mTxtTop + 20
   Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "This WHR can not be pledged to Banks/Financial Institutions, unless a 'Pledge Confirmation Letter' issued by NCMSL in their letter head is attached alongwith this")
   mTxtTop = mTxtTop + 10
   Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "WHR with complete details of applicable terms & conditions of storage.")
     
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   
   mTxtTop = mTxtTop + mRowHeight
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Date")
   
        
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "Name and Address Of Lender")
    
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, "Pledge Number")
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 10 + mRowHeight, "as per StoreMan")
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, " Pledge Quantity")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 10 + mRowHeight, "(in MT)")
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 5, "Pledge Released on")
            
    'Commented on 07 Jan 2012 1205 PM by Parag giving error while pledge qty is present but not release entry
    
    'tmp = "SELECT  TPC.PledgeDate, convert(varchar,TPC.PledgeID)+'~'+TPC.PledgeReqNo as PledgeID, max(TPCG.PledgeQuantity)+sum(TPRG.ReleasedQuantity)PledgeQuantity,"
    'tmp = tmp & " max(TPR.PledgeReleaseDate)PledgeReleaseDate, MB.BankName, MB.Address"
    'tmp = tmp & " FROM  Txn_WHR_Detail TWD"
    'tmp = tmp & " inner join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
    'tmp = tmp & " inner join Txn_PledgeCreation_GSL TPCG on  TPC.PledgeID = TPCG.PledgeID And TPC.SM_Prefix = TPCG.SM_Prefix"
    'tmp = tmp & " left Join Txn_PledgeRelease TPR on TPR.PledgeID = TPC.PledgeID And TPC.SM_Prefix = TPR.SM_Prefix"
    'tmp = tmp & " left Join Txn_PledgeRelease_GSL TPRG on TPR.PledgeReleaseID = TPRG.PledgeReleaseID And TPR.SM_Prefix = TPRG.SM_Prefix"
    'tmp = tmp & " inner Join Mst_Branch MBr On TPC.BranchID = MBr.BranchID"
    'tmp = tmp & " inner Join Mst_Bank MB On TPC.BankID = MB.BankID"
    'tmp = tmp & " Where TWD.WHRNo = " & RsTxn_WHRDetail!WHRNO
    'tmp = tmp & " Group By TPC.PledgeDate, TPC.PledgeID,TPC.PledgeReqNo, MB.BankName, MB.Address"
    
    tmp = " SELECT  TPC.PledgeDate, convert(varchar,TPC.Txn_PledgeCreation_ID) as Txn_PledgeCreation_ID,"
    tmp = tmp & " Sum(TPCG.PledgeQuantity) PledgeQuantity, "
    tmp = tmp & "  MB.BankName, MBr.Address,TPC.Txn_PledgeCreation_ID as OriPledgeID,TPC.SM_Prefix  "
    tmp = tmp & " FROM  Txn_WHR_Detail TWD"
    tmp = tmp & " inner join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
    tmp = tmp & " inner join Txn_PledgeCreation_GSL TPCG on  TPC.Txn_PledgeCreation_ID = TPCG.Txn_PledgeCreation_ID And TPC.SM_Prefix = TPCG.SM_Prefix"
    tmp = tmp & " inner Join Mst_Branch MBr On TPC.BranchID = MBr.BranchID"
    tmp = tmp & " inner Join Mst_Bank MB On TPC.BankID = MB.BankID"
    tmp = tmp & " Where TWD.WHRNo = " & RsTxn_WHRDetail!WHRNo & " "
    tmp = tmp & " Group By TPC.PledgeDate, TPC.Txn_PledgeCreation_ID,TPC.PledgeReqNo,TPC.SM_Prefix, MB.BankName, MBr.Address"
    
    Call Tmp4Table
    
    'mPQty = TmpRs4!PledgeQuantity
    
    objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop + mRowHeight - 5
    objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop + mRowHeight - 5

    objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop + mRowHeight - 5
    objPDF.PDFDrawLine mGlobalLeft + 478, mTxtTop - mRowHeight, mGlobalLeft + 478, mTxtTop + mRowHeight - 5
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop + mRowHeight - 5, objPDF.PDFGetPageWidth - 18, mTxtTop + mRowHeight - 5
    
    'objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
        
    mTxtTop = mTxtTop + mRowHeight + (mRowHeight) - 5
    
    mPledgeQty = 0
    
    
    For k = 1 To TmpRs4.RecordCount
        mPledgeQty = TmpRs4!PledgeQuantity
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, Format(TmpRs4!PledgeDate, "dd-MMM-yyyy"))
        
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, TmpRs4!BankName)
        
        mMax = 70
        mPara = 7
        t1 = 1
        mLen = Len(TmpRs4!Address)
        
        If mLen > 70 Then
           mLen = RoundUp(mLen / mMax)
        Else
           mLen = 1
        End If
        For T = 1 To mLen
            If T > 1 Then
               t1 = mMax + t1
               mPara = mPara + 7
            End If
            Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop + mPara - 5, mID$(TmpRs4!Address, t1, mMax))
        Next
        
        a1 = mTxtTop + mPara - 5

        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, TmpRs4!Txn_PledgeCreation_ID)
                
        mRPledge = 0
                
        tmp = "Select Sum(isnull(ReleasedQuantity,0)) RQty,Max(PledgeReleaseDate) RDate"
        tmp = tmp & " From Txn_PledgeRelease TPR"
        tmp = tmp & " left Join Txn_PledgeRelease_GSL TPRG on TPR.Txn_PledgeRelease_ID = TPRG.Txn_PledgeRelease_ID And TPR.SM_Prefix = TPRG.SM_Prefix"
        tmp = tmp & " Where TPR.Txn_PledgeCreation_ID = " & TmpRs4!OriPledgeID & " And TPR.SM_Prefix = '" & TmpRs4!SM_Prefix & "'"
                
        Call Tmp5Table
        
        If TmpRs5.RecordCount > 0 Then
           If IsNull(TmpRs5!RQty) = False Then
              mRPledge = TmpRs5!RQty
           End If
           If Val(RsTxn_WHRDetail!WHRQuantity) >= Val(mRPledge) Then
              Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 5, IIf(IsNull(TmpRs5!RDate), "", Format(TmpRs5!RDate, "dd-MMM-yyyy")))
           End If
        Else
           Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 5, "")
        End If
                    
        Dim mTotPledge As String
        mTotPledge = Val(mPledgeQty) + Val(mRPledge)
                
                
        'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, IIf(IsNull(TmpRs4!PledgeQuantity), 0, TmpRs4!PledgeQuantity))
        
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, mTotPledge)
        
        'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 5, IIf(IsNull(TmpRs4!PledgeReleaseDate), "", Format(TmpRs4!PledgeReleaseDate, "dd-MMM-yyyy")))
    
        objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, Val(a1) + 5 'mTxtTop + mRowHeight - 5
        objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, Val(a1) + 5  'mTxtTop + mRowHeight - 5
        objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, Val(a1) + 5 'mTxtTop + mRowHeight - 5
        objPDF.PDFDrawLine mGlobalLeft + 478, mTxtTop - mRowHeight, mGlobalLeft + 478, Val(a1) + 5 'mTxtTop + mRowHeight - 5
        objPDF.PDFDrawLine mGlobalLeft, Val(a1) + 5, objPDF.PDFGetPageWidth - 18, Val(a1) + 5  'mTxtTop + mRowHeight - 5
         
        mTxtTop = Val(a1) + mRowHeight + 5
        
        TmpRs4.MoveNext
        
    Next
    mTxtTop = Val(a1) + mRowHeight - 5
End If

mTxtTop = mTxtTop + (mRowHeight) - 5

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Withdrawal History")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "The good mentioned below are hereby released for delivery from the Warehouse. The unreleased balance is subject to lien for unpaid charges")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 10 + mRowHeight, "and borrowings from lending institutions.")

objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop + mRowHeight, objPDF.PDFGetPageWidth - 18, mTxtTop + mRowHeight

mTxtTop = mTxtTop + (mRowHeight) + (mRowHeight)


Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 240, mTxtTop - 5, "Quantity Released")
 
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, "Quantity Balance")

objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 348, mTxtTop - mRowHeight, mGlobalLeft + 348, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 518, mTxtTop - mRowHeight, mGlobalLeft + 518, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop, mGlobalLeft + 518, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Date")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "CITF No")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop - 5, "No. of Bags")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 270, mTxtTop - 5, "Quantity in MT")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop - 5, "No. Of Bags")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 440, mTxtTop - 5, "Quantity in  MT")

''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 580, mTxtTop - 15, "Signature & Stamp of ")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 5, "Warehouse Manager")
objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 268, mTxtTop - mRowHeight, mGlobalLeft + 268, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 348, mTxtTop - mRowHeight, mGlobalLeft + 348, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 438, mTxtTop - mRowHeight, mGlobalLeft + 438, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 518, mTxtTop - mRowHeight, mGlobalLeft + 518, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, mGlobalLeft + 518, mTxtTop '' objPDF.PDFGetPageWidth - 18, mTxtTop
  
mTxtTop = mTxtTop + (mRowHeight)
  
''tmp = "select CxTFDate,sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
''tmp = tmp & " from Txn_CxTF_GSL b"
''tmp = tmp & " inner join  txn_cxtf_details a on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
''tmp = tmp & " Where WHRNo= " & RsTxn_WHRDetail!WHRNO & "  And a.TxnType='W' And a.SM_Prefix='" & RsTxn_WHRDetail!SM_Prefix & "'"
''tmp = tmp & " Group By CxTFDate"
''
''Call Tmp4Table

Dim mWHRWt As Double
Dim mWHRBag As Double

mWHRWt = RsTxn_WHRDetail!WHRQuantity
mWHRBag = RsTxn_WHRDetail!WHRNoOfBags

If MSHFlexGrid3.Rows > 3 Then
   For k = 2 To MSHFlexGrid3.Rows - 2
       mWHRWt = Format(mWHRWt - Val(MSHFlexGrid3.TextMatrix(k, 3)), "#####0.000")
       mWHRBag = Format(mWHRBag - Val(MSHFlexGrid3.TextMatrix(k, 4)), "#####0")
       
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, Format(MSHFlexGrid3.TextMatrix(k, 2), "dd-MMM-yyyy"))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, MSHFlexGrid3.TextMatrix(k, 1))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop - 5, Format(Val(MSHFlexGrid3.TextMatrix(k, 4)), "#####0"))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTxtTop - 5, Format(Val(MSHFlexGrid3.TextMatrix(k, 3)), "#####0.000"))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop - 5, Str(mWHRBag))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 440, mTxtTop - 5, Str(mWHRWt))
  
       objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 268, mTxtTop - mRowHeight, mGlobalLeft + 268, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 348, mTxtTop - mRowHeight, mGlobalLeft + 348, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 438, mTxtTop - mRowHeight, mGlobalLeft + 438, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 518, mTxtTop - mRowHeight, mGlobalLeft + 518, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft, mTxtTop, mGlobalLeft + 518, mTxtTop
       
        mTxtTop = mTxtTop + (mRowHeight)
    Next
    mTxtTop = mTxtTop - 5
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
Else
    objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 268, mTxtTop - mRowHeight, mGlobalLeft + 268, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 348, mTxtTop - mRowHeight, mGlobalLeft + 348, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 438, mTxtTop - mRowHeight, mGlobalLeft + 438, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 518, mTxtTop - mRowHeight, mGlobalLeft + 518, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, mGlobalLeft + 518, mTxtTop
    
    mTxtTop = mTxtTop + mRowHeight
    
    objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 268, mTxtTop - mRowHeight, mGlobalLeft + 268, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 348, mTxtTop - mRowHeight, mGlobalLeft + 348, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 438, mTxtTop - mRowHeight, mGlobalLeft + 438, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 518, mTxtTop - mRowHeight, mGlobalLeft + 518, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, mGlobalLeft + 518, mTxtTop
    
    mTxtTop = mTxtTop + mRowHeight

    objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 268, mTxtTop - mRowHeight, mGlobalLeft + 268, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 348, mTxtTop - mRowHeight, mGlobalLeft + 348, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 438, mTxtTop - mRowHeight, mGlobalLeft + 438, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 518, mTxtTop - mRowHeight, mGlobalLeft + 518, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
    
    mTxtTop = mTxtTop + 10
    
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

End If

mTxtTop = mTxtTop + (mRowHeight)

'mTxtTop = mTxtTop + (mRowHeight) + 10

'Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "The Warehouse Receipt is valid for transfer by endorsement, and/or for  marking of plegde, subject to confirmation by Warehouse Manager")
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'
'mTxtTop = mTxtTop + (mRowHeight)
If mNCMLFinance = 0 Then
   Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "This Warehouse Receipt is not valid without an official Hologram of NCMSL affixed below.")
End If

objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "If the commodities are not withdrawn by 'Goods are Accepted for storage UpTo' date mentioned in the WHR,  NCMSL shall be entitled to dispose the same at the risk ")
mTxtTop = mTxtTop + 10
Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "and responsibility of the depositor and remit the proceeds to the depositor after deducting the applicable fees and charges.")
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop


mTxtTop = mTxtTop + (mRowHeight) + (mRowHeight)
If mTxtTop > objPDF.PDFGetPageHeight Then
   Call CheckForNewPage
End If
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop + 10, "Signature & Stamp of Warehouse")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop + mRowHeight + 5, "Manager")

Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, "")
If mNCMLFinance = 0 Then
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop + mRowHeight + 5, "Hologram No:- " & RsTxn_WHRDetail!HologramNo)

   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop + mRowHeight + 5, "Affix Hologram Here")
End If

objPDF.PDFDrawLine mGlobalLeft + 132.5, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 132.5, mTxtTop + (mRowHeight) + 8

objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 276, mTxtTop + (mRowHeight) + 8
objPDF.PDFDrawLine mGlobalLeft + 422, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 422, mTxtTop + (mRowHeight) + 8
objPDF.PDFDrawLine mGlobalLeft, mTxtTop + (mRowHeight) + 8, objPDF.PDFGetPageWidth - 18, mTxtTop + (mRowHeight) + 8

objPDF.PDFDrawLine mGlobalLeft, mTxtTop + (mRowHeight) + 10, objPDF.PDFGetPageWidth - 18, mTxtTop + (mRowHeight) + 10

mTxtTop = mTxtTop + (mRowHeight) + 10 + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "WHR Versions :")
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)
If mTxtTop > objPDF.PDFGetPageHeight Then
   Call CheckForNewPage
End If
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Hologram No")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 135, mTxtTop - 5, " Date")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "WHR Type")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 380, mTxtTop - 5, "Pledgee Bank Name")

objPDF.PDFDrawLine mGlobalLeft + 132, mTxtTop - mRowHeight, mGlobalLeft + 132, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 376, mTxtTop - mRowHeight, mGlobalLeft + 376, mTxtTop

objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop '' objPDF.PDFGetPageWidth - 18, mTxtTop
  
mTxtTop = mTxtTop + (mRowHeight)
  
'tmp = "Select * from Log_WHRPrint Where WHRNo= " & RsTxn_WHRDetail!WHRNO & ""

tmp = "Select LW.HologramNo,LW.PrintedDate 'PrintDate',LW.Version,MB.BankName"
tmp = tmp & " From Log_WHRPrint LW"
tmp = tmp & " left join Txn_PledgeCreation TP on TP.Txn_PledgeCreation_ID=LW.Txn_PledgeCreation_ID And TP.SM_Prefix=LW.SM_Prefix"
tmp = tmp & " left Join Mst_Bank MB on MB.BankID = TP.BankID"
tmp = tmp & " Where LW.WHRNo = " & RsTxn_WHRDetail!WHRNo
    
Call Tmp4Table
  
If TmpRs4.RecordCount > 0 Then
   For k = 1 To TmpRs4.RecordCount
       If mTxtTop > objPDF.PDFGetPageHeight Then
          Call CheckForNewPage
       End If
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, TmpRs4!HologramNo)
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 135, mTxtTop - 5, Format(TmpRs4!PrintDate, "dd-MMM-yyyy"))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, TmpRs4!Version)
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 380, mTxtTop - 5, IIf(IsNull(TmpRs4!BankName), "", TmpRs4!BankName))
           
       objPDF.PDFDrawLine mGlobalLeft + 132, mTxtTop - mRowHeight, mGlobalLeft + 132, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 376, mTxtTop - mRowHeight, mGlobalLeft + 376, mTxtTop
            
       objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
       mTxtTop = mTxtTop + (mRowHeight)
       TmpRs4.MoveNext
  Next
Else
  If mTxtTop > objPDF.PDFGetPageHeight Then
     Call CheckForNewPage
  End If
  objPDF.PDFDrawLine mGlobalLeft + 132, mTxtTop - mRowHeight, mGlobalLeft + 132, mTxtTop
  objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
  objPDF.PDFDrawLine mGlobalLeft + 376, mTxtTop - mRowHeight, mGlobalLeft + 376, mTxtTop
  objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
  mTxtTop = mTxtTop + (mRowHeight)
End If


mTxtTop = mTxtTop + (mRowHeight)
If mTxtTop > objPDF.PDFGetPageHeight Then
   Call CheckForNewPage
End If
Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Terms & Conditions can be viewed in http://www.ncmsl.com/whrtNc.html")


tmp = "SELECT  REPLACE(CONVERT(CHAR(10), Concurrency, 103),'/','') + REPLACE(CONVERT(CHAR(12), Concurrency,114),':','') as Concurrency from Txn_WHR_Detail Where WHRNo=" & TxtWHRNo
Call Tmp1Table

Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 410, mTxtTop - 5, "Validation Code:" & IIf(IsNull(RsTxn_WHRDetail!CreatedBy), "", RsTxn_WHRDetail!CreatedBy) & IIf(IsNull(TmpRs1!Concurrency), "", TmpRs1!Concurrency))

If LCase(mTypeFlag) = "np" Then
   Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 510, mTxtTop + (20 * 7), "F&F 131A")
ElseIf LCase(mTypeFlag) = "p" Then
   Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 510, mTxtTop + (20 * 7), "F&F 131B")
End If

If mLotWiseWHR = "1" Then
   Call CheckForNewPage
   mTxtTop = mTxtTop + mRowHeight
   
   Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
   
   mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
   mStrLength = Len(GetCMPAddress(RsTxn_WHRDetail!CMPID, "D") & " Annexure to WHR for Lot Details")
   mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
   mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
   mFontSize = 12
   
   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, GetCMPAddress(RsTxn_WHRDetail!CMPID, "D") & " Annexure to WHR for Lot Details")
   
   mHeight = mTop + mRowHeight
   mLeft = mGlobalLeft
   
   Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
   
   mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
   mStrLength = Len("Warehouse Receipt")
   mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
   mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
   mFontSize = 7
   
   '''Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mTxtLeft - 26, mTxtTop - 3, "Non Negotiable Warehouse Receipt")
   mHeadingTop = mTxtTop - 3
   mHeadingLeft = mTxtLeft - 26
   
   mHeight = mTop + mRowHeight
   mLeft = mGlobalLeft
   
'   mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'   mStrLength = Len("Depositor Details")
'   mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'   mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'   mFontSize = 7
  
   Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Receipt No")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, TxtWHRNoNew)
   '--
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Receipt Date")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, Format(WHRDate, "dd-MMM-yyyy"))
   '--
   mTxtTop = mTxtTop + mRowHeight
   Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Warehouse License No")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, CStr(mLicenseNo))
   '--
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "License Expiry Date")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, Format(CDate(mLicenseExpiryDate), "dd-MMM-yyyy"))
   '--
   
'   mTxtTop = mTxtTop + mRowHeight
'   Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)
'
'   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Commodity")
'   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, CmbWHRCommodity.Text)
   
   
   mTxtTop = mTxtTop + mRowHeight
   
   Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Location & Address of Warehouse")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, GetLocationFromCMP(RsMst_Godown!CMPID))
   '--
   
   objPDF.PDFDrawLine mGlobalLeft + 132.5, mTop - mRowHeight, mGlobalLeft + 132.5, mTop 'mRowHeight  'mRowHeight 'mHeight - 5 'mRowHeight + 100
   
   mTxtTop = mTxtTop + mRowHeight
   
   Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Godown No")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, GetGodownName(RsTxn_WHRDetail!GodownID, "N"))
   
   objPDF.PDFDrawLine mGlobalLeft + 132.5, mTop - mRowHeight, mGlobalLeft + 132.5, mTop
   
   mTxtTop = mTxtTop + mRowHeight
   '    Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Godown Address")
   
'   Dim mGodownAddr As String
   mGodownAddr = GetGodownName(RsTxn_WHRDetail!GodownID, "A")
'   Dim mLen As Integer, mMax As Integer, mPara As Integer
'   Dim t1 As Double
   mMax = 120
   mPara = 0
   t1 = 1
   mLen = Len(mGodownAddr)
   mLen = RoundUp(mLen / mMax) '+ 1
   For T = 1 To mLen
       If T > 1 Then
          t1 = mMax + t1
           'mMax '= mMax + 35
          mPara = mPara + 7
       End If
       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara - 5, mID$(mGodownAddr, t1, mMax))
   Next
   
   mTxtTop = mTxtTop + (mRowHeight)
   
'   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft + 132.5, mTop - (mRowHeight * (T - 1)), mGlobalLeft + 132.5, mTxtTop 'mRowHeight
   
   mTxtTop = mTxtTop + (mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Name and Address of the Depositor")
       
'   Dim mDepName As String
   ans = GetClientNameByRow(mClientIDRow)
   
   If ans(4) <> "" Then
      mDepName = TxtClientID & "-" & Trim(ans(4))
   Else
      mDepName = TxtClientID
   End If
   
   mMax = 145
   mPara = 0
   t1 = 1
   mLen = Len(mDepName)
   mLen = RoundUp(mLen / mMax) '+ 1
   For T = 1 To mLen
       If T > 1 Then
          t1 = mMax + t1
          'mMax '= mMax + 35
          mPara = mPara + 7
       End If
       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara - 5, mID$(mDepName, t1, mMax))
   Next
   
   mTxtTop = mTxtTop + (mRowHeight)
    
   objPDF.PDFDrawLine mGlobalLeft + 132.5, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 132.5, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
 
   mTxtTop = mTxtTop + (mRowHeight)
    
    'Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)
    
   If ChkSaidToContain.Value = 1 Then
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Commodity (Said to contain)")
   Else
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Commodity")
   End If
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, TxtWHRCommodityId.Text)
    
    
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Condition of Commodity")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, CmbCommodityCondition)
   
   objPDF.PDFDrawLine mGlobalLeft + 132.5, mTxtTop - mRowHeight, mGlobalLeft + 132.5, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft + 420, mTxtTop - mRowHeight, mGlobalLeft + 420, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

 
   mTxtTop = mTxtTop + (mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "GSL Details")
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   
   mTxtTop = mTxtTop + (mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Stack No")
   ' Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5 + mRowHeight, "Manager")
       
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "Lot No")
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop - 5, "No Of Bags")
    
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, "Quantity in MT")
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, "WHR No")
    
   objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
   
   objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
    
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   
   mTxtTop = mTxtTop + (mRowHeight)
    
   For k = 2 To MSHFlexGrid4.Rows - 2
       
       If mTxtTop > objPDF.PDFGetPageHeight Then
          Call CheckForNewPage
       End If
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, Trim(MSHFlexGrid4.TextMatrix(k, 1)))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, Trim(MSHFlexGrid4.TextMatrix(k, 2)))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop - 5, Val(MSHFlexGrid4.TextMatrix(k, 5)))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, Val(MSHFlexGrid4.TextMatrix(k, 6)))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, Trim(TxtWHRNoNew))
  
       objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
       
        mTxtTop = mTxtTop + (mRowHeight)
   Next
 
End If
objPDF.PDFEndDoc

End Sub

Private Sub PDF_Print_RptLetterHead()

Dim mHeadingTop As Double
Dim mHeadingLeft As Double

Dim ans As Variant
Set objPDF = New mjwPDF

I = 2
' Create a simple PDF file using the mjwPDF class

Gen_mTimeStamp = GetTimeStamp
'Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDTF-" & Trim(TxtPreFix) & "" & Trim(TxtCxTFNo) & ".xls")
    
    
' Set the PDF title and filename
objPDF.PDFTitle = "WHR PDF Document"
objPDF.PDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_WHR-" & Trim(TxtWHRNoNew) & ".pdf"
mPDFFileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "_WHR-" & Trim(TxtWHRNoNew) & ".pdf"
' We must tell the class where the PDF fonts are located
objPDF.PDFLoadAfm = App.Path & "\Fonts"

' Set the file properties
objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
objPDF.PDFFormatPage = FORMAT_A4
objPDF.PDFOrientation = ORIENT_PORTRAIT
objPDF.PDFSetUnit = UNIT_PT
objPDF.PDFSetViewerPreferences = VIEW_HIDEMENUBAR & VIEW_HIDETOOLBAR & VIEW_HIDEWINDOWUI & VIEW_HIDERIGHTCLICK '& VIEW_DISABLEFILEMENU


' Lets us set see the bookmark pane when we view the PDF
'objPDF.PDFUseOutlines = True

' View the PDF file after we create it
objPDF.PDFView = True
'    objPDF.PDFWatermark = "NCMSL"
' Begin our PDF document

objPDF.PDFBeginDoc
    ' Lets add a heading
objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True

objPDF.PDFCell " ", 15, 80, _
objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 125

If PntPrv = "Y" Then
   objPDF.PDFWatermark = "Preview"
Else
   objPDF.PDFWatermark = "NCML" '"NCMSL"
End If

objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
objPDF.PDFSetDrawColor = vbWhite
objPDF.PDFSetTextColor = vbBlack
objPDF.PDFSetAlignement = ALIGN_CENTER
objPDF.PDFSetBorder = BORDER_ALL
objPDF.PDFSetFill = True

pgval = ((objPDF.PDFGetPageWidth - 18))
 
mRowHeight = 15
mGlobalLeft = 12
'objPDF.PDFDrawPolygon Array(12, 15, objPDF.PDFGetPageWidth - 18, 15, objPDF.PDFGetPageWidth - 18, 30, 12, 30)
mLeft = mGlobalLeft
mTop = 80
mHeight = mTop + mRowHeight
mTxtFromTop = mRowHeight + 2
 
 'Call DrawTable(10, 2, 15, 30, 12)

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len(GetCMPAddress(RsTxn_WHRDetail!CMPID, "D"))
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 12

Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, GetCMPAddress(RsTxn_WHRDetail!CMPID, "D"))

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft


Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Warehouse Receipt")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
mFontSize = 10

''Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mTxtLeft - 26, mTxtTop - 3, "Non Negotiable Warehouse Receipt")
mHeadingTop = mTxtTop - 3
mHeadingLeft = mTxtLeft - 26

mHeight = mTop + mRowHeight
mLeft = mGlobalLeft

mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
mStrLength = Len("Depositor Details")
mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
mFontSize = 7
Dim mDispTextSize  As Integer
mDispTextSize = 7

mTxtTop = mTxtTop + mRowHeight

Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Receipt No")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, TxtWHRNoNew)
'--
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Receipt Date")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, Format(WHRDate, "dd-MMM-yyyy"))
'--
mTxtTop = mTxtTop + mRowHeight

Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)

If mLotWiseWHR = 0 Then
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Warehouse License No")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, MSHFlexGrid2.TextMatrix(2, 3))
'--
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "License Expiry Date")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, Format(RsMst_License!LicenseExpiryDate, "dd-MMM-yyyy"))
'--
Else
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Warehouse License No")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, CStr(mLicenseNo))
   '--
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "License Expiry Date")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, Format(CDate(mLicenseExpiryDate), "dd-MMM-yyyy"))
End If
mTxtTop = mTxtTop + mRowHeight

Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Location & Address of Warehouse")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, GetLocationFromCMP(RsMst_Godown!CMPID))
'--

objPDF.PDFDrawLine mGlobalLeft + 132.5, mTop - mRowHeight, mGlobalLeft + 132.5, mTop 'mRowHeight  'mRowHeight 'mHeight - 5 'mRowHeight + 100

mTxtTop = mTxtTop + mRowHeight


Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Godown No")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, GetGodownName(RsTxn_WHRDetail!GodownID, "N"))
'--

objPDF.PDFDrawLine mGlobalLeft + 132.5, mTop - mRowHeight, mGlobalLeft + 132.5, mTop

mTxtTop = mTxtTop + mRowHeight
'    Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Godown Address")

Dim mGodownAddr As String
mGodownAddr = GetGodownName(RsTxn_WHRDetail!GodownID, "A")
Dim mLen As Integer, mMax As Integer, mPara As Integer
Dim t1 As Double
mMax = 120
mPara = 0
t1 = 1
mLen = Len(mGodownAddr)
mLen = RoundUp(mLen / mMax) '+ 1
For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       'mMax '= mMax + 35
       mPara = mPara + 7
    End If
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara - 5, mID$(mGodownAddr, t1, mMax))
Next

mTxtTop = mTxtTop + (mRowHeight)

objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 132.5, mTop - (mRowHeight * (T - 1)), mGlobalLeft + 132.5, mTxtTop 'mRowHeight

mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Name and Address of the Depositor")
   
 
Dim mDepName As String
ans = GetClientNameByRow(mClientIDRow)

If ans(4) <> "" Then
   mDepName = TxtClientID & "-" & Trim(ans(4))
Else
   mDepName = TxtClientID
End If

mMax = 145
mPara = 0
t1 = 1
mLen = Len(mDepName)
mLen = RoundUp(mLen / mMax) '+ 1
For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       'mMax '= mMax + 35
       mPara = mPara + 7
    End If
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara - 5, mID$(mDepName, t1, mMax))
Next

mTxtTop = mTxtTop + (mRowHeight)
 
objPDF.PDFDrawLine mGlobalLeft + 132.5, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 132.5, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)

'Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)

If ChkSaidToContain.Value = 1 Then
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Commodity (Said to contain)")
Else
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Commodity")
End If
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, TxtWHRCommodityId.Text)
 
 
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Condition of Commodity")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, CmbCommodityCondition)


 
objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop


mTxtTop = mTxtTop + (mRowHeight)


Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "No of Bags")
If mLotWiseWHR = 0 Then
   Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, "[" & Format(MSHFlexGrid2.TextMatrix(I, 7), "#######0") & "]")
Else
   Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, "[" & Format(MSHFlexGrid4.TextMatrix(0, 5), "#######0") & "]")
End If
If ChkEstimatedWgt.Value = 1 Then
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Gross Quantity(MT) (Said to Contain)")
Else
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Gross Quantity(MT)")
End If
 
If mLotWiseWHR = 0 Then
   Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "[" & Format(MSHFlexGrid2.TextMatrix(I, 8), "#######0.000") & "]")
Else
   Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "[" & Format(MSHFlexGrid4.TextMatrix(0, 6), "#######0.000") & "]")
End If
objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop


mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Name and License No of Weigher")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, IIf(TxtWeigher = "", "---", TxtWeigher))
 
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Private Marks on Packaging")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, TxtGrade)
 
objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop


mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Name & License No of Assayer")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, IIf(TxtAssayer = "", "---", TxtAssayer))

objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   

mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Quality As per Grader")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, IIf(TxtAssayerCerti = "", "---", TxtAssayerCerti))

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Grade Valid Till")

If TxtAssayer <> "" Then
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, Format(GradeValidUpto, "dd-mmm-yyyy"))
Else
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "---")
End If

objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop


mTxtTop = mTxtTop + (mRowHeight)
If OptSRPerBag.Value = True Then
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Storage Charge(Rate)-excluding Taxes")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10 + mRowHeight, "(Rupees per bag/month)")
Else
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Storage Charge(Rate)-excluding Taxes ")
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10 + mRowHeight, "(Rupees per MT/month)")
End If
Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, Format(TxtStorageCharges, "#########0.00"))
     
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Storage Charge - excluding Taxes ")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 10 + mRowHeight, "(Total Amount) per month")

Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "[" & Format(TxtStorageAmount, "###,###,###,###,##0") & "]")
 
objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop + mRowHeight - 5
objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop + mRowHeight - 5
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop + mRowHeight - 5
objPDF.PDFDrawLine mGlobalLeft, mTxtTop + mRowHeight - 5, objPDF.PDFGetPageWidth - 18, mTxtTop + mRowHeight - 5

mTxtTop = mTxtTop + (mRowHeight) + (mRowHeight) - 5

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Goods are Accepted for storage From")

Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, Format(StorageFrom, "dd-mmm-yyyy"))

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Goods are Accepted for storage UpTo")
 
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, Format(StorageTo, "dd-mmm-yyyy"))
 
objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop


mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Market Rate at the time of deposit (Rs/MT)")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10 + mRowHeight, "(Rs/MT)")

Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, "[" & Format(TxtRate, "###,###,###,##0") & "]")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Value of Commodities at time of deposit(Amount in Rs) ")
'Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 10 + mRowHeight, "")
 
Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "[" & Format(TxtValueofCommodity, "###,###,###,##0") & "]")
 
objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 466, mTxtTop - mRowHeight, mGlobalLeft + 466, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Gross Quantity in Words")

Dim mqtyStr, str1, str2, str3 As String
str1 = ""
str2 = ""
str3 = ""

'mqtyStr = Format(MSHFlexGrid2.TextMatrix(I, 8), "#####0.000")

If mLotWiseWHR = 0 Then
   mqtyStr = Format(MSHFlexGrid2.TextMatrix(I, 8), "#####0.000") 'Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "[" & Format(MSHFlexGrid2.TextMatrix(I, 8), "#######0.000") & "]")
Else
   mqtyStr = Format(MSHFlexGrid4.TextMatrix(0, 6), "#####0.000") 'Call PrintText(8, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, "[" & Format(MSHFlexGrid4.TextMatrix(0, 6), "#######0.000") & "]")
End If

str1 = NumStrConv(mID(mqtyStr, 1, InStr(mqtyStr, ".") - 1))
str2 = NumStrConv(mID(mqtyStr, InStr(mqtyStr, ".") + 1, 3))
str1 = Replace(str1, "Only", "MT")
str2 = Replace(str2, "Only", "KG")
If str2 = "" Then
   str3 = str1 & " Only"
ElseIf str1 = "" Then
   str3 = str2 & " Only"
Else
   str3 = str1 & " And " & str2 & " Only"
End If

Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, str3)

objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   
mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Value of Commodity in Words")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, NumStrConv(Val(Format(TxtValueofCommodity.Text, "######0"))))

objPDF.PDFDrawLine mGlobalLeft + 156, mTxtTop - mRowHeight, mGlobalLeft + 156, mTxtTop

objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop


''mTxtTop = mTxtTop + (mRowHeight) + (mRowHeight)
''
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop + 10, "Signature & Stamp of Warehouse")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop + mRowHeight + 5, "Manager")
''
''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, "")
''
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop + mRowHeight + 5, "Hologram No:- " & RsTxn_WHRDetail!HologramNo)
''
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop + mRowHeight + 5, "Affix Hologram Here")
''
''objPDF.PDFDrawLine mGlobalLeft + 132.5, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 132.5, mTxtTop + (mRowHeight) + 8
''
''objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 276, mTxtTop + (mRowHeight) + 8
''objPDF.PDFDrawLine mGlobalLeft + 422, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 422, mTxtTop + (mRowHeight) + 8
''objPDF.PDFDrawLine mGlobalLeft, mTxtTop + (mRowHeight) + 8, objPDF.PDFGetPageWidth - 18, mTxtTop + (mRowHeight) + 8
''
''objPDF.PDFDrawLine mGlobalLeft, mTxtTop + (mRowHeight) + 10, objPDF.PDFGetPageWidth - 18, mTxtTop + (mRowHeight) + 10
''
''
''mTxtTop = mTxtTop + (mRowHeight) + (mRowHeight) + 10
mTxtTop = mTxtTop + (mRowHeight)
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Insurance Details")
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop


mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Valid From - Valid to")
' Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5 + mRowHeight, "Manager")
    
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "Policy & Cover Note No")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop - 5, "Insurance Co.")
 
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, "Policy Amount (in Rs.)")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, "Risk Covered")
 
objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
 
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop


mTxtTop = mTxtTop + (mRowHeight)
 
If ValidFrom.Value <> "1/1/1900" Or ValidTo <> "1/1/1900" Then
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, Format(ValidFrom, "DD-MMM-YYYY") & " To ")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 10 + mRowHeight, Format(ValidTo, "DD-MMM-YYYY"))
End If
Dim a1, a2, a3 As Variant
mMax = 26
mPara = 0
t1 = 1
mLen = Len(TxtPolicyNo)
mLen = RoundUp(mLen / mMax) '+ 1
For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       mPara = mPara + 7
    End If
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop + mPara - 5, mID$(TxtPolicyNo, t1, mMax))
Next
a1 = mTxtTop + mPara - 5
mMax = 35
mPara = 0
t1 = 1
mLen = Len(TxtInsuranceCo)
mLen = RoundUp(mLen / mMax) '+ 1
For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       mPara = mPara + 7
     End If
     Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop + mPara - 5, mID$(TxtInsuranceCo, t1, mMax))
Next
'a2 = mTxtTop + mPara - 5
If TxtAmount <> "0" Then
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, TxtAmount)
End If

mMax = 50
mPara = 0
t1 = 1
mLen = Len(TxtRisks)
mLen = RoundUp(mLen / mMax) '+ 1
For T = 1 To mLen
    If T > 1 Then
       t1 = mMax + t1
       mPara = mPara + 7
    End If
    Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop + mPara - 5, mID$(TxtRisks, t1, mMax))
Next

a3 = mTxtTop + mPara - 5
 
If a1 > a3 Then
   objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, Val(a1) + 10 'mTxtTop + (mRowHeight) + a1

   objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, Val(a1) + 10 'mTxtTop + (mRowHeight) + a1
   objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, Val(a1) + 10 'mTxtTop + (mRowHeight) + a1
   objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, Val(a1) + 10 'mTxtTop + (mRowHeight) + a1
   objPDF.PDFDrawLine mGlobalLeft, Val(a1) + 10, objPDF.PDFGetPageWidth - 18, Val(a1) + 10
   mTxtTop = Val(a1) + 10
Else
   objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, Val(a3) + 10
   objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, Val(a3) + 10
   objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, Val(a3) + 10
   objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, Val(a3) + 10
   objPDF.PDFDrawLine mGlobalLeft, Val(a3) + 10, objPDF.PDFGetPageWidth - 18, Val(a3) + 10
   mTxtTop = Val(a3) + 10
End If


'''mTxtTop = mTxtTop + (mRowHeight)
'''
'''
'''
'''
'''Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "The Warehouse Receipt is valid for transfer by endorsement, and/or for  marking of plegde, subject to confirmation by Warehouse Manager")
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'''
'''mTxtTop = mTxtTop + (mRowHeight)
'''
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Pledge details: Any lien marking of this WHR is invalid without an authorised Pledge number under 'Lien/Pledge number as per StoreMan' column below")
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'''mTxtTop = mTxtTop + (mRowHeight)
'''
'''Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "This WHR can not be pledged to Banks/Financial Institutions, unless a 'Pledge Confirmation Letter' is issued by Corporate Office of NCMSL in their letter head.")
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'''mTxtTop = mTxtTop + (mRowHeight)
'''
'''
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Date")
'''' Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5 + mRowHeight, "Manager")
'''
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "Name and Address Of Lender")
'''
'''
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, "Lien/Pledge Number")
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 10 + mRowHeight, "as per StoreMan")
'''
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, "Pledge Released on")
'''
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 5, "Signature & Stamp of")
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 10 + mRowHeight, "Warehouse Manager")
'''
'''objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop + mRowHeight - 5
'''
'''objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop + mRowHeight - 5
'''objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop + mRowHeight - 5
'''objPDF.PDFDrawLine mGlobalLeft + 478, mTxtTop - mRowHeight, mGlobalLeft + 478, mTxtTop + mRowHeight - 5
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop + mRowHeight - 5, objPDF.PDFGetPageWidth - 18, mTxtTop + mRowHeight - 5
'''
'''mTxtTop = mTxtTop + mRowHeight - 5 + (mRowHeight)
'''
'''objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 478, mTxtTop - mRowHeight, mGlobalLeft + 478, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'''
'''mTxtTop = mTxtTop + (mRowHeight)
'''
'''objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 478, mTxtTop - mRowHeight, mGlobalLeft + 478, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'''
'''mTxtTop = mTxtTop + (mRowHeight)
'''
'''objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 478, mTxtTop - mRowHeight, mGlobalLeft + 478, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'''mTxtTop = mTxtTop + (mRowHeight)
'''
'''objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 478, mTxtTop - mRowHeight, mGlobalLeft + 478, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'''
'''mTxtTop = mTxtTop + (mRowHeight)

''tmp = "SELECT  TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity, sum(ReleasedQuantity)ReleasedQuantity"
''tmp = tmp & " FROM  Txn_WHR_Detail TWD"
''tmp = tmp & " inner join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
''tmp = tmp & " left Join Txn_PledgeRelease TPR on TPR.PledgeID = TPC.PledgeID And TPC.SM_Prefix = TPR.SM_Prefix"
''tmp = tmp & " left Join Txn_PledgeRelease_GSL TPRG on TPRG.PledgeReleaseID = TPR.PledgeReleaseID And TPRG.SM_Prefix = TPR.SM_Prefix"
''tmp = tmp & " Where TWD.WHRNo = " & RsTxn_WHRDetail!WHRNO
''tmp = tmp & " Group By TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity"
''tmp = tmp & " Having TWD.WHRQuantity > Sum(ReleasedQuantity)"

tmp = " SELECT  TWD.Txn_PledgeCreation_ID,TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity,isnull(sum(PledgeQuantity),0) PledgeQuantity"
tmp = tmp & " FROM  Txn_WHR_Detail TWD"
tmp = tmp & " left join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
tmp = tmp & " left Join Txn_PledgeCreation_GSL TPRG on TPRG.Txn_PledgeCreation_ID = TPC.Txn_PledgeCreation_ID And TPRG.SM_Prefix = TPC.SM_Prefix"
tmp = tmp & " Where TWD.WHRNo =  " & RsTxn_WHRDetail!WHRNo & " And TPC.Txn_PledgeCreation_ID = " & IIf(IsNull(RsTxn_WHRDetail!Txn_PledgeCreation_ID), 0, RsTxn_WHRDetail!Txn_PledgeCreation_ID)
tmp = tmp & " Group By TWD.Txn_PledgeCreation_ID,TWD.WHRNo,TWD.SM_Prefix,TWD.WHRQuantity"
tmp = tmp & " Having IsNull(Sum(PledgeQuantity), 0) > 0"

Call Tmp4Table

If TmpRs4.RecordCount = 0 Then
   'Commented on 12 April 2012 Mail by Mayur Sir for removal of Non Negotiable
   'Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mHeadingLeft, mHeadingTop, "Non Pledgable Non Negotiable Warehouse Receipt")
   
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mHeadingLeft, mHeadingTop, "            Non Pledgable Warehouse Receipt")
   
   mTxtTop = mTxtTop + (mRowHeight)
    
   Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "This WHR can not be pledged to Banks/Financial Institutions.") ', subject to confirmation by Warehouse Manager")
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   mTxtTop = mTxtTop + 5
   mTypeFlag = "NP"
Else
'''    mTxtTop = mTxtTop + (mRowHeight)
'''
'''    Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "The Warehouse Receipt is valid for transfer by endorsement, and/or for  marking of plegde.") ', subject to confirmation by Warehouse Manager")
'''    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

    'Commented on 12 April 2012 Mail by Mayur Sir for removal of Non Negotiable
    'Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mHeadingLeft, mHeadingTop, " Non Negotiable Warehouse Receipt")
    mTypeFlag = "P"
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mHeadingLeft, mHeadingTop, "            Warehouse Receipt")
    
    mTxtTop = mTxtTop + (mRowHeight)
    
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Pledge details: Any lien marking of this WHR is invalid without an authorised Pledge number under 'Lien/Pledge number as per StoreMan' column below")
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
    mTxtTop = mTxtTop + (mRowHeight)
    
    'mTxtTop = mTxtTop + 20
    
    Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "This WHR can not be pledged to Banks/Financial Institutions, unless a 'Pledge Confirmation Letter' issued by NCMSL in their letter head is attached alongwith this")
    mTxtTop = mTxtTop + 10
    Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "WHR with complete details of applicable terms & conditions of storage.")
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
    
    mTxtTop = mTxtTop + mRowHeight
    
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Date")
    
        
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "Name and Address Of Lender")
    
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, "Pledge Number")
    
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 10 + mRowHeight, "as per StoreMan")
    
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, " Quantity")
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 10 + mRowHeight, "(in MT)")
        
    
    Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 5, "Pledge Released on")
    
''    tmp = "SELECT  TPC.PledgeDate, convert(varchar,TPC.PledgeID)+'~'+TPC.PledgeReqNo as PledgeID, TPC.PledgeQuantity,"
''    tmp = tmp & " max(TPR.PledgeReleaseDate)PledgeReleaseDate, MB.BankName, MB.Address"
''    tmp = tmp & " FROM  Txn_WHR_Detail TWD"
''    tmp = tmp & " inner join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
''    tmp = tmp & " left Join Txn_PledgeRelease TPR on TPR.PledgeID = TPC.PledgeID And TPC.SM_Prefix = TPR.SM_Prefix"
''    tmp = tmp & " inner Join Mst_Branch MBr On TPC.BranchID = MBr.BranchID"
''    tmp = tmp & " inner Join Mst_Bank MB On TPC.BankID = MB.BankID"
''    tmp = tmp & " Where TWD.WHRNo=" & RsTxn_WHRDetail!WHRNo
''    tmp = tmp & " Group By TPC.PledgeDate, TPC.PledgeID,TPC.PledgeReqNo, MB.BankName, MB.Address, TPC.PledgeQuantity"
    
'    tmp = "SELECT  TPC.PledgeDate, convert(varchar,TPC.PledgeID)+'~'+TPC.PledgeReqNo as PledgeID, max(TPCG.PledgeQuantity)+sum(TPRG.ReleasedQuantity)PledgeQuantity,"
'    tmp = tmp & " max(TPR.PledgeReleaseDate)PledgeReleaseDate, MB.BankName, MB.Address"
'    tmp = tmp & " FROM  Txn_WHR_Detail TWD"
'    tmp = tmp & " inner join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
'    tmp = tmp & " inner join Txn_PledgeCreation_GSL TPCG on  TPC.PledgeID = TPCG.PledgeID And TPC.SM_Prefix = TPCG.SM_Prefix"
'    tmp = tmp & " left Join Txn_PledgeRelease TPR on TPR.PledgeID = TPC.PledgeID And TPC.SM_Prefix = TPR.SM_Prefix"
'    tmp = tmp & " left Join Txn_PledgeRelease_GSL TPRG on TPR.PledgeReleaseID = TPRG.PledgeReleaseID And TPR.SM_Prefix = TPRG.SM_Prefix"
'    tmp = tmp & " inner Join Mst_Branch MBr On TPC.BranchID = MBr.BranchID"
'    tmp = tmp & " inner Join Mst_Bank MB On TPC.BankID = MB.BankID"
'    tmp = tmp & " Where TWD.WHRNo = " & RsTxn_WHRDetail!WHRNO
'    tmp = tmp & " Group By TPC.PledgeDate, TPC.PledgeID,TPC.PledgeReqNo, MB.BankName, MB.Address"
    
    
    tmp = " SELECT  TPC.PledgeDate, convert(varchar,TPC.Txn_PledgeCreation_ID) as Txn_PledgeCreation_ID,"
    tmp = tmp & " Sum(TPCG.PledgeQuantity) PledgeQuantity, "
    tmp = tmp & "  MB.BankName, MBr.Address,TPC.Txn_PledgeCreation_ID as OriPledgeID,TPC.SM_Prefix  "
    tmp = tmp & " FROM  Txn_WHR_Detail TWD"
    tmp = tmp & " inner join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
    tmp = tmp & " inner join Txn_PledgeCreation_GSL TPCG on  TPC.Txn_PledgeCreation_ID = TPCG.Txn_PledgeCreation_ID And TPC.SM_Prefix = TPCG.SM_Prefix"
    tmp = tmp & " inner Join Mst_Branch MBr On TPC.BranchID = MBr.BranchID"
    tmp = tmp & " inner Join Mst_Bank MB On TPC.BankID = MB.BankID"
    tmp = tmp & " Where TWD.WHRNo = " & RsTxn_WHRDetail!WHRNo & " "
    tmp = tmp & " Group By TPC.PledgeDate, TPC.Txn_PledgeCreation_ID,TPC.PledgeReqNo,TPC.SM_Prefix, MB.BankName, MBr.Address"
    
    
    Call Tmp4Table
    
    objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop + mRowHeight - 5
    objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop + mRowHeight - 5

    objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop + mRowHeight - 5
    objPDF.PDFDrawLine mGlobalLeft + 478, mTxtTop - mRowHeight, mGlobalLeft + 478, mTxtTop + mRowHeight - 5
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop + mRowHeight - 5, objPDF.PDFGetPageWidth - 18, mTxtTop + mRowHeight - 5
    
    'objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
    
    mTxtTop = mTxtTop + mRowHeight + (mRowHeight) - 5
    
    mPledgeQty = 0
    
    
    
    For k = 1 To TmpRs4.RecordCount
        mPledgeQty = TmpRs4!PledgeQuantity
        
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, Format(TmpRs4!PledgeDate, "dd-MMM-yyyy"))
        
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, TmpRs4!BankName)
        
        mMax = 70
        mPara = 7
        t1 = 1
        mLen = Len(TmpRs4!Address)
        
        If mLen > 70 Then
           mLen = RoundUp(mLen / mMax)
        Else
           mLen = 1
        End If
        For T = 1 To mLen
            If T > 1 Then
               t1 = mMax + t1
               mPara = mPara + 7
            End If
            Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop + mPara - 5, mID$(TmpRs4!Address, t1, mMax))
        Next
        
        a1 = mTxtTop + mPara - 5

        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, TmpRs4!Txn_PledgeCreation_ID)
        
        mRPledge = 0
                
        tmp = "Select Sum(isnull(ReleasedQuantity,0)) RQty,Max(PledgeReleaseDate) RDate"
        tmp = tmp & " From Txn_PledgeRelease TPR"
        tmp = tmp & " left Join Txn_PledgeRelease_GSL TPRG on TPR.Txn_PledgeRelease_ID = TPRG.Txn_PledgeRelease_ID And TPR.SM_Prefix = TPRG.SM_Prefix"
        tmp = tmp & " Where TPR.Txn_PledgeCreation_ID = " & TmpRs4!OriPledgeID & " And TPR.SM_Prefix = '" & TmpRs4!SM_Prefix & "'"
                
        Call Tmp5Table
        
        If TmpRs5.RecordCount > 0 Then
           If IsNull(TmpRs5!RQty) = False Then
              mRPledge = TmpRs5!RQty
           End If
           If Val(RsTxn_WHRDetail!WHRQuantity) >= Val(mRPledge) Then
              Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 5, IIf(IsNull(TmpRs5!RDate), "", Format(TmpRs5!RDate, "dd-MMM-yyyy")))
           End If
        Else
           Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 5, "")
        End If
                    
        Dim mTotPledge As String
        mTotPledge = Val(mPledgeQty) + Val(mRPledge)
                
                
        'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, IIf(IsNull(TmpRs4!PledgeQuantity), 0, TmpRs4!PledgeQuantity))
        
        Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, mTotPledge)
                
        'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, IIf(IsNull(TmpRs4!PledgeQuantity), 0, TmpRs4!PledgeQuantity))
        
        'Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 5, IIf(IsNull(TmpRs4!PledgeReleaseDate), "", Format(TmpRs4!PledgeReleaseDate, "dd-MMM-yyyy")))
    
        objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, Val(a1) + 5 'mTxtTop + mRowHeight - 5
        objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, Val(a1) + 5  'mTxtTop + mRowHeight - 5
        objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, Val(a1) + 5 'mTxtTop + mRowHeight - 5
        objPDF.PDFDrawLine mGlobalLeft + 478, mTxtTop - mRowHeight, mGlobalLeft + 478, Val(a1) + 5 'mTxtTop + mRowHeight - 5
        objPDF.PDFDrawLine mGlobalLeft, Val(a1) + 5, objPDF.PDFGetPageWidth - 18, Val(a1) + 5  'mTxtTop + mRowHeight - 5
         
        mTxtTop = Val(a1) + mRowHeight + 5
        
        TmpRs4.MoveNext
        
    Next
    mTxtTop = Val(a1) + mRowHeight - 5
End If

mTxtTop = mTxtTop + (mRowHeight) - 5

'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Withdrawal History")
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "The good mentioned below are hereby released for delivery from the Warehouse. The unreleased balance is subject to lien for unpaid charges")
'''Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 10 + mRowHeight, "and borrowings from lending institutions.")
'''
'''objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop + mRowHeight
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop + mRowHeight, objPDF.PDFGetPageWidth - 18, mTxtTop + mRowHeight
'''
'''mTxtTop = mTxtTop + (mRowHeight) + (mRowHeight)
'''
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "Quantity Released")
'''
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, "Quantity Balance")
'''
'''objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
'''
'''objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
'''
'''objPDF.PDFDrawLine mGlobalLeft + 478, mTxtTop - mRowHeight, mGlobalLeft + 478, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop, mGlobalLeft + 478, mTxtTop
'''
'''
'''
'''mTxtTop = mTxtTop + (mRowHeight)
'''
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Date")
'''
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "No. of Bags")
'''
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop - 5, "Weight in MT")
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, "No. Of Bags")
'''
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, "Weight in  MT")
'''
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 15, "Signature & Stamp of ")
'''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 5, "Warehouse Manager")
'''
'''
'''objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 478, mTxtTop - mRowHeight, mGlobalLeft + 478, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'''
'''mTxtTop = mTxtTop + (mRowHeight)
'''
'''objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 478, mTxtTop - mRowHeight, mGlobalLeft + 478, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'''
'''mTxtTop = mTxtTop + (mRowHeight)
'''
'''objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft + 478, mTxtTop - mRowHeight, mGlobalLeft + 478, mTxtTop
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'''
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop + 10, objPDF.PDFGetPageWidth - 18, mTxtTop + 10
'''
'''mTxtTop = mTxtTop + (mRowHeight) + 10
'''
''''Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "The Warehouse Receipt is valid for transfer by endorsement, and/or for  marking of plegde, subject to confirmation by Warehouse Manager")
''''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
''''
''''mTxtTop = mTxtTop + (mRowHeight)
'''
'''Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "This Warehouse Receipt is not valid without an official Hologram of NCMSL affixed above.")
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'''
'''mTxtTop = mTxtTop + (mRowHeight)
'''
'''Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "If the commodities are not withdrawn by 'Goods are Accepted for storage UpTo' date mentioned in the WHR,  NCMSL shall be entitled to dispose the same at the risk ")
'''mTxtTop = mTxtTop + 10
'''Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "and responsibility of the depositor and remit the proceeds to the depositor after deducting the applicable fees and charges.")
'''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'''
''''mTxtTop = mTxtTop + 20
''''Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "This WHR can not be pledged to Banks/Financial Institutions, unless a 'Pledge Confirmation Letter' is issued by Corporate Office of NCMSL in their letter head.")
''''objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'''
'''mTxtTop = mTxtTop + 15
''''mTxtTop = mTxtTop + 45
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Withdrawal History")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "The good mentioned below are hereby released for delivery from the Warehouse. The unreleased balance is subject to lien for unpaid charges")
Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 10 + mRowHeight, "and borrowings from lending institutions.")

objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop + mRowHeight
objPDF.PDFDrawLine mGlobalLeft, mTxtTop + mRowHeight, objPDF.PDFGetPageWidth - 18, mTxtTop + mRowHeight

mTxtTop = mTxtTop + (mRowHeight) + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 240, mTxtTop - 5, "Quantity Released")
 
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, "Quantity Balance")

objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 348, mTxtTop - mRowHeight, mGlobalLeft + 348, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 518, mTxtTop - mRowHeight, mGlobalLeft + 518, mTxtTop

objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop, mGlobalLeft + 518, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Date")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "CITF No")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop - 5, "No. of Bags")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 270, mTxtTop - 5, "Quantity in MT")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop - 5, "No. Of Bags")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 440, mTxtTop - 5, "Quantity in  MT")

''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 580, mTxtTop - 15, "Signature & Stamp of ")
''Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 480, mTxtTop - 5, "Warehouse Manager")
objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 268, mTxtTop - mRowHeight, mGlobalLeft + 268, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 348, mTxtTop - mRowHeight, mGlobalLeft + 348, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 438, mTxtTop - mRowHeight, mGlobalLeft + 438, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 518, mTxtTop - mRowHeight, mGlobalLeft + 518, mTxtTop
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, mGlobalLeft + 518, mTxtTop '' objPDF.PDFGetPageWidth - 18, mTxtTop
  
mTxtTop = mTxtTop + (mRowHeight)
  
''tmp = "select CxTFDate,sum(b.NoOfBags) as NoOfBags,sum(b.GrossWeight) as GrossWeight"
''tmp = tmp & " from Txn_CxTF_GSL b"
''tmp = tmp & " inner join  txn_cxtf_details a on  a.SM_Prefix = b.SM_Prefix and a.CxTFNo  = b.CxTFNo and a.TxnType = b.TxnType"
''tmp = tmp & " Where WHRNo= " & RsTxn_WHRDetail!WHRNO & "  And a.TxnType='W' And a.SM_Prefix='" & RsTxn_WHRDetail!SM_Prefix & "'"
''tmp = tmp & " Group By CxTFDate"
''
''Call Tmp4Table

Dim mWHRWt As Double
Dim mWHRBag As Double

mWHRWt = RsTxn_WHRDetail!WHRQuantity
mWHRBag = RsTxn_WHRDetail!WHRNoOfBags

If MSHFlexGrid3.Rows > 3 Then
   For k = 2 To MSHFlexGrid3.Rows - 2
       mWHRWt = Format(mWHRWt - Val(MSHFlexGrid3.TextMatrix(k, 3)), "#####0.000")
       mWHRBag = Format(mWHRBag - Val(MSHFlexGrid3.TextMatrix(k, 4)), "#####0")
       
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, Format(MSHFlexGrid3.TextMatrix(k, 2), "dd-MMM-yyyy"))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, MSHFlexGrid3.TextMatrix(k, 1))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop - 5, Format(Val(MSHFlexGrid3.TextMatrix(k, 4)), "#####0"))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 270, mTxtTop - 5, Format(Val(MSHFlexGrid3.TextMatrix(k, 3)), "#####0.000"))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 350, mTxtTop - 5, Str(Format(mWHRBag, "#####0")))
'       If mWHRWt < 1 Then
'          Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 440, mTxtTop - 5, "0" & Str(Format(mWHRWt, "#####0.000")))
'       Else
          Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 440, mTxtTop - 5, Str(Format(mWHRWt, "#####0.000")))
'       End If
       
       objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 268, mTxtTop - mRowHeight, mGlobalLeft + 268, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 348, mTxtTop - mRowHeight, mGlobalLeft + 348, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 438, mTxtTop - mRowHeight, mGlobalLeft + 438, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 518, mTxtTop - mRowHeight, mGlobalLeft + 518, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft, mTxtTop, mGlobalLeft + 518, mTxtTop
       
        mTxtTop = mTxtTop + (mRowHeight)
    Next
    mTxtTop = mTxtTop - 5
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
Else
    objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 268, mTxtTop - mRowHeight, mGlobalLeft + 268, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 348, mTxtTop - mRowHeight, mGlobalLeft + 348, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 438, mTxtTop - mRowHeight, mGlobalLeft + 438, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 518, mTxtTop - mRowHeight, mGlobalLeft + 518, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, mGlobalLeft + 518, mTxtTop
    
    mTxtTop = mTxtTop + mRowHeight
    
    objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 268, mTxtTop - mRowHeight, mGlobalLeft + 268, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 348, mTxtTop - mRowHeight, mGlobalLeft + 348, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 438, mTxtTop - mRowHeight, mGlobalLeft + 438, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 518, mTxtTop - mRowHeight, mGlobalLeft + 518, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, mGlobalLeft + 518, mTxtTop
    
    mTxtTop = mTxtTop + mRowHeight

    objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 268, mTxtTop - mRowHeight, mGlobalLeft + 268, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 348, mTxtTop - mRowHeight, mGlobalLeft + 348, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 438, mTxtTop - mRowHeight, mGlobalLeft + 438, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft + 518, mTxtTop - mRowHeight, mGlobalLeft + 518, mTxtTop
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
    
    mTxtTop = mTxtTop + 10
    
    objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
    
End If

mTxtTop = mTxtTop + (mRowHeight)

'mTxtTop = mTxtTop + (mRowHeight) + 10

'Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "The Warehouse Receipt is valid for transfer by endorsement, and/or for  marking of plegde, subject to confirmation by Warehouse Manager")
'objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
'
'mTxtTop = mTxtTop + (mRowHeight)
If mNCMLFinance = 0 Then
   Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "This Warehouse Receipt is not valid without an official Hologram of NCMSL affixed below.")
End If
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)

If mTxtTop > objPDF.PDFGetPageHeight - 50 Then
   Call CheckForNewPageLetter
End If
Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "If the commodities are not withdrawn by 'Goods are Accepted for storage UpTo' date mentioned in the WHR,  NCMSL shall be entitled to dispose the same at the risk ")
mTxtTop = mTxtTop + 10
Call PrintText(7, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "and responsibility of the depositor and remit the proceeds to the depositor after deducting the applicable fees and charges.")
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop


mTxtTop = mTxtTop + (mRowHeight) + (mRowHeight)
If mTxtTop > objPDF.PDFGetPageHeight - 50 Then
   Call CheckForNewPageLetter
End If

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop + 10, "Signature & Stamp of Warehouse")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop + mRowHeight + 5, "Manager")

Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop, "")

If mNCMLFinance = 0 Then
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop + mRowHeight + 5, "Hologram No:- " & RsTxn_WHRDetail!HologramNo)
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop + mRowHeight + 5, "Affix Hologram Here")
End If

objPDF.PDFDrawLine mGlobalLeft + 132.5, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 132.5, mTxtTop + (mRowHeight) + 8

objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 276, mTxtTop + (mRowHeight) + 8
objPDF.PDFDrawLine mGlobalLeft + 422, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 422, mTxtTop + (mRowHeight) + 8
objPDF.PDFDrawLine mGlobalLeft, mTxtTop + (mRowHeight) + 8, objPDF.PDFGetPageWidth - 18, mTxtTop + (mRowHeight) + 8

objPDF.PDFDrawLine mGlobalLeft, mTxtTop + (mRowHeight) + 10, objPDF.PDFGetPageWidth - 18, mTxtTop + (mRowHeight) + 10

mTxtTop = mTxtTop + (mRowHeight) + 10 + (mRowHeight)



Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "WHR Versions :")
objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

mTxtTop = mTxtTop + (mRowHeight)
If mTxtTop > objPDF.PDFGetPageHeight - 50 Then
   Call CheckForNewPageLetter
End If
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Hologram No")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 135, mTxtTop - 5, " Date")

Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "WHR Type")
Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 380, mTxtTop - 5, "Pledgee Bank Name")

objPDF.PDFDrawLine mGlobalLeft + 132, mTxtTop - mRowHeight, mGlobalLeft + 132, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
objPDF.PDFDrawLine mGlobalLeft + 376, mTxtTop - mRowHeight, mGlobalLeft + 376, mTxtTop

objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop '' objPDF.PDFGetPageWidth - 18, mTxtTop
  
mTxtTop = mTxtTop + (mRowHeight)
  
'tmp = "Select * from Log_WHRPrint Where WHRNo= " & RsTxn_WHRDetail!WHRNO & ""

tmp = "Select LW.HologramNo,LW.PrintedDate,LW.Version,MB.BankName"
tmp = tmp & " From Log_WHRPrint LW"
tmp = tmp & " left join Txn_PledgeCreation TP on TP.Txn_PledgeCreation_ID=LW.Txn_PledgeCreation_ID And TP.SM_Prefix=LW.SM_Prefix"
tmp = tmp & " left Join Mst_Bank MB on MB.BankID = TP.BankID"
tmp = tmp & " Where LW.WHRNo = " & RsTxn_WHRDetail!WHRNo

Call Tmp4Table

If TmpRs4.RecordCount > 0 Then
   For k = 1 To TmpRs4.RecordCount
       If mTxtTop > objPDF.PDFGetPageHeight - 50 Then
          Call CheckForNewPageLetter
       End If
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, TmpRs4!HologramNo)
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 135, mTxtTop - 5, Format(TmpRs4!PrintedDate, "dd-MMM-yyyy"))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, TmpRs4!Version)
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 380, mTxtTop - 5, IIf(IsNull(TmpRs4!BankName), "", TmpRs4!BankName))
       
       objPDF.PDFDrawLine mGlobalLeft + 132, mTxtTop - mRowHeight, mGlobalLeft + 132, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 376, mTxtTop - mRowHeight, mGlobalLeft + 376, mTxtTop
        
       objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
       
       mTxtTop = mTxtTop + (mRowHeight)
       TmpRs4.MoveNext
   Next
Else
   If mTxtTop > objPDF.PDFGetPageHeight - 50 Then
      Call CheckForNewPageLetter
   End If
   objPDF.PDFDrawLine mGlobalLeft + 132, mTxtTop - mRowHeight, mGlobalLeft + 132, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft + 376, mTxtTop - mRowHeight, mGlobalLeft + 376, mTxtTop
    
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
    
   mTxtTop = mTxtTop + (mRowHeight)
End If

mTxtTop = mTxtTop + (mRowHeight)

If mTxtTop > objPDF.PDFGetPageHeight - 50 Then
   Call CheckForNewPageLetter
End If

Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Terms & Conditions can be viewed in http://www.ncmsl.com/whrtNc.html")

tmp = "SELECT  REPLACE(CONVERT(CHAR(10), Concurrency, 103),'/','') + REPLACE(CONVERT(CHAR(12), Concurrency,114),':','') as Concurrency from Txn_WHR_Detail Where WHRNo=" & TxtWHRNo
Call Tmp1Table

Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 410, mTxtTop - 5, "Validation Code:" & IIf(IsNull(RsTxn_WHRDetail!CreatedBy), "", RsTxn_WHRDetail!CreatedBy) & IIf(IsNull(TmpRs1!Concurrency), "", TmpRs1!Concurrency))

If LCase(mTypeFlag) = "np" Then
   Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 510, mTxtTop + (20 * 2), "F&F 131A")
ElseIf LCase(mTypeFlag) = "p" Then
   Call PrintText(7, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 510, mTxtTop + (20 * 2), "F&F 131B")
End If


If mLotWiseWHR = "1" Then
   Call CheckForNewPageLetter
   mTxtTop = mTxtTop + mRowHeight
   
   Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
   
   mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
   mStrLength = Len(GetCMPAddress(RsTxn_WHRDetail!CMPID, "D") & " Annexure to WHR for Lot Details")
   mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
   mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
   mFontSize = 12
   
   Call PrintText(mFontSize, FONT_NORMAL, ALIGN_CENTER, mTxtLeft, mTxtTop, GetCMPAddress(RsTxn_WHRDetail!CMPID, "D") & " Annexure to WHR for Lot Details")
   
   mHeight = mTop + mRowHeight
   mLeft = mGlobalLeft
   
   Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
   
   mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
   mStrLength = Len("Warehouse Receipt")
   mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
   mTxtLeft = mPrintPoint - (mStrLength * 2) '* 5) '270
   mFontSize = 7
   
   '''Call PrintText(mFontSize, FONT_BOLD, ALIGN_CENTER, mTxtLeft - 26, mTxtTop - 3, "Non Negotiable Warehouse Receipt")
   mHeadingTop = mTxtTop - 3
   mHeadingLeft = mTxtLeft - 26
   
   mHeight = mTop + mRowHeight
   mLeft = mGlobalLeft
   
'   mTxtTop = mHeight - (mTxtFromTop) 'mTop ' + 15
'   mStrLength = Len("Depositor Details")
'   mPrintPoint = ((pgval - mGlobalLeft) - mStrLength) / 2
'   mTxtLeft = mPrintPoint '- (mStrLength) '* 5) '270
'   mFontSize = 7
  
   Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Receipt No")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, TxtWHRNoNew)
   '--
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Receipt Date")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, Format(WHRDate, "dd-MMM-yyyy"))
   '--
   mTxtTop = mTxtTop + mRowHeight
   Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Warehouse License No")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, CStr(mLicenseNo))
   '--
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "License Expiry Date")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 430, mTxtTop - 5, Format(CDate(mLicenseExpiryDate), "dd-MMM-yyyy"))
   '--
   
'   mTxtTop = mTxtTop + mRowHeight
'   Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)
'
'   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Commodity")
'   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, CmbWHRCommodity.Text)
   
   
   mTxtTop = mTxtTop + mRowHeight
   
   Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Location & Address of Warehouse")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, GetLocationFromCMP(RsMst_Godown!CMPID))
   '--
   
   objPDF.PDFDrawLine mGlobalLeft + 132.5, mTop - mRowHeight, mGlobalLeft + 132.5, mTop 'mRowHeight  'mRowHeight 'mHeight - 5 'mRowHeight + 100
   
   mTxtTop = mTxtTop + mRowHeight
   
   Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Godown No")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop - 5, GetGodownName(RsTxn_WHRDetail!GodownID, "N"))
   
   objPDF.PDFDrawLine mGlobalLeft + 132.5, mTop - mRowHeight, mGlobalLeft + 132.5, mTop
   
   mTxtTop = mTxtTop + mRowHeight
   '    Call DrawTable(, , mTop, mHeight, mLeft, mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Godown Address")
   
'   Dim mGodownAddr As String
   mGodownAddr = GetGodownName(RsTxn_WHRDetail!GodownID, "A")
'   Dim mLen As Integer, mMax As Integer, mPara As Integer
'   Dim t1 As Double
   mMax = 120
   mPara = 0
   t1 = 1
   mLen = Len(mGodownAddr)
   mLen = RoundUp(mLen / mMax) '+ 1
   For T = 1 To mLen
       If T > 1 Then
          t1 = mMax + t1
           'mMax '= mMax + 35
          mPara = mPara + 7
       End If
       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara - 5, mID$(mGodownAddr, t1, mMax))
   Next
   
   mTxtTop = mTxtTop + (mRowHeight)
   
'   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft + 132.5, mTop - (mRowHeight * (T - 1)), mGlobalLeft + 132.5, mTxtTop 'mRowHeight
   
   mTxtTop = mTxtTop + (mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Name and Address of the Depositor")
       
'   Dim mDepName As String
   ans = GetClientNameByRow(mClientIDRow)
   
   If ans(4) <> "" Then
      mDepName = TxtClientID & "-" & Trim(ans(4))
   Else
      mDepName = TxtClientID
   End If
   
   mMax = 145
   mPara = 0
   t1 = 1
   mLen = Len(mDepName)
   mLen = RoundUp(mLen / mMax) '+ 1
   For T = 1 To mLen
       If T > 1 Then
          t1 = mMax + t1
          'mMax '= mMax + 35
          mPara = mPara + 7
       End If
       Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 140, mTxtTop + mPara - 5, mID$(mDepName, t1, mMax))
   Next
   
   mTxtTop = mTxtTop + (mRowHeight)
    
   objPDF.PDFDrawLine mGlobalLeft + 132.5, mTxtTop - mRowHeight - mRowHeight, mGlobalLeft + 132.5, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
 
   mTxtTop = mTxtTop + (mRowHeight)
    
    'Call DrawTable(, 4, mTop, mHeight, mLeft, mRowHeight)
    
   If ChkSaidToContain.Value = 1 Then
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Commodity (Said to contain)")
   Else
      Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Commodity")
   End If
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 160, mTxtTop - 5, TxtWHRCommodityId.Text)
    
    
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 280, mTxtTop - 5, "Condition of Commodity")
   Call PrintText(mDispTextSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 470, mTxtTop - 5, CmbCommodityCondition)
   
   objPDF.PDFDrawLine mGlobalLeft + 132.5, mTxtTop - mRowHeight, mGlobalLeft + 132.5, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft + 276, mTxtTop - mRowHeight, mGlobalLeft + 276, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft + 420, mTxtTop - mRowHeight, mGlobalLeft + 420, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop

 
   mTxtTop = mTxtTop + (mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "GSL Details")
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   
   mTxtTop = mTxtTop + (mRowHeight)
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, "Stack No")
   ' Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5 + mRowHeight, "Manager")
       
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, "Lot No")
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop - 5, "No Of Bags")
    
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, "Quantity in MT")
   
   Call PrintText(mFontSize, FONT_BOLD, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, "WHR No")
    
   objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
   
   objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
   objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
    
   objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
   
   mTxtTop = mTxtTop + (mRowHeight)
    
   For k = 2 To MSHFlexGrid4.Rows - 2
       
       If mTxtTop > objPDF.PDFGetPageHeight - 50 Then
          Call CheckForNewPageLetter
       End If
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft, mTxtTop - 5, Trim(MSHFlexGrid4.TextMatrix(k, 1)))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 80, mTxtTop - 5, Trim(MSHFlexGrid4.TextMatrix(k, 2)))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 180, mTxtTop - 5, Val(MSHFlexGrid4.TextMatrix(k, 5)))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 320, mTxtTop - 5, Val(MSHFlexGrid4.TextMatrix(k, 6)))
       Call PrintText(mFontSize, FONT_NORMAL, ALIGN_LEFT, mGlobalLeft + 400, mTxtTop - 5, Trim(TxtWHRNoNew))
  
       objPDF.PDFDrawLine mGlobalLeft + 78, mTxtTop - mRowHeight, mGlobalLeft + 78, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 178, mTxtTop - mRowHeight, mGlobalLeft + 178, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 318, mTxtTop - mRowHeight, mGlobalLeft + 318, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft + 398, mTxtTop - mRowHeight, mGlobalLeft + 398, mTxtTop
       objPDF.PDFDrawLine mGlobalLeft, mTxtTop, objPDF.PDFGetPageWidth - 18, mTxtTop
       
        mTxtTop = mTxtTop + (mRowHeight)
   Next
 
End If



objPDF.PDFEndDoc

End Sub


Private Sub SetScrollBars()
'    Frame1.Width = 10000
'    Frame1.Height = 20000
VScroll1.Top = 0
HScroll1.Left = 0
VScroll1.Left = Me.Width - (VScroll1.Width + 120)
HScroll1.Top = Me.Height - ((HScroll1.Height * 2) + 250)
VScroll1.Height = Me.Height - 500
HScroll1.Width = Me.Width - 500
VScroll1.Min = 0
HScroll1.Min = 0
VScroll1.Max = Me.ScaleHeight - Frame4.Height
VScroll1.LargeChange = Me.ScaleHeight
VScroll1.SmallChange = Me.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame4.Width
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True
If Frame4.Height < Me.Height Then VScroll1.Enabled = False
If Frame4.Width >= Me.Width Then HScroll1.Enabled = True
If Frame4.Width < Me.Width Then HScroll1.Enabled = False
End Sub

Private Sub VScroll1_Change()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

'Frame3.Top = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()

   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = VScroll1.Value
   a = Me.Height - Frame4.Height
   l = (a * x) / 1000
   Frame4.Top = -l

'Frame3.Top = VScroll1.Value
End Sub

Private Sub HScroll1_Change()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   l = (a * x) / 1000
   Frame4.Left = -l
End Sub

Private Sub HScroll1_Scroll()
   Dim l As Double
   Dim a As Double
   Dim x As Double
   
   x = HScroll1.Value
   a = Me.Width - Frame4.Width
   l = (a * x) / 1000
   Frame4.Left = -l

End Sub
Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub CheckForNewPageLetter()
   objPDF.PDFWatermark = ""
   objPDF.PDFEndPage
   objPDF.PDFNewPage
   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
   objPDF.PDFSetDrawColor = vbWhite
   objPDF.PDFSetTextColor = vbBlack
   objPDF.PDFSetAlignement = ALIGN_CENTER
   objPDF.PDFSetBorder = BORDER_ALL
   objPDF.PDFSetFill = True

   objPDF.PDFCell " ", 15, 80, _
   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 125

   If PntPrv = "Y" Then
      objPDF.PDFWatermark = "Preview"
   Else
      objPDF.PDFWatermark = "NCML" '"NCMSL"
   End If
   
   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
   objPDF.PDFSetDrawColor = vbWhite
   objPDF.PDFSetTextColor = vbBlack
   objPDF.PDFSetAlignement = ALIGN_CENTER
   objPDF.PDFSetBorder = BORDER_ALL
   objPDF.PDFSetFill = True

   mLeft = mGlobalLeft
   mTxtTop = 95
   mHeight = mTxtTop + mRowHeight
   mTxtFromTop = mRowHeight + 2
End Sub

Private Sub CheckForNewPage()
   objPDF.PDFWatermark = ""
   objPDF.PDFEndPage
   objPDF.PDFNewPage
   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
   objPDF.PDFSetDrawColor = vbWhite
   objPDF.PDFSetTextColor = vbBlack
   objPDF.PDFSetAlignement = ALIGN_CENTER
   objPDF.PDFSetBorder = BORDER_ALL
   objPDF.PDFSetFill = True

   objPDF.PDFCell " ", 15, 15, _
   objPDF.PDFGetPageWidth - 30, objPDF.PDFGetPageHeight - 20
   If PntPrv = "Y" Then
      objPDF.PDFWatermark = "Preview"
   Else
      objPDF.PDFWatermark = "NCML" '"NCMSL"
   End If
   
   objPDF.PDFSetFont FONT_ARIAL, 12, FONT_BOLD
   objPDF.PDFSetDrawColor = vbWhite
   objPDF.PDFSetTextColor = vbBlack
   objPDF.PDFSetAlignement = ALIGN_CENTER
   objPDF.PDFSetBorder = BORDER_ALL
   objPDF.PDFSetFill = True

   mLeft = mGlobalLeft
   mTxtTop = 30
   mHeight = mTxtTop + mRowHeight
   mTxtFromTop = mRowHeight + 2
End Sub

Private Sub InDataPledgeDetail()

'tmp = "SELECT  convert(varchar,TPC.PledgeID)+'~'+TPC.PledgeReqNo as PledgeID,TPC.PledgeDate,isnull(sum(PledgeQuantity),0) PledgeQuantity,isnull(sum(PledgeNoOfBags),0) PledgeBags,MB.BankName,MBr.BranchName"
tmp = "SELECT  TPC.Txn_PledgeCreation_ID,TPC.PledgeDate,isnull(sum(PledgeQuantity),0) PledgeQuantity,isnull(sum(PledgeNoOfBags),0) PledgeBags,MB.BankName,MBr.BranchName,isnull(TPC.NCMLFinance,0) 'NCMLFinance'"
tmp = tmp & " FROM  Txn_WHR_Detail TWD"
tmp = tmp & " left join Txn_PledgeCreation TPC on  TWD.WHRNo = TPC.WHRNo And TWD.SM_Prefix = TPC.SM_Prefix And TPC.ExchangeTypeID = 1"
tmp = tmp & " left Join Txn_PledgeCreation_GSL TPRG on TPRG.Txn_PledgeCreation_ID = TPC.Txn_PledgeCreation_ID And TPRG.SM_Prefix = TPC.SM_Prefix"
tmp = tmp & " left Join Mst_Bank MB on MB.BankID = TPC.BankID"
tmp = tmp & " left Join Mst_Branch MBr on MBr.BranchID = TPC.BranchID"
tmp = tmp & " Where TWD.WHRNo = " & RsTxn_WHRDetail!WHRNo & " And TPC.Txn_PledgeCreation_ID = " & IIf(IsNull(RsTxn_WHRDetail!Txn_PledgeCreation_ID), 0, RsTxn_WHRDetail!Txn_PledgeCreation_ID)
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
   CmdPledge.Enabled = True
   mNCMLFinance = 0
End If

End Sub

Private Sub IndataCARDetails()


TxtCARDate.Text = ""
If IsNull(RsTxn_WHRDetail!CARNo) = False And RsTxn_WHRDetail!CARNo <> 0 Then
   TxtCARNo.Text = ""
   
   tmp = "Select CARNo, CARDate from Txn_CAR Where CARNo=" & IIf(IsNull(RsTxn_WHRDetail!CARNo), 0, RsTxn_WHRDetail!CARNo)
   Call TmpTableCMPro
   If TmpRsCMPro.RecordCount > 0 Then
      TxtCARNo.Text = TmpRsCMPro!CARNo
      TxtCARDate.Text = Format(TmpRsCMPro!CARDate, "dd-MMM-yyyy")
   End If
Else
   If Val(TxtCARNo) > 0 Then
      tmp = "Select CARNo, CARDate from Txn_CAR Where CARNo=" & Val(TxtCARNo)
      Call TmpTableCMPro
      If TmpRsCMPro.RecordCount > 0 Then
         TxtCARNo.Text = TmpRsCMPro!CARNo
         TxtCARDate.Text = Format(TmpRsCMPro!CARDate, "dd-MMM-yyyy")
      End If
   End If
End If

End Sub

Private Sub MSHFlexGrid3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyC And Shift = vbCtrlMask Then
   With MSHFlexGrid3
        .Redraw = False              'Find range to copy
        min_row = .row:   max_row = .RowSel
        If min_row > max_row Then tmp = max_row: max_row = min_row: min_row = tmp
        min_col = .Col:   max_col = .ColSel
        If min_col > max_col Then tmp = max_col: max_col = min_col: min_col = tmp

        my_text = ""                   'Get the text
        For I = min_row To max_row
            For j = min_col To max_col
                my_text = my_text & .TextMatrix(I, j)
                If j < max_col Then my_text = my_text & vbTab
            Next j
            If I < max_row Then my_text = my_text & vbCrLf
        Next I
        .Redraw = True
   End With
   If my_text <> "" Then
      Clipboard.Clear
      Clipboard.SetText (my_text)
   Else
      Beep
      MsgBox "Please select some text.", vbExclamation, "Nothing selected!"
   End If
End If
End Sub

Private Sub ExportToCSV()

Dim Path, strFilName, strFilLocation As Variant
Path = App.Path
Gen_mTimeStamp = GetTimeStamp
strFilName = "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_WHRDetails.csv"
strFilLocation = Path + strFilName

If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation

Open strFilLocation For Output As #1

For I = 0 To MSHFlexGrid1.Rows - 1
    tmp = ""
    For C = 0 To MSHFlexGrid1.Cols - 1
        If tmp = "" Then
           tmp = tmp & MSHFlexGrid1.TextMatrix(I, C)
        Else
           tmp = tmp & "," & MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
    Print #1, tmp
Next

MsgBox "CSV File Created!!!"

Close #1

End Sub

Private Sub CmbSSCMFlag_LostFocus()
If CmbSSCMFlag.Text = "" Then Exit Sub

If LCase(CmbSSCMFlag.Text) <> "yes" And LCase(CmbSSCMFlag.Text) <> "no" Then
   MsgBox "Invalid Selection!!!"
   CmbSSCMFlag.SetFocus
   Exit Sub
End If

End Sub


Private Sub CmbGodownID_GotFocus()
If CmbCMPID.Text = "" Then
   MsgBox "Select CMP first !!! "
   CmbCMPID.SetFocus
   Exit Sub
End If
tmp = CmbGodownID.Text
Call TableMst_Godown("Where CloseDate is null And CMPID = " & mCMPID)
CmbGodownID.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Godown.RecordCount
    CmbGodownID.AddItem RsMst_Godown!GodownNo
    RsMst_Godown.MoveNext
Next
CmbGodownID.Text = tmp
End Sub
Private Sub CmbGodownID_LostFocus()
If CmbGodownID.Text = "" Then
   mGodownID = Null
   mLotWiseWHR = Null
   Call Grid_HeadNew
   Call Grid_Head
   Exit Sub
End If

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownID.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid selection "
   CmbGodownID.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
TxtAgreementEndDate = Format(RsMst_Godown!AgreementEndDate, "dd-mmm-yyyy")
mLicenseID = RsMst_Godown!LicenseID
ans = MsgBox("Click 'Yes' to generate WHR against lots which are having 'O' flag with zero WHR Bags and click 'No' to generate WHR against lots having 'E'or 'O' flag having partial WHR generated ", vbYesNo) '  WHR to be generated for New Lots or Old Lots . Click 'Yes for New' and 'No' for Old WHR pattern", vbYesNo)
If ans = vbYes Then
   mLotWiseWHR = "1"
   Frame2.Visible = False
   MSHFlexGrid2.Visible = False
   Call Grid_Head
   MSHFlexGrid4.Visible = True
   Call Grid_HeadNew
Else
   mLotWiseWHR = "0"
   Frame2.Visible = True
   MSHFlexGrid2.Visible = True
   Call Grid_Head
   MSHFlexGrid4.Visible = False
   Call Grid_HeadNew
End If

Call GetGridData(mExchangeTypeID, mCommodityID, mClientIDRow)

End Sub

Public Sub Grid_HeadNew()

MSHFlexGrid4.Clear
MSHFlexGrid4.Cols = 10
MSHFlexGrid4.Rows = 3
MSHFlexGrid4.FixedRows = 2

MSHFlexGrid4.TextMatrix(1, 0) = "Selection"
MSHFlexGrid4.TextMatrix(1, 1) = "Stack No "
MSHFlexGrid4.TextMatrix(1, 2) = "Lot No"
MSHFlexGrid4.TextMatrix(1, 3) = "Available No Of Bags"
MSHFlexGrid4.TextMatrix(1, 4) = "Available Electronic Weight"
MSHFlexGrid4.TextMatrix(1, 5) = "WHR No Of Bags"
MSHFlexGrid4.TextMatrix(1, 6) = "WHR Quantity"
MSHFlexGrid4.TextMatrix(1, 7) = "Delivered Qty"
MSHFlexGrid4.TextMatrix(1, 8) = "GSL ID"
MSHFlexGrid4.TextMatrix(1, 9) = "Prefix"

MSHFlexGrid4.ColWidth(0) = 1000
MSHFlexGrid4.ColWidth(1) = 1400
MSHFlexGrid4.ColWidth(2) = 1400
MSHFlexGrid4.ColWidth(3) = 1400
MSHFlexGrid4.ColWidth(4) = 1400
MSHFlexGrid4.ColWidth(5) = 1400
MSHFlexGrid4.ColWidth(6) = 1400
MSHFlexGrid4.ColWidth(7) = 2000
MSHFlexGrid4.ColWidth(8) = 1400
MSHFlexGrid4.ColWidth(9) = 1400

MSHFlexGrid4.RowHeight(1) = 500

End Sub

Private Sub GetGridInDataLotWise()

k = 2

tmp = " Select * from Mst_GSL Where WHRNO = " & TxtWHRNo & " And Status not in ('Z') "
Call TmpTable

For I = 1 To TmpRs.RecordCount
    MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + 1
    MSHFlexGrid4.TextMatrix(k, 1) = TmpRs!StackNo
    MSHFlexGrid4.TextMatrix(k, 2) = TmpRs!LotNo
    MSHFlexGrid4.TextMatrix(k, 3) = 0
    MSHFlexGrid4.TextMatrix(k, 4) = 0
    MSHFlexGrid4.TextMatrix(k, 5) = TmpRs!WHRBags
    MSHFlexGrid4.TextMatrix(k, 6) = TmpRs!WHRWeight
    MSHFlexGrid4.TextMatrix(k, 7) = IIf(IsNull(TmpRs!CDTFWeight), 0, TmpRs!CDTFWeight)
    MSHFlexGrid4.TextMatrix(k, 8) = TmpRs!Mst_GSL_ID
    MSHFlexGrid4.TextMatrix(k, 9) = TmpRs!SM_Prefix
    MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(k, 5))
    MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(k, 6))
    MSHFlexGrid4.row = k
    MSHFlexGrid4.Col = 0
    MSHFlexGrid4.CellFontName = "Wingdings"
    MSHFlexGrid4.CellFontSize = 12
    MSHFlexGrid4.CellAlignment = flexAlignCenterCenter
    MSHFlexGrid4.Text = strChecked
    k = k + 1
    TmpRs.MoveNext
Next

End Sub
