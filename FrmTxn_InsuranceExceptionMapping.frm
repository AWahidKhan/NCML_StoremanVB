VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_InsuranceExceptionMapping 
   Caption         =   "Insurance Exception Mapping"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5085
      Left            =   1080
      TabIndex        =   33
      Top             =   6240
      Visible         =   0   'False
      Width           =   14355
      Begin VB.TextBox TxtSIEMID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12240
         MaxLength       =   12
         TabIndex        =   52
         Top             =   840
         Width           =   1500
      End
      Begin VB.ComboBox CmbSLocation 
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
         ItemData        =   "FrmTxn_InsuranceExceptionMapping.frx":0000
         Left            =   3210
         List            =   "FrmTxn_InsuranceExceptionMapping.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   832
         Width           =   3000
      End
      Begin VB.ComboBox CmbSClient 
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
         Left            =   6240
         Sorted          =   -1  'True
         TabIndex        =   50
         Top             =   832
         Width           =   2940
      End
      Begin VB.ComboBox CmbSCommodity 
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
         TabIndex        =   49
         Top             =   832
         Width           =   2940
      End
      Begin VB.OptionButton OptEndDt 
         Caption         =   "End Date"
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
         Left            =   1800
         TabIndex        =   48
         Top             =   240
         Width           =   1635
      End
      Begin VB.OptionButton OptStartDt 
         Caption         =   "Start Date"
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
         TabIndex        =   47
         Top             =   240
         Width           =   1395
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
         Left            =   13260
         TabIndex        =   35
         Top             =   1440
         Width           =   1000
      End
      Begin VB.CommandButton CmdSearch 
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
         Height          =   350
         Left            =   12240
         TabIndex        =   34
         Top             =   1440
         Width           =   1000
      End
      Begin MSComCtl2.DTPicker TxtSEndDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   825
         Width           =   1440
         _ExtentX        =   2540
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   50462723
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker TxtSStartDate 
         Height          =   375
         Left            =   240
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   825
         Width           =   1410
         _ExtentX        =   2487
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   50462723
         CurrentDate     =   36494
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2895
         Left            =   120
         TabIndex        =   38
         Top             =   1920
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   5106
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
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   10110
         TabIndex        =   46
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   7110
         TabIndex        =   45
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label23 
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
         Left            =   4328
         TabIndex        =   44
         Top             =   615
         Width           =   765
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
         Left            =   9480
         TabIndex        =   43
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label Label28 
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
         Left            =   7080
         TabIndex        =   42
         Top             =   1440
         Width           =   2415
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
         Left            =   458
         TabIndex        =   41
         Top             =   615
         Width           =   975
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
         Left            =   1973
         TabIndex        =   40
         Top             =   615
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "IEMID "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12690
         TabIndex        =   39
         Top             =   600
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   120
      TabIndex        =   24
      Top             =   5040
      Width           =   12375
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1935
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   3413
         _Version        =   393216
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
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   12375
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
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   2295
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
         Left            =   2535
         TabIndex        =   6
         Top             =   240
         Width           =   2295
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
         Left            =   4830
         TabIndex        =   7
         Top             =   240
         Width           =   2295
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
         Left            =   7125
         TabIndex        =   8
         Top             =   240
         Width           =   2295
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
         Left            =   9420
         TabIndex        =   29
         Top             =   240
         Width           =   2295
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
         Left            =   240
         TabIndex        =   10
         Top             =   615
         Width           =   2295
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
         Left            =   2535
         TabIndex        =   28
         Top             =   615
         Width           =   2295
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
         Left            =   4830
         TabIndex        =   27
         Top             =   615
         Width           =   2295
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
         Left            =   7125
         TabIndex        =   26
         Top             =   615
         Width           =   2295
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
         Height          =   375
         Left            =   9420
         TabIndex        =   9
         Top             =   615
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3810
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12375
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
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1800
         Width           =   10620
      End
      Begin VB.ComboBox CmbNonNcdexClient 
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
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1320
         Width           =   10620
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
         Left            =   1320
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   10620
      End
      Begin VB.TextBox TxtClient 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   10620
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2730
         Width           =   10590
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3225
         Width           =   10590
      End
      Begin VB.TextBox TxtIEMID 
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   1260
      End
      Begin MSComCtl2.DTPicker TxtStartDate 
         Height          =   360
         Left            =   1320
         TabIndex        =   4
         Top             =   2265
         Width           =   1275
         _ExtentX        =   2249
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   50462723
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtEndDate 
         Height          =   330
         Left            =   4200
         TabIndex        =   5
         Top             =   2280
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
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
         CustomFormat    =   "MMM-yyyy"
         Format          =   50462723
         CurrentDate     =   39850
         MinDate         =   38353
      End
      Begin VB.TextBox TxtLocation 
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
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   75
         TabIndex        =   31
         Top             =   840
         Width           =   10620
      End
      Begin VB.TextBox TxtCommodity 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1800
         Visible         =   0   'False
         Width           =   10620
      End
      Begin VB.Label Label3 
         Caption         =   "End Date"
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
         Left            =   3240
         TabIndex        =   23
         Top             =   2318
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Start Date"
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
         TabIndex        =   22
         Top             =   2295
         Width           =   975
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1853
         Width           =   1200
      End
      Begin VB.Label Label19 
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
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1373
         Width           =   1200
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
         Left            =   120
         TabIndex        =   18
         Top             =   2790
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
         Left            =   120
         TabIndex        =   17
         Top             =   3285
         Width           =   975
      End
      Begin VB.Label LblLocationID 
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
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   893
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
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
         TabIndex        =   15
         Top             =   390
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmTxn_InsuranceExceptionMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mSM_Prefix, mCommodityID, mClientIDRow, mSLocationID, msSM_prefix, mSClientIDRow, mSCommodityID As Variant
Dim Edit_Flg As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"

Frame1.Enabled = True
Frame2.Enabled = False

CmbLocation.Visible = True
TxtLocation.Visible = False

CmbCommodity.Visible = True
TxtCommodity.Visible = False

CmbNonNcdexClient.Visible = True
TxtClient.Visible = False

TxtIEMID.Text = ""
CmbLocation.Text = ""
CmbCommodity.Text = ""
CmbNonNcdexClient.Text = ""
'TxtRemark.Text = ""
TxtCreated = ""
TxtUpdated = ""

TxtStartDate.Value = Date
TxtEndDate.Value = Date


TxtStartDate.Value = DateSerial(Year(TxtStartDate.Value), Month(TxtStartDate.Value), 1)
TxtEndDate.Value = DateSerial(Year(TxtEndDate.Value), Month(TxtEndDate.Value) + 1, 0)


Call Grid_Head

Call GButtonLock(Me, False)

End Sub

Private Sub CmbCommodity_GotFocus()
If CmbLocation.Text = "" Then
   MsgBox "Blank Location not allowed "
   CmbLocation.SetFocus
   Exit Sub
End If
Call TableMst_Commodity
tmp = "select MstLCM.LocationID, MstLCM.CommodityID, MC.Commodity  From Mst_LocationCommodityMapping as MstLCM"
tmp = tmp & " Inner Join Mst_Commodity as MC on MC.CommodityID = MstLCM.CommodityID"
tmp = tmp & " Inner Join Mst_Location as ML on ML.LocationID = MstLCM.LocationID"
tmp = tmp & " Where MstLCM.LocationID =" & mLocationID
Call Tmp1Table

tmp = CmbCommodity.Text

CmbCommodity.Clear
If TmpRs1.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To TmpRs1.RecordCount
 CmbCommodity.AddItem TmpRs1!Commodity
 TmpRs1.MoveNext
Next
CmbCommodity.Text = tmp
End Sub


Private Sub CmbCommodity_LostFocus()
If CmbCommodity.Text = "" Then
   mCommodityID = Null
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodity.SetFocus
   Exit Sub
End If

mCommodityID = RsMst_Commodity!CommodityID
End Sub


Private Sub Grid_Head()
With MSHFlexGrid1
     .Clear
     .Rows = 2
     .Cols = 8
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .WordWrap = True
     '.RowHeight(1) = 700
     .TextMatrix(0, 0) = "ID"
     .TextMatrix(0, 1) = "Location Name"
     .TextMatrix(0, 2) = "Client Name"
     .TextMatrix(0, 3) = "Commodity"
     .TextMatrix(0, 4) = "Start Date"
     .TextMatrix(0, 5) = "End Date"
     .TextMatrix(0, 6) = "Concurrency"
     .TextMatrix(0, 7) = "Created By"
    
     .ColWidth(0) = 800
     .ColWidth(1) = 1500
     .ColWidth(2) = 3000
     .ColWidth(3) = 1500
     .ColWidth(4) = 1200
     .ColWidth(5) = 1200
     .ColWidth(6) = 1500
     .ColWidth(7) = 1200

End With
End Sub
Private Sub Grid_Head1()
With MSHFlexGrid2
     .Clear
     .Rows = 2
     .Cols = 6
     .AllowUserResizing = flexResizeBoth
     .Font.Size = 10
     .WordWrap = True
     '.RowHeight(1) = 700
     .TextMatrix(0, 0) = "ID"
     .TextMatrix(0, 1) = "Location Name"
     .TextMatrix(0, 2) = "Client Name"
     .TextMatrix(0, 3) = "Commodity"
     .TextMatrix(0, 4) = "Start Date"
     .TextMatrix(0, 5) = "End Date"
'     .TextMatrix(0, 6) = "Concurrency"
'     .TextMatrix(0, 7) = "Created By"
    
     .ColWidth(0) = 800
     .ColWidth(1) = 2500
     .ColWidth(2) = 4500
     .ColWidth(3) = 3000
     .ColWidth(4) = 1500
     .ColWidth(5) = 1500


End With
End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found!!! "
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
   mLocationID = Null
   'Call ClearFrame
   Call Grid_Head
'   Frame4.Enabled = False
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbLocationID.SetFocus
   Exit Sub
End If
If mLocationID <> RsMst_Location!LocationID Then
   Call Grid_Head
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix

End Sub




Private Sub CmbLocation_GotFocus()
tmp = CmbLocation.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocation.Text = tmp
End Sub


Private Sub CmbLocation_LostFocus()
If CmbLocation.Text = "" Then
   mLocationID = 0
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbLocation.SetFocus
   Exit Sub
End If
If mLocationID <> RsMst_Location!LocationID Then
   'CmbNonNcdexClient.Text = ""
   CmbCommodity.Text = ""
End If
mLocationID = RsMst_Location!LocationID
mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmbSClient_GotFocus()
tmp = CmbSClient.Text
Call TableMst_Client("where ExchangeTypeId = 1 ") '" & mExchangeTypeID)
CmbSClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbSClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbSClient.Text = tmp
End Sub

Private Sub CmbSClient_LostFocus()
If CmbSClient.Text = "" Then
   mSClientIDRow = Null
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbSClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbSClient.SetFocus
   Exit Sub
End If
mSClientIDRow = RsMst_Client!ClientIDRow
End Sub

Private Sub CmbSCommodity_GotFocus()
'If CmbSLocation.Text = "" Then
'   MsgBox "Blank Location not allowed "
'   CmbSLocation.SetFocus
'   Exit Sub
'End If

Call TableMst_Commodity

'tmp = "select MstLCM.LocationID, MstLCM.CommodityID, MC.Commodity  From Mst_LocationCommodityMapping as MstLCM"
'tmp = tmp & " Inner Join Mst_Commodity as MC on MC.CommodityID = MstLCM.CommodityID"
'tmp = tmp & " Inner Join Mst_Location as ML on ML.LocationID = MstLCM.LocationID"
'tmp = tmp & " Where MstLCM.LocationID =" & mSLocationId
'Call Tmp1Table

tmp = CmbSCommodity.Text

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
   mSCommodityID = Null
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

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text

If Gen_WcLocation = "" Then
   Call TableMst_Location("")
Else
   Call TableMst_Location("where " & Gen_WcLocation)
End If

CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp
End Sub

Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then
   mSLocationID = 0
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!! "
   CmbSLocation.SetFocus
   Exit Sub
End If
If mSLocationID <> RsMst_Location!LocationID Then
   'CmbNonNcdexClient.Text = ""
   CmbCommodity.Text = ""
End If
mSLocationID = RsMst_Location!LocationID
msSM_prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmdSearch_Click()
Dim wc As String
Dim mFromDt As Date
Dim mToDt As Date
wc = ""
Call Grid_Head1
TxtSStartDate.Value = DateSerial(Year(TxtSStartDate.Value), Month(TxtSStartDate.Value), 1)
TxtSEndDate.Value = DateSerial(Year(TxtSEndDate.Value), Month(TxtSEndDate.Value) + 1, 0)

tmp = "Select Txn_IEM.IEMID, Txn_IEM.LocationID, ML.SM_Prefix, ML.LocationName, Txn_IEM.CommodityID, MCO.Commodity, Txn_IEM.ClientIDRow, MC.ClientName, Txn_IEM.StartDate, Txn_IEM.EndDate, Txn_IEM.Concurrency, Txn_IEM.CreatedBy "
tmp = tmp & " From  Txn_InsuranceExceptionMapping as Txn_IEM"
tmp = tmp & " Inner Join  Mst_Location as ML on ML.LocationID = Txn_IEM.LocationID"
tmp = tmp & " Left Join Mst_Client as MC on MC.ClientIDRow = Txn_IEM.ClientIDRow"
tmp = tmp & " Left Join Mst_Commodity as MCO on MCO.CommodityID = Txn_IEM.CommodityID"
''tmp = tmp & " Where Txn_IEM.StartDate > = '" & Format(TxtSStartDate.Value, Gen_DatFormat) & "' And Txn_IEM.EndDate <='" & Format(TxtSEndDate.Value, Gen_DatFormat) & "'"
If OptStartDt.Value = True Then
   'mFromDt
   wc = " Where Txn_IEM.StartDate Between '" & Format(TxtSStartDate.Value, Gen_DatFormat) & "' And '" & Format(TxtSEndDate, Gen_DatFormat) & "'"
  ' Wc = " Where Txn_IEM.StartDate Between " & TxtSStartDate.Value & " And " & TxtSEndDate & " "
ElseIf OptEndDt.Value = True Then
   wc = " Where Txn_IEM.EndDate Between '" & Format(TxtSStartDate.Value, Gen_DatFormat) & "' And '" & Format(TxtSEndDate, Gen_DatFormat) & "'"
End If
If CmbSLocation.Text <> "" Then
   wc = wc & " And Txn_IEM.LocationID =" & mSLocationID
End If
If CmbSClient.Text <> "" Then
   wc = wc & " And Txn_IEM.ClientIDRow =" & mSClientIDRow
End If
If CmbSCommodity.Text <> "" Then
   wc = wc & " And Txn_IEM.CommodityID =" & mSCommodityID
End If
If TxtSIEMID.Text <> "" Then
   wc = wc & " And Txn_IEM.IEMID =" & Val(TxtSIEMID)
End If

tmp = tmp & wc
tmp = tmp & " Order by  Txn_IEM.IEMID"
Call Tmp2Table

If Not TmpRs2.EOF Then
   Label22.Caption = TmpRs2.RecordCount
   With MSHFlexGrid2
        For I = 1 To TmpRs2.RecordCount
            .TextMatrix(I, 0) = TmpRs2!IEMID
            .TextMatrix(I, 1) = TmpRs2!LocationName
            .TextMatrix(I, 2) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
            .TextMatrix(I, 3) = IIf(IsNull(TmpRs2!Commodity), "", TmpRs2!Commodity)
            .TextMatrix(I, 4) = Format(TmpRs2!StartDate, "dd-MMM-yyyy")
            .TextMatrix(I, 5) = Format(TmpRs2!EndDate, "dd-MMM-yyyy")
            TmpRs2.MoveNext
            .Rows = .Rows + 1
        Next
   End With
End If

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsTxn_InsuranceExceptionMapping.Delete
   RsTxn_InsuranceExceptionMapping.Update
   If RsTxn_InsuranceExceptionMapping.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_InsuranceExceptionMapping.MoveNext
   If RsTxn_InsuranceExceptionMapping.EOF() Then
      RsTxn_InsuranceExceptionMapping.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_State.CancelUpdate
End Sub


Private Sub EditCmd_Click()

Edit_Flg = "E"

Call GButtonLock(Me, False)

Frame1.Enabled = True
'TxtHologramDate.Enabled = False
Frame2.Enabled = True
CmbLocation.Text = TxtLocation.Text
CmbLocation.Visible = True
TxtLocation.Visible = False

CmbCommodity.Text = TxtCommodity.Text
CmbCommodity.Visible = True
TxtCommodity.Visible = False

CmbNonNcdexClient.Text = TxtClient.Text
CmbNonNcdexClient.Visible = True
TxtClient.Visible = False

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_InsuranceExceptionMapping.MoveFirst
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
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
TxtStartDate.Value = Date
TxtEndDate.Value = Date
TxtSStartDate.Value = Date
TxtSEndDate.Value = Date
OptStartDt.Value = True
Call Grid_Head1
Call TableTxn_InsuranceExceptionMapping("  Order by IEMID ")
If RsTxn_InsuranceExceptionMapping.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "FrmTxn_InsuranceExceptionMapping.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "FrmTxn_InsuranceExceptionMapping.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
'ProgressBar1.Value = 0
'ProgressBar1.Max = MSHFlexGrid2.Rows

For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For C = 1 To MSHFlexGrid2.Cols - 1
        If InStr(MSHFlexGrid2.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C) = Format(MSHFlexGrid2.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C) = MSHFlexGrid2.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C) = MSHFlexGrid2.TextMatrix(I, C)
        End If
        MSHFlexGrid2.row = I
        MSHFlexGrid2.Col = C
                
        If MSHFlexGrid2.CellBackColor = vbRed Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbRed
        ElseIf MSHFlexGrid2.CellBackColor = vbGreen Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbGreen
        ElseIf MSHFlexGrid2.CellBackColor = vbYellow Then
           oWS.Range(oWS.Cells(mRow, C), oWS.Cells(mRow, C)).Interior.Color = vbYellow
        End If
              
    Next
'    ProgressBar1.Value = ProgressBar1.Value + 1
Next
'ProgressBar1.Value = ProgressBar1.Max
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text
Call TableMst_Client("where ExchangeTypeId = 1 ") '" & mExchangeTypeID)
CmbNonNcdexClient.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbNonNcdexClient.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbNonNcdexClient.Text = tmp
End Sub
Private Sub CmbNonNcdexClient_LostFocus()
If CmbNonNcdexClient.Text = "" Then
   mClientIDRow = Null
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbNonNcdexClient.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientIDRow

End Sub

Private Sub LastCmd_Click()
RsTxn_InsuranceExceptionMapping.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_InsuranceExceptionMapping.MoveFirst
RsTxn_InsuranceExceptionMapping.Find "IEMID= " & Val(MSHFlexGrid1.TextMatrix(I, 0)) & ""
If Not RsTxn_InsuranceExceptionMapping.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub MSHFlexGrid2_Click()
Dim I As Variant
I = MSHFlexGrid2.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_InsuranceExceptionMapping.MoveFirst
RsTxn_InsuranceExceptionMapping.Find "IEMID= " & Val(MSHFlexGrid2.TextMatrix(I, 0)) & ""
If Not RsTxn_InsuranceExceptionMapping.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_InsuranceExceptionMapping.MoveNext
If RsTxn_InsuranceExceptionMapping.EOF Then
   RsTxn_InsuranceExceptionMapping.MoveLast
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
RsTxn_InsuranceExceptionMapping.MovePrevious
If RsTxn_InsuranceExceptionMapping.BOF Then
   RsTxn_InsuranceExceptionMapping.MoveFirst
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
If CmbLocation.Text = "" Then
   MsgBox "Select Location Name!!"
   CmbLocation.SetFocus
   Exit Sub
End If

If CmbNonNcdexClient.Text = "" And CmbCommodity.Text = "" Then
   MsgBox "Blank Client Or Commodity not allowed!!!"
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If

If CmbNonNcdexClient.Text <> "" And CmbCommodity.Text <> "" Then
   MsgBox "Either Client or Commodity Should be Selected "
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If


If TxtStartDate.Value > TxtEndDate.Value Then
   MsgBox "'From Date' Must be less than 'To Date'!!!"
   TxtStartDate.SetFocus
   Exit Sub
End If


'If CDate(Format(TxtStartDate.Value, Gen_DatFormat)) > Date Then
'    MsgBox "Future Date not Allowed!!!"
'    TxtStartDate.SetFocus
'    Exit Sub
'End If
'If CDate(Format(TxtEndDate.Value, Gen_DatFormat)) > Date Then
'    MsgBox "Future Date not Allowed!!!"
'    TxtEndDate.SetFocus
'    Exit Sub
'End If
   

'tmp = "Select * From Txn_InsuranceExceptionMapping Where LocationID =" & mLocationID & " And ClientIDRow =" & IIf(IsNull(mClientIDRow), "Null", mClientIDRow) & " And CommodityID =" & IIf(IsNull(mCommodityID), "Null", mCommodityID)
tmp = "Select * From Txn_InsuranceExceptionMapping Where LocationID =" & mLocationID & " "
'--And (ClientIDRow =" & IIf(IsNull(mClientIDRow), "Null", mClientIDRow) & " And CommodityID =" & IIf(IsNull(mCommodityID), "Null", mCommodityID) & ")"
'if mClientIDRow <> "Null" Then
If CmbNonNcdexClient.Text <> "" Then
   tmp = tmp & "And ClientIDRow =" & mClientIDRow
'ElseIf mCommodityID <> "Null" Then
End If
If CmbCommodity <> "" Then
   tmp = tmp & "And CommodityID =" & mCommodityID
 End If
tmp = tmp & " And (( '" & CDate(Format(TxtStartDate.Value, Gen_DatFormat)) & "'" & " >= StartDate And "
tmp = tmp & "'" & CDate(Format(TxtStartDate.Value, Gen_DatFormat)) & "'" & "  <= EndDate) Or"
tmp = tmp & "'" & CDate(Format(TxtEndDate.Value, Gen_DatFormat)) & "'" & "  >= StartDate And"
tmp = tmp & "'" & CDate(Format(TxtEndDate.Value, Gen_DatFormat)) & "'" & "  <= EndDate Or"
tmp = tmp & " (" & "'" & CDate(Format(TxtStartDate.Value, Gen_DatFormat)) & "'" & "  <= StartDate And"
tmp = tmp & "'" & CDate(Format(TxtEndDate.Value, Gen_DatFormat)) & "'" & "  >= EndDate))"
Call Tmp4Table
If TmpRs4.RecordCount > 0 Then
   If CmbNonNcdexClient <> "" Then
      MsgBox "Data For Selected Location and Client For Given Data Range " & Format(TxtStartDate.Value, Gen_DatFormat) & " And " & Format(TxtEndDate.Value, Gen_DatFormat) & " Already Exists"
   Else
      MsgBox "Data For Selected Location and Commodity For Given Data Range " & Format(TxtStartDate.Value, Gen_DatFormat) & " And " & Format(TxtEndDate.Value, Gen_DatFormat) & " Already Exists"
   End If
   Exit Sub
End If
'If Not TmpRs4.EOF Then
'   For i = 1 To TmpRs4.RecordCount
'       If (CDate(Format(TxtStartDate.Value, "MM/dd/yyyy")) >= CDate(TmpRs4!StartDate) And CDate(Format(TxtStartDate.Value, "MM/dd/yyyy")) <= CDate(TmpRs4!EndDate)) Or CDate(Format(TxtEndDate.Value, "MM/dd/yyyy")) >= CDate(TmpRs4!StartDate) And CDate(Format(TxtEndDate.Value, "MM/dd/yyyy")) <= CDate(TmpRs4!EndDate) Or (CDate(Format(TxtStartDate.Value, "MM/dd/yyyy")) <= CDate(TmpRs4!StartDate) And CDate(Format(TxtEndDate.Value, "MM/dd/yyyy")) >= CDate(TmpRs4!EndDate)) Then
'          MsgBox "Can not Enter"
'          Exit Sub
'       End If
'       TmpRs4.MoveNext
'   Next
'End If

If Edit_Flg = "A" Then
   RsTxn_InsuranceExceptionMapping.AddNew
   RsTxn_InsuranceExceptionMapping!IEMID = getIEMID
   RsTxn_InsuranceExceptionMapping!G_UID = GetGUID
   TxtIEMID = RsTxn_InsuranceExceptionMapping!IEMID
Else
   RsTxn_InsuranceExceptionMapping.MoveFirst
   RsTxn_InsuranceExceptionMapping.Find "IEMID = " & TxtIEMID
End If
RsTxn_InsuranceExceptionMapping!LocationID = mLocationID

If CmbNonNcdexClient.Text <> "" Then
   RsTxn_InsuranceExceptionMapping!ClientIDRow = mClientIDRow
End If

If CmbCommodity.Text <> "" Then
   RsTxn_InsuranceExceptionMapping!CommodityID = mCommodityID
End If
RsTxn_InsuranceExceptionMapping!StartDate = Format(TxtStartDate.Value, Gen_DatFormat)
RsTxn_InsuranceExceptionMapping!EndDate = Format(TxtEndDate.Value, Gen_DatFormat)
If IsNull(RsTxn_InsuranceExceptionMapping!CreatedBy) = True Or RsTxn_InsuranceExceptionMapping!CreatedBy = "" Then
   RsTxn_InsuranceExceptionMapping!CreatedBy = Gen_UserLoginID
   RsTxn_InsuranceExceptionMapping!CreatedDate = Now
Else
   RsTxn_InsuranceExceptionMapping!UpdatedBy = Gen_UserLoginID
   RsTxn_InsuranceExceptionMapping!UpdatedDate = Now
End If
RsTxn_InsuranceExceptionMapping.Update
RsTxn_InsuranceExceptionMapping.Requery

RsTxn_InsuranceExceptionMapping.MoveFirst
RsTxn_InsuranceExceptionMapping.Find "IEMID = " & TxtIEMID
If RsTxn_InsuranceExceptionMapping.EOF Then
   RsTxn_InsuranceExceptionMapping.MoveFirst
End If
Call GButtonLock(Me, True)
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_InsuranceExceptionMapping.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_InsuranceExceptionMapping.MoveFirst
   RsTxn_InsuranceExceptionMapping.Find "IEMID  = " & Val(TxtIEMID) & ""
   If RsTxn_InsuranceExceptionMapping.EOF Then
      RsTxn_InsuranceExceptionMapping.MoveLast
   End If
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 200
Frame0.Visible = True

MSHFlexGrid2.Width = Frame0.Width - 200
MSHFlexGrid2.Height = Frame0.Height - 2500

RsTxn_InsuranceExceptionMapping.MoveFirst
'RsTxn_InsuranceExceptionMapping.Filter = ""
End Sub


Private Sub SStartDate_LostFocus()
TxtStartDate.Value = DateSerial(Year(TxtStartDate.Value), Month(TxtStartDate.Value), 1)
End Sub


Private Sub TxtSEndDate_LostFocus()
TxtSEndDate.Value = DateSerial(Year(TxtSEndDate.Value), Month(TxtSEndDate.Value) + 1, 0)
End Sub


Private Sub TxtSStartDate_LostFocus()
TxtSStartDate.Value = DateSerial(Year(TxtSStartDate.Value), Month(TxtSStartDate.Value), 1)
End Sub


Private Sub TxtEndDate_LostFocus()
TxtEndDate.Value = DateSerial(Year(TxtEndDate.Value), Month(TxtEndDate.Value) + 1, 0)
End Sub

Private Sub TxtStartDate_LostFocus()
TxtStartDate.Value = DateSerial(Year(TxtStartDate.Value), Month(TxtStartDate.Value), 1)
End Sub
Public Sub Indata()

Frame1.Enabled = False
Frame2.Enabled = True

CmbLocation.Visible = False
TxtLocation.Visible = True


CmbNonNcdexClient.Visible = False
TxtClient.Visible = True

CmbCommodity.Visible = False
TxtCommodity.Visible = True

TxtIEMID = RsTxn_InsuranceExceptionMapping!IEMID

mLocationID = RsTxn_InsuranceExceptionMapping!LocationID
TxtLocation.Text = GetLocationName(RsTxn_InsuranceExceptionMapping!LocationID)

mClientIDRow = IIf(IsNull(RsTxn_InsuranceExceptionMapping!ClientIDRow), "", RsTxn_InsuranceExceptionMapping!ClientIDRow)
If mClientIDRow = "" Then
   TxtClient.Text = ""
Else
   TxtClient.Text = GetClientName(RsTxn_InsuranceExceptionMapping!ClientIDRow)
End If
'TxtClient.Text = IIf(IsNull(RsTxn_InsuranceExceptionMapping!ClientIDRow), "", GetClientName(RsTxn_InsuranceExceptionMapping!ClientIDRow))

'mCommodityID = RsTxn_InsuranceExceptionMapping!CommodityID
mCommodityID = IIf(IsNull(RsTxn_InsuranceExceptionMapping!CommodityID), "", RsTxn_InsuranceExceptionMapping!CommodityID)
If mCommodityID = "" Then
   TxtCommodity.Text = ""
Else
   TxtCommodity.Text = GetCommodityName(RsTxn_InsuranceExceptionMapping!CommodityID)
End If
'TxtCommodity.Text = IIf(IsNull(RsTxn_InsuranceExceptionMapping!CommodityID), "", GetCommodityName(RsTxn_InsuranceExceptionMapping!CommodityID))

TxtStartDate.Value = Format(RsTxn_InsuranceExceptionMapping!StartDate, "dd/MMM/yyyy")
TxtEndDate.Value = Format(RsTxn_InsuranceExceptionMapping!EndDate, "dd/MMM/yyyy")


TxtCreated = IIf(IsNull(RsTxn_InsuranceExceptionMapping!CreatedBy), "", RsTxn_InsuranceExceptionMapping!CreatedBy) & " :- " & IIf(IsNull(RsTxn_InsuranceExceptionMapping!CreatedDate), "", RsTxn_InsuranceExceptionMapping!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_InsuranceExceptionMapping!UpdatedBy), "", RsTxn_InsuranceExceptionMapping!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_InsuranceExceptionMapping!UpdatedDate), "", RsTxn_InsuranceExceptionMapping!UpdatedDate)


tmp = "Select Txn_IEM.IEMID, Txn_IEM.LocationID, ML.SM_Prefix,  ML.LocationName, Txn_IEM.CommodityID, MCO.Commodity, Txn_IEM.ClientIDRow, MC.ClientName, Txn_IEM.StartDate, Txn_IEM.EndDate, Txn_IEM.Concurrency, Txn_IEM.CreatedBy"
tmp = tmp & " From  Txn_InsuranceExceptionMapping as Txn_IEM"
tmp = tmp & " Inner Join  Mst_Location as ML on ML.LocationID = Txn_IEM.LocationID"
tmp = tmp & " Left Join Mst_Client as MC on MC.ClientIDRow = Txn_IEM.ClientIDRow"
tmp = tmp & " Left Join Mst_Commodity as MCO on MCO.CommodityID = Txn_IEM.CommodityID"
tmp = tmp & " Where Txn_IEM.LocationID=" & RsTxn_InsuranceExceptionMapping!LocationID
If IsNull(RsTxn_InsuranceExceptionMapping!ClientIDRow) = False Then
   tmp = tmp & " And Txn_IEM.ClientIDRow = " & RsTxn_InsuranceExceptionMapping!ClientIDRow
Else
   tmp = tmp & " And Txn_IEM.ClientIDRow is NULL"
End If
If IsNull(RsTxn_InsuranceExceptionMapping!CommodityID) = False Then
   tmp = tmp & " And Txn_IEM.CommodityID =" & RsTxn_InsuranceExceptionMapping!CommodityID
Else
   tmp = tmp & " And Txn_IEM.CommodityID is NULL"
End If
tmp = tmp & " And Txn_IEM.IEMID <>" & Val(TxtIEMID)
''And (Txn_IEM.ClientIDRow=" & IIf(IsNull(RsTxn_InsuranceExceptionMapping!ClientIDRow), 0, RsTxn_InsuranceExceptionMapping!ClientIDRow) & "  or Txn_IEM.CommodityID='" & IIf(IsNull(RsTxn_InsuranceExceptionMapping!CommodityID), 0, RsTxn_InsuranceExceptionMapping!CommodityID) & "' And Txn_IEM.IEMID <>" & Val(TxtIEMID) & ")"
tmp = tmp & " Order by Txn_IEM.IEMID"
Call Tmp3Table
Call Grid_Head
If TmpRs3.RecordCount > 0 Then
   TmpRs3.MoveFirst
'TmpRs3.Filter = "LocationID='" & RsTxn_InsuranceExceptionMapping!LocationID & "' And ClientIDRow=" & IIf(IsNull(RsTxn_InsuranceExceptionMapping!ClientIDRow), 0, RsTxn_InsuranceExceptionMapping!ClientIDRow) & "  And CommodityID='" & RsTxn_InsuranceExceptionMapping!CommodityID & "'"
   With MSHFlexGrid1
        For I = 1 To TmpRs3.RecordCount
            .TextMatrix(I, 0) = TmpRs3!IEMID
            .TextMatrix(I, 1) = TmpRs3!LocationName
            .TextMatrix(I, 2) = IIf(IsNull(TmpRs3!ClientName), "", TmpRs3!ClientName)
            .TextMatrix(I, 3) = IIf(IsNull(TmpRs3!Commodity), "", TmpRs3!Commodity)
            .TextMatrix(I, 4) = Format(TmpRs3!StartDate, "dd-MMM-yyyy")
            .TextMatrix(I, 5) = Format(TmpRs3!EndDate, "dd-MMM-yyyy")
            .TextMatrix(I, 6) = Format(TmpRs3!Concurrency, "dd-MMM-yyyy")
            .TextMatrix(I, 7) = TmpRs3!CreatedBy
            TmpRs3.MoveNext
            .Rows = .Rows + 1
        Next
   End With
End If

If RsTxn_InsuranceExceptionMapping.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub
Private Function getIEMID() As Double
Dim Retval As Double
tmp = "Select max(IEMID) from Txn_InsuranceExceptionMapping"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getIEMID = Retval
End Function
