VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMst_NCDEXAccreditedCapacity 
   Caption         =   "NCDEX Accredited Capacity"
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
      Height          =   3135
      Left            =   1560
      TabIndex        =   51
      Top             =   4200
      Visible         =   0   'False
      Width           =   14175
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
         ItemData        =   "FrmMst_NCDEXAccreditedCapacity.frx":0000
         Left            =   6720
         List            =   "FrmMst_NCDEXAccreditedCapacity.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   600
         Width           =   1425
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
         Left            =   3000
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   600
         Width           =   3600
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
         Left            =   11880
         TabIndex        =   22
         Top             =   720
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
         Left            =   12960
         TabIndex        =   23
         Top             =   720
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   120
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1155
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   3916
         _Version        =   393216
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
      Begin MSComCtl2.DTPicker TxtFromDate 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   115081217
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtToDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   115081217
         CurrentDate     =   39884
      End
      Begin VB.Label Label3 
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
         Left            =   7155
         TabIndex        =   58
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   315
         TabIndex        =   57
         Top             =   300
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   1860
         TabIndex        =   56
         Top             =   300
         Width           =   735
      End
      Begin VB.Label LblLocationID 
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
         Index           =   1
         Left            =   4418
         TabIndex        =   55
         Top             =   300
         Width           =   765
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
         Left            =   9120
         TabIndex        =   54
         Top             =   405
         Width           =   2370
      End
      Begin VB.Label lblTotalRecord 
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
         Left            =   11640
         TabIndex        =   53
         Top             =   420
         Width           =   90
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1110
      Left            =   120
      TabIndex        =   37
      Top             =   7980
      Width           =   14250
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
         Left            =   11370
         TabIndex        =   17
         Top             =   615
         Width           =   2820
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
         Left            =   11370
         TabIndex        =   12
         Top             =   195
         Width           =   2820
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
         Left            =   8550
         TabIndex        =   16
         Top             =   615
         Width           =   2820
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
         Left            =   8550
         TabIndex        =   11
         Top             =   195
         Width           =   2820
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
         Left            =   5730
         TabIndex        =   15
         Top             =   615
         Width           =   2820
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
         Left            =   5730
         TabIndex        =   10
         Top             =   195
         Width           =   2820
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
         Left            =   2895
         TabIndex        =   14
         Top             =   615
         Width           =   2820
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
         Left            =   90
         TabIndex        =   13
         Top             =   615
         Width           =   2820
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
         Left            =   90
         TabIndex        =   8
         Top             =   195
         Width           =   2820
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "&Save"
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
         Left            =   2895
         TabIndex        =   9
         Top             =   195
         Width           =   2820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3285
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   14220
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
         Height          =   1035
         Left            =   1800
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2160
         Width           =   12240
      End
      Begin VB.TextBox TxtAccreditedDescription 
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
         Left            =   1800
         MaxLength       =   250
         TabIndex        =   2
         Top             =   720
         Width           =   12240
      End
      Begin VB.TextBox TxtAccreditedID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   203
         Width           =   2055
      End
      Begin VB.TextBox TxtPreFix 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   375
         Left            =   13560
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker TxtAccreditedDate 
         Height          =   375
         Left            =   5520
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   203
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   115081217
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtStartDate 
         Height          =   375
         Left            =   9720
         TabIndex        =   0
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   115081217
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker TxtCloseDate 
         Height          =   375
         Left            =   12720
         TabIndex        =   1
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   115081217
         CurrentDate     =   39884
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
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1215
         Width           =   11760
      End
      Begin VB.CommandButton CmdLocationID 
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
         Left            =   13680
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1245
         Width           =   375
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
         ItemData        =   "FrmMst_NCDEXAccreditedCapacity.frx":0020
         Left            =   5520
         List            =   "FrmMst_NCDEXAccreditedCapacity.frx":002A
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1687
         Width           =   1425
      End
      Begin VB.TextBox TxtCapacity 
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
         Left            =   1800
         TabIndex        =   4
         Top             =   1680
         Width           =   2055
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1215
         Width           =   12240
      End
      Begin VB.TextBox TxtStatusFlag 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
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
         TabIndex        =   50
         Top             =   2557
         Width           =   945
      End
      Begin VB.Label LblISINDescription 
         Caption         =   "Accredited Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   90
         TabIndex        =   46
         Top             =   623
         Width           =   1800
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Close Date"
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
         Left            =   11400
         TabIndex        =   45
         Top             =   270
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Start Date"
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
         Left            =   8640
         TabIndex        =   44
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Accredited Date"
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
         Left            =   3960
         TabIndex        =   36
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2970
         TabIndex        =   35
         Top             =   3780
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Till Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5970
         TabIndex        =   34
         Top             =   3780
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label LblWHRDate 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   33
         Top             =   3780
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label LblWHRNo 
         AutoSize        =   -1  'True
         Caption         =   "Accredited ID"
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
         Left            =   90
         TabIndex        =   32
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label LblLocationID 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   31
         Top             =   1275
         Width           =   900
      End
      Begin VB.Label Label6 
         Caption         =   "SM_Prefix"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   13365
         TabIndex        =   30
         Top             =   1770
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   4755
         TabIndex        =   49
         Top             =   1747
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Capacity"
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
         Left            =   90
         TabIndex        =   47
         Top             =   1747
         Width           =   930
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4620
      Left            =   120
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3360
      Width           =   14250
      _ExtentX        =   25135
      _ExtentY        =   8149
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Godown Connected Details"
      TabPicture(0)   =   "FrmMst_NCDEXAccreditedCapacity.frx":0040
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSHFlexGrid2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ChkCheckAll"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Log History"
      TabPicture(1)   =   "FrmMst_NCDEXAccreditedCapacity.frx":005C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(2)=   "TxtUpdated"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "TxtCreated"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.CheckBox ChkCheckAll 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1095
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
         Left            =   -73695
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   540
         Width           =   12060
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
         Left            =   -73695
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1005
         Width           =   12090
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3675
         Left            =   120
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   840
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   6482
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
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
         TabIndex        =   43
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   -74760
         TabIndex        =   42
         Top             =   1065
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmMst_NCDEXAccreditedCapacity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLOcationID, Edit_Flg As Variant
Const strChecked = "þ"
Const strUnChecked = "q"

Private Sub AddCmd_Click()
Call ClearControl
Call CmbTxtVisible(True, False)
Edit_Flg = "A"
Call GButtonLock(Me, False)
If RsMst_NCDEXAccreditedCapacity.RecordCount > 0 Then
   TxtStartDate.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmbTxtVisible(var1 As Boolean, var2 As Boolean)
TxtStartDate.Enabled = var1
TxtCloseDate.Enabled = var1
CmbLocationID.Visible = var1
CmdLocationID.Visible = var1
CmbStatusFlag.Visible = var1
TxtAccreditedDescription.Locked = var2
TxtCapacity.Locked = var2
TxtRemark.Locked = var2
ChkCheckAll.Enabled = var1

TxtLocationID.Visible = var2
TxtStatusFlag.Visible = var2
End Sub
Private Sub ChkCheckAll_Click()
If SaveCmd.Enabled = False Then Exit Sub
If MSHFlexGrid2.Rows = 2 Then
   ChkCheckAll.Value = 0
   Exit Sub
End If

If ChkCheckAll.Value = vbChecked Then
   For I = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.row = I
       MSHFlexGrid2.TextMatrix(I, 0) = strChecked
   Next
Else
   For I = 1 To MSHFlexGrid2.Rows - 2
       MSHFlexGrid2.row = I
       MSHFlexGrid2.TextMatrix(I, 0) = strUnChecked
   Next
End If
End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location(Gen_WcLocation)
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub

Private Sub CmbLocationID_LostFocus()
Call Grid_Head1
If CmbLocationID.Text = "" Then Exit Sub
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
mLOcationID = RsMst_Location!LocationID
Call GetGodownDetails
End Sub

Private Sub GetGodownDetails()
Call Grid_Head1

tmp = " Select * From Log_Mst_NCDEXAccreditedCapacity Where AccreditedID = " & Val(TxtAccreditedID)
Call TmpTable

If TmpRs.RecordCount > 0 Then
    tmp = " Select"
    tmp = tmp & " MCMP.CMPName , MG.GodownNo, MG.Address, MG.GodownID, LMNAC.AccreditedID"
    tmp = tmp & " From Mst_CMP MCMP"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.CMPID = MCMP.CMPID"
    tmp = tmp & " Inner Join Mst_NCDEXAccreditedCapacity MNAC On MNAC.LocationID = MCMP.LocationID"
    tmp = tmp & " Left Join Log_Mst_NCDEXAccreditedCapacity LMNAC on LMNAC.AccreditedID = MNAC.AccreditedID And LMNAC.GodownID = MG.GodownID"
    tmp = tmp & " Where MNAC.AccreditedID = " & TxtAccreditedID
    tmp = tmp & " Order by MCMP.CMPName, MG.GodownNo, MG.Address, MG.GodownID"
    Call TmpTable
Else
    tmp = " Select"
    tmp = tmp & " MCMP.CMPName , MG.GodownNo, MG.Address, MG.GodownID, MG.AccreditedID"
    tmp = tmp & " From Mst_CMP MCMP"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.CMPID = MCMP.CMPID"
    tmp = tmp & " Where LOcationID = " & mLOcationID
    tmp = tmp & " Order by MCMP.CMPName, MG.GodownNo, MG.Address, MG.GodownID"
    Call TmpTable
End If

If TmpRs.RecordCount > 0 Then
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
    ChkCheckAll.Value = vbChecked
    For I = 1 To TmpRs.RecordCount
        MSHFlexGrid2.TextMatrix(I, 1) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
        MSHFlexGrid2.TextMatrix(I, 2) = IIf(IsNull(TmpRs!godownno), "", TmpRs!godownno)
        MSHFlexGrid2.TextMatrix(I, 3) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
        MSHFlexGrid2.TextMatrix(I, 4) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
        MSHFlexGrid2.row = I
        MSHFlexGrid2.Col = 0
        MSHFlexGrid2.CellFontName = "Wingdings"
        MSHFlexGrid2.CellFontSize = 12
        MSHFlexGrid2.CellAlignment = flexAlignCenterCenter
        If Edit_Flg = "A" Then
           MSHFlexGrid2.Text = strUnChecked
           ChkCheckAll.Value = vbUnchecked
        Else
           If IsNull(TmpRs!AccreditedID) = True Then
              MSHFlexGrid2.Text = strUnChecked
              ChkCheckAll.Value = vbUnchecked
           Else
              If TmpRs!AccreditedID = TxtAccreditedID.Text Then
                 MSHFlexGrid2.Text = strChecked
              Else
                 MSHFlexGrid2.Text = strUnChecked
              End If
           End If
        End If
        TmpRs.MoveNext
    Next
End If
End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
Call TableMst_Location(Gen_WcLocation)
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp

End Sub

Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then Exit Sub
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection!!!"
   CmbSLocation.SetFocus
   Exit Sub
End If
mLOcationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub
If LCase(CmbSStatus.Text) <> "active" And LCase(CmbSStatus.Text) <> "deactive" Then
    MsgBox "Invalid Selection!!!"
    CmbSStatus.SetFocus
    Exit Sub
End If
End Sub

Private Sub CmbStatusFlag_LostFocus()
If CmbStatusFlag.Text = "" Then Exit Sub

If LCase(CmbStatusFlag.Text) <> "active" And LCase(CmbStatusFlag.Text) <> "deactive" Then
    MsgBox "Invalid Selection!!!"
    CmbStatusFlag.SetFocus
    Exit Sub
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_NCDEXAccreditedCapacity.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_NCDEXAccreditedCapacity.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
oWS.Range("a:a").ColumnWidth = 9.57
oWS.Range("b:b").ColumnWidth = 9.86
oWS.Range("c:c").ColumnWidth = 9.14
oWS.Range("d:d").ColumnWidth = 9.71
oWS.Range("e:e").ColumnWidth = 30
oWS.Range("f:f").ColumnWidth = 16.71
oWS.Range("g:g").ColumnWidth = 8.43
oWS.Range("h:h").ColumnWidth = 30
oWS.Range("i:i").ColumnWidth = 8.29

tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, c + 1) = "NCDEX Accredited Capacity"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, c + 1).WrapText = True

For I = 0 To MSHFlexGrid1.Rows - 2
    mRow = mRow + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)

        If I = 0 Then
           oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, c + 1).Font.Bold = True
           oWS.Cells(mRow, c + 1).RowHeight = 30
           oWS.Cells(mRow, c + 1).WrapText = True
        End If
        oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGo_Click()
Dim Wc As Variant
Call Grid_Head
If TxtFromDate.Value > Date Then
    MsgBox "Future From Date not Allowed!!!"
    TxtFromDate.SetFocus
    Exit Sub
End If
If TxtFromDate.Value > TxtToDate.Value Then
    MsgBox "From Date should not greater than To Date!!!"
    TxtToDate.SetFocus
    Exit Sub
End If
If Gen_DBType = "MDB" Then
    Wc = " WHERE AccreditedDate Between #" & TxtFromDate.Value & "# AND #" & TxtToDate.Value & "#"
Else
    Wc = " WHERE AccreditedDate Between '" & TxtFromDate.Value & "' AND '" & TxtToDate.Value & "'"
End If
If CmbSLocation.Text <> "" Then
    Wc = Wc & " And LocationID = " & mLOcationID
End If
If CmbSStatus.Text <> "" Then
    Wc = Wc & " And Status = '" & Left(CmbSStatus.Text, 1) & "'"
End If

tmp = "Select * From Mst_NCDEXAccreditedCapacity"
tmp = tmp & Wc
Call TmpTable

LblTotalRecord.Caption = TmpRs.RecordCount
If TmpRs.RecordCount > 0 Then
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
    For I = 1 To TmpRs.RecordCount
        With MSHFlexGrid1
            .TextMatrix(I, 0) = IIf(IsNull(TmpRs!AccreditedID), "", TmpRs!AccreditedID)
            .TextMatrix(I, 1) = IIf(IsNull(TmpRs!AccreditedDate), "", Format(TmpRs!AccreditedDate, "dd-MMM-yyyy"))
            .TextMatrix(I, 2) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy"))
            .TextMatrix(I, 3) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-yyyy"))
            .TextMatrix(I, 4) = IIf(IsNull(TmpRs!AccreditedDescription), "", TmpRs!AccreditedDescription)
            .TextMatrix(I, 5) = GetLocationName(IIf(IsNull(TmpRs!LocationID), 0, TmpRs!LocationID))
            .TextMatrix(I, 6) = IIf(IsNull(TmpRs!Capacity), "", TmpRs!Capacity)
            .TextMatrix(I, 7) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
            If TmpRs!Status = "A" Then
               .TextMatrix(I, 8) = "Active"
            Else
               .TextMatrix(I, 8) = "DeActive"
            End If
        End With
        TmpRs.MoveNext
    Next
    CmdExcel.Enabled = True
'Else
'    MsgBox "No Record Found!!!"
End If
End Sub

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   
   RsMst_NCDEXAccreditedCapacity.Delete
   RsMst_NCDEXAccreditedCapacity.Update
   If RsMst_NCDEXAccreditedCapacity.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_NCDEXAccreditedCapacity.MoveNext
   If RsMst_NCDEXAccreditedCapacity.EOF() Then
      RsMst_NCDEXAccreditedCapacity.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_NCDEXAccreditedCapacity.CancelUpdate
End Sub

Private Sub EditCmd_Click()
If TxtCloseDate.Value < Date Then
    MsgBox "NCDEX Accredited is Expired, You can not edit!!!"
    Exit Sub
End If
Edit_Flg = "E"
Call GButtonLock(Me, False)
Call CmbTxtVisible(True, False)
TxtLocationID.Visible = True
CmdLocationID.Visible = False
CmbLocationID.Visible = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_NCDEXAccreditedCapacity.MoveFirst
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
    ButtonFrm.Enabled = False
    MsgBox "Access denied !!!!!!!!!"
    Exit Sub
End If

Frame1.Enabled = True
ButtonFrm.Enabled = True

Call Grid_Head
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
CmbSLocation.Text = ""
CmbSStatus.Text = ""

Call TableMst_NCDEXAccreditedCapacity(" Order by AccreditedID")
If RsMst_NCDEXAccreditedCapacity.RecordCount = 0 Then
    Call AddCmd_Click
    Exit Sub
End If

Call GButtonLock(Me, True)
RsMst_NCDEXAccreditedCapacity.MoveLast
Call Indata
End Sub

Private Sub ClearControl()
TxtAccreditedID.Text = ""
TxtAccreditedDate.Value = Date
TxtStartDate.Value = Date
TxtCloseDate.Value = Date + 30
TxtAccreditedDescription.Text = ""
TxtLocationID.Text = ""
CmbLocationID.Text = ""
TxtCapacity.Text = ""
CmbStatusFlag.Text = ""
TxtStatusFlag.Text = ""
TxtRemark.Text = ""
ChkCheckAll.Value = vbUnchecked
Call Grid_Head1
End Sub

Private Sub Grid_Head1()

With MSHFlexGrid2
    .Clear
    .Rows = 2
    .Cols = 5
    .FixedCols = 0
    .FixedRows = 1
    '.WordWrap = True
    
    '.AllowUserResizing = flexResizeBoth
    .TextMatrix(0, 0) = "Select"
    .TextMatrix(0, 1) = "CMP"
    .TextMatrix(0, 2) = "Godown No"
    .TextMatrix(0, 3) = "Godown Address"
    .TextMatrix(0, 4) = "Godown ID"
    
    .ColWidth(0) = 700
    .ColWidth(1) = 2000
    .ColWidth(2) = 1200
    .ColWidth(3) = 9500
    .ColWidth(4) = 0
End With
End Sub

Private Sub LastCmd_Click()
RsMst_NCDEXAccreditedCapacity.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
If mCaId = "" Then Exit Sub
RsMst_NCDEXAccreditedCapacity.MoveFirst

If RsMst_NCDEXAccreditedCapacity.RecordCount = 0 Then
    Call AddCmd_Click
    Exit Sub
End If
RsMst_NCDEXAccreditedCapacity.MoveFirst
RsMst_NCDEXAccreditedCapacity.Find "AccreditedID = " & Val(mCaId) & ""
If RsMst_NCDEXAccreditedCapacity.EOF Then
    RsMst_NCDEXAccreditedCapacity.MoveLast
End If
Frame0.Visible = False
Call Indata
End Sub

Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
With MSHFlexGrid2
    If .Col = 0 Then  '.Row = 1 And
        If .TextMatrix(.row, .Col) = "" Then Exit Sub
        If .TextMatrix(.row, .Col) = strUnChecked Then
           .TextMatrix(.row, .Col) = strChecked
        Else
           .TextMatrix(.row, .Col) = strUnChecked
        End If
    End If
End With
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_NCDEXAccreditedCapacity.MoveNext
If RsMst_NCDEXAccreditedCapacity.EOF Then
   RsMst_NCDEXAccreditedCapacity.MoveLast
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
RsMst_NCDEXAccreditedCapacity.MovePrevious
If RsMst_NCDEXAccreditedCapacity.BOF Then
   RsMst_NCDEXAccreditedCapacity.MoveFirst
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
If TxtStartDate.Value > Date Then
    MsgBox "Future Start Date not Allowed!!!"
    TxtStartDate.SetFocus
    Exit Sub
End If
If TxtStartDate.Value > TxtCloseDate.Value Then
    MsgBox "Start Date should not greater than CloseDate!!!"
    TxtCloseDate.SetFocus
    Exit Sub
End If
If TxtAccreditedDescription.Text = "" Then
    MsgBox "Blank Accredited Description not Allowed!!!"
    TxtAccreditedDescription.SetFocus
    Exit Sub
End If
If CmbLocationID.Text = "" Then
    MsgBox "Blank Location not Allowed!!!"
    CmbLocationID.SetFocus
    Exit Sub
End If
If TxtCapacity.Text = "" Then
    MsgBox "Blank Capacity not Allowed!!!"
    TxtCapacity.SetFocus
    Exit Sub
End If
If CmbStatusFlag.Text = "" Then
    MsgBox "Blank Status not Allowed!!!"
    CmbStatusFlag.SetFocus
    Exit Sub
End If
If TxtRemark.Text = "" Then
    MsgBox "Blank Remarks not Allowed!!!"
    TxtRemark.SetFocus
    Exit Sub
End If

If Edit_Flg = "A" Then
    If RsMst_NCDEXAccreditedCapacity.RecordCount > 0 Then
        tmp = "Select * From Mst_NCDEXAccreditedCapacity Where LocationID = " & mLOcationID
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
            If Not TmpRs.EOF Then
                TmpRs.MoveLast
                If TmpRs!CloseDate > Date Then
                    MsgBox "Duplicate Location Not Allowed !!! "
                    CmbLocationID.SetFocus
                    Exit Sub
                Else
                    tmp = " Select * From Mst_Godown Where AccreditedID in (Select Max(AccreditedID) from Mst_NCDEXAccreditedCapacity Where LocationID  = " & mLOcationID & " )"
                    Call Tmp1Table
                    
                    tmp = "Select * From Log_Mst_NCDEXAccreditedCapacity"
                    Call TmpTable
                    
                    For I = 1 To TmpRs1.RecordCount
                        TmpRs.AddNew
                        TmpRs!AccreditedID = RsMst_NCDEXAccreditedCapacity!AccreditedID
                        TmpRs!GodownID = TmpRs1!GodownID
                        TmpRs!CreatedDate = RsMst_NCDEXAccreditedCapacity!CreatedDate
                        TmpRs!CreatedBy = RsMst_NCDEXAccreditedCapacity!CreatedBy
                        TmpRs!UpdatedDate = RsMst_NCDEXAccreditedCapacity!UpdatedDate
                        TmpRs!UpdatedBy = RsMst_NCDEXAccreditedCapacity!UpdatedBy
                        TmpRs.Update
                        TmpRs.Requery
                        TmpRs1.MoveNext
                    Next
                End If
            End If
        End If
    End If
    RsMst_NCDEXAccreditedCapacity.AddNew
    RsMst_NCDEXAccreditedCapacity!AccreditedID = GetAccreditedID
    TxtAccreditedID = RsMst_NCDEXAccreditedCapacity!AccreditedID
    RsMst_NCDEXAccreditedCapacity!G_UID = GetGUID
Else
    RsMst_NCDEXAccreditedCapacity.MoveFirst
    RsMst_NCDEXAccreditedCapacity.Find "AccreditedID= '" & TxtAccreditedID.Text & "'"
End If

RsMst_NCDEXAccreditedCapacity!AccreditedDate = TxtAccreditedDate.Value
RsMst_NCDEXAccreditedCapacity!StartDate = TxtStartDate.Value
RsMst_NCDEXAccreditedCapacity!CloseDate = TxtCloseDate.Value
RsMst_NCDEXAccreditedCapacity!AccreditedDescription = TxtAccreditedDescription.Text
RsMst_NCDEXAccreditedCapacity!LocationID = mLOcationID
RsMst_NCDEXAccreditedCapacity!Capacity = Val(TxtCapacity.Text)
RsMst_NCDEXAccreditedCapacity!Status = Left(CmbStatusFlag.Text, 1)
RsMst_NCDEXAccreditedCapacity!Remarks = TxtRemark.Text

If IsNull(RsMst_NCDEXAccreditedCapacity!CreatedBy) = True Or RsMst_NCDEXAccreditedCapacity!CreatedBy = "" Then
   RsMst_NCDEXAccreditedCapacity!CreatedBy = Gen_UserLoginID
   RsMst_NCDEXAccreditedCapacity!CreatedDate = Now
Else
   RsMst_NCDEXAccreditedCapacity!UpdatedBy = Gen_UserLoginID
   RsMst_NCDEXAccreditedCapacity!UpdatedDate = Now
End If

RsMst_NCDEXAccreditedCapacity.Update
RsMst_NCDEXAccreditedCapacity.Requery

Call SaveGodownDetail
Call GButtonLock(Me, True)

'RsMst_NCDEXAccreditedCapacity.MoveLast
RsMst_NCDEXAccreditedCapacity.Find "AccreditedID = " & TxtAccreditedID.Text & ""
Call Indata

End Sub

Private Sub Indata()
Edit_Flg = ""
Call CmbTxtVisible(False, True)
TxtAccreditedID.Text = IIf(IsNull(RsMst_NCDEXAccreditedCapacity!AccreditedID), "", RsMst_NCDEXAccreditedCapacity!AccreditedID)
TxtAccreditedDate.Value = IIf(IsNull(RsMst_NCDEXAccreditedCapacity!AccreditedDate), "01/01/2011", RsMst_NCDEXAccreditedCapacity!AccreditedDate)
TxtStartDate.Value = IIf(IsNull(RsMst_NCDEXAccreditedCapacity!StartDate), "01/01/2011", RsMst_NCDEXAccreditedCapacity!StartDate)
TxtCloseDate.Value = IIf(IsNull(RsMst_NCDEXAccreditedCapacity!CloseDate), "01/01/2011", RsMst_NCDEXAccreditedCapacity!CloseDate)
TxtAccreditedDescription.Text = IIf(IsNull(RsMst_NCDEXAccreditedCapacity!AccreditedDescription), "", RsMst_NCDEXAccreditedCapacity!AccreditedDescription)
mLOcationID = IIf(IsNull(RsMst_NCDEXAccreditedCapacity!LocationID), 0, RsMst_NCDEXAccreditedCapacity!LocationID)
Call GetGodownDetails
TxtLocationID.Text = GetLocationName(Val(mLOcationID))
CmbLocationID.Text = TxtLocationID.Text
TxtCapacity.Text = IIf(IsNull(RsMst_NCDEXAccreditedCapacity!Capacity), "", RsMst_NCDEXAccreditedCapacity!Capacity)
If RsMst_NCDEXAccreditedCapacity!Status = "A" Then
   TxtStatusFlag = "Active"
Else
   TxtStatusFlag = "DeActive"
End If
CmbStatusFlag.Text = TxtStatusFlag.Text
TxtRemark.Text = IIf(IsNull(RsMst_NCDEXAccreditedCapacity!Remarks), "", RsMst_NCDEXAccreditedCapacity!Remarks)
TxtCreated = IIf(IsNull(RsMst_NCDEXAccreditedCapacity!CreatedBy), "", RsMst_NCDEXAccreditedCapacity!CreatedBy) & " :- " & IIf(IsNull(RsMst_NCDEXAccreditedCapacity!CreatedDate), "", RsMst_NCDEXAccreditedCapacity!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_NCDEXAccreditedCapacity!UpdatedBy), "", RsMst_NCDEXAccreditedCapacity!UpdatedBy) & " :- " & IIf(IsNull(RsMst_NCDEXAccreditedCapacity!UpdatedDate), "", RsMst_NCDEXAccreditedCapacity!UpdatedDate)
End Sub

Private Sub SaveGodownDetail()
For I = 1 To MSHFlexGrid2.Rows - 2
    tmp = "Select * From Mst_Godown Where GodownID = " & MSHFlexGrid2.TextMatrix(I, 4)
    Call TmpTable
    If MSHFlexGrid2.TextMatrix(I, 0) = strChecked Then
        TmpRs!AccreditedID = TxtAccreditedID.Text
    Else
        TmpRs!AccreditedID = Null
    End If
    TmpRs.Update
Next
End Sub

Private Function GetAccreditedID() As Double
Dim Retval As Double
tmp = "Select max(AccreditedID) from Mst_NCDEXAccreditedCapacity"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetAccreditedID = Retval
End Function

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_NCDEXAccreditedCapacity.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_NCDEXAccreditedCapacity.EOF Then
      RsMst_NCDEXAccreditedCapacity.MoveFirst
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
MSHFlexGrid1.Height = Frame0.Height - 1200
'TxtSQualityParameter.Text = ""
'Label22.Caption = ""
RsMst_NCDEXAccreditedCapacity.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head
'TxtFromDate.Value = Date - 30
'TxtToDate.Value = Date
'CmbSLocation.Text = ""
'CmbSStatus.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub Grid_Head()
CmdExcel.Enabled = False
LblTotalRecord.Caption = ""
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 9
    .FixedCols = 1
    .FixedRows = 1
    '.WordWrap = True
    
    '.AllowUserResizing = flexResizeBoth
    .TextMatrix(0, 0) = "Accredited ID"
    .TextMatrix(0, 1) = "Accredited Date"
    .TextMatrix(0, 2) = "Start Date"
    .TextMatrix(0, 3) = "Close Date"
    .TextMatrix(0, 4) = "Accredited Description"
    .TextMatrix(0, 5) = "Location"
    .TextMatrix(0, 6) = "Capacity"
    .TextMatrix(0, 7) = "Remark"
    .TextMatrix(0, 8) = "Status"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1200
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 3500
    .ColWidth(5) = 1600
    .ColWidth(6) = 1200
    .ColWidth(7) = 2400
    .ColWidth(8) = 1000
    .RowHeight(0) = 700
    .WordWrap = True
End With
End Sub

Private Sub TxtCapacity_LostFocus()
If TxtCapacity = "" Then Exit Sub
If Not IsNumeric(TxtCapacity) Or Val(TxtCapacity) < 0 Then
   MsgBox "Invalid Number Foramt !!!!"
   TxtCapacity.SetFocus
   Exit Sub
End If
End Sub
