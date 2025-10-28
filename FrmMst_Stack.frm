VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Stack 
   Caption         =   "Stack Master"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5940
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   5940
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   3285
      Left            =   120
      TabIndex        =   34
      Top             =   9600
      Visible         =   0   'False
      Width           =   15000
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
         ItemData        =   "FrmMst_Stack.frx":0000
         Left            =   12840
         List            =   "FrmMst_Stack.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   450
         Width           =   1995
      End
      Begin VB.TextBox TxtSStack 
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
         Left            =   10920
         MaxLength       =   10
         TabIndex        =   38
         Top             =   450
         Width           =   1875
      End
      Begin VB.ComboBox CmbSGodownID 
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
         Left            =   7560
         Sorted          =   -1  'True
         TabIndex        =   37
         Top             =   450
         Width           =   3315
      End
      Begin VB.ComboBox CmbSCMPID 
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
         Left            =   3840
         Sorted          =   -1  'True
         TabIndex        =   36
         Top             =   450
         Width           =   3675
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
         Left            =   9070
         TabIndex        =   40
         Top             =   855
         Width           =   885
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
         Left            =   9960
         TabIndex        =   41
         Top             =   855
         Width           =   885
      End
      Begin VB.ComboBox CmbSLocationID 
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
         Sorted          =   -1  'True
         TabIndex        =   35
         Top             =   450
         Width           =   3675
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   975
         Left            =   75
         TabIndex        =   42
         Top             =   1320
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   1720
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         Appearance      =   0
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Status Flag"
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
         Left            =   13335
         TabIndex        =   50
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label6 
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
         Left            =   11445
         TabIndex        =   49
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Godown Name"
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
         Left            =   8542
         TabIndex        =   48
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   5160
         TabIndex        =   46
         Top             =   180
         Width           =   1035
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
         Left            =   13605
         TabIndex        =   45
         Top             =   915
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
         Left            =   11205
         TabIndex        =   44
         Top             =   915
         Width           =   2355
      End
      Begin VB.Label Label19 
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
         Left            =   1575
         TabIndex        =   43
         Top             =   180
         Width           =   765
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   32
      Top             =   4560
      Width           =   7200
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
         Left            =   1500
         TabIndex        =   16
         Top             =   240
         Width           =   1400
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
         Left            =   105
         TabIndex        =   15
         Top             =   240
         Width           =   1400
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
         Left            =   105
         TabIndex        =   20
         Top             =   645
         Width           =   1400
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
         Left            =   1500
         TabIndex        =   21
         Top             =   645
         Width           =   1400
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
         Left            =   2895
         TabIndex        =   17
         Top             =   240
         Width           =   1400
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
         Height          =   405
         Left            =   2895
         TabIndex        =   22
         Top             =   645
         Width           =   1400
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
         Left            =   4290
         TabIndex        =   18
         Top             =   240
         Width           =   1400
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
         Left            =   4290
         TabIndex        =   23
         Top             =   645
         Width           =   1400
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
         Left            =   5685
         TabIndex        =   19
         Top             =   240
         Width           =   1400
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
         Left            =   5685
         TabIndex        =   24
         Top             =   645
         Width           =   1400
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   7200
      Begin VB.TextBox TxtNoofStack 
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
         Left            =   5400
         MaxLength       =   3
         TabIndex        =   9
         Top             =   2000
         Width           =   1680
      End
      Begin VB.TextBox TxtStackFrom 
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
         Left            =   2085
         MaxLength       =   10
         TabIndex        =   8
         Top             =   2000
         Width           =   1680
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
         Height          =   615
         Left            =   2085
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2850
         Width           =   4995
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
         Left            =   2085
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1550
         Width           =   4995
      End
      Begin VB.TextBox TxtStackNo 
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
         Left            =   2085
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1995
         Width           =   4995
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
         Left            =   2085
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   675
         Width           =   4995
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
         ItemData        =   "FrmMst_Stack.frx":0021
         Left            =   2085
         List            =   "FrmMst_Stack.frx":002B
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   2415
         Width           =   4995
      End
      Begin VB.ComboBox CmbCMPID 
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
         Left            =   2085
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1120
         Width           =   4995
      End
      Begin VB.TextBox TxtLocationName 
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
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   690
         Visible         =   0   'False
         Width           =   4995
      End
      Begin VB.TextBox TxtStackID 
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
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1680
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
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4035
         Width           =   4995
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
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3570
         Width           =   4995
      End
      Begin VB.TextBox TxtCMPName 
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
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1110
         Visible         =   0   'False
         Width           =   4995
      End
      Begin VB.TextBox TxtGodownName 
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
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1560
         Visible         =   0   'False
         Width           =   4995
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
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2415
         Visible         =   0   'False
         Width           =   4995
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "No of Stack"
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
         TabIndex        =   54
         Top             =   2060
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Stack Start From"
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
         TabIndex        =   53
         Top             =   2055
         Width           =   1740
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label26 
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
         Height          =   210
         Left            =   120
         TabIndex        =   51
         Top             =   3052
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Godown Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   47
         Top             =   1600
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Stack No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   33
         Top             =   2070
         Width           =   2280
      End
      Begin VB.Label LblStateName 
         Caption         =   "Location Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   31
         Top             =   750
         Width           =   1800
      End
      Begin VB.Label LblStackID 
         Caption         =   "Stack ID"
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
         Left            =   120
         TabIndex        =   30
         Top             =   308
         Width           =   1800
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
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   3645
         Width           =   1695
      End
      Begin VB.Label Label1 
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
         Height          =   210
         Left            =   120
         TabIndex        =   28
         Top             =   1155
         Width           =   1800
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
         TabIndex        =   27
         Top             =   4095
         Width           =   1455
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
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
         TabIndex        =   26
         Top             =   2445
         Width           =   480
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   9495
      Left            =   7440
      TabIndex        =   52
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   16748
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      Appearance      =   0
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
Attribute VB_Name = "FrmMst_Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As String
Dim mLocationID, mCMPID, mGodownID As Double
Dim mSLocationID, mSCMPID, mSGodownID As Double

Private Sub CmbCMPID_GotFocus()
If CmbLocationID = "" Then
    MsgBox "Select Location.", vbOKOnly + vbInformation, App.Title
    CmbLocationID.SetFocus
    Exit Sub
End If

TxtStackNo = ""

tmp = CmbCMPID.Text

If Gen_WcCMP <> "" Then
    Call TableMst_CMP(" Where CloseDate Is Null AND " & Gen_WcCMP & " AND LocationID = " & mLocationID & " ")
Else
    Call TableMst_CMP(" Where CloseDate Is Null AND LocationID = " & mLocationID & " ")
End If
CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
   CmbCMPID.AddItem RsMst_CMP!CMPName
   RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp
End Sub
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   CmbGodownID.Text = ""
   Exit Sub
End If

If tmp <> CmbCMPID.Text Then
   CmbGodownID.Text = ""
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbGodownID_GotFocus()
If CmbCMPID = "" Then
    MsgBox "Select CMP.", vbOKOnly + vbInformation, App.Title
    CmbCMPID.SetFocus
    Exit Sub
End If

tmp = CmbGodownID.Text
Call TableMst_Godown(" Where CloseDate Is Null AND CMPID = " & mCMPID & " ")
CmbGodownID.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For i = 1 To RsMst_Godown.RecordCount
   CmbGodownID.AddItem RsMst_Godown!GodownNo
   RsMst_Godown.MoveNext
Next
CmbGodownID.Text = tmp
End Sub

Private Sub CmbGodownID_LostFocus()
If CmbGodownID = "" Then Exit Sub

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbGodownID.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbGodownID.SetFocus
   Exit Sub
End If
mGodownID = RsMst_Godown!GodownID
''-- Call Stack List
Call StackListUnderGodown(CDbl(mGodownID))

If Edit_Flg = "A" And MSHFlexGrid2.Rows = 2 Then
   Label8.Visible = True: TxtStackFrom.Visible = True: TxtStackFrom = ""
   Label11.Visible = True: TxtNoOfStack.Visible = True: TxtNoOfStack = ""
   Label3.Visible = False: TxtStackNo.Visible = False: TxtStackNo = ""
   TxtStackFrom.SetFocus
Else
   Label8.Visible = False: TxtStackFrom.Visible = False: TxtStackFrom = ""
   Label11.Visible = False: TxtNoOfStack.Visible = False: TxtNoOfStack = ""
   Label3.Visible = True: TxtStackNo.Visible = True: TxtStackNo = ""
   
   Me.MousePointer = vbHourglass
   tmp = " SELECT ISNULL(MAX(Convert(Decimal,StackNo)),0) +1  FROM Mst_Stack WHERE GodownID = " & mGodownID & " AND Isnumeric(StackNo) = 1 "
   Call TmpTable
   Me.MousePointer = vbDefault
   If TmpRs.RecordCount > 0 Then
      TxtStackNo = TmpRs.Fields(0).Value
      TxtStackNo.SetFocus
   End If
End If
''-And MSHFlexGrid2.Rows > 2
End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
End If
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
   CmbLocationID.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub

Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   CmbCMPID.Text = ""
   CmbGodownID.Text = ""
   Exit Sub
End If

If tmp <> CmbLocationID.Text Then
   CmbCMPID.Text = ""
   CmbGodownID.Text = ""
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSStatus_Change()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "active" And LCase(CmbSStatus.Text) <> "de-active" Then
   MsgBox "Invalid Selection.", vbOKOnly + vbInformation, App.Title
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdGo_Click()
Dim wc As Variant
wc = ""

tmp = " Select MS.StackID,ML.LocationName,MC.CMPName,MG.GodownNo,MS.StackNo,MS.Flag,MS.Remark "
tmp = tmp & " from Mst_Stack MS "
tmp = tmp & " Inner Join Mst_Godown MG On MS.GodownID=MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC On MG.CMPID=MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID=ML.LocationID"

'tmp = tmp & " Left Join Mst_Location ML On MS.LocationID=ML.LocationID "
'tmp = tmp & " Left Join Mst_CMP MC On MS.CMPID=MC.CMPID "
'tmp = tmp & " Left Join Mst_Godown MG On MS.GodownID=MG.GodownID "

If CmbSLocationID <> "" Then
   If wc = "" Then
      wc = " Where ML.LocationID = " & mSLocationID & " "
   End If
End If

If CmbSCMPID <> "" Then
   If wc = "" Then
      wc = " Where MC.CMPID =" & mSCMPID & " "
   Else
      wc = wc & " AND MC.CMPID =" & mSCMPID & " "
   End If
End If

If CmbSGodownID <> "" Then
   If wc = "" Then
      wc = " Where MG.GodownID =" & mSGodownID & " "
   Else
      wc = wc & " AND MG.GodownID =" & mSGodownID & " "
   End If
End If

If TxtSStack <> "" Then
   If wc = "" Then
      wc = " Where MS.StackNo ='" & TxtSStack & "'"
   Else
      wc = wc & " AND MS.StackNo ='" & TxtSStack & "'"
   End If
End If

If CmbSStatus <> "" Then
   If wc = "" Then
      wc = " Where MS.Flag ='" & Left(CmbSStatus.Text, 1) & "' "
   Else
      wc = wc & " AND MS.Flag ='" & Left(CmbSStatus.Text, 1) & "' "
   End If
End If

tmp = tmp & wc & " Order By MS.StackID "
Call TmpTable

Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    With MSHFlexGrid1
        .TextMatrix(iu, 0) = IIf(IsNull(TmpRs!StackID), "", TmpRs!StackID)
        .TextMatrix(iu, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
        .TextMatrix(iu, 2) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
        .TextMatrix(iu, 3) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
        .TextMatrix(iu, 4) = IIf(IsNull(TmpRs!StackNo), "", TmpRs!StackNo)
        .TextMatrix(iu, 5) = IIf(IsNull(TmpRs!Flag), "", IIf(LCase(TmpRs!Flag) = "a", "Active", IIf(LCase(TmpRs!Flag) = "d", "De-Active", "")))
        .TextMatrix(iu, 6) = IIf(IsNull(TmpRs!Remark), "", TmpRs!Remark)
        .Rows = .Rows + 1
        .Refresh
    End With
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError

tmp = " SELECT * FROM MST_GSL Where GodownID = " & mGodownID & " AND StackNo = '" & Trim(TxtStackNo) & "' "
Call TmpTable
If TmpRs.RecordCount > 0 Then
    MsgBox "Stack Details are present with GSL Master.", vbOKOnly + vbInformation, App.Title
    Exit Sub
End If

ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_Stack.Delete
   RsMst_Stack.Update
   If RsMst_Stack.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Stack.MoveNext
   If RsMst_Stack.EOF() Then
      RsMst_Stack.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Stack.CancelUpdate
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_HeadStackList
Call Grid_Head
CmbSLocationID = ""
CmbSCMPID = ""
CmbSGodownID = ""
TxtSStack = ""
CmbSStatus = ""
Label22.Caption = ""

Call TableMst_Stack
If RsMst_Stack.RecordCount = 0 Then
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
Call Indata
End Sub

Private Sub Grid_Head()
    With MSHFlexGrid1
        .Clear
        .Rows = 2
        .Cols = 7
        .TextMatrix(0, 0) = "StackID":          .ColWidth(0) = 1000
        .TextMatrix(0, 1) = "Location Name":    .ColWidth(1) = 3000
        .TextMatrix(0, 2) = "CMP Name":         .ColWidth(2) = 3000
        .TextMatrix(0, 3) = "Godown Name":      .ColWidth(3) = 2500
        .TextMatrix(0, 4) = "Stack No":         .ColWidth(4) = 1200
        .TextMatrix(0, 5) = "Status Flag":      .ColWidth(5) = 1800
        .TextMatrix(0, 6) = "Remark":           .ColWidth(6) = 2500
    End With
End Sub

Private Sub Grid_HeadStackList()
    With MSHFlexGrid2
        .Clear
        .Rows = 2
        .Cols = 6
        .TextMatrix(0, 0) = "Sr No":            .ColWidth(0) = 900
        .TextMatrix(0, 1) = "Stack No":         .ColWidth(1) = 2000
        .TextMatrix(0, 2) = "Status Flag":      .ColWidth(2) = 1200
        .TextMatrix(0, 3) = "Balance Bags":     .ColWidth(3) = 1500
        .TextMatrix(0, 4) = "Balance Weight":   .ColWidth(4) = 1800
        .TextMatrix(0, 5) = "StackID":          .ColWidth(5) = 0
    End With
End Sub

Private Sub MSHFlexGrid1_Click()
    Dim i As Variant
    i = MSHFlexGrid1.RowSel
    If i <= 0 Then Exit Sub
    If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
    RsMst_Stack.MoveFirst
    RsMst_Stack.Find "StackID = " & MSHFlexGrid1.TextMatrix(i, 0)
    If Not RsMst_Stack.EOF Then
       Call Indata
       Frame0.Visible = False
    End If
End Sub

Private Sub MSHFlexGrid2_Click()
    If SaveCmd.Enabled = False Then
       If MSHFlexGrid2.Rows > 2 Then
        If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5) <> "" Then
            RsMst_Stack.MoveFirst
            RsMst_Stack.Find "StackID = " & Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 5))
            'Call Indata
            ''------------
            Me.MousePointer = vbHourglass
                
                mGodownID = IIf(IsNull(RsMst_Stack!GodownID), 0, RsMst_Stack!GodownID)
                mCMPID = GetCMPIDFromGodown(mGodownID)
                mLocationID = GetLocationIDFromCMP(CDbl(mCMPID))
                
                TxtStackId = IIf(IsNull(RsMst_Stack!StackID), "", RsMst_Stack!StackID)
                TxtLocationName = GetLocationName(CDbl(mLocationID))
                TxtLocationName.Visible = True: CmbLocationID.Visible = False
                TxtCMPName = GetCMPName(CDbl(mCMPID))
                TxtCMPName.Visible = True: CmbCMPID.Visible = False
                
                TxtGodownName = GetGodownName(CDbl(mGodownID), "N")
                TxtGodownName.Visible = True: CmbGodownID.Visible = False
                
                TxtStackNo = IIf(IsNull(RsMst_Stack!StackNo), "", RsMst_Stack!StackNo)
                TxtStackNo.Locked = True
                Label3.Visible = True: TxtStackNo.Visible = True
                
                Label8.Visible = False: TxtStackFrom.Visible = False: TxtStackFrom = ""
                Label11.Visible = False: TxtNoOfStack.Visible = False: TxtNoOfStack = ""
                
                If LCase(RsMst_Stack!Flag) = "a" Then
                   TxtFlag.Text = "Active"
                ElseIf LCase(RsMst_Stack!Flag) = "d" Then
                   TxtFlag.Text = "De-Active"
                Else
                   TxtFlag.Text = ""
                End If
                
                TxtFlag.Visible = True: CmbFlag.Visible = False
                
                TxtCreated = IIf(IsNull(RsMst_Stack!CreatedBy), "", RsMst_Stack!CreatedBy) & " :- " & IIf(IsNull(RsMst_Stack!CreatedDate), "", RsMst_Stack!CreatedDate)
                TxtUpdated = IIf(IsNull(RsMst_Stack!UpdatedBy), "", RsMst_Stack!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Stack!UpdatedDate), "", RsMst_Stack!UpdatedDate)
                
                TxtRemarks = IIf(IsNull(RsMst_Stack!Remark), "", RsMst_Stack!Remark)
                TxtRemarks.Locked = True
                
                If RsMst_Stack.RecordCount = 1 Then
                   Call GButtonLock_1(Me, True)
                Else
                   Call GButtonLock(Me, True)
                End If
                
            Me.MousePointer = vbDefault
            ''------------
        End If
       End If
    End If
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Stack.MoveNext
If RsMst_Stack.EOF Then
   RsMst_Stack.MoveLast
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
RsMst_Stack.MovePrevious
If RsMst_Stack.BOF Then
   RsMst_Stack.MoveFirst
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
RsMst_Stack.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub LastCmd_Click()
RsMst_Stack.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Stack.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Stack.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(i, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmbFlag_LostFocus()
If CmbFlag.Text = "" Then Exit Sub

If LCase(CmbFlag.Text) <> "active" And LCase(CmbFlag.Text) <> "de-active" Then
   MsgBox "Invalid Selection"
   CmbFlag.SetFocus
   Exit Sub
End If
End Sub

Public Sub Indata()
Dim mR As Double
Me.MousePointer = vbHourglass

mGodownID = IIf(IsNull(RsMst_Stack!GodownID), 0, RsMst_Stack!GodownID)
mCMPID = GetCMPIDFromGodown(mGodownID)
mLocationID = GetLocationIDFromCMP(CDbl(mCMPID))

TxtStackId = IIf(IsNull(RsMst_Stack!StackID), "", RsMst_Stack!StackID)
'mLocationID = IIf(IsNull(RsMst_Stack!LocationID), 0, RsMst_Stack!LocationID)
TxtLocationName = GetLocationName(CDbl(mLocationID))
TxtLocationName.Visible = True
CmbLocationID.Visible = False
'mCMPID = IIf(IsNull(RsMst_Stack!CMPID), 0, RsMst_Stack!CMPID)
TxtCMPName = GetCMPName(CDbl(mCMPID))
TxtCMPName.Visible = True
CmbCMPID.Visible = False

TxtGodownName = GetGodownName(CDbl(mGodownID), "N")
TxtGodownName.Visible = True
CmbGodownID.Visible = False

Call StackListUnderGodown(CDbl(mGodownID))

TxtStackNo = IIf(IsNull(RsMst_Stack!StackNo), "", RsMst_Stack!StackNo)
TxtStackNo.Locked = True

Label3.Visible = True
TxtStackNo.Visible = True

Label8.Visible = False
Label11.Visible = False
TxtStackFrom.Visible = False: TxtStackFrom = ""
TxtNoOfStack.Visible = False: TxtNoOfStack = ""

If LCase(RsMst_Stack!Flag) = "a" Then
   TxtFlag.Text = "Active"
ElseIf LCase(RsMst_Stack!Flag) = "d" Then
   TxtFlag.Text = "De-Active"
Else
   TxtFlag.Text = ""
End If
TxtFlag.Visible = True
CmbFlag.Visible = False

TxtCreated = IIf(IsNull(RsMst_Stack!CreatedBy), "", RsMst_Stack!CreatedBy) & " :- " & IIf(IsNull(RsMst_Stack!CreatedDate), "", RsMst_Stack!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Stack!UpdatedBy), "", RsMst_Stack!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Stack!UpdatedDate), "", RsMst_Stack!UpdatedDate)

TxtRemarks = IIf(IsNull(RsMst_Stack!Remark), "", RsMst_Stack!Remark)
TxtRemarks.Locked = True

If RsMst_Stack.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

Me.MousePointer = vbDefault
End Sub

Private Sub StackListUnderGodown(mGodown_ As Double)
Me.MousePointer = vbHourglass
    Call Grid_HeadStackList
    ''tmp = " SELECT * FROM Mst_Stack Where GodownID = " & mGodown_ & IIf(Val(TxtStackID) = 0, "", " AND StackID NOT IN (" & Val(TxtStackID) & ")")
    tmp = " SELECT MS.StackID,MG.StackNo,MS.Flag,SUM(MG.BalanceBags)BalanceBags,SUM(MG.BalanceWeight)BalanceWeight "
    tmp = tmp & " FROM Mst_Stack as MS Left Join Mst_GSL as MG On MS.GodownID = MG.GodownID AND MS.StackNo = MG.StackNo "
    tmp = tmp & " Where MG.GodownID = " & mGodown_ & " Group By MS.Flag,MG.StackNo,MS.StackID "
    ''tmp = tmp & " Where MG.GodownID = " & mGodown_ & "" & IIf(Val(TxtStackNo) = 0, "", " AND StackID NOT IN (" & Val(TxtStackID) & ") ") & " Group By MS.Flag,MG.StackNo Order By MG.StackNo "
    tmp = tmp & " Union"
    tmp = tmp & " SELECT MS.StackID,MS.StackNo,MS.Flag,0,0 FROM Mst_Stack as MS Where MS.GodownID = " & mGodown_ & ""
    tmp = tmp & " AND MS.StackID NOT IN ( SELECT MS.StackID FROM Mst_Stack as MS Left Join Mst_GSL as MG On MS.GodownID = MG.GodownID AND MS.StackNo = MG.StackNo Where MG.GodownID = " & mGodown_ & " Group By MS.Flag,MG.StackNo,MS.StackID ) "
    ''Order By MG.StackNo
    
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        With MSHFlexGrid2
            For mR = 1 To TmpRs.RecordCount
                .TextMatrix(mR, 0) = mR
                .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!StackNo), "", TmpRs!StackNo)
                .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!Flag), "", IIf(LCase(TmpRs!Flag) = "a", "Active", IIf(LCase(TmpRs!Flag) = "d", "De-Active", "")))
                .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags)
                .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
                .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!StackID), "", TmpRs!StackID)
                TmpRs.MoveNext
                .Rows = .Rows + 1
                .Refresh
            Next mR
        End With
    End If
Me.MousePointer = vbDefault
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtStackId = ""
TxtLocationName = "": TxtLocationName.Visible = False
CmbLocationID = "": CmbLocationID.Visible = True

TxtCMPName = "": TxtCMPName.Visible = False
CmbCMPID = "": CmbCMPID.Visible = True

TxtGodownName = "": TxtGodownName.Visible = False
CmbGodownID = "": CmbGodownID.Visible = True

TxtStackNo = "": TxtStackNo.Locked = False

Label8.Visible = False
Label11.Visible = False
TxtStackFrom.Visible = False
TxtNoOfStack.Visible = False

TxtFlag = "Active": TxtFlag.Visible = True
CmbFlag = "Active": CmbFlag.Visible = False

TxtCreated = ""
TxtUpdated = ""
TxtRemarks = "": TxtRemarks.Locked = False

If RsMst_Stack.RecordCount > 0 Then
   CmbLocationID.SetFocus
End If
Call Grid_HeadStackList
End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

CmbLocationID = TxtLocationName
CmbCMPID = TxtCMPName
CmbGodownID = TxtGodownName

tmp = " SELECT * FROM MST_GSL Where GodownID = " & mGodownID & " AND StackNo = '" & Trim(TxtStackNo) & "' "
Call TmpTable
If TmpRs.RecordCount = 0 Then
    CmbLocationID.Visible = True
    TxtLocationName.Visible = False
    CmbCMPID.Visible = True
    TxtCMPName.Visible = False
    CmbGodownID.Visible = True
    TxtGodownName.Visible = False
    TxtStackNo.Locked = False
End If

CmbFlag = TxtFlag
CmbFlag.Visible = True
TxtFlag.Visible = False
TxtRemarks.Locked = False

End Sub

Private Sub SaveCmd_Click()
Dim mStackFrom, mNoofStack As Variant
Dim mStackCount As Double

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location Name Not Allowed !!! "
   CmbLocationID.SetFocus
   Exit Sub
End If
If CmbCMPID.Text = "" Then
   MsgBox "Blank CMP Name Not Allowed !!! "
   CmbCMPID.SetFocus
   Exit Sub
End If
If CmbGodownID.Text = "" Then
   MsgBox "Blank Godown Name Not Allowed !!! "
   CmbGodownID.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" And MSHFlexGrid2.Rows = 2 Then
    If TxtStackFrom = "" Then
       MsgBox "Blank Stack Start From Not Allowed !!! "
       TxtStackFrom.SetFocus
       Exit Sub
    End If
    If Val(TxtStackFrom) = 0 Then
       MsgBox "Stack No Start From Zero Not Allowed !!! "
       TxtStackFrom.SetFocus
       Exit Sub
    End If

    If TxtNoOfStack = "" Then
       MsgBox "Blank No of Stack Not Allowed !!! "
       TxtNoOfStack.SetFocus
       Exit Sub
    End If
    If Val(TxtNoOfStack) = 0 Then
       MsgBox "Zero No of Stack Not Allowed !!! "
       TxtNoOfStack.SetFocus
       Exit Sub
    End If
Else
    If TxtStackNo = "" Then
       MsgBox "Blank Stack No Not Allowed !!! "
       TxtStackNo.SetFocus
       Exit Sub
    End If
End If

If IsClean(TxtStackNo) = False Then
    MsgBox "Special characters not allowed.", vbOKOnly + vbInformation, App.Title
    TxtStackNo.SetFocus
    Exit Sub
End If

If CmbFlag.Text = "" Then
   MsgBox "Blank Flag Not Allowed !!! "
   CmbFlag.SetFocus
   Exit Sub
End If

If Edit_Flg = "E" Then
    If TxtRemarks = "" Then
     If LCase(Left(CmbFlag.Text, 1)) = "d" Then
        MsgBox "Blank remark not allowed.", vbOKOnly + vbInformation, App.Title
        TxtRemarks.SetFocus
        Exit Sub
     End If
    End If
End If

If Edit_Flg = "A" And MSHFlexGrid2.Rows = 2 Then
    If Val(TxtStackFrom) + (Val(TxtNoOfStack) - 1) > 9999999999# Then
        MsgBox "More then 10 digit Stack No not allowed.", vbOKOnly + vbInformation, App.Title
        TxtStackFrom.SetFocus
        Exit Sub
    End If
End If

If Edit_Flg = "A" Then
    If Edit_Flg = "A" And MSHFlexGrid2.Rows = 2 Then
         ''-- Bulk Stack Entry
         mNoofStack = Val(TxtNoOfStack)
         mStackFrom = Val(TxtStackFrom)
         mStackCount = 0
         
        tmp = " SELECT * FROM Mst_Stack WHERE GodownID = " & mGodownID & " AND Isnumeric(StackNo) = 1 AND Convert(Decimal,StackNo) BETWEEN " & Val(TxtStackFrom) & " AND " & Val(TxtStackFrom) + (Val(TxtNoOfStack) - 1)
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
           MsgBox "Some duplicate Stack No found in given rang.", vbOKOnly + vbInformation, App.Title
           TxtNoOfStack.SetFocus
           Exit Sub
        End If
         
         Me.MousePointer = vbHourglass
         For mStackCount = 0 To mNoofStack - 1
            
            RsMst_Stack.AddNew
            RsMst_Stack!StackID = GetMaxStackID
            'TxtStackId = RsMst_Stack!StackID
            RsMst_Stack!G_UID = GetGUID
            RsMst_Stack!GodownID = mGodownID
            RsMst_Stack!StackNo = Val(mStackFrom) + mStackCount
            RsMst_Stack!Flag = Left(CmbFlag.Text, 1)
            RsMst_Stack!Remark = Trim(TxtRemarks)
            If IsNull(RsMst_Stack!CreatedBy) = True Or RsMst_Stack!CreatedBy = "" Then
               RsMst_Stack!CreatedBy = Gen_UserLoginID
               RsMst_Stack!CreatedDate = Now
            Else
               RsMst_Stack!UpdatedBy = Gen_UserLoginID
               RsMst_Stack!UpdatedDate = Now
            End If
            
            RsMst_Stack.Update
            'RsMst_Stack.Requery
            
         Next mStackCount
         Me.MousePointer = vbDefault
         
         TxtStackId = RsMst_Stack!StackID
         RsMst_Stack.Requery
         RsMst_Stack.MoveFirst
         RsMst_Stack.Find "StackID = " & TxtStackId
         Call Indata
         Exit Sub
         ''------------------
    ElseIf Edit_Flg = "A" And MSHFlexGrid2.Rows > 2 Then
        tmp = " SELECT * FROM Mst_Stack WHERE GodownID = " & mGodownID & " AND StackNo = '" & Trim(TxtStackNo) & "'"
        Call TmpTable
        If TmpRs.RecordCount > 0 Then
           MsgBox "Duplicate Entry found for same Godown.", vbOKOnly + vbInformation, App.Title
           TxtStackNo.SetFocus
           Exit Sub
        End If
        RsMst_Stack.AddNew
        RsMst_Stack!StackID = GetMaxStackID
        TxtStackId = RsMst_Stack!StackID
        RsMst_Stack!G_UID = GetGUID
    End If
Else
    'tmp = " SELECT * FROM Mst_Stack WHERE LocationID = " & mLocationID & " AND CMPID = " & mCMPID & " AND GodownID = " & mGodownID & " "
    tmp = " SELECT * FROM Mst_Stack WHERE GodownID = " & mGodownID & " AND StackNo = '" & Trim(TxtStackNo) & "' "
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
       If TmpRs!StackID <> Val(TxtStackId.Text) Then
        MsgBox "Duplicate Entry found for same Godown.", vbOKOnly + vbInformation, App.Title
        TxtStackNo.SetFocus
        Exit Sub
       End If
    End If
End If

''RsMst_Stack!LocationID = mLocationID
''RsMst_Stack!CMPID = mCMPID
RsMst_Stack!GodownID = mGodownID
RsMst_Stack!StackNo = Trim(TxtStackNo)
RsMst_Stack!Flag = Left(CmbFlag.Text, 1)
RsMst_Stack!Remark = Trim(TxtRemarks)

If IsNull(RsMst_Stack!CreatedBy) = True Or RsMst_Stack!CreatedBy = "" Then
   RsMst_Stack!CreatedBy = Gen_UserLoginID
   RsMst_Stack!CreatedDate = Now
Else
   RsMst_Stack!UpdatedBy = Gen_UserLoginID
   RsMst_Stack!UpdatedDate = Now
End If

RsMst_Stack.Update
RsMst_Stack.Requery

RsMst_Stack.MoveFirst
RsMst_Stack.Find "StackID = " & TxtStackId
Call Indata

End Sub

Private Function GetMaxStackID() As Double
Dim Retval As Double
tmp = "Select MAX(StackID) from Mst_Stack"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxStackID = Retval
End Function

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_Stack.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_Stack.EOF Then
      RsMst_Stack.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 650
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 400
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1400
RsMst_Stack.MoveFirst

End Sub

''---- Search
Private Sub CmbSLocationID_GotFocus()
tmp = CmbSLocationID.Text
If Gen_WcLocation = "" Then
   Call TableMst_Location
Else
   Call TableMst_Location(" Where " & Gen_WcLocation)
End If
CmbSLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
   CmbSLocationID.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbSLocationID.Text = tmp
End Sub
Private Sub CmbSLocationID_LostFocus()
If CmbSLocationID.Text = "" Then
   CmbSCMPID.Text = ""
   CmbSGodownID.Text = ""
   Exit Sub
End If

If tmp <> CmbSLocationID.Text Then
   CmbSCMPID.Text = ""
   CmbSGodownID.Text = ""
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbSLocationID.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSCMPID_GotFocus()
If CmbSLocationID = "" Then
    MsgBox "Select Location.", vbOKOnly + vbInformation, App.Title
    CmbSLocationID.SetFocus
    Exit Sub
End If

tmp = CmbSCMPID.Text
If Gen_WcCMP <> "" Then
    Call TableMst_CMP(" Where CloseDate Is Null AND " & Gen_WcCMP & " AND  LocationID = " & mSLocationID & " ")
Else
    Call TableMst_CMP(" Where CloseDate Is Null AND LocationID = " & mSLocationID & " ")
End If
CmbSCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
   CmbSCMPID.AddItem RsMst_CMP!CMPName
   RsMst_CMP.MoveNext
Next
CmbSCMPID.Text = tmp
End Sub
Private Sub CmbSCMPID_LostFocus()
If CmbSCMPID.Text = "" Then
   CmbSGodownID.Text = ""
   Exit Sub
End If

If tmp <> CmbSCMPID.Text Then
   CmbSGodownID.Text = ""
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbSCMPID.SetFocus
   Exit Sub
End If
mSCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbSGodownID_GotFocus()
If CmbSCMPID = "" Then
    MsgBox "Select CMP.", vbOKOnly + vbInformation, App.Title
    CmbSCMPID.SetFocus
    Exit Sub
End If

tmp = CmbSGodownID.Text
Call TableMst_Godown(" Where CloseDate Is Null AND CMPID = " & mSCMPID & " ")
CmbSGodownID.Clear
If RsMst_Godown.RecordCount = 0 Then
   MsgBox "No Record found!!!"
   Exit Sub
End If
For i = 1 To RsMst_Godown.RecordCount
   CmbSGodownID.AddItem RsMst_Godown!GodownNo
   RsMst_Godown.MoveNext
Next
CmbSGodownID.Text = tmp
End Sub
Private Sub CmbSGodownID_LostFocus()
If CmbSGodownID = "" Then Exit Sub

RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownNo = '" & CmbSGodownID.Text & "'"
If RsMst_Godown.EOF Then
   MsgBox "Invalid  selection!!!"
   CmbSGodownID.SetFocus
   Exit Sub
End If
mSGodownID = RsMst_Godown!GodownID
End Sub

Private Sub TxtNoofStack_LostFocus()
    If TxtNoOfStack = "" Then Exit Sub
    If IsNumeric(TxtNoOfStack) = False Then
        MsgBox "Only numbers allowed.", vbOKOnly + vbInformation, App.Title
        TxtNoOfStack.SetFocus
        Exit Sub
    End If
    If Trim(TxtNoOfStack) = "0" Then
        MsgBox "Zero numbers of stack not allowed.", vbOKOnly + vbInformation, App.Title
        TxtNoOfStack.SetFocus
        Exit Sub
    End If
End Sub

Private Sub TxtStackFrom_LostFocus()
    If TxtStackFrom = "" Then Exit Sub
    If IsNumeric(TxtStackFrom) = False Then
        MsgBox "Only Numeric Stack Numbers allowed.", vbOKOnly + vbInformation, App.Title
        TxtStackFrom.SetFocus
        Exit Sub
    End If
    If Trim(TxtStackFrom) = "0" Then
        MsgBox "Stack No can not start from Zero.", vbOKOnly + vbInformation, App.Title
        TxtStackFrom.SetFocus
        Exit Sub
    End If
End Sub

Private Sub TxtStackNo_LostFocus()
''-- Check for duplicate
''-- if found then going to Indata
'tmp = " SELECT * FROM Mst_Stack WHERE LocationID = " & mLocationID & " AND CMPID  = " & mCMPID & " AND GodownID = " & mGodownID & " AND StackID <> " & Val(TxtStackId) & " "

If IsClean(TxtStackNo) = False Then
    MsgBox "Special characters not allowed.", vbOKOnly + vbInformation, App.Title
    TxtStackNo.SetFocus
    Exit Sub
End If

tmp = " SELECT * FROM Mst_Stack WHERE GodownID = " & Val(mGodownID) & " AND StackNo = '" & Trim(TxtStackNo) & "' AND StackID <> " & Val(TxtStackId) & " "
Call TmpTable
If TmpRs.RecordCount > 0 Then
   MsgBox "Duplicate Entry found for same Godown.", vbOKOnly + vbInformation, App.Title
   TxtStackNo.SetFocus
   Exit Sub
'   If RsMst_Stack.RecordCount > 0 Then
'      RsMst_Stack.MoveFirst
'      RsMst_Stack.Find "StackID = " & TmpRs!StackID
'      If RsMst_Stack.EOF = False Then
'        Call Indata
'      End If
'   End If
End If
End Sub
