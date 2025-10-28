VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_CommtrackInvoice 
   Caption         =   "CommtrackInvoice"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   45
      Top             =   9600
      Width           =   15015
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9975
      Left            =   15000
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame Frame4 
      Height          =   9255
      Left            =   120
      TabIndex        =   46
      Top             =   120
      Width           =   14895
      Begin VB.Frame Frame0 
         Height          =   1575
         Left            =   360
         TabIndex        =   47
         Top             =   5640
         Visible         =   0   'False
         Width           =   14130
         Begin VB.TextBox TxtSPostInvoiceNo 
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
            IMEMode         =   3  'DISABLE
            Left            =   11520
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   473
            Width           =   2535
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   350
            Left            =   12330
            TabIndex        =   39
            Top             =   885
            Width           =   1000
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   13335
            TabIndex        =   40
            Top             =   885
            Width           =   1000
         End
         Begin VB.ComboBox CmbSLocation 
            Height          =   315
            Left            =   2985
            Sorted          =   -1  'True
            TabIndex        =   35
            Top             =   473
            Width           =   3045
         End
         Begin VB.ComboBox CmbNonNcdexClient 
            Height          =   315
            Left            =   6090
            Sorted          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   473
            Width           =   3390
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2040
            TabIndex        =   38
            Top             =   900
            Width           =   225
         End
         Begin VB.ComboBox CmbSStatus 
            Height          =   315
            ItemData        =   "FrmTxn_CommtrackInvoice.frx":0000
            Left            =   9600
            List            =   "FrmTxn_CommtrackInvoice.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   37
            Top             =   473
            Width           =   1875
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1095
            Left            =   120
            TabIndex        =   48
            Top             =   1320
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   1931
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   360
            Left            =   1575
            TabIndex        =   34
            Top             =   450
            Width           =   1245
            _ExtentX        =   2196
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
            Format          =   109838337
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   360
            Left            =   120
            TabIndex        =   33
            Top             =   450
            Width           =   1245
            _ExtentX        =   2196
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
            Format          =   109838337
            CurrentDate     =   36494
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Post Invoice No"
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
            Left            =   12082
            TabIndex        =   85
            Top             =   240
            Width           =   1410
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
            Left            =   9315
            TabIndex        =   56
            Top             =   900
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
            Left            =   6840
            TabIndex        =   55
            Top             =   900
            Width           =   2415
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Location"
            Height          =   195
            Left            =   4155
            TabIndex        =   54
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "CP"
            Height          =   195
            Left            =   7575
            TabIndex        =   53
            Top             =   240
            Width           =   210
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "To Date"
            Height          =   195
            Left            =   1845
            TabIndex        =   52
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "From Date"
            Height          =   195
            Left            =   405
            TabIndex        =   51
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label30 
            Caption         =   "Approved By FAG"
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
            Top             =   900
            Width           =   1680
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Reverse Flag"
            Height          =   195
            Left            =   10065
            TabIndex        =   49
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H008080FF&
         Caption         =   "Reverse Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   4920
         TabIndex        =   57
         Top             =   3000
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox TxtReverseRemark 
            Appearance      =   0  'Flat
            Height          =   480
            Left            =   1560
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   840
            Width           =   4680
         End
         Begin VB.CommandButton CmdSaveReverse 
            BackColor       =   &H00E0E0E0&
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
            Left            =   1575
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1440
            Width           =   1440
         End
         Begin VB.CommandButton CmdCancelReverse 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cancel"
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
            Left            =   3015
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   1440
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker TxtReverseDate 
            Height          =   375
            Left            =   1560
            TabIndex        =   41
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   109838337
            CurrentDate     =   39884
         End
         Begin VB.Label Label5 
            BackColor       =   &H008080FF&
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
            Height          =   300
            Left            =   120
            TabIndex        =   59
            Top             =   930
            Width           =   1260
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Reverse Date"
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
            TabIndex        =   58
            Top             =   390
            Width           =   1800
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1065
         Left            =   120
         TabIndex        =   81
         Top             =   8070
         Width           =   14625
         Begin VB.CommandButton CmdPrint 
            Caption         =   "Print Annexure"
            Height          =   795
            Left            =   12285
            TabIndex        =   83
            Top             =   195
            Width           =   2100
         End
         Begin VB.CommandButton CmdViewInvoice 
            Caption         =   "View Invoice"
            Height          =   400
            Left            =   10575
            TabIndex        =   32
            Top             =   600
            Width           =   1710
         End
         Begin VB.CommandButton CmdReverse 
            Caption         =   "Reverse"
            Height          =   400
            Left            =   10575
            TabIndex        =   26
            Top             =   195
            Width           =   1710
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   8505
            TabIndex        =   31
            Top             =   600
            Width           =   2070
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   6435
            TabIndex        =   30
            Top             =   600
            Width           =   2070
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   4365
            TabIndex        =   29
            Top             =   600
            Width           =   2070
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   2295
            TabIndex        =   28
            Top             =   600
            Width           =   2070
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   225
            TabIndex        =   27
            Top             =   600
            Width           =   2070
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   225
            TabIndex        =   21
            Top             =   195
            Width           =   2070
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   2295
            TabIndex        =   22
            Top             =   195
            Width           =   2070
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   4365
            TabIndex        =   23
            Top             =   195
            Width           =   2070
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   6435
            TabIndex        =   24
            Top             =   195
            Width           =   2070
         End
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   8505
            TabIndex        =   25
            Top             =   195
            Width           =   2070
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4965
         Left            =   120
         TabIndex        =   60
         Top             =   120
         Width           =   14655
         Begin VB.TextBox TxtPostInvoiceNo 
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
            IMEMode         =   3  'DISABLE
            Left            =   1935
            Locked          =   -1  'True
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   4440
            Width           =   5415
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
            Height          =   645
            Left            =   1920
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   2625
            Width           =   12570
         End
         Begin VB.TextBox TxtServiceTax 
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
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   2205
            Width           =   5280
         End
         Begin VB.TextBox TxtBalAmt 
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
            IMEMode         =   3  'DISABLE
            Left            =   10890
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   4005
            Width           =   1815
         End
         Begin VB.TextBox TxtEntryNo 
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
            IMEMode         =   3  'DISABLE
            Left            =   12705
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   4005
            Width           =   1815
         End
         Begin VB.TextBox TxtRecAmt 
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
            IMEMode         =   3  'DISABLE
            Left            =   9075
            Locked          =   -1  'True
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   4005
            Width           =   1815
         End
         Begin VB.TextBox TxtTISTAmt 
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
            IMEMode         =   3  'DISABLE
            Left            =   7380
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   4005
            Width           =   1695
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
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   3315
            Width           =   5250
         End
         Begin VB.TextBox TxtTINo 
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
            IMEMode         =   3  'DISABLE
            Left            =   1935
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   4005
            Width           =   1815
         End
         Begin VB.TextBox TxtClient 
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
            IMEMode         =   3  'DISABLE
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1208
            Width           =   5295
         End
         Begin VB.TextBox TxtInvoiceNo 
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
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TxtInvoiceDate 
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
            Left            =   5385
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox TxtTIDate 
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
            IMEMode         =   3  'DISABLE
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   4005
            Width           =   1815
         End
         Begin VB.TextBox TxtBillingAmount 
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
            IMEMode         =   3  'DISABLE
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1740
            Width           =   1815
         End
         Begin VB.TextBox TxtTIAmount 
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
            IMEMode         =   3  'DISABLE
            Left            =   5565
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   4005
            Width           =   1815
         End
         Begin VB.TextBox TxtLocation 
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
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   728
            Width           =   5295
         End
         Begin VB.TextBox TxtInvoiceMonth 
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
            Left            =   9120
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox TxtNAVCode 
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
            IMEMode         =   3  'DISABLE
            Left            =   9120
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1208
            Width           =   1815
         End
         Begin VB.TextBox TxtServiceTaxAmount 
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
            IMEMode         =   3  'DISABLE
            Left            =   5385
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1740
            Width           =   1815
         End
         Begin VB.TextBox TxtTotalAmount 
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
            IMEMode         =   3  'DISABLE
            Left            =   9120
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1740
            Width           =   1815
         End
         Begin VB.TextBox TxtBatchID 
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
            IMEMode         =   3  'DISABLE
            Left            =   9120
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2198
            Width           =   1815
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Post Invoice No"
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
            TabIndex        =   87
            Top             =   4507
            Width           =   1410
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Balance Amount"
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
            Left            =   11070
            TabIndex        =   80
            Top             =   3735
            Width           =   1470
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Entry No"
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
            Left            =   13200
            TabIndex        =   79
            Top             =   3735
            Width           =   765
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Received Amount"
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
            Left            =   9180
            TabIndex        =   78
            Top             =   3735
            Width           =   1605
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "TI ST Amount"
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
            Left            =   7680
            TabIndex        =   77
            Top             =   3735
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Invoice No"
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
            TabIndex        =   76
            Top             =   307
            Width           =   960
         End
         Begin VB.Label Label9 
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
            Left            =   120
            TabIndex        =   75
            Top             =   795
            Width           =   765
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "CP ID and Name"
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
            TabIndex        =   74
            Top             =   1275
            Width           =   1500
         End
         Begin VB.Label Label2 
            Caption         =   "Invoice Month and Year"
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
            Left            =   7680
            TabIndex        =   73
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Billing Amount"
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
            TabIndex        =   72
            Top             =   1807
            Width           =   1260
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "TI No"
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
            Left            =   2610
            TabIndex        =   71
            Top             =   3735
            Width           =   495
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "TI Amount"
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
            Left            =   5880
            TabIndex        =   70
            Top             =   3735
            Width           =   900
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Invoice Date"
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
            TabIndex        =   69
            Top             =   307
            Width           =   1125
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "TI Date"
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
            TabIndex        =   68
            Top             =   3735
            Width           =   660
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "NAV Code"
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
            Left            =   7680
            TabIndex        =   66
            Top             =   1275
            Width           =   960
         End
         Begin VB.Label Label46 
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
            Left            =   120
            TabIndex        =   64
            Top             =   3382
            Width           =   855
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
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
            TabIndex        =   63
            Top             =   2947
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label3 
            Caption         =   "Service Tax Amount"
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
            Left            =   3840
            TabIndex        =   62
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Total Amount"
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
            Left            =   7680
            TabIndex        =   61
            Top             =   1807
            Width           =   1185
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Batch ID"
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
            Left            =   7680
            TabIndex        =   67
            Top             =   2265
            Width           =   750
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Service Tax"
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
            TabIndex        =   65
            Top             =   2265
            Width           =   1080
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   2925
         Left            =   120
         TabIndex        =   82
         Top             =   5160
         Width           =   14625
         _ExtentX        =   25797
         _ExtentY        =   5159
         _Version        =   393216
         Rows            =   3
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
End
Attribute VB_Name = "FrmTxn_CommtrackInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSLocationID, mDepositoryID, mMonthListindex, mSM_Prefix, mLocationID, mCPID, mPostInvoiceNo As Variant

Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text
Call TableMst_Depository("")
CmbNonNcdexClient.Clear
If RsMst_Depository.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Depository.RecordCount
   CmbNonNcdexClient.AddItem RsMst_Depository!DepositoryName
   RsMst_Depository.MoveNext
Next
CmbNonNcdexClient.Text = tmp
End Sub

Private Sub CmbNonNcdexClient_LostFocus()
If CmbNonNcdexClient.Text = "" Then
   Exit Sub
End If
RsMst_Depository.MoveFirst
RsMst_Depository.Find "DepositoryName = '" & CmbNonNcdexClient.Text & "'"
If RsMst_Depository.EOF Then
   MsgBox "Invalid selection "
   CmbNonNcdexClient.SetFocus
   Exit Sub
End If
mDepositoryID = RsMst_Depository!DepositoryID
End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text

If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location
End If
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp
End Sub

Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
mSLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbSStatus_LostFocus()
If CmbSStatus.Text = "" Then Exit Sub

If LCase(CmbSStatus.Text) <> "yes" And LCase(CmbSStatus.Text) <> "no" Then
   MsgBox "Invalid selection!!!"
   CmbSStatus.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdCancelReverse_Click()
TxtReverseDate.Value = Date
TxtReverseRemark = ""
Frame6.Visible = False
ButtonFrm.Enabled = True
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_CommtrackInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_CommtrackInvoice.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
oWS.Range("a:a").ColumnWidth = 11.29
oWS.Range("b:b").ColumnWidth = 11.29
oWS.Range("c:c").ColumnWidth = 10.43
oWS.Range("d:d").ColumnWidth = 13.57
oWS.Range("e:e").ColumnWidth = 40
oWS.Range("f:f").ColumnWidth = 9.57
oWS.Range("g:g").ColumnWidth = 13.57
oWS.Range("h:h").ColumnWidth = 9.86
oWS.Range("i:i").ColumnWidth = 12.43
oWS.Range("j:j").ColumnWidth = 8
oWS.Range("k:k").ColumnWidth = 12.43
oWS.Range("l:l").ColumnWidth = 9.14
oWS.Range("m:m").ColumnWidth = 9.86
oWS.Range("n:n").ColumnWidth = 9
oWS.Range("o:o").ColumnWidth = 9

tmp = "a" & mRow & ":" & "o" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, C + 1) = "Commtrack Invoice"
oWS.Cells(mRow, C + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, C + 1).Font.Bold = True
oWS.Cells(mRow, C + 1).Font.Size = 14
oWS.Cells(mRow, C + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, C + 1).WrapText = True

For I = 0 To MSHFlexGrid1.Rows - 2
    mRow = mRow + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        If I = 0 Then
           oWS.Cells(mRow, C + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, C + 1).Font.Bold = True
           oWS.Cells(mRow, C + 1).RowHeight = 30
        End If
        oWS.Cells(mRow, C + 1).WrapText = True
        oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
    Next
Next

oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGo_Click()

Call Grid_Head


If TxtFromDate.Value > TxtToDate Then
    MsgBox "From Date should not greater than To Date!!!"
    TxtToDate.SetFocus
    Exit Sub
End If

mFromDate = Format(TxtFromDate.Value, "yyyy-mm-dd")
mToDate = Format(TxtToDate.Value, "yyyy-mm-dd")

tmp = " Select Distinct TMCI.InvoiceID, TMCI.MonthYear, ML.LocationName, TMCI.CPNavCode,"
tmp = tmp & " TMCI.Amt, TMCI.STAmt, TMCI.TotalAmt, TMCI.PINo, TMCI.BatchID, TMCI.Flag, TMCI.InvoiceNo, TMCI.InvoiceDate,"
tmp = tmp & " TMCI.InvSTAmt,TMCI.InvoiceAmt, TMCI.CreatedDate,"
tmp = tmp & " (Select top 1 MD.DepositoryName from Mst_Depository MD Where TMCI.CPNavCode = MD.NavCode) DepositoryName,"
tmp = tmp & " (Select top 1 MD.DepositoryID from Mst_Depository MD Where TMCI.CPNavCode = MD.NavCode) DepositoryID,"
tmp = tmp & " TMCI.CPName,TMCI.CPID ,TI.TIID , TI.TINo, TI.TIDate, TI.TIAmount, TI.PostInvoiceNo, TI.EntryNo, TI.TISTAmount,TMCI.ReverseFlag,MSI.StateIndiaName ,MSI.GSTIN 'NCMLGSTNo',MCSG.GSTNo 'ClientGSTNo',MCSG.GSTNavCode,'' 'TaxApplied'"
tmp = tmp & " From Txn_MonthlyCommtrackInvoice TMCI"
tmp = tmp & " LEFT JOIN Mst_Location ML On TMCI.LocationID = ML.LocationID"
tmp = tmp & " LEFT JOIN Mst_Depository MD On MD.NavCode = TMCI.CPNavCode"
tmp = tmp & " LEFT JOIN Txn_TIDetails TI On TI.PIID = TMCI.PINo And PIFlag ='Commtrack'"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID "
tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID "
tmp = tmp & " Left  Join Mst_DepositoryStateWiseGST MCSG On TMCI.CPID = MCSG.DepositoryID ANd MSI.StateIndiaID = MCSG.StateINdiaID "


If Gen_DBType = "MDB" Then
   tmp = tmp & " Where TMCI.MonthYear Between #" & mFromDate & "# And #" & mToDate & "# " 'And ExchangeTypeID = 1 And (TRI.ContractID = 0 Or TRI.ContractID is Null)"
Else
   tmp = tmp & " Where TMCI.MonthYear Between '" & mFromDate & "' And '" & mToDate & "'" 'And ExchangeTypeID = 1 And (TRI.ContractID = 0 Or TRI.ContractID is Null)"
End If
If CmbSLocation.Text <> "" Then
   tmp = tmp & " And ML.LocationID = " & mSLocationID
   Else
   If Gen_WcLocation <> "" Then
      tmp = tmp & " And ML." & Gen_WcLocation
   End If
End If
If CmbNonNcdexClient.Text <> "" Then
   tmp = tmp & " And MD.DepositoryID = '" & mDepositoryID & "'"
End If
If ChkBatchNo.Value = 0 Then
   tmp = tmp & " And TMCI.BatchID = 0 "
End If
If CmbSStatus.Text <> "" Then
   tmp = tmp & " And TMCI.ReverseFlag = '" & Left$(CmbSStatus, 1) & "'"
End If

If Trim(TxtSPostInvoiceNo) <> "" Then
   tmp = tmp & " And TI.PostInvoiceNo like '%" & Trim(TxtSPostInvoiceNo) & "%' "
End If
tmp = tmp & " Order By TMCI.InvoiceID"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!InvoiceID), "", TmpRs1!InvoiceID)
       MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!CreatedDate), "", Format(TmpRs1!CreatedDate, "dd-MMM-yyyy")) 'InvoiceDate
       MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!MonthYear), "", Format(TmpRs1!MonthYear, "MMM-yyyy")) 'MonthYear
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!Locationname), "", TmpRs1!Locationname) 'Location
       If IsNull(TmpRs1!CPID) = True Then
          MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!DepositoryName), "", TmpRs1!DepositoryName + "~" + TmpRs1!DepositoryID) 'Depository
       Else
          MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!CPName), "", TmpRs1!CPName + "~" + TmpRs1!CPID) 'Depository
       End If
       MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs1!CPNAVCode), "", TmpRs1!CPNAVCode) 'NAVCode
       MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!Amt), "", TmpRs1!Amt)
       MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!STAmt), "", TmpRs1!STAmt)
       MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs1!TotalAmt), "", TmpRs1!TotalAmt)
       MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs1!BatchID), "", TmpRs1!BatchID)
       MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs1!TINo), "", TmpRs1!TINo)
       MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs1!TIDate), "", Format(TmpRs1!TIDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs1!TIAmount), "", TmpRs1!TIAmount)
       MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs1!TISTAmount), "", TmpRs1!TISTAmount)
       MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs1!EntryNo), "", TmpRs1!EntryNo)
       MSHFlexGrid1.TextMatrix(I, 15) = IIf(IsNull(TmpRs1!PostInvoiceNo), "", TmpRs1!PostInvoiceNo)
       MSHFlexGrid1.TextMatrix(I, 16) = IIf(IsNull(TmpRs1!ReverseFlag), "N", TmpRs1!ReverseFlag)
       MSHFlexGrid1.TextMatrix(I, 17) = IIf(IsNull(TmpRs1!StateIndiaName), "", TmpRs1!StateIndiaName)
       MSHFlexGrid1.TextMatrix(I, 18) = IIf(IsNull(TmpRs1!NCMLGSTNo), "", TmpRs1!NCMLGSTNo)
       MSHFlexGrid1.TextMatrix(I, 19) = IIf(IsNull(TmpRs1!ClientGSTNo), "", TmpRs1!ClientGSTNo)
       MSHFlexGrid1.TextMatrix(I, 20) = IIf(IsNull(TmpRs1!GSTNavCode), "", TmpRs1!GSTNavCode)
       If Trim(MSHFlexGrid1.TextMatrix(I, 18)) <> "" And Trim(MSHFlexGrid1.TextMatrix(I, 19)) <> "" Then
          If Left(MSHFlexGrid1.TextMatrix(I, 18), 2) = Left(MSHFlexGrid1.TextMatrix(I, 19), 2) Then
             MSHFlexGrid1.TextMatrix(I, 21) = "Intra"
          Else
             MSHFlexGrid1.TextMatrix(I, 21) = "Inter"
          End If
       End If
       
       TmpRs1.MoveNext
   Next
End If

Label22.Caption = ""
Label22.Caption = TmpRs1.RecordCount
Label22.Refresh

End Sub

Private Sub CmdPrint_Click()

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "CommtrackInvoice-" & TxtInvoiceNo & "-Annexure.xls")
Call Xls_Detail_RptAnnexure

End Sub

Private Sub CmdReverse_Click()
Frame6.Visible = True
ButtonFrm.Enabled = False
TxtReverseDate.Value = Date
TxtReverseRemark = ""
End Sub

Private Sub CmdSaveReverse_Click()
If TxtReverseDate.Value > Date Then
   MsgBox "Future Date not allowed!!"
   TxtReverseDate.SetFocus
   Exit Sub
End If
If Format(Gen_BlockingDate, "yyyy-mm") >= Format(TxtReverseDate.Value, "yyyy-mm") Then
   MsgBox "Reverse Date allowed for next month of " & Format(Gen_BlockingDate, "MMM-yyyy") & "!!! "
   TxtReverseDate.SetFocus
   Exit Sub
End If
If TxtReverseRemark.Text = "" Then
   MsgBox "Blank Remark not allowed!!"
   TxtReverseRemark.SetFocus
   Exit Sub
End If
TxtReverseRemark = Replace(TxtReverseRemark, ",", "-")
RsTxn_MonthlyCommtrackInvoice!ReverseDate = TxtReverseDate.Value
RsTxn_MonthlyCommtrackInvoice!ReverseFlag = "Y"
RsTxn_MonthlyCommtrackInvoice!ReverseRemark = TxtReverseRemark.Text
RsTxn_MonthlyCommtrackInvoice!UpdatedBy = Gen_UserLoginID
RsTxn_MonthlyCommtrackInvoice!UpdatedDate = Now
RsTxn_MonthlyCommtrackInvoice.Update

tmp = " Update Txn_TIDetails Set TIAmount=0, ReverseFlag='Y',UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = getdate() Where PIID=" & RsTxn_MonthlyCommtrackInvoice!PINo & " and PIFLag='Commtrack'"
Call TmpTable

Frame6.Visible = False
ButtonFrm.Enabled = True
CmdReverse.Enabled = False
End Sub

Private Sub CmdViewInvoice_Click()
'On Error GoTo Error:
''Shell (App.Path & "\Invoices\" & TxtEntryNo & " ")
'If mEntryNo <> "" Then
'   wsPathAdobe = "\\10.0.0.8\Reader\AcroRD32.exe" '"C:\Program Files\Adobe\Acrobat 6.0\Reader\AcroRd32.exe" '"\\10.0.0.8\Reader\AcroRD32.exe"
'   mFileName = App.Path & "\NavisionInvoice\" & TxtEntryNo & ".pdf"
'   'mFileName = App.Path & "\Invoices\" & Replace(TxtPostInvoiceNo, "/", "") & ".pdf"
'   If wsPathAdobe <> "" Then
'      iRet = Shell(wsPathAdobe & " " & mFileName, vbMaximizedFocus)
'   End If
'End If
'
'Exit Sub
'Error:
'      MsgBox "Please contact Technology department!!!"



Dim fso As New FileSystemObject
On Error GoTo Error:
'Shell (App.Path & "\Invoices\" & TxtEntryNo & " ")
If TxtEntryNo <> "" Then
   wsPathAdobe = Gen_ReadPathConfigINI("PDFPATH") '"\\10.0.0.8\Reader\AcroRD32.exe" '"C:\Program Files\Adobe\Acrobat 6.0\Reader\AcroRd32.exe" '"\\10.0.0.8\Reader\AcroRD32.exe"
   mFileName = App.Path & "\NavisionInvoice\" & TxtEntryNo & ".pdf"
   'mFileName = App.Path & "\Invoices\" & Replace(TxtPostInvoiceNo, "/", "") & ".pdf"
   If fso.FileExists(mFileName) = True Then
      If wsPathAdobe <> "" Then
         iRet = Shell(wsPathAdobe & " " & mFileName, vbMaximizedFocus)
         Call Log_Print_History(mSM_Prefix, "COMM", TxtOriginalInvoiceNo, TxtEntryNo)
      End If
   Else
      MsgBox "Invoice not found. Please contact CMG Department for same!!!"
   End If
End If
Exit Sub

Error:
      MsgBox "Please contact Technology department!!!"


End Sub

Private Sub EditCmd_Click()
Frame1.Enabled = True
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_MonthlyCommtrackInvoice.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
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


Private Sub Form_Load()
Dim wc As Variant

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   CmdReverse.Enabled = False
   CmdViewInvoice.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head

'CmbMonth.Text = ""
'CmbYear.Text = ""
TxtFromDate.Value = Date - 30
TxtToDate.Value = Date
CmbSLocation.Text = ""
CmbNonNcdexClient.Text = ""
Label22.Caption = ""

If Gen_WcLocation <> "" Then
    Call TableTxn_MonthlyCommtrackInvoice(" Where " & Gen_WcLocation & "  Order by InvoiceID")
Else
    Call TableTxn_MonthlyCommtrackInvoice(" Order by InvoiceID")
End If
If RsTxn_MonthlyCommtrackInvoice.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call GButtonLockAD(Me)
   Exit Sub
End If

Call GButtonLock(Me, True)
Call Indata
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 22 '17 '16
    .Font.Size = 10
    .TextMatrix(0, 0) = "Invoice No"
    .TextMatrix(0, 1) = "Invoice Date"
    .TextMatrix(0, 2) = "Year Month"
    .TextMatrix(0, 3) = "Location Name"
    .TextMatrix(0, 4) = "Depository Name"
    .TextMatrix(0, 5) = "NAVCode"
    .TextMatrix(0, 6) = "Billing Amount"
    .TextMatrix(0, 7) = "Tax Amount"
    .TextMatrix(0, 8) = "Total Amount"
    .TextMatrix(0, 9) = "Batch ID"
    .TextMatrix(0, 10) = "TI No"
    .TextMatrix(0, 11) = "TI Date"
    .TextMatrix(0, 12) = "TI Amount"
    .TextMatrix(0, 13) = "TI ST Amount"
    .TextMatrix(0, 14) = "TI Entry No"
    .TextMatrix(0, 15) = "TI Post Invoice No"
    .TextMatrix(0, 16) = "Reverse Flag"
    .TextMatrix(0, 17) = "State India Name"
    .TextMatrix(0, 18) = "NCML GST No"
    .TextMatrix(0, 19) = "Client GST No"
    .TextMatrix(0, 20) = "Client GST NAVCode"
    .TextMatrix(0, 21) = "Tax Applied (Inter/Intra)"
        
        
    .ColWidth(0) = 1500
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 2500
    .ColWidth(4) = 2500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500
    .ColWidth(9) = 1000
    .ColWidth(10) = 1500
    .ColWidth(11) = 1500
    .ColWidth(12) = 1500
    .ColWidth(13) = 1500
    .ColWidth(14) = 1500
    .ColWidth(15) = 1500
    
    .RowHeight(0) = 600
    .WordWrap = True
End With
End Sub

Private Sub Indata()
Frame1.Enabled = False
CmdReverse.Enabled = False
TxtInvoiceNo.Text = RsTxn_MonthlyCommtrackInvoice!InvoiceID
Call Grid_Head2
CmdPrint.Enabled = False
If IsNull(RsTxn_MonthlyCommtrackInvoice!CPID) = True Then
   tmp = " Select Distinct TMCI.InvoiceID, TMCI.MonthYear, ML.LocationName, TMCI.CPNavCode,"
   tmp = tmp & " TMCI.Amt, TMCI.STAmt, TMCI.TotalAmt, TMCI.PINo, TMCI.BatchID, TMCI.Flag, TMCI.InvoiceNo, TMCI.InvoiceDate,"
   tmp = tmp & " TMCI.InvSTAmt,TMCI.InvoiceAmt, TMCI.CreatedDate, TMCI.ReverseRemark,ML.SM_Prefix,"
   tmp = tmp & " (Select top 1 MD.DepositoryName from Mst_Depository MD Where TMCI.CPNavCode = MD.NavCode) DepositoryName,"
   tmp = tmp & " (Select top 1 MD.DepositoryID from Mst_Depository MD Where TMCI.CPNavCode = MD.NavCode) DepositoryID,ML.LocationID"
   tmp = tmp & " From Txn_MonthlyCommtrackInvoice TMCI"
   tmp = tmp & " Inner Join Mst_Location ML On TMCI.LocationID = ML.LocationID"
   tmp = tmp & " Where TMCI.InvoiceID = " & Val(TxtInvoiceNo)
Else
   tmp = " Select Distinct TMCI.InvoiceID, TMCI.MonthYear, ML.LocationName, TMCI.CPNavCode,"
   tmp = tmp & " TMCI.Amt, TMCI.STAmt, TMCI.TotalAmt, TMCI.PINo, TMCI.BatchID, TMCI.Flag, TMCI.InvoiceNo, TMCI.InvoiceDate,"
   tmp = tmp & " TMCI.InvSTAmt,TMCI.InvoiceAmt, TMCI.CreatedDate, TMCI.ReverseRemark,ML.SM_Prefix,"
   tmp = tmp & " MD.DepositoryName,MD.DepositoryID,ML.LocationID "
   tmp = tmp & " From Txn_MonthlyCommtrackInvoice TMCI"
   tmp = tmp & " Inner Join Mst_Location ML On TMCI.LocationID = ML.LocationID"
   tmp = tmp & " Inner Join Mst_Depository MD On TMCI.CPID = MD.DepositoryID"
   tmp = tmp & " Where TMCI.InvoiceID = " & Val(TxtInvoiceNo)
End If

Call TmpTable

If TmpRs.RecordCount > 0 Then
    TxtInvoiceDate.Text = IIf(IsNull(TmpRs!CreatedDate), "", Format(TmpRs!CreatedDate, "dd-MMM-yyyy"))
    TxtInvoiceMonth.Text = IIf(IsNull(TmpRs!MonthYear), "", Format(TmpRs!MonthYear, "MMM-yyyy"))
    TxtLocation.Text = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
    mSM_Prefix = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
    mLocationID = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
    TxtClient.Text = IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName + "~" + TmpRs!DepositoryID)
    mCPID = IIf(IsNull(TmpRs!DepositoryID), "", TmpRs!DepositoryID)
    TxtNAVCode.Text = IIf(IsNull(TmpRs!CPNAVCode), "", TmpRs!CPNAVCode)
    TxtBillingAmount.Text = IIf(IsNull(TmpRs!Amt), "", TmpRs!Amt)
    TxtServiceTaxAmount.Text = IIf(IsNull(TmpRs!STAmt), "", TmpRs!STAmt)
    TxtTotalAmount.Text = IIf(IsNull(TmpRs!TotalAmt), "", TmpRs!TotalAmt)
    If Val(TxtServiceTaxAmount) > 0 Then
       TxtServiceTax.Text = "Service Tax" 'GetGLAccNAVCode("CommtrackSTCode") 'IIf(IsNull(TmpRs!InvoiceNo), "", TmpRs!InvoiceNo)
    Else
       TxtServiceTax.Text = "No Service Tax"
    End If
    TxtBatchID.Text = IIf(IsNull(TmpRs!BatchID), "", TmpRs!BatchID)
    TxtRemarks.Text = IIf(IsNull(TmpRs!ReverseRemark), "", TmpRs!ReverseRemark)
    Dim ans As Variant
    ans = GetTIDetails(Val(IIf(IsNull(TmpRs!PINo), "", TmpRs!PINo)), "Commtrack")
    TxtTINo = ans(1)
    TxtTIDate = Format(ans(2), "dd-MMM-yyyy")
    TxtTIAmount = ans(3)
    TxtEntryNo = ans(5)
    TxtTISTAmt = ans(7)
    mPostInvoiceNo = ans(4)
    TxtPostInvoiceNo = ans(4)
    TxtRecAmt = Val(ans(8)) + Val(ans(9)) + Val(ans(10))
    If Val(TxtRecAmt) > 0 Then
       TxtBalAmt = Format(Val(TxtTIAmount) - Val(TxtRecAmt), "#####0.00")
    Else
       TxtBalAmt = TxtTIAmount
    End If
    If TxtEntryNo <> "" And (LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin" Or LCase(Gen_UserRole) = "admin" Or LCase(Gen_UserRole) = "superadmin" Or LCase(Gen_UserRole) = "headcmg") Then
       If IIf(IsNull(RsTxn_MonthlyCommtrackInvoice!ReverseFlag), "N", RsTxn_MonthlyCommtrackInvoice!ReverseFlag) <> "Y" Then
          If CheckTIInCR(Val(ans(0))) = False Then
             CmdReverse.Enabled = True
          End If
       End If
    End If
    TxtCreated = IIf(IsNull(RsTxn_MonthlyCommtrackInvoice!CreatedBy), "", RsTxn_MonthlyCommtrackInvoice!CreatedBy) & " :- " & IIf(IsNull(RsTxn_MonthlyCommtrackInvoice!CreatedDate), "", RsTxn_MonthlyCommtrackInvoice!CreatedDate)
    If IsNull(RsTxn_MonthlyCommtrackInvoice!CPID) = False Then
       mFromDate = Format(Year(TmpRs!MonthYear) & "-" & Month(TmpRs!MonthYear) & "-" & "01", "yyyy-mm-dd")
       tmp = " Select Distinct a.LotId,a.CPID,a.CPName,a.ClientID,a.ClientName,a.CommodityCode,a.Qty,a.StartDate,a.EndDate,a.AmtIncST"
       tmp = tmp & " from Txn_MonthlyCommtrackFileDetail a"
       tmp = tmp & " Inner Join Mst_ISIN b On a.LOTID = b.ISINID"
       tmp = tmp & " Where MOnthYear = '" & CDate(mFromDate) & "' And b.LocationID =  " & mLocationID & " And a.CPID = '" & mCPID & "'"
       If Val(TxtServiceTaxAmount) > 0 Then
          tmp = tmp & " And Amount < AmtIncST "
       Else
          tmp = tmp & " And Amount = AmtIncST "
       End If
'a.LotId,a.CPID,a.CPName,a.ClientID,a.ClientName,a.CommodityCode,a.Qty,a.StartDate,a.EndDate,a.AmtIncST
       Call Tmp1Table
       x = 2
       If TmpRs1.RecordCount > 0 Then
          CmdPrint.Enabled = True
       End If
       For I = 1 To TmpRs1.RecordCount
           MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
           MSHFlexGrid2.TextMatrix(x, 0) = IIf(IsNull(TmpRs1!LotId), "", TmpRs1!LotId)
           MSHFlexGrid2.TextMatrix(x, 1) = IIf(IsNull(TmpRs1!CPID), "", TmpRs1!CPID)
           MSHFlexGrid2.TextMatrix(x, 2) = IIf(IsNull(TmpRs1!CPName), "", TmpRs1!CPName)
           MSHFlexGrid2.TextMatrix(x, 3) = IIf(IsNull(TmpRs1!ClientID), "", TmpRs1!ClientID)
           MSHFlexGrid2.TextMatrix(x, 4) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
           MSHFlexGrid2.TextMatrix(x, 5) = IIf(IsNull(TmpRs1!CommodityCode), "", TmpRs1!CommodityCode)
           MSHFlexGrid2.TextMatrix(x, 6) = IIf(IsNull(TmpRs1!Qty), "", TmpRs1!Qty)
           MSHFlexGrid2.TextMatrix(x, 7) = IIf(IsNull(TmpRs1!StartDate), "", Format(TmpRs1!StartDate, "dd-mmm-yyyy"))
           MSHFlexGrid2.TextMatrix(x, 8) = IIf(IsNull(TmpRs1!EndDate), "", Format(TmpRs1!EndDate, "dd-mmm-yyyy"))
           MSHFlexGrid2.TextMatrix(x, 9) = IIf(IsNull(TmpRs1!AmtIncST), "", TmpRs1!AmtIncST)
           MSHFlexGrid2.TextMatrix(x, 10) = TxtInvoiceNo
           MSHFlexGrid2.TextMatrix(x, 11) = mPostInvoiceNo
           MSHFlexGrid2.TextMatrix(0, 9) = Val(MSHFlexGrid2.TextMatrix(0, 9)) + Val(MSHFlexGrid2.TextMatrix(x, 9))
           x = x + 1
           TmpRs1.MoveNext
       Next
       
    End If
     
End If
End Sub


Private Sub LastCmd_Click()
RsTxn_MonthlyCommtrackInvoice.MoveLast
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
RsTxn_MonthlyCommtrackInvoice.MoveFirst
RsTxn_MonthlyCommtrackInvoice.Find "InvoiceID = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'"
If Not RsTxn_MonthlyCommtrackInvoice.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_MonthlyCommtrackInvoice.MoveNext
If RsTxn_MonthlyCommtrackInvoice.EOF Then
   RsTxn_MonthlyCommtrackInvoice.MoveLast
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
RsTxn_MonthlyCommtrackInvoice.MovePrevious
If RsTxn_MonthlyCommtrackInvoice.BOF Then
   RsTxn_MonthlyCommtrackInvoice.MoveFirst
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

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_MonthlyCommtrackInvoice.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_MonthlyCommtrackInvoice.MoveFirst
 
   RsTxn_MonthlyCommtrackInvoice.Find "InvoiceID = '" & TxtInvoiceNo.Text & "'"
   If RsTxn_MonthlyCommtrackInvoice.EOF Then
      RsTxn_MonthlyCommtrackInvoice.MoveLast
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Frame4.Height - 200 'Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Frame4.Width - 200
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1400

RsTxn_MonthlyCommtrackInvoice.MoveFirst

End Sub

Private Sub TxtFromDate_LostFocus()
TxtFromDate.Value = GetLastDateOfMonth(TxtFromDate.Value)
End Sub

Private Sub TxtToDate_LostFocus()
TxtToDate.Value = GetLastDateOfMonth(TxtToDate.Value)
End Sub
Private Function GetLastDateOfMonth(mDate_ As Date)
Dim mSMonth, mSYear, mSTotalDays As Variant
mSMonth = Month(mDate_)
mSYear = Year(mDate_)
mSTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
GetLastDateOfMonth = CDate(mSYear & "-" & mSMonth & "-" & mSTotalDays)
End Function

Private Sub Grid_Head2()
With MSHFlexGrid2
    .Clear
    .Rows = 3
    .FixedRows = 2
    .Cols = 12 '16
    .Font.Size = 10
    .WordWrap = True
    .TextMatrix(1, 0) = "LOT ID"
    .TextMatrix(1, 1) = "CP Id"
    .TextMatrix(1, 2) = "CP Name"
    .TextMatrix(1, 3) = "Client Id"
    .TextMatrix(1, 4) = "Client Name"
    .TextMatrix(1, 5) = "Commodity Code"
    .TextMatrix(1, 6) = "Actual Qty"
    .TextMatrix(1, 7) = "Start Date"
    .TextMatrix(1, 8) = "End Date"
    .TextMatrix(1, 9) = "Amt. Inc ST"
    .TextMatrix(1, 10) = "Proforma Invoice No"
    .TextMatrix(1, 11) = "Tax Invoice No"

        
    .ColWidth(0) = 1500
    .ColWidth(1) = 900
    .ColWidth(2) = 2400
    .ColWidth(3) = 1000
    .ColWidth(4) = 2000
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1000
    .ColWidth(10) = 700
    
    .RowHeight(0) = 600
    .WordWrap = True
End With
End Sub

Public Sub Xls_Detail_RptAnnexure()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "CommtrackInvoice-" & TxtInvoiceNo & "-Annexure.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

For I = 0 To MSHFlexGrid2.Rows - 2
    mRow = mRow + 1 ' + 6
    For C = 0 To MSHFlexGrid2.Cols - 1
        If C = 0 Then
           oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid2.TextMatrix(I, C)
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
        End If
        oWS.Cells(mRow, C + 1).WrapText = True
        oWS.Cells(mRow, C + 1).Borders.Color = RGB(0, 0, 0)
    Next
Next

oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
