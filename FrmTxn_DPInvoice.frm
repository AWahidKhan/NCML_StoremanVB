VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_DPInvoice 
   Caption         =   "DP Invoices"
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
      Height          =   9975
      Left            =   15000
      TabIndex        =   1
      Top             =   -120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   9240
      Width           =   15015
   End
   Begin VB.Frame Frame4 
      Height          =   9135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   14895
      Begin VB.Frame Frame0 
         Height          =   2415
         Left            =   0
         TabIndex        =   10
         Top             =   6960
         Visible         =   0   'False
         Width           =   14850
         Begin VB.ComboBox CmbSStatus 
            Height          =   315
            ItemData        =   "FrmTxn_DPInvoice.frx":0000
            Left            =   9600
            List            =   "FrmTxn_DPInvoice.frx":000A
            Sorted          =   -1  'True
            TabIndex        =   80
            Top             =   480
            Width           =   1875
         End
         Begin VB.CheckBox ChkBatchNo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2040
            TabIndex        =   78
            Top             =   900
            Width           =   225
         End
         Begin VB.ComboBox CmbNonNcdexClient 
            Height          =   315
            Left            =   6090
            Sorted          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   480
            Width           =   3390
         End
         Begin VB.ComboBox CmbSLocation 
            Height          =   315
            Left            =   2985
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   480
            Width           =   3045
         End
         Begin VB.CommandButton CmdExcel 
            Caption         =   "Excel"
            Height          =   350
            Left            =   13515
            TabIndex        =   16
            Top             =   852
            Width           =   1000
         End
         Begin VB.CommandButton CmdGo 
            Caption         =   "Go"
            Height          =   350
            Left            =   12450
            TabIndex        =   15
            Top             =   852
            Width           =   1000
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   1095
            Left            =   120
            TabIndex        =   17
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
            TabIndex        =   12
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
            Format          =   107413505
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   360
            Left            =   120
            TabIndex        =   11
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
            Format          =   107413505
            CurrentDate     =   36494
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Reverse Flag"
            Height          =   195
            Left            =   10065
            TabIndex        =   81
            Top             =   240
            Width           =   945
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
            TabIndex        =   79
            Top             =   900
            Width           =   1680
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "From Date"
            Height          =   195
            Left            =   405
            TabIndex        =   77
            Top             =   225
            Width           =   735
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "To Date"
            Height          =   195
            Left            =   1845
            TabIndex        =   76
            Top             =   225
            Width           =   585
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "DP"
            Height          =   195
            Left            =   7575
            TabIndex        =   21
            Top             =   225
            Width           =   225
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Location"
            Height          =   195
            Left            =   4155
            TabIndex        =   20
            Top             =   225
            Width           =   615
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
            TabIndex        =   19
            Top             =   900
            Width           =   2415
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
            TabIndex        =   18
            Top             =   900
            Width           =   90
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
         Left            =   4560
         TabIndex        =   3
         Top             =   2760
         Visible         =   0   'False
         Width           =   6375
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
            TabIndex        =   6
            Top             =   1440
            Width           =   1455
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
            TabIndex        =   5
            Top             =   1440
            Width           =   1440
         End
         Begin VB.TextBox TxtReverseRemark 
            Appearance      =   0  'Flat
            Height          =   480
            Left            =   1560
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   840
            Width           =   4680
         End
         Begin MSComCtl2.DTPicker TxtReverseDate 
            Height          =   375
            Left            =   1560
            TabIndex        =   7
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
            Format          =   107413505
            CurrentDate     =   39884
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
            TabIndex        =   9
            Top             =   390
            Width           =   1800
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
            TabIndex        =   8
            Top             =   930
            Width           =   1260
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4725
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   14655
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
            TabIndex        =   74
            Top             =   1740
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
            TabIndex        =   72
            Top             =   1740
            Width           =   1815
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
            Height          =   765
            Left            =   1920
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   70
            Top             =   2685
            Width           =   12570
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
            TabIndex        =   38
            Top             =   1208
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
            TabIndex        =   37
            Top             =   2198
            Width           =   1815
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
            TabIndex        =   36
            Top             =   240
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
            TabIndex        =   35
            Top             =   728
            Width           =   5295
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
            TabIndex        =   34
            Top             =   4245
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
            TabIndex        =   33
            Top             =   1740
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
            TabIndex        =   32
            Top             =   4245
            Width           =   1815
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
            TabIndex        =   31
            Top             =   240
            Width           =   1815
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
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
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
            TabIndex        =   29
            Top             =   1208
            Width           =   5295
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
            TabIndex        =   28
            Top             =   4245
            Width           =   1815
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
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   3555
            Width           =   5250
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
            TabIndex        =   26
            Top             =   4245
            Width           =   1695
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
            TabIndex        =   25
            Top             =   4245
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
            TabIndex        =   24
            Top             =   4245
            Width           =   1815
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
            TabIndex        =   23
            Top             =   4245
            Width           =   1815
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
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   2205
            Width           =   5280
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
            TabIndex        =   75
            Top             =   1807
            Width           =   1185
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
            TabIndex        =   73
            Top             =   1680
            Width           =   1215
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
            TabIndex        =   71
            Top             =   2947
            Width           =   975
            WordWrap        =   -1  'True
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
            TabIndex        =   56
            Top             =   3615
            Width           =   855
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
            TabIndex        =   55
            Top             =   2265
            Width           =   1080
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
            TabIndex        =   54
            Top             =   1275
            Width           =   960
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
            TabIndex        =   53
            Top             =   2265
            Width           =   750
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
            TabIndex        =   52
            Top             =   3975
            Width           =   660
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
            TabIndex        =   51
            Top             =   307
            Width           =   1125
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
            TabIndex        =   50
            Top             =   3975
            Width           =   900
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
            TabIndex        =   49
            Top             =   3975
            Width           =   495
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
            TabIndex        =   48
            Top             =   1807
            Width           =   1260
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
            TabIndex        =   47
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "DP ID and Name"
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
            TabIndex        =   46
            Top             =   1275
            Width           =   1515
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
            TabIndex        =   45
            Top             =   795
            Width           =   765
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
            TabIndex        =   44
            Top             =   307
            Width           =   960
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
            TabIndex        =   43
            Top             =   3975
            Width           =   1215
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
            TabIndex        =   42
            Top             =   3975
            Width           =   1605
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
            TabIndex        =   41
            Top             =   3975
            Width           =   765
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
            TabIndex        =   40
            Top             =   3975
            Width           =   1470
         End
      End
      Begin VB.Frame ButtonFrm 
         Height          =   1185
         Left            =   120
         TabIndex        =   57
         Top             =   4830
         Width           =   14625
         Begin VB.CommandButton SearchCmd 
            Caption         =   "Search"
            Height          =   400
            Left            =   9945
            TabIndex        =   58
            Top             =   195
            Width           =   2430
         End
         Begin VB.CommandButton DeleteCmd 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   7515
            TabIndex        =   69
            Top             =   195
            Width           =   2430
         End
         Begin VB.CommandButton EditCmd 
            Caption         =   "&Edit"
            Height          =   400
            Left            =   5085
            TabIndex        =   68
            Top             =   195
            Width           =   2430
         End
         Begin VB.CommandButton SaveCmd 
            Caption         =   "Save"
            Height          =   400
            Left            =   2655
            TabIndex        =   67
            Top             =   195
            Width           =   2430
         End
         Begin VB.CommandButton AddCmd 
            Caption         =   "&Add New"
            Height          =   400
            Left            =   225
            TabIndex        =   66
            Top             =   195
            Width           =   2430
         End
         Begin VB.CommandButton NextCmd 
            Caption         =   "&Next"
            Height          =   400
            Left            =   225
            TabIndex        =   60
            Top             =   600
            Width           =   2430
         End
         Begin VB.CommandButton PreviousCmd 
            Caption         =   "&Previous"
            Height          =   400
            Left            =   2655
            TabIndex        =   61
            Top             =   600
            Width           =   2430
         End
         Begin VB.CommandButton FirstCmd 
            Caption         =   "&First"
            Height          =   400
            Left            =   5085
            TabIndex        =   62
            Top             =   600
            Width           =   2430
         End
         Begin VB.CommandButton LastCmd 
            Caption         =   "&Last"
            Height          =   400
            Left            =   7515
            TabIndex        =   63
            Top             =   600
            Width           =   2430
         End
         Begin VB.CommandButton ExitCmd 
            Caption         =   "E&xit"
            Height          =   400
            Left            =   9945
            TabIndex        =   64
            Top             =   600
            Width           =   2430
         End
         Begin VB.CommandButton CmdReverse 
            Caption         =   "Reverse"
            Height          =   400
            Left            =   12390
            TabIndex        =   59
            Top             =   195
            Width           =   2070
         End
         Begin VB.CommandButton CmdViewInvoice 
            Caption         =   "View Invoice"
            Height          =   400
            Left            =   12390
            TabIndex        =   65
            Top             =   600
            Width           =   2070
         End
      End
   End
End
Attribute VB_Name = "FrmTxn_DPInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mSLocationID, mDepositoryID, mMonthListindex, mSM_Prefix As Variant

'Private Sub CmbMonth_GotFocus()
'tmp = CmbMonth.Text
'CmbMonth.Clear
'For I = 1 To 12
'    CmbMonth.AddItem MonthName(I), I - 1
'Next
'CmbMonth.Text = tmp
'End Sub
'
'Private Sub CmbMonth_LostFocus()
'
'If CmbMonth.Text = "" Then Exit Sub
'mStatus = False
'For I = 0 To 11
'    If Trim(CmbMonth.Text) = CmbMonth.List(I) Then
'       CmbMonth.Text = CmbMonth.List(I)
'       mStatus = True
'       Exit For
'    End If
'Next
'
'If mStatus = False Then
'   MsgBox "Invalid Selection"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'End Sub

Private Sub CmbNonNcdexClient_GotFocus()
tmp = CmbNonNcdexClient.Text
Call TableMst_Depository("") '" & mExchangeTypeID)
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
   Call TableMst_Location(" Where " & Gen_WcLocation)   '(" Where LocationId > 0 ")
Else
   Call TableMst_Location
End If
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No record found !!!! "
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

'Private Sub CmbYear_GotFocus()
'tmp = CmbYear.Text
'CmbYear.Clear
'For I = 2011 To Year(Now)
'    CmbYear.AddItem I
'Next
'CmbYear.Text = tmp
'End Sub
'
'Private Sub CmbYear_LostFocus()
'If CmbYear.Text = "" Then Exit Sub
'Dim flg As Variant
'flg = "N"
'For I = 2011 To Year(Now)
'    If CmbYear.Text = I Then
'       flg = "Y"
'       Exit For
'    End If
'Next
'If flg = "N" Then
'   MsgBox "Invalid Selection!!!"
'   CmbYear.SetFocus
'   Exit Sub
'End If
'
'End Sub

Private Sub CmdCancelReverse_Click()
TxtReverseDate.Value = Date
TxtReverseRemark = ""
Frame6.Visible = False
ButtonFrm.Enabled = True
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_DPInvoice.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_DPInvoice.xls")
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
oWS.Cells(mRow, C + 1) = "DP Invoice"
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
''========================Month and Year Search==============================
'If CmbMonth.Text = "" Then
'   MsgBox " Please Select Month"
'   CmbMonth.SetFocus
'   Exit Sub
'End If
'If CmbYear.Text = "" Then
'   MsgBox " Please Select Year"
'   CmbYear.SetFocus
'   Exit Sub
'End If
'If mMonthListindex < 10 Then
'   mMonthListindex = "0" & mMonthListindex
'End If
'mMonthListindex = CmbMonth.ListIndex + 1
'vMonth = Month(Now)
'mYear = Val(CmbYear.Text)
'mTotalDays = Day(DateSerial(mYear, mMonthListindex + 1, 0))
'mFromDate = Format(CmbYear.Text & "-" & mMonthListindex & "-" & "01", "yyyy-mm-dd")
'mToDate = Format(CmbYear.Text & "-" & mMonthListindex & "-" & mTotalDays, "yyyy-mm-dd")
''========================Month and Year Search==============================

tmp = " Select Distinct MDPI.DPInvoiceID, MDPI.MonthYear, ML.LocationName, MDPI.DPNavCode,"
tmp = tmp & " MDPI.Amt, MDPI.STAmt, MDPI.TotalAmt, MDPI.PINo, MDPI.BatchID, MDPI.Flag, MDPI.InvoiceNo, MDPI.InvoiceDate,"
tmp = tmp & " MDPI.InvSTAmt,MDPI.InvoiceAmt, MDPI.CreatedDate,"
tmp = tmp & " (Select top 1 MD.DepositoryName from Mst_Depository MD Where MDPI.DPNAVCOde = MD.NavCode) DepositoryName,"
tmp = tmp & " (Select top 1 MD.DepositoryID from Mst_Depository MD Where MDPI.DPNAVCOde = MD.NavCode) DepositoryID,"
tmp = tmp & " TI.TIID , TI.TINo, TI.TIDate, TI.TIAmount, TI.PostInvoiceNo, TI.EntryNo, TI.TISTAmount,MDPI.ReverseFlag "
tmp = tmp & " From Mst_DPInvoiceUpload MDPI"
tmp = tmp & " Inner Join Mst_Location ML On MDPI.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_Depository MD On MD.NavCode = MDPI.DPNAVCode"
tmp = tmp & " Inner Join Txn_TIDetails TI On TI.PIID = MDPI.PINo And PIFlag ='DP'"

If Gen_DBType = "MDB" Then
   tmp = tmp & " Where MDPI.MonthYear Between #" & mFromDate & "# And #" & mToDate & "# " 'And ExchangeTypeID = 1 And (TRI.ContractID = 0 Or TRI.ContractID is Null)"
Else
   tmp = tmp & " Where MDPI.MonthYear Between '" & mFromDate & "' And '" & mToDate & "'" 'And ExchangeTypeID = 1 And (TRI.ContractID = 0 Or TRI.ContractID is Null)"
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
   tmp = tmp & " And MDPI.BatchID = 0 "
End If

If CmbSStatus.Text <> "" Then
   tmp = tmp & " And MDPI.ReverseFlag = '" & Left$(CmbSStatus, 1) & "'"
End If


tmp = tmp & " Order By DPInvoiceID"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       MSHFlexGrid1.TextMatrix(I, 0) = TmpRs1!DPInvoiceID 'InvoiceNo
       MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!CreatedDate), "", Format(TmpRs1!CreatedDate, "dd-MMM-yyyy")) 'InvoiceDate
       MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!MonthYear), "", Format(TmpRs1!MonthYear, "MMM-yyyy")) 'MonthYear
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName) 'Location
       MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!DepositoryName), "", TmpRs1!DepositoryName + "~" + TmpRs1!DepositoryID) 'Depository
       MSHFlexGrid1.TextMatrix(I, 5) = IIf(IsNull(TmpRs1!DPNAVCode), "", TmpRs1!DPNAVCode) 'NAVCode
       MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!Amt), "", TmpRs1!Amt)
       MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!STAmt), "", TmpRs1!STAmt)
       MSHFlexGrid1.TextMatrix(I, 8) = IIf(IsNull(TmpRs1!TotalAmt), "", TmpRs1!TotalAmt)
       MSHFlexGrid1.TextMatrix(I, 9) = IIf(IsNull(TmpRs1!BatchID), "", TmpRs1!BatchID)
       MSHFlexGrid1.TextMatrix(I, 10) = IIf(IsNull(TmpRs1!TINo), "", TmpRs1!TINo)
       MSHFlexGrid1.TextMatrix(I, 11) = IIf(IsNull(TmpRs1!TIDate), "", Format(TmpRs1!TIDate, "dd-MMM-yyyy"))
       MSHFlexGrid1.TextMatrix(I, 12) = IIf(IsNull(TmpRs1!TIAmount), "", TmpRs1!TIAmount)
       MSHFlexGrid1.TextMatrix(I, 13) = IIf(IsNull(TmpRs1!TISTAmount), "", TmpRs1!TISTAmount)
       MSHFlexGrid1.TextMatrix(I, 14) = IIf(IsNull(TmpRs1!EntryNo), "", TmpRs1!EntryNo)
       MSHFlexGrid1.TextMatrix(I, 15) = IIf(IsNull(TmpRs1!ReverseFlag), "N", TmpRs1!ReverseFlag)
       TmpRs1.MoveNext
   Next
End If

Label22.Caption = ""
Label22.Caption = TmpRs1.RecordCount
Label22.Refresh

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

'If TxtReverseDate.Value < RsMst_DPInvoiceUpload!BillingDate Then
'   MsgBox "'Reverse Date' should not less than 'Billing Date' !!"
'   TxtReverseDate.SetFocus
'   Exit Sub
'End If
If TxtReverseRemark.Text = "" Then
   MsgBox "Blank Remark not allowed!!"
   TxtReverseRemark.SetFocus
   Exit Sub
End If
TxtReverseRemark = Replace(TxtReverseRemark, ",", "-")
RsMst_DPInvoiceUpload!ReverseDate = TxtReverseDate.Value
RsMst_DPInvoiceUpload!ReverseFlag = "Y"
RsMst_DPInvoiceUpload!ReverseRemark = TxtReverseRemark.Text
RsMst_DPInvoiceUpload.Update

tmp = " Update Txn_TIDetails Set TIAmount=0,ReverseFlag='Y' Where PIID=" & RsMst_DPInvoiceUpload!PINo & " and PIFLag='DP'"

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
'
'End If
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
         Call Log_Print_History(mSM_Prefix, "DP", TxtOriginalInvoiceNo, TxtEntryNo)
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
RsMst_DPInvoiceUpload.MoveFirst
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
    Call TableMst_DPInvoiceUpload(" Where " & Gen_WcLocation & "  Order by DPInvoiceID")
Else
    Call TableMst_DPInvoiceUpload(" Order by DPInvoiceID")
End If
If RsMst_DPInvoiceUpload.RecordCount = 0 Then
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
    .Cols = 16
    .Font.Size = 10
    .TextMatrix(0, 0) = "Invoice No"
    .TextMatrix(0, 1) = "Invoice Date"
    .TextMatrix(0, 2) = "Year Month"
    .TextMatrix(0, 3) = "Location Name"
    .TextMatrix(0, 4) = "Depository Name"
    .TextMatrix(0, 5) = "NAVCode"
    .TextMatrix(0, 6) = "Billing Amount"
    .TextMatrix(0, 7) = "Service Tax Amount"
    .TextMatrix(0, 8) = "Total Amount"
    .TextMatrix(0, 9) = "Batch ID"
    .TextMatrix(0, 10) = "TI No"
    .TextMatrix(0, 11) = "TI Date"
    .TextMatrix(0, 12) = "TI Amount"
    .TextMatrix(0, 13) = "TI ST Amount"
    .TextMatrix(0, 14) = "TI Entry No"
    .TextMatrix(0, 15) = "Reverse Flag"
    'TxtTINo = ans(1)
    'TxtTIDate = Format(ans(2), "dd-MMM-yyyy")
    'TxtTIAmount = ans(3)
    'TxtEntryNo = ans(5)
    'TxtTISTAmt = ans(7)
    'TxtRecAmt = Val(ans(8)) + Val(ans(9)) + Val(ans(10))
    
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
TxtInvoiceNo.Text = RsMst_DPInvoiceUpload!DPInvoiceID
tmp = " Select Distinct MDPI.DPInvoiceID, MDPI.MonthYear, ML.LocationName, MDPI.DPNavCode,"
tmp = tmp & " MDPI.Amt, MDPI.STAmt, MDPI.TotalAmt, MDPI.PINo, MDPI.BatchID, MDPI.Flag, MDPI.InvoiceNo, MDPI.InvoiceDate,"
tmp = tmp & " MDPI.InvSTAmt,MDPI.InvoiceAmt, MDPI.CreatedDate, MDPI.ReverseRemark,ML.SM_Prefix,"
tmp = tmp & " (Select top 1 MD.DepositoryName from Mst_Depository MD Where MDPI.DPNAVCOde = MD.NavCode) DepositoryName,"
tmp = tmp & " (Select top 1 MD.DepositoryID from Mst_Depository MD Where MDPI.DPNAVCOde = MD.NavCode) DepositoryID"
tmp = tmp & " From Mst_DPInvoiceUpload MDPI"
tmp = tmp & " Inner Join Mst_Location ML On MDPI.LocationID = ML.LocationID"
tmp = tmp & " Where MDPI.DPInvoiceID = " & Val(TxtInvoiceNo)
Call TmpTable

If TmpRs.RecordCount > 0 Then
    TxtInvoiceDate.Text = IIf(IsNull(TmpRs!CreatedDate), "", Format(TmpRs!CreatedDate, "dd-MMM-yyyy"))
    TxtInvoiceMonth.Text = IIf(IsNull(TmpRs!MonthYear), "", Format(TmpRs!MonthYear, "MMM-yyyy"))
    TxtLocation.Text = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
    TxtClient.Text = IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName + "~" + TmpRs!DepositoryID)
    TxtNAVCode.Text = IIf(IsNull(TmpRs!DPNAVCode), "", TmpRs!DPNAVCode)
    TxtBillingAmount.Text = IIf(IsNull(TmpRs!Amt), "", TmpRs!Amt)
    TxtServiceTaxAmount.Text = IIf(IsNull(TmpRs!STAmt), "", TmpRs!STAmt)
    TxtTotalAmount.Text = IIf(IsNull(TmpRs!TotalAmt), "", TmpRs!TotalAmt)
    mSM_Prefix = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
    If Val(TxtServiceTaxAmount) > 0 Then
       TxtServiceTax.Text = "Service Tax" 'GetGLAccNAVCode("DPSTCode") 'IIf(IsNull(TmpRs!InvoiceNo), "", TmpRs!InvoiceNo)
    Else
       TxtServiceTax.Text = "No Service Tax"
    End If
    TxtBatchID.Text = IIf(IsNull(TmpRs!BatchID), "", TmpRs!BatchID)
    TxtRemarks.Text = IIf(IsNull(TmpRs!ReverseRemark), "", TmpRs!ReverseRemark)
    Dim ans As Variant
    ans = GetTIDetails(Val(IIf(IsNull(TmpRs!PINo), "", TmpRs!PINo)), "DP")
    TxtTINo = ans(1)
    TxtTIDate = Format(ans(2), "dd-MMM-yyyy")
    TxtTIAmount = ans(3)
    TxtEntryNo = ans(5)
    TxtTISTAmt = ans(7)
    TxtRecAmt = Val(ans(8)) + Val(ans(9)) + Val(ans(10))
    If Val(TxtRecAmt) > 0 Then
       TxtBalAmt = Format(Val(TxtTIAmount) - Val(TxtRecAmt), "#####0.00")
    Else
       TxtBalAmt = TxtTIAmount
    End If
    If TxtEntryNo <> "" And (LCase(Gen_UserRole) = "fag" Or LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "fagadmin") Then
       If IIf(IsNull(RsMst_DPInvoiceUpload!ReverseFlag), "N", RsMst_DPInvoiceUpload!ReverseFlag) <> "Y" Then
          If CheckTIInCR(Val(ans(0))) = False Then
             CmdReverse.Enabled = True
          End If
       End If
    End If
    TxtCreated = IIf(IsNull(RsMst_DPInvoiceUpload!CreatedBy), "", RsMst_DPInvoiceUpload!CreatedBy) & " :- " & IIf(IsNull(RsMst_DPInvoiceUpload!CreatedDate), "", RsMst_DPInvoiceUpload!CreatedDate)

End If
End Sub


Private Sub LastCmd_Click()
RsMst_DPInvoiceUpload.MoveLast
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
RsMst_DPInvoiceUpload.MoveFirst
RsMst_DPInvoiceUpload.Find "DPInvoiceID = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'"
If Not RsMst_DPInvoiceUpload.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_DPInvoiceUpload.MoveNext
If RsMst_DPInvoiceUpload.EOF Then
   RsMst_DPInvoiceUpload.MoveLast
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
RsMst_DPInvoiceUpload.MovePrevious
If RsMst_DPInvoiceUpload.BOF Then
   RsMst_DPInvoiceUpload.MoveFirst
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
   If RsMst_DPInvoiceUpload.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsMst_DPInvoiceUpload.MoveFirst
 
   RsMst_DPInvoiceUpload.Find "InvoiceNo = '" & TxtInvoiceNo.Text & "'"
   If RsMst_DPInvoiceUpload.EOF Then
      RsMst_DPInvoiceUpload.MoveLast
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

RsMst_DPInvoiceUpload.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head
''CmbMonth.Text = ""
''CmbYear.Text = ""
'TxtFromDate.Value = Date - 30
'TxtToDate.Value = Date
'CmbSLocation.Text = ""
'CmbNonNcdexClient.Text = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

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
