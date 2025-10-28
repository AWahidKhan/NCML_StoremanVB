VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_PledgeRelease 
   Caption         =   "Pledge Release Order"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14475
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   14475
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   7605
      Left            =   120
      ScaleHeight     =   7545
      ScaleWidth      =   14205
      TabIndex        =   12
      Top             =   120
      Width           =   14265
      Begin TabDlg.SSTab SSTab1 
         Height          =   5160
         Left            =   120
         TabIndex        =   28
         Top             =   2280
         Width           =   14040
         _ExtentX        =   24765
         _ExtentY        =   9102
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "GSL Detail"
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame1"
         Tab(0).Control(1)=   "MSHFlexGrid2"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Pledge Detail"
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   -74880
            TabIndex        =   50
            Top             =   360
            Width           =   13725
            Begin VB.TextBox TxtElectronicWeight 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   8280
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   420
               Width           =   1965
            End
            Begin VB.TextBox TxtNoOfBags 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   6315
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   420
               Width           =   1905
            End
            Begin VB.TextBox TxtCMP 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   420
               Width           =   3315
            End
            Begin VB.TextBox TxtGSL 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               Height          =   360
               Left            =   3480
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   420
               Width           =   2775
            End
            Begin VB.CommandButton CmdUpdate 
               Caption         =   "Update List"
               Height          =   360
               Left            =   12135
               TabIndex        =   55
               Top             =   870
               Width           =   1515
            End
            Begin VB.TextBox TxtPledgeNoOfBags 
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
               Left            =   10290
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   420
               Width           =   1785
            End
            Begin VB.TextBox Text1 
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
               Left            =   12120
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   420
               Width           =   1560
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
               Left            =   6300
               TabIndex        =   62
               Top             =   135
               Width           =   1965
            End
            Begin VB.Label Label10 
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
               Left            =   1267
               TabIndex        =   61
               Top             =   165
               Width           =   1020
            End
            Begin VB.Label Label12 
               Caption         =   "G-S-L"
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
               Left            =   4537
               TabIndex        =   60
               Top             =   120
               Width           =   660
            End
            Begin VB.Label Label13 
               Caption         =   "Eletronic Weight"
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
               Left            =   8520
               TabIndex        =   59
               Top             =   120
               Width           =   1500
            End
            Begin VB.Label Label3 
               Caption         =   "Pledge No Of Bags"
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
               Left            =   10320
               TabIndex        =   54
               Top             =   120
               Width           =   1815
            End
            Begin VB.Label Label4 
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
               Height          =   255
               Left            =   12165
               TabIndex        =   53
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.Frame Frame3 
            Height          =   3555
            Left            =   120
            TabIndex        =   31
            Top             =   390
            Width           =   13815
            Begin VB.Frame Frame2 
               Caption         =   "Pledge Receipt Details "
               Height          =   1890
               Left            =   120
               TabIndex        =   44
               Top             =   1530
               Width           =   13485
               Begin VB.TextBox TxtDepositeBankACNo 
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
                  TabIndex        =   69
                  TabStop         =   0   'False
                  Top             =   1170
                  Width           =   1950
               End
               Begin VB.CommandButton CmdDepositeBank 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "Arial Black"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   9000
                  TabIndex        =   67
                  Top             =   720
                  Width           =   375
               End
               Begin VB.ComboBox CmbDepositeBank 
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
                  Sorted          =   -1  'True
                  TabIndex        =   65
                  Top             =   720
                  Width           =   6840
               End
               Begin VB.TextBox TxtChequeAmount 
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
                  Left            =   9120
                  TabIndex        =   64
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   2040
               End
               Begin VB.TextBox TxtChequeNo 
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
                  Left            =   1815
                  TabIndex        =   45
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1950
               End
               Begin MSComCtl2.DTPicker ChequeDate 
                  Height          =   375
                  Left            =   5280
                  TabIndex        =   46
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   661
                  _Version        =   393216
                  Format          =   49086465
                  CurrentDate     =   39884
               End
               Begin VB.TextBox TxtDepositeBank 
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
                  TabIndex        =   68
                  Top             =   720
                  Width           =   6840
               End
               Begin MSComCtl2.DTPicker DepositeDate 
                  Height          =   375
                  Left            =   7050
                  TabIndex        =   71
                  TabStop         =   0   'False
                  Top             =   1155
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   661
                  _Version        =   393216
                  Format          =   49086465
                  CurrentDate     =   39884
               End
               Begin VB.Label Label17 
                  Caption         =   "Date Of Deposite"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   5115
                  TabIndex        =   72
                  Top             =   1200
                  Width           =   1875
               End
               Begin VB.Label Label16 
                  Caption         =   "Deposite Bank A/C No"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   120
                  TabIndex        =   70
                  Top             =   1200
                  Width           =   1695
               End
               Begin VB.Label Label14 
                  Caption         =   "Deposite Bank Name"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   120
                  TabIndex        =   66
                  Top             =   690
                  Width           =   1695
               End
               Begin VB.Label Label40 
                  Caption         =   "Cheque Date"
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
                  Left            =   3960
                  TabIndex        =   49
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label41 
                  Caption         =   "Cheque No"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   120
                  TabIndex        =   48
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.Label Label42 
                  Caption         =   "Cheque Amount"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   7440
                  TabIndex        =   47
                  Top             =   240
                  Width           =   1650
               End
            End
            Begin VB.TextBox TxtPledgeFee 
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
               Left            =   6720
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   1080
               Width           =   2040
            End
            Begin VB.TextBox TxtPledgeReqNo 
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
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   1095
               Width           =   1935
            End
            Begin VB.ComboBox CmbBankName 
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
               Sorted          =   -1  'True
               TabIndex        =   35
               Top             =   240
               Width           =   6825
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
               Left            =   1920
               TabIndex        =   36
               Top             =   240
               Width           =   6810
            End
            Begin VB.ComboBox CmbBranchName 
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
               Sorted          =   -1  'True
               TabIndex        =   34
               Top             =   660
               Width           =   6825
            End
            Begin VB.CommandButton CmdBankName 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   9120
               TabIndex        =   33
               Top             =   255
               Width           =   375
            End
            Begin VB.CommandButton CmdBranchName 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   9120
               TabIndex        =   32
               Top             =   660
               Width           =   375
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
               Height          =   375
               Left            =   1920
               TabIndex        =   37
               Top             =   660
               Width           =   6825
            End
            Begin VB.Label Label11 
               Caption         =   "Pledge Fee Receviable"
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
               Left            =   4245
               TabIndex        =   43
               Top             =   1110
               Width           =   2535
            End
            Begin VB.Label Label5 
               Caption         =   "Pledge Req No"
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
               TabIndex        =   41
               Top             =   1125
               Width           =   1815
            End
            Begin VB.Label Label38 
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
               TabIndex        =   39
               Top             =   285
               Width           =   1455
            End
            Begin VB.Label Label39 
               Caption         =   "Branch Name"
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
               Left            =   120
               TabIndex        =   38
               Top             =   705
               Width           =   1815
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   3225
            Left            =   -74880
            TabIndex        =   29
            Top             =   1800
            Width           =   13785
            _ExtentX        =   24315
            _ExtentY        =   5689
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   2025
         Left            =   120
         ScaleHeight     =   1965
         ScaleWidth      =   13920
         TabIndex        =   13
         Top             =   120
         Width           =   13980
         Begin VB.ComboBox CmbCommodityCondition 
            Height          =   315
            Left            =   1935
            TabIndex        =   27
            Top             =   555
            Width           =   2670
         End
         Begin VB.TextBox TxtClientName 
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
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   960
            Width           =   7710
         End
         Begin VB.TextBox TxtCommodity 
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
            TabIndex        =   17
            Top             =   1395
            Width           =   2670
         End
         Begin VB.TextBox TxtClientID 
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
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   960
            Width           =   2670
         End
         Begin VB.TextBox TxtPledgeID 
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
            TabIndex        =   15
            Top             =   120
            Width           =   2655
         End
         Begin VB.TextBox TxtLocationID 
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
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1395
            Width           =   7725
         End
         Begin MSComCtl2.DTPicker CDCDate 
            Height          =   375
            Left            =   6120
            TabIndex        =   19
            Top             =   120
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   49086465
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtWHRNo 
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
            Top             =   540
            Width           =   2670
         End
         Begin VB.Label Label1 
            Caption         =   "WHR No"
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
            TabIndex        =   26
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label8 
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
            Left            =   4800
            TabIndex        =   25
            Top             =   1020
            Width           =   1365
         End
         Begin VB.Label Label7 
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
            Left            =   135
            TabIndex        =   24
            Top             =   1455
            Width           =   1455
         End
         Begin VB.Label Label26 
            Caption         =   "Pledge  Date"
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
            Left            =   4800
            TabIndex        =   23
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Client ID"
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
            TabIndex        =   22
            Top             =   1020
            Width           =   1590
         End
         Begin VB.Label Label9 
            Caption         =   "Pledge ID"
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
            TabIndex        =   21
            Top             =   180
            Width           =   1335
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   4800
            TabIndex        =   20
            Top             =   1455
            Width           =   1095
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   14205
      TabIndex        =   0
      Top             =   7800
      Width           =   14265
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   2415
         TabIndex        =   10
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4725
         TabIndex        =   9
         Top             =   120
         Width           =   2280
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   7020
         TabIndex        =   8
         Top             =   120
         Width           =   2280
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   9300
         TabIndex        =   7
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   2415
         TabIndex        =   5
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   4710
         TabIndex        =   4
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   7005
         TabIndex        =   3
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   9300
         TabIndex        =   2
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   750
         Left            =   11595
         TabIndex        =   1
         Top             =   120
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frm_PledgeRelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CDCDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub ExitCmd_Click()
  Unload Me
End Sub

Public Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 10
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.TextMatrix(1, 0) = "  "
MSHFlexGrid2.TextMatrix(1, 1) = "CMP Name"
MSHFlexGrid2.TextMatrix(1, 2) = "Godown No"
MSHFlexGrid2.TextMatrix(1, 3) = "Stack No"
MSHFlexGrid2.TextMatrix(1, 4) = "Lot No"
MSHFlexGrid2.TextMatrix(1, 5) = "Availabe Electronic No Of Bags"
MSHFlexGrid2.TextMatrix(1, 6) = "Availabe Electronic Quantity"
MSHFlexGrid2.TextMatrix(1, 7) = "Pledge No Of Bags"
MSHFlexGrid2.TextMatrix(1, 8) = "Pledge Quantity"
MSHFlexGrid2.TextMatrix(1, 9) = "PledgeGSLID"

MSHFlexGrid2.ColWidth(0) = 400
MSHFlexGrid2.ColWidth(1) = 3200
MSHFlexGrid2.ColWidth(2) = 1200
MSHFlexGrid2.ColWidth(3) = 1000
MSHFlexGrid2.ColWidth(4) = 1000
MSHFlexGrid2.ColWidth(5) = 1800
MSHFlexGrid2.ColWidth(6) = 1800
MSHFlexGrid2.ColWidth(7) = 1600
MSHFlexGrid2.ColWidth(8) = 1600
MSHFlexGrid2.ColWidth(9) = 0

MSHFlexGrid2.RowHeight(1) = 400
End Sub

Private Sub Form_Load()
  Call Grid_Head
End Sub


