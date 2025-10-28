VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_GodownWiseProfitability 
   Caption         =   "Godown Wise Profitability"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtUnallocatedSal 
      Height          =   855
      Left            =   120
      TabIndex        =   74
      Top             =   9240
      Width           =   15000
   End
   Begin VB.Frame Frame11 
      Height          =   2325
      Left            =   120
      TabIndex        =   9
      Top             =   -30
      Width           =   15000
      Begin VB.CheckBox ChkInsurance 
         Caption         =   "Insurance Charges"
         Enabled         =   0   'False
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
         Left            =   10800
         TabIndex        =   68
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox ChkInterest 
         Caption         =   "Interest Charges"
         Enabled         =   0   'False
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
         Left            =   12960
         TabIndex        =   67
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CheckBox ChkSpotIncome 
         Caption         =   "Spot Income"
         Enabled         =   0   'False
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
         Left            =   9000
         TabIndex        =   66
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox ChkCMIncome 
         Caption         =   "CM Income"
         Enabled         =   0   'False
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
         Left            =   6720
         TabIndex        =   65
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox ChkPledgeIncome 
         Caption         =   "Pledge Income"
         Enabled         =   0   'False
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
         TabIndex        =   64
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox ChkEmployeeClaim 
         Caption         =   "Employee Claim"
         Enabled         =   0   'False
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
         Left            =   10800
         TabIndex        =   63
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox ChkVendorExpense 
         Caption         =   "Vendor Expense"
         Enabled         =   0   'False
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
         Left            =   12960
         TabIndex        =   62
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox ChkSecurityExpense 
         Caption         =   "Security Expense"
         Enabled         =   0   'False
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
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CheckBox ChkFumigationExpense 
         Caption         =   "Fumigation Expense"
         Enabled         =   0   'False
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
         Left            =   2160
         TabIndex        =   60
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CheckBox ChkGodownRent 
         Caption         =   "Godown Rent"
         Enabled         =   0   'False
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
         TabIndex        =   59
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CheckBox ChkFranchiseeFees 
         Caption         =   "Franchisee Fees"
         Enabled         =   0   'False
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
         Left            =   6720
         TabIndex        =   58
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CheckBox ChkSalary 
         Caption         =   "Salary"
         Enabled         =   0   'False
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
         Left            =   9000
         TabIndex        =   57
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox ChkReservationIncome 
         Caption         =   "Reservation Income"
         Enabled         =   0   'False
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
         TabIndex        =   56
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox ChkSevenDaysIncome 
         Caption         =   "Seven Days Income"
         Enabled         =   0   'False
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
         Left            =   6720
         TabIndex        =   55
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox ChkDPIncome 
         Caption         =   "DP Income"
         Enabled         =   0   'False
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
         Left            =   9000
         TabIndex        =   54
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox ChkRejectedIncome 
         Caption         =   "Rejected Income"
         Enabled         =   0   'False
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
         Left            =   10800
         TabIndex        =   53
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CheckBox ChkOtherIncome 
         Caption         =   "Other Income"
         Enabled         =   0   'False
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
         Left            =   12960
         TabIndex        =   52
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox ChkSCMIncome 
         Caption         =   "SCM Income"
         Enabled         =   0   'False
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
         TabIndex        =   51
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox ChkSpillageIncome 
         Caption         =   "Spillage Income"
         Enabled         =   0   'False
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
         Left            =   2160
         TabIndex        =   50
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CheckBox ChkNNNRIncome 
         Caption         =   "NNNR Income"
         Enabled         =   0   'False
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
         Left            =   2160
         TabIndex        =   49
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox ChkStockDetail 
         Caption         =   "Stock Detail"
         Enabled         =   0   'False
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
         TabIndex        =   48
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox ChkFromSavedData 
         Caption         =   "From Saved Data"
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
         Left            =   7680
         TabIndex        =   47
         Top             =   795
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.OptionButton Opt_Stock 
         Caption         =   "Stock Wise"
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
         TabIndex        =   30
         Top             =   472
         Width           =   1695
      End
      Begin VB.OptionButton Opt_Capacity 
         Caption         =   "Capacity Wise"
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
         TabIndex        =   4
         Top             =   150
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton Cmd_Process 
         Caption         =   "Save Report"
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
         Left            =   13560
         TabIndex        =   8
         Top             =   510
         Width           =   1335
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
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
         Left            =   11145
         TabIndex        =   6
         Top             =   510
         Width           =   1215
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
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
         Left            =   9930
         TabIndex        =   5
         Top             =   510
         Width           =   1215
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
         Height          =   375
         Left            =   12360
         TabIndex        =   7
         Top             =   510
         Width           =   1215
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
         Left            =   4800
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   510
         Width           =   2640
      End
      Begin VB.ComboBox CmbStateID 
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
         Left            =   2250
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   510
         Width           =   2445
      End
      Begin VB.ComboBox CmbMonth 
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
         TabIndex        =   0
         Top             =   510
         Width           =   1170
      End
      Begin VB.ComboBox CmbYear 
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
         ItemData        =   "FrmRpt_GodownWiseProfitability.frx":0000
         Left            =   1335
         List            =   "FrmRpt_GodownWiseProfitability.frx":0002
         TabIndex        =   1
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Month"
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
         Left            =   390
         TabIndex        =   15
         Top             =   195
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Year"
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
         Left            =   1507
         TabIndex        =   14
         Top             =   195
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Record Found :-"
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
         Left            =   10680
         TabIndex        =   13
         Top             =   195
         Width           =   2280
      End
      Begin VB.Label Label5 
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
         Left            =   13080
         TabIndex        =   12
         Top             =   195
         Width           =   90
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
         Left            =   5738
         TabIndex        =   11
         Top             =   195
         Width           =   765
      End
      Begin VB.Label LblStateID 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
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
         Left            =   2940
         TabIndex        =   10
         Top             =   195
         Width           =   1065
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   16
      Tab             =   15
      TabsPerRow      =   16
      TabHeight       =   1058
      TabCaption(0)   =   "Godown Wise Profitability"
      TabPicture(0)   =   "FrmRpt_GodownWiseProfitability.frx":0004
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "MSHFlexGrid1"
      Tab(0).Control(1)=   "ChkAll"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "NCDEX or Non NCDEX Stock"
      TabPicture(1)   =   "FrmRpt_GodownWiseProfitability.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSHFlexGrid2"
      Tab(1).Control(1)=   "CheckStock"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "NNNR Income"
      TabPicture(2)   =   "FrmRpt_GodownWiseProfitability.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSHFlexGrid3"
      Tab(2).Control(1)=   "CheckNNNR"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Reservation Income"
      TabPicture(3)   =   "FrmRpt_GodownWiseProfitability.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "MSHFlexGrid4"
      Tab(3).Control(1)=   "CheckReservation"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Invoice Income"
      TabPicture(4)   =   "FrmRpt_GodownWiseProfitability.frx":0074
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "MSHFlexGrid5"
      Tab(4).Control(1)=   "CheckInvoice"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "DP Income as per Invoice"
      TabPicture(5)   =   "FrmRpt_GodownWiseProfitability.frx":0090
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "MSHFlexGrid6"
      Tab(5).Control(1)=   "CheckDPIncome"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Pledge Income"
      TabPicture(6)   =   "FrmRpt_GodownWiseProfitability.frx":00AC
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "MSHFlexGrid7"
      Tab(6).Control(1)=   "CheckPledge"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Vendor Expense"
      TabPicture(7)   =   "FrmRpt_GodownWiseProfitability.frx":00C8
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "MSHFlexGrid8"
      Tab(7).Control(1)=   "CheckVendor"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Security Expense"
      TabPicture(8)   =   "FrmRpt_GodownWiseProfitability.frx":00E4
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "MSHFlexGrid9"
      Tab(8).Control(1)=   "CheckSecurity"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "Godown Rent Expense"
      TabPicture(9)   =   "FrmRpt_GodownWiseProfitability.frx":0100
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "MSHFlexGrid10"
      Tab(9).Control(1)=   "CheckGodownRent"
      Tab(9).ControlCount=   2
      TabCaption(10)  =   "DP Income as per GSL"
      TabPicture(10)  =   "FrmRpt_GodownWiseProfitability.frx":011C
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "MSHFlexGrid11"
      Tab(10).Control(1)=   "CheckDPGSL"
      Tab(10).ControlCount=   2
      TabCaption(11)  =   "Franchisee Fees"
      TabPicture(11)  =   "FrmRpt_GodownWiseProfitability.frx":0138
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "MSHFlexGrid12"
      Tab(11).Control(1)=   "CheckFranchisee"
      Tab(11).ControlCount=   2
      TabCaption(12)  =   "Interest Charges"
      TabPicture(12)  =   "FrmRpt_GodownWiseProfitability.frx":0154
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "MSHFlexGrid13"
      Tab(12).Control(1)=   "CheckInterest"
      Tab(12).ControlCount=   2
      TabCaption(13)  =   "Reversal Detail"
      TabPicture(13)  =   "FrmRpt_GodownWiseProfitability.frx":0170
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "MSHFlexGrid14"
      Tab(13).Control(1)=   "MSHFlexGrid15"
      Tab(13).Control(2)=   "CheckReversal"
      Tab(13).ControlCount=   3
      TabCaption(14)  =   "Claim Expense from HR Pro"
      TabPicture(14)  =   "FrmRpt_GodownWiseProfitability.frx":018C
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "MSHFlexGrid16"
      Tab(14).Control(1)=   "CheckClaimExpenseHRPro"
      Tab(14).ControlCount=   2
      TabCaption(15)  =   "Salary of Employee "
      TabPicture(15)  =   "FrmRpt_GodownWiseProfitability.frx":01A8
      Tab(15).ControlEnabled=   -1  'True
      Tab(15).Control(0)=   "MSHFlexGrid17"
      Tab(15).Control(0).Enabled=   0   'False
      Tab(15).Control(1)=   "CheckSalaryofEmployee"
      Tab(15).Control(1).Enabled=   0   'False
      Tab(15).ControlCount=   2
      Begin VB.CheckBox CheckSalaryofEmployee 
         Caption         =   "Get Detail"
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
         TabIndex        =   72
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox CheckClaimExpenseHRPro 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   70
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox ChkAll 
         Caption         =   "With All Detail"
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
         Left            =   -74880
         TabIndex        =   69
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox CheckReversal 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   46
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox CheckInterest 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   45
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox CheckFranchisee 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   44
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox CheckDPGSL 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   43
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox CheckGodownRent 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   42
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox CheckSecurity 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   41
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox CheckVendor 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   40
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox CheckPledge 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   39
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox CheckDPIncome 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   38
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox CheckInvoice 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   37
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox CheckReservation 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   36
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox CheckNNNR 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   35
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox CheckStock 
         Caption         =   "Get Detail"
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
         Left            =   -74880
         TabIndex        =   34
         Top             =   840
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   17
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   18
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   20
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   21
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   22
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   23
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid7 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   24
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid8 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   25
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid9 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   26
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid10 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   27
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid11 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   28
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid12 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   29
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid13 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   31
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid15 
         Height          =   4950
         Left            =   -74880
         TabIndex        =   33
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8731
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid14 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   32
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   6376
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid16 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   71
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8493
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid17 
         Height          =   4935
         Left            =   120
         TabIndex        =   73
         Top             =   1200
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   8705
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
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   8760
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "FrmRpt_GodownWiseProfitability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mStateID, mLocationIDs As Variant
Dim mStartDate, mEndDate, mMonth, mYear, mTotalDays, Wc As Variant
Dim mUnAllocatedSalary As Double
Enum E
   MonthYear = 0
   StateName = 1
   LocationName = 2
   AMname = 3
   RMName = 4
   CMPName = 5
   GodownNo = 6
   UnitType = 7
   GodownStartDate = 8
   GodownCloseDate = 9
   GodownCapacity = 10
   NcdexStock = 11
   NonNCDEXStock = 12
   NNNRIncomeAdhoc = 13
   NNNRIncome = 14
   ReservationIncomeAdhoc = 15
   ReservationIncome = 16
   SevenDaysIncome = 17
   DPIncomeAdHoc = 18
   FinalDPIncome = 19
   RejectedIncome = 20
   SpotIncomeAdHoc = 21
   OtherIncome = 22
   SCMIncome = 23
   CMUnbilledIncome = 24
   CMPendingIncome = 25
   CMApprovedIncome = 26
   LeftOverStock = 27
   PledgeIncome = 28
   
   LeaseRentPaid = 29
   LeaseRentFAG = 30
   LeaseRentDRC = 31
   EmpClaimFAG = 32
   EmpClaimAccrual = 33
   VendorExpenses = 34
   VendorExpensesAccrued = 35
   SecurityFeesPaid = 36
   SecurityFeesAccrual = 37
   FumigationFeesVendor = 38
   FumigationFeesPaid = 39
   FumigationFeesAccrual = 40
   FranchiseeFeesPaid = 41
   FranchiseeFeesFAG = 42
   FranchiseeFeesCMG = 43
   Salary = 44
   SalaryFieldStaff = 45
   NCDEXCharges = 46
   CoExpense = 47
   InsurancePremium = 48
   InterestCharges = 49
   TotalIncome = 50
   TotalExpense = 51
   GrossIncome = 52
   NetIncome = 53
   StateID = 54
   LocationID = 55
   GodownID = 56
   DPIncomeasperGSL = 57
   Difference = 58
   DPIncomeasperInvoice = 59

End Enum

Private Sub ChkAll_Click()
If ChkAll.Value = vbChecked Then
   CheckStock.Value = vbChecked
   CheckNNNR.Value = vbChecked
   CheckReservation.Value = vbChecked
   CheckInvoice.Value = vbChecked
   CheckDPIncome.Value = vbChecked
   CheckPledge.Value = vbChecked
   CheckVendor.Value = vbChecked
   CheckSecurity.Value = vbChecked
   CheckGodownRent.Value = vbChecked
   CheckDPGSL.Value = vbChecked
   CheckFranchisee.Value = vbChecked
   CheckInterest.Value = vbChecked
   CheckReversal.Value = vbChecked
   CheckClaimExpenseHRPro.Value = vbChecked
   CheckSalaryofEmployee.Value = vbChecked
Else
   CheckStock.Value = vbUnchecked
   CheckNNNR.Value = vbUnchecked
   CheckReservation.Value = vbUnchecked
   CheckInvoice.Value = vbUnchecked
   CheckDPIncome.Value = vbUnchecked
   CheckPledge.Value = vbUnchecked
   CheckVendor.Value = vbUnchecked
   CheckSecurity.Value = vbUnchecked
   CheckGodownRent.Value = vbUnchecked
   CheckDPGSL.Value = vbUnchecked
   CheckFranchisee.Value = vbUnchecked
   CheckInterest.Value = vbUnchecked
   CheckReversal.Value = vbUnchecked
   CheckClaimExpenseHRPro.Value = vbUnchecked
   CheckSalaryofEmployee.Value = vbUnchecked
End If
End Sub

Private Sub ChkFromSavedData_Click()
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "head cmg" Then Exit Sub
         
If ChkFromSavedData.Value = vbChecked Then
   ChkStockDetail.Enabled = True
   ChkNNNRIncome.Enabled = True
   ChkReservationIncome.Enabled = True
   ChkSevenDaysIncome.Enabled = True
   ChkDPIncome.Enabled = True
   ChkRejectedIncome.Enabled = True
   ChkOtherIncome.Enabled = True
   ChkSCMIncome.Enabled = True
   ChkSpillageIncome.Enabled = True
   ChkPledgeIncome.Enabled = True
   ChkCMIncome.Enabled = True
   ChkSpotIncome.Enabled = True
   ChkEmployeeClaim.Enabled = True
   ChkVendorExpense.Enabled = True
   ChkSecurityExpense.Enabled = True
   ChkFumigationExpense.Enabled = True
   ChkGodownRent.Enabled = True
   ChkFranchiseeFees.Enabled = True
   ChkSalary.Enabled = True
   ChkInsurance.Enabled = True
   ChkInterest.Enabled = True
Else
   ChkStockDetail.Enabled = False
   ChkNNNRIncome.Enabled = False
   ChkReservationIncome.Enabled = False
   ChkSevenDaysIncome.Enabled = False
   ChkDPIncome.Enabled = False
   ChkRejectedIncome.Enabled = False
   ChkOtherIncome.Enabled = False
   ChkSCMIncome.Enabled = False
   ChkSpillageIncome.Enabled = False
   ChkPledgeIncome.Enabled = False
   ChkCMIncome.Enabled = False
   ChkSpotIncome.Enabled = False
   ChkEmployeeClaim.Enabled = False
   ChkVendorExpense.Enabled = False
   ChkSecurityExpense.Enabled = False
   ChkFumigationExpense.Enabled = False
   ChkGodownRent.Enabled = False
   ChkFranchiseeFees.Enabled = False
   ChkSalary.Enabled = False
   ChkInsurance.Enabled = False
   ChkInterest.Enabled = False
   
   ChkStockDetail.Value = vbUnchecked
   ChkNNNRIncome.Value = vbUnchecked
   ChkReservationIncome.Value = vbUnchecked
   ChkSevenDaysIncome.Value = vbUnchecked
   ChkDPIncome.Value = vbUnchecked
   ChkRejectedIncome.Value = vbUnchecked
   ChkOtherIncome.Value = vbUnchecked
   ChkSCMIncome.Value = vbUnchecked
   ChkSpillageIncome.Value = vbUnchecked
   ChkPledgeIncome.Value = vbUnchecked
   ChkCMIncome.Value = vbUnchecked
   ChkSpotIncome.Value = vbUnchecked
   ChkEmployeeClaim.Value = vbUnchecked
   ChkVendorExpense.Value = vbUnchecked
   ChkSecurityExpense.Value = vbUnchecked
   ChkFumigationExpense.Value = vbUnchecked
   ChkGodownRent.Value = vbUnchecked
   ChkFranchiseeFees.Value = vbUnchecked
   ChkSalary.Value = vbUnchecked
   ChkInsurance.Value = vbUnchecked
   ChkInterest.Value = vbUnchecked
   
End If
End Sub

Private Sub CmbMonth_GotFocus()
tmp = CmbMonth.Text
CmbMonth.Clear
For i = 1 To 12
    CmbMonth.AddItem MonthName(i)
Next
CmbMonth.Text = tmp
End Sub

Private Sub CmbMonth_LostFocus()
If Trim(CmbMonth.Text) = "" Then Exit Sub
mStatus = False
For i = 0 To 11
   If Trim(CmbMonth.Text) = CmbMonth.List(i) Then
      CmbMonth.Text = CmbMonth.List(i)
      mStatus = True
      mMonth = i + 1
      Exit For
   End If
Next
If mStatus = False Then
   MsgBox "Invalid Selection"
   CmbMonth.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
Call TableMst_State
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub
'----
Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   CmbLocationID.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If

If mStateID <> RsMst_State!StateID Then
   CmbLocationID.Text = ""
End If
mStateID = RsMst_State!StateID
End Sub
Private Sub CmbLocationID_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Please Select State Name !!"
   CmbStateID.SetFocus
   Exit Sub
End If

tmp = CmbLocationID.Text
Call TableMst_Location("Where StateID=" & mStateID & "")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbStateID.SetFocus
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
   CmbLocationID.AddItem RsMst_Location!LocationName
   RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()

If CmbLocationID.Text = "" Then Exit Sub

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbYear_GotFocus()
If CmbMonth.Text = "" Then
   MsgBox "Select Month!!!"
   CmbMonth.SetFocus
   Exit Sub
End If
tmp = CmbYear.Text
CmbYear.Clear
For i = 2010 To Year(Now)
   CmbYear.AddItem i
Next

If Month(Now) = 12 Then
   CmbYear.AddItem Year(Now) + 1
End If

CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
If Trim(CmbYear.Text) = "" Then Exit Sub
Dim flg As Variant
Dim myr As Variant

flg = "N"

If Month(Now) = 12 Then
   myr = Year(Now) + 1
Else
   myr = Year(Now)
End If

For i = 2010 To myr 'Year(Now)
   If CmbYear.Text = i Then
      flg = "Y"
      Exit For
   End If
Next
If flg = "N" Then
   MsgBox "Invalid Selection!!!"
   CmbYear.SetFocus
   Exit Sub
End If

mYear = CmbYear.Text
End Sub

Private Sub Cmd_Process_Click()

If Opt_Capacity.Value = True Then
    mReportFlag = "C"
ElseIf Opt_Stock.Value = True Then
    mReportFlag = "S"
End If

tmp = "Delete From Txn_MonGodownWiseProfitability Where MonthYear = '" & mStartDate & "' And ReportFlag = '" & mReportFlag & "'"
Call TmpTable

tmp = "Select * From Txn_MonGodownWiseProfitability Where MonthYear = '" & mStartDate & "' And ReportFlag = '" & mReportFlag & "'"
Call Tmp1Table

ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
For i = 2 To MSHFlexGrid1.Rows - 2
   TmpRs1.AddNew
   TmpRs1!MGWPID = GetMGWPID
   TmpRs1!G_UID = GetGUID
   TmpRs1!MonthYear = CDate(mStartDate)
   If MSHFlexGrid1.TextMatrix(i, E.GodownID) <> "" Then
      TmpRs1!StateID = MSHFlexGrid1.TextMatrix(i, E.StateID)
      TmpRs1!LocationID = MSHFlexGrid1.TextMatrix(i, E.LocationID)
      TmpRs1!GodownID = MSHFlexGrid1.TextMatrix(i, E.GodownID)
   End If
   TmpRs1!AM = MSHFlexGrid1.TextMatrix(i, E.AMname)
   TmpRs1!RM = MSHFlexGrid1.TextMatrix(i, E.RMName)
   TmpRs1!Capacity = Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity))
   TmpRs1!NcdexStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock))
   TmpRs1!NNStock = Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
   TmpRs1!NNNRIncomeAdhoc = Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncomeAdhoc))
   TmpRs1!NNNRIncome = Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncome))
   TmpRs1!ReservationIncomeAdhoc = Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc))
   TmpRs1!ReservationIncome = Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome))
   TmpRs1!SevenDaysIncome = Val(MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome))
   TmpRs1!DPIncomeAdHoc = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeAdHoc))
   TmpRs1!FinalDPIncome = Val(MSHFlexGrid1.TextMatrix(i, E.FinalDPIncome))
   TmpRs1!DPIncomeGSL = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperGSL))
   TmpRs1!Diff = Val(MSHFlexGrid1.TextMatrix(i, E.Difference))
   TmpRs1!DPIncomeInvoice = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice))
   TmpRs1!RejectedIncome = Val(MSHFlexGrid1.TextMatrix(i, E.RejectedIncome))
   TmpRs1!SpotIncomeAdHoc = Val(MSHFlexGrid1.TextMatrix(i, E.SpotIncomeAdHoc))
   TmpRs1!OtherIncome = Val(MSHFlexGrid1.TextMatrix(i, E.OtherIncome))
   TmpRs1!SCMIncome = Val(MSHFlexGrid1.TextMatrix(i, E.SCMIncome))
   TmpRs1!CMUnbilledIncome = Val(MSHFlexGrid1.TextMatrix(i, E.CMUnbilledIncome))
   TmpRs1!CMPendingIncome = Val(MSHFlexGrid1.TextMatrix(i, E.CMPendingIncome))
   TmpRs1!CMApprovedIncome = Val(MSHFlexGrid1.TextMatrix(i, E.CMApprovedIncome))
   TmpRs1!LeftOverStock = Val(MSHFlexGrid1.TextMatrix(i, E.LeftOverStock))
   TmpRs1!PledgeIncome = Val(MSHFlexGrid1.TextMatrix(i, E.PledgeIncome))
   TmpRs1!LeaseRentPaid = Val(MSHFlexGrid1.TextMatrix(i, E.LeaseRentPaid))
   TmpRs1!LeaseRentFAG = Val(MSHFlexGrid1.TextMatrix(i, E.LeaseRentFAG))
   TmpRs1!LeaseRentDRC = Val(MSHFlexGrid1.TextMatrix(i, E.LeaseRentDRC))
   TmpRs1!EmpClaimFAG = Val(MSHFlexGrid1.TextMatrix(i, E.EmpClaimFAG))
   TmpRs1!EmpClaimAccrual = Val(MSHFlexGrid1.TextMatrix(i, E.EmpClaimAccrual))
   TmpRs1!VendorExpenses = Val(MSHFlexGrid1.TextMatrix(i, E.VendorExpenses))
   TmpRs1!VendorExpensesAccrued = Val(MSHFlexGrid1.TextMatrix(i, E.VendorExpensesAccrued))
   TmpRs1!SecurityFeesPaid = Val(MSHFlexGrid1.TextMatrix(i, E.SecurityFeesPaid))
   TmpRs1!SecurityFeesAccrual = Val(MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual))
   TmpRs1!FumigationFeesVendor = Val(MSHFlexGrid1.TextMatrix(i, E.FumigationFeesVendor))
   TmpRs1!FumigationFeesPaid = Val(MSHFlexGrid1.TextMatrix(i, E.FumigationFeesPaid))
   TmpRs1!FumigationFeesAccrual = Val(MSHFlexGrid1.TextMatrix(i, E.FumigationFeesAccrual))
   TmpRs1!FranchiseeFeesPaid = Val(MSHFlexGrid1.TextMatrix(i, E.FranchiseeFeesPaid))
   TmpRs1!FranchiseeFeesFAG = Val(MSHFlexGrid1.TextMatrix(i, E.FranchiseeFeesFAG))
   TmpRs1!FranchiseeFeesCMG = Val(MSHFlexGrid1.TextMatrix(i, E.FranchiseeFeesCMG))
   TmpRs1!Salary = Val(MSHFlexGrid1.TextMatrix(i, E.Salary))
   TmpRs1!SalaryFieldStaff = Val(MSHFlexGrid1.TextMatrix(i, E.SalaryFieldStaff))
   TmpRs1!NCDEXCharges = Val(MSHFlexGrid1.TextMatrix(i, E.NCDEXCharges))
   TmpRs1!CoExpense = Val(MSHFlexGrid1.TextMatrix(i, E.CoExpense))
   TmpRs1!InsurancePremium = Val(MSHFlexGrid1.TextMatrix(i, E.InsurancePremium))
   TmpRs1!InterestCharges = Val(MSHFlexGrid1.TextMatrix(i, E.InterestCharges))
   TmpRs1!ReportFlag = mReportFlag
   TmpRs1!CreatedDate = Date
   TmpRs1!CreatedBy = Gen_UserLoginID
   TmpRs1.Update
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done!!!"
Cmd_Process.Enabled = False
End Sub

Private Function GetMGWPID() As Double
Dim Retval As Double
tmp = "Select max(MGWPID) from Txn_MonGodownWiseProfitability "
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMGWPID = Retval
End Function

Private Sub CmdClear_Click()
Call Grid_Head

'TxtFromDate.Value = CDate(Year(Date) & "-" & Month(Date) & "-01")
'mTotalDays = Day(DateSerial(TxtFromDate.Year, TxtFromDate.Month + 1, 0))
'TxtToDate.Value = CDate(Year(Date) & "-" & Month(Date) & "-" & mTotalDays)

CmbMonth.Text = ""
CmbYear.Text = ""
CmbStateID.Text = ""
CmbLocationID.Text = ""
End Sub

Private Sub CmdExcel_Click()
'Gen_mTimeStamp = GetTimeStamp
'Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownWiseProfitability.xls")
'Call Xls_Detail_Rpt

Dim Path, strFilName, strFilLocation As Variant
Gen_mTimeStamp = GetTimeStamp
tmp = ""
Ans = MsgBox("Do you want to Genearate CSV File", vbYesNo)
If Ans = vbNo Then
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownWiseProfitability.xls")
   Call Xls_Detail_Rpt
Else
   Path = App.Path
   strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "GodownWiseProfitability" & Gen_UserName & "GWP" & " .csv"
   strFilLocation = Path + strFilName
   If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
   Open strFilLocation For Output As #1
   For i = 0 To MSHFlexGrid1.Rows - 2
       tmp = ""
       For c = 0 To MSHFlexGrid1.Cols - 7
          If c = 0 Then
             tmp = tmp & Replace(Replace(MSHFlexGrid1.TextMatrix(i, c), vbNewLine, ""), ",", "~")
          ElseIf c = 6 Then
             tmp = tmp & "," & "'" & Replace(Replace(MSHFlexGrid1.TextMatrix(i, c), vbNewLine, ""), ",", "~")
          Else
             tmp = tmp & "," & Replace(Replace(MSHFlexGrid1.TextMatrix(i, c), vbNewLine, ""), ",", "~")
          End If
       Next
       Print #1, tmp
   Next
   Close #1
   
   
   If CheckStock.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "NCDEX OR Non NCDEX" & Gen_UserName & "Stock" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid2.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid2.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid2.TextMatrix(i, c), vbNewLine, ""), ",", "~")
             ElseIf c = 3 Then
                tmp = tmp & "," & "'" & Replace(Replace(MSHFlexGrid2.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid2.TextMatrix(i, c), vbNewLine, ""), ",", "~")
                 
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckNNNR.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "NNNRIncome" & Gen_UserName & "NNNR" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid3.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid3.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid3.TextMatrix(i, c), vbNewLine, ""), ",", "~") 'MSHFlexGrid3.TextMatrix(i, c)
             ElseIf c = 3 Then
                tmp = tmp & "," & "'" & Replace(Replace(MSHFlexGrid3.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid3.TextMatrix(i, c), vbNewLine, ""), ",", "~") 'Replace(Replace(MSHFlexGrid3.TextMatrix(i, c), ",", "~"), vbNewLine, "")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckReservation.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "ReservationIncome" & Gen_UserName & "RI" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid4.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid4.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid4.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid4.TextMatrix(i, c), vbNewLine, ""), ",", "~") 'Replace(MSHFlexGrid4.TextMatrix(i, c), ",", "~")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckInvoice.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "InvoiceIncome" & Gen_UserName & "INI" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid5.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid5.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid5.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid5.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckDPIncome.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "DPIncome" & Gen_UserName & "DPI" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid6.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid6.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid6.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid6.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckPledge.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "PledgeIncome" & Gen_UserName & "PI" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid7.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid7.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid7.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid7.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckVendor.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "VendorExpense" & Gen_UserName & "VE" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid8.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid8.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid8.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid8.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckSecurity.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "SecurityExpense" & Gen_UserName & "SE" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid9.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid9.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid9.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid9.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckGodownRent.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "RentExpense" & Gen_UserName & "GRE" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid10.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid10.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid10.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              ElseIf c = 3 Then
                 tmp = tmp & "," & "'" & Replace(Replace(MSHFlexGrid10.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid10.TextMatrix(i, c), vbNewLine, ""), ",", "~") 'Replace(MSHFlexGrid10.TextMatrix(i, c), ",", "~")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckDPGSL.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "DPGSLInvoice" & Gen_UserName & "DPGSLI" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid11.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid11.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid11.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              ElseIf c = 3 Then
                 tmp = tmp & "," & "'" & Replace(Replace(MSHFlexGrid11.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid11.TextMatrix(i, c), vbNewLine, ""), ",", "~") 'Replace(MSHFlexGrid11.TextMatrix(i, c), ",", "~")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckFranchisee.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "FranchiseeFees" & Gen_UserName & "FF" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid12.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid12.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid12.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              ElseIf c = 3 Then
                 tmp = tmp & "," & "'" & Replace(Replace(MSHFlexGrid11.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid12.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckInterest.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "InterestCharges" & Gen_UserName & "Interest" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid13.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid13.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid13.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid13.TextMatrix(i, c), vbNewLine, ""), ",", "~") 'Replace(MSHFlexGrid13.TextMatrix(i, c), ",", "~")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckInterest.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "ReversalDetail" & Gen_UserName & "RD" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid15.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid15.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid15.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid15.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckClaimExpenseHRPro.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "ClaimExpenseFromHRPro" & Gen_UserName & "CEHPro" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid16.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid16.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid16.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid16.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   If CheckSalaryofEmployee.Value = vbChecked Then
      strFilName = ""
      strFilLocation = ""
      strFilName = "\Excel\" & CmbMonth & CmbYear & Gen_mTimeStamp & "SalaryofEmployee" & Gen_UserName & "SOE" & " .csv"
      strFilLocation = Path + strFilName
      If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
      Open strFilLocation For Output As #1
      For i = 0 To MSHFlexGrid17.Rows - 2
          tmp = ""
          For c = 0 To MSHFlexGrid17.Cols - 1
              If c = 0 Then
                 tmp = tmp & Replace(Replace(MSHFlexGrid17.TextMatrix(i, c), vbNewLine, ""), ",", "~")
              Else
                 tmp = tmp & "," & Replace(Replace(MSHFlexGrid17.TextMatrix(i, c), vbNewLine, ""), ",", "~") 'Replace(MSHFlexGrid17.TextMatrix(i, c), ",", "~")
              End If
          Next
          Print #1, tmp
      Next
      Close #1
   End If
   MsgBox "Done !!!"
End If


End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Variant
Dim mCol As Variant
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GodownWiseProfitability.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Godown Wise Profitability"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid1.Rows + 2
mRow = 1

'oWS.Range("o3:o3").Interior.Color = vbGreen
'oWS.Range("p3:p3").Interior.Color = vbGreen
'oWS.Range("q3:q3").Interior.Color = vbGreen
'oWS.Range("r3:r3").Interior.Color = vbGreen
'oWS.Range("s3:s3").Interior.Color = vbGreen
'oWS.Range("t3:t3").Interior.Color = vbGreen
'oWS.Range("u3:u3").Interior.Color = vbGreen
'oWS.Range("v3:v3").Interior.Color = vbGreen
'oWS.Range("w3:w3").Interior.Color = vbGreen
'oWS.Range("x3:x3").Interior.Color = vbGreen
'oWS.Range("y3:y3").Interior.Color = vbGreen
'oWS.Range("z3:z3").Interior.Color = vbGreen
'oWS.Range("aa3:aa3").Interior.Color = vbGreen
'oWS.Range("ab3:ab3").Interior.Color = vbGreen
'oWS.Range("ac3:ac3").Interior.Color = vbGreen

'oWS.Range("ad3:at3").Interior.Color = vbRed
'oWS.Range("ae3:ae3").Interior.Color = vbRed
'oWS.Range("af3:af3").Interior.Color = vbRed
'oWS.Range("ag3:ag3").Interior.Color = vbRed
'oWS.Range("ah3:ah3").Interior.Color = vbRed
'oWS.Range("ai3:ai3").Interior.Color = vbRed
'oWS.Range("aj3:aj3").Interior.Color = vbRed
'oWS.Range("ak3:ak3").Interior.Color = vbRed
'oWS.Range("al3:al3").Interior.Color = vbRed
'oWS.Range("am3:am3").Interior.Color = vbRed
'oWS.Range("an3:an3").Interior.Color = vbRed
'oWS.Range("ao3:ao3").Interior.Color = vbRed
'oWS.Range("ap3:ap3").Interior.Color = vbRed
'oWS.Range("aq3:aq3").Interior.Color = vbRed
'oWS.Range("ar3:ar3").Interior.Color = vbRed
'oWS.Range("as3:as3").Interior.Color = vbRed
'oWS.Range("at3:at3").Interior.Color = vbRed

'oWS.Range("au3:aw3").Interior.Color = &HC0FF&
'oWS.Range("av3:av3").Interior.Color = &HC0FF&
'oWS.Range("aw3:aw3").Interior.Color = &HC0FF&

'oWS.Range("ax3:ba3").Interior.Color = &HF3EEDB
'oWS.Range("ay3:ay3").Interior.Color = &HF3EEDB
'oWS.Range("az3:az3").Interior.Color = &HF3EEDB
'oWS.Range("ba3:ba3").Interior.Color = &HF3EEDB

tmp = "a" & mRow & ":" & "bb" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Godown Wise Profitability"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

For i = 0 To MSHFlexGrid1.Rows - 2
   mRow = mRow + 1 ' + 6
   For c = 0 To MSHFlexGrid1.Cols - 7
      oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
      If i = 0 Then
         oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
         oWS.Cells(mRow, c + 1).Font.Bold = True
      ElseIf i = 1 Then
         If c >= 13 And c <= 28 Then
            oWS.Cells(mRow, c + 1).Interior.Color = vbGreen
         ElseIf c >= 29 And c <= 45 Then
            oWS.Cells(mRow, c + 1).Interior.Color = vbRed
         ElseIf c >= 46 And c <= 49 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HC0FF&
         ElseIf c >= 50 And c <= 53 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HF3EEDB
         Else
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
         End If
         oWS.Cells(mRow, c + 1).Font.Bold = True
         oWS.Cells(mRow, c + 1).RowHeight = 45
      End If
      oWS.Cells(mRow, c + 1).WrapText = True
      oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
   Next
   ProgressBar1.Value = ProgressBar1.Value + 1
Next
ProgressBar1.Value = ProgressBar1.Max

Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "NCDEX OR Non NCDEX Stock"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid2.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 19
'oWS.Range("d:d").ColumnWidth = 14
'oWS.Range("e:e").ColumnWidth = 40
'oWS.Range("f:f").ColumnWidth = 8
'oWS.Range("g:g").ColumnWidth = 8
'oWS.Range("h:h").ColumnWidth = 25
'oWS.Range("i:i").ColumnWidth = 14
'oWS.Range("j:j").ColumnWidth = 11

tmp = "a" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "NCDEX OR Non NCDEX Stock"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckStock.Value = vbChecked Then
   For i = 0 To MSHFlexGrid2.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid2.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Set oWS = oWB.Worksheets("Sheet3")
oWS.Name = "NNNR Income"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid3.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 19
'oWS.Range("d:d").ColumnWidth = 14
'oWS.Range("e:e").ColumnWidth = 40
'oWS.Range("f:f").ColumnWidth = 8
'oWS.Range("g:g").ColumnWidth = 8
'oWS.Range("h:h").ColumnWidth = 19
'oWS.Range("i:i").ColumnWidth = 25
'oWS.Range("j:j").ColumnWidth = 11.43

tmp = "a" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "NNNR Income"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckNNNR.Value = vbChecked Then
   For i = 0 To MSHFlexGrid3.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid3.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Set oWS = oWB.Worksheets("Sheet4")
oWS.Name = "Reservation Income"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid4.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 19
'oWS.Range("d:d").ColumnWidth = 14
'oWS.Range("e:e").ColumnWidth = 40
'oWS.Range("f:f").ColumnWidth = 8
'oWS.Range("g:g").ColumnWidth = 8
'oWS.Range("h:h").ColumnWidth = 19
'oWS.Range("i:i").ColumnWidth = 25
'oWS.Range("j:j").ColumnWidth = 10
'oWS.Range("k:k").ColumnWidth = 11
'oWS.Range("l:l").ColumnWidth = 11.43

tmp = "a" & mRow & ":" & "e" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Reservation Income"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckReservation.Value = vbChecked Then
   For i = 0 To MSHFlexGrid4.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid4.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid4.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Set oWS = oWB.Worksheets("Sheet5")
oWS.Name = "Invoice Income"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid5.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 14
'oWS.Range("d:d").ColumnWidth = 19
'oWS.Range("e:e").ColumnWidth = 25
'oWS.Range("f:f").ColumnWidth = 19
'oWS.Range("g:g").ColumnWidth = 8
'oWS.Range("h:h").ColumnWidth = 9
'oWS.Range("i:i").ColumnWidth = 11
'oWS.Range("j:j").ColumnWidth = 11.43

tmp = "a" & mRow & ":" & "j" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Invoice Income"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckInvoice.Value = vbChecked Then
   For i = 0 To MSHFlexGrid5.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid5.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid5.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Set oWS = oWB.Worksheets("Sheet6")
oWS.Name = "DP Income"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid6.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 11.43
'oWS.Range("d:d").ColumnWidth = 8
'oWS.Range("e:e").ColumnWidth = 10.71
'oWS.Range("f:f").ColumnWidth = 11.14
'oWS.Range("g:g").ColumnWidth = 11
'oWS.Range("h:h").ColumnWidth = 11.43
'oWS.Range("i:i").ColumnWidth = 11.43

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "DP Income"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckDPIncome.Value = vbChecked Then
   For i = 0 To MSHFlexGrid6.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid6.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid6.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Set oWS = oWB.Worksheets("Sheet7")
oWS.Name = "Pledge Income"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid7.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 10
'oWS.Range("d:d").ColumnWidth = 20
'oWS.Range("e:e").ColumnWidth = 20
'oWS.Range("f:f").ColumnWidth = 15
'oWS.Range("g:g").ColumnWidth = 8
'oWS.Range("h:h").ColumnWidth = 9
'oWS.Range("i:i").ColumnWidth = 11.43
'oWS.Range("j:j").ColumnWidth = 11.43
'oWS.Range("k:k").ColumnWidth = 11.43
'oWS.Range("l:l").ColumnWidth = 11.43

tmp = "a" & mRow & ":" & "l" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Pledge Income"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckPledge.Value = vbChecked Then
   For i = 0 To MSHFlexGrid7.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid7.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid7.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Set oWS = oWB.Worksheets("Sheet8")
oWS.Name = "Vendor Expense"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid8.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 15
'oWS.Range("d:d").ColumnWidth = 70
'oWS.Range("e:e").ColumnWidth = 12

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Vendor Expense"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckVendor.Value = vbChecked Then
   For i = 0 To MSHFlexGrid8.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid8.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid8.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Set oWS = oWB.Worksheets("Sheet9")
oWS.Name = "Security Expense"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid9.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 19
'oWS.Range("d:d").ColumnWidth = 20
'oWS.Range("e:e").ColumnWidth = 20
'oWS.Range("f:f").ColumnWidth = 15
'oWS.Range("g:g").ColumnWidth = 8
'oWS.Range("h:h").ColumnWidth = 9
'oWS.Range("i:i").ColumnWidth = 11.43
'oWS.Range("j:j").ColumnWidth = 11.43
'oWS.Range("k:k").ColumnWidth = 11.43

tmp = "a" & mRow & ":" & "l" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Security Expense"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckSecurity.Value = vbChecked Then
   For i = 0 To MSHFlexGrid9.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid9.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid9.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Set oWS = oWB.Worksheets("Sheet10")
oWS.Name = "Rent Expense"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid10.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 19
'oWS.Range("d:d").ColumnWidth = 14
'oWS.Range("e:e").ColumnWidth = 40
'oWS.Range("f:f").ColumnWidth = 12
'oWS.Range("g:g").ColumnWidth = 20
'oWS.Range("h:h").ColumnWidth = 25
'oWS.Range("i:i").ColumnWidth = 11.43

tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Rent Expense"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckGodownRent.Value = vbChecked Then
   For i = 0 To MSHFlexGrid10.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid10.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid10.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Set oWS = oWB.Worksheets("Sheet11")
oWS.Name = "DP GSL Invoice"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid11.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 19
'oWS.Range("d:d").ColumnWidth = 14
'oWS.Range("e:e").ColumnWidth = 40
'oWS.Range("f:f").ColumnWidth = 25
'oWS.Range("g:g").ColumnWidth = 20
'oWS.Range("h:h").ColumnWidth = 10
'oWS.Range("i:i").ColumnWidth = 10
'oWS.Range("j:j").ColumnWidth = 10
'oWS.Range("k:k").ColumnWidth = 10
'oWS.Range("l:l").ColumnWidth = 11.43

tmp = "a" & mRow & ":" & "l" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "DP GSL Invoice"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckDPGSL.Value = vbChecked Then
   For i = 0 To MSHFlexGrid11.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid11.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid11.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Set oWS = oWB.Worksheets("Sheet12")
oWS.Name = "Franchisee Fees"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid12.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 19
'oWS.Range("d:d").ColumnWidth = 14
'oWS.Range("e:e").ColumnWidth = 40
'oWS.Range("f:f").ColumnWidth = 25
'oWS.Range("g:g").ColumnWidth = 20
'oWS.Range("h:h").ColumnWidth = 10
'oWS.Range("i:i").ColumnWidth = 10
'oWS.Range("j:j").ColumnWidth = 10
'oWS.Range("k:k").ColumnWidth = 10
'oWS.Range("l:l").ColumnWidth = 11.43

tmp = "a" & mRow & ":" & "q" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Franchisee Fees"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckFranchisee.Value = vbChecked Then
   For i = 0 To MSHFlexGrid12.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid12.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid12.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Set oWS = oWB.Worksheets("Sheet13")
oWS.Name = "Interest Charges"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid13.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 19
'oWS.Range("d:d").ColumnWidth = 14
'oWS.Range("e:e").ColumnWidth = 40
'oWS.Range("f:f").ColumnWidth = 25
'oWS.Range("g:g").ColumnWidth = 20
'oWS.Range("h:h").ColumnWidth = 10
'oWS.Range("i:i").ColumnWidth = 10
'oWS.Range("j:j").ColumnWidth = 10
'oWS.Range("k:k").ColumnWidth = 10
'oWS.Range("l:l").ColumnWidth = 11.43

tmp = "a" & mRow & ":" & "l" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Interest Charges"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckInterest.Value = vbChecked Then
   For i = 0 To MSHFlexGrid13.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid13.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid13.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Set oWS = oWB.Worksheets("Sheet14")
oWS.Name = "Reversal Detail"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid15.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 19
'oWS.Range("d:d").ColumnWidth = 14
'oWS.Range("e:e").ColumnWidth = 40
'oWS.Range("f:f").ColumnWidth = 25
'oWS.Range("g:g").ColumnWidth = 20
'oWS.Range("h:h").ColumnWidth = 10

tmp = "a" & mRow & ":" & "h" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Reversal Detail"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckInterest.Value = vbChecked Then
   For i = 0 To MSHFlexGrid15.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid15.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid15.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

Set oWS = oWB.Worksheets("Sheet15")
oWS.Name = "Claim Expense from HR Pro"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid16.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 19
'oWS.Range("d:d").ColumnWidth = 14
'oWS.Range("e:e").ColumnWidth = 40
'oWS.Range("f:f").ColumnWidth = 25'
'oWS.Range("g:g").ColumnWidth = 20
'oWS.Range("h:h").ColumnWidth = 10

tmp = "a" & mRow & ":" & "m" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Claim Expense from HR Pro"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckClaimExpenseHRPro.Value = vbChecked Then
   For i = 0 To MSHFlexGrid16.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid16.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid16.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

oWB.Worksheets.Add , oWB.Worksheets("Claim Expense from HR Pro"), 1

Set oWS = oWB.Worksheets(16)
oWS.Name = "Salary of Employee"
oWS.ClearArrows
ProgressBar1.Value = 0
ProgressBar1.Max = MSHFlexGrid17.Rows + 2
mRow = 1
c = 0

'oWS.Range("a:a").ColumnWidth = 15
'oWS.Range("b:b").ColumnWidth = 15
'oWS.Range("c:c").ColumnWidth = 19
'oWS.Range("d:d").ColumnWidth = 14
'oWS.Range("e:e").ColumnWidth = 40
'oWS.Range("f:f").ColumnWidth = 25'
'oWS.Range("g:g").ColumnWidth = 20
'oWS.Range("h:h").ColumnWidth = 10

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Salary of Employee"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

If CheckSalaryofEmployee.Value = vbChecked Then
   For i = 0 To MSHFlexGrid17.Rows - 2
      mRow = mRow + 1 ' + 6
      For c = 0 To MSHFlexGrid17.Cols - 1
         oWS.Cells(mRow, c + 1) = MSHFlexGrid17.TextMatrix(i, c)
         If i = 0 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &HD9D9FF
            oWS.Cells(mRow, c + 1).Font.Bold = True
         ElseIf i = 1 Then
            oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
            oWS.Cells(mRow, c + 1).Font.Bold = True
            oWS.Cells(mRow, c + 1).RowHeight = 45
         End If
         oWS.Cells(mRow, c + 1).WrapText = True
         oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
      Next
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
End If

oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
'Cmd_Process.Enabled = False
Call Grid_Head

'If (MsgBox("Please check that Salary is Processed in HR-Pro, CMP-CM Share Income is processed in CM-PRo and Franchisee Fees Stock Op. Balance in StoreMan?", vbYesNo) = vbNo) Then Exit Sub

If Trim(CmbMonth.Text) = "" Then
   MsgBox "Select Month!!!"
   CmbMonth.SetFocus
   Exit Sub
End If

If Trim(CmbYear.Text) = "" Then
   MsgBox "Select Year!!!"
   CmbYear.SetFocus
   Exit Sub
End If

mTotalDays = Day(DateSerial(mYear, mMonth + 1, 0))

Wc = ""
mYearMonth = CmbMonth.Text & "-" & CmbYear.Text
If CmbLocationID.Text <> "" Then
   Wc = " Where ML.LocationID = " & mLocationID
Else
   If Gen_WcLocation <> "" Then
      Wc = " ML." & Gen_WcLocation
   Else
      If CmbStateID.Text <> "" Then
         Wc = " Where MS.StateID = " & mStateID
      Else
         If Gen_WcState <> "" Then
            Wc = " Where ML." & Gen_WcLocation
         End If
      End If
   End If
End If

mStartDate = mYear & "-" & mMonth & "-1"
mEndDate = mYear & "-" & mMonth & "-" & mTotalDays

If Opt_Stock.Value = True Then
    tmp = " Select MCMP.LocationID, sum(TIMG.BalanceWeight) BalanceWeight"
    tmp = tmp & " From Txn_InvMonGSLDetail TIMG"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
    tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
    tmp = tmp & " Where TIMG.ExchangeTypeID in (1, 2) And TIMG.BalanceWeight > 0 And TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
    tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
    tmp = tmp & " Group by MCMP.LocationID"
    Call Tmp2Table

ElseIf Opt_Capacity.Value = True Then
    
    tmp = " Select ML.LocationID, sum(MG.GodownCapacity) Capacity"
    tmp = tmp & " From Mst_Godown MG"
    tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
    tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
    tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
    tmp = tmp & " group by ML.LocationID"
    Call Tmp2Table
End If

If ChkFromSavedData.Value = vbChecked Then
   ''==========================Taking Data From Inserted Table Txn_MonGodownWiseProfitability===============
   tmp = " Select TMGWP.*, ML.LocationName, MS.StateName, MG.GodownCapacity, MG.GodownNo, MG.Address,  MG.StartDate, MG.CloseDate, MUT.UnitType, MCMP.CMPNAme From Txn_MonGodownWiseProfitability TMGWP"
   tmp = tmp & " Left Join Mst_Location ML On ML.LocationID = TMGWP.LocationID"
   tmp = tmp & " Left Join Mst_State MS On MS.StateID = TMGWP.StateID"
   tmp = tmp & " Left Join Mst_Godown MG On MG.GodownID = TMGWP.GodownID"
   tmp = tmp & " Left Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UNitTypeID"
   tmp = tmp & " Left Join Mst_CMP MCMP On MCMP.CMPID = MG.CMPID"
   If Wc <> "" Then
       tmp = tmp & " " & Wc & " And TMGWP.MonthYear = '" & CDate(mStartDate) & "'"
   Else
       tmp = tmp & " Where TMGWP.MonthYear = '" & CDate(mStartDate) & "'"
   End If
   If Opt_Capacity.Value = True Then
       tmp = tmp & " And TMGWP.ReportFlag = 'C'"
   ElseIf Opt_Stock.Value = True Then
       tmp = tmp & " And TMGWP.ReportFlag = 'S'"
   End If
   tmp = tmp & " Order by MGWPID"
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
'      tmp = " Select MS.StateName, ML.LocationID, ML.LocationName, MCMP.CMPName, MG.GodownID, MG.GodownNo, MG.Address, MG.StartDate, MG.CloseDate, MUT.UnitType,  MG.GodownCapacity, MS.StateID "
'      tmp = tmp & " From Mst_Godown MG"
'      tmp = tmp & " Inner Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UNitTypeID"
'      tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
'      tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
'      tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
'      If wc <> "" Then
'         tmp = tmp & wc & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
'      Else
'         tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
'      End If
'      tmp = tmp & " Order by MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo"
'      Call Tmp1Table
'      If TmpRs.RecordCount <> TmpRs1.RecordCount Then GoTo GodownData
    
      Label5.Caption = TmpRs.RecordCount
      ProgressBar1.Value = 0
      ProgressBar1.Max = MSHFlexGrid1.Rows + TmpRs.RecordCount
      MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
      CmdExcel.Enabled = True
      If (LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "head cmg") And CmbLocationID.Text = "" And CmbStateID.Text = "" Then
         Cmd_Process.Enabled = True
      End If
      
      For i = 2 To TmpRs.RecordCount + 1
         With MSHFlexGrid1
            .TextMatrix(i, E.MonthYear) = mYearMonth
            .TextMatrix(i, E.StateName) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
            .TextMatrix(i, E.LocationName) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
            .TextMatrix(i, E.AMname) = IIf(IsNull(TmpRs!AM), "", TmpRs!AM)
            .TextMatrix(i, E.RMName) = IIf(IsNull(TmpRs!RM), "", TmpRs!RM)
            .TextMatrix(i, E.CMPName) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
            .TextMatrix(i, E.GodownNo) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
            .TextMatrix(i, E.GodownCapacity) = IIf(IsNull(TmpRs!Capacity), "", TmpRs!Capacity)
            .TextMatrix(i, E.NcdexStock) = IIf(IsNull(TmpRs!NcdexStock), "", TmpRs!NcdexStock)
            .TextMatrix(i, E.NonNCDEXStock) = IIf(IsNull(TmpRs!NNStock), "", TmpRs!NNStock)
            .TextMatrix(i, E.NNNRIncomeAdhoc) = IIf(IsNull(TmpRs!NNNRIncomeAdhoc), "", TmpRs!NNNRIncomeAdhoc)
            .TextMatrix(i, E.NNNRIncome) = IIf(IsNull(TmpRs!NNNRIncome), "", TmpRs!NNNRIncome)
            .TextMatrix(i, E.ReservationIncomeAdhoc) = IIf(IsNull(TmpRs!ReservationIncomeAdhoc), "", TmpRs!ReservationIncomeAdhoc)
            .TextMatrix(i, E.ReservationIncome) = IIf(IsNull(TmpRs!ReservationIncome), "", TmpRs!ReservationIncome)
            .TextMatrix(i, E.SevenDaysIncome) = IIf(IsNull(TmpRs!SevenDaysIncome), "", TmpRs!SevenDaysIncome)
            .TextMatrix(i, E.DPIncomeAdHoc) = IIf(IsNull(TmpRs!DPIncomeAdHoc), "", TmpRs!DPIncomeAdHoc)
            .TextMatrix(i, E.FinalDPIncome) = IIf(IsNull(TmpRs!FinalDPIncome), "", TmpRs!FinalDPIncome)
            .TextMatrix(i, E.DPIncomeasperGSL) = IIf(IsNull(TmpRs!DPIncomeGSL), "", TmpRs!DPIncomeGSL)
            .TextMatrix(i, E.Difference) = IIf(IsNull(TmpRs!Diff), "", TmpRs!Diff)
            .TextMatrix(i, E.DPIncomeasperInvoice) = IIf(IsNull(TmpRs!DPIncomeInvoice), "", TmpRs!DPIncomeInvoice)
            .TextMatrix(i, E.RejectedIncome) = IIf(IsNull(TmpRs!RejectedIncome), "", TmpRs!RejectedIncome)
            .TextMatrix(i, E.SpotIncomeAdHoc) = IIf(IsNull(TmpRs!SpotIncomeAdHoc), "", TmpRs!SpotIncomeAdHoc)
            .TextMatrix(i, E.OtherIncome) = IIf(IsNull(TmpRs!OtherIncome), "", TmpRs!OtherIncome)
            .TextMatrix(i, E.SCMIncome) = IIf(IsNull(TmpRs!SCMIncome), "", TmpRs!SCMIncome)
            .TextMatrix(i, E.CMUnbilledIncome) = IIf(IsNull(TmpRs!CMUnbilledIncome), "", TmpRs!CMUnbilledIncome)
            .TextMatrix(i, E.CMPendingIncome) = IIf(IsNull(TmpRs!CMPendingIncome), "", TmpRs!CMPendingIncome)
            .TextMatrix(i, E.CMApprovedIncome) = IIf(IsNull(TmpRs!CMApprovedIncome), "", TmpRs!CMApprovedIncome)
            .TextMatrix(i, E.LeftOverStock) = IIf(IsNull(TmpRs!LeftOverStock), "", TmpRs!LeftOverStock)
            .TextMatrix(i, E.PledgeIncome) = IIf(IsNull(TmpRs!PledgeIncome), "", TmpRs!PledgeIncome)
            .TextMatrix(i, E.LeaseRentPaid) = IIf(IsNull(TmpRs!LeaseRentPaid), "", TmpRs!LeaseRentPaid)
            .TextMatrix(i, E.LeaseRentFAG) = IIf(IsNull(TmpRs!LeaseRentFAG), "", TmpRs!LeaseRentFAG)
            .TextMatrix(i, E.LeaseRentDRC) = IIf(IsNull(TmpRs!LeaseRentDRC), "", TmpRs!LeaseRentDRC)
            .TextMatrix(i, E.EmpClaimFAG) = IIf(IsNull(TmpRs!EmpClaimFAG), "", TmpRs!EmpClaimFAG)
            .TextMatrix(i, E.EmpClaimAccrual) = IIf(IsNull(TmpRs!EmpClaimAccrual), "", TmpRs!EmpClaimAccrual)
            .TextMatrix(i, E.VendorExpenses) = IIf(IsNull(TmpRs!VendorExpenses), "", TmpRs!VendorExpenses)
            .TextMatrix(i, E.VendorExpensesAccrued) = IIf(IsNull(TmpRs!VendorExpensesAccrued), "", TmpRs!VendorExpensesAccrued)
            .TextMatrix(i, E.SecurityFeesPaid) = IIf(IsNull(TmpRs!SecurityFeesPaid), "", TmpRs!SecurityFeesPaid)
            .TextMatrix(i, E.SecurityFeesAccrual) = IIf(IsNull(TmpRs!SecurityFeesAccrual), "", TmpRs!SecurityFeesAccrual)
            .TextMatrix(i, E.FumigationFeesVendor) = IIf(IsNull(TmpRs!FumigationFeesVendor), "", TmpRs!FumigationFeesVendor)
            .TextMatrix(i, E.FumigationFeesPaid) = IIf(IsNull(TmpRs!FumigationFeesPaid), "", TmpRs!FumigationFeesPaid)
            .TextMatrix(i, E.FumigationFeesAccrual) = IIf(IsNull(TmpRs!FumigationFeesAccrual), "", TmpRs!FumigationFeesAccrual)
            .TextMatrix(i, E.FranchiseeFeesPaid) = IIf(IsNull(TmpRs!FranchiseeFeesPaid), "", TmpRs!FranchiseeFeesPaid)
            .TextMatrix(i, E.FranchiseeFeesCMG) = IIf(IsNull(TmpRs!FranchiseeFeesCMG), "", TmpRs!FranchiseeFeesCMG)
            .TextMatrix(i, E.FranchiseeFeesFAG) = IIf(IsNull(TmpRs!FranchiseeFeesFAG), "", TmpRs!FranchiseeFeesFAG)
            .TextMatrix(i, E.Salary) = IIf(IsNull(TmpRs!Salary), "", TmpRs!Salary)
            .TextMatrix(i, E.SalaryFieldStaff) = IIf(IsNull(TmpRs!SalaryFieldStaff), "", TmpRs!SalaryFieldStaff)
            .TextMatrix(i, E.NCDEXCharges) = IIf(IsNull(TmpRs!NCDEXCharges), "", TmpRs!NCDEXCharges)
            .TextMatrix(i, E.CoExpense) = IIf(IsNull(TmpRs!CoExpense), "", TmpRs!CoExpense)
            .TextMatrix(i, E.InsurancePremium) = IIf(IsNull(TmpRs!InsurancePremium), "", TmpRs!InsurancePremium)
            .TextMatrix(i, E.InterestCharges) = IIf(IsNull(TmpRs!InterestCharges), "", TmpRs!InterestCharges)
            .TextMatrix(i, E.GodownStartDate) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy"))
            .TextMatrix(i, E.GodownCloseDate) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-yyyy"))
            .TextMatrix(i, E.UnitType) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
            .TextMatrix(i, E.StateID) = IIf(IsNull(TmpRs!StateID), "", TmpRs!StateID)
            .TextMatrix(i, E.LocationID) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
            .TextMatrix(i, E.GodownID) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
            
            .TextMatrix(0, E.NcdexStock) = Val(.TextMatrix(i, E.NcdexStock)) + Val(.TextMatrix(0, E.NcdexStock))
            .TextMatrix(0, E.GodownCapacity) = Val(.TextMatrix(i, E.GodownCapacity)) + Val(.TextMatrix(0, E.GodownCapacity))
            .TextMatrix(0, E.NonNCDEXStock) = Val(.TextMatrix(i, E.NonNCDEXStock)) + Val(.TextMatrix(0, E.NonNCDEXStock))
            .TextMatrix(0, E.NNNRIncomeAdhoc) = Val(.TextMatrix(i, E.NNNRIncomeAdhoc)) + Val(.TextMatrix(0, E.NNNRIncomeAdhoc))
            .TextMatrix(0, E.NNNRIncome) = Val(.TextMatrix(i, E.NNNRIncome)) + Val(.TextMatrix(0, E.NNNRIncome))
            .TextMatrix(0, E.ReservationIncomeAdhoc) = Val(.TextMatrix(i, E.ReservationIncomeAdhoc)) + Val(.TextMatrix(0, E.ReservationIncomeAdhoc))
            .TextMatrix(0, E.ReservationIncome) = Val(.TextMatrix(i, E.ReservationIncome)) + Val(.TextMatrix(0, E.ReservationIncome))
            .TextMatrix(0, E.SevenDaysIncome) = Val(.TextMatrix(i, E.SevenDaysIncome)) + Val(.TextMatrix(0, E.SevenDaysIncome))
            .TextMatrix(0, E.DPIncomeAdHoc) = Val(.TextMatrix(i, E.DPIncomeAdHoc)) + Val(.TextMatrix(0, E.DPIncomeAdHoc))
            .TextMatrix(0, E.FinalDPIncome) = Val(.TextMatrix(i, E.FinalDPIncome)) + Val(.TextMatrix(0, E.FinalDPIncome))
            .TextMatrix(0, E.DPIncomeasperGSL) = Val(.TextMatrix(i, E.DPIncomeasperGSL)) + Val(.TextMatrix(0, E.DPIncomeasperGSL))
            .TextMatrix(0, E.Difference) = Val(.TextMatrix(i, E.Difference)) + Val(.TextMatrix(0, E.Difference))
            .TextMatrix(0, E.DPIncomeasperInvoice) = Val(.TextMatrix(i, E.DPIncomeasperInvoice)) + Val(.TextMatrix(0, E.DPIncomeasperInvoice))
            .TextMatrix(0, E.RejectedIncome) = Val(.TextMatrix(i, E.RejectedIncome)) + Val(.TextMatrix(0, E.RejectedIncome))
            .TextMatrix(0, E.SpotIncomeAdHoc) = Val(.TextMatrix(i, E.SpotIncomeAdHoc)) + Val(.TextMatrix(0, E.SpotIncomeAdHoc))
            .TextMatrix(0, E.OtherIncome) = Val(.TextMatrix(i, E.OtherIncome)) + Val(.TextMatrix(0, E.OtherIncome))
            .TextMatrix(0, E.SCMIncome) = Val(.TextMatrix(i, E.SCMIncome)) + Val(.TextMatrix(0, E.SCMIncome))
            .TextMatrix(0, E.CMUnbilledIncome) = Val(.TextMatrix(i, E.CMUnbilledIncome)) + Val(.TextMatrix(0, E.CMUnbilledIncome))
            .TextMatrix(0, E.CMPendingIncome) = Val(.TextMatrix(i, E.CMPendingIncome)) + Val(.TextMatrix(0, E.CMPendingIncome))
            .TextMatrix(0, E.CMApprovedIncome) = Val(.TextMatrix(i, E.CMApprovedIncome)) + Val(.TextMatrix(0, E.CMApprovedIncome))
            .TextMatrix(0, E.LeftOverStock) = Val(.TextMatrix(i, E.LeftOverStock)) + Val(.TextMatrix(0, E.LeftOverStock))
            .TextMatrix(0, E.PledgeIncome) = Val(.TextMatrix(i, E.PledgeIncome)) + Val(.TextMatrix(0, E.PledgeIncome))
            .TextMatrix(0, E.LeaseRentPaid) = Val(.TextMatrix(i, E.LeaseRentPaid)) + Val(.TextMatrix(0, E.LeaseRentPaid))
            .TextMatrix(0, E.LeaseRentFAG) = Val(.TextMatrix(i, E.LeaseRentFAG)) + Val(.TextMatrix(0, E.LeaseRentFAG))
            .TextMatrix(0, E.LeaseRentDRC) = Val(.TextMatrix(i, E.LeaseRentDRC)) + Val(.TextMatrix(0, E.LeaseRentDRC))
            .TextMatrix(0, E.EmpClaimFAG) = Val(.TextMatrix(i, E.EmpClaimFAG)) + Val(.TextMatrix(0, E.EmpClaimFAG))
            .TextMatrix(0, E.EmpClaimAccrual) = Val(.TextMatrix(i, E.EmpClaimAccrual)) + Val(.TextMatrix(0, E.EmpClaimAccrual))
            .TextMatrix(0, E.VendorExpenses) = Val(.TextMatrix(i, E.VendorExpenses)) + Val(.TextMatrix(0, E.VendorExpenses))
            .TextMatrix(0, E.VendorExpensesAccrued) = Val(.TextMatrix(i, E.VendorExpensesAccrued)) + Val(.TextMatrix(0, E.VendorExpensesAccrued))
            .TextMatrix(0, E.SecurityFeesPaid) = Val(.TextMatrix(i, E.SecurityFeesPaid)) + Val(.TextMatrix(0, E.SecurityFeesPaid))
            .TextMatrix(0, E.SecurityFeesAccrual) = Val(.TextMatrix(i, E.SecurityFeesAccrual)) + Val(.TextMatrix(0, E.SecurityFeesAccrual))
            .TextMatrix(0, E.FumigationFeesVendor) = Val(.TextMatrix(i, E.FumigationFeesVendor)) + Val(.TextMatrix(0, E.FumigationFeesVendor))
            .TextMatrix(0, E.FumigationFeesPaid) = Val(.TextMatrix(i, E.FumigationFeesPaid)) + Val(.TextMatrix(0, E.FumigationFeesPaid))
            .TextMatrix(0, E.FumigationFeesAccrual) = Val(.TextMatrix(i, E.FumigationFeesAccrual)) + Val(.TextMatrix(0, E.FumigationFeesAccrual))
            .TextMatrix(0, E.FranchiseeFeesPaid) = Val(.TextMatrix(i, E.FranchiseeFeesPaid)) + Val(.TextMatrix(0, E.FranchiseeFeesPaid))
            .TextMatrix(0, E.FranchiseeFeesCMG) = Val(.TextMatrix(i, E.FranchiseeFeesCMG)) + Val(.TextMatrix(0, E.FranchiseeFeesCMG))
            .TextMatrix(0, E.FranchiseeFeesFAG) = Val(.TextMatrix(i, E.FranchiseeFeesFAG)) + Val(.TextMatrix(0, E.FranchiseeFeesFAG))
            .TextMatrix(0, E.Salary) = Val(.TextMatrix(i, E.Salary)) + Val(.TextMatrix(0, E.Salary))
            .TextMatrix(0, E.SalaryFieldStaff) = Val(.TextMatrix(i, E.SalaryFieldStaff)) + Val(.TextMatrix(0, E.SalaryFieldStaff))
            .TextMatrix(0, E.NCDEXCharges) = Val(.TextMatrix(i, E.NCDEXCharges)) + Val(.TextMatrix(0, E.NCDEXCharges))
            .TextMatrix(0, E.CoExpense) = Val(.TextMatrix(i, E.CoExpense)) + Val(.TextMatrix(0, E.CoExpense))
            .TextMatrix(0, E.InsurancePremium) = Val(.TextMatrix(i, E.InsurancePremium)) + Val(.TextMatrix(0, E.InsurancePremium))
            .TextMatrix(0, E.InterestCharges) = Val(.TextMatrix(i, E.InterestCharges)) + Val(.TextMatrix(0, E.InterestCharges))
         End With
         TmpRs.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next i
      
      If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "head cmg" Then
         If ChkStockDetail.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.NcdexStock) = 0
            MSHFlexGrid1.TextMatrix(0, E.NonNCDEXStock) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.NcdexStock) = 0
               MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock) = 0
            Next
            Call GetNCDEXStock
            Call GetNonNCDEXStock
         End If
         If ChkNNNRIncome.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.NNNRIncome) = 0
            MSHFlexGrid1.TextMatrix(0, E.NNNRIncomeAdhoc) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.NNNRIncome) = 0
               MSHFlexGrid1.TextMatrix(i, E.NNNRIncomeAdhoc) = 0
            Next
            Call GetNNNRIncome
            Call GetNNNRIncomethroughAdhoc
         End If
         If ChkReservationIncome.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.ReservationIncome) = 0
            MSHFlexGrid1.TextMatrix(0, E.ReservationIncomeAdhoc) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.ReservationIncome) = 0
               MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc) = 0
            Next
            Call GetReservationIncome
            Call GetReservationIncomethroughAdhoc
         End If
         If ChkSevenDaysIncome.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.SevenDaysIncome) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome) = 0
            Next
            Call GetSevenDaysIncome
         End If
         If ChkDPIncome.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.DPIncomeAdHoc) = 0
            MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperGSL) = 0
            MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice) = 0
            MSHFlexGrid1.TextMatrix(0, E.Difference) = 0
            MSHFlexGrid1.TextMatrix(0, E.FinalDPIncome) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.DPIncomeAdHoc) = 0
               MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperGSL) = 0
               MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice) = 0
               MSHFlexGrid1.TextMatrix(i, E.Difference) = 0
               MSHFlexGrid1.TextMatrix(i, E.FinalDPIncome) = 0
            Next
            Call GetDPIncome
            Call GetDPIncomethroughAdhoc
         End If
         If ChkRejectedIncome.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.RejectedIncome) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.RejectedIncome) = 0
            Next
            Call GetNCDEXRejectedIncome
         End If
         If ChkOtherIncome.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.OtherIncome) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.OtherIncome) = 0
            Next
            Call GetOtherIncome
         End If
         If ChkSCMIncome.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.SCMIncome) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.SCMIncome) = 0
            Next
            Call GetSCMIncome
         End If
         If ChkSpillageIncome.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.LeftOverStock) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.LeftOverStock) = 0
            Next
            Call GetSpillageIncome
         End If
         If ChkPledgeIncome.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.PledgeIncome) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.PledgeIncome) = 0
            Next
            Call GetPledgeIncome
         End If
         If ChkEmployeeClaim.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.EmpClaimAccrual) = 0
            MSHFlexGrid1.TextMatrix(0, E.EmpClaimFAG) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.EmpClaimAccrual) = 0
               MSHFlexGrid1.TextMatrix(i, E.EmpClaimFAG) = 0
            Next
            Call GetExpenseEmployeeClaim
         End If
         If ChkVendorExpense.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.VendorExpenses) = 0
            MSHFlexGrid1.TextMatrix(0, E.VendorExpensesAccrued) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.VendorExpenses) = 0
               MSHFlexGrid1.TextMatrix(i, E.VendorExpensesAccrued) = 0
            Next
            Call GetVendorExpense
         End If
         If ChkSecurityExpense.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.SecurityFeesAccrual) = 0
            MSHFlexGrid1.TextMatrix(0, E.SecurityFeesPaid) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.SecurityFeesPaid) = 0
               MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual) = 0
            Next
            Call GetSecurityExpense
         End If
         If ChkFumigationExpense.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.FumigationFeesAccrual) = 0
            MSHFlexGrid1.TextMatrix(0, E.FumigationFeesPaid) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.FumigationFeesPaid) = 0
               MSHFlexGrid1.TextMatrix(i, E.FumigationFeesAccrual) = 0
            Next
            Call GetFumigationExpense
         End If
         If ChkGodownRent.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.LeaseRentDRC) = 0
            MSHFlexGrid1.TextMatrix(0, E.LeaseRentFAG) = 0
            MSHFlexGrid1.TextMatrix(0, E.LeaseRentPaid) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.LeaseRentPaid) = 0
               MSHFlexGrid1.TextMatrix(i, E.LeaseRentFAG) = 0
               MSHFlexGrid1.TextMatrix(i, E.LeaseRentDRC) = 0
            Next
            Call GetRentExpense
         End If
         If ChkFranchiseeFees.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.FranchiseeFeesPaid) = 0
            MSHFlexGrid1.TextMatrix(0, E.FranchiseeFeesCMG) = 0
            MSHFlexGrid1.TextMatrix(0, E.FranchiseeFeesFAG) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.FranchiseeFeesPaid) = 0
               MSHFlexGrid1.TextMatrix(i, E.FranchiseeFeesCMG) = 0
               MSHFlexGrid1.TextMatrix(i, E.FranchiseeFeesFAG) = 0
            Next
            Call GetFranchiseeExpense
         End If
         If ChkSalary.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.Salary) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.Salary) = 0
            Next
            MSHFlexGrid1.TextMatrix(0, E.SalaryFieldStaff) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.SalaryFieldStaff) = 0
            Next
'            Call GetSalaryExpense
             Call GetSalaryExpenseFromIncomeforFieldStaff
             Call GetSalaryExpenseFromIncomeforCOStaff
         End If
         If ChkInsurance.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.InsurancePremium) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.InsurancePremium) = 0
            Next
            Call GetInsuranceCharges
         End If
         If ChkInterest.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.InterestCharges) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.InterestCharges) = 0
            Next
            Call GetOutstandingInterest
         End If
         If ChkSpotIncome.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.SpotIncomeAdHoc) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.SpotIncomeAdHoc) = 0
            Next
            Call GetSpotIncomethroughAdhoc
         End If
         If ChkCMIncome.Value = vbChecked Then
            MSHFlexGrid1.TextMatrix(0, E.CMApprovedIncome) = 0
            MSHFlexGrid1.TextMatrix(0, E.CMPendingIncome) = 0
            MSHFlexGrid1.TextMatrix(0, E.CMUnbilledIncome) = 0
            For i = 2 To MSHFlexGrid1.Rows - 2
               MSHFlexGrid1.TextMatrix(i, E.CMApprovedIncome) = 0
               MSHFlexGrid1.TextMatrix(i, E.CMPendingIncome) = 0
               MSHFlexGrid1.TextMatrix(i, E.CMUnbilledIncome) = 0
            Next
            Call GetCMIncome
         End If
         MSHFlexGrid1.TextMatrix(0, E.CoExpense) = 0
         For i = 2 To MSHFlexGrid1.Rows - 2
            MSHFlexGrid1.TextMatrix(i, E.CoExpense) = 0
         Next
         Call GetCoExpense
      End If
      
      GoTo Detail:
   Else
      MsgBox "No record found!!!"
      Exit Sub
   End If
End If

tmp = " Select MS.StateName, ML.LocationID, ML.LocationName, MCMP.CMPName, MG.GodownID, MG.GodownNo, MG.Address, MG.GodownCapacity, MUT.UnitType, MG.StartDate, MG.CloseDate,  MS.StateID "
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Inner Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UNitTypeID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
If Wc <> "" Then
    tmp = tmp & Wc & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
Else
    tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
End If
tmp = tmp & " Order by MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo"
Call TmpTable

Label5.Caption = TmpRs.RecordCount
If TmpRs.RecordCount > 0 Then
   ProgressBar1.Value = 0
   ProgressBar1.Max = MSHFlexGrid1.Rows + TmpRs.RecordCount
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
   CmdExcel.Enabled = True
   If (LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "head cmg") And CmbLocationID.Text = "" And CmbStateID.Text = "" Then
      Cmd_Process.Enabled = True
   End If
   For i = 2 To TmpRs.RecordCount + 1
      With MSHFlexGrid1
         For j = 10 To 48
            .TextMatrix(i, j) = 0
         Next
         .TextMatrix(i, E.MonthYear) = mYearMonth
         .TextMatrix(i, E.StateName) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
         .TextMatrix(i, E.LocationName) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         .TextMatrix(i, E.CMPName) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
         .TextMatrix(i, E.GodownNo) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
         .TextMatrix(i, E.GodownCapacity) = IIf(IsNull(TmpRs!GodownCapacity), "", TmpRs!GodownCapacity)
         .TextMatrix(0, E.GodownCapacity) = Val(.TextMatrix(0, E.GodownCapacity)) + Val(.TextMatrix(i, E.GodownCapacity))
         .TextMatrix(i, E.GodownStartDate) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-MMM-yyyy"))
         .TextMatrix(i, E.GodownCloseDate) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-MMM-yyyy"))
         .TextMatrix(i, E.UnitType) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
         .TextMatrix(i, E.StateID) = IIf(IsNull(TmpRs!StateID), "", TmpRs!StateID)
         .TextMatrix(i, E.LocationID) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
         .TextMatrix(i, E.GodownID) = IIf(IsNull(TmpRs!GodownID), "", TmpRs!GodownID)
      End With
      TmpRs.MoveNext
      ProgressBar1.Value = ProgressBar1.Value + 1
   Next i
   Call GetAMRMName
   Call GetNCDEXStock
   Call GetNonNCDEXStock
   Call GetNNNRIncome
   Call GetNNNRIncomethroughAdhoc
   Call GetReservationIncome
   Call GetReservationIncomethroughAdhoc
   Call GetSevenDaysIncome
   Call GetDPIncome
   Call GetDPIncomethroughAdhoc
   Call GetNCDEXRejectedIncome
   Call GetOtherIncome
   Call GetSCMIncome
   Call GetSpillageIncome
   Call GetPledgeIncome
   Call GetExpenseEmployeeClaim
   Call GetVendorExpense
   Call GetSecurityExpense
   Call GetFumigationExpense
   Call GetRentExpense
   Call GetFranchiseeExpense
'   Call GetSalaryExpense
   Call GetSalaryExpenseFromIncomeforFieldStaff
   Call GetSalaryExpenseFromIncomeforCOStaff
   Call GetInsuranceCharges
   Call GetOutstandingInterest
   Call GetSpotIncomethroughAdhoc
   Call GetCMIncome
   Call GetCoExpense
Detail:
   
   '//////////////////*Claim Amount from HR Pro Locations not open in current month
If ChkFromSavedData.Value = vbUnchecked Or ChkEmployeeClaim.Value = vbChecked Then
   tmp = "Select distinct MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
   Call Tmp7Table
   If TmpRs7.RecordCount > 0 Then
      mLocationIDs = "(0"
      For k = 1 To TmpRs7.RecordCount
         mLocationIDs = mLocationIDs & "," & IIf(IsNull(TmpRs7!LocationID), 0, TmpRs7!LocationID)
         TmpRs7.MoveNext
      Next
      mLocationIDs = mLocationIDs & ")"
   End If
   'tmp = " Select TCS.Flag, Sum(TCSD.Amount) Amount"
   tmp = " Select Sum(TCSD.Amount) Amount"
   tmp = tmp & " From Txn_ClaimSegmentDistribution TCSD"
   tmp = tmp & " Inner Join Txn_ClaimSubmission TCS on TCS.ClaimID = TCSD.ClaimID"
   tmp = tmp & " Where TCSD.SegmentID = 1 And TCS.Flag not in ('R')"
   tmp = tmp & " And TCSD.LocationID not in " & mLocationIDs
   tmp = tmp & " And '" & mStartDate & "' Between TCS.ExpenseFrom And TCS.ExpenseTo "
   'group by TCS.Flag"
   Call TmpTableHRPro
   
   If TmpRsHRPro.RecordCount > 0 Then
      With MSHFlexGrid1
         If ChkFromSavedData.Value = vbChecked And .TextMatrix(.Rows - 2, E.AMname) = "CMG-EmployeeClaim of Closed Locations" Then
            k = .Rows - 2
            .TextMatrix(0, E.EmpClaimFAG) = Val(.TextMatrix(0, E.EmpClaimFAG)) - Val(.TextMatrix(k, E.EmpClaimFAG))
            .TextMatrix(0, E.EmpClaimAccrual) = Val(.TextMatrix(0, E.EmpClaimAccrual)) - Val(.TextMatrix(k, E.EmpClaimAccrual))
         ElseIf ChkFromSavedData.Value = vbChecked And .TextMatrix(.Rows - 3, E.AMname) = "CMG-EmployeeClaim of Closed Locations" Then
            k = .Rows - 3
            .TextMatrix(0, E.EmpClaimFAG) = Val(.TextMatrix(0, E.EmpClaimFAG)) - Val(.TextMatrix(k, E.EmpClaimFAG))
            .TextMatrix(0, E.EmpClaimAccrual) = Val(.TextMatrix(0, E.EmpClaimAccrual)) - Val(.TextMatrix(k, E.EmpClaimAccrual))
         Else
            k = .Rows - 1
            .Rows = .Rows + 1
         End If
         .TextMatrix(k, E.MonthYear) = CmbMonth.Text & "-" & CmbYear.Text
         '.TextMatrix(k, E.StateName) = "CMG-EmployeeClaim of Closed Locations"
         '.TextMatrix(k, E.LocationName) = "CMG-EmployeeClaim of Closed Locations"
         .TextMatrix(k, E.AMname) = "CMG-EmployeeClaim of Closed Locations"
         .TextMatrix(k, E.RMName) = "CMG-EmployeeClaim of Closed Locations"
         '.TextMatrix(k, E.CMPName) = "CMG-EmployeeClaim of Closed Locations"
         '.TextMatrix(k, E.GodownNo) = "CMG-EmployeeClaim of Closed Locations"
         .TextMatrix(k, E.EmpClaimFAG) = IIf(IsNull(TmpRsHRPro!Amount), "", TmpRsHRPro!Amount)
         .TextMatrix(0, E.EmpClaimFAG) = Val(.TextMatrix(0, E.EmpClaimFAG)) + Val(.TextMatrix(k, E.EmpClaimFAG))
         If TmpRsHRPro.RecordCount > 1 Then
            TmpRsHRPro.MoveNext
            .TextMatrix(k, E.EmpClaimAccrual) = IIf(IsNull(TmpRsHRPro!Amount), "", TmpRsHRPro!Amount)
            .TextMatrix(0, E.EmpClaimAccrual) = Val(.TextMatrix(0, E.EmpClaimAccrual)) + Val(.TextMatrix(k, E.EmpClaimAccrual))
         End If
      End With
   End If
End If
   'Vendor Expense Amount from HR Pro Locations not open in current month*//////////////////'
If ChkFromSavedData.Value = vbUnchecked Or ChkVendorExpense.Value = vbChecked Then
   tmp = "Select distinct MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
   Call Tmp7Table
   If TmpRs7.RecordCount > 0 Then
      mLocationIDs = "(0"
      For k = 1 To TmpRs7.RecordCount
         mLocationIDs = mLocationIDs & "," & IIf(IsNull(TmpRs7!LocationID), 0, TmpRs7!LocationID)
         TmpRs7.MoveNext
      Next
      mLocationIDs = mLocationIDs & ")"
   End If
   
   tmp = " Select TCS.Flag, Sum(TCSD.Amount) Amount"
   tmp = tmp & " From Txn_VendorClaimSegmentDistribution TCSD"
   tmp = tmp & " Inner Join Txn_VendorClaim TCS on TCS.ClaimID = TCSD.ClaimID"
   tmp = tmp & " Where TCSD.SegmentID = 1 And TCS.Flag in ('A', 'P') " 'And TCSD.Flag = 'L'"
   tmp = tmp & " And TCSD.LocationID not in " & mLocationIDs
   tmp = tmp & " And '" & mStartDate & "' Between TCS.ExpenseFrom And TCS.ExpenseTo And TCSD.ExpenseTypeID not in (107,66) group by TCS.Flag"
   Call TmpTableHRPro
   If TmpRsHRPro.RecordCount > 0 Then
      With MSHFlexGrid1
         If ChkFromSavedData.Value = vbChecked And .TextMatrix(.Rows - 2, E.AMname) = "CMG-VendorClaim of Closed Locations" Then
            k = .Rows - 2
            .TextMatrix(0, E.VendorExpenses) = Val(.TextMatrix(0, E.VendorExpenses)) - Val(.TextMatrix(k, E.VendorExpenses))
            .TextMatrix(0, E.VendorExpensesAccrued) = Val(.TextMatrix(0, E.VendorExpensesAccrued)) - Val(.TextMatrix(k, E.VendorExpensesAccrued))
         ElseIf ChkFromSavedData.Value = vbChecked And .TextMatrix(.Rows - 3, E.AMname) = "CMG-VendorClaim of Closed Locations" Then
            k = .Rows - 3
            .TextMatrix(0, E.VendorExpenses) = Val(.TextMatrix(0, E.VendorExpenses)) - Val(.TextMatrix(k, E.VendorExpenses))
            .TextMatrix(0, E.VendorExpensesAccrued) = Val(.TextMatrix(0, E.VendorExpensesAccrued)) - Val(.TextMatrix(k, E.VendorExpensesAccrued))
         Else
            k = .Rows - 1
            .Rows = .Rows + 1
         End If
         
         .TextMatrix(k, E.MonthYear) = CmbMonth.Text & "-" & CmbYear.Text
         '.TextMatrix(k, E.StateName) = "CMG-VendorClaim of Closed Locations"
         '.TextMatrix(k, E.LocationName) = "CMG-VendorClaim of Closed Locations"
         .TextMatrix(k, E.AMname) = "CMG-VendorClaim of Closed Locations"
         .TextMatrix(k, E.RMName) = "CMG-VendorClaim of Closed Locations"
         '.TextMatrix(k, E.CMPName) = "CMG-VendorClaim of Closed Locations"
         '.TextMatrix(k, E.GodownNo) = "CMG-VendorClaim of Closed Locations"
         .TextMatrix(k, E.VendorExpenses) = IIf(IsNull(TmpRsHRPro!Amount), "", TmpRsHRPro!Amount)
         .TextMatrix(0, E.VendorExpenses) = Val(.TextMatrix(0, E.VendorExpenses)) + Val(.TextMatrix(k, E.VendorExpenses))
         If TmpRsHRPro.RecordCount > 1 Then
            TmpRsHRPro.MoveNext
            .TextMatrix(k, E.VendorExpensesAccrued) = IIf(IsNull(TmpRsHRPro!Amount), "", TmpRsHRPro!Amount)
            .TextMatrix(0, E.VendorExpensesAccrued) = Val(.TextMatrix(0, E.VendorExpensesAccrued)) + Val(.TextMatrix(k, E.VendorExpensesAccrued))
         End If
      End With
   End If


''------------------------
'
'
'   tmp = "Select distinct MG.GodownID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
'   Call Tmp7Table
'   If TmpRs7.RecordCount > 0 Then
'      mLocationIDs = "(0"
'      For k = 1 To TmpRs7.RecordCount
'         mLocationIDs = mLocationIDs & "," & IIf(IsNull(TmpRs7!GodownID), 0, TmpRs7!GodownID)
'         TmpRs7.MoveNext
'      Next
'      mLocationIDs = mLocationIDs & ")"
'   End If
'
'   tmp = " Select TCS.Flag, Sum(TCSD.Amount) Amount"
'   tmp = tmp & " From Txn_VendorClaimSegmentDistribution TCSD"
'   tmp = tmp & " Inner Join Txn_VendorClaim TCS on TCS.ClaimID = TCSD.ClaimID"
'   tmp = tmp & " Where TCSD.SegmentID = 1 And TCS.Flag in ('A', 'P') And TCSD.Flag = 'G'"
'   tmp = tmp & " And TCSD.LocationID not in " & mLocationIDs
'   tmp = tmp & " And '" & mStartDate & "' Between TCS.ExpenseFrom And TCS.ExpenseTo And TCSD.ExpenseTypeID not in (107,66) group by TCS.Flag"
'   Call TmpTableHRPro
'   If TmpRsHRPro.RecordCount > 0 Then
'      With MSHFlexGrid1
'         If ChkFromSavedData.Value = vbChecked And .TextMatrix(.Rows - 2, E.AMname) = "CMG-VendorClaim of Closed Locations" Then
'            k = .Rows - 2
'            .TextMatrix(0, E.VendorExpenses) = Val(.TextMatrix(0, E.VendorExpenses)) - Val(.TextMatrix(k, E.VendorExpenses))
'            .TextMatrix(0, E.VendorExpensesAccrued) = Val(.TextMatrix(0, E.VendorExpensesAccrued)) - Val(.TextMatrix(k, E.VendorExpensesAccrued))
'         ElseIf ChkFromSavedData.Value = vbChecked And .TextMatrix(.Rows - 3, E.AMname) = "CMG-VendorClaim of Closed Locations" Then
'            k = .Rows - 3
'            .TextMatrix(0, E.VendorExpenses) = Val(.TextMatrix(0, E.VendorExpenses)) - Val(.TextMatrix(k, E.VendorExpenses))
'            .TextMatrix(0, E.VendorExpensesAccrued) = Val(.TextMatrix(0, E.VendorExpensesAccrued)) - Val(.TextMatrix(k, E.VendorExpensesAccrued))
'         Else
'            k = .Rows - 1
'            .Rows = .Rows + 1
'         End If
'
'         .TextMatrix(k, E.MonthYear) = CmbMonth.Text & "-" & CmbYear.Text
'         '.TextMatrix(k, E.StateName) = "CMG-VendorClaim of Closed Locations"
'         '.TextMatrix(k, E.LocationName) = "CMG-VendorClaim of Closed Locations"
'         .TextMatrix(k, E.AMname) = "CMG-VendorClaim of Closed Locations"
'         .TextMatrix(k, E.RMName) = "CMG-VendorClaim of Closed Locations"
'         '.TextMatrix(k, E.CMPName) = "CMG-VendorClaim of Closed Locations"
'         '.TextMatrix(k, E.GodownNo) = "CMG-VendorClaim of Closed Locations"
'         .TextMatrix(k, E.VendorExpenses) = IIf(IsNull(TmpRsHRPro!Amount), "", TmpRsHRPro!Amount)
'         .TextMatrix(0, E.VendorExpenses) = Val(.TextMatrix(0, E.VendorExpenses)) + Val(.TextMatrix(k, E.VendorExpenses))
'         If TmpRsHRPro.RecordCount > 1 Then
'            TmpRsHRPro.MoveNext
'            .TextMatrix(k, E.VendorExpensesAccrued) = IIf(IsNull(TmpRsHRPro!Amount), "", TmpRsHRPro!Amount)
'            .TextMatrix(0, E.VendorExpensesAccrued) = Val(.TextMatrix(0, E.VendorExpensesAccrued)) + Val(.TextMatrix(k, E.VendorExpensesAccrued))
'         End If
'      End With
'   End If


End If

With MSHFlexGrid1
   For i = 2 To MSHFlexGrid1.Rows - 2
      .TextMatrix(i, E.TotalIncome) = Val(.TextMatrix(i, E.NNNRIncomeAdhoc)) + Val(.TextMatrix(i, E.NNNRIncome)) + Val(.TextMatrix(i, E.ReservationIncomeAdhoc)) + Val(.TextMatrix(i, E.ReservationIncome)) + Val(.TextMatrix(i, E.SevenDaysIncome)) + Val(.TextMatrix(i, E.DPIncomeAdHoc)) + Val(.TextMatrix(i, E.FinalDPIncome)) + Val(.TextMatrix(i, E.RejectedIncome)) + Val(.TextMatrix(i, E.SpotIncomeAdHoc)) + Val(.TextMatrix(i, E.OtherIncome)) + Val(.TextMatrix(i, E.SCMIncome)) + Val(.TextMatrix(i, E.CMUnbilledIncome)) + Val(.TextMatrix(i, E.CMPendingIncome)) + Val(.TextMatrix(i, E.CMApprovedIncome)) + Val(.TextMatrix(i, E.LeftOverStock)) + Val(.TextMatrix(i, E.PledgeIncome))
      .TextMatrix(i, E.TotalIncome) = Format(.TextMatrix(i, E.TotalIncome), "############.00")
      .TextMatrix(i, E.TotalExpense) = Val(.TextMatrix(i, E.LeaseRentPaid)) + Val(.TextMatrix(i, E.LeaseRentFAG)) + Val(.TextMatrix(i, E.LeaseRentDRC)) + Val(.TextMatrix(i, E.EmpClaimFAG)) + Val(.TextMatrix(i, E.EmpClaimAccrual)) + Val(.TextMatrix(i, E.VendorExpenses)) + Val(.TextMatrix(i, E.VendorExpensesAccrued)) + Val(.TextMatrix(i, E.SecurityFeesPaid)) + Val(.TextMatrix(i, E.SecurityFeesAccrual)) + Val(.TextMatrix(i, E.FumigationFeesVendor)) + Val(.TextMatrix(i, E.FumigationFeesPaid)) + Val(.TextMatrix(i, E.FumigationFeesAccrual)) + Val(.TextMatrix(i, E.FranchiseeFeesPaid)) + Val(.TextMatrix(i, E.FranchiseeFeesCMG)) + Val(.TextMatrix(i, E.FranchiseeFeesFAG)) + Val(.TextMatrix(i, E.Salary)) + Val(.TextMatrix(i, E.SalaryFieldStaff)) + Val(.TextMatrix(i, E.NCDEXCharges))
      .TextMatrix(i, E.TotalExpense) = Format(.TextMatrix(i, E.TotalExpense), "############.00")
      .TextMatrix(i, E.GrossIncome) = Val(.TextMatrix(i, E.TotalIncome)) - Val(.TextMatrix(i, E.TotalExpense))
      .TextMatrix(i, E.GrossIncome) = Format(.TextMatrix(i, E.GrossIncome), "############.00")
      .TextMatrix(i, E.NetIncome) = Val(.TextMatrix(i, E.GrossIncome)) - Val(.TextMatrix(i, E.CoExpense)) - Val(.TextMatrix(i, E.InsurancePremium)) - Val(.TextMatrix(i, E.InterestCharges))
      .TextMatrix(i, E.NetIncome) = Format(.TextMatrix(i, E.NetIncome), "############.00")
      .TextMatrix(0, E.TotalIncome) = Format(Val(.TextMatrix(0, E.TotalIncome)) + Val(.TextMatrix(i, E.TotalIncome)), "############.00")
      .TextMatrix(0, E.TotalExpense) = Format(Val(.TextMatrix(0, E.TotalExpense)) + Val(.TextMatrix(i, E.TotalExpense)), "############.00")
      .TextMatrix(0, E.GrossIncome) = Format(Val(.TextMatrix(0, E.GrossIncome)) + Val(.TextMatrix(i, E.GrossIncome)), "############.00")
      .TextMatrix(0, E.NetIncome) = Format(Val(.TextMatrix(0, E.NetIncome)) + Val(.TextMatrix(i, E.NetIncome)), "############.00")
   Next
End With

   If CheckStock.Value = vbChecked Then
      Call GetStockDetail
   End If
   If CheckNNNR.Value = vbChecked Then
      Call GetNNNRIncomeDetail
   End If
   If CheckReservation.Value = vbChecked Then
      Call GetReservationIncomeDetail
   End If
   If CheckInvoice.Value = vbChecked Then
      Call GetSevenDaysIncomeDetail
   End If
   If CheckDPIncome.Value = vbChecked Then
      Call GetDPIncomeDetail
   End If
   If CheckPledge.Value = vbChecked Then
      Call GetPledgeIncomeDetail
   End If
   If CheckVendor.Value = vbChecked Then
      Call GetVendorExpenseDetail
   End If
   If CheckSecurity.Value = vbChecked Then
      Call GetSecurityExpenseDetail
   End If
   If CheckGodownRent.Value = vbChecked Then
      Call GetRentExpenseDetail
   End If
   If CheckDPGSL.Value = vbChecked Then
      Call GetDPGSLInvoiceDetail
   End If
   If CheckFranchisee.Value = vbChecked Then
      Call GetFranchiseeFeesDetail
   End If
   If MSHFlexGrid13.Rows = 3 And CheckInterest.Value = vbChecked Then
      MSHFlexGrid1.TextMatrix(0, E.InterestCharges) = 0
      For i = 2 To MSHFlexGrid1.Rows - 2
         MSHFlexGrid1.TextMatrix(i, E.InterestCharges) = 0
      Next
      Call GetOutstandingInterest
   End If
   If CheckReversal.Value = vbChecked Then
      Call GetReversalDetail
   End If
   If CheckClaimExpenseHRPro.Value = vbChecked Then
      Call GetClaimExpenseHRPro
   End If
   If CheckSalaryofEmployee.Value = vbChecked Then
      Call GetSalaryofEmployee
   End If
   ProgressBar1.Value = ProgressBar1.Max
   MsgBox "Done!!!"
Else
    MsgBox "No Record Found!!!"
End If

End Sub

Private Sub GetClaimExpenseHRPro()

tmp = " Select LocationID,LocationName,StateName from Mst_Location ML"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
Call Tmp7Table

tmp = " Select ML.LocationID,LocationName,StateName,MG.GodownID from Mst_Location ML"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Inner Join Mst_CMP MC On ML.LocationID = MC.LocationID"
tmp = tmp & " Inner Join Mst_Godown MG On MC.CMPID = MG.CMPID"
Call Tmp6Table


tmp = " Select TC.ClaimID,MB.SegmentName,TCA.SegmentID,TCA.LocationID, TC.ExpenseFrom ,MEx.ExpenseType, TCA.Amount,"
tmp = tmp & " TC.Flag ,ME.FullName + '~' + ME.EmployeeNo  as Employee, MD.Designation, 'Employee' EmpVendor, TC.ApprovalRemarks,"
tmp = tmp & " TC.ApprovalDate, TC.BatchID,'L' 'DisFlag' From Txn_ClaimSegmentDistribution TCA"
tmp = tmp & " inner Join Txn_ClaimSubmission TC on TCA.ClaimID=TC.ClaimID"
tmp = tmp & " Inner Join Mst_Employee ME On TC.EmployeeID = ME.EmployeeID"
tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID"
tmp = tmp & " inner Join Mst_ExpenseType MEx on TCA.ExpenseTypeID=MEx.ExpenseTypeID"
tmp = tmp & " Inner Join Mst_BusinessSegment MB On TCA.SegmentID = MB.BusinessSegmentID "
tmp = tmp & " Where TC.Flag not in ('R') and TC.ExpenseFrom >= '" & mStartDate & "' And TC.ExpenseTo <= '" & mEndDate & "' And TCA.SegmentID =1"
tmp = tmp & " Union"
tmp = tmp & " Select TC.ClaimID,MB.SegmentName,TCA.SegmentID,TCA.LocationID, TC.ExpenseFrom ,MEx.ExpenseType, TCA.Amount,"
tmp = tmp & " TC.Flag ,MV.VendorName Employee, '' Designation, 'Vendor' EmpVendor, TC.ApprovalRemarks, TC.ApprovalDate, "
tmp = tmp & " TC.BatchID,TCA.Flag 'DisFlag' From Txn_VendorClaimSegmentDistribution TCA "
tmp = tmp & " inner Join Txn_VendorClaim TC on TCA.ClaimID=TC.ClaimID "
tmp = tmp & " Inner Join Mst_Vendor MV On TC.VendorID = MV.VendorID "
tmp = tmp & " inner Join Mst_ExpenseType MEx on TCA.ExpenseTypeID=MEx.ExpenseTypeID "
tmp = tmp & " Inner Join Mst_BusinessSegment MB On TCA.SegmentID = MB.BusinessSegmentID"
tmp = tmp & " Where TC.Flag in ('A', 'P') and TC.ExpenseFrom >= '" & mStartDate & "' And TC.ExpenseTo <= '" & mEndDate & "' And TCA.SegmentID =1 And TCA.ExpenseTypeID not in (107,66)" 'And TCA.ExpenseTypeID not in (107)"
Call TmpTableHRPro

Call Grid_Head16
ProgressBar1.Value = 0

If TmpRsHRPro.RecordCount > 0 Then
   MSHFlexGrid16.Rows = MSHFlexGrid16.Rows + TmpRsHRPro.RecordCount
   ProgressBar1.Max = MSHFlexGrid16.Rows
   With MSHFlexGrid16
      For i = 2 To TmpRsHRPro.RecordCount + 1
         If LCase(TmpRsHRPro!DisFlag) = "l" Then
            If TmpRs7.RecordCount > 0 Then TmpRs7.MoveFirst
            TmpRs7.Find " LocationID=" & TmpRsHRPro!LocationID
            If TmpRs7.EOF = False Then
               .TextMatrix(i, 3) = TmpRs7!LocationName
               .TextMatrix(i, 2) = TmpRs7!StateName
            End If
         ElseIf LCase(TmpRsHRPro!DisFlag) = "g" Then
            If TmpRs6.RecordCount > 0 Then TmpRs6.MoveFirst
            TmpRs6.Find " GodownID=" & TmpRsHRPro!LocationID
            If TmpRs6.EOF = False Then
               .TextMatrix(i, 3) = TmpRs6!LocationName
               .TextMatrix(i, 2) = TmpRs6!StateName
            End If
         End If
         .TextMatrix(i, 0) = IIf(IsNull(TmpRsHRPro!ClaimID), "", TmpRsHRPro!ClaimID)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRsHRPro!SegmentName), "", TmpRsHRPro!SegmentName)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRsHRPro!Employee), "", TmpRsHRPro!Employee)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRsHRPro!Designation), "", TmpRsHRPro!Designation)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRsHRPro!ExpenseFrom), "", Format(TmpRsHRPro!ExpenseFrom, "dd-MMM-yyyy"))
         .TextMatrix(i, 7) = IIf(IsNull(TmpRsHRPro!ExpenseType), "", TmpRsHRPro!ExpenseType)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRsHRPro!Amount), 0, TmpRsHRPro!Amount)
         .TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) + Val(.TextMatrix(i, 8))
         .TextMatrix(i, 9) = IIf(IsNull(TmpRsHRPro!Flag), "", TmpRsHRPro!Flag)
         .TextMatrix(i, 10) = IIf(IsNull(TmpRsHRPro!EmpVendor), "", TmpRsHRPro!EmpVendor)
         .TextMatrix(i, 11) = IIf(IsNull(TmpRsHRPro!ApprovalDate), "", Format(TmpRsHRPro!ApprovalDate, "dd-MMM-yyyy"))
         .TextMatrix(i, 12) = IIf(IsNull(TmpRsHRPro!ApprovalRemarks), "", TmpRsHRPro!ApprovalRemarks)
         .TextMatrix(i, 13) = IIf(IsNull(TmpRsHRPro!DisFlag), "", TmpRsHRPro!DisFlag)
         If .TextMatrix(i, 9) = "A" Then
            .TextMatrix(i, 9) = "Approved"
         ElseIf .TextMatrix(i, 9) = "P" Then
            .TextMatrix(i, 9) = "Pending"
         End If
         TmpRsHRPro.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
End If
End Sub

Private Sub GetSalaryofEmployee()

tmp = " select me.fullname+'~'+me.employeeno employeename, mbs.segmentname, md.department, mdg.designation, mct.ContractType , IsNull(tmsd.expallocation / 100, 0) 'countt'"
tmp = tmp & " From Txn_MonthEmp_Salary tms"
tmp = tmp & " inner join mst_employee me on me.employeeid =  tms.employeeid"
tmp = tmp & " inner join mst_contracttype mct on mct.contracttypeid = me.contracttypeid"
tmp = tmp & " inner join mst_department md on md.departmentid = me.departmentid"
tmp = tmp & " inner join mst_designation mdg on mdg.designationid = me.designationid"
tmp = tmp & " inner join Txn_EmpSegmentMapping tes on tes.employeeid = tms.employeeid"
tmp = tmp & " inner join Txn_EmpSegmentMappingDetail tmsd on  tmsd.EmpSegMapID = tes.EmpSegMapID"
tmp = tmp & " inner join mst_businesssegment mbs on mbs.businesssegmentid = tmsd.segmentid"
tmp = tmp & " where tms.monthyear Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " and IsNull(tmsd.expallocation / 100, 0) > 0"
tmp = tmp & " and tmsd.segmentid = 1"
Call TmpTableHRPro

Call Grid_Head17
ProgressBar1.Value = 0

If TmpRsHRPro.RecordCount > 0 Then
   MSHFlexGrid17.Rows = MSHFlexGrid17.Rows + TmpRsHRPro.RecordCount
   ProgressBar1.Max = MSHFlexGrid17.Rows
   With MSHFlexGrid17
      For i = 2 To TmpRsHRPro.RecordCount + 1
         .TextMatrix(i, 0) = IIf(IsNull(TmpRsHRPro!EmployeeName), "", TmpRsHRPro!EmployeeName)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRsHRPro!SegmentName), "", TmpRsHRPro!SegmentName)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRsHRPro!Department), "", TmpRsHRPro!Department)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRsHRPro!Designation), "", TmpRsHRPro!Designation)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRsHRPro!ContractType), "", TmpRsHRPro!ContractType)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRsHRPro!Countt), "", TmpRsHRPro!Countt)
         .TextMatrix(0, 5) = Val(.TextMatrix(0, 5)) + Val(.TextMatrix(i, 5))
         TmpRsHRPro.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
End If
End Sub


Private Sub GetReversalDetail()
tmp = " Select"
tmp = tmp & " ms.StateName , ML.LocationName, MET.ExpenseType, TCAD.TIID, TCAD.RecTDS, TCAD.RecOtherDed, MCL.ClientName, MD.DepositoryName"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCAD"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCA on TCA.TDSAID = TCAD.TDSAID And TCA.SM_Prefix = TCAD.SM_Prefix"
tmp = tmp & " Inner Join mst_expensetype MET on MET.ExpenseTypeID = TCAD.ExpenseAccID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TCA.LocationID"
tmp = tmp & " Inner JOin Mst_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Left Join Mst_Client MCL on MCL.ClientIDRow = TCA.ClientIDRow"
tmp = tmp & " Left Join Mst_Depository MD on MD.DepositoryID = TCA.DPID"
If Wc <> "" Then
   tmp = tmp & Wc & " And "
Else
   tmp = tmp & " Where "
End If
tmp = tmp & " TCA.Flag in ('A', 'P') And AdjDate Between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And TCAD.ExpenseAccID in (88,89,90,91,92)"
Call Tmp1Table

Call Grid_Head15
ProgressBar1.Value = 0

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid15.Rows = MSHFlexGrid15.Rows + TmpRs1.RecordCount
   ProgressBar1.Max = MSHFlexGrid15.Rows
   With MSHFlexGrid15
      For i = 2 To TmpRs1.RecordCount + 1
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!ExpenseType), "", TmpRs1!ExpenseType)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs1!DepositoryName), "", TmpRs1!DepositoryName)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs1!TIID), "", TmpRs1!TIID)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs1!RecTDS), "", TmpRs1!RecTDS)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs1!RecOtherDed), "", TmpRs1!RecOtherDed)
         .TextMatrix(0, 6) = Val(.TextMatrix(0, 6)) + Val(.TextMatrix(i, 6))
         .TextMatrix(0, 7) = Val(.TextMatrix(0, 7)) + Val(.TextMatrix(i, 7))
         TmpRs1.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
End If
End Sub


Private Sub GetDPGSLInvoiceDetail()
tmp = " Select MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo, MG.Address, MCL.ClientName, MCom.Commodity, TIMG.StackNo,"
tmp = tmp & " TIMG.LotNo, TIMG.BalanceBags, TIMG.BalanceWeight, isnull(TIMG.InvoiceAmount,0) IncomeAmt, isnull(TIMG.DPIncome,0) DPIncome,  isnull(TIMG.CommtrackIncome,0) CommtrackIncome From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_Location ML On TIMG.SM_Prefix =ML.SM_Prefix"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = TIMG.CMPID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Inner Join Mst_Client MCL on MCL.ClientIDRow = TIMG.ClientIDRow"
tmp = tmp & " Inner Join Mst_Commodity MCom on MCom.CommodityID = TIMG.CommodityID"
If Wc <> "" Then
    tmp = tmp & Wc & " And "
Else
    tmp = tmp & " Where "
End If
tmp = tmp & " TIMG.YearMonth between '" & mStartDate & "' and '" & mEndDate & "' And TIMG.ExchangeTypeID = 2 "
tmp = tmp & " And TIMG.BalanceWeight > 0"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " And (TIMG.InvoiceAmount > 0 or TIMG.DPIncome > 0 Or TIMG.CommtrackIncome > 0 )"

Call Tmp1Table

Call Grid_Head11
ProgressBar1.Value = 0

If TmpRs1.RecordCount > 0 Then
    MSHFlexGrid11.Rows = MSHFlexGrid11.Rows + TmpRs1.RecordCount
    ProgressBar1.Max = MSHFlexGrid11.Rows
    With MSHFlexGrid11
        For i = 2 To TmpRs1.RecordCount + 1
            .TextMatrix(i, 0) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
            .TextMatrix(i, 1) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
            .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!CMPName), "", TmpRs1!CMPName)
            .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!GodownNo), "", TmpRs1!GodownNo)
            .TextMatrix(i, 4) = IIf(IsNull(TmpRs1!Address), "", TmpRs1!Address)
            .TextMatrix(i, 5) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
            .TextMatrix(i, 6) = IIf(IsNull(TmpRs1!Commodity), "", TmpRs1!Commodity)
            .TextMatrix(i, 7) = IIf(IsNull(TmpRs1!StackNo), "", TmpRs1!StackNo)
            .TextMatrix(i, 8) = IIf(IsNull(TmpRs1!LotNo), "", TmpRs1!LotNo)
            .TextMatrix(i, 9) = IIf(IsNull(TmpRs1!BalanceBags), "", TmpRs1!BalanceBags)
            .TextMatrix(i, 10) = IIf(IsNull(TmpRs1!BalanceWeight), "", TmpRs1!BalanceWeight)
            .TextMatrix(i, 11) = IIf(IsNull(TmpRs1!IncomeAmt), "", TmpRs1!IncomeAmt)
            .TextMatrix(0, 11) = Val(.TextMatrix(0, 11)) + Val(.TextMatrix(i, 11))
            .TextMatrix(i, 12) = IIf(IsNull(TmpRs1!DPIncome), "", TmpRs1!DPIncome)
            .TextMatrix(0, 12) = Val(.TextMatrix(0, 12)) + Val(.TextMatrix(i, 12))
            .TextMatrix(i, 13) = IIf(IsNull(TmpRs1!CommtrackIncome), "", TmpRs1!CommtrackIncome)
            .TextMatrix(0, 13) = Val(.TextMatrix(0, 13)) + Val(.TextMatrix(i, 13))
            TmpRs1.MoveNext
            ProgressBar1.Value = ProgressBar1.Value + 1
        Next
    End With
End If
End Sub

Private Sub GetRentExpenseDetail()
tmp = " Select MS.StateName, ML.LocationID, ML.LocationName, MCMP.CMPName, MG.GodownID, MG.GodownNo, MG.Address, MG.StartDate, MG.CloseDate, MUT.UnitType,  MG.GodownCapacity, MS.StateID, "
tmp = tmp & " (select top 1 advancegiven from Txn_LessorAgreement TLA inner join Txn_GodownLessorAgreeDetail tlad on tlad.laid = tla.laid where tlad.godownid = mg.godownid order by startdate desc) securitydeposite"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Inner Join Mst_UnitType MUT on MUT.UnitTypeID = MG.UNitTypeID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
If Wc <> "" Then
    tmp = tmp & Wc & " And "
Else
    tmp = tmp & " Where "
End If
tmp = tmp & " (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " And MUT.UnitTypeID in (1,6) " ' 1 and 6 means franchisee or leased franchisee
tmp = tmp & " Order by MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo"
Call Tmp1Table

Call Grid_Head10
ProgressBar1.Value = 0

If TmpRs1.RecordCount > 0 Then
    MSHFlexGrid10.Rows = MSHFlexGrid10.Rows + TmpRs1.RecordCount
    ProgressBar1.Max = MSHFlexGrid10.Rows
    With MSHFlexGrid10
        For i = 2 To TmpRs1.RecordCount + 1
            .TextMatrix(i, 0) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName) '"State Name"
            .TextMatrix(i, 1) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName) '"Location Name"
            .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!CMPName), "", TmpRs1!CMPName) '"CMP Name"
            .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!GodownNo), "", TmpRs1!GodownNo) '"Godown No"
            .TextMatrix(i, 4) = IIf(IsNull(TmpRs1!StartDate), "", Format(TmpRs1!StartDate, "dd-MMM-yyyy")) '"Start Date"
            .TextMatrix(i, 5) = IIf(IsNull(TmpRs1!CloseDate), "", Format(TmpRs1!CloseDate, "dd-MMM-yyyy")) '"Close Date"
            .TextMatrix(i, 6) = IIf(IsNull(TmpRs1!Address), "", TmpRs1!Address) '"Godown Address"
            .TextMatrix(i, 7) = IIf(IsNull(TmpRs1!GodownID), "", TmpRs1!GodownID) '"Godown ID"
            TmpRs1.MoveNext
            ProgressBar1.Value = ProgressBar1.Value + 1
        Next
    End With
End If

'mTotalDays = Day(DateSerial(Year(TxtFromDate.Value), Month(TxtFromDate.Value) + 1 + x_ - 9, 0))

tmp = " Select TGRD.GodownID, sum(TGRD.Amount) As IncomeAmt"
tmp = tmp & " From Txn_GodownRentDetail TGRD"
tmp = tmp & " Inner Join Txn_GodownRent TGR on TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TGRD.GodownID"
tmp = tmp & " Where TGR.Flag in ('A', 'P') And  TGR.ExpenseFrom Between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group by TGRD.GodownID"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
    For i = 2 To MSHFlexGrid10.Rows - 2
        TmpRs1.MoveFirst
        TmpRs1.Find "GodownID = " & MSHFlexGrid10.TextMatrix(i, 7)
        If Not TmpRs1.EOF Then
            MSHFlexGrid10.TextMatrix(i, 8) = Val(MSHFlexGrid10.TextMatrix(i, 8)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)
            MSHFlexGrid10.TextMatrix(0, 8) = Val(MSHFlexGrid10.TextMatrix(0, 8)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)
        End If
    Next
End If

tmp = " select TLAD.GodownID, MG.CLoseDate, TLA.StartDate , TLA.EXpiryDate, sum(TLAD.RentAmount) Amount  from Txn_GodownLessorAgreeDetail TLAD"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TLAD.GodownID"
tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TLA.LAID = TLAD.LAID"
tmp = tmp & " Where TLA.ExpiryDate >= '" & mStartDate & "' and TLA.StartDate <= '" & mEndDate & "'  and  TLA.FLag in ('A', 'P')"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " And TLA.LAID not in"
tmp = tmp & " (Select LAID from Txn_GodownRentDetail TGRD"
tmp = tmp & " Inner Join Txn_GodownRent TGR on TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TGRD.GodownID"
tmp = tmp & " Where TGR.Flag in ('A', 'P') And"
tmp = tmp & " TGR.ExpenseFrom = '" & mStartDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'  )"
tmp = tmp & " Group by TLAD.GodownID, TLA.EXpiryDate, TLA.StartDate , MG.CLoseDate "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
    For i = 2 To MSHFlexGrid10.Rows - 2
        TmpRs1.MoveFirst
        TmpRs1.Find "GodownID = " & MSHFlexGrid10.TextMatrix(i, 7)
        If Not TmpRs1.EOF Then
            If IsNull(TmpRs1!CloseDate) = True And Format(TmpRs1!StartDate, "yyyy-mm-dd") <= Format(mStartDate, "yyyy-mm-dd") Then
                MSHFlexGrid10.TextMatrix(i, 8) = Val(MSHFlexGrid10.TextMatrix(i, 8)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                MSHFlexGrid10.TextMatrix(0, 8) = Val(MSHFlexGrid10.TextMatrix(0, 8)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
            Else
                If Format(TmpRs1!ExpiryDate, "yyyy-mm-dd") > Format(TmpRs1!CloseDate, "yyyy-mm-dd") Then
                    If Format(TmpRs1!ExpiryDate, "yyyy-mm-dd") >= Format(mEndDate, "yyyy-mm-dd") Then
                        If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mStartDate, "yyyy-mm-dd") Then
                            mdays = DateDiff("d", TmpRs1!StartDate, CDate(mEndDate)) + 1
                            MSHFlexGrid10.TextMatrix(i, 8) = Val(MSHFlexGrid10.TextMatrix(i, 8)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                            MSHFlexGrid10.TextMatrix(0, 8) = Val(MSHFlexGrid10.TextMatrix(0, 8)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                        Else
                            MSHFlexGrid10.TextMatrix(i, 8) = Val(MSHFlexGrid10.TextMatrix(i, 8)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                            MSHFlexGrid10.TextMatrix(0, 8) = Val(MSHFlexGrid10.TextMatrix(0, 8)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                        End If
                    Else
                        If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mStartDate, "yyyy-mm-dd") Then
                            mdays = DateDiff("d", TmpRs1!StartDate, TmpRs1!ExpiryDate) + 1
                            MSHFlexGrid10.TextMatrix(i, 8) = Val(MSHFlexGrid10.TextMatrix(i, 8)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                            MSHFlexGrid10.TextMatrix(0, 8) = Val(MSHFlexGrid10.TextMatrix(0, 8)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                        Else
                            mdays = DateDiff("d", CDate(mStartDate), TmpRs1!ExpiryDate) + 1
                            MSHFlexGrid10.TextMatrix(i, 8) = Val(MSHFlexGrid10.TextMatrix(i, 8)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                            MSHFlexGrid10.TextMatrix(0, 8) = Val(MSHFlexGrid10.TextMatrix(0, 8)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                        End If
                    End If
                Else
                    If Format(TmpRs1!CloseDate, "yyyy-mm-dd") >= Format(mEndDate, "yyyy-mm-dd") Then
                        If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mStartDate, "yyyy-mm-dd") Then
                            mdays = DateDiff("d", TmpRs1!StartDate, CDate(mEndDate)) + 1
                            MSHFlexGrid10.TextMatrix(i, 8) = Val(MSHFlexGrid10.TextMatrix(i, 8)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                            MSHFlexGrid10.TextMatrix(0, 8) = Val(MSHFlexGrid10.TextMatrix(0, 8)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                        Else
                            MSHFlexGrid10.TextMatrix(i, 8) = Val(MSHFlexGrid10.TextMatrix(i, 8)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                            MSHFlexGrid10.TextMatrix(0, 8) = Val(MSHFlexGrid10.TextMatrix(0, 8)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                        End If
                    Else
                        If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mStartDate, "yyyy-mm-dd") Then
                            mdays = DateDiff("d", TmpRs1!StartDate, TmpRs1!CloseDate) + 1
                            MSHFlexGrid10.TextMatrix(i, 8) = Val(MSHFlexGrid10.TextMatrix(i, 8)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                            MSHFlexGrid10.TextMatrix(0, 8) = Val(MSHFlexGrid10.TextMatrix(0, 8)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                        Else
                            mdays = DateDiff("d", CDate(mStartDate), TmpRs1!CloseDate) + 1
                            MSHFlexGrid10.TextMatrix(i, 8) = Val(MSHFlexGrid10.TextMatrix(i, 8)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                            MSHFlexGrid10.TextMatrix(0, 8) = Val(MSHFlexGrid10.TextMatrix(0, 8)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                        End If
                    End If
                End If
            End If
        End If
    Next
End If
End Sub

Private Sub GetSecurityExpenseDetail()
tmp = " Select MS.StateName, ML.LocationName, TSBD.BillTxnID, TSB.MonthYear, MSA.AgencyName, MSS.SecurityShifts,"
tmp = tmp & " MSD.SecurityDesignation, NoofGodown, RatePerMonth, MST.Description, TSBD.BillDetailAmt IncomeAmt, 'B' Flag"
tmp = tmp & " From Txn_SecurityBilling TSB"
tmp = tmp & " Inner Join Txn_SecurityBillingDetail TSBD on TSBD.BillTxnID = TSB.BillTxnID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TSBD.LocationID And TSBD.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
tmp = tmp & " Inner Join MSt_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Inner Join Mst_ServiceTax MST on MST.ServiceTaxID =  TSB.STaxID"
tmp = tmp & " Inner Join Mst_SecurityAgency MSA on MSA.AgencyID = TSB.SecurityAgencyID"
tmp = tmp & " Inner Join Mst_SecurityDesignation MSD on MSD.SecurityDesignationID = TSBD.SecurityDesignationID"
tmp = tmp & " Inner Join Mst_SecurityShifts MSS on MSS.SecurityShiftsID = TSBD.SecurityTypeID"
If Wc <> "" Then
    tmp = tmp & Wc & " And TSB.Status in ('A', 'P') And TSB.MonthYear Between  '" & mStartDate & "' And '" & mEndDate & "'"
Else
    tmp = tmp & " Where TSB.Status in ('A', 'P') And TSB.MonthYear Between  '" & mStartDate & "' And '" & mEndDate & "'"
End If
tmp = tmp & " Union All"
tmp = tmp & " select MS.StateName, ML.LocationName, MSLD.TxnDetailID, '" & mStartDate & "' MonthYear, MSA.AgencyName, MSS.SecurityShifts,"
tmp = tmp & " MSD.SecurityDesignation, MSLD.NoofGodown, MSLD.RatePerMonth, '' Description, MSLD.RatePerMonth incomeamt, 'UB' Flag"
tmp = tmp & " from Mst_SeclocMappingDetail  MSLD"
tmp = tmp & " Inner Join Mst_SecLocMapping MSL on MSL.TxnID = MSLD.TxnID"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = MSL.LocationID And MSL.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
tmp = tmp & " Inner Join MSt_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Inner Join Mst_SecurityAgency MSA on MSA.AgencyID = MSL.SecurityAgencyID"
tmp = tmp & " Inner Join Mst_SecurityDesignation MSD on MSD.SecurityDesignationID = MSLD.SecurityDesignationID"
tmp = tmp & " Inner Join Mst_SecurityShifts MSS on MSS.SecurityShiftsID = MSLD.SecurityTypeID"
If Wc <> "" Then
    tmp = tmp & Wc & " and "
Else
   tmp = tmp & " Where "
End If
tmp = tmp & " '" & mStartDate & "' Between MSLD.StartDate And MSLD.EndDate"
tmp = tmp & " And MSLD.TxnDetailID not in (Select SecLocMapDetailID from Txn_MonthwiseSecLocDetail  TMSLD"
tmp = tmp & " Inner Join Txn_SecurityBillingDetail TSBD on TSBD.MonthwiseTxnDetailID = TMSLD.MonthwiseTxnDetailID"
tmp = tmp & " Inner Join Txn_SecurityBilling TSB on TSB.BillTxnID = TSBD.BillTxnID"
tmp = tmp & " Where TSB.Status in ('A', 'P', 'R') And TSB.MonthYear Between  '" & mStartDate & "' And '" & mEndDate & "')"
tmp = tmp & " And MSLD.TxnDetailID not in (select txndetailid from Txn_SecurityAttendance "
tmp = tmp & " where AttendanceDate between '" & mStartDate & "' And '" & mEndDate & "' And Status in ('L','R'))"
tmp = tmp & " Union All"
tmp = tmp & " select  MS.StateName, ML.LocationName, MSLD.TxnDetailID, '" & mStartDate & "' MonthYear, MSA.AgencyName, MSS.SecurityShifts,"
tmp = tmp & " MSD.SecurityDesignation, MSLD.NoofGodown, MSLD.RatePerMonth, '' Description, ((msld.ratepermonth * count(attendanceid)) / datediff(d, '" & mStartDate & "','" & mEndDate & "')+1) Incomeamt, 'UB' Flag"
tmp = tmp & " from mst_seclocmappingdetail msld"
tmp = tmp & " inner join txn_securityattendance tsa on msld.txndetailid = tsa.txndetailid"
tmp = tmp & " inner join mst_seclocmapping msl on msld.txnid = msl.txnid"
tmp = tmp & " Inner Join Mst_Location ML On ML.LocationID = MSL.LocationID And MSL.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
tmp = tmp & " Inner Join MSt_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Inner Join Mst_SecurityAgency MSA on MSA.AgencyID = MSL.SecurityAgencyID"
tmp = tmp & " Inner Join Mst_SecurityDesignation MSD on MSD.SecurityDesignationID = MSLD.SecurityDesignationID"
tmp = tmp & " Inner Join Mst_SecurityShifts MSS on MSS.SecurityShiftsID = MSLD.SecurityTypeID"
tmp = tmp & " where tsa.status  = 'A' and tsa.attendancedate between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " and msld.txndetailid not in (select isnull(msld.SecLocMapDetailID,0) from Txn_MonthwiseSecLocDetail msld"
tmp = tmp & " inner join Txn_MonthwiseSecLoc msl on msl.monthwisetxnid = msld.monthwisetxnid"
tmp = tmp & " where monthyear Between  '" & mStartDate & "' And '" & mEndDate & "')"
tmp = tmp & " group by MS.StateName, ML.LocationName, MSLD.TxnDetailID, MSA.AgencyName, MSS.SecurityShifts, MSD.SecurityDesignation, MSLD.NoofGodown, MSLD.RatePerMonth"

Call Tmp1Table

Call Grid_Head9
ProgressBar1.Value = 0

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid9.Rows = MSHFlexGrid9.Rows + TmpRs1.RecordCount
   ProgressBar1.Max = MSHFlexGrid9.Rows
   With MSHFlexGrid9
      For i = 2 To TmpRs1.RecordCount + 1
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!BillTxnID), "", TmpRs1!BillTxnID)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!MonthYear), "", Format(TmpRs1!MonthYear, "dd-MMM-yyyy"))
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs1!AgencyName), "", TmpRs1!AgencyName)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs1!SecurityShifts), "", TmpRs1!SecurityShifts)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs1!SecurityDesignation), "", TmpRs1!SecurityDesignation)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs1!NoofGodown), "", TmpRs1!NoofGodown)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs1!RatePerMonth), "", TmpRs1!RatePerMonth)
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs1!Description), "", TmpRs1!Description)
         .TextMatrix(i, 10) = IIf(IsNull(TmpRs1!IncomeAmt), "", TmpRs1!IncomeAmt)
         .TextMatrix(0, 10) = Val(.TextMatrix(0, 10)) + Val(.TextMatrix(i, 10))
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs1!Flag), "", TmpRs1!Flag)
         TmpRs1.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
End If
End Sub

Private Sub GetVendorExpenseDetail()
tmp = " Select MS.StateName, ML.LocationName, MIET.ExpenseType, TVCD.Particulars, TVCD.Amount as IncomeAmt, case TVC.Flag when 'A' then 'Approved' when 'P' then 'Pending' end as Flag"
tmp = tmp & " From Txn_VendorClaimDetail TVCD"
tmp = tmp & " Inner Join Txn_VendorClaim TVC on TVC.ClaimID = TVCD.ClaimID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TVCD.LocationID And TVCD.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
tmp = tmp & " Left Join Mst_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Inner Join Mst_ExpenseType MIET on MIET.ExpenseTypeID = TVCD.ExpenseTypeID"
If Wc <> "" Then
    tmp = tmp & Wc & " And TVC.Flag in ('A', 'P') And  TVCD.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
Else
    tmp = tmp & " Where TVC.Flag in ('A', 'P') And TVCD.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
End If
tmp = tmp & " And TVCD.ExpenseTypeID not in (17, 18, 8, 69)" ''8 and  69 means Franchisee expese type and 18 means Fumigation and 17 means Spraying Charges
tmp = tmp & " And isnull(MIET.MISExpenseGroup, '') <> 'NA'"
Call Tmp1Table

Call Grid_Head8
ProgressBar1.Value = 0

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid8.Rows = MSHFlexGrid8.Rows + TmpRs1.RecordCount
   ProgressBar1.Max = MSHFlexGrid8.Rows
   With MSHFlexGrid8
      For i = 2 To TmpRs1.RecordCount + 1
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!ExpenseType), "", TmpRs1!ExpenseType)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!Particulars), "", TmpRs1!Particulars)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs1!IncomeAmt), "", TmpRs1!IncomeAmt)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs1!Flag), "", TmpRs1!Flag)
         .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) + Val(.TextMatrix(i, 4))
         TmpRs1.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
End If

End Sub

Private Sub GetPledgeIncomeDetail()
tmp = " Select TPCR.PCRID, MS.StateName, ML.LocationName, TPCR.InstrumentType, TPCR.DrawnOn, TPCR.BranchName, TPCR.InstrumentNo, TPCR.InstrumentDate,"
tmp = tmp & " TPCR.receivedFrom, TPCR.ExcessAmount, TPCR.Remark, TPCR.InstrumentAmount IncomeAmt"
tmp = tmp & " from Txn_PledgeCashReceipt TPCR"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TPCR.LocationID"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
If Wc <> "" Then
    tmp = tmp & Wc & " And TPCR.InstrumentDate  Between  '" & mStartDate & "' And '" & mEndDate & "'"
Else
    tmp = tmp & " Where TPCR.InstrumentDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
End If

Call Tmp1Table

Call Grid_Head7
ProgressBar1.Value = 0
If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid7.Rows = MSHFlexGrid7.Rows + TmpRs1.RecordCount
   ProgressBar1.Max = MSHFlexGrid7.Rows
   With MSHFlexGrid7
      For i = 2 To TmpRs1.RecordCount + 1
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!PCRID), "", TmpRs1!PCRID)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!InstrumentType), "", TmpRs1!InstrumentType)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs1!DrawnOn), "", TmpRs1!DrawnOn)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs1!BranchName), "", TmpRs1!BranchName)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs1!InstrumentNo), "", TmpRs1!InstrumentNo)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs1!InstrumentDate), "", Format(TmpRs1!InstrumentDate, "dd-MMM-yyyy"))
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs1!ReceivedFrom), "", TmpRs1!ReceivedFrom)
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs1!ExcessAmount), "", TmpRs1!ExcessAmount)
         .TextMatrix(0, 9) = Val(.TextMatrix(0, 9)) + Val(.TextMatrix(i, 9))
         .TextMatrix(i, 10) = IIf(IsNull(TmpRs1!Remark), "", TmpRs1!Remark)
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs1!IncomeAmt), "", TmpRs1!IncomeAmt)
         .TextMatrix(0, 11) = Val(.TextMatrix(0, 11)) + Val(.TextMatrix(i, 11))
         TmpRs1.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
End If
End Sub

Private Sub GetDPIncomeDetail()

tmp = " Select MS.StateName, ML.LocationName, MDP.DPInvoiceID, MDP.PINo, MDP.MonthYear, MDP.DPNAvcode, MDP.Amt" ', MDP.STAmt,"
'tmp = tmp & " MDP.TotalAmt IncomeAmt "
tmp = tmp & " from Mst_DPInvoiceUpload MDP"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MDP.LocationID And MDP.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
tmp = tmp & " Inner Join Mst_State MS  on MS.StateID = ML.StateID"
If Wc <> "" Then
    tmp = tmp & Wc & " And isnull(MDP.ReverseFlag,'')<>'Y' And  MDP.MonthYear Between  '" & mStartDate & "' And '" & mEndDate & "'"
Else
    tmp = tmp & " Where isnull(MDP.ReverseFlag,'')<>'Y' And  MDP.MonthYear Between  '" & mStartDate & "' And '" & mEndDate & "'"
End If
tmp = tmp & " Order By MS.StateName, ML.LocationName"

Call Tmp1Table

Call Grid_Head6
ProgressBar1.Value = 0

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + TmpRs1.RecordCount
   ProgressBar1.Max = MSHFlexGrid6.Rows
   With MSHFlexGrid6
      For i = 2 To TmpRs1.RecordCount + 1
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!DPInvoiceID), "", TmpRs1!DPInvoiceID)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!PINo), "", TmpRs1!PINo)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs1!MonthYear), "", Format(TmpRs1!MonthYear, "dd-MMM-yyyy"))
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs1!DPNAVCode), "", TmpRs1!DPNAVCode)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs1!Amt), "", TmpRs1!Amt)
         .TextMatrix(0, 6) = Val(.TextMatrix(0, 6)) + Val(.TextMatrix(i, 6))
         '.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!STAmt), "", TmpRs1!STAmt)
         '.TextMatrix(0, 7) = Val(.TextMatrix(0, 7)) + Val(.TextMatrix(I, 7))
         '.TextMatrix(I, 8) = IIf(IsNull(TmpRs1!IncomeAmt), "", TmpRs1!IncomeAmt)
         '.TextMatrix(0, 8) = Val(.TextMatrix(0, 8)) + Val(.TextMatrix(I, 8))
         TmpRs1.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
End If
End Sub

Private Sub GetSevenDaysIncomeDetail()
tmp = " Select MS.StateName, ML.LocationName, TI.PINo, TI.PIDate, MIET.IEType, MET.ExchangeType, MCom.Commodity,"
tmp = tmp & " MCL.ClientName, TI.InvoiceDate, TI.Amount IncomeAmt"
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " Inner Join Mst_IncomeExpenseType MIET on MIET.IETypeID = TI.InvoiceTypeID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TI.LocationID  And TI.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Left Join Mst_ExchangeType MET on MET.ExchangeTypeID = TI.ExchangeTypeID"
tmp = tmp & " Left Join Mst_Commodity MCom on MCom.CommodityID = TI.CommodityID"
tmp = tmp & " Left Join Mst_Client MCL on MCL.ClientIDRow = TI.ClientIDRow"
If Wc <> "" Then
    tmp = tmp & Wc & " And "
Else
    tmp = tmp & " Where "
End If
tmp = tmp & " TI.Flag in ('A', 'P') And isnuLL(TI.ReverseFlag,'') <> 'Y' And TI.InvoiceTypeID in (1,2,3,4,5,6,7,8,11) "
tmp = tmp & " And TI.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Order By MS.StateName, ML.LocationName, MIET.IEType "

Call Tmp1Table

Call Grid_Head5
ProgressBar1.Value = 0

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + TmpRs1.RecordCount
   ProgressBar1.Max = MSHFlexGrid5.Rows
   With MSHFlexGrid5
      For i = 2 To TmpRs1.RecordCount + 1
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!ExchangeType), "", TmpRs1!ExchangeType)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!Commodity), "", TmpRs1!Commodity)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs1!IEType), "", TmpRs1!IEType)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs1!PINo), "", TmpRs1!PINo)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs1!PIDate), "", Format(TmpRs1!PIDate, "dd-MMM-yyyy"))
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs1!InvoiceDate), "", Format(TmpRs1!InvoiceDate, "dd-MMM-yyyy"))
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs1!IncomeAmt), "", TmpRs1!IncomeAmt)
         .TextMatrix(0, 9) = Val(.TextMatrix(0, 9)) + Val(.TextMatrix(i, 9))
         TmpRs1.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
End If
End Sub

Private Sub GetReservationIncomeDetail()
tmp = " Select MS.StateName, ML.LocationName, "
tmp = tmp & " TI.InvoiceNo, TI.BillingDate, TI.ActualAmount IncomeAmt"
tmp = tmp & " from Txn_ReservationInvoice TI"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TI.LocationID And TI.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
If Wc <> "" Then
    tmp = tmp & Wc & " And "
Else
    tmp = tmp & " Where "
End If
tmp = tmp & " isnull(TI.ReverseFlag, '') <> 'Y' And  TI.Flag in ('A', 'P') And TI.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
'tmp = tmp & " isnull(TI.ReverseFlag, '') <> 'Y' And TI.Flag = 'A' And TI.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " Order By MS.StateName, ML.LocationName"
Call Tmp1Table

Call Grid_Head4
ProgressBar1.Value = 0

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + TmpRs1.RecordCount
   ProgressBar1.Max = MSHFlexGrid4.Rows
   With MSHFlexGrid4
      For i = 2 To TmpRs1.RecordCount + 1
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!InvoiceNo), "", TmpRs1!InvoiceNo)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!BillingDate), "", Format(TmpRs1!BillingDate, "dd-MMM-yyyy"))
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs1!IncomeAmt), "", TmpRs1!IncomeAmt)
         .TextMatrix(0, 4) = Val(.TextMatrix(0, 4)) + Val(.TextMatrix(i, 4))
         TmpRs1.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
       Next
   End With
End If
End Sub

Private Sub GetFranchiseeFeesDetail()
tmp = " Select MS.StateName, ML.LOcationName, MCMP.CMPName, MG.GodownNo, MG.Address, MET.ExchangeType, TIMG.BalanceBags, TIMG.BalanceWeight, TIMG.FranchiseeRate, MB.BillingCycle, TIMG.FranchBillingUnit, TIMG.FranchNoofDays, TIMG.GSLID, TIMG.StackNo, TIMG.LotNo, "
tmp = tmp & " isnull(TFR.ActualFee * TIMG.FranchInvoiceAmount/TFR.Amount, TIMG.FranchInvoiceAmount) FranchInvoiceAmount, TFR.Flag ClaimFlag From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Inner Join MSt_LOcation ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " Inner Join MSt_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Left JOin Mst_ExchangeType MET on MET.ExchangeTypeID = TIMG.ExchangeTypeID"
tmp = tmp & " Inner Join Txn_FranchiseeRent TFR on TFR.ClaimID = TIMG.FranchClaimID and TFR.flag in ('A', 'P')"
tmp = tmp & " Left Join Mst_BillingCycle MB on TIMG.FranchBillingCycleID = MB.BillingCycleID"
tmp = tmp & " Where TIMG.FranchiseeRate is not null And TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And TIMG.FranchInvoiceAmount > 0 And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "')"
tmp = tmp & " And MG.StartDate <= '" & mEndDate & "' and isnull(TIMG.FranchClaimID, 0) <> 0"
tmp = tmp & " Union All"
tmp = tmp & " Select MS.StateName, ML.LOcationName, MCMP.CMPName, MG.GodownNo, MG.Address, MET.ExchangeType, TIMG.BalanceBags, TIMG.BalanceWeight, TIMG.FranchiseeRate, MB.BillingCycle, TIMG.FranchBillingUnit, TIMG.FranchNoofDays, TIMG.GSLID, TIMG.StackNo, TIMG.LotNo,"
tmp = tmp & " TIMG.FranchInvoiceAmount, 'UB' ClaimFlag From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Inner Join MSt_LOcation ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " Inner Join MSt_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Left JOin Mst_ExchangeType MET on MET.ExchangeTypeID = TIMG.ExchangeTypeID"
tmp = tmp & " Left Join Mst_BillingCycle MB on TIMG.FranchBillingCycleID = MB.BillingCycleID"
tmp = tmp & " Where TIMG.FranchiseeRate is not null And TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And TIMG.FranchInvoiceAmount > 0 And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "')"
tmp = tmp & " And MG.StartDate <= '" & mEndDate & "' and isnull(TIMG.FranchClaimID, 0) = 0"
Call Tmp1Table

Call Grid_Head12
ProgressBar1.Value = 0

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid12.Rows = MSHFlexGrid12.Rows + TmpRs1.RecordCount
   ProgressBar1.Max = MSHFlexGrid12.Rows
   With MSHFlexGrid12
      For i = 2 To TmpRs1.RecordCount + 1
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!CMPName), "", TmpRs1!CMPName)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!GodownNo), "", TmpRs1!GodownNo)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs1!Address), "", TmpRs1!Address)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs1!ExchangeType), "", TmpRs1!ExchangeType)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs1!BalanceBags), "", TmpRs1!BalanceBags)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs1!BalanceWeight), "", TmpRs1!BalanceWeight)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs1!FranchiseeRate), "", TmpRs1!FranchiseeRate)
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs1!BillingCycle), "", TmpRs1!BillingCycle)
         .TextMatrix(i, 10) = IIf(IsNull(TmpRs1!FranchBillingUnit), "", TmpRs1!FranchBillingUnit)
         If LCase(.TextMatrix(i, 10)) = "p" Then
             .TextMatrix(i, 10) = "Percentage"
         ElseIf LCase(.TextMatrix(i, 10)) = "b" Then
             .TextMatrix(i, 10) = "Bags"
         ElseIf LCase(.TextMatrix(i, 10)) = "q" Then
             .TextMatrix(i, 10) = "Quantity"
         ElseIf LCase(.TextMatrix(i, 10)) = "f" Then
             .TextMatrix(i, 10) = "Fixed"
         End If
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs1!FranchNoofDays), "", TmpRs1!FranchNoofDays)
         .TextMatrix(i, 12) = IIf(IsNull(TmpRs1!GSLID), "", TmpRs1!GSLID)
         .TextMatrix(i, 13) = IIf(IsNull(TmpRs1!StackNo), "", TmpRs1!StackNo)
         .TextMatrix(i, 14) = IIf(IsNull(TmpRs1!LotNo), "", TmpRs1!LotNo)
         .TextMatrix(i, 15) = IIf(IsNull(TmpRs1!FranchInvoiceAmount), "", TmpRs1!FranchInvoiceAmount)
         .TextMatrix(0, 15) = Val(.TextMatrix(0, 15)) + Val(.TextMatrix(i, 15))
         .TextMatrix(i, 16) = IIf(IsNull(TmpRs1!ClaimFlag), "", TmpRs1!ClaimFlag)
         TmpRs1.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
End If

End Sub


Private Sub GetNNNRIncomeDetail()

tmp = " Select MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo, MG.Address, TIG.StackNo, TIG.LotNo, MCom.Commodity, MCL.ClientName,"
tmp = tmp & " TIG.InvoiceAmount IncomeAmt"
tmp = tmp & " from Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIG.GodownID"
tmp = tmp & " Inner Join Txn_ProInvoice TPI on TPI.InvoiceNo = TIG.PINo And isnull(TPI.ReverseFlag, '') <> 'Y'"
tmp = tmp & " Inner Join Mst_Commodity MCom on MCom.CommodityID = TIG.CommodityID"
tmp = tmp & " Inner Join Mst_Client MCL on MCL.ClientIDRow = TIG.ClientIDRow"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = TIG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.SM_Prefix = TIG.SM_prefix"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
If Wc <> "" Then
   tmp = tmp & Wc & " And TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null) And PINo Is Not Null "
Else
   tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null) And PINo Is Not Null "
End If
tmp = tmp & " And TIG.BalanceWeight > 0 And TIG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Order by MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo"

Call Tmp1Table

Call Grid_Head3
ProgressBar1.Value = 0

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + TmpRs1.RecordCount
   ProgressBar1.Max = MSHFlexGrid3.Rows
   With MSHFlexGrid3
        For i = 2 To TmpRs1.RecordCount + 1
            .TextMatrix(i, 0) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
            .TextMatrix(i, 1) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
            .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!CMPName), "", TmpRs1!CMPName)
            .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!GodownNo), "", TmpRs1!GodownNo)
            .TextMatrix(i, 4) = IIf(IsNull(TmpRs1!Address), "", TmpRs1!Address)
            .TextMatrix(i, 5) = IIf(IsNull(TmpRs1!StackNo), "", TmpRs1!StackNo)
            .TextMatrix(i, 6) = IIf(IsNull(TmpRs1!LotNo), "", TmpRs1!LotNo)
            .TextMatrix(i, 7) = IIf(IsNull(TmpRs1!Commodity), "", TmpRs1!Commodity)
            .TextMatrix(i, 8) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
            .TextMatrix(i, 9) = IIf(IsNull(TmpRs1!IncomeAmt), "", TmpRs1!IncomeAmt)
            .TextMatrix(0, 9) = Val(.TextMatrix(0, 9)) + Val(.TextMatrix(i, 9))
            TmpRs1.MoveNext
            ProgressBar1.Value = ProgressBar1.Value + 1
        Next
   End With
End If

End Sub

Private Sub GetStockDetail()
tmp = " Select MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo, MG.Address, MET.ExchangeType, TIMG.StackNo, TIMG.LotNo, MCL.ClientName, "
tmp = tmp & " TIMG.BalanceWeight"
tmp = tmp & " From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_ExchangeType MET on MET.ExchangeTypeID = TIMG.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = TIMG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.SM_Prefix = TIMG.SM_Prefix"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Inner Join Mst_Client MCL on MCL.ClientIDRow = TIMG.ClientIDRow"
If Wc <> "" Then
    tmp = tmp & Wc & " And "
Else
    tmp = tmp & " Where "
End If
tmp = tmp & " TIMG.ExchangeTypeID in (1, 2) And TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And TIMG.BalanceWeight > 0 And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Order by MET.ExchangeType, MS.StateName, ML.LocationName, MCMP.CMPName, MG.GodownNo"
Call Tmp1Table

Call Grid_Head2
ProgressBar1.Value = 0

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs1.RecordCount
   ProgressBar1.Max = MSHFlexGrid2.Rows
   With MSHFlexGrid2
      For i = 2 To TmpRs1.RecordCount + 1
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs1!StateName), "", TmpRs1!StateName)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
         .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!CMPName), "", TmpRs1!CMPName)
         .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!GodownNo), "", TmpRs1!GodownNo)
         .TextMatrix(i, 4) = IIf(IsNull(TmpRs1!Address), "", TmpRs1!Address)
         .TextMatrix(i, 5) = IIf(IsNull(TmpRs1!StackNo), "", TmpRs1!StackNo)
         .TextMatrix(i, 6) = IIf(IsNull(TmpRs1!LotNo), "", TmpRs1!LotNo)
         .TextMatrix(i, 7) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
         .TextMatrix(i, 8) = IIf(IsNull(TmpRs1!ExchangeType), "", TmpRs1!ExchangeType)
         .TextMatrix(i, 9) = IIf(IsNull(TmpRs1!BalanceWeight), "", TmpRs1!BalanceWeight)
         .TextMatrix(0, 9) = Val(.TextMatrix(0, 9)) + Val(.TextMatrix(i, 9))
         TmpRs1.MoveNext
         ProgressBar1.Value = ProgressBar1.Value + 1
      Next
   End With
End If
End Sub

Private Sub GetExpenseEmployeeClaim()
tmp = " Select TCSD.LocationID, Sum(TCSD.Amount) Amount"
tmp = tmp & " From Txn_ClaimSegmentDistribution TCSD"
tmp = tmp & " Inner Join Txn_ClaimSubmission TCS on TCS.ClaimID = TCSD.ClaimID"
tmp = tmp & " Where TCSD.SegmentID = 1 And TCS.Flag = 'S'"
tmp = tmp & " And '" & mStartDate & "' Between TCS.ExpenseFrom And TCS.ExpenseTo"
tmp = tmp & " group by TCSD.LocationID"
Call TmpTableHRPro

If TmpRsHRPro.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRsHRPro.MoveFirst
      TmpRsHRPro.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRsHRPro.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.EmpClaimFAG) = IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.EmpClaimFAG) = IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.EmpClaimFAG) = Val(MSHFlexGrid1.TextMatrix(i, E.EmpClaimFAG)) + Val(MSHFlexGrid1.TextMatrix(0, E.EmpClaimFAG))
   Next
End If

'''''''''''''Pending Claim================================

tmp = " Select TCSD.LocationID, Sum(TCSD.Amount) Amount"
tmp = tmp & " From Txn_ClaimSegmentDistribution TCSD"
tmp = tmp & " Inner Join Txn_ClaimSubmission TCS on TCS.ClaimID = TCSD.ClaimID"
tmp = tmp & " Where TCSD.SegmentID = 1 And TCS.Flag not in ('R', 'S')"
tmp = tmp & " And '" & mStartDate & "' Between TCS.ExpenseFrom And TCS.ExpenseTo"
tmp = tmp & " group by TCSD.LocationID"
'tmp = tmp & " And TCSD.LocationID = " & TmpRs!LocationID
Call TmpTableHRPro

If TmpRsHRPro.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRsHRPro.MoveFirst
      TmpRsHRPro.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRsHRPro.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.EmpClaimAccrual) = IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.EmpClaimAccrual) = IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.EmpClaimAccrual) = Val(MSHFlexGrid1.TextMatrix(i, E.EmpClaimAccrual)) + Val(MSHFlexGrid1.TextMatrix(0, E.EmpClaimAccrual))
   Next
End If
End Sub

Private Sub GetAMRMName()
tmp = " Select ECM.LocationID, EMP2.EmployeeName+'~'+EMP2.EmployeeNo  as 'AMName'"
tmp = tmp & " , EMP3.EmployeeName +'~'+EMP3.EmployeeNo as 'RMName'"
tmp = tmp & " From Mst_EMPCMPMapping ECM"
tmp = tmp & " Right Outer Join Mst_Employee EMP On ECM.EmployeeID = EMP.EMployeeID And EMP.Flag = 'Y'"
tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP.DirectReportingTo = EMP2.EmployeeID And EMP2.Flag = 'Y'"
tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID And EMP3.Flag = 'Y'"
tmp = tmp & " Where EMP2.EmployeeName Is Not Null"
tmp = tmp & " Group By ECM.LocationID, EMP2.EmployeeName, EMP2.EmployeeNo"
tmp = tmp & " ,EMP3.EmployeeName ,EMP3.EmployeeNo"
Call Tmp1Table

For x_ = 2 To MSHFlexGrid1.Rows - 2
   TmpRs1.Filter = "LocationID = " & MSHFlexGrid1.TextMatrix(x_, E.LocationID)
   If TmpRs1.RecordCount > 0 Then
      For j = 1 To TmpRs1.RecordCount
         If MSHFlexGrid1.TextMatrix(x_, E.AMname) = "" Then
            MSHFlexGrid1.TextMatrix(x_, E.AMname) = IIf(IsNull(TmpRs1!AMname), "", TmpRs1!AMname)
         Else
            If InStr(1, MSHFlexGrid1.TextMatrix(x_, E.AMname), IIf(IsNull(TmpRs1!AMname), "", TmpRs1!AMname)) <= 0 Then
               MSHFlexGrid1.TextMatrix(x_, E.AMname) = IIf(IsNull(TmpRs1!AMname), "", TmpRs1!AMname) & " , " & MSHFlexGrid1.TextMatrix(x_, E.AMname)
            End If
         End If
         If MSHFlexGrid1.TextMatrix(x_, E.RMName) = "" Then
            MSHFlexGrid1.TextMatrix(x_, E.RMName) = IIf(IsNull(TmpRs1!RMName), "", TmpRs1!RMName)
         Else
            If InStr(1, MSHFlexGrid1.TextMatrix(x_, E.RMName), IIf(IsNull(TmpRs1!RMName), "", TmpRs1!RMName)) <= 0 Then
               MSHFlexGrid1.TextMatrix(x_, E.RMName) = IIf(IsNull(TmpRs1!RMName), "", TmpRs1!RMName) & " , " & MSHFlexGrid1.TextMatrix(x_, E.RMName)
            End If
         End If
         TmpRs1.MoveNext
      Next j
   End If
   TmpRs1.Filter = ""
Next
End Sub

Private Sub GetNCDEXStock()

tmp = " Select TIMG.GodownID, sum(TIMG.BalanceWeight) BalanceWeight"
tmp = tmp & " From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Where TIMG.ExchangeTypeID = 2 And TIMG.BalanceWeight > 0 And TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group by TIMG.GodownID"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.NcdexStock) = IIf(IsNull(TmpRs1!BalanceWeight), 0, TmpRs1!BalanceWeight)
      End If
      MSHFlexGrid1.TextMatrix(0, E.NcdexStock) = Val(MSHFlexGrid1.TextMatrix(0, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock))
   Next
End If
End Sub

Private Sub GetNonNCDEXStock()
tmp = " Select TIMG.GodownID, sum(TIMG.BalanceWeight) BalanceWeight"
tmp = tmp & " From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Where TIMG.ExchangeTypeID = 1 And TIMG.BalanceWeight > 0 And TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group by TIMG.GodownID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock) = IIf(IsNull(TmpRs1!BalanceWeight), 0, TmpRs1!BalanceWeight)
      End If
      MSHFlexGrid1.TextMatrix(0, E.NonNCDEXStock) = Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock)) + Val(MSHFlexGrid1.TextMatrix(0, E.NonNCDEXStock))
   Next
End If

End Sub

Private Sub GetNNNRIncome()

tmp = " Select TIG.GodownID, Sum(TIG.InvoiceAmount) IncomeAmt"
tmp = tmp & " from Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIG.GodownID"
tmp = tmp & " Inner Join Txn_ProInvoice TPI on TPI.InvoiceNo = TIG.PINo And isnull(TPI.ReverseFlag, '') <> 'Y'"
tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null)"
tmp = tmp & " And PINo is not Null And TIG.BalanceWeight > 0"
tmp = tmp & " And TIG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group By TIG.GodownID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
       TmpRs1.MoveFirst
       TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
       If Not TmpRs1.EOF Then
          MSHFlexGrid1.TextMatrix(i, E.NNNRIncome) = IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)
       End If
       MSHFlexGrid1.TextMatrix(0, E.NNNRIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.NNNRIncome))
    Next
End If

''========================Discount And Premium

tmp = " Select TIG.PINo, TIG.GodownID, Sum(TIG.InvoiceAmount) IncomeAmt"
tmp = tmp & " from Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIG.GodownID"
tmp = tmp & " Inner Join Txn_ProInvoice TPI on TPI.InvoiceNo = TIG.PINo And isnull(TPI.ReverseFlag, '') <> 'Y'"
tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null)"
tmp = tmp & " And PINo is not Null And TIG.BalanceWeight > 0"
tmp = tmp & " And TIG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group By TIG.GodownID, TIG.PINo"
Call Tmp1Table

tmp = " select Amount, InvoiceNo, Discount, Premium from txn_proinvoice Where (Discount > 0 Or Premium > 0) "
tmp = tmp & " and InvoiceNo in ("
tmp = tmp & " Select TIG.PINo from Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIG.GodownID"
tmp = tmp & " Inner Join Txn_ProInvoice TPI on TPI.InvoiceNo = TIG.PINo And isnull(TPI.ReverseFlag, '') <> 'Y'"
tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID = 0 Or TIG.ContractID Is Null)"
tmp = tmp & " And PINo is not Null And TIG.BalanceWeight > 0"
tmp = tmp & " And TIG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " )"

Call Tmp3Table

If TmpRs1.RecordCount > 0 And TmpRs3.RecordCount > 0 Then
   For j = 1 To TmpRs1.RecordCount
      TmpRs3.MoveFirst
      TmpRs3.Find " InvoiceNo = " & TmpRs1!PINo
      If Not TmpRs3.EOF Then
         mDiscount = (TmpRs3!Discount * TmpRs1!IncomeAmt) / TmpRs3!Amount
         mPremium = (TmpRs3!Premium * TmpRs1!IncomeAmt) / TmpRs3!Amount
         For i = 2 To MSHFlexGrid1.Rows - 2
            If MSHFlexGrid1.TextMatrix(i, E.GodownID) = TmpRs1!GodownID Then
               MSHFlexGrid1.TextMatrix(i, E.NNNRIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncome)) + Format(Val(mPremium) - Val(mDiscount), "##########.00")
               MSHFlexGrid1.TextMatrix(0, E.NNNRIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.NNNRIncome)) + Format(Val(mPremium) - Val(mDiscount), "##########.00")
               Exit For
            End If
         Next
      End If
      TmpRs1.MoveNext
   Next
End If

'''''Reversal minus

tmp = " Select  TCA.LocationID,sum(RecOtherDed) RecOtherDed"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCAD"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCA on TCA.TDSAID = TCAD.TDSAID And TCA.SM_Prefix = TCAD.SM_Prefix"
tmp = tmp & " Where TCA.Flag in ('A', 'P') And AdjDate Between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And TCAD.ExpenseAccID = 88 " '88 means NNNR Reversal"
tmp = tmp & " group by TCA.LocationID"
Call Tmp4Table

mTotalNNNRIncome = MSHFlexGrid1.TextMatrix(0, E.NNNRIncome)
If Val(mTotalNNNRIncome) > 0 Then
    For i = 2 To MSHFlexGrid1.Rows - 2
        If TmpRs4.RecordCount > 0 Then
           TmpRs4.MoveFirst
           TmpRs4.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
           If Not TmpRs4.EOF Then
              mRecOtherDed = Format(Val(Val(IIf(IsNull(TmpRs4!RecOtherDed), "", TmpRs4!RecOtherDed)) * Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncome))) / Val(mTotalNNNRIncome), "########.00")
              MSHFlexGrid1.TextMatrix(i, E.NNNRIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncome)) - Val(mRecOtherDed)
              MSHFlexGrid1.TextMatrix(0, E.NNNRIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.NNNRIncome)) - Val(mRecOtherDed)
           End If
        End If
    Next
End If
End Sub

Private Sub GetReservationIncome()


'tmp = " Select TI.InvoiceNo, TI.LocationID, sum(TI.ActualAmount) As 'IncomeAmt'"
'tmp = tmp & " from Txn_ReservationInvoice TI"
'tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TI.LocationID And TI.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
'tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
'If Wc <> "" Then
'    tmp = tmp & Wc & " And "
'Else
'    tmp = tmp & " Where "
'End If
'tmp = tmp & " isnull(TI.ReverseFlag, '') <> 'Y' And TI.Flag in ('A', 'P') And TI.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
''tmp = tmp & " isnull(TI.ReverseFlag, '') <> 'Y' And TI.Flag = 'A' And TI.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
'tmp = tmp & " Group by TI.InvoiceNo, TI.LocationID " ''--TI.ContractID, TI.LocationID,
'tmp = tmp & " Order by  TI.InvoiceNo, TI.LocationID " ''--TI.LocationID,
'Call Tmp1Table
'
'If TmpRs1.RecordCount > 0 Then
'   For j = 1 To TmpRs1.RecordCount
'      tmp = " Select MGSL.GodownID, Sum(MGSL.BalanceWeight) BalanceWeight, sum(MGSL.BalanceBags) BalanceBags From Txn_INvMonGSLDetail MGSL"
'      tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MGSL.GodownID"
'      tmp = tmp & " Where MGSL.ContractID in"
'      tmp = tmp & " (Select ContractID From Txn_ReservationInvoice TI"
'      tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TI.LocationID"
'      tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
'      If Wc <> "" Then
'         tmp = tmp & Wc & " And "
'      Else
'         tmp = tmp & " Where "
'      End If
'      tmp = tmp & " isnull(TI.ReverseFlag, '') <> 'Y' And (TI.InvoiceNo = " & IIf(IsNull(TmpRs1!InvoiceNo), 0, TmpRs1!InvoiceNo) & " Or TI.PrinciplePINo = " & IIf(IsNull(TmpRs1!InvoiceNo), 0, TmpRs1!InvoiceNo) & ")"
'      'tmp = tmp & " And MGSL.BalanceWeight > 0 And MGSL.BalanceBags > 0 And TI.Flag = 'A' And TI.YearMonth Between '" & mStartDate & "'  and '" & mEndDate & "')" '')"
'      tmp = tmp & " And MGSL.BalanceWeight > 0 And MGSL.BalanceBags > 0 And TI.Flag in ('A', 'P') And TI.YearMonth Between '" & mStartDate & "'  and '" & mEndDate & "')" '')"
'      tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "' ) And MG.StartDate <= '" & mEndDate & "'"
'      tmp = tmp & " And MGSL.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
'      tmp = tmp & " Group By MGSL.GodownID"
'      Call Tmp4Table
'
'      If TmpRs4.RecordCount > 0 Then
'         mBalanceWeight = 0
'         For k = 1 To TmpRs4.RecordCount
'            mBalanceWeight = Val(mBalanceWeight) + IIf(IsNull(TmpRs4!BalanceWeight), 0, TmpRs4!BalanceWeight)
'            TmpRs4.MoveNext
'         Next
'
'         TmpRs4.MoveFirst
'         For k = 1 To TmpRs4.RecordCount
'            For i = 2 To MSHFlexGrid1.Rows - 2
'               If MSHFlexGrid1.TextMatrix(i, E.GodownID) = TmpRs4!GodownID Then
'                  mGodownWeight = IIf(IsNull(TmpRs4!BalanceWeight), 0, TmpRs4!BalanceWeight)
'                  MSHFlexGrid1.TextMatrix(i, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome)) + Format((IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt) * mGodownWeight) / mBalanceWeight, "###########.00")
'                  MSHFlexGrid1.TextMatrix(0, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) + Format((IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt) * mGodownWeight) / mBalanceWeight, "###########.00")
'                  Exit For
'               End If
'            Next
'            TmpRs4.MoveNext
'         Next
'      Else
'         For k = 2 To MSHFlexGrid1.Rows - 2
'            If MSHFlexGrid1.TextMatrix(k, E.LocationID) = TmpRs1!LocationID Then
'               TmpRs2.MoveFirst
'               TmpRs2.Find "LocationID = " & MSHFlexGrid1.TextMatrix(k, E.LocationID)
'               If Not TmpRs2.EOF Then
'                  If Opt_Stock.Value = True Then
'                     mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
'                     MSHFlexGrid1.TextMatrix(k, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(k, E.ReservationIncome)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.000"))
'                     MSHFlexGrid1.TextMatrix(0, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.000"))
'                  ElseIf Opt_Capacity.Value = True Then
'                     MSHFlexGrid1.TextMatrix(k, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(k, E.ReservationIncome)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(k, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
'                     MSHFlexGrid1.TextMatrix(0, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(k, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
'                  End If
'               End If
'            End If
'         Next
'      End If
'      TmpRs1.MoveNext
'   Next
'End If
'
'
''''''For Reversal==========================================
'tmp = " Select  sum(RecOtherDed) RecOtherDed" 'TCA.LocationID,
'tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCAD"
'tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCA on TCA.TDSAID = TCAD.TDSAID And TCA.SM_Prefix = TCAD.SM_Prefix"
'tmp = tmp & " Where TCA.Flag in ('A', 'P') And AdjDate Between '" & mStartDate & "' And '" & mEndDate & "'"
'tmp = tmp & " And TCAD.ExpenseAccID  = 90 " '90 means Reservation Reversal"
'Call Tmp5Table
'
'mTotalReservationIncome = MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)
'If Val(mTotalReservationIncome) > 0 Then
'    For i = 2 To MSHFlexGrid1.Rows - 2
'        If TmpRs5.RecordCount > 0 Then
'            'TmpRs5.MoveFirst
'            'TmpRs5.Find "LocationID = " & val( MSHFlexGrid1.TextMatrix(I, E.LocationID))
'            'If Not TmpRs5.EOF Then
'                mRecOtherDed = Format(Val(Val(IIf(IsNull(TmpRs5!RecOtherDed), "", TmpRs5!RecOtherDed)) * Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome))) / Val(mTotalReservationIncome), "########.00")
'                MSHFlexGrid1.TextMatrix(i, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome)) - Val(mRecOtherDed)
'                MSHFlexGrid1.TextMatrix(0, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) - Val(mRecOtherDed)
'            'End If
'        End If
'    Next
'End If


'-------------------------------------------------------------------------------------------------------------------------------


tmp = " Select TIG.GodownID, Sum(TIG.InvoiceAmount) IncomeAmt"
tmp = tmp & " from Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIG.GodownID"
tmp = tmp & " Inner Join Txn_ReservationInvoice TPI on TPI.InvoiceNo = TIG.PINo And isnull(TPI.ReverseFlag, '') <> 'Y'"
tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID <> 0 And TIG.ContractID Is not Null)"
tmp = tmp & " And PINo is not Null And TIG.BalanceWeight > 0"
tmp = tmp & " And TIG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group By TIG.GodownID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
       TmpRs1.MoveFirst
       TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
       If Not TmpRs1.EOF Then
          MSHFlexGrid1.TextMatrix(i, E.ReservationIncome) = IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)
       End If
       MSHFlexGrid1.TextMatrix(0, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome))
   Next
End If


tmp = "Select ML.LocationID, Sum(TIG.InvoiceAmount) IncomeAmt"
tmp = tmp & " from Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MC on MG.CMPID = MC.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on MC.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Txn_ProInvoice TPI on TPI.InvoiceNo = TIG.PINo And isnull(TPI.ReverseFlag, '') <> 'Y'"
tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID <> 0 And TIG.ContractID Is Not Null)"
tmp = tmp & " And PINo is not Null And TIG.BalanceWeight > 0"
tmp = tmp & " And TIG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And TIG.GodownID not in  ("
tmp = tmp & " Select TIG.GodownID"
tmp = tmp & " from Txn_InvMonGSLDetail TIG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIG.GodownID"
tmp = tmp & " Inner Join Txn_ReservationInvoice TPI on TPI.InvoiceNo = TIG.PINo And isnull(TPI.ReverseFlag, '') <> 'Y'"
tmp = tmp & " Where TIG.ExchangeTypeID = 1 And (TIG.ContractID <> 0 And TIG.ContractID Is Not Null)"
tmp = tmp & " And PINo is not Null And TIG.BalanceWeight > 0"
tmp = tmp & " And TIG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " )"
tmp = tmp & " Group By ML.LocationID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
'   For i = 2 To MSHFlexGrid1.Rows - 2
'       TmpRs1.MoveFirst
'       TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
'       If Not TmpRs1.EOF Then
'          MSHFlexGrid1.TextMatrix(i, E.ReservationIncome) = IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)
'       End If
'       MSHFlexGrid1.TextMatrix(0, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome))
'   Next
   For i = 1 To TmpRs1.RecordCount
       For k = 2 To MSHFlexGrid1.Rows - 2
           If MSHFlexGrid1.TextMatrix(k, E.LocationID) = TmpRs1!LocationID Then
              TmpRs2.MoveFirst
              TmpRs2.Find "LocationID = " & MSHFlexGrid1.TextMatrix(k, E.LocationID)
              If Not TmpRs2.EOF Then
                 If Opt_Stock.Value = True Then
                    mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
                    MSHFlexGrid1.TextMatrix(k, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(k, E.ReservationIncome)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.000"))
                    MSHFlexGrid1.TextMatrix(0, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.000"))
                 ElseIf Opt_Capacity.Value = True Then
                    MSHFlexGrid1.TextMatrix(k, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(k, E.ReservationIncome)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(k, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
                    MSHFlexGrid1.TextMatrix(0, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(k, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
                 End If
              End If
           End If
       Next
       TmpRs1.MoveNext
   Next
End If


tmp = " Select TI.InvoiceNo, TI.LocationID, sum(TI.ActualAmount) As 'IncomeAmt'"
tmp = tmp & " from Txn_ReservationInvoice TI"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TI.LocationID And TI.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
If Wc <> "" Then
    tmp = tmp & Wc & " And "
Else
    tmp = tmp & " Where "
End If
tmp = tmp & " isnull(TI.ReverseFlag, '') <> 'Y' And TI.Flag in ('A', 'P') And TI.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
'tmp = tmp & " isnull(TI.ReverseFlag, '') <> 'Y' And TI.Flag = 'A' And TI.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And InvoiceNo not in (Select Distinct PINo from Txn_InvMonGSLDetail Where ExchangeTypeID = 1  And YearMonth = '" & mStartDate & "'"
tmp = tmp & " And ContractID <> 0 And ContractID is not null And PINO is not Null)"
tmp = tmp & " Group by TI.InvoiceNo, TI.LocationID " ''--TI.ContractID, TI.LocationID,
tmp = tmp & " Order by  TI.InvoiceNo, TI.LocationID " ''--TI.LocationID,
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For j = 1 To TmpRs1.RecordCount
      tmp = " Select MGSL.GodownID, Sum(MGSL.BalanceWeight) BalanceWeight, sum(MGSL.BalanceBags) BalanceBags From Txn_INvMonGSLDetail MGSL"
      tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MGSL.GodownID"
      tmp = tmp & " Where MGSL.ContractID in"
      tmp = tmp & " (Select ContractID From Txn_ReservationInvoice TI"
      tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TI.LocationID"
      tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
      If Wc <> "" Then
         tmp = tmp & Wc & " And "
      Else
         tmp = tmp & " Where "
      End If
      tmp = tmp & " isnull(TI.ReverseFlag, '') <> 'Y' And (TI.InvoiceNo = " & IIf(IsNull(TmpRs1!InvoiceNo), 0, TmpRs1!InvoiceNo) & " Or TI.PrinciplePINo = " & IIf(IsNull(TmpRs1!InvoiceNo), 0, TmpRs1!InvoiceNo) & ")"
      'tmp = tmp & " And MGSL.BalanceWeight > 0 And MGSL.BalanceBags > 0 And TI.Flag = 'A' And TI.YearMonth Between '" & mStartDate & "'  and '" & mEndDate & "')" '')"
      tmp = tmp & " And MGSL.BalanceWeight > 0 And MGSL.BalanceBags > 0 And TI.Flag in ('A', 'P') And TI.YearMonth Between '" & mStartDate & "'  and '" & mEndDate & "')" '')"
      tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "' ) And MG.StartDate <= '" & mEndDate & "'"
      tmp = tmp & " And MGSL.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
      tmp = tmp & " Group By MGSL.GodownID"
      Call Tmp4Table

      If TmpRs4.RecordCount > 0 Then
         mBalanceWeight = 0
         For k = 1 To TmpRs4.RecordCount
            mBalanceWeight = Val(mBalanceWeight) + IIf(IsNull(TmpRs4!BalanceWeight), 0, TmpRs4!BalanceWeight)
            TmpRs4.MoveNext
         Next

         TmpRs4.MoveFirst
         For k = 1 To TmpRs4.RecordCount
            For i = 2 To MSHFlexGrid1.Rows - 2
               If MSHFlexGrid1.TextMatrix(i, E.GodownID) = TmpRs4!GodownID Then
                  mGodownWeight = IIf(IsNull(TmpRs4!BalanceWeight), 0, TmpRs4!BalanceWeight)
                  MSHFlexGrid1.TextMatrix(i, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome)) + Format((IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt) * mGodownWeight) / mBalanceWeight, "###########.00")
                  MSHFlexGrid1.TextMatrix(0, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) + Format((IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt) * mGodownWeight) / mBalanceWeight, "###########.00")
                  Exit For
               End If
            Next
            TmpRs4.MoveNext
         Next
      Else
         For k = 2 To MSHFlexGrid1.Rows - 2
            If MSHFlexGrid1.TextMatrix(k, E.LocationID) = TmpRs1!LocationID Then
               TmpRs2.MoveFirst
               TmpRs2.Find "LocationID = " & MSHFlexGrid1.TextMatrix(k, E.LocationID)
               If Not TmpRs2.EOF Then
                  If Opt_Stock.Value = True Then
                     mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
                     MSHFlexGrid1.TextMatrix(k, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(k, E.ReservationIncome)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.000"))
                     MSHFlexGrid1.TextMatrix(0, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.000"))
                  ElseIf Opt_Capacity.Value = True Then
                     MSHFlexGrid1.TextMatrix(k, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(k, E.ReservationIncome)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(k, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
                     MSHFlexGrid1.TextMatrix(0, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) + IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(k, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
                  End If
               End If
            End If
         Next
      End If
      TmpRs1.MoveNext
   Next
End If


'''''For Reversal==========================================
tmp = " Select TCA.LocationID,sum(RecOtherDed) RecOtherDed"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCAD"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCA on TCA.TDSAID = TCAD.TDSAID And TCA.SM_Prefix = TCAD.SM_Prefix"
tmp = tmp & " Where TCA.Flag in ('A', 'P') And AdjDate Between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And TCAD.ExpenseAccID  = 90 " '90 means Reservation Reversal"
tmp = tmp & " Group by TCA.LocationID "
Call Tmp5Table

mTotalReservationIncome = MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)
If Val(mTotalReservationIncome) > 0 Then
    For i = 2 To MSHFlexGrid1.Rows - 2
        If TmpRs5.RecordCount > 0 Then
            TmpRs5.MoveFirst
            TmpRs5.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
            If Not TmpRs5.EOF Then
                mRecOtherDed = Format(Val(Val(IIf(IsNull(TmpRs5!RecOtherDed), "", TmpRs5!RecOtherDed)) * Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome))) / Val(mTotalReservationIncome), "########.00")
                MSHFlexGrid1.TextMatrix(i, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome)) - Val(mRecOtherDed)
                MSHFlexGrid1.TextMatrix(0, E.ReservationIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) - Val(mRecOtherDed)
            End If
        End If
    Next
End If





End Sub

Private Sub GetSevenDaysIncome()

tmp = " Select TI.LocationID, sum(TI.Amount) IncomeAmt"
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " Where TI.Flag in ('A', 'P') And TI.InvoiceTypeID = 1 And isnuLL(TI.ReverseFlag,'') <> 'Y'"
tmp = tmp & " And TI.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Group by TI.LocationID"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRs1.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, 12) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.SevenDaysIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SevenDaysIncome))
   Next
End If

tmp = " Select TCA.LocationID,sum(RecOtherDed) RecOtherDed"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCAD"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCA on TCA.TDSAID = TCAD.TDSAID And TCA.SM_Prefix = TCAD.SM_Prefix"
tmp = tmp & " Where TCA.Flag in ('A', 'P') And AdjDate Between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And TCAD.ExpenseAccID  = 91 " '91 means 7 Days Reversal"
tmp = tmp & " group by TCA.LocationID"
Call Tmp5Table

mTotal7DaysIncome = MSHFlexGrid1.TextMatrix(0, E.SevenDaysIncome)
If Val(mTotal7DaysIncome) <> 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      If TmpRs5.RecordCount > 0 Then
         TmpRs5.MoveFirst
         TmpRs5.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs5.EOF Then
            mRecOtherDed = Format(Val(Val(IIf(IsNull(TmpRs5!RecOtherDed), "", TmpRs5!RecOtherDed)) * Val(MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome)) / Val(mTotal7DaysIncome)), "########.00")
            MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome)) - Val(mRecOtherDed)
            MSHFlexGrid1.TextMatrix(0, E.SevenDaysIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.SevenDaysIncome)) - Val(mRecOtherDed)
         End If
      End If
   Next
End If

tmp = " Select MGSL.GodownID, sum(TBGSL.Amount) Amount from Txn_CDTFBOE TB"
tmp = tmp & " Inner Join Txn_CDTFBOE_GSL TBGSL on TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Inner Join Mst_GSL MGSL on MGSL.SM_Prefix = TBGSL.SM_Prefix And MGSL.GSLID = TBGSL.GSLID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MGSL.GodownID"
tmp = tmp & " Where TB.BillingMonth Between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "' ) And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " And TB.TxnType = 'D'"
tmp = tmp & " Group by MGSL.GodownID"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
         MSHFlexGrid1.TextMatrix(0, E.SevenDaysIncome) = IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount) + Val(MSHFlexGrid1.TextMatrix(0, E.SevenDaysIncome))
      End If
   Next
End If

End Sub

Private Sub GetNCDEXRejectedIncome()

tmp = " Select TI.LocationID, sum(TI.Amount) IncomeAmt " '' /(select count(*) from Mst_Godown MG inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID Where MCMP.LocationID = " & TmpRs!LocationID & ") As IncomeAmt "
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " Where TI.Flag in ('A', 'P') And TI.InvoiceTypeID = 2 And isnuLL(TI.ReverseFlag,'') <> 'Y'"
tmp = tmp & " And TI.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
'tmp = tmp & " And TI.LocationID = " & TmpRs!LocationID
tmp = tmp & " Group by TI.LocationID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRs1.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
            If Not TmpRs2.EOF Then
               If Opt_Stock.Value = True Then
                  mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
                  MSHFlexGrid1.TextMatrix(i, E.RejectedIncome) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
               ElseIf Opt_Capacity.Value = True Then
                  MSHFlexGrid1.TextMatrix(i, E.RejectedIncome) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
               End If
            End If
         End If
      MSHFlexGrid1.TextMatrix(0, E.RejectedIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.RejectedIncome))
   Next
End If

tmp = " Select TCA.LocationID,sum(RecOtherDed) RecOtherDed" '
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCAD"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCA on TCA.TDSAID = TCAD.TDSAID And TCA.SM_Prefix = TCAD.SM_Prefix"
tmp = tmp & " Where TCA.Flag in ('A', 'P') And AdjDate Between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And TCAD.ExpenseAccID  = 92 " '92 means Rejected Reversal"
tmp = tmp & " group by TCA.LocationID"
Call Tmp5Table

mTotalRejectedIncome = MSHFlexGrid1.TextMatrix(0, E.RejectedIncome)

If Val(mTotalRejectedIncome) <> 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      If TmpRs5.RecordCount > 0 Then
         TmpRs5.MoveFirst
         TmpRs5.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs5.EOF Then
            mRecOtherDed = Format(Val(Val(IIf(IsNull(TmpRs5!RecOtherDed), "", TmpRs5!RecOtherDed)) * Val(MSHFlexGrid1.TextMatrix(i, E.RejectedIncome))) / Val(mTotalRejectedIncome), "########.00")
            MSHFlexGrid1.TextMatrix(i, E.RejectedIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.RejectedIncome)) - Val(mRecOtherDed)
            MSHFlexGrid1.TextMatrix(0, E.RejectedIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.RejectedIncome)) - Val(mRecOtherDed)
         End If
      End If
   Next
End If

tmp = " Select MGSL.GodownID, sum(TBGSL.Amount) Amount from Txn_CDTFBOE TB"
tmp = tmp & " Inner Join Txn_CDTFBOE_GSL TBGSL on TBGSL.BOEID = TB.BOEID"
tmp = tmp & " Inner Join Mst_GSL MGSL on MGSL.SM_Prefix = TBGSL.SM_Prefix And MGSL.GSLID = TBGSL.GSLID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MGSL.GodownID"
tmp = tmp & " Where TB.BillingMonth Between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "' ) And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " And TB.TxnType = 'W'"
tmp = tmp & " Group by MGSL.GodownID"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.RejectedIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.RejectedIncome)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
         MSHFlexGrid1.TextMatrix(0, E.RejectedIncome) = IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount) + Val(MSHFlexGrid1.TextMatrix(0, E.RejectedIncome))
      End If
   Next
End If
End Sub

Private Sub GetCMIncome()
tmp = " select convert(varchar, cmpid)+'-'+convert(varchar, whrno-cmpid*1000000) whrno, GodownID, CARNo from txn_whr_detail"
Call Tmp1Table

tmp = " select whrno, carno, status, sum(invoiceamt) IncomeAmt from Txn_NLGBillingAmount Where YearMonth = '" & mStartDate & "' group by whrno, carno, status"
Call TmpTableCMPro
'a = 0
If TmpRsCMPro.RecordCount > 0 And TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.Filter = "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      For j = 1 To TmpRs1.RecordCount
         TmpRsCMPro.Filter = ""
         TmpRsCMPro.Filter = "WHRNo = '" & TmpRs1!whrno & "'" ' OR CARNo = " & TmpRs1!CARNo
         If TmpRsCMPro.RecordCount > 0 Then
            For k = 1 To TmpRsCMPro.RecordCount
               'Commented on 04 Jan 2014 after mail from Mayur Sir and approval from Sankar Sir for 60:40 Ratio instead of 50:50 ratio after MCM meeting
               'If TmpRsCMPro!Status = "U" Then
               '   MSHFlexGrid1.TextMatrix(i, E.CMUnbilledIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.CMUnbilledIncome)) + IIf(IsNull(TmpRsCMPro!IncomeAmt), 0, TmpRsCMPro!IncomeAmt * 0.5)
               '   MSHFlexGrid1.TextMatrix(0, E.CMUnbilledIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.CMUnbilledIncome)) + IIf(IsNull(TmpRsCMPro!IncomeAmt), 0, TmpRsCMPro!IncomeAmt * 0.5)
               'ElseIf TmpRsCMPro!Status = "P" Then
               '   MSHFlexGrid1.TextMatrix(i, E.CMPendingIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.CMPendingIncome)) + IIf(IsNull(TmpRsCMPro!IncomeAmt), 0, TmpRsCMPro!IncomeAmt * 0.5)
               '   MSHFlexGrid1.TextMatrix(0, E.CMPendingIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.CMPendingIncome)) + IIf(IsNull(TmpRsCMPro!IncomeAmt), 0, TmpRsCMPro!IncomeAmt * 0.5)
               'ElseIf TmpRsCMPro!Status = "A" Then
               '   MSHFlexGrid1.TextMatrix(i, E.CMApprovedIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.CMApprovedIncome)) + IIf(IsNull(TmpRsCMPro!IncomeAmt), 0, TmpRsCMPro!IncomeAmt * 0.5)
               '   MSHFlexGrid1.TextMatrix(0, E.CMApprovedIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.CMApprovedIncome)) + IIf(IsNull(TmpRsCMPro!IncomeAmt), 0, TmpRsCMPro!IncomeAmt * 0.5)
               'End If
               
               If TmpRsCMPro!Status = "U" Then
                  MSHFlexGrid1.TextMatrix(i, E.CMUnbilledIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.CMUnbilledIncome)) + IIf(IsNull(TmpRsCMPro!IncomeAmt), 0, TmpRsCMPro!IncomeAmt * 0.6)
                  MSHFlexGrid1.TextMatrix(0, E.CMUnbilledIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.CMUnbilledIncome)) + IIf(IsNull(TmpRsCMPro!IncomeAmt), 0, TmpRsCMPro!IncomeAmt * 0.6)
               ElseIf TmpRsCMPro!Status = "P" Then
                  MSHFlexGrid1.TextMatrix(i, E.CMPendingIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.CMPendingIncome)) + IIf(IsNull(TmpRsCMPro!IncomeAmt), 0, TmpRsCMPro!IncomeAmt * 0.6)
                  MSHFlexGrid1.TextMatrix(0, E.CMPendingIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.CMPendingIncome)) + IIf(IsNull(TmpRsCMPro!IncomeAmt), 0, TmpRsCMPro!IncomeAmt * 0.6)
               ElseIf TmpRsCMPro!Status = "A" Then
                  MSHFlexGrid1.TextMatrix(i, E.CMApprovedIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.CMApprovedIncome)) + IIf(IsNull(TmpRsCMPro!IncomeAmt), 0, TmpRsCMPro!IncomeAmt * 0.6)
                  MSHFlexGrid1.TextMatrix(0, E.CMApprovedIncome) = Val(MSHFlexGrid1.TextMatrix(0, E.CMApprovedIncome)) + IIf(IsNull(TmpRsCMPro!IncomeAmt), 0, TmpRsCMPro!IncomeAmt * 0.6)
               End If
               
               TmpRsCMPro.MoveNext
            Next
         Else
            'a = a
            'a = Val(a) + IIf(IsNull(TmpRsCMPro!IncomeAmt), 0, TmpRsCMPro!IncomeAmt * 0.6)
         End If
         TmpRs1.MoveNext
      Next
      TmpRs1.Filter = ""
   Next
End If

End Sub

Private Sub GetOtherIncome()


'Commented on 30 Jan 2014 to introduce Loading Unloading charges in Other Income
'tmp = tmp & " Where TI.Flag in ('A', 'P') And TI.InvoiceTypeID = 5 And isnuLL(TI.ReverseFlag,'') <> 'Y'"

tmp = " Select TI.LocationID, sum(TI.Amount) IncomeAmt " ''''/(select count(*) from Mst_Godown MG inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID Where MCMP.LocationID = " & TmpRs!LocationID & ") As IncomeAmt "
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " Where TI.Flag in ('A', 'P') And TI.InvoiceTypeID in (5,10) And isnuLL(TI.ReverseFlag,'') <> 'Y'"
tmp = tmp & " And TI.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Group by TI.LocationID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRs1.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.OtherIncome) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.OtherIncome) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.OtherIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.OtherIncome))
   Next
End If

End Sub

Private Sub GetSCMIncome()
tmp = " Select TI.LocationID, sum(TI.Amount) IncomeAmt " ''''/(select count(*) from Mst_Godown MG inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID Where MCMP.LocationID = " & TmpRs!LocationID & ") As IncomeAmt "
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " Where TI.Flag in ('A', 'P') And TI.InvoiceTypeID = 11 And isnuLL(TI.ReverseFlag,'') <> 'Y'"
tmp = tmp & " And TI.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Group by TI.LocationID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRs1.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.SCMIncome) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.SCMIncome) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.SCMIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SCMIncome))
   Next
End If
End Sub

Private Sub GetNNNRIncomethroughAdhoc()
tmp = " Select TI.LocationID, sum(TI.Amount) IncomeAmt " ''''/(select count(*) from Mst_Godown MG inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID Where MCMP.LocationID = " & TmpRs!LocationID & ") As IncomeAmt "
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " Where TI.Flag in ('A', 'P') And TI.InvoiceTypeID = 7 And isnuLL(TI.ReverseFlag,'') <> 'Y'"
tmp = tmp & " And TI.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Group by TI.LocationID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRs1.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.NNNRIncomeAdhoc) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.NNNRIncomeAdhoc) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.NNNRIncomeAdhoc) = Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.NNNRIncomeAdhoc))
   Next
End If
End Sub

Private Sub GetDPIncomethroughAdhoc()
tmp = " Select TI.LocationID, sum(TI.Amount) IncomeAmt " ''''/(select count(*) from Mst_Godown MG inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID Where MCMP.LocationID = " & TmpRs!LocationID & ") As IncomeAmt "
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " Where TI.Flag in ('A', 'P') And TI.InvoiceTypeID = 8 And isnuLL(TI.ReverseFlag,'') <> 'Y'"
tmp = tmp & " And TI.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Group by TI.LocationID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRs1.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.DPIncomeAdHoc) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.DPIncomeAdHoc) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.DPIncomeAdHoc) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeAdHoc))
   Next
End If

End Sub

Private Sub GetSpotIncomethroughAdhoc()
tmp = " Select TI.LocationID, sum(TI.Amount) IncomeAmt " ''''/(select count(*) from Mst_Godown MG inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID Where MCMP.LocationID = " & TmpRs!LocationID & ") As IncomeAmt "
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " Where TI.Flag in ('A', 'P') And TI.InvoiceTypeID = 3 And isnuLL(TI.ReverseFlag,'') <> 'Y'"
tmp = tmp & " And TI.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Group by TI.LocationID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRs1.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.SpotIncomeAdHoc) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.SpotIncomeAdHoc) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.SpotIncomeAdHoc) = Val(MSHFlexGrid1.TextMatrix(i, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.SpotIncomeAdHoc))
   Next
End If
End Sub

Private Sub GetReservationIncomethroughAdhoc()

'tmp = " Select TI.LocationID, sum(TI.Amount) IncomeAmt " ''''/(select count(*) from Mst_Godown MG inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID Where MCMP.LocationID = " & TmpRs!LocationID & ") As IncomeAmt "
'tmp = tmp & " from Txn_TempInvoice TI"
'tmp = tmp & " Where TI.Flag in ('A', 'P') And TI.InvoiceTypeID = 4 And isnuLL(TI.ReverseFlag,'') <> 'Y'"
'tmp = tmp & " And TI.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
'tmp = tmp & " Group by TI.LocationID "

tmp = " Select TISGD.GodownID, sum(TISGD.InvoiceAmt) IncomeAmt"
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " Inner Join Txn_TempInvSRGodownDetails TISGD on TI.PINo = TISGD.PINo"
tmp = tmp & " Where TI.Flag in ('A', 'P') And TI.InvoiceTypeID = 4 And isnuLL(TI.ReverseFlag,'') <> 'Y'"
tmp = tmp & " And TI.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Group by TISGD.GodownID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
'   For i = 2 To MSHFlexGrid1.Rows - 2
''       TmpRs1.MoveFirst
''       TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
''       If Not TmpRs1.EOF Then
''          TmpRs2.MoveFirst
''          TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
''          If Not TmpRs2.EOF Then
''             If Opt_Stock.Value = True Then
''                mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
''                MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
''              ElseIf Opt_Capacity.Value = True Then
''                MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
'' '            End If
'' '         End If
''      End If
''      MSHFlexGrid1.TextMatrix(0, E.ReservationIncomeAdhoc) = Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncomeAdhoc))
'
'       TmpRs1.MoveFirst
'       TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
'       If Not TmpRs1.EOF Then
'          If Opt_Stock.Value = True Then
'             mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
'             MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc) = TmpRs1!IncomeAmt  'IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
'          ElseIf Opt_Capacity.Value = True Then
'             MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc) = TmpRs1!IncomeAmt 'IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
'          End If
'       Else
'          TxtUnallocatedSal = TxtUnallocatedSal & MSHFlexGrid1.TextMatrix(i, E.GodownID) & vbCrLf
'       End If
'       MSHFlexGrid1.TextMatrix(0, E.ReservationIncomeAdhoc) = Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncomeAdhoc))
'   Next
   TxtUnallocatedSal = ""
   For X = 1 To TmpRs1.RecordCount
       mStatus = 0
       For i = 2 To MSHFlexGrid1.Rows - 2
           If MSHFlexGrid1.TextMatrix(i, E.GodownID) = TmpRs1!GodownID Then
              mStatus = 1
              MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc) = TmpRs1!IncomeAmt 'IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
              MSHFlexGrid1.TextMatrix(0, E.ReservationIncomeAdhoc) = Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncomeAdhoc))
           End If
       Next
       If mStatus = 0 Then
          TxtUnallocatedSal = TxtUnallocatedSal & TmpRs1!GodownID & vbCrLf
       End If
       TmpRs1.MoveNext
   Next
End If

End Sub

Private Sub GetSpillageIncome()
tmp = " Select TI.LocationID, sum(TI.Amount) IncomeAmt"
tmp = tmp & " from Txn_TempInvoice TI"
tmp = tmp & " Where TI.Flag in ('A', 'P') And TI.InvoiceTypeID = 6 And isnuLL(TI.ReverseFlag,'') <> 'Y'"
tmp = tmp & " And TI.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Group by TI.LocationID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRs1.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.LeftOverStock) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.LeftOverStock) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.LeftOverStock) = Val(MSHFlexGrid1.TextMatrix(i, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(0, E.LeftOverStock))
   Next
End If
End Sub

Private Sub GetPledgeIncome()
tmp = " Select TPCR.LocationID, sum(TPCR.InstrumentAmount) IncomeAmt"
tmp = tmp & " from Txn_PledgeCashReceipt TPCR"
tmp = tmp & " Where TPCR.InstrumentDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " group by TPCR.LocationID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRs1.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.PledgeIncome) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.PledgeIncome) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.PledgeIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.PledgeIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.PledgeIncome))
   Next
End If
End Sub

Private Sub GetVendorExpense()

tmp = " Select TVCD.LocationID,  sum(TVCD.Amount) IncomeAmt  "
tmp = tmp & " From Txn_VendorClaimDetail TVCD"
tmp = tmp & " Inner Join Txn_VendorClaim TVC on TVC.ClaimID = TVCD.ClaimID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TVCD.LocationID And TVCD.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
tmp = tmp & " Inner Join Mst_ExpenseType MIET on MIET.ExpenseTypeID = TVCD.ExpenseTypeID"
tmp = tmp & " Where TVC.Flag = 'A' And TVCD.ExpenseTypeID not in (17, 18, 8, 69)  And TVCD.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And isnull(MIET.MISExpenseGroup, '') <> 'NA'"
tmp = tmp & " group by TVCD.LocationID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRs1.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.VendorExpenses) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.VendorExpenses) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.VendorExpenses) = Val(MSHFlexGrid1.TextMatrix(i, E.VendorExpenses)) + Val(MSHFlexGrid1.TextMatrix(0, E.VendorExpenses))
   Next
End If

tmp = " Select TVCD.LocationID,  sum(TVCD.Amount) IncomeAmt  "
tmp = tmp & " From Txn_VendorClaimDetail TVCD"
tmp = tmp & " Inner Join Txn_VendorClaim TVC on TVC.ClaimID = TVCD.ClaimID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TVCD.LocationID And TVCD.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
tmp = tmp & " Inner Join Mst_ExpenseType MIET on MIET.ExpenseTypeID = TVCD.ExpenseTypeID"
tmp = tmp & " Where TVC.Flag = 'P' And TVCD.ExpenseTypeID not in (17, 18, 8, 69)  And TVCD.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And isnull(MIET.MISExpenseGroup, '') <> 'NA'"
tmp = tmp & " group by TVCD.LocationID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRs1.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.VendorExpensesAccrued) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.VendorExpensesAccrued) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.VendorExpensesAccrued) = Val(MSHFlexGrid1.TextMatrix(i, E.VendorExpensesAccrued)) + Val(MSHFlexGrid1.TextMatrix(0, E.VendorExpensesAccrued))
   Next
End If

tmp = " Select TCSD.LocationID, Sum(TCSD.Amount) IncomeAmt"
tmp = tmp & " From Txn_VendorClaimSegmentDistribution TCSD"
tmp = tmp & " Inner Join Txn_VendorClaim TCS on TCS.ClaimID = TCSD.ClaimID"
tmp = tmp & " Where TCSD.SegmentID = 1 And TCS.Flag = 'A' And TCSD.FLag = 'L' "
tmp = tmp & " And '" & mStartDate & "' Between TCS.ExpenseFrom And TCS.ExpenseTo And TCSD.ExpenseTypeID not in (107,66) " 'And TCSD.ExpenseTypeID not in (107) "
tmp = tmp & " group by TCSD.LocationID"
Call TmpTableHRPro

If TmpRsHRPro.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRsHRPro.MoveFirst
      TmpRsHRPro.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRsHRPro.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.VendorExpenses) = Val(MSHFlexGrid1.TextMatrix(i, E.VendorExpenses)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
               MSHFlexGrid1.TextMatrix(0, E.VendorExpenses) = Val(MSHFlexGrid1.TextMatrix(0, E.VendorExpenses)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.VendorExpenses) = Val(MSHFlexGrid1.TextMatrix(i, E.VendorExpenses)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
               MSHFlexGrid1.TextMatrix(0, E.VendorExpenses) = Val(MSHFlexGrid1.TextMatrix(0, E.VendorExpenses)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
   Next
End If

tmp = " Select TCSD.LocationID, Sum(TCSD.Amount) IncomeAmt"
tmp = tmp & " From Txn_VendorClaimSegmentDistribution TCSD"
tmp = tmp & " Inner Join Txn_VendorClaim TCS on TCS.ClaimID = TCSD.ClaimID"
tmp = tmp & " Where TCSD.SegmentID = 1 And TCS.Flag = 'P'  And TCSD.FLag = 'L'  "
tmp = tmp & " And '" & mStartDate & "' Between TCS.ExpenseFrom And TCS.ExpenseTo And TCSD.ExpenseTypeID not in (107,66) " 'And TCSD.ExpenseTypeID not in (107) "
tmp = tmp & " group by TCSD.LocationID"

Call TmpTableHRPro

If TmpRsHRPro.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRsHRPro.MoveFirst
      TmpRsHRPro.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRsHRPro.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.VendorExpensesAccrued) = Val(MSHFlexGrid1.TextMatrix(i, E.VendorExpensesAccrued)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
               MSHFlexGrid1.TextMatrix(0, E.VendorExpensesAccrued) = Val(MSHFlexGrid1.TextMatrix(0, E.VendorExpensesAccrued)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.VendorExpensesAccrued) = Val(MSHFlexGrid1.TextMatrix(i, E.VendorExpensesAccrued)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
               MSHFlexGrid1.TextMatrix(0, E.VendorExpensesAccrued) = Val(MSHFlexGrid1.TextMatrix(0, E.VendorExpensesAccrued)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
   Next
End If



'------------------------


tmp = " Select TCSD.LocationID, Sum(TCSD.Amount) IncomeAmt"
tmp = tmp & " From Txn_VendorClaimSegmentDistribution TCSD"
tmp = tmp & " Inner Join Txn_VendorClaim TCS on TCS.ClaimID = TCSD.ClaimID"
tmp = tmp & " Where TCSD.SegmentID = 1 And TCS.Flag = 'A' And TCSD.FLag = 'G' "
tmp = tmp & " And '" & mStartDate & "' Between TCS.ExpenseFrom And TCS.ExpenseTo And TCSD.ExpenseTypeID not in (107,66) " 'And TCSD.ExpenseTypeID not in (107) "
tmp = tmp & " group by TCSD.LocationID"
Call TmpTableHRPro

If TmpRsHRPro.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRsHRPro.MoveFirst
      TmpRsHRPro.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRsHRPro.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.VendorExpenses) = Val(MSHFlexGrid1.TextMatrix(i, E.VendorExpenses)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
               MSHFlexGrid1.TextMatrix(0, E.VendorExpenses) = Val(MSHFlexGrid1.TextMatrix(0, E.VendorExpenses)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.VendorExpenses) = Val(MSHFlexGrid1.TextMatrix(i, E.VendorExpenses)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
               MSHFlexGrid1.TextMatrix(0, E.VendorExpenses) = Val(MSHFlexGrid1.TextMatrix(0, E.VendorExpenses)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
   Next
End If

tmp = " Select TCSD.LocationID, Sum(TCSD.Amount) IncomeAmt"
tmp = tmp & " From Txn_VendorClaimSegmentDistribution TCSD"
tmp = tmp & " Inner Join Txn_VendorClaim TCS on TCS.ClaimID = TCSD.ClaimID"
tmp = tmp & " Where TCSD.SegmentID = 1 And TCS.Flag = 'P' And TCSD.FLag = 'G'  "
tmp = tmp & " And '" & mStartDate & "' Between TCS.ExpenseFrom And TCS.ExpenseTo And TCSD.ExpenseTypeID not in (107,66) " 'And TCSD.ExpenseTypeID not in (107) "
tmp = tmp & " group by TCSD.LocationID"

Call TmpTableHRPro

If TmpRsHRPro.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRsHRPro.MoveFirst
      TmpRsHRPro.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRsHRPro.EOF Then
         TmpRs2.MoveFirst
         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
         If Not TmpRs2.EOF Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.VendorExpensesAccrued) = Val(MSHFlexGrid1.TextMatrix(i, E.VendorExpensesAccrued)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
               MSHFlexGrid1.TextMatrix(0, E.VendorExpensesAccrued) = Val(MSHFlexGrid1.TextMatrix(0, E.VendorExpensesAccrued)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.VendorExpensesAccrued) = Val(MSHFlexGrid1.TextMatrix(i, E.VendorExpensesAccrued)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
               MSHFlexGrid1.TextMatrix(0, E.VendorExpensesAccrued) = Val(MSHFlexGrid1.TextMatrix(0, E.VendorExpensesAccrued)) + IIf(IsNull(TmpRsHRPro!IncomeAmt), 0, Format(TmpRsHRPro!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            End If
         End If
      End If
   Next
End If

End Sub

Private Sub GetSecurityExpense()
tmp = " Select MSLD.TxnDetailID, TSBD.LocationID, sum(TSBD.BillDetailAmt) Amount"
tmp = tmp & " From Txn_SecurityBilling TSB"
tmp = tmp & " Inner Join Txn_SecurityBillingDetail TSBD on TSBD.BillTxnID = TSB.BillTxnID"
tmp = tmp & " Inner Join Txn_MonthwiseSecLocDetail TMSLD on TMSLD.MonthwiseTxnDetailID = TSBD.MonthwiseTxnDetailID"
tmp = tmp & " Inner Join Mst_SeclocMappingDetail MSLD on MSLD.TxnDetailID = TMSLD.SecLocMapDetailID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TSBD.LocationID And TSBD.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
If Wc <> "" Then
    tmp = tmp & Wc & " And "
Else
    tmp = tmp & " Where "
End If
tmp = tmp & " TSB.Status in ('A', 'P') And TSB.MonthYear Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " group by MSLD.TxnDetailID, TSBD.LocationID "
Call Tmp1Table

tmp = " Select"
tmp = tmp & " MSLGMD.TxnDetailID , Sum(MG.GodownCapacity) Capacity"
tmp = tmp & " From Mst_SeclocGodMappingDetail MSLGMD"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MSLGMD.GodownID"
tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "' ) And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " group by MSLGMD.TxnDetailID"
Call Tmp3Table

tmp = " Select"
tmp = tmp & " MSLGMD.TxnDetailID , MG.GodownCapacity, MG.GodownID"
tmp = tmp & " From Mst_SeclocGodMappingDetail MSLGMD"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MSLGMD.GodownID"
tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "' ) And MG.StartDate <= '" & mEndDate & "'"
Call Tmp4Table

'If TmpRs3.RecordCount > 0 Then
   For j = 1 To TmpRs1.RecordCount
      TmpRs3.MoveFirst
      TmpRs3.Find "TxnDetailID = " & TmpRs1!TxnDetailID
      If Not TmpRs3.EOF Then
         tmp = " Select"
         tmp = tmp & " MSLGMD.TxnDetailID , MG.GodownCapacity, MG.GodownID"
         tmp = tmp & " From Mst_SeclocGodMappingDetail MSLGMD"
         tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MSLGMD.GodownID"
         tmp = tmp & " Where MSLGMD.TxnDetailId = " & TmpRs1!TxnDetailID
         tmp = tmp & " and (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "' ) And MG.StartDate <= '" & mEndDate & "'"
         Call Tmp4Table
         
         For i = 2 To MSHFlexGrid1.Rows - 2
            TmpRs4.MoveFirst
            TmpRs4.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
            If Not TmpRs4.EOF Then
               MSHFlexGrid1.TextMatrix(i, E.SecurityFeesPaid) = Val(MSHFlexGrid1.TextMatrix(i, E.SecurityFeesPaid)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs3!Capacity, "#############.00"))
               MSHFlexGrid1.TextMatrix(0, E.SecurityFeesPaid) = Val(MSHFlexGrid1.TextMatrix(0, E.SecurityFeesPaid)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs3!Capacity, "#############.00"))
            End If
         Next
      Else
         For i = 2 To MSHFlexGrid1.Rows - 2
            If MSHFlexGrid1.TextMatrix(i, E.LocationID) = TmpRs1!LocationID Then
               TmpRs2.MoveFirst
               TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
               If Not TmpRs2.EOF Then
                  If Opt_Stock.Value = True Then
                     mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
                     MSHFlexGrid1.TextMatrix(i, E.SecurityFeesPaid) = IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amountt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
                  ElseIf Opt_Capacity.Value = True Then
                     MSHFlexGrid1.TextMatrix(i, E.SecurityFeesPaid) = Val(MSHFlexGrid1.TextMatrix(i, E.SecurityFeesPaid)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
                     MSHFlexGrid1.TextMatrix(0, E.SecurityFeesPaid) = Val(MSHFlexGrid1.TextMatrix(0, E.SecurityFeesPaid)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
                  End If
               End If
            End If
         Next
      End If
      TmpRs1.MoveNext
   Next
'End If

tmp = " select  mslm.locationid, mslmd.txndetailid, ((mslmd.ratepermonth * count(attendanceid)) / datediff(d, '" & mStartDate & "','" & mEndDate & "')+1)  Amount"
tmp = tmp & " from mst_seclocmappingdetail mslmd"
tmp = tmp & " inner join txn_securityattendance tsa on mslmd.txndetailid = tsa.txndetailid"
tmp = tmp & " inner join mst_seclocmapping mslm on mslmd.txnid = mslm.txnid"
tmp = tmp & " where tsa.status  = 'A' and tsa.attendancedate between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " and mslmd.txndetailid not in (select isnull(msld.SecLocMapDetailID,0) from Txn_MonthwiseSecLocDetail msld"
tmp = tmp & " inner join Txn_MonthwiseSecLoc msl on msl.monthwisetxnid = msld.monthwisetxnid"
tmp = tmp & " where monthyear Between  '" & mStartDate & "' And '" & mEndDate & "')"
tmp = tmp & " group by mslm.locationid, mslmd.txndetailid, mslmd.enddate, mslmd.ratepermonth"
Call Tmp1Table

'If TmpRs3.RecordCount > 0 Then
   For j = 1 To TmpRs1.RecordCount
      TmpRs3.MoveFirst
      TmpRs3.Find "TxnDetailID = " & TmpRs1!TxnDetailID
      If Not TmpRs3.EOF Then
         tmp = " Select"
         tmp = tmp & " MSLGMD.TxnDetailID , MG.GodownCapacity, MG.GodownID"
         tmp = tmp & " From Mst_SeclocGodMappingDetail MSLGMD"
         tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MSLGMD.GodownID"
         tmp = tmp & " Where TxnDetailId = " & TmpRs1!TxnDetailID
         Call Tmp4Table
            
         For i = 2 To MSHFlexGrid1.Rows - 2
            TmpRs4.MoveFirst
            TmpRs4.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
            If Not TmpRs4.EOF Then
               MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs3!Capacity, "#############.00"))
               MSHFlexGrid1.TextMatrix(0, E.SecurityFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(0, E.SecurityFeesAccrual)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs3!Capacity, "#############.00"))
            End If
         Next
      Else
         For i = 2 To MSHFlexGrid1.Rows - 2
            If MSHFlexGrid1.TextMatrix(i, E.LocationID) = TmpRs1!LocationID Then
               TmpRs2.MoveFirst
               TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
               If Not TmpRs2.EOF Then
                  If Opt_Stock.Value = True Then
                     mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
                     MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
                     MSHFlexGrid1.TextMatrix(0, E.SecurityFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(0, E.SecurityFeesAccrual)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
                  ElseIf Opt_Capacity.Value = True Then
                     MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
                     MSHFlexGrid1.TextMatrix(0, E.SecurityFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(0, E.SecurityFeesAccrual)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
                  End If
               End If
            End If
         Next
      End If
      TmpRs1.MoveNext
   Next
'End If

tmp = " select MSLD.TxnDetailID, MSL.LocationID, sum(MSLD.RatePerMonth) Amount from Mst_SeclocMappingDetail  MSLD"
tmp = tmp & " Inner Join Mst_SecLocMapping MSL on MSL.TxnID = MSLD.TxnID And MSL.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
tmp = tmp & " where '" & mStartDate & "' Between MSLD.StartDate And MSLD.EndDate"
tmp = tmp & " And MSLD.TxnDetailID not in (Select SecLocMapDetailID from Txn_MonthwiseSecLocDetail  TMSLD"
tmp = tmp & " Inner Join Txn_SecurityBillingDetail TSBD on TSBD.MonthwiseTxnDetailID = TMSLD.MonthwiseTxnDetailID"
tmp = tmp & " Inner Join Txn_SecurityBilling TSB on TSB.BillTxnID = TSBD.BillTxnID"
tmp = tmp & " Where TSB.Status in ('A', 'P', 'R') And TSB.MonthYear Between  '" & mStartDate & "' And '" & mEndDate & "')"
tmp = tmp & " And MSLD.TxnDetailID not in (select txndetailid from Txn_SecurityAttendance "
tmp = tmp & " where attendancedate between '" & mStartDate & "' And '" & mEndDate & "' And Status in ('L','R'))"
tmp = tmp & " group by MSLD.TxnDetailID, MSL.LocationID"
Call Tmp1Table

'If TmpRs3.RecordCount > 0 Then
   For j = 1 To TmpRs1.RecordCount
      TmpRs3.MoveFirst
      TmpRs3.Find "TxnDetailID = " & TmpRs1!TxnDetailID
      If Not TmpRs3.EOF Then
         tmp = " Select"
         tmp = tmp & " MSLGMD.TxnDetailID , MG.GodownCapacity, MG.GodownID"
         tmp = tmp & " From Mst_SeclocGodMappingDetail MSLGMD"
         tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = MSLGMD.GodownID"
         tmp = tmp & " Where TxnDetailId = " & TmpRs1!TxnDetailID
         Call Tmp4Table
            
         For i = 2 To MSHFlexGrid1.Rows - 2
            TmpRs4.MoveFirst
            TmpRs4.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
            If Not TmpRs4.EOF Then
               MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs3!Capacity, "#############.00"))
               MSHFlexGrid1.TextMatrix(0, E.SecurityFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(0, E.SecurityFeesAccrual)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs3!Capacity, "#############.00"))
            End If
         Next
      Else
         For i = 2 To MSHFlexGrid1.Rows - 2
            If MSHFlexGrid1.TextMatrix(i, E.LocationID) = TmpRs1!LocationID Then
               TmpRs2.MoveFirst
               TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
               If Not TmpRs2.EOF Then
                  If Opt_Stock.Value = True Then
                     mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
                     MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
                     MSHFlexGrid1.TextMatrix(0, E.SecurityFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(0, E.SecurityFeesAccrual)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
                  ElseIf Opt_Capacity.Value = True Then
                     MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(i, E.SecurityFeesAccrual)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
                     MSHFlexGrid1.TextMatrix(0, E.SecurityFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(0, E.SecurityFeesAccrual)) + IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
                  End If
               End If
            End If
         Next
      End If
      TmpRs1.MoveNext
   Next
'End If
   
End Sub

Private Sub GetFumigationExpense()
tmp = " Select TVCD.LocationID, sum(TVCD.Amount) IncomeAmt "
tmp = tmp & " From Txn_VendorClaimDetail TVCD"
tmp = tmp & " Inner Join Txn_VendorClaim TVC on TVC.ClaimID = TVCD.ClaimID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = TVCD.LocationID And TVCD.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
tmp = tmp & " Where TVC.Flag in ('A', 'P') And TVCD.ExpenseTypeID = 18 And TVCD.InvoiceDate Between  '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Group by TVCD.LocationID "
Call Tmp1Table

tmp = " Select MCMP.LocationID, sum(TIMG.BalanceWeight) BalanceWeight"
tmp = tmp & " From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Where TIMG.ExchangeTypeID in (1, 2) And TIMG.BalanceWeight > 0 And TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group by MCMP.LocationID"
Call Tmp3Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
      If Not TmpRs1.EOF Then
         TmpRs3.MoveFirst
         TmpRs3.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
         If Not TmpRs3.EOF Then
            'If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.FumigationFeesVendor) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs3!BalanceWeight, "#############.00"))
            'ElseIf Opt_Capacity.Value = True Then
               'MSHFlexGrid1.TextMatrix(I, E.FumigationFeesVendor) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(I, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
            'End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.FumigationFeesVendor) = Val(MSHFlexGrid1.TextMatrix(i, E.FumigationFeesVendor)) + Val(MSHFlexGrid1.TextMatrix(0, E.FumigationFeesVendor))
   Next
End If

'tmp = " Select ML.LocationName,MC.CMPName,  MG.GodownID, MG.GodownNo,TF.VendorRate,TF.VendorQuantity,TF.VendorRate*TF.VendorQuantity as Amount"
tmp = " Select MG.GodownID, sum(TFD.VendorRate*TFD.VendorQuantity) as Amount"
tmp = tmp & " from Txn_Fumigation TF"
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TF.GodownID"
tmp = tmp & " Inner Join Txn_FumigationDetail TFD on TFD.FumigationID = TF.FumigationID"
'tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
'tmp = tmp & " Inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where TF.FumigationDate between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And TF.MappingID in ( select mappingid from Txn_FumigationApproval Where MonthYear between '" & mStartDate & "' And '" & mEndDate & "' And Flag in ('A', 'P'))"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group by MG.GodownID"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.FumigationFeesPaid) = IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
      End If
      MSHFlexGrid1.TextMatrix(0, E.FumigationFeesPaid) = Val(MSHFlexGrid1.TextMatrix(i, E.FumigationFeesPaid)) + Val(MSHFlexGrid1.TextMatrix(0, E.FumigationFeesPaid))
   Next
End If

tmp = " Select MG.GodownID, sum(TS.VendorRate*TS.VendorQuantity) as Amount"
tmp = tmp & " from Txn_Spraying TS"
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TS.GodownID"
'tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
'tmp = tmp & " Inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where TS.MappingID>0 And TS.SprayingDate between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group by MG.GodownID"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.FumigationFeesPaid) = Val(MSHFlexGrid1.TextMatrix(i, E.FumigationFeesPaid)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
         MSHFlexGrid1.TextMatrix(0, E.FumigationFeesPaid) = IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount) + Val(MSHFlexGrid1.TextMatrix(0, E.FumigationFeesPaid))
      End If
   Next
End If

tmp = " Select MG.GodownID, sum(TFD.Amount) Amount"
tmp = tmp & " from Txn_Fumigation TF"
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TF.GodownID"
tmp = tmp & " Inner Join Txn_FumigationDetail TFD on TFD.FumigationID = TF.FumigationID"
'tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
'tmp = tmp & " Inner join Mst_Location ML on MC.LocationID=ML.LocationID"
'tmp = tmp & " Where IsNull(TF.MappingID, 0) = 0 And isnull(TF.BillFlag,'')='' And TF.FumigationDate between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Where IsNull(TF.MappingID, 0) = 0 And isnull(TFD.BillableFlag,'Y')='Y' And TF.FumigationDate between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group by MG.GodownID"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.FumigationFeesAccrual) = IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
      End If
      MSHFlexGrid1.TextMatrix(0, E.FumigationFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(i, E.FumigationFeesAccrual)) + Val(MSHFlexGrid1.TextMatrix(0, E.FumigationFeesAccrual))
   Next
End If

tmp = " Select MG.GodownID, sum(TS.Amount) Amount"
tmp = tmp & " from Txn_Spraying TS"
tmp = tmp & " Inner join Mst_Godown MG on MG.GodownID=TS.GodownID"
'tmp = tmp & " Inner join Mst_CMP MC on MC.CMPID=MG.CMPID"
'tmp = tmp & " Inner join Mst_Location ML on MC.LocationID=ML.LocationID"
tmp = tmp & " Where IsNull(TS.MappingID, 0) = 0 And isnull(TS.BillFlag,'')=''"
tmp = tmp & " And TS.SprayingDate between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group by MG.GodownID"
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.FumigationFeesAccrual) = Val(MSHFlexGrid1.TextMatrix(i, E.FumigationFeesAccrual)) + IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
         MSHFlexGrid1.TextMatrix(0, E.FumigationFeesAccrual) = IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount) + Val(MSHFlexGrid1.TextMatrix(0, E.FumigationFeesAccrual))
      End If
   Next
End If

End Sub

Private Sub GetInsuranceCharges()

For i = 2 To MSHFlexGrid1.Rows - 2
   tmp = " Select TIMG.CommodityID, sum(TIMG.BalanceBags) BalanceBags, sum(TIMG.BalanceWeight) BalanceQty"
   tmp = tmp & " From Txn_InvMonGSLDetail TIMG"
   tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
   tmp = tmp & " Where TIMG.BalanceWeight > 0 And TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
   tmp = tmp & " And TIMG.GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
   tmp = tmp & " Group by TIMG.CommodityID"
   Call TmpTable
    
   mBalanceBags = 0
   mBalanceQty = 0
   mTotalAUM = 0
   
   If TmpRs.RecordCount > 0 Then
      For j = 1 To TmpRs.RecordCount
         mBalanceBags = Val(mBalanceBags) + IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
         mBalanceQty = Val(mBalanceQty) + IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)
            
         tmp = " Select MCP.* From Mst_CommodityPrice MCP"
         tmp = tmp & " Inner Join Mst_State MS on MS.StateID = MCP.StateID"
         tmp = tmp & " Inner Join Mst_Location ML on ML.StateID = MS.StateID"
         tmp = tmp & " Where ML.LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID)) & " And MCP.CommodityID = " & TmpRs!CommodityID
         tmp = tmp & " And YearMonth = ("
         tmp = tmp & " Select Max(YearMonth) From Mst_CommodityPrice MCP"
         tmp = tmp & " Inner Join Mst_State MS on MS.StateID = MCP.StateID"
         tmp = tmp & " Inner Join Mst_Location ML on ML.StateID = MS.StateID"
         tmp = tmp & " Where YearMonth <= '" & Format(mStartDate, "yyyy-mm") & "' And ML.LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID)) & " And MCP.CommodityID = " & TmpRs!CommodityID
         tmp = tmp & " )"
         Call Tmp1Table
            
         If TmpRs1.RecordCount > 0 Then
            If IsNull(TmpRs1!Day31) = False Then
               mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day31)
            ElseIf IsNull(TmpRs1!Day30) = False Then
               mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day30)
            ElseIf IsNull(TmpRs1!Day29) = False Then
               mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day29)
            ElseIf IsNull(TmpRs1!Day28) = False Then
               mTotalAUM = Val(mTotalAUM) + ((IIf(IsNull(TmpRs!BalanceQty), 0, TmpRs!BalanceQty)) * TmpRs1!Day28)
            End If
         End If
         TmpRs.MoveNext
      Next
   End If
   If Format(mStartDate, "yyyy-mm") >= "2012-04" And Format(mStartDate, "yyyy-mm") <= "2013-03" Then
      mInsuranceCharges = Val(mTotalAUM) * 0.0003 / 12
   Else
      mInsuranceCharges = Val(mTotalAUM) * 0.0006 / 12
   End If
   
   MSHFlexGrid1.TextMatrix(i, E.InsurancePremium) = Format(Val(mInsuranceCharges), "###########.00")
   MSHFlexGrid1.TextMatrix(0, E.InsurancePremium) = Val(MSHFlexGrid1.TextMatrix(0, E.InsurancePremium)) + Val(MSHFlexGrid1.TextMatrix(i, E.InsurancePremium))
Next
End Sub

Private Sub GetRentExpense()

'Commented on 30 Jan 2014 After discussion to encorportae arrears amount
'tmp = " Select TGRD.GodownID, sum(TGRD.Amount) As IncomeAmt "

tmp = " Select TGRD.GodownID, sum(TGRD.Amount)+ Sum(isnull(TGRD.Arrears,0)) As IncomeAmt "
tmp = tmp & " From Txn_GodownRentDetail TGRD"
tmp = tmp & " Inner Join Txn_GodownRent TGR on TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TGRD.GodownID"
tmp = tmp & " Where TGR.Flag = 'A' And  TGR.ExpenseFrom Between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group by TGRD.GodownID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.LeaseRentPaid) = IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)
      End If
      MSHFlexGrid1.TextMatrix(0, E.LeaseRentPaid) = Val(MSHFlexGrid1.TextMatrix(i, E.LeaseRentPaid)) + Val(MSHFlexGrid1.TextMatrix(0, E.LeaseRentPaid))
   Next
End If

'Commented on 30 Jan 2014 After discussion to encorportae arrears amount
'tmp = " Select TGRD.GodownID, sum(TGRD.Amount) As IncomeAmt "

tmp = " Select TGRD.GodownID, sum(TGRD.Amount)+ Sum(isnull(TGRD.Arrears,0)) As IncomeAmt "
tmp = tmp & " From Txn_GodownRentDetail TGRD"
tmp = tmp & " Inner Join Txn_GodownRent TGR on TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TGRD.GodownID"
tmp = tmp & " Where TGR.Flag = 'P' And  TGR.ExpenseFrom Between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group by TGRD.GodownID "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs1.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.LeaseRentFAG) = IIf(IsNull(TmpRs1!IncomeAmt), 0, TmpRs1!IncomeAmt)
      End If
      MSHFlexGrid1.TextMatrix(0, E.LeaseRentFAG) = Val(MSHFlexGrid1.TextMatrix(i, E.LeaseRentFAG)) + Val(MSHFlexGrid1.TextMatrix(0, E.LeaseRentFAG))
   Next
End If

tmp = " select TLAD.GodownID, MG.CLoseDate, TLA.StartDate , TLA.EXpiryDate, sum(TLAD.RentAmount) Amount  from Txn_GodownLessorAgreeDetail TLAD"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TLAD.GodownID"
tmp = tmp & " Inner Join Txn_LessorAgreement TLA on TLA.LAID = TLAD.LAID"
tmp = tmp & " Where TLA.ExpiryDate >= '" & mStartDate & "' and TLA.StartDate <= '" & mEndDate & "'  and  TLA.FLag in ('A', 'P')"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " And TLA.LAID not in"
tmp = tmp & " (Select LAID from Txn_GodownRentDetail TGRD"
tmp = tmp & " Inner Join Txn_GodownRent TGR on TGR.ClaimID = TGRD.ClaimID"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TGRD.GodownID"
tmp = tmp & " Where TGR.Flag in ('A', 'P') And"
tmp = tmp & " TGR.ExpenseFrom = '" & mStartDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'  )"
tmp = tmp & " Group by TLAD.GodownID, TLA.EXpiryDate, TLA.StartDate , MG.CLoseDate "
Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs1.MoveFirst
      TmpRs1.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs1.EOF Then
         If IsNull(TmpRs1!CloseDate) = True And Format(TmpRs1!StartDate, "yyyy-mm-dd") <= Format(mStartDate, "yyyy-mm-dd") Then
            MSHFlexGrid1.TextMatrix(i, E.LeaseRentDRC) = IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
         Else
            If Format(TmpRs1!ExpiryDate, "yyyy-mm-dd") > Format(TmpRs1!CloseDate, "yyyy-mm-dd") Then
               If Format(TmpRs1!ExpiryDate, "yyyy-mm-dd") >= Format(mEndDate, "yyyy-mm-dd") Then
                  If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mStartDate, "yyyy-mm-dd") Then
                     mdays = DateDiff("d", TmpRs1!StartDate, CDate(mEndDate)) + 1
                     MSHFlexGrid1.TextMatrix(i, E.LeaseRentDRC) = IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                  Else
                     MSHFlexGrid1.TextMatrix(i, E.LeaseRentDRC) = IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                  End If
               Else
                  If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mStartDate, "yyyy-mm-dd") Then
                     mdays = DateDiff("d", TmpRs1!StartDate, TmpRs1!ExpiryDate) + 1
                     MSHFlexGrid1.TextMatrix(i, E.LeaseRentDRC) = IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                  Else
                     mdays = DateDiff("d", CDate(mStartDate), TmpRs1!ExpiryDate) + 1
                     MSHFlexGrid1.TextMatrix(i, E.LeaseRentDRC) = IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                  End If
               End If
            Else
               If Format(TmpRs1!CloseDate, "yyyy-mm-dd") >= Format(mEndDate, "yyyy-mm-dd") Then
                  If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mStartDate, "yyyy-mm-dd") Then
                     mdays = DateDiff("d", TmpRs1!StartDate, CDate(mEndDate)) + 1
                     MSHFlexGrid1.TextMatrix(i, E.LeaseRentDRC) = IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                  Else
                     MSHFlexGrid1.TextMatrix(i, E.LeaseRentDRC) = IIf(IsNull(TmpRs1!Amount), 0, TmpRs1!Amount)
                  End If
               Else
                  If Format(TmpRs1!StartDate, "yyyy-mm-dd") > Format(mStartDate, "yyyy-mm-dd") Then
                     mdays = DateDiff("d", TmpRs1!StartDate, TmpRs1!CloseDate) + 1
                     MSHFlexGrid1.TextMatrix(i, E.LeaseRentDRC) = IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                  Else
                     mdays = DateDiff("d", CDate(mStartDate), TmpRs1!CloseDate) + 1
                     MSHFlexGrid1.TextMatrix(i, E.LeaseRentDRC) = IIf(IsNull(TmpRs1!Amount), 0, Format(TmpRs1!Amount * Val(mdays) / Val(mTotalDays), "#########.00"))
                  End If
               End If
            End If
         End If
      End If
      MSHFlexGrid1.TextMatrix(0, E.LeaseRentDRC) = Val(MSHFlexGrid1.TextMatrix(i, E.LeaseRentDRC)) + Val(MSHFlexGrid1.TextMatrix(0, E.LeaseRentDRC))
   Next
End If
End Sub

Private Sub GetFranchiseeExpense()
'tmp = " Select TIMG.GodownID, Sum(isnull(TFR.ActualFee * TIMG.FranchInvoiceAmount/TFR.Amount, TIMG.FranchInvoiceAmount)) IncomeAmt From Txn_InvMonGSLDetail TIMG"
tmp = " Select TIMG.GodownID, Sum(TFR.ActualFee * TIMG.FranchInvoiceAmount/TFR.Amount) IncomeAmt From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Inner Join Txn_FranchiseeRent TFR on TFR.ClaimID = TIMG.FranchClaimID And TFR.flag = 'A'"
tmp = tmp & " Where TIMG.FranchiseeRate is not null And TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " And TIMG.FranchInvoiceAmount > 0"
tmp = tmp & " Group By TIMG.GodownID"
Call Tmp3Table

If TmpRs3.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs3.MoveFirst
      TmpRs3.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs3.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.FranchiseeFeesPaid) = IIf(IsNull(TmpRs3!IncomeAmt), 0, Format(TmpRs3!IncomeAmt, "#############.00"))
      End If
      MSHFlexGrid1.TextMatrix(0, E.FranchiseeFeesPaid) = Val(MSHFlexGrid1.TextMatrix(i, E.FranchiseeFeesPaid)) + Val(MSHFlexGrid1.TextMatrix(0, E.FranchiseeFeesPaid))
   Next
End If

tmp = " Select TIMG.GodownID, Sum(isnull(TFR.ActualFee * TIMG.FranchInvoiceAmount/TFR.Amount, TIMG.FranchInvoiceAmount)) IncomeAmt From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Inner Join Txn_FranchiseeRent TFR on TFR.ClaimID = TIMG.FranchClaimID And TFR.flag = 'P'"
tmp = tmp & " Where TIMG.FranchiseeRate is not null And TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " And TIMG.FranchInvoiceAmount > 0"
tmp = tmp & " Group By TIMG.GodownID"
Call Tmp3Table

If TmpRs3.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs3.MoveFirst
      TmpRs3.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs3.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.FranchiseeFeesFAG) = IIf(IsNull(TmpRs3!IncomeAmt), 0, Format(TmpRs3!IncomeAmt, "#############.00"))
      End If
      MSHFlexGrid1.TextMatrix(0, E.FranchiseeFeesFAG) = Val(MSHFlexGrid1.TextMatrix(i, E.FranchiseeFeesFAG)) + Val(MSHFlexGrid1.TextMatrix(0, E.FranchiseeFeesFAG))
   Next
End If

tmp = " Select TIMG.GodownID, Sum(TIMG.FranchInvoiceAmount) IncomeAmt From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Where TIMG.FranchiseeRate is not null And TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " And TIMG.FranchInvoiceAmount > 0 And isnull(TIMG.FranchClaimID, 0) = 0"
tmp = tmp & " Group By TIMG.GodownID"
Call Tmp3Table

If TmpRs3.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs3.MoveFirst
      TmpRs3.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs3.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.FranchiseeFeesCMG) = IIf(IsNull(TmpRs3!IncomeAmt), 0, Format(TmpRs3!IncomeAmt, "#############.00"))
      End If
      MSHFlexGrid1.TextMatrix(0, E.FranchiseeFeesCMG) = Val(MSHFlexGrid1.TextMatrix(i, E.FranchiseeFeesCMG)) + Val(MSHFlexGrid1.TextMatrix(0, E.FranchiseeFeesCMG))
   Next
End If

End Sub

Private Sub GetSalaryExpenseDetail()
Call Grid_Head14

tmp = " select msd.locationid, me.employeeno, me.fullname, msd.amount from Txn_SalaryDistribution msd"
tmp = tmp & " inner join mst_employee me on me.employeeno = msd.employeeno"
tmp = tmp & " where msd.monthyear between '" & mStartDate & "' And '" & mEndDate & "' And SegmentID = 1"
tmp = tmp & " order by msd.locationid"
Call TmpTableHRPro

For k = 1 To TmpRsHRPro.RecordCount
   i = MSHFlexGrid14.Rows - 1
   TmpRs2.MoveFirst
   TmpRs2.Find "LocationID = " & TmpRsHRPro!LocationID
   If Not TmpRs.EOF Then
      TmpRs.Filter = ""
      TmpRs.Filter = "LocationID = " & TmpRsHRPro!LocationID
      For j = 1 To TmpRs.RecordCount
         MSHFlexGrid14.Rows = MSHFlexGrid14.Rows + 1
         MSHFlexGrid14.TextMatrix(i, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
         MSHFlexGrid14.TextMatrix(i, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         MSHFlexGrid14.TextMatrix(i, 2) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
         MSHFlexGrid14.TextMatrix(i, 3) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
         MSHFlexGrid14.TextMatrix(i, 4) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
         MSHFlexGrid14.TextMatrix(i, 5) = IIf(IsNull(TmpRsHRPro!FullName), "", TmpRsHRPro!FullName)
         If Opt_Stock.Value = True Then
            mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
            MSHFlexGrid1.TextMatrix(i, 6) = IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
         ElseIf Opt_Capacity.Value = True Then
            MSHFlexGrid14.TextMatrix(i, 6) = Format((TmpRsHRPro!Amount * TmpRs!GodownCapacity) / TmpRs2!GodownCapacity, "#########.00")
         End If
         MSHFlexGrid14.TextMatrix(0, 6) = Val(MSHFlexGrid14.TextMatrix(i, 6)) + Val(MSHFlexGrid14.TextMatrix(0, 6))
         i = i + 1
         TmpRs.MoveNext
      Next
   Else
      TmpRs.Filter = ""
        
      For j = 1 To TmpRs.RecordCount
         MSHFlexGrid14.Rows = MSHFlexGrid14.Rows + 1
         MSHFlexGrid14.TextMatrix(i, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
         MSHFlexGrid14.TextMatrix(i, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         MSHFlexGrid14.TextMatrix(i, 2) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
         MSHFlexGrid14.TextMatrix(i, 3) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
         MSHFlexGrid14.TextMatrix(i, 4) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
         MSHFlexGrid14.TextMatrix(i, 5) = IIf(IsNull(TmpRsHRPro!FullName), "", TmpRsHRPro!FullName)
         If Opt_Stock.Value = True Then
            MSHFlexGrid14.TextMatrix(i, 6) = Format(TmpRsHRPro!Amount / (MSHFlexGrid1.Rows - 3), "#########.00")
         ElseIf Opt_Capacity.Value = True Then
            MSHFlexGrid14.TextMatrix(i, 6) = Format((TmpRsHRPro!Amount * TmpRs!GodownCapacity) / MSHFlexGrid1.TextMatrix(0, E.GodownCapacity), "#########.00")
         End If
         MSHFlexGrid14.TextMatrix(0, 6) = Val(MSHFlexGrid14.TextMatrix(i, 6)) + Val(MSHFlexGrid14.TextMatrix(0, 6))
         i = i + 1
         TmpRs.MoveNext
      Next
   End If
   TmpRsHRPro.MoveNext
Next
End Sub

Private Sub GetSalaryExpense()
tmp = " Select LocationID, sum(isnull(BasicMonthly,0)) + sum(isnull(HRAMonthly,0)) + sum(isnull(ConveyanceMonthly,0)) + "
tmp = tmp & " sum(isnull(OtherAllMonthly,0)) + sum(isnull(MEDICAL,0)) + sum(isnull(SplAllMonthly,0)) + sum(isnull(LTA,0)) +"
tmp = tmp & " Sum (IsNull(LaptopAllowMonthly, 0)) + Sum(IsNull(LeaveEncashment, 0)) + Sum(IsNull(GainShare, 0)) + Sum(IsNull(Incentive, 0))"
tmp = tmp & " + sum(isnull(Perquisites,0)) + sum(isnull(ARREARSBASIC,0)) + sum(isnull(ARREARSHRA,0)) +"
tmp = tmp & " sum(isnull(ARREARSCONVEYANCE,0)) + sum(isnull(ARREARSMEDICAL,0)) + sum(isnull(PERFORMANCEBONUS,0)) +"
tmp = tmp & " Sum (IsNull(PROFESSIONALFEES, 0)) + Sum(IsNull(STIPEND, 0)) + Sum(IsNull(SERVICECHARGES, 0))"
tmp = tmp & " + sum(isnull(ACCRUEDVARIABLEPAY,0)) + sum(isnull(ACCRUEDRETIRALS,0))+sum(isnull(employersesic,0))"
tmp = tmp & " + sum(isnull(carmaintandpetrol,0))+ sum(isnull(booksandperiodical,0))+ sum(isnull(mealallowance,0))"
tmp = tmp & " + sum(isnull(LWFEMPLOYER,0)) + sum(isnull(EMPLOYERSPF,0)) + sum(isnull(STaxandEducationcess,0)) Amount "
tmp = tmp & " From Txn_SalaryDistribution  MSD"
tmp = tmp & " Where MSD.MonthYear Between  '" & mStartDate & "' And '" & mEndDate & "' And SegmentID = 1"
tmp = tmp & " Group by LocationID Order by LocationID"
Call TmpTableHRPro

If TmpRsHRPro.RecordCount > 0 Then
   For j = 1 To TmpRsHRPro.RecordCount
      mStatus = 0
      For i = 2 To MSHFlexGrid1.Rows - 2
'           If TmpRsHRPro!LocationID = 263 Then
'              a = a
'           End If
           
          If MSHFlexGrid1.TextMatrix(i, E.LocationID) = TmpRsHRPro!LocationID Then
             mStatus = 1
             TmpRs2.MoveFirst
             TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
             If Not TmpRs2.EOF Then
                If Opt_Stock.Value = True Then
                   mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
                   MSHFlexGrid1.TextMatrix(i, E.Salary) = Val(MSHFlexGrid1.TextMatrix(i, E.Salary)) + IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
                   MSHFlexGrid1.TextMatrix(0, E.Salary) = Val(MSHFlexGrid1.TextMatrix(0, E.Salary)) + IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
                ElseIf Opt_Capacity.Value = True Then
                   MSHFlexGrid1.TextMatrix(i, E.Salary) = Val(MSHFlexGrid1.TextMatrix(i, E.Salary)) + IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
                   MSHFlexGrid1.TextMatrix(0, E.Salary) = Val(MSHFlexGrid1.TextMatrix(0, E.Salary)) + IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
                End If
             End If
          End If
      Next
      If mStatus = 0 Then
         tmp = " Select MS.StateID, sum(MG.GodownCapacity) Capacity"
         tmp = tmp & " From Mst_Godown MG"
         tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
         tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
         tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
         tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
         tmp = tmp & " And MS.StateID in (Select StateID from Mst_Location Where LocationID = " & TmpRsHRPro!LocationID & " ) group by MS.StateID"
         Call Tmp3Table
         
         If TmpRs3.RecordCount = 0 Then
            tmp = " Select MS.StateID, sum(MG.GodownCapacity) Capacity"
            tmp = tmp & " From Mst_Godown MG"
            tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
            tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
            tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
            tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
            tmp = tmp & " And MS.RegionID in (Select MS.RegionID from Mst_State MS Inner Join Mst_Location ML on ML.StateID = MS.StateID Where ML.LocationID = " & TmpRsHRPro!LocationID & " ) group by MS.StateID"
            Call Tmp3Table
            
            tmp = " Select sum(MG.GodownCapacity) Capacity"
            tmp = tmp & " From Mst_Godown MG"
            tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
            tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
            tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
            tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
            tmp = tmp & " And MS.RegionID in (Select MS.RegionID from Mst_State MS Inner Join Mst_Location ML on ML.StateID = MS.StateID Where ML.LocationID = " & TmpRsHRPro!LocationID & " )"
            Call Tmp4Table
                
            For i = 2 To MSHFlexGrid1.Rows - 2
               TmpRs3.MoveFirst
               TmpRs3.Find "StateID = " & Val(MSHFlexGrid1.TextMatrix(i, E.StateID))
               If Not TmpRs3.EOF Then
                  MSHFlexGrid1.TextMatrix(i, E.Salary) = Val(MSHFlexGrid1.TextMatrix(i, E.Salary)) + IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs4!Capacity, "#############.00"))
                  MSHFlexGrid1.TextMatrix(0, E.Salary) = Val(MSHFlexGrid1.TextMatrix(0, E.Salary)) + IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs4!Capacity, "#############.00"))
               End If
            Next
         Else
            For i = 2 To MSHFlexGrid1.Rows - 2
'               If Val(MSHFlexGrid1.TextMatrix(i, E.StateID)) = 20 Then
'                  a = a
'               End If
               TmpRs3.MoveFirst
               TmpRs3.Find "StateID = " & Val(MSHFlexGrid1.TextMatrix(i, E.StateID))
               If Not TmpRs3.EOF Then
'                  If Val(MSHFlexGrid1.TextMatrix(i, E.StateID)) = 20 Then
'                      a = a
'                  End If
                  MSHFlexGrid1.TextMatrix(i, E.Salary) = Val(MSHFlexGrid1.TextMatrix(i, E.Salary)) + IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs3!Capacity, "#############.00"))
                  MSHFlexGrid1.TextMatrix(0, E.Salary) = Val(MSHFlexGrid1.TextMatrix(0, E.Salary)) + IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs3!Capacity, "#############.00"))
               End If
            Next
         End If
      End If
      TmpRsHRPro.MoveNext
   Next
End If
End Sub

Private Sub GetCoExpense()

mCountMonth = 1
tmp = " Select sum(ExpenseAmt * " & mCountMonth & ")/12 As IncomeAmt "
tmp = tmp & " From Mst_CoExpense MCE"
tmp = tmp & " Where '" & mStartDate & "' Between MCE.FromDate And MCE.Todate"
Call Tmp1Table
mTotalIncome = Val(MSHFlexGrid1.TextMatrix(0, E.NNNRIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(0, E.PledgeIncome)) '+ Val(MSHFlexGrid1.TextMatrix(0, E.Salary))
'mTotalIncome = Val(MSHFlexGrid1.TextMatrix(0, E.TotalIncome)) '+ Val(MSHFlexGrid1.TextMatrix(0, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(0, E.PledgeIncome)) '+ Val(MSHFlexGrid1.TextMatrix(0, E.Salary))

If TmpRs1.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      mGodownIncome = Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.PledgeIncome)) '+ Val(MSHFlexGrid1.TextMatrix(i, E.Salary))
      'mGodownIncome = Val(MSHFlexGrid1.TextMatrix(i, E.TotalIncome))
      MSHFlexGrid1.TextMatrix(i, E.CoExpense) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format((TmpRs1!IncomeAmt * mGodownIncome) / mTotalIncome, "#############.00"))
      MSHFlexGrid1.TextMatrix(0, E.CoExpense) = Val(MSHFlexGrid1.TextMatrix(i, E.CoExpense)) + Val(MSHFlexGrid1.TextMatrix(0, E.CoExpense))
   Next
End If

End Sub

Private Sub GetDPIncome()
'''===================Previous Query for finding DP Income=================================
'tmp = " Select MDP.LocationID, sum(MDP.Amt) IncomeAmt "
'tmp = tmp & " from Mst_DPInvoiceUpload MDP"
'tmp = tmp & " Where isnull(MDP.ReverseFlag,'')<>'Y' And MDP.MonthYear Between '" & mStartDate & "' And '" & mEndDate & "'" '''MDP.ReverseFlag = 'N' And
'tmp = tmp & " And MDP.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
'tmp = tmp & " Group By MDP.LocationID "
'Call Tmp1Table
'
'If TmpRs1.RecordCount > 0 Then
'   For I = 2 To MSHFlexGrid1.Rows - 2
'      TmpRs1.MoveFirst
'      TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(I, E.LocationID))
'      If Not TmpRs1.EOF Then
'         TmpRs2.MoveFirst
'         TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(I, E.LocationID))
'         If Not TmpRs2.EOF Then
'            If Opt_Stock.Value = True Then
'               mStock = Val(MSHFlexGrid1.TextMatrix(I, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(I, E.NonNCDEXStock))
'               MSHFlexGrid1.TextMatrix(I, E.DPIncomeasperInvoice) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(mStock) / TmpRs2!BalanceWeight, "#############.00"))
'            ElseIf Opt_Capacity.Value = True Then
'               MSHFlexGrid1.TextMatrix(I, E.DPIncomeasperInvoice) = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt * Val(MSHFlexGrid1.TextMatrix(I, E.GodownCapacity)) / TmpRs2!Capacity, "#############.00"))
'            End If
'         End If
'      End If
'      MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice) = Val(MSHFlexGrid1.TextMatrix(I, E.DPIncomeasperInvoice)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice))
'   Next
'End If

''commented when added commtrackamount on 1 july 2013
'tmp = " Select TIMG.GodownID, Sum(TIMG.DPIncome) IncomeAmt From Txn_InvMonGSLDetail TIMG"

tmp = " Select Sum(isnull(TIMG.DPIncome,0))'DPIncome',sum(isnull(TIMG.CommtrackIncome,0)) 'CommIncome' From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Where TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And TIMG.ExchangeTypeID = 2 And TIMG.BalanceWeight > 0"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"

Call Tmp3Table

If TmpRs3!DPIncome > 0 And TmpRs3!CommIncome > 0 Then

   tmp = " Select TIMG.GodownID, Sum(isnull(TIMG.DPIncome,0))+sum(isnull(TIMG.CommtrackIncome,0)) IncomeAmt From Txn_InvMonGSLDetail TIMG"
   tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
   tmp = tmp & " Where TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
   tmp = tmp & " And TIMG.ExchangeTypeID = 2 And TIMG.BalanceWeight > 0"
   tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
'tmp = tmp & " And (TIMG.DPIncome > 0 Or TIMG.CommtrackIncome > 0)"
   tmp = tmp & " Group By TIMG.GodownID"
   Call Tmp3Table

   If TmpRs3.RecordCount > 0 Then
      For i = 2 To MSHFlexGrid1.Rows - 2
          TmpRs3.MoveFirst
          TmpRs3.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
          If Not TmpRs3.EOF Then
             MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice) = IIf(IsNull(TmpRs3!IncomeAmt), 0, Format(TmpRs3!IncomeAmt, "#############.00"))
          End If
          MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice))
      Next
   End If


   tmp = " select distinct sm_prefix from txn_invmongsldetail"
   tmp = tmp & " where exchangetypeid = 2 and isnull(InvoiceAmount,0) > 0 and yearmonth between '" & mStartDate & "' and '" & mEndDate & "'"
   tmp = tmp & " and isnull(commtrackflag,0) = 0"
   Call Tmp3Table
   mSM_Prefix = "(''"
   For i = 1 To TmpRs3.RecordCount
      mSM_Prefix = mSM_Prefix & " ,'" & IIf(IsNull(TmpRs3!SM_Prefix), "", TmpRs3!SM_Prefix) & "'"
      TmpRs3.MoveNext
   Next
   mSM_Prefix = mSM_Prefix & " )"
   tmp = " select mdp.locationid, sum(TotalAmt) IncomeAmt from Mst_DPInvoiceUpload mdp"
   tmp = tmp & " inner join mst_location ml on ml.locationid = mdp.locationid"
   tmp = tmp & " where MDP.MonthYear between '" & mStartDate & "' And '" & mEndDate & "'"
   tmp = tmp & " and ml.sm_prefix not in " & mSM_Prefix
   tmp = tmp & " group by mdp.locationid"
   Call Tmp3Table
   
   tmp = " Select MCMP.LocationID, sum(TIMG.BalanceWeight) BalanceWeight"
   tmp = tmp & " From Txn_InvMonGSLDetail TIMG"
   tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
   tmp = tmp & " Where TIMG.ExchangeTypeID in (1, 2) And TIMG.BalanceWeight > 0 And TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
   tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
   tmp = tmp & " Group by MCMP.LocationID"
   Call Tmp1Table
    
    
    If TmpRs3.RecordCount > 0 Then
       For i = 2 To MSHFlexGrid1.Rows - 2
          TmpRs3.MoveFirst
          TmpRs3.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
          If Not TmpRs3.EOF Then
             TmpRs1.MoveFirst
             TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
             If Not TmpRs1.EOF Then
                mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
                MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice)) + IIf(IsNull(TmpRs3!IncomeAmt), 0, Format(TmpRs3!IncomeAmt * Val(mStock) / TmpRs1!BalanceWeight, "#############.00"))
             End If
          End If
          MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice))
       Next
    End If

ElseIf TmpRs3!DPIncome > 0 And TmpRs3!CommIncome = 0 Then

    tmp = " Select TIMG.GodownID, Sum(isnull(TIMG.DPIncome,0)) IncomeAmt From Txn_InvMonGSLDetail TIMG"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
    tmp = tmp & " Where TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
    tmp = tmp & " And TIMG.ExchangeTypeID = 2 And TIMG.BalanceWeight > 0"
    tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
    tmp = tmp & " And TIMG.CommtrackFlag = '0' "
    tmp = tmp & " Group By TIMG.GodownID"
    Call Tmp3Table

    If TmpRs3.RecordCount > 0 Then
       For i = 2 To MSHFlexGrid1.Rows - 2
           TmpRs3.MoveFirst
           TmpRs3.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
           If Not TmpRs3.EOF Then
              MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice) = IIf(IsNull(TmpRs3!IncomeAmt), 0, Format(TmpRs3!IncomeAmt, "#############.00"))
           End If
           MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice))
       Next
    End If


    tmp = " Select TIMG.GodownID, Sum(isnull(TIMG.InvoiceAmount,0)) IncomeAmt From Txn_InvMonGSLDetail TIMG"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
    tmp = tmp & " Where TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
    tmp = tmp & " And TIMG.ExchangeTypeID = 2 And TIMG.BalanceWeight > 0"
    tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
    tmp = tmp & " And TIMG.CommtrackFLag = '1' And ISINID not in (Select Distinct ISINID from Mst_RRN Where CommtrackFlag = '0')"
    tmp = tmp & " Group By TIMG.GodownID"
    Call Tmp3Table

    If TmpRs3.RecordCount > 0 Then
       For i = 2 To MSHFlexGrid1.Rows - 2
           TmpRs3.MoveFirst
           TmpRs3.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
           If Not TmpRs3.EOF Then
              MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice)) + IIf(IsNull(TmpRs3!IncomeAmt), 0, Format(TmpRs3!IncomeAmt, "#############.00"))
           End If
           MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice))
       Next
    End If

   tmp = " select distinct sm_prefix from txn_invmongsldetail"
   tmp = tmp & " where exchangetypeid = 2 and isnull(InvoiceAmount,0) > 0 and yearmonth between '" & mStartDate & "' and '" & mEndDate & "'"
   tmp = tmp & " and isnull(commtrackflag,0) = 0"
   Call Tmp3Table
   mSM_Prefix = "(''"
   For i = 1 To TmpRs3.RecordCount
      mSM_Prefix = mSM_Prefix & " ,'" & IIf(IsNull(TmpRs3!SM_Prefix), "", TmpRs3!SM_Prefix) & "'"
      TmpRs3.MoveNext
   Next
   mSM_Prefix = mSM_Prefix & " )"
   tmp = " select mdp.locationid, sum(TotalAmt) IncomeAmt from Mst_DPInvoiceUpload mdp"
   tmp = tmp & " inner join mst_location ml on ml.locationid = mdp.locationid"
   tmp = tmp & " where MDP.MonthYear between '" & mStartDate & "' And '" & mEndDate & "'"
   tmp = tmp & " and ml.sm_prefix not in " & mSM_Prefix
   tmp = tmp & " group by mdp.locationid"
   Call Tmp3Table
   
   tmp = " Select MCMP.LocationID, sum(TIMG.BalanceWeight) BalanceWeight"
   tmp = tmp & " From Txn_InvMonGSLDetail TIMG"
   tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
   tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
   tmp = tmp & " Where TIMG.ExchangeTypeID in (1, 2) And TIMG.BalanceWeight > 0 And TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
   tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
   tmp = tmp & " Group by MCMP.LocationID"
   Call Tmp1Table
    
    
    If TmpRs3.RecordCount > 0 Then
       For i = 2 To MSHFlexGrid1.Rows - 2
          TmpRs3.MoveFirst
          TmpRs3.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
          If Not TmpRs3.EOF Then
             TmpRs1.MoveFirst
             TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
             If Not TmpRs1.EOF Then
                mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
                MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice)) + IIf(IsNull(TmpRs3!IncomeAmt), 0, Format(TmpRs3!IncomeAmt * Val(mStock) / TmpRs1!BalanceWeight, "#############.00"))
             End If
          End If
          MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice))
       Next
    End If


ElseIf TmpRs3!DPIncome = 0 And TmpRs3!CommIncome > 0 Then

    tmp = " Select TIMG.GodownID, Sum(isnull(TIMG.CommtrackIncome,0)) IncomeAmt From Txn_InvMonGSLDetail TIMG"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
    tmp = tmp & " Where TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
    tmp = tmp & " And TIMG.ExchangeTypeID = 2 And TIMG.BalanceWeight > 0"
    tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
    tmp = tmp & " And TIMG.CommtrackFlag = '1' "
    tmp = tmp & " Group By TIMG.GodownID"
    Call Tmp3Table

    If TmpRs3.RecordCount > 0 Then
       For i = 2 To MSHFlexGrid1.Rows - 2
           TmpRs3.MoveFirst
           TmpRs3.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
           If Not TmpRs3.EOF Then
              MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice) = IIf(IsNull(TmpRs3!IncomeAmt), 0, Format(TmpRs3!IncomeAmt, "#############.00"))
           End If
           MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice))
       Next
    End If


    tmp = " Select TIMG.GodownID, Sum(isnull(TIMG.InvoiceAmount,0)) IncomeAmt From Txn_InvMonGSLDetail TIMG"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
    tmp = tmp & " Where TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
    tmp = tmp & " And TIMG.ExchangeTypeID = 2 And TIMG.BalanceWeight > 0"
    tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
    tmp = tmp & " And TIMG.CommtrackFLag = '0'"
    tmp = tmp & " Group By TIMG.GodownID"
    Call Tmp3Table

    If TmpRs3.RecordCount > 0 Then
       For i = 2 To MSHFlexGrid1.Rows - 2
           TmpRs3.MoveFirst
           TmpRs3.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
           If Not TmpRs3.EOF Then
              MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice)) + IIf(IsNull(TmpRs3!IncomeAmt), 0, Format(TmpRs3!IncomeAmt, "#############.00"))
           End If
           MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice))
       Next
    End If


ElseIf TmpRs3!DPIncome = 0 And TmpRs3!CommIncome = 0 Then
    tmp = " Select TIMG.GodownID, Sum(isnull(TIMG.InvoiceAmount,0)) IncomeAmt From Txn_InvMonGSLDetail TIMG"
    tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
    tmp = tmp & " Where TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
    tmp = tmp & " And TIMG.ExchangeTypeID = 2 And TIMG.BalanceWeight > 0"
    tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
    tmp = tmp & " Group By TIMG.GodownID"
    Call Tmp3Table
    
    If TmpRs3.RecordCount > 0 Then
       For i = 2 To MSHFlexGrid1.Rows - 2
           TmpRs3.MoveFirst
           TmpRs3.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
           If Not TmpRs3.EOF Then
              MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice) = IIf(IsNull(TmpRs3!IncomeAmt), 0, Format(TmpRs3!IncomeAmt, "#############.00"))
           End If
           MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperInvoice))
       Next
    End If

End If


'''===================Previous Query for finding DP Income=================================

tmp = " Select TIMG.GodownID, Sum(TIMG.InvoiceAmount) IncomeAmt From Txn_InvMonGSLDetail TIMG"
tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
tmp = tmp & " Where TIMG.YearMonth Between '" & mStartDate & "' and '" & mEndDate & "'"
tmp = tmp & " And TIMG.ExchangeTypeID = 2 And TIMG.BalanceWeight > 0"
tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
tmp = tmp & " Group By TIMG.GodownID"

Call Tmp3Table

If TmpRs3.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      TmpRs3.MoveFirst
      TmpRs3.Find "GodownID = " & Val(MSHFlexGrid1.TextMatrix(i, E.GodownID))
      If Not TmpRs3.EOF Then
         MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperGSL) = IIf(IsNull(TmpRs3!IncomeAmt), 0, Format(TmpRs3!IncomeAmt, "#############.00"))
      End If
      MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperGSL) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperGSL)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperGSL))
   Next
End If

'tmp = " Select MCMP.LocationID, sum(TIMG.BalanceWeight) BalanceWeight,  Sum(TIMG.InvoiceAmount) IncomeAmt From Txn_InvMonGSLDetail TIMG"
'tmp = tmp & " Inner Join Mst_CMP MCMP on TIMG.CMPID = MCMP.CMPID"
'tmp = tmp & " Inner Join Mst_Godown MG on MG.GodownID = TIMG.GodownID"
'tmp = tmp & " Where TIMG.YearMonth  Between '" & mStartDate & "' and '" & mEndDate & "'"
'tmp = tmp & " And TIMG.ExchangeTypeID = 2 And TIMG.BalanceWeight > 0"
'tmp = tmp & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
'tmp = tmp & " Group By MCMP.LocationID"
'
'Call Tmp4Table
'
''''''For Reversal==========================================
'tmp = " Select TCA.LocationID, sum(RecOtherDed) RecOtherDed" 'TCA.LocationID,
'tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCAD"
'tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCA on TCA.TDSAID = TCAD.TDSAID And TCA.SM_Prefix = TCAD.SM_Prefix And TCA.LocationID in (Select MCMP.LocationID from Mst_CMP MCMP Inner Join MSt_Godown MG on MG.CMPID = MCMP.CMPID And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "' )"
'tmp = tmp & " Where TCA.Flag in ('A', 'P') And AdjDate Between '" & mStartDate & "' And '" & mEndDate & "'"
'tmp = tmp & " And TCAD.ExpenseAccID  = 89 " '89 means DP Reversal"
'tmp = tmp & " group by TCA.LocationID"
'Call Tmp5Table
''''''For Reversal==========================================
'
'If TmpRs4.RecordCount > 0 Then
'   For I = 2 To MSHFlexGrid1.Rows - 2
'      mDPIncome = 0
'      mGSLDPIncome = 0
'      mLocationAmount = 0
'      TmpRs4.MoveFirst
'      TmpRs4.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(I, E.LocationID))
'      If Not TmpRs4.EOF Then
'         mGSLDPIncome = IIf(IsNull(TmpRs4!IncomeAmt), 0, Format(TmpRs4!IncomeAmt, "#############.00"))
'         mLocationAmount = IIf(IsNull(TmpRs4!IncomeAmt), 0, TmpRs4!IncomeAmt)
'         '''''For Reversal==========================================
'         If TmpRs5.RecordCount > 0 Then
'            TmpRs5.MoveFirst
'            TmpRs5.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(I, E.LocationID))
'            If Not TmpRs5.EOF Then
'               mRecOtherDed = Format(Val(Val(IIf(IsNull(TmpRs5!RecOtherDed), "", TmpRs5!RecOtherDed)) * Val(MSHFlexGrid1.TextMatrix(I, E.DPIncomeasperGSL))) / TmpRs4!IncomeAmt, "########.00")
'               MSHFlexGrid1.TextMatrix(I, E.DPIncomeasperGSL) = Val(MSHFlexGrid1.TextMatrix(I, E.DPIncomeasperGSL)) - Val(mRecOtherDed)
'               MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperGSL) = Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeasperGSL)) - Val(mRecOtherDed)
'               mGSLDPIncome = IIf(IsNull(TmpRs4!IncomeAmt), 0, Format(TmpRs4!IncomeAmt, "#############.00")) - IIf(IsNull(TmpRs5!RecOtherDed), 0, TmpRs5!RecOtherDed)
'               mLocationAmount = IIf(IsNull(TmpRs4!IncomeAmt), 0, TmpRs4!IncomeAmt) - IIf(IsNull(TmpRs5!RecOtherDed), 0, TmpRs5!RecOtherDed)
'            End If
'         End If
'         '''''For Reversal==========================================
'      End If
'      If TmpRs1.RecordCount > 0 Then
'         TmpRs1.MoveFirst
'         TmpRs1.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(I, E.LocationID))
'         If Not TmpRs1.EOF Then
'            mDPIncome = IIf(IsNull(TmpRs1!IncomeAmt), 0, Format(TmpRs1!IncomeAmt, "#############.00"))
'         End If
'      End If
'      mDifference = Val(mGSLDPIncome) - Val(mDPIncome)
'
'      If Val(mLocationAmount) > 0 Then
'         MSHFlexGrid1.TextMatrix(I, E.Difference) = Format((Val(mDifference) * Val(MSHFlexGrid1.TextMatrix(I, E.DPIncomeasperGSL))) / Val(mLocationAmount), "#############.00")
'      ElseIf Val(MSHFlexGrid1.TextMatrix(I, E.DPIncomeasperInvoice)) > 0 Then
'         MSHFlexGrid1.TextMatrix(I, E.Difference) = -Val(MSHFlexGrid1.TextMatrix(I, E.DPIncomeasperInvoice))
'      End If
'      MSHFlexGrid1.TextMatrix(0, E.Difference) = Val(MSHFlexGrid1.TextMatrix(I, E.Difference)) + Val(MSHFlexGrid1.TextMatrix(0, E.Difference))
'   Next
'End If
'
'For I = 2 To MSHFlexGrid1.Rows - 2
'    MSHFlexGrid1.TextMatrix(I, E.FinalDPIncome) = Val(MSHFlexGrid1.TextMatrix(I, E.DPIncomeasperGSL)) - Val(MSHFlexGrid1.TextMatrix(I, E.Difference))
'    MSHFlexGrid1.TextMatrix(I, E.NCDEXCharges) = Format(Val(MSHFlexGrid1.TextMatrix(I, E.FinalDPIncome)) * 0.02, "###########.00")
'    MSHFlexGrid1.TextMatrix(0, E.FinalDPIncome) = Val(MSHFlexGrid1.TextMatrix(I, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.FinalDPIncome))
'    MSHFlexGrid1.TextMatrix(0, E.NCDEXCharges) = Val(MSHFlexGrid1.TextMatrix(I, E.NCDEXCharges)) + Val(MSHFlexGrid1.TextMatrix(0, E.NCDEXCharges))
'Next





For i = 2 To MSHFlexGrid1.Rows - 2
    MSHFlexGrid1.TextMatrix(i, E.FinalDPIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeasperInvoice))
    MSHFlexGrid1.TextMatrix(i, E.NCDEXCharges) = Format(Val(MSHFlexGrid1.TextMatrix(i, E.FinalDPIncome)) * 0.02, "###########.00")
    MSHFlexGrid1.TextMatrix(0, E.FinalDPIncome) = Val(MSHFlexGrid1.TextMatrix(i, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.FinalDPIncome))
    MSHFlexGrid1.TextMatrix(0, E.NCDEXCharges) = Val(MSHFlexGrid1.TextMatrix(i, E.NCDEXCharges)) + Val(MSHFlexGrid1.TextMatrix(0, E.NCDEXCharges))
Next

End Sub

Private Sub Grid_Head2()
With MSHFlexGrid2
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 10
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "CMP Name"
   .TextMatrix(1, 3) = "Godown No"
   .TextMatrix(1, 4) = "Godown Address"
   .TextMatrix(1, 5) = "Stack No"
   .TextMatrix(1, 6) = "Lot No"
   .TextMatrix(1, 7) = "Client Name"
   .TextMatrix(1, 8) = "Exchange Type"
   .TextMatrix(1, 9) = "Balance Weight"
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 3500
   .ColWidth(3) = 1500
   .ColWidth(4) = 3500
   .ColWidth(5) = 1800
   .ColWidth(6) = 1800
   .ColWidth(7) = 2500
   .ColWidth(8) = 2000
   .ColWidth(9) = 1900
   
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head3()
With MSHFlexGrid3
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 10
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "CMP Name"
   .TextMatrix(1, 3) = "Godown No"
   .TextMatrix(1, 4) = "Godown Address"
   .TextMatrix(1, 5) = "Stack No"
   .TextMatrix(1, 6) = "Lot No"
   .TextMatrix(1, 7) = "Commodity"
   .TextMatrix(1, 8) = "Client Name"
   .TextMatrix(1, 9) = "Income Amount"
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 3500
   .ColWidth(3) = 1500
   .ColWidth(4) = 3500
   .ColWidth(5) = 1500
   .ColWidth(6) = 1500
   .ColWidth(7) = 2200
   .ColWidth(8) = 3500
   .ColWidth(9) = 1500
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head4()
With MSHFlexGrid4
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 5 '12
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "Invoice No"
   .TextMatrix(1, 3) = "Billing Date"
   .TextMatrix(1, 4) = "Income Amount"
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 2000
   .ColWidth(3) = 1800
   .ColWidth(4) = 1500
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head5()
With MSHFlexGrid5
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 10
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "Exchange Type"
   .TextMatrix(1, 3) = "Commodity"
   .TextMatrix(1, 4) = "Client Name"
   .TextMatrix(1, 5) = "Invoice Type"
   .TextMatrix(1, 6) = "PI No"
   .TextMatrix(1, 7) = "PI Date"
   .TextMatrix(1, 8) = "Invoice Date"
   .TextMatrix(1, 9) = "Income Amount"
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 1900
   .ColWidth(3) = 1900
   .ColWidth(4) = 3500
   .ColWidth(5) = 1800
   .ColWidth(6) = 1800
   .ColWidth(7) = 1800
   .ColWidth(8) = 1800
   .ColWidth(9) = 2000
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head6()
With MSHFlexGrid6
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 7
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "DPInvoice ID"
   .TextMatrix(1, 3) = "PINo"
   .TextMatrix(1, 4) = "Month Year"
   .TextMatrix(1, 5) = "DP NavCode"
   .TextMatrix(1, 6) = "Amount"
   
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 1900
   .ColWidth(3) = 1800
   .ColWidth(4) = 3500
   .ColWidth(5) = 1800
   .ColWidth(6) = 1800
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head7()
With MSHFlexGrid7
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 12
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "PCRID"
   .TextMatrix(1, 3) = "Instrument Type"
   .TextMatrix(1, 4) = "Drawn On"
   .TextMatrix(1, 5) = "Branch Name"
   .TextMatrix(1, 6) = "Instrument No"
   .TextMatrix(1, 7) = "Instrument Date"
   .TextMatrix(1, 8) = "Received From"
   .TextMatrix(1, 9) = "Excess Amount"
   .TextMatrix(1, 10) = "Remark"
   .TextMatrix(1, 11) = "Income Amount"
   
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 1900
   .ColWidth(3) = 1800
   .ColWidth(4) = 2200
   .ColWidth(5) = 1800
   .ColWidth(6) = 1800
   .ColWidth(7) = 1800
   .ColWidth(8) = 1800
   .ColWidth(9) = 1800
   .ColWidth(10) = 1800
   .ColWidth(11) = 1800
   
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head8()
With MSHFlexGrid8
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 6
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "Expense Type"
   .TextMatrix(1, 3) = "Particulars"
   .TextMatrix(1, 4) = "Amount"
   .TextMatrix(1, 5) = "Flag"
   
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 1900
   .ColWidth(3) = 1800
   .ColWidth(4) = 1800
   .ColWidth(5) = 1800
   
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head9()
With MSHFlexGrid9
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 12
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "BillTxn Id"
   .TextMatrix(1, 3) = "Month Year"
   .TextMatrix(1, 4) = "Agency Name"
   .TextMatrix(1, 5) = "Security Type"
   .TextMatrix(1, 6) = "Security Designation"
   .TextMatrix(1, 7) = "No Of Godown"
   .TextMatrix(1, 8) = "Rate Per Month"
   .TextMatrix(1, 9) = "Description"
   .TextMatrix(1, 10) = "Amount"
   .TextMatrix(1, 11) = "Flag"
   
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 1900
   .ColWidth(3) = 1800
   .ColWidth(4) = 2500
   .ColWidth(5) = 2500
   .ColWidth(6) = 2500
   .ColWidth(7) = 1500
   .ColWidth(8) = 1800
   .ColWidth(9) = 2500
   .ColWidth(10) = 1800
   .ColWidth(11) = 1000
   
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head10()
With MSHFlexGrid10
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 9
   
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "CMP Name"
   .TextMatrix(1, 3) = "Godown No"
   .TextMatrix(1, 4) = "Start Date"
   .TextMatrix(1, 5) = "Close Date"
   .TextMatrix(1, 6) = "Address"
   .TextMatrix(1, 7) = "GodownID"
   .TextMatrix(1, 8) = "Amount"
     
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 1900
   .ColWidth(3) = 1800
   .ColWidth(4) = 2500
   .ColWidth(5) = 2500
   .ColWidth(6) = 2500
   .ColWidth(7) = 1500
   .ColWidth(8) = 1800
   
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head11()
With MSHFlexGrid11
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 14
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "CMP Name"
   .TextMatrix(1, 3) = "Godown No"
   .TextMatrix(1, 4) = "Address"
   .TextMatrix(1, 5) = "Client Name"
   .TextMatrix(1, 6) = "Commodity"
   .TextMatrix(1, 7) = "Stack No"
   .TextMatrix(1, 8) = "Lot No"
   .TextMatrix(1, 9) = "Balance Bags"
   .TextMatrix(1, 10) = "Balance Weight"
   .TextMatrix(1, 11) = "DPIncome as per GSL"
   .TextMatrix(1, 12) = "DPIncome as per NCDEX"
   .TextMatrix(1, 13) = "Commtrack Income"
   
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 1900
   .ColWidth(3) = 1800
   .ColWidth(4) = 2500
   .ColWidth(5) = 2500
   .ColWidth(6) = 2500
   .ColWidth(7) = 1500
   .ColWidth(8) = 1500
   .ColWidth(9) = 1500
   .ColWidth(10) = 1500
   .ColWidth(11) = 1800
   
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head12()
With MSHFlexGrid12
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 17
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "CMP Name"
   .TextMatrix(1, 3) = "Godown No"
   .TextMatrix(1, 4) = "Address"
   .TextMatrix(1, 5) = "Exchange Type"
   .TextMatrix(1, 6) = "Balance Bags"
   .TextMatrix(1, 7) = "Balance Weight"
   .TextMatrix(1, 8) = "Franchisee Rate"
   .TextMatrix(1, 9) = "Billing Cycle"
   .TextMatrix(1, 10) = "Billing Unit"
   .TextMatrix(1, 11) = "No of Days"
   .TextMatrix(1, 12) = "GSLID"
   .TextMatrix(1, 13) = "Stack No"
   .TextMatrix(1, 14) = "Lot No"
   .TextMatrix(1, 15) = "Franchisee Invoice Amount"
   .TextMatrix(1, 16) = "Claim Flag"
       
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 1900
   .ColWidth(3) = 1800
   .ColWidth(4) = 2500
   .ColWidth(5) = 2500
   .ColWidth(6) = 2500
   .ColWidth(7) = 1500
   .ColWidth(8) = 1500
   .ColWidth(9) = 1500
   .ColWidth(10) = 1500
   .ColWidth(11) = 1500
   .ColWidth(12) = 1500
   
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head()
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4
Call Grid_Head5
Call Grid_Head6
Call Grid_Head7
Call Grid_Head8
Call Grid_Head9
Call Grid_Head10
Call Grid_Head11
Call Grid_Head12
Call Grid_Head13
Call Grid_Head14
Call Grid_Head15
Call Grid_Head16
Call Grid_Head17

Label5.Caption = ""
ProgressBar1.Value = 0
Cmd_Process.Enabled = False
CmdExcel.Enabled = False
With MSHFlexGrid1
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 60
   
   .TextMatrix(1, E.MonthYear) = "Month Year"
   .TextMatrix(1, E.StateName) = "State Name"
   .TextMatrix(1, E.LocationName) = "Location Name"
   .TextMatrix(1, E.AMname) = "AM Name"
   .TextMatrix(1, E.RMName) = "RM Name"
   .TextMatrix(1, E.CMPName) = "CMP Name"
   .TextMatrix(1, E.GodownNo) = "Godown No"
   .TextMatrix(1, E.GodownCapacity) = "Godown Capacity"
   .TextMatrix(1, E.NcdexStock) = "NCDEX Stock "
   .TextMatrix(1, E.NonNCDEXStock) = "Non NCDEX Stock"
   .TextMatrix(1, E.NNNRIncome) = "NNNR Income "
   .TextMatrix(1, E.ReservationIncome) = "Reservation Income "
   .TextMatrix(1, E.SevenDaysIncome) = "Seven Days Income "
   .TextMatrix(1, E.FinalDPIncome) = "Final DP Income for the Godown" ''"DP Income "
   .TextMatrix(1, E.DPIncomeasperGSL) = "DP Income as per GSL" '"DP Invoice Amount From GSL Detail"
   .TextMatrix(1, E.Difference) = "Difference between DP (GSL - Invoice)" '"Difference Distribution"
   .TextMatrix(1, E.DPIncomeasperInvoice) = "DP Income as per DP Invoice" '"Amount After Distribution"
   .TextMatrix(1, E.RejectedIncome) = "Rejected Income " '" NCDEX WithDrawal Income "
   .TextMatrix(1, E.OtherIncome) = "Other Income "
   .TextMatrix(1, E.SCMIncome) = "SCM Income "
   .TextMatrix(1, E.LeftOverStock) = "Left over Stocks"
   .TextMatrix(1, E.PledgeIncome) = "Pledge Income"
   .TextMatrix(1, E.EmpClaimFAG) = "Employee Claims - approved by FAG"
   .TextMatrix(1, E.VendorExpenses) = "Vendor Expenses Paid"
   .TextMatrix(1, E.VendorExpensesAccrued) = "Vendor Expenses Accrued"
   .TextMatrix(1, E.SecurityFeesPaid) = "Security Fees - paid "
   .TextMatrix(1, E.FumigationFeesVendor) = "Fumigation Fees thru Vendor module - paid"
   .TextMatrix(1, E.LeaseRentPaid) = "Lease Rent Paid"
   .TextMatrix(1, E.LeaseRentFAG) = "Lease Rent - pending with FAG"
   .TextMatrix(1, E.LeaseRentDRC) = "Lease Rent - pending with DRC"
   .TextMatrix(1, E.FranchiseeFeesPaid) = "Franchisee Fees - Paid by FAG"
   .TextMatrix(1, E.FranchiseeFeesFAG) = "Franchisee Fees - Processed but not paid by FAG"
   .TextMatrix(1, E.FranchiseeFeesCMG) = "Franchisee Fees - Unprocessed by CMG"
   .TextMatrix(1, E.Salary) = "Salary Others"
   .TextMatrix(1, E.SalaryFieldStaff) = "Salary Field Staff"
   .TextMatrix(1, E.CoExpense) = "CO Expenses - Allocated"
   .TextMatrix(1, E.NCDEXCharges) = "NCDEX Charges - 2% of DP Fees"
   .TextMatrix(1, E.GodownStartDate) = "Godown Start Date"
   .TextMatrix(1, E.GodownCloseDate) = "Godown Close Date"
   .TextMatrix(1, E.UnitType) = "Unit Type"
   .TextMatrix(1, E.StateID) = "StateID"
   .TextMatrix(1, E.LocationID) = "LocationID"
   .TextMatrix(1, E.GodownID) = "GodownID"
   .TextMatrix(1, E.SecurityFeesAccrual) = "Security Fees - Accrued"
   .TextMatrix(1, E.FumigationFeesPaid) = "Fumigation & Spraying Fees - paid"
   .TextMatrix(1, E.FumigationFeesAccrual) = "Fumigation & Spraying Fees - accrued"
   .TextMatrix(1, E.InsurancePremium) = "Insurance Premium"
   .TextMatrix(1, E.InterestCharges) = "Interest Charges"
   .TextMatrix(1, E.EmpClaimAccrual) = "Employee Claim - accrued"
   .TextMatrix(1, E.CMUnbilledIncome) = "Income from CM Unbilled"
   .TextMatrix(1, E.CMPendingIncome) = "Income from CM Pending"
   .TextMatrix(1, E.CMApprovedIncome) = "Income from CM Approved"
   .TextMatrix(1, E.ReservationIncomeAdhoc) = "Reservation Income through Ad-Hoc"
   .TextMatrix(1, E.DPIncomeAdHoc) = "DP Income through Ad-Hoc"
   .TextMatrix(1, E.NNNRIncomeAdhoc) = "NNNR Income through Ad-Hoc"
   .TextMatrix(1, E.SpotIncomeAdHoc) = "Spot Income through Ad-Hoc"
   .TextMatrix(1, E.TotalIncome) = "Total Income"
   .TextMatrix(1, E.TotalExpense) = "Total Expense"
   .TextMatrix(1, E.GrossIncome) = "Gross Income"
   .TextMatrix(1, E.NetIncome) = "Net Income"

   
   .ColWidth(0) = 1500
   .ColWidth(1) = 1800
   .ColWidth(2) = 1800
   .ColWidth(3) = 2800
   .ColWidth(4) = 2800
   .ColWidth(5) = 1800
   .ColWidth(6) = 1000
   For j = 7 To 50
       .ColWidth(j) = 1000
   Next
   For j = 51 To 55
       .ColWidth(j) = 0
   Next
   .WordWrap = True
   .RowHeight(1) = 700
End With
End Sub

Private Sub Grid_Head13()
With MSHFlexGrid13
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 12
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "Opening TI Amount"
   .TextMatrix(1, 3) = "Opening Received Amount"
   .TextMatrix(1, 4) = "Opeing Advance Amount"
   .TextMatrix(1, 5) = "Opening Balance"
   .TextMatrix(1, 6) = "This Month TI Amount"
   .TextMatrix(1, 7) = "This Month Advance"
   .TextMatrix(1, 8) = "This Month Received Amount"
   .TextMatrix(1, 9) = "Closing Balance"
   .TextMatrix(1, 10) = "Interest Amount"
   .TextMatrix(1, 11) = "LocationID"
   
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 1900
   .ColWidth(3) = 1800
   .ColWidth(4) = 2500
   .ColWidth(5) = 2500
   .ColWidth(6) = 2500
   .ColWidth(7) = 1500
   .ColWidth(8) = 1500
   .ColWidth(9) = 1500
   .ColWidth(10) = 1500
   .ColWidth(11) = 0
      
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head14()
With MSHFlexGrid14
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 7
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "CMP Name"
   .TextMatrix(1, 3) = "Godown No"
   .TextMatrix(1, 4) = "Godown Address"
   .TextMatrix(1, 5) = "Employee Name"
   .TextMatrix(1, 6) = "Salary Amount"
   
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 1900
   .ColWidth(3) = 1800
   .ColWidth(4) = 2500
   .ColWidth(5) = 2500
   .ColWidth(6) = 2500
      
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head15()
With MSHFlexGrid15
   .Clear
   .Rows = 3
   .FixedRows = 2
   .Cols = 8
   .TextMatrix(1, 0) = "State Name"
   .TextMatrix(1, 1) = "Location Name"
   .TextMatrix(1, 2) = "Expense Type"
   .TextMatrix(1, 3) = "Client Name"
   .TextMatrix(1, 4) = "Depository Name"
   .TextMatrix(1, 5) = "TIID"
   .TextMatrix(1, 6) = "Rec TDS"
   .TextMatrix(1, 7) = "Rec Other Ded."
   
   .ColWidth(0) = 1900
   .ColWidth(1) = 1900
   .ColWidth(2) = 1900
   .ColWidth(3) = 1800
   .ColWidth(4) = 2500
   .ColWidth(5) = 2500
   .ColWidth(6) = 2500
   .ColWidth(7) = 2500
      
   .WordWrap = True
   .RowHeight(1) = 600
End With
End Sub

Private Sub Grid_Head16()
With MSHFlexGrid16
   .Clear
   .Rows = 3
   .Cols = 14
   .FixedCols = 1
   .FixedRows = 2
   .WordWrap = True
   .TextMatrix(1, 0) = "ClaimID"
   .TextMatrix(1, 1) = "Segment"
   .TextMatrix(1, 2) = "State"
   .TextMatrix(1, 3) = "Location"
   .TextMatrix(1, 4) = "Employee/Vendor Name"
   .TextMatrix(1, 5) = "Current Designation"
   .TextMatrix(1, 6) = "Expense for the Month"
   .TextMatrix(1, 7) = "Expense Type"
   .TextMatrix(1, 8) = "Claimed Amount"
   .TextMatrix(1, 9) = "Status"
   .TextMatrix(1, 10) = "Employee/ Vendor"
   .TextMatrix(1, 11) = "Approval Date"
   .TextMatrix(1, 12) = "Approval Remarks"
   .TextMatrix(1, 13) = "Distribution Flag"
   .RowHeight(1) = 800
End With
End Sub

Private Sub Grid_Head17()
With MSHFlexGrid17
   .Clear
   .Rows = 3
   .Cols = 6
   .FixedCols = 1
   .FixedRows = 2
   .WordWrap = True
     
   .TextMatrix(1, 0) = "EMPLOYEE NAME"
   .TextMatrix(1, 1) = "SEGMENT "
   .TextMatrix(1, 2) = "DEPARTMENT"
   .TextMatrix(1, 3) = "DESIGNATION"
   .TextMatrix(1, 4) = "CONTRACT TYPE"
   .TextMatrix(1, 5) = "COUNT"
      
   .RowHeight(1) = 800
End With
End Sub


Private Sub GetOutstandingInterest()
Call Grid_Head13
tmp = " Select distinct MS.StateName,ML.LocationName, ML.LocationID"
tmp = tmp & " From Mst_Godown MG"
tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
If Wc <> "" Then
   tmp = tmp & Wc & " And (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
Else
   tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
End If
tmp = tmp & " Order by MS.StateName,ML.LocationName, ML.LocationID"
Call TmpTable

tmp = "Select LocationName,Sum(TIAmount)TIAmount,Sum(RecAmount)+Sum(RecTDSAmount)+Sum(RecOtherDed) RecAmount"
tmp = tmp & " From (Select ML.LocationName,TI.TIID,TI.TIAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.RejectFlag <> 'Y'"
tmp = tmp & " And TCR.CRDate<'" & mStartDate & "'"
tmp = tmp & " )+ isnull(TI.AdjAmountFAG,0) as RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.RejectFlag <> 'Y'"
tmp = tmp & " And TCR.CRDate<'" & mStartDate & "'"
tmp = tmp & " ) +(Select isnull(Sum(TCRT.RecTDS),0) as RecTDS"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.Flag<>'R'And TIID=TI.TIID"
tmp = tmp & " And TCR.AdjDate<'" & mStartDate & "'"
tmp = tmp & "  )RecTDSAmount,(Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.RejectFlag <> 'Y'"
tmp = tmp & " And TCR.CRDate<'" & mStartDate & "'"
tmp = tmp & " )+"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.Flag<>'R'And TIID=TI.TIID And TCR.AdjDate<'" & mStartDate & "'"
tmp = tmp & " )RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Where  TINo<>'' And TI.TIDate<'" & mStartDate & "' )A"
tmp = tmp & " Group By A.LocationNAme"
tmp = tmp & " Order By A.LocationNAme"
Call Tmp1Table

tmp = "Select LocationName,Sum(TIAmount)TIAmount,Sum(RecAmount)+Sum(RecTDSAmount)+Sum(RecOtherDed) RecAmount"
tmp = tmp & " From ("
tmp = tmp & " Select ML.LocationName,TI.TIID,TI.TIAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.RejectFlag <> 'Y'"
tmp = tmp & " And TCR.CRDate>='" & mStartDate & "' And TCR.CRDate<='" & mEndDate & "'"
tmp = tmp & " )+ isnull(TI.AdjAmountFAG,0) as RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.RejectFlag <> 'Y'"
tmp = tmp & " And TCR.CRDate>='" & mStartDate & "' And TCR.CRDate<='" & mEndDate & "'"
tmp = tmp & " ) +"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.Flag<>'R'And TIID=TI.TIID"
tmp = tmp & " And TCR.AdjDate>='" & mStartDate & "' And TCR.AdjDate<='" & mEndDate & "'"
tmp = tmp & " )RecTDSAmount,(Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCR.RejectFlag <> 'Y'"
tmp = tmp & " And TCR.CRDate>='" & mStartDate & "' And TCR.CRDate<='" & mEndDate & "'"
tmp = tmp & " )+"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed"
tmp = tmp & " From Txn_CRTDSAdjustmentDetail TCRT"
tmp = tmp & " Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " Where TCR.Flag<>'R'And TIID=TI.TIID And TCR.AdjDate>='" & mStartDate & "' And TCR.AdjDate<='" & mEndDate & "'"
tmp = tmp & " )RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Where  TINo<>'' And TI.TIDate<='" & mEndDate & "' )A"
tmp = tmp & " Group By A.LocationNAme"
tmp = tmp & " Order By A.LocationNAme"

Call Tmp6Table

tmp = " Select LocationName,Sum(TIAmount)TIAmount,Sum(RecAmount)+Sum(RecTDS)+Sum(RecOtherDed) RecAmount from"
tmp = tmp & " (select ML.LocationName,TI.TIID,TI.TIAmount, isnull(TI.AdjAmountFAG,0)+"
tmp = tmp & " (select isnull(Sum(TCRT.RecAmount),0)+isnull(Sum(TCRT.RecTDS),0)+isnull(Sum(TCRT.RecOtherDed),0) from Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCRT.TIID = TI.TIID and TCR.ReturnFlag<>'Y'"
tmp = tmp & " And TCR.RejectFlag <> 'Y' And TCR.CRDate between '" & mStartDate & "' and '" & mEndDate & "') RecAmount,"
tmp = tmp & " (select isnull(Sum(TCRTA.RecTDS),0) from Txn_CRTDSAdjustmentDetail TCRTA"
tmp = tmp & " inner Join Txn_CRTDSAdjustment TCRA on TCRA.TDSAID= TCRTA.TDSAID and TCRA.SM_Prefix=TCRTA.SM_Prefix"
tmp = tmp & " Where TCRTA.TIID = TI.TIID and TCRA.Flag<>'R'And TCRTA.TIID=TI.TIID And TCRA.AdjDate between '" & mStartDate & "' And '" & mEndDate & "') as RecTDS,"
tmp = tmp & " (select isnull(Sum(TCRTA.RecOtherDed),0) from Txn_CRTDSAdjustmentDetail TCRTA"
tmp = tmp & " inner Join Txn_CRTDSAdjustment TCRA on TCRA.TDSAID= TCRTA.TDSAID and TCRA.SM_Prefix=TCRTA.SM_Prefix"
tmp = tmp & " Where TCRTA.TIID = TI.TIID and TCRA.Flag<>'R'And TCRTA.TIID=TI.TIID And TCRA.AdjDate between '" & mStartDate & "' And '" & mEndDate & "') as RecOtherDed"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Where  TINo<>'' And TI.TIDate between '" & mStartDate & "' And '" & mEndDate & "'"
tmp = tmp & " Group by ML.LocationName,TI.TIID,TI.TIAmount, TI.AdjAmountFAG"
tmp = tmp & " ) a"
tmp = tmp & " group by locationname"
Call Tmp3Table

tmp = "Select LocationName,Sum(AdvanceAmount) AdvanceAmount"
tmp = tmp & " From (Select ML.LocationName,TCRCD.Amount -"
tmp = tmp & " (select isnull(sum(isnull(TCTD.RecAmount,0))+sum(isnull(TCTD.RecTDS,0))+sum(isnull(TCTD.RecOtherDed,0)),0)"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCTD"
tmp = tmp & " Inner Join Txn_TIDetails TI on TI.TIID=TCTD.TIID"
tmp = tmp & " Where TCTD.CRID = CR.CRID And TCTD.SM_Prefix = CR.SM_Prefix And (TCTD.ClientIDRow = TCRCD.ClientIDRow Or TCTD.DPID = TCRCD.DPID) And TCTD.CRAAID > 0"
tmp = tmp & " And TI.TIDate<'" & mStartDate & "'"
tmp = tmp & " ) as AdvanceAmount"
tmp = tmp & " From Txn_NewCashReceipt CR"
tmp = tmp & " Inner Join Mst_Location ML on CR.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Txn_NewCashReceiptClientDetail TCRCD on CR.CRID = TCRCD.CRID And CR.SM_Prefix = TCRCD.SM_Prefix"
tmp = tmp & " where CR.AdvanceFlag = 1 And TCRCD.Amount > 0 And"
tmp = tmp & " (TCRCD.Amount  -"
tmp = tmp & " (select sum(isnull(TCTD.RecAmount,0))+sum(isnull(TCTD.RecTDS,0))+sum(isnull(TCTD.RecOtherDed,0))"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCTD"
tmp = tmp & " Inner Join Txn_TIDetails TI on TI.TIID=TCTD.TIID"
tmp = tmp & " Where TCTD.CRID= CR.CRID And TCTD.SM_Prefix = CR.SM_Prefix And (TCTD.ClientIDrow = TCRCD.ClientIDRow Or TCTD.DPID = TCRCD.DPID) And TCTD.CRAAID>0 )) >0"
tmp = tmp & " And (CR.RejectFlag is Null or CR.RejectFlag = 'N') And (isnull(CR.ReturnFlag,'') <> 'Y')"
tmp = tmp & " And CR.CRDate<'" & mStartDate & "'"
tmp = tmp & " )A"
tmp = tmp & " Group By LocationName"

Call Tmp4Table

tmp = "Select LocationName,Sum(AdvanceAmount) AdvanceAmount"
tmp = tmp & " From (Select ML.LocationName,TCRCD.Amount -"
tmp = tmp & " (select isnull(sum(isnull(TCTD.RecAmount,0))+sum(isnull(TCTD.RecTDS,0))+sum(isnull(TCTD.RecOtherDed,0)),0)"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCTD"
tmp = tmp & " Inner Join Txn_TIDetails TI on TI.TIID=TCTD.TIID"
tmp = tmp & " Where TCTD.CRID = CR.CRID And TCTD.SM_Prefix = CR.SM_Prefix And (TCTD.ClientIDRow = TCRCD.ClientIDRow Or TCTD.DPID = TCRCD.DPID) And TCTD.CRAAID > 0"
tmp = tmp & " And TI.TIDate<='" & mEndDate & "'"
tmp = tmp & " ) as AdvanceAmount"
tmp = tmp & " From Txn_NewCashReceipt CR"
tmp = tmp & " Inner Join Mst_Location ML on CR.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_State MS on ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Txn_NewCashReceiptClientDetail TCRCD on CR.CRID = TCRCD.CRID And CR.SM_Prefix = TCRCD.SM_Prefix"
tmp = tmp & " where CR.AdvanceFlag = 1 And TCRCD.Amount > 0 And"
tmp = tmp & " (TCRCD.Amount  -"
tmp = tmp & " (select isnull(sum(TCTD.RecAmount),0)+isnull(sum(TCTD.RecTDS),0)+isnull(sum(TCTD.RecOtherDed),0)"
tmp = tmp & " From Txn_NewCashReceiptTIDetail TCTD"
tmp = tmp & " Inner Join Txn_TIDetails TI on TI.TIID=TCTD.TIID"
tmp = tmp & " Where TCTD.CRID= CR.CRID And TCTD.SM_Prefix = CR.SM_Prefix And (TCTD.ClientIDrow = TCRCD.ClientIDRow Or TCTD.DPID = TCRCD.DPID) And CRAAID>0 And TI.TIDate<='" & mEndDate & "')) >0"
tmp = tmp & " And (CR.RejectFlag is Null or CR.RejectFlag = 'N') And (isnull(CR.ReturnFlag,'') <> 'Y')"
tmp = tmp & " And CR.CRDate>='" & mStartDate & "' And CR.CRDate<='" & mEndDate & "'"
tmp = tmp & " )A"
tmp = tmp & " Group By LocationName"

Call Tmp5Table

Call Grid_Head13

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid13.Rows = MSHFlexGrid13.Rows + TmpRs.RecordCount
   With MSHFlexGrid13
      For i = 2 To TmpRs.RecordCount + 1
         .TextMatrix(i, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
         .TextMatrix(i, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
         If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
         TmpRs1.Find "LocationName='" & TmpRs!LocationName & "'"
         If TmpRs1.EOF = False Then
            .TextMatrix(i, 2) = IIf(IsNull(TmpRs1!TIAmount), 0, TmpRs1!TIAmount)
            .TextMatrix(i, 3) = IIf(IsNull(TmpRs1!RecAmount), 0, TmpRs1!RecAmount)
         End If
         If TmpRs4.RecordCount > 0 Then TmpRs4.MoveFirst
         TmpRs4.Find "LocationName='" & TmpRs!LocationName & "'"
         If TmpRs4.EOF = False Then
            .TextMatrix(i, 4) = IIf(IsNull(TmpRs4!AdvanceAmount), 0, TmpRs4!AdvanceAmount)
         End If
         .TextMatrix(i, 5) = Format(Val(.TextMatrix(i, 2)) - (Val(.TextMatrix(i, 3)) + Val(.TextMatrix(i, 4))), "######.00")
         If TmpRs6.RecordCount > 0 Then TmpRs6.MoveFirst
         TmpRs6.Find "LocationName='" & TmpRs!LocationName & "'"
         If TmpRs6.EOF = False Then
            .TextMatrix(i, 8) = IIf(IsNull(TmpRs6!RecAmount), 0, TmpRs6!RecAmount)
         End If
         If TmpRs3.RecordCount > 0 Then TmpRs3.MoveFirst
         TmpRs3.Find "LocationName='" & TmpRs!LocationName & "'"
         If TmpRs3.EOF = False Then
            .TextMatrix(i, 6) = IIf(IsNull(TmpRs3!TIAmount), 0, TmpRs3!TIAmount)
         End If
         If TmpRs5.RecordCount > 0 Then TmpRs5.MoveFirst
         TmpRs5.Find "LocationName='" & TmpRs!LocationName & "'"
         If TmpRs5.EOF = False Then
            .TextMatrix(i, 7) = IIf(IsNull(TmpRs5!AdvanceAmount), 0, TmpRs5!AdvanceAmount)
         End If
         
         .TextMatrix(i, 9) = Format(Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6)) - (Val(.TextMatrix(i, 7)) + Val(.TextMatrix(i, 8))), "######.00")
         .TextMatrix(i, 10) = Format(((((Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 9))) / 2) * 45 / 100) * 11 / 100) / 12, "###########.00")
         .TextMatrix(i, 11) = IIf(IsNull(TmpRs!LocationID), "", TmpRs!LocationID)
         .TextMatrix(0, 2) = Format(Val(.TextMatrix(i, 2)) + Val(.TextMatrix(0, 2)), "######.00")
         .TextMatrix(0, 3) = Format(Val(.TextMatrix(i, 3)) + Val(.TextMatrix(0, 3)), "######.00")
         .TextMatrix(0, 4) = Format(Val(.TextMatrix(i, 4)) + Val(.TextMatrix(0, 4)), "######.00")
         .TextMatrix(0, 5) = Format(Val(.TextMatrix(i, 5)) + Val(.TextMatrix(0, 5)), "######.00")
         .TextMatrix(0, 6) = Format(Val(.TextMatrix(i, 6)) + Val(.TextMatrix(0, 6)), "######.00")
         .TextMatrix(0, 7) = Format(Val(.TextMatrix(i, 7)) + Val(.TextMatrix(0, 7)), "######.00")
         .TextMatrix(0, 8) = Format(Val(.TextMatrix(i, 8)) + Val(.TextMatrix(0, 8)), "######.00")
         .TextMatrix(0, 9) = Format(Val(.TextMatrix(i, 9)) + Val(.TextMatrix(0, 9)), "######.00")
         .TextMatrix(0, 10) = Format(Val(.TextMatrix(i, 10)) + Val(.TextMatrix(0, 10)), "######.00")
         TmpRs.MoveNext
      Next
   End With
End If

For i = 2 To MSHFlexGrid1.Rows - 2
   TmpRs2.MoveFirst
   TmpRs2.Find "LocationID = " & Val(MSHFlexGrid1.TextMatrix(i, E.LocationID))
   'If Val(MSHFlexGrid1.TextMatrix(I, E.LocationID)) = 111 Then MsgBox "stop"
   If Not TmpRs2.EOF Then
      For j = 2 To MSHFlexGrid13.Rows - 2
         If MSHFlexGrid1.TextMatrix(i, E.LocationID) = MSHFlexGrid13.TextMatrix(j, 11) Then
            If Opt_Stock.Value = True Then
               mStock = Val(MSHFlexGrid1.TextMatrix(i, E.NcdexStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.NonNCDEXStock))
               MSHFlexGrid1.TextMatrix(i, E.InterestCharges) = Format(Val(MSHFlexGrid13.TextMatrix(j, 10)) * Val(mStock) / TmpRs2!BalanceWeight, "########.00")
               MSHFlexGrid1.TextMatrix(0, E.InterestCharges) = Format(Val(MSHFlexGrid13.TextMatrix(j, 10)) * Val(mStock) / TmpRs2!BalanceWeight, "########.00") + Val(MSHFlexGrid1.TextMatrix(0, E.InterestCharges))
            ElseIf Opt_Capacity.Value = True Then
               MSHFlexGrid1.TextMatrix(i, E.InterestCharges) = Format(Val(MSHFlexGrid13.TextMatrix(j, 10)) * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "########.00")
               MSHFlexGrid1.TextMatrix(0, E.InterestCharges) = Format(Val(MSHFlexGrid13.TextMatrix(j, 10)) * Val(MSHFlexGrid1.TextMatrix(i, E.GodownCapacity)) / TmpRs2!Capacity, "########.00") + Val(MSHFlexGrid1.TextMatrix(0, E.InterestCharges))
            End If
            'MSHFlexGrid1.TextMatrix(0, E.InterestCharges) = Val(MSHFlexGrid1.TextMatrix(I, E.InterestCharges)) + Val(MSHFlexGrid1.TextMatrix(0, E.InterestCharges))
            Exit For
         End If
      Next j
   End If
Next i
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame11.Enabled = False
   SSTab1.Enabled = False
   Exit Sub
End If
If LCase(Gen_UserRole) <> "power" And LCase(Gen_UserRole) <> "head cmg" Then
   ChkFromSavedData.Value = vbChecked
   ChkFromSavedData.Enabled = False
End If
Frame11.Enabled = True
SSTab1.Enabled = True
Call CmdClear_Click
Call ChkFromSavedData_Click
End Sub

Private Sub TxtFromDate_LostFocus()
TxtFromDate.Day = 1
End Sub

Private Sub TxtToDate_LostFocus()
mTotalDays = Day(DateSerial(TxtToDate.Year, TxtToDate.Month + 1, 0))
TxtToDate.Day = Val(mTotalDays)
End Sub


Private Sub GetSalaryExpenseFromIncomeforFieldStaff()

Dim mTotalLocIncomeforSal, mSalLocationID, mSalStateID As Variant
Dim mChkRow As Boolean

mChkRow = False

tmp = " Select ME.EmployeeNo,ME.fullname,MD.Designation,MDep.Department ,sum(isnull(BasicMonthly,0)) + sum(isnull(HRAMonthly,0)) + sum(isnull(ConveyanceMonthly,0)) +"
tmp = tmp & " sum(isnull(OtherAllMonthly,0)) + sum(isnull(MEDICAL,0)) + sum(isnull(SplAllMonthly,0)) + sum(isnull(LTA,0)) +"
tmp = tmp & " Sum (IsNull(LaptopAllowMonthly, 0)) + Sum(IsNull(LeaveEncashment, 0)) + Sum(IsNull(GainShare, 0)) + Sum(IsNull(Incentive, 0))"
tmp = tmp & " + sum(isnull(Perquisites,0)) + sum(isnull(ARREARSBASIC,0)) + sum(isnull(ARREARSHRA,0)) +"
tmp = tmp & " sum(isnull(ARREARSCONVEYANCE,0)) + sum(isnull(ARREARSMEDICAL,0)) + sum(isnull(PERFORMANCEBONUS,0)) +"
tmp = tmp & " Sum (IsNull(PROFESSIONALFEES, 0)) + Sum(IsNull(STIPEND, 0)) + Sum(IsNull(SERVICECHARGES, 0))"
tmp = tmp & " + sum(isnull(ACCRUEDVARIABLEPAY,0)) + sum(isnull(ACCRUEDRETIRALS,0))+sum(isnull(employersesic,0))"
tmp = tmp & " + sum(isnull(carmaintandpetrol,0))+ sum(isnull(booksandperiodical,0))+ sum(isnull(mealallowance,0))"
tmp = tmp & " + sum(isnull(LWFEMPLOYER,0)) + sum(isnull(EMPLOYERSPF,0)) + sum(isnull(STaxandEducationcess,0)) Amount"
tmp = tmp & " From Txn_SalaryDistribution  MSD"
tmp = tmp & " Inner Join Mst_Employee ME on MSD.EmployeeID = ME.EmployeeID"
tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID"
tmp = tmp & " Inner Join Mst_Department MDep On ME.DepartmentID = MDep.DepartmentID"
tmp = tmp & " Where MSD.MonthYear Between  '" & mStartDate & "' And '" & mEndDate & "' And SegmentID = 1"
tmp = tmp & " And MD.Designation + MDep.Department   in ( 'Area ExecutiveBusiness & Operations'   ,"
tmp = tmp & " 'Area ManagerBusiness & Operations'     ,"
tmp = tmp & " 'Business Dev ExecutiveBusiness & Operations'   ,"
tmp = tmp & " 'Business Development ManagerBusiness & Operations'     ,"
tmp = tmp & " 'Regional Manager - NorthBusiness & Operations'     ,"
tmp = tmp & " 'Regional Manager - RajasthanBusiness & Operations'     ,"
tmp = tmp & " 'Regional Manager - SouthBusiness & Operations'     ,"
tmp = tmp & " 'Regional Manager - WestBusiness & Operations'  ,"
tmp = tmp & " 'Regional Manager -EastBusiness & Operations'   ,"
tmp = tmp & " 'State HeadBusiness & Operations'   ,"
tmp = tmp & "     'CMP ManagerStorage And Preservation'   ,"
tmp = tmp & "     'CMP Manager cum DRCStorage And Preservation'   ,"
tmp = tmp & "     'ConsultantStorage And Preservation'    ,"
tmp = tmp & "     'DRCStorage And Preservation'   ,"
tmp = tmp & "     'ExecutiveStorage And Preservation' ,"
tmp = tmp & "     'OfficerStorage And Preservation'   ,"
tmp = tmp & "     'Operations ControllerStorage And Preservation' ,"
tmp = tmp & "     'QQSStorage And Preservation'   ,"
tmp = tmp & "     'Senior Executive Storage And Preservation')"
tmp = tmp & " Group by ME.EmployeeNo,ME.fullname,MD.Designation,MDep.Department"
tmp = tmp & " Order by ME.fullname,MD.Designation,MDep.Department"

Call TmpTableHRPro

If TmpRsHRPro.RecordCount > 0 Then
   For j = 1 To TmpRsHRPro.RecordCount
       mStatus = 0
       mSalEmployee = 0
       mSalEmployee = TmpRsHRPro!Amount
  
       tmp = " Select MG.GodownID, sum(MG.GodownCapacity) Capacity ,"
       tmp = tmp & " sum(MGPL.NNNRIncomeAdhoc + MGPL.NNNRIncome + MGPL.ReservationIncomeAdhoc +"
       tmp = tmp & " MGPL.ReservationIncome + MGPL.SevenDaysIncome + MGPL.DPIncomeAdHoc + MGPL.FinalDPIncome  +"
       tmp = tmp & " MGPL.RejectedIncome + MGPL.SpotIncomeAdHoc + MGPL.OtherIncome + MGPL.SCMIncome +"
       tmp = tmp & " MGPL.CMUnbilledIncome + MGPL.CMPendingIncome + MGPL.CMApprovedIncome + MGPL.LeftOverStock)"
       tmp = tmp & " From Mst_Godown MG"
       tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
       tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
       tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
       tmp = tmp & " Inner Join Txn_MonGodownWiseProfitability MGPL on MGPL.GodownID = MG.GodownID and MGPL.MonthYear = '" & mStartDate & "'"
       tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') "
       tmp = tmp & " And MG.StartDate <= '" & mEndDate & "' and MCMP.LocationID in ("
       tmp = tmp & " Select distinct MEAD.AllocationID LocationID From Mst_EmployeeAllocationDetails MEAD"
       tmp = tmp & " Inner Join Mst_Employee ME on ME.EmployeeID = MEAD.EmployeeID"
       tmp = tmp & " Where ME.EmployeeNo = '" & TmpRsHRPro!EmployeeNo & "' And AllocationType = 'L'"
       tmp = tmp & " Union"
       tmp = tmp & " Select distinct ML.LocationID From Mst_EmployeeAllocationDetails MEAD"
       tmp = tmp & " Inner Join Mst_Employee ME on ME.EmployeeID = MEAD.EmployeeID"
       tmp = tmp & " Inner Join Mst_State MS on MS.StateID = MEAD.AllocationID"
       tmp = tmp & " Inner Join Mst_Location ML on MS.StateID = ML.StateID"
       tmp = tmp & " Where ME.EmployeeNo = '" & TmpRsHRPro!EmployeeNo & "' And AllocationType = 'S'"
       tmp = tmp & " Union"
       tmp = tmp & " Select distinct ML.LocationID From Mst_EmployeeAllocationDetails MEAD"
       tmp = tmp & " Inner Join Mst_Employee ME on ME.EmployeeID = MEAD.EmployeeID"
       tmp = tmp & " Inner Join Mst_State MS on MS.RegionID = MEAD.AllocationID"
       tmp = tmp & " Inner Join Mst_Location ML on MS.StateID = ML.StateID"
       tmp = tmp & " Where ME.EmployeeNo = '" & TmpRsHRPro!EmployeeNo & "' And AllocationType = 'R'"
       tmp = tmp & " )"
       tmp = tmp & " group by MG.GodownID"
'--and mg.godownid = 2324
       Call Tmp1Table
'---
       If TmpRs1.RecordCount > 0 Then
          mTotalGodIncomeforSal = 0
          For X = 1 To TmpRs1.RecordCount
              For i = 2 To MSHFlexGrid1.Rows - 2
                  If MSHFlexGrid1.TextMatrix(i, E.GodownID) = TmpRs1!GodownID Then
                     mTotalGodIncomeforSal = mTotalGodIncomeforSal + Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.PledgeIncome))
                  End If
              Next
              TmpRs1.MoveNext
          Next
          TmpRs1.MoveFirst
          For X = 1 To TmpRs1.RecordCount
              For i = 2 To MSHFlexGrid1.Rows - 2
                  If MSHFlexGrid1.TextMatrix(i, E.GodownID) = TmpRs1!GodownID Then
'                     mTotalGodIncomeforSal = mTotalGodIncomeforSal + Val(MSHFlexGrid1.TextMatrix(x, E.NNNRIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(x, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(x, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(x, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(x, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(x, E.PledgeIncome))
                      mGodownIncome = Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.PledgeIncome)) '+ Val(MSHFlexGrid1.TextMatrix(i, E.Salary))
                      If mTotalGodIncomeforSal = 0 Then
                         MSHFlexGrid1.TextMatrix(i, E.SalaryFieldStaff) = Val(MSHFlexGrid1.TextMatrix(i, E.SalaryFieldStaff)) + IIf(IsNull(mSalEmployee), 0, Format(mSalEmployee / TmpRs1.RecordCount, "#############.00"))
                         MSHFlexGrid1.TextMatrix(0, E.SalaryFieldStaff) = Val(MSHFlexGrid1.TextMatrix(0, E.SalaryFieldStaff)) + IIf(IsNull(mSalEmployee), 0, Format(mSalEmployee / TmpRs1.RecordCount, "#############.00"))
                      Else
                         MSHFlexGrid1.TextMatrix(i, E.SalaryFieldStaff) = Val(MSHFlexGrid1.TextMatrix(i, E.SalaryFieldStaff)) + IIf(IsNull(mSalEmployee), 0, Format(mSalEmployee * Val(mGodownIncome) / Val(mTotalGodIncomeforSal), "#############.00"))
                         MSHFlexGrid1.TextMatrix(0, E.SalaryFieldStaff) = Val(MSHFlexGrid1.TextMatrix(0, E.SalaryFieldStaff)) + IIf(IsNull(mSalEmployee), 0, Format(mSalEmployee * Val(mGodownIncome) / Val(mTotalGodIncomeforSal), "#############.00"))
                      End If
                  End If
              Next
              TmpRs1.MoveNext
          Next
       Else
          If TmpRsHRPro!Amount <> 0 Then
             TxtUnallocatedSal = Trim(TxtUnallocatedSal) & TmpRsHRPro!EmployeeNo & "-" & TmpRsHRPro!Amount & " " & vbCrLf
             mUnAllocatedSalary = mUnAllocatedSalary + TmpRsHRPro!Amount
          
             If ChkFromSavedData.Value = vbUnchecked Or ChkSalary.Value = vbChecked Then
                With MSHFlexGrid1
                     If ChkFromSavedData.Value = vbChecked And .TextMatrix(.Rows - 2, E.AMname) = "CMG-SalaryExpense of Closed Locations" Then
                        k = .Rows - 2
                        .TextMatrix(0, E.SalaryFieldStaff) = Val(.TextMatrix(0, E.SalaryFieldStaff)) - Val(.TextMatrix(k, E.SalaryFieldStaff))
                     ElseIf ChkFromSavedData.Value = vbChecked And .TextMatrix(.Rows - 3, E.AMname) = "CMG-SalaryExpense of Closed Locations" Then
                        k = .Rows - 3
                        .TextMatrix(0, E.SalaryFieldStaff) = Val(.TextMatrix(0, E.SalaryFieldStaff)) - Val(.TextMatrix(k, E.SalaryFieldStaff))
                     Else
                        If mChkRow = False Then
                           k = .Rows - 1
                           .Rows = .Rows + 1
                           mChkRow = True
                        End If
                     End If
                     .TextMatrix(k, E.MonthYear) = CmbMonth.Text & "-" & CmbYear.Text
                     .TextMatrix(k, E.AMname) = "CMG-SalaryExpense of Closed Locations"
                     .TextMatrix(k, E.RMName) = "CMG-SalaryExpense of Closed Locations"
                     .TextMatrix(k, E.SalaryFieldStaff) = Val(.TextMatrix(k, E.SalaryFieldStaff)) + IIf(IsNull(TmpRsHRPro!Amount), "", TmpRsHRPro!Amount)
                     .TextMatrix(0, E.SalaryFieldStaff) = Val(.TextMatrix(0, E.SalaryFieldStaff)) + Val(.TextMatrix(k, E.SalaryFieldStaff))
                  End With
            End If
          End If
       End If
       TmpRsHRPro.MoveNext
   Next
End If

'      If mStatus = 0 Then
'         'tmp = " Select MS.StateID, sum(MG.GodownCapacity) Capacity"
'         tmp = " Select Distinct MS.StateID"
'         tmp = tmp & " From Mst_Godown MG"
'         tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
'         tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
'         tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
'         tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
'         tmp = tmp & " And MS.StateID in (Select StateID from Mst_Location Where LocationID = " & TmpRsHRPro!LocationID & " ) group by MS.StateID"
'         Call Tmp3Table
'
'         If TmpRs3.RecordCount = 0 Then
'            tmp = " Select Distinct MS.StateID "
'            tmp = tmp & " From Mst_Godown MG"
'            tmp = tmp & " Inner Join Mst_CMP MCMP on MCMP.CMPID = MG.CMPID"
'            tmp = tmp & " Inner Join Mst_Location ML on ML.LocationID = MCMP.LocationID"
'            tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
'            tmp = tmp & " Where (MG.CloseDate is null Or MG.CloseDate >= '" & mStartDate & "') And MG.StartDate <= '" & mEndDate & "'"
'            tmp = tmp & " And MS.RegionID in (Select MS.RegionID from Mst_State MS Inner Join Mst_Location ML on ML.StateID = MS.StateID Where ML.LocationID = " & TmpRsHRPro!LocationID & " ) group by MS.StateID"
'            Call Tmp3Table
'
'            mSalStateID = ""
'            For i = 2 To MSHFlexGrid1.Rows - 2
'                TmpRs3.MoveFirst
'                TmpRs3.Find "StateID = " & Val(MSHFlexGrid1.TextMatrix(i, E.StateID))
'                If Not TmpRs3.EOF Then
'                   If MSHFlexGrid1.TextMatrix(i, E.StateID) = TmpRs3!StateID Then
'                      If mSalStateID <> TmpRs3!StateID Then
'                         mSalStateID = TmpRs3!StateID
'                         mTotalLocIncomeforSal = 0
'                         For x = 2 To MSHFlexGrid1.Rows - 2
'                             If mSalStateID = Val(MSHFlexGrid1.TextMatrix(x, E.StateID)) Then
'                                mTotalLocIncomeforSal = Val(mTotalLocIncomeforSal) + Val(MSHFlexGrid1.TextMatrix(x, E.NNNRIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(x, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(x, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(x, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(x, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(x, E.PledgeIncome)) '+ Val(MSHFlexGrid1.TextMatrix(0, E.Salary))
'                             End If
'                         Next
'                      End If
'                      mGodownIncome = Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.PledgeIncome)) '+ Val(MSHFlexGrid1.TextMatrix(i, E.Salary))
'                      MSHFlexGrid1.TextMatrix(i, E.Salary) = Val(MSHFlexGrid1.TextMatrix(i, E.Salary)) + IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(mGodownIncome) / Val(mTotalLocIncomeforSal), "#############.00"))
'                      MSHFlexGrid1.TextMatrix(0, E.Salary) = Val(MSHFlexGrid1.TextMatrix(0, E.Salary)) + IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(mGodownIncome) / Val(mTotalLocIncomeforSal), "#############.00"))
'                   End If
'                End If
'            Next
'         Else
'            mSalStateID = ""
'            For i = 2 To MSHFlexGrid1.Rows - 2
'                TmpRs3.MoveFirst
'                TmpRs3.Find "StateID = " & Val(MSHFlexGrid1.TextMatrix(i, E.StateID))
'                If Not TmpRs3.EOF Then
'                   If MSHFlexGrid1.TextMatrix(i, E.StateID) = TmpRs3!StateID Then
'                      If mSalStateID <> TmpRs3!StateID Then
'                         mSalStateID = TmpRs3!StateID
'                         mTotalLocIncomeforSal = 0
'                         For x = 2 To MSHFlexGrid1.Rows - 2
'                             If mSalStateID = Val(MSHFlexGrid1.TextMatrix(x, E.StateID)) Then
'                                mTotalLocIncomeforSal = Val(mTotalLocIncomeforSal) + Val(MSHFlexGrid1.TextMatrix(x, E.NNNRIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(x, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(x, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(x, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(x, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(x, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(x, E.PledgeIncome)) '+ Val(MSHFlexGrid1.TextMatrix(0, E.Salary))
'                             End If
'                         Next
'                      End If
'                      mGodownIncome = Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.PledgeIncome)) '+ Val(MSHFlexGrid1.TextMatrix(i, E.Salary))
'                      MSHFlexGrid1.TextMatrix(i, E.Salary) = Val(MSHFlexGrid1.TextMatrix(i, E.Salary)) + IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(mGodownIncome) / Val(mTotalLocIncomeforSal), "#############.00"))
'                      MSHFlexGrid1.TextMatrix(0, E.Salary) = Val(MSHFlexGrid1.TextMatrix(0, E.Salary)) + IIf(IsNull(TmpRsHRPro!Amount), 0, Format(TmpRsHRPro!Amount * Val(mGodownIncome) / Val(mTotalLocIncomeforSal), "#############.00"))
'                   End If
'                End If
'            Next
'         End If
'      End If
'   Next
'End If

End Sub

Private Sub GetSalaryExpenseFromIncomeforCOStaff()

'tmp = "Select sum(isnull(BasicMonthly,0)) + sum(isnull(HRAMonthly,0)) + sum(isnull(ConveyanceMonthly,0)) +"
'tmp = tmp & " sum(isnull(OtherAllMonthly,0)) + sum(isnull(MEDICAL,0)) + sum(isnull(SplAllMonthly,0)) + sum(isnull(LTA,0)) +"
'tmp = tmp & " Sum (IsNull(LaptopAllowMonthly, 0)) + Sum(IsNull(LeaveEncashment, 0)) + Sum(IsNull(GainShare, 0)) + Sum(IsNull(Incentive, 0))"
'tmp = tmp & " + sum(isnull(Perquisites,0)) + sum(isnull(ARREARSBASIC,0)) + sum(isnull(ARREARSHRA,0)) +"
'tmp = tmp & " sum(isnull(ARREARSCONVEYANCE,0)) + sum(isnull(ARREARSMEDICAL,0)) + sum(isnull(PERFORMANCEBONUS,0)) +"
'tmp = tmp & " Sum (IsNull(PROFESSIONALFEES, 0)) + Sum(IsNull(STIPEND, 0)) + Sum(IsNull(SERVICECHARGES, 0))"
'tmp = tmp & " + sum(isnull(ACCRUEDVARIABLEPAY,0)) + sum(isnull(ACCRUEDRETIRALS,0))+sum(isnull(employersesic,0))"
'tmp = tmp & " + sum(isnull(carmaintandpetrol,0))+ sum(isnull(booksandperiodical,0))+ sum(isnull(mealallowance,0))"
'tmp = tmp & " + sum(isnull(LWFEMPLOYER,0)) + sum(isnull(EMPLOYERSPF,0)) + sum(isnull(STaxandEducationcess,0)) Amount"
'tmp = tmp & " From Txn_SalaryDistribution  MSD"
'tmp = tmp & " Inner Join Mst_Employee ME on MSD.EmployeeID = ME.EmployeeID"
'tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID"
'tmp = tmp & " Inner Join Mst_Department MDep On ME.DepartmentID = MDep.DepartmentID"
'tmp = tmp & " Where MSD.MonthYear Between  '" & mStartDate & "' And '" & mEndDate & "' And SegmentID = 1"
''tmp = tmp & " And MDep.Department not in ('Business & Operations ','Storage And Preservation')"
'tmp = tmp & " And MD.Designation not in ('Area Executive','Area Manager','Business Dev Executive','Business Development Manager','Regional Manager - North',"
'tmp = tmp & " 'Regional Manager - Rajasthan','Regional Manager - South','Regional Manager - West','Regional Manager -East',"
'tmp = tmp & " 'State Head','CMP Manager','CMP Manager cum DRC','Consultant','DRC','Executive','Officer','Operations Controller','QQS',"
'tmp = tmp & " 'Senior Executive')"


tmp = "Select sum(isnull(BasicMonthly,0)) + sum(isnull(HRAMonthly,0)) + sum(isnull(ConveyanceMonthly,0)) + "
tmp = tmp & " sum(isnull(OtherAllMonthly,0)) + sum(isnull(MEDICAL,0)) + sum(isnull(SplAllMonthly,0)) + sum(isnull(LTA,0)) +"
tmp = tmp & " Sum (IsNull(LaptopAllowMonthly, 0)) + Sum(IsNull(LeaveEncashment, 0)) + Sum(IsNull(GainShare, 0)) + Sum(IsNull(Incentive, 0))"
tmp = tmp & " + sum(isnull(Perquisites,0)) + sum(isnull(ARREARSBASIC,0)) + sum(isnull(ARREARSHRA,0)) +"
tmp = tmp & " sum(isnull(ARREARSCONVEYANCE,0)) + sum(isnull(ARREARSMEDICAL,0)) + sum(isnull(PERFORMANCEBONUS,0)) +"
tmp = tmp & " Sum (IsNull(PROFESSIONALFEES, 0)) + Sum(IsNull(STIPEND, 0)) + Sum(IsNull(SERVICECHARGES, 0))"
tmp = tmp & " + sum(isnull(ACCRUEDVARIABLEPAY,0)) + sum(isnull(ACCRUEDRETIRALS,0))+sum(isnull(employersesic,0))"
tmp = tmp & " + sum(isnull(carmaintandpetrol,0))+ sum(isnull(booksandperiodical,0))+ sum(isnull(mealallowance,0))"
tmp = tmp & " + sum(isnull(LWFEMPLOYER,0)) + sum(isnull(EMPLOYERSPF,0)) + sum(isnull(STaxandEducationcess,0)) Amount "
tmp = tmp & " From Txn_SalaryDistribution  MSD"
tmp = tmp & " Inner Join Mst_Employee ME on MSD.EmployeeID = ME.EmployeeID"
tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID"
tmp = tmp & " Inner Join Mst_Department MDep On ME.DepartmentID = MDep.DepartmentID  "
tmp = tmp & " Where MSD.MonthYear Between  '" & mStartDate & "' And '" & mEndDate & "' And SegmentID = 1  "
tmp = tmp & " And MD.Designation + MDep.Department  not in (    'Area ExecutiveBusiness & Operations'   ,"
tmp = tmp & "   'Area ManagerBusiness & Operations'     ,"
tmp = tmp & "   'Business Dev ExecutiveBusiness & Operations'   ,"
tmp = tmp & "   'Business Development ManagerBusiness & Operations'     ,"
tmp = tmp & "   'Regional Manager - NorthBusiness & Operations'     ,"
tmp = tmp & "   'Regional Manager - RajasthanBusiness & Operations'     ,"
tmp = tmp & "   'Regional Manager - SouthBusiness & Operations'     ,"
tmp = tmp & "   'Regional Manager - WestBusiness & Operations'  ,"
tmp = tmp & "   'Regional Manager -EastBusiness & Operations'   ,"
tmp = tmp & "   'State HeadBusiness & Operations'   ,"
tmp = tmp & "   'CMP ManagerStorage And Preservation'   ,"
tmp = tmp & "   'CMP Manager cum DRCStorage And Preservation'   ,"
tmp = tmp & "   'ConsultantStorage And Preservation'    ,"
tmp = tmp & "   'DRCStorage And Preservation'   ,"
tmp = tmp & "   'ExecutiveStorage And Preservation' ,"
tmp = tmp & "   'OfficerStorage And Preservation'   ,"
tmp = tmp & "   'Operations ControllerStorage And Preservation' ,"
tmp = tmp & "   'QQSStorage And Preservation'   ,"
tmp = tmp & "   'Senior Executive Storage And Preservation')    "

Call TmpTableHRPro

mTotalIncome = Val(MSHFlexGrid1.TextMatrix(0, E.NNNRIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(0, E.PledgeIncome)) '+ Val(MSHFlexGrid1.TextMatrix(0, E.Salary))
mSalaryCo = TmpRsHRPro!Amount + mUnAllocatedSalary '476065.38
MSHFlexGrid1.TextMatrix(0, E.Salary) = 0
'mTotalIncome = Val(MSHFlexGrid1.TextMatrix(0, E.TotalIncome)) '+ Val(MSHFlexGrid1.TextMatrix(0, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(0, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(0, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(0, E.PledgeIncome)) '+ Val(MSHFlexGrid1.TextMatrix(0, E.Salary))
If TmpRsHRPro.RecordCount > 0 Then
   For i = 2 To MSHFlexGrid1.Rows - 2
      mGodownIncome = Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.NNNRIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncomeAdhoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.ReservationIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SevenDaysIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.DPIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.FinalDPIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.RejectedIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SpotIncomeAdHoc)) + Val(MSHFlexGrid1.TextMatrix(i, E.OtherIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.SCMIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMUnbilledIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMPendingIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.CMApprovedIncome)) + Val(MSHFlexGrid1.TextMatrix(i, E.LeftOverStock)) + Val(MSHFlexGrid1.TextMatrix(i, E.PledgeIncome)) '+ Val(MSHFlexGrid1.TextMatrix(i, E.Salary))
      'mGodownIncome = Val(MSHFlexGrid1.TextMatrix(i, E.TotalIncome))
      MSHFlexGrid1.TextMatrix(i, E.Salary) = IIf(IsNull(mSalaryCo), 0, Format((mSalaryCo * mGodownIncome) / mTotalIncome, "#############.00"))
      MSHFlexGrid1.TextMatrix(0, E.Salary) = Val(MSHFlexGrid1.TextMatrix(i, E.Salary)) + Val(MSHFlexGrid1.TextMatrix(0, E.Salary))
   Next
End If

End Sub
