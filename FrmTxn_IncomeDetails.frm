VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_IncomeDetails 
   Caption         =   "Income Details"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   6255
      Left            =   120
      TabIndex        =   53
      Top             =   7080
      Visible         =   0   'False
      Width           =   15135
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   10005
         TabIndex        =   63
         Top             =   360
         Width           =   1000
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   9000
         TabIndex        =   62
         Top             =   360
         Width           =   1005
      End
      Begin VB.ComboBox CmbSearchTransactionType 
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
         ItemData        =   "FrmTxn_IncomeDetails.frx":0000
         Left            =   6000
         List            =   "FrmTxn_IncomeDetails.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   59
         Top             =   360
         Width           =   2880
      End
      Begin VB.ComboBox CmbSearchMode 
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
         ItemData        =   "FrmTxn_IncomeDetails.frx":0023
         Left            =   960
         List            =   "FrmTxn_IncomeDetails.frx":0039
         Sorted          =   -1  'True
         TabIndex        =   58
         Top             =   360
         Width           =   2880
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   120
         TabIndex        =   54
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3916
         _Version        =   393216
         Rows            =   3
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label22 
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
         TabIndex        =   65
         Top             =   360
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
         Left            =   11715
         TabIndex        =   64
         Top             =   375
         Width           =   2355
      End
      Begin VB.Label Label16 
         Caption         =   "Mode"
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
         Left            =   360
         TabIndex        =   61
         Top             =   405
         Width           =   600
      End
      Begin VB.Label Label12 
         Caption         =   "Transaction Type"
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
         Left            =   4320
         TabIndex        =   60
         Top             =   360
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   15015
      Begin VB.CheckBox ChkCxTFDate 
         Caption         =   "CxTF Date"
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
         Left            =   10710
         TabIndex        =   57
         Top             =   285
         Width           =   1335
      End
      Begin VB.TextBox TxtG_UID 
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
         Left            =   13680
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   55
         Top             =   5280
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox TxtNoOfBags 
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
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   12
         Top             =   3960
         Width           =   2640
      End
      Begin VB.TextBox TxtStorageCharges 
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
         Height          =   360
         Left            =   8520
         MaxLength       =   100
         TabIndex        =   15
         Top             =   3360
         Width           =   3120
      End
      Begin VB.TextBox TxtRate 
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
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   13
         Top             =   4560
         Width           =   2640
      End
      Begin VB.TextBox TxtServiceTax 
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
         Height          =   360
         Left            =   8520
         MaxLength       =   100
         TabIndex        =   16
         Top             =   3960
         Width           =   3120
      End
      Begin VB.TextBox TxtPINO 
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
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   14
         Top             =   5160
         Width           =   2640
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
         TabIndex        =   6
         Top             =   1440
         Width           =   4440
      End
      Begin VB.TextBox TxtNoOfdays 
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
         Height          =   360
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   11
         Top             =   3360
         Width           =   2640
      End
      Begin VB.TextBox TxtAmtInclOfServiceTax 
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
         Height          =   360
         Left            =   8520
         MaxLength       =   100
         TabIndex        =   17
         Top             =   4560
         Width           =   3120
      End
      Begin VB.CommandButton CmdClientIDRow 
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
         Height          =   300
         Left            =   13560
         TabIndex        =   9
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox TxtCXTFNO 
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
         Left            =   8520
         MaxLength       =   100
         TabIndex        =   2
         Top             =   240
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
         Height          =   480
         Left            =   1800
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2640
         Width           =   11640
      End
      Begin VB.TextBox TxtAmtInclOfSerTaxInsurance 
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
         Height          =   360
         Left            =   8520
         MaxLength       =   100
         TabIndex        =   18
         Top             =   5160
         Width           =   3120
      End
      Begin MSComCtl2.DTPicker TxtBillingMonthStartDate 
         Height          =   360
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   55443457
         CurrentDate     =   39864
      End
      Begin MSComCtl2.DTPicker TxtBillingMonthEndDate 
         Height          =   360
         Left            =   4920
         TabIndex        =   1
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   55443457
         CurrentDate     =   39864
      End
      Begin MSComCtl2.DTPicker TxtCXTFDate 
         Height          =   360
         Left            =   12120
         TabIndex        =   3
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   55443457
         CurrentDate     =   40124
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
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1440
         Width           =   4440
      End
      Begin VB.ComboBox CmbMode 
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
         ItemData        =   "FrmTxn_IncomeDetails.frx":009D
         Left            =   1800
         List            =   "FrmTxn_IncomeDetails.frx":00B3
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   4440
      End
      Begin VB.ComboBox CmbTransactionType 
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
         ItemData        =   "FrmTxn_IncomeDetails.frx":0117
         Left            =   8520
         List            =   "FrmTxn_IncomeDetails.frx":0121
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   4920
      End
      Begin VB.ComboBox CmbClientIDRow 
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
         TabIndex        =   8
         Top             =   2040
         Width           =   11640
      End
      Begin VB.TextBox TxtClientIDRow 
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
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2040
         Width           =   11640
      End
      Begin VB.TextBox TxtMode 
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
         MaxLength       =   100
         TabIndex        =   52
         Top             =   840
         Width           =   4440
      End
      Begin VB.TextBox TxtTransactionType 
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
         Left            =   8520
         MaxLength       =   100
         TabIndex        =   51
         Top             =   840
         Width           =   4920
      End
      Begin VB.ComboBox CmbCommodityID 
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
         Left            =   8520
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   4920
      End
      Begin VB.TextBox TxtCommodityID 
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
         Left            =   8520
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1440
         Width           =   4920
      End
      Begin VB.Label Label3 
         Caption         =   "G_UID"
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
         Left            =   12840
         TabIndex        =   56
         Top             =   5340
         Visible         =   0   'False
         Width           =   720
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
         Height          =   300
         Left            =   6840
         TabIndex        =   49
         Top             =   1470
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "No of Bags"
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
         TabIndex        =   47
         Top             =   4020
         Width           =   1080
      End
      Begin VB.Label Label8 
         Caption         =   "Storage Charges"
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
         Left            =   6600
         TabIndex        =   45
         Top             =   3420
         Width           =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "Rate"
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
         Top             =   4620
         Width           =   945
      End
      Begin VB.Label Label14 
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
         Left            =   7080
         TabIndex        =   43
         Top             =   4020
         Width           =   1080
      End
      Begin VB.Label Label15 
         Caption         =   "PI NO"
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
         TabIndex        =   42
         Top             =   5220
         Width           =   945
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
         Height          =   300
         Left            =   120
         TabIndex        =   41
         Top             =   1470
         Width           =   960
      End
      Begin VB.Label LblStartDate 
         Caption         =   "Billing Month  Start Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   120
         TabIndex        =   40
         Top             =   158
         Width           =   1440
      End
      Begin VB.Label Label13 
         Caption         =   "Transaction Type"
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
         Left            =   6840
         TabIndex        =   39
         Top             =   840
         Width           =   1680
      End
      Begin VB.Label LblMode 
         Caption         =   "Mode"
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
         Left            =   120
         TabIndex        =   38
         Top             =   878
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "No of Days"
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
         TabIndex        =   37
         Top             =   3420
         Width           =   1320
      End
      Begin VB.Label Label5 
         Caption         =   "Amt Incl of Service Tax"
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
         Left            =   6240
         TabIndex        =   36
         Top             =   4620
         Width           =   2040
      End
      Begin VB.Label LblCloseDate 
         Caption         =   "Billing Month  End Date"
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
         Left            =   3480
         TabIndex        =   35
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   34
         Top             =   2070
         Width           =   1320
      End
      Begin VB.Label Label10 
         Caption         =   "CxTF NO"
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
         Left            =   6840
         TabIndex        =   33
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label11 
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
         TabIndex        =   32
         Top             =   2760
         Width           =   1200
      End
      Begin VB.Label Label9 
         Caption         =   "Amt Incl of Service Tax and Insurance"
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
         Left            =   4920
         TabIndex        =   31
         Top             =   5220
         Width           =   3360
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   19
      Top             =   5880
      Width           =   15045
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   225
         TabIndex        =   29
         Top             =   240
         Width           =   2880
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   225
         TabIndex        =   28
         Top             =   645
         Width           =   2880
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   3120
         TabIndex        =   27
         Top             =   240
         Width           =   2880
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   3105
         TabIndex        =   26
         Top             =   645
         Width           =   2880
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   5985
         TabIndex        =   25
         Top             =   240
         Width           =   2880
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   5985
         TabIndex        =   24
         Top             =   645
         Width           =   2880
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   8865
         TabIndex        =   23
         Top             =   240
         Width           =   2880
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   8865
         TabIndex        =   22
         Top             =   645
         Width           =   2880
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   11745
         TabIndex        =   21
         Top             =   240
         Width           =   2880
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   11745
         TabIndex        =   20
         Top             =   645
         Width           =   2880
      End
   End
End
Attribute VB_Name = "FrmTxn_IncomeDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mLocationID, mCommodityId, mClientIDRow As Variant
Dim mOld, mNew, mCxTFDate As Boolean


Private Sub ChkCxTFDate_Click()
If ChkCxTFDate.Value = vbChecked Then
   TxtCXTFDate.Enabled = True
   mCxTFDate = True
Else
   TxtCXTFDate.Enabled = False
   mCxTFDate = False
End If
End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
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
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
End Sub
Private Sub CmbCommodityID_GotFocus()

tmp = CmbCommodityID.Text
Call TableMst_Commodity
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Commodity.RecordCount
    CmbCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodityID.Text = tmp


End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   Exit Sub
End If

RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If
mCommodityId = RsMst_Commodity!CommodityID
End Sub

Private Sub CmbClientIDRow_GotFocus()
tmp = CmbClientIDRow.Text
'mExchangeTypeID = GetExchangeID("NON NCDEX")
Call TableMst_Client '("where ClientIdRow = " & mExchangeTypeID)
CmbClientIDRow.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For i = 1 To RsMst_Client.RecordCount
    CmbClientIDRow.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbClientIDRow.Text = tmp
End Sub

Private Sub CmbClientIDRow_LostFocus()
If CmbClientIDRow.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbClientIDRow.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection!!! "
   CmbClientIDRow.SetFocus
   Exit Sub
End If
mClientIDRow = RsMst_Client!ClientIDRow
End Sub

Private Sub CmbMode_LostFocus()
If CmbMode.Text = "" Then
   Exit Sub
End If
If LCase(CmbMode.Text) = "corporate" Then
    Exit Sub
ElseIf LCase(CmbMode.Text) = "ncdex withdrawal" Then
   Exit Sub
ElseIf LCase(CmbMode.Text) = "non-ncdex deposit" Then
   Exit Sub
ElseIf LCase(CmbMode.Text) = "non-ncdex op. bal" Then
   Exit Sub
ElseIf LCase(CmbMode.Text) = "non-ncdex withdrawal" Then
   Exit Sub
ElseIf LCase(CmbMode.Text) = "s & w" Then
   Exit Sub
Else
   MsgBox "Invalid Selection"
   CmbMode.SetFocus
   Exit Sub
End If
End Sub



Private Sub CmbTransactionType_LostFocus()

If CmbTransactionType.Text = "" Then
   Exit Sub
End If

If LCase(CmbTransactionType) = "deposit" Then
   Exit Sub
ElseIf LCase(CmbTransactionType) = "withdrawal" Then
   Exit Sub
Else
   MsgBox "Invalid Selection"
   CmbTransactionType.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdClientIDRow_Click()
FrmMst_Client.Show
End Sub

Private Sub CmdGo_Click()
If CmbSearchMode.Text = "" And CmbSearchTransactionType.Text = "" Then Exit Sub

'tmp = "Select Mode,TxnType,LocationID,BillingMonthStartDate,BillingMonthEndDate,ClientIDRow,"
'tmp = tmp & " CommodityID,OpeningNoofBags,StorageCharges,ServiceTax,AmtInclOfSerTaxnInsurance,"
'tmp = tmp & " PINO,Remarks,CxTFNO,CxTFDate,Rate,NoOfDays,ChargesCollected,AmtInclOfSerTax,G_UID"
'tmp = tmp & " From Txn_IncomeDetails"
MSHFlexGrid1.Clear
Call GridHead

tmp = "Select * from Txn_IncomeDetails "

If CmbSearchMode.Text <> "" And CmbSearchTransactionType.Text <> "" Then
   tmp = tmp & " where Mode = '" & CmbSearchMode.Text & "' and  TxnType = '" & CmbSearchTransactionType.Text & "' "
   Call TmpTable
ElseIf CmbSearchMode.Text <> "" Then
   tmp = tmp & " where Mode = '" & CmbSearchMode.Text & "'"
   Call TmpTable
ElseIf CmbSearchTransactionType.Text <> "" Then
   tmp = tmp & " Where TxnType = '" & CmbSearchTransactionType.Text & "'"
   Call TmpTable
End If

If TmpRs.RecordCount > 0 Then
   MsgBox " Wait till Next Message!!"
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 1
   Label22.Refresh
   Label22.Caption = TmpRs.RecordCount
   For i = 1 To TmpRs.RecordCount
       MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(TmpRs!Mode), "", TmpRs!Mode)
       MSHFlexGrid1.TextMatrix(i, 1) = IIf(IsNull(TmpRs!TxnType), "", TmpRs!TxnType)
       MSHFlexGrid1.TextMatrix(i, 2) = IIf(IsNull(TmpRs!LocationID), "", GetLocationName(TmpRs!LocationID))
       MSHFlexGrid1.TextMatrix(i, 3) = IIf(IsNull(TmpRs!BillingMonthStartDate), "", TmpRs!BillingMonthStartDate)
       MSHFlexGrid1.TextMatrix(i, 4) = IIf(IsNull(TmpRs!BillingMonthEndDate), "", TmpRs!BillingMonthEndDate)
       MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(TmpRs!ClientIDRow), "", GetClientName(TmpRs!ClientIDRow))
       If IsNull(TmpRs!CommodityID) = False Then
          MSHFlexGrid1.TextMatrix(i, 6) = GetCommodityName(TmpRs!CommodityID)
       Else
          MSHFlexGrid1.TextMatrix(i, 6) = ""
       End If
       MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(TmpRs!OpeningNoofBags), "", TmpRs!OpeningNoofBags)
       MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(TmpRs!StorageCharges), "", TmpRs!StorageCharges)
       MSHFlexGrid1.TextMatrix(i, 9) = IIf(IsNull(TmpRs!ServiceTax), "", TmpRs!ServiceTax)
       MSHFlexGrid1.TextMatrix(i, 10) = IIf(IsNull(TmpRs!AmtInclOfSerTaxnInsurance), "", TmpRs!AmtInclOfSerTaxnInsurance)
       If IsNull(TmpRs!OldPINO) = False Then
          MSHFlexGrid1.TextMatrix(i, 11) = (TmpRs!OldPINO)
       ElseIf IsNull(TmpRs!NewPINO) = False Then
          MSHFlexGrid1.TextMatrix(i, 11) = (TmpRs!NewPINO)
       End If
       'MSHFlexGrid1.TextMatrix(0, 11) = IIf(IsNull(RsTxn_IncomeDetails!PINO), "", RsTxn_IncomeDetails!PINO)
       MSHFlexGrid1.TextMatrix(i, 12) = IIf(IsNull(TmpRs!Remarks), "", TmpRs!Remarks)
       MSHFlexGrid1.TextMatrix(i, 13) = IIf(IsNull(TmpRs!CxTFNO), "", TmpRs!CxTFNO)
       MSHFlexGrid1.TextMatrix(i, 14) = IIf(IsNull(TmpRs!CxTFDate), "", TmpRs!CxTFDate)
       MSHFlexGrid1.TextMatrix(i, 15) = IIf(IsNull(TmpRs!Rate), "", TmpRs!Rate)
       MSHFlexGrid1.TextMatrix(i, 16) = IIf(IsNull(TmpRs!NoOfDaysChargesCollected), "", TmpRs!NoOfDaysChargesCollected)
       MSHFlexGrid1.TextMatrix(i, 17) = IIf(IsNull(TmpRs!AmtInclOfSerTax), "", TmpRs!AmtInclOfSerTax)
       MSHFlexGrid1.TextMatrix(i, 18) = IIf(IsNull(TmpRs!G_UID), "", TmpRs!G_UID)
       TmpRs.MoveNext
   Next
   MsgBox "Done!!"
Else
   MsgBox "Record Not Found!!!"
End If

End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

ChkCxTFDate.Enabled = True
TxtCXTFDate.Enabled = False
TxtBillingMonthStartDate.Value = Now
TxtBillingMonthEndDate.Value = Now
TxtCXTFDate.Value = Now

TxtMode.Visible = False
TxtTransactionType.Visible = False
TxtLocationID.Visible = False
TxtCommodityID.Visible = False
TxtClientIDRow.Visible = False

CmbMode.Visible = True
CmbTransactionType.Visible = True
CmbLocationID.Visible = True
CmbCommodityID.Visible = True
CmbClientIDRow.Visible = True

CmbMode.Text = ""
CmbTransactionType.Text = ""
CmbLocationID.Text = ""
CmbCommodityID.Text = ""
CmbClientIDRow.Text = ""

TxtCXTFNO.Text = ""
TxtRemarks.Text = ""
TxtNoOfdays.Text = ""
TxtNoOfBags.Text = ""
TxtRate.Text = ""
TxtPINO.Text = ""
TxtStorageCharges.Text = ""
TxtServiceTax.Text = ""
TxtAmtInclOfServiceTax.Text = ""
TxtAmtInclOfSerTaxInsurance.Text = ""

TxtCXTFDate.Visible = True
Frame1.Enabled = True

If RsTxn_IncomeDetails.RecordCount > 0 Then
   TxtBillingMonthStartDate.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub DeleteCmd_Click()
Dim Ans As Variant
On Error GoTo msgError
Ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If Ans = vbYes Then
   RsTxn_IncomeDetails.Delete
   RsTxn_IncomeDetails.Update
   'MsgBox "Record Successfully Removed"
   If RsTxn_IncomeDetails.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsTxn_IncomeDetails.MoveNext
   If RsTxn_IncomeDetails.EOF() Then
      RsTxn_IncomeDetails.MoveLast
   End If
   Call InData
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsTxn_IncomeDetails.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
ChkCxTFDate.Enabled = True
TxtCXTFDate.Enabled = True


TxtMode.Visible = False
TxtTransactionType.Visible = False
TxtLocationID.Visible = False
TxtCommodityID.Visible = False
TxtClientIDRow.Visible = False

CmbMode.Visible = True
CmbTransactionType.Visible = True
CmbLocationID.Visible = True
CmbCommodityID.Visible = True
CmbClientIDRow.Visible = True

CmbMode.Text = TxtMode.Text
CmbTransactionType.Text = TxtTransactionType.Text
CmbLocationID.Text = TxtLocationID.Text
CmbCommodityID.Text = TxtCommodityID.Text
CmbClientIDRow.Text = TxtClientIDRow.Text




Call GButtonLock(Me, False)
'TxtBranchName.Locked = False

CmdClientIDRow.Visible = True
Frame1.Enabled = True
TxtBillingMonthStartDate.SetFocus
End Sub


Private Sub FirstCmd_Click()
RsTxn_IncomeDetails.MoveFirst
Call InData
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
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call TableTxn_IncomeDetails '("  Order by BankID,BranchName ")

If RsTxn_IncomeDetails.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call InData
Call GButtonLock(Me, True)
End Sub

Public Sub InData()
Frame1.Enabled = False
mOld = False
mNew = False
'mBankID = RsTxn_IncomeDetails!BankID
mLocationID = RsTxn_IncomeDetails!LocationID
mCommodityId = RsTxn_IncomeDetails!CommodityID
mClientIDRow = RsTxn_IncomeDetails!ClientIDRow

TxtMode.Text = IIf(IsNull(RsTxn_IncomeDetails!Mode), "", RsTxn_IncomeDetails!Mode)
TxtTransactionType.Text = IIf(IsNull(RsTxn_IncomeDetails!TxnType), "", RsTxn_IncomeDetails!TxnType)
TxtLocationID.Text = IIf(IsNull(RsTxn_IncomeDetails!LocationID), "", GetLocationName(RsTxn_IncomeDetails!LocationID))
TxtBillingMonthStartDate.Value = IIf(IsNull(RsTxn_IncomeDetails!BillingMonthStartDate), "", RsTxn_IncomeDetails!BillingMonthStartDate)
TxtBillingMonthEndDate.Value = IIf(IsNull(RsTxn_IncomeDetails!BillingMonthEndDate), "", RsTxn_IncomeDetails!BillingMonthEndDate)
TxtClientIDRow.Text = IIf(IsNull(RsTxn_IncomeDetails!ClientIDRow), "", GetClientName(RsTxn_IncomeDetails!ClientIDRow))
If IsNull(RsTxn_IncomeDetails!CommodityID) = False Then
   TxtCommodityID.Text = GetCommodityName(RsTxn_IncomeDetails!CommodityID)
End If
TxtNoOfBags.Text = IIf(IsNull(RsTxn_IncomeDetails!OpeningNoofBags), "", RsTxn_IncomeDetails!OpeningNoofBags)
TxtStorageCharges.Text = IIf(IsNull(RsTxn_IncomeDetails!StorageCharges), "", RsTxn_IncomeDetails!StorageCharges)
TxtServiceTax.Text = IIf(IsNull(RsTxn_IncomeDetails!ServiceTax), "", RsTxn_IncomeDetails!ServiceTax)
TxtAmtInclOfSerTaxInsurance.Text = IIf(IsNull(RsTxn_IncomeDetails!AmtInclOfSerTaxnInsurance), "", RsTxn_IncomeDetails!AmtInclOfSerTaxnInsurance)
If IsNull(RsTxn_IncomeDetails!OldPINO) = False Then
   TxtPINO.Text = RsTxn_IncomeDetails!OldPINO
   mOld = True
ElseIf IsNull(RsTxn_IncomeDetails!NewPINO) = False Then
   TxtPINO.Text = RsTxn_IncomeDetails!NewPINO
   mNew = True
Else
   TxtPINO.Text = ""
   mNew = True
End If
TxtRemarks.Text = IIf(IsNull(RsTxn_IncomeDetails!Remarks), "", RsTxn_IncomeDetails!Remarks)
TxtCXTFNO.Text = IIf(IsNull(RsTxn_IncomeDetails!CxTFNO), "", RsTxn_IncomeDetails!CxTFNO)
If IsNull(RsTxn_IncomeDetails!CxTFDate) = False Then
   TxtCXTFDate.Value = RsTxn_IncomeDetails!CxTFDate
   TxtCXTFDate.Enabled = True
   ChkCxTFDate.Enabled = True
   ChkCxTFDate.Value = vbChecked
Else
   TxtCXTFDate.Enabled = False
   ChkCxTFDate.Enabled = False
   ChkCxTFDate.Value = 0
End If
TxtRate.Text = IIf(IsNull(RsTxn_IncomeDetails!Rate), "", RsTxn_IncomeDetails!Rate)
TxtNoOfdays.Text = IIf(IsNull(RsTxn_IncomeDetails!NoOfDaysChargesCollected), "", RsTxn_IncomeDetails!NoOfDaysChargesCollected)
TxtAmtInclOfServiceTax.Text = IIf(IsNull(RsTxn_IncomeDetails!AmtInclOfSerTax), "", RsTxn_IncomeDetails!AmtInclOfSerTax)
TxtG_UID = IIf(IsNull(RsTxn_IncomeDetails!G_UID), "", RsTxn_IncomeDetails!G_UID)

TxtMode.Visible = True
TxtTransactionType.Visible = True
TxtLocationID.Visible = True
TxtCommodityID.Visible = True
TxtClientIDRow.Visible = True

CmbMode.Visible = False
CmbTransactionType.Visible = False
CmbLocationID.Visible = False
CmbCommodityID.Visible = False
CmbClientIDRow.Visible = False

CmdClientIDRow.Visible = False

Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsTxn_IncomeDetails.MoveLast
Call InData
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
If MSHFlexGrid1.TextMatrix(x, 0) = "" And MSHFlexGrid1.TextMatrix(x, 1) = "" Then Exit Sub
mCaId = MSHFlexGrid1.TextMatrix(x, 18)

RsTxn_IncomeDetails.MoveFirst

If RsTxn_IncomeDetails.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsTxn_IncomeDetails.MoveFirst
RsTxn_IncomeDetails.Find "G_UID = " & Val(mCaId) & "" ''''Changesssss
If RsTxn_IncomeDetails.EOF Then
   RsTxn_IncomeDetails.MoveLast
End If
Frame0.Visible = False
Call InData

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_IncomeDetails.MoveNext
If RsTxn_IncomeDetails.EOF Then
   RsTxn_IncomeDetails.MoveLast
   LF_Flag = "Y"
End If
Call InData
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
RsTxn_IncomeDetails.MovePrevious
If RsTxn_IncomeDetails.BOF Then
   RsTxn_IncomeDetails.MoveFirst
   LF_Flag = "Y"
End If
Call InData
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

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location Name not Allowed"
   CmbLocationID.SetFocus
   Exit Sub
End If
'If CmbCommodityID.Text = "" Then
'   MsgBox "Blank Commodity not Allowed"
'   CmbCommodityID.SetFocus
'   Exit Sub
'End If
'If CmbClientIDRow.Text = "" Then
'   MsgBox "Blank Client Name not Allowed"
'   CmbClientIDRow.SetFocus
'   Exit Sub
'End If
'If TxtNoofBags.Text = "" Then
'   MsgBox "Blank Bag Number not Allowed"
'   TxtNoofBags.SetFocus
'   Exit Sub
'End If


If Edit_Flg = "A" Then
'   If RsTxn_IncomeDetails.RecordCount > 0 Then
'      RsTxn_IncomeDetails.MoveFirst
'      RsTxn_IncomeDetails.Find "BranchName = '" & TxtBranchName & "'"
'      If Not RsTxn_IncomeDetails.EOF Then
'         MsgBox "Duplicate Branch Name not Allowed"
'         Exit Sub
'      End If
'   End If
   mNew = True
   RsTxn_IncomeDetails.AddNew
   RsTxn_IncomeDetails!G_UID = GetGUID
   TxtG_UID = RsTxn_IncomeDetails!G_UID
Else
'   RsTxn_IncomeDetails.MoveFirst
'   RsTxn_IncomeDetails.Find "gu= '" & TxtBranchName & "'"
'   If Not RsTxn_IncomeDetails.EOF Then
'      If RsTxn_IncomeDetails!BranchID <> TxtBranchID Then
'         MsgBox "Duplicate Branch Name not Allowed"
'         Exit Sub
'      End If
'   End If
'   RsTxn_IncomeDetails.MoveFirst
   RsTxn_IncomeDetails.Find "G_UID = '" & TxtG_UID & "'"
End If

RsTxn_IncomeDetails!Mode = CmbMode.Text
RsTxn_IncomeDetails!TxnType = CmbTransactionType.Text
RsTxn_IncomeDetails!LocationID = mLocationID
RsTxn_IncomeDetails!BillingMonthStartDate = TxtBillingMonthStartDate.Value
RsTxn_IncomeDetails!BillingMonthEndDate = TxtBillingMonthEndDate.Value
RsTxn_IncomeDetails!ClientIDRow = mClientIDRow
RsTxn_IncomeDetails!CommodityID = mCommodityId
RsTxn_IncomeDetails!OpeningNoofBags = IIf(TxtNoOfBags.Text = "", 0, TxtNoOfBags.Text)
RsTxn_IncomeDetails!StorageCharges = IIf(TxtStorageCharges.Text = "", 0, TxtStorageCharges.Text)
RsTxn_IncomeDetails!ServiceTax = IIf(TxtServiceTax.Text = "", 0, TxtServiceTax.Text)
RsTxn_IncomeDetails!AmtInclOfSerTaxnInsurance = IIf(TxtAmtInclOfSerTaxInsurance.Text = "", 0, TxtAmtInclOfSerTaxInsurance.Text)
If TxtPINO.Text <> "" Then
   If mOld = True Then
      RsTxn_IncomeDetails!OldPINO = TxtPINO.Text
   ElseIf mNew = True Then
      RsTxn_IncomeDetails!NewPINO = TxtPINO.Text
   ElseIf TxtPINO.Text = "" Then
      RsTxn_IncomeDetails!NewPINO = GetMaxNewPINo
   End If
End If
RsTxn_IncomeDetails!Remarks = TxtRemarks.Text
RsTxn_IncomeDetails!CxTFNO = IIf(TxtCXTFNO.Text = "", 0, TxtCXTFNO.Text)
RsTxn_IncomeDetails!CxTFDate = IIf(mCxTFDate = True, TxtCXTFDate.Value, Null)
RsTxn_IncomeDetails!Rate = IIf(TxtRate.Text = "", 0, TxtRate.Text)
RsTxn_IncomeDetails!NoOfDaysChargesCollected = IIf(TxtNoOfdays.Text = "", 0, TxtNoOfdays.Text)
RsTxn_IncomeDetails!AmtInclOfSerTax = IIf(TxtAmtInclOfServiceTax.Text = "", 0, TxtAmtInclOfServiceTax.Text)



RsTxn_IncomeDetails.Update
RsTxn_IncomeDetails.Requery

'RsTxn_IncomeDetails.MoveFirst
RsTxn_IncomeDetails.Find "G_UID = '" & TxtG_UID & "'"
If RsTxn_IncomeDetails.EOF Then
   RsTxn_IncomeDetails.MoveFirst
End If
Call GButtonLock(Me, True)

Call InData
End Sub
Private Function GetMaxNewPINo() As Double
Dim RetVal As Double
tmp = "Select Max(NewPINo) From Txn_IncomeDetails "
Call Tmp1Table
RetVal = IIf(IsNull(TmpRs1.Fields(0)), 0, TmpRs1.Fields(0))
RetVal = RetVal + 1
TmpRs1.Close
GetMaxNewPINo = RetVal
End Function

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsTxn_IncomeDetails.EOF Then
      RsTxn_IncomeDetails.MoveFirst
   End If
   Call InData
   Exit Sub
End If

Frame0.Height = Me.Height - 250
Frame0.Left = Frame1.Left - 120
Frame0.Top = Frame1.Top '120
Frame0.Width = Me.Width - 100
Frame0.Visible = True
MSHFlexGrid1.Clear

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1200

RsTxn_IncomeDetails.MoveFirst
Call GridHead
CmbSearchMode.Text = ""
CmbSearchTransactionType.Text = ""
Label22.Caption = ""

'MSHFlexGrid1.Rows = RsTxn_IncomeDetails.RecordCount + 1

'For i = 1 To RsTxn_IncomeDetails.RecordCount
'    MSHFlexGrid1.TextMatrix(i, 0) = IIf(IsNull(RsTxn_IncomeDetails!Mode), "", RsTxn_IncomeDetails!Mode)
'    MSHFlexGrid1.TextMatrix(i, 1) = IIf(IsNull(RsTxn_IncomeDetails!TxnType), "", RsTxn_IncomeDetails!TxnType)
'    MSHFlexGrid1.TextMatrix(i, 2) = IIf(IsNull(RsTxn_IncomeDetails!LocationID), "", GetLocationName(RsTxn_IncomeDetails!LocationID))
'    MSHFlexGrid1.TextMatrix(i, 3) = IIf(IsNull(RsTxn_IncomeDetails!BillingMonthStartDate), "", RsTxn_IncomeDetails!BillingMonthStartDate)
'    MSHFlexGrid1.TextMatrix(i, 4) = IIf(IsNull(RsTxn_IncomeDetails!BillingMonthEndDate), "", RsTxn_IncomeDetails!BillingMonthEndDate)
'    MSHFlexGrid1.TextMatrix(i, 5) = IIf(IsNull(RsTxn_IncomeDetails!ClientIDRow), "", GetClientName(RsTxn_IncomeDetails!ClientIDRow))
'    If IsNull(RsTxn_IncomeDetails!CommodityID) = False Then
'       MSHFlexGrid1.TextMatrix(i, 6) = GetCommodityName(RsTxn_IncomeDetails!CommodityID)
'    Else
'       MSHFlexGrid1.TextMatrix(i, 6) = ""
'    End If
'    MSHFlexGrid1.TextMatrix(i, 7) = IIf(IsNull(RsTxn_IncomeDetails!OpeningNoofBags), "", RsTxn_IncomeDetails!OpeningNoofBags)
'    MSHFlexGrid1.TextMatrix(i, 8) = IIf(IsNull(RsTxn_IncomeDetails!StorageCharges), "", RsTxn_IncomeDetails!StorageCharges)
'    MSHFlexGrid1.TextMatrix(i, 9) = IIf(IsNull(RsTxn_IncomeDetails!ServiceTax), "", RsTxn_IncomeDetails!ServiceTax)
'    MSHFlexGrid1.TextMatrix(i, 10) = IIf(IsNull(RsTxn_IncomeDetails!AmtInclOfSerTaxnInsurance), "", RsTxn_IncomeDetails!AmtInclOfSerTaxnInsurance)
'    If IsNull(RsTxn_IncomeDetails!OldPINO) = False Then
'       MSHFlexGrid1.TextMatrix(i, 11) = (RsTxn_IncomeDetails!OldPINO)
'    ElseIf IsNull(RsTxn_IncomeDetails!NewPINO) = False Then
'       MSHFlexGrid1.TextMatrix(i, 11) = (RsTxn_IncomeDetails!NewPINO)
'    End If
'    'MSHFlexGrid1.TextMatrix(0, 11) = IIf(IsNull(RsTxn_IncomeDetails!PINO), "", RsTxn_IncomeDetails!PINO)
'    MSHFlexGrid1.TextMatrix(i, 12) = IIf(IsNull(RsTxn_IncomeDetails!Remarks), "", RsTxn_IncomeDetails!Remarks)
'    MSHFlexGrid1.TextMatrix(i, 13) = IIf(IsNull(RsTxn_IncomeDetails!CxTFNO), "", RsTxn_IncomeDetails!CxTFNO)
'    MSHFlexGrid1.TextMatrix(i, 14) = IIf(IsNull(RsTxn_IncomeDetails!CxTFDate), "", RsTxn_IncomeDetails!CxTFDate)
'    MSHFlexGrid1.TextMatrix(i, 15) = IIf(IsNull(RsTxn_IncomeDetails!Rate), "", RsTxn_IncomeDetails!Rate)
'    MSHFlexGrid1.TextMatrix(i, 16) = IIf(IsNull(RsTxn_IncomeDetails!NoOfDaysChargesCollected), "", RsTxn_IncomeDetails!NoOfDaysChargesCollected)
'    MSHFlexGrid1.TextMatrix(i, 17) = IIf(IsNull(RsTxn_IncomeDetails!AmtInclOfSerTax), "", RsTxn_IncomeDetails!AmtInclOfSerTax)
'    MSHFlexGrid1.TextMatrix(i, 18) = IIf(IsNull(RsTxn_IncomeDetails!G_UID), "", RsTxn_IncomeDetails!G_UID)
'    RsTxn_IncomeDetails.MoveNext
'Next


End Sub
Private Sub GridHead()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 19
MSHFlexGrid1.TextMatrix(0, 0) = "Mode"
MSHFlexGrid1.TextMatrix(0, 1) = "TxnType"
MSHFlexGrid1.TextMatrix(0, 2) = "Location"
MSHFlexGrid1.TextMatrix(0, 3) = "BillingMonthStartDate"
MSHFlexGrid1.TextMatrix(0, 4) = "BillingMonthEndDate"
MSHFlexGrid1.TextMatrix(0, 5) = "Client"
MSHFlexGrid1.TextMatrix(0, 6) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 7) = "No of Bags"
MSHFlexGrid1.TextMatrix(0, 8) = "Storage Charges"
MSHFlexGrid1.TextMatrix(0, 9) = "Service Tax"
MSHFlexGrid1.TextMatrix(0, 10) = "Amt Incl Of SerTax & Insurance"
MSHFlexGrid1.TextMatrix(0, 11) = "PI NO"
MSHFlexGrid1.TextMatrix(0, 12) = "Remarks"
MSHFlexGrid1.TextMatrix(0, 13) = "CxTF NO"
MSHFlexGrid1.TextMatrix(0, 14) = "CxTF Date"
MSHFlexGrid1.TextMatrix(0, 15) = "Rate"
MSHFlexGrid1.TextMatrix(0, 16) = "No Of Days "
MSHFlexGrid1.TextMatrix(0, 17) = "Amt Incl Of Service Tax"
MSHFlexGrid1.TextMatrix(0, 18) = "G_UID"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 1200
MSHFlexGrid1.ColWidth(4) = 1200
MSHFlexGrid1.ColWidth(5) = 2000
MSHFlexGrid1.ColWidth(6) = 1500
MSHFlexGrid1.ColWidth(7) = 1000
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 1000
MSHFlexGrid1.ColWidth(12) = 2000
MSHFlexGrid1.ColWidth(13) = 1000
MSHFlexGrid1.ColWidth(14) = 1000
MSHFlexGrid1.ColWidth(15) = 1000
MSHFlexGrid1.ColWidth(16) = 1000
MSHFlexGrid1.ColWidth(17) = 1200
MSHFlexGrid1.ColWidth(18) = 0
End Sub
Private Sub TxtAmtInclOfSerTaxInsurance_LostFocus()
If TxtAmtInclOfSerTaxInsurance = "" Then Exit Sub
If IsNumeric(TxtAmtInclOfSerTaxInsurance) = False Or Val(TxtAmtInclOfSerTaxInsurance) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtAmtInclOfSerTaxInsurance.SetFocus
   Exit Sub
End If
End Sub
Private Sub TxtAmtInclOfServiceTax_LostFocus()
If TxtAmtInclOfServiceTax = "" Then Exit Sub
If IsNumeric(TxtAmtInclOfServiceTax) = False Or Val(TxtAmtInclOfServiceTax) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtAmtInclOfServiceTax.SetFocus
   Exit Sub
End If
End Sub
Private Sub TxtBillingMonthStartDate_LostFocus()
Sdate = Month(TxtBillingMonthStartDate.Value) & "/" & "1" & "/" & Year(TxtBillingMonthStartDate.Value)
TxtBillingMonthStartDate.Value = Format(Sdate, "MM/DD/YYYY")
TxtBillingMonthEndDate.Value = GetEndDate(Month(TxtBillingMonthStartDate.Value), TxtBillingMonthStartDate)
End Sub
Private Sub TxtBillingMonthEndDate_LostFocus()
'If TxtBillingMonthEndDate.Value < TxtBillingMonthStartDate.Value Then
'   MsgBox "Billing End Date Should Not Be Less Than Billing Start Date!!! "
'   TxtBillingMonthEndDate.SetFocus
'Else
'   'TxtBillingMonthEndDate.Value = GetMonthEnd(TxtBillingMonthEndDate.Value)
'   TxtBillingMonthEndDate.Value = GetEndDate(Month(TxtBillingMonthStartDate.Value), (TxtBillingMonthStartDate))
'End If
End Sub
Private Sub TxtNoOfBags_LostFocus()
If TxtNoOfBagst = "" Then Exit Sub
If IsNumeric(TxtNoOfBags) = False Or Val(TxtNoOfBags) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtNoOfBags.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtNoOfdays_LostFocus()
If TxtNoOfdays = "" Then Exit Sub
If IsNumeric(TxtNoOfdays) = False Or Val(TxtNoOfdays) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtNoOfdays.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtPINO_LostFocus()
If TxtPINO = "" Then Exit Sub
If IsNumeric(TxtPINO) = False Or Val(TxtPINO) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtPINO.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtRate_LostFocus()
If TxtRate = "" Then Exit Sub
If IsNumeric(TxtRate) = False Or Val(TxtRate) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtRate.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtServiceTax_LostFocus()
If TxtServiceTax = "" Then Exit Sub
If IsNumeric(TxtServiceTax) = False Or Val(TxtServiceTax) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtServiceTax.SetFocus
   Exit Sub
End If
End Sub

Private Sub TxtStorageCharges_LostFocus()
If TxtStorageCharges = "" Then Exit Sub
If IsNumeric(TxtStorageCharges) = False Or Val(TxtStorageCharges) < 0 Then
   MsgBox "Invalid numeric format!!!"
   TxtStorageCharges.SetFocus
   Exit Sub
End If
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_IncomeDetails.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_IncomeDetails.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

