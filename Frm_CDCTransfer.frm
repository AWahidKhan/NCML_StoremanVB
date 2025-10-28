VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_CDCTransfer 
   Caption         =   "CDC Payout"
   ClientHeight    =   8835
   ClientLeft      =   -3000
   ClientTop       =   555
   ClientWidth     =   14715
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   14715
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   7260
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   14460
      Begin VB.TextBox TxtCommodityCondition 
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
         TabIndex        =   79
         Top             =   1500
         Width           =   1830
      End
      Begin VB.TextBox TxtNoofDays 
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
         Left            =   13800
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame2 
         Height          =   3600
         Left            =   120
         TabIndex        =   39
         Top             =   2760
         Width           =   14175
         Begin VB.CommandButton CmdClient 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3720
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   2640
            Width           =   330
         End
         Begin VB.TextBox TxtGodownID 
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
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   870
            Width           =   2070
         End
         Begin VB.TextBox TxtStackId 
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
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   1305
            Width           =   2070
         End
         Begin VB.TextBox TxtLotId 
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
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1740
            Width           =   2070
         End
         Begin VB.TextBox TxtNoOfBags 
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
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   2175
            Width           =   2070
         End
         Begin VB.TextBox TxtQuantity 
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
            Left            =   5775
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   870
            Width           =   825
         End
         Begin VB.TextBox TxtNewDPID 
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
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   3030
            Width           =   2070
         End
         Begin VB.TextBox TxtNewClientId 
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
            TabIndex        =   2
            Top             =   2595
            Width           =   1860
         End
         Begin VB.TextBox TxtNewDPName 
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
            Left            =   5775
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   3030
            Width           =   8250
         End
         Begin VB.ComboBox CmbClientDP 
            Height          =   315
            Left            =   5775
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   2595
            Visible         =   0   'False
            Width           =   8280
         End
         Begin VB.TextBox TxtStorageRate 
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
            Left            =   12735
            TabIndex        =   4
            Top             =   870
            Width           =   1275
         End
         Begin VB.TextBox TxtNetWeight 
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
            Left            =   5775
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   2175
            Width           =   2265
         End
         Begin VB.TextBox TxtStdDeduction 
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
            Left            =   5775
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1740
            Width           =   2265
         End
         Begin VB.TextBox TxtGunnyWeight 
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
            Left            =   5775
            Locked          =   -1  'True
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   1305
            Width           =   2265
         End
         Begin VB.TextBox TxtGodownAddress 
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
            Height          =   690
            Left            =   1815
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   120
            Width           =   12255
         End
         Begin VB.TextBox TxtStorageAmount 
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
            Left            =   10095
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   2175
            Width           =   2265
         End
         Begin VB.TextBox TxtQuantityKG 
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
            Left            =   6735
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   870
            Width           =   1305
         End
         Begin MSComCtl2.DTPicker StorageFrom 
            Height          =   375
            Left            =   10095
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   1305
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   57475073
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker StorageTo 
            Height          =   375
            Left            =   12735
            TabIndex        =   54
            Top             =   1305
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   57475073
            CurrentDate     =   39884
         End
         Begin VB.TextBox TxtNewClientName 
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
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   2640
            Width           =   8250
         End
         Begin VB.Label lblNameDid 
            Caption         =   "ClientDPID"
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
            Left            =   4080
            TabIndex        =   75
            Top             =   3285
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Storage From"
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
            Left            =   8295
            TabIndex        =   73
            Top             =   1365
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "Payout Date"
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
            Left            =   11415
            TabIndex        =   72
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label LblStorageDays 
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
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
            Left            =   10095
            TabIndex        =   71
            Top             =   1800
            Width           =   75
         End
         Begin VB.Label Label27 
            Caption         =   "Additional Storage Days :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   8280
            TabIndex        =   70
            Top             =   1770
            Width           =   1710
         End
         Begin VB.Label Label13 
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
            Height          =   255
            Left            =   135
            TabIndex        =   69
            Top             =   930
            Width           =   1455
         End
         Begin VB.Label Label30 
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
            Height          =   255
            Left            =   135
            TabIndex        =   68
            Top             =   1365
            Width           =   1455
         End
         Begin VB.Label Label31 
            Caption         =   "Lot No"
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
            TabIndex        =   67
            Top             =   1800
            Width           =   750
         End
         Begin VB.Label Label44 
            Caption         =   "No Of Bags"
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
            TabIndex        =   66
            Top             =   2235
            Width           =   1215
         End
         Begin VB.Label Label45 
            Caption         =   "Qty in MT - KG"
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
            Left            =   4095
            TabIndex        =   65
            Top             =   930
            Width           =   1455
         End
         Begin VB.Label Label46 
            Caption         =   "Trading Member Name "
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
            Left            =   4095
            TabIndex        =   64
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Label Label48 
            Caption         =   "Trading Member ID "
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
            Left            =   120
            TabIndex        =   63
            Top             =   3030
            Width           =   1455
         End
         Begin VB.Label Label49 
            Caption         =   "Buyer Client Name"
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
            Left            =   4095
            TabIndex        =   62
            Top             =   2670
            Width           =   1695
         End
         Begin VB.Label Label50 
            Caption         =   "Buyer Client ID"
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
            Left            =   135
            TabIndex        =   61
            Top             =   2655
            Width           =   1605
         End
         Begin VB.Label Label4 
            Caption         =   "Storage Rate per MT per day"
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
            Left            =   10065
            TabIndex        =   60
            Top             =   930
            Width           =   3495
         End
         Begin VB.Label Label52 
            Caption         =   "Net Weight"
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
            Left            =   4095
            TabIndex        =   59
            Top             =   2235
            Width           =   1215
         End
         Begin VB.Label Label53 
            Caption         =   "Std Deduction"
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
            Left            =   4095
            TabIndex        =   58
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label54 
            Caption         =   "Gunny Weight"
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
            Left            =   4095
            TabIndex        =   57
            Top             =   1365
            Width           =   1455
         End
         Begin VB.Label Label55 
            Caption         =   "Godown Address"
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
            TabIndex        =   56
            Top             =   180
            Width           =   1575
         End
         Begin VB.Label Label56 
            Caption         =   "Storage Amount"
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
            Left            =   8280
            TabIndex        =   55
            Top             =   2265
            Width           =   1575
         End
      End
      Begin VB.TextBox TxtCDTFDate 
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
         Left            =   12840
         TabIndex        =   37
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TxtCDTFNo 
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
         Left            =   11040
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   615
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2355
         Width           =   5880
      End
      Begin VB.TextBox TxtCMPID 
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
         Top             =   1920
         Width           =   5865
      End
      Begin VB.TextBox TxtCMPAddress 
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
         Height          =   825
         Left            =   9105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1920
         Width           =   4575
      End
      Begin VB.TextBox TxtCDCID 
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
         Left            =   6120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtDPName 
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1080
         Width           =   7575
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   660
         Width           =   1830
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
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1500
         Width           =   7590
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
         TabIndex        =   17
         Top             =   645
         Width           =   7575
      End
      Begin VB.TextBox TxtDPID 
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
         Top             =   1080
         Width           =   1830
      End
      Begin MSComCtl2.DTPicker CDCDate 
         Height          =   375
         Left            =   8520
         TabIndex        =   21
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   57475073
         CurrentDate     =   39884
      End
      Begin VB.ComboBox CmbPreFix 
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
         TabIndex        =   81
         Top             =   240
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.TextBox TxtPreFix 
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
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label5 
         Caption         =   "Location ID"
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
         TabIndex        =   80
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Condition Of Commodity"
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
         Left            =   120
         TabIndex        =   78
         Top             =   1455
         Width           =   1695
      End
      Begin VB.Label Label58 
         Caption         =   "CDTF Date"
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
         Left            =   11640
         TabIndex        =   38
         Top             =   300
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label57 
         Caption         =   "CDTF NO"
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
         Left            =   9960
         TabIndex        =   36
         Top             =   300
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label51 
         Caption         =   "WH Address (CMP)"
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
         Left            =   7920
         TabIndex        =   34
         Top             =   1920
         Width           =   1455
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
         Left            =   135
         TabIndex        =   33
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "WH Name (CMP)"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1965
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Trading Member ID"
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
         TabIndex        =   28
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Trading Member Name"
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
         Left            =   3960
         TabIndex        =   27
         Top             =   1080
         Width           =   2160
      End
      Begin VB.Label Label9 
         Caption         =   "CDC NO"
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
         Left            =   3960
         TabIndex        =   26
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Seller Client ID"
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
         TabIndex        =   25
         Top             =   705
         Width           =   1590
      End
      Begin VB.Label Label26 
         Caption         =   "CDC Date"
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
         Left            =   7560
         TabIndex        =   24
         Top             =   300
         Width           =   1215
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
         Left            =   3975
         TabIndex        =   23
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Seller Client Name"
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
         Left            =   3960
         TabIndex        =   22
         Top             =   705
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   14430
      TabIndex        =   0
      Top             =   7320
      Width           =   14490
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   11580
         TabIndex        =   6
         Top             =   495
         Width           =   2775
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   8805
         TabIndex        =   14
         Top             =   495
         Width           =   2775
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   5910
         TabIndex        =   13
         Top             =   495
         Width           =   2895
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   3015
         TabIndex        =   12
         Top             =   495
         Width           =   2895
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   495
         Width           =   2895
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   11580
         TabIndex        =   10
         Top             =   120
         Width           =   2775
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   8820
         TabIndex        =   9
         Top             =   120
         Width           =   2760
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   5925
         TabIndex        =   8
         Top             =   120
         Width           =   2880
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   3015
         TabIndex        =   5
         Top             =   120
         Width           =   2895
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Frm_CDCTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCDCID As Double
Dim mClientRowID As Double
Dim mNewClientRowID As Double
Dim mCMPID As Double
Dim mCommodityID As Double
Dim mExchangeTypeID As Double
Dim mGSLID As Double
Dim mNewCDCNo As Double
Dim mCDTFNo As Double
Dim mCDTFDate As Date
Dim Edit_Flg, mSM_Prefix As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call TableMst_Config
CmbPreFix.Visible = True
TxtPreFix.Visible = False

TxtNoofDays = IIf(IsNull(RsMst_Config!SportStorageChargesDays), 0, RsMst_Config!SportStorageChargesDays)
Call TableTxn_CDC("Where Flag = 'O'")
Call ClearData
Call GButtonLock(Me, False)
'CmdPrint.Enabled = False
TxtCDCID.Locked = False
TxtStorageRate.Locked = False
TxtNewClientID.Locked = False
If RsTxn_CDC.RecordCount <> 0 Then
' TxtCDCID.SetFocus
End If


End Sub

Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text
Call TableMst_StoreManLocations("")

CmbPreFix.Clear
If RsMst_StoreManLocations.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbPreFix.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_StoreManLocations.RecordCount
    CmbPreFix.AddItem RsMst_StoreManLocations!SM_Prefix
    RsMst_StoreManLocations.MoveNext
Next
CmbPreFix.Text = tmp
End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   Exit Sub
End If
RsMst_StoreManLocations.MoveFirst
RsMst_StoreManLocations.Find "SM_Prefix = '" & CmbPreFix.Text & "'"

If RsMst_StoreManLocations.EOF Then
   MsgBox "Invalid PreFix Selection !!!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
mSM_Prefix = RsMst_StoreManLocations!SM_Prefix

End Sub

Private Sub CmdClient_Click()
FrmMst_Client.Show
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant

ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbNo Then
   Exit Sub
End If

Call TableMst_GSL("Where  sm_Prefix ='" & TxtPreFix.Text & "' AND GSLID = " & mGSLID & "")

RsTxn_CDC_GSLDetail.MoveFirst
RsTxn_CDC_GSLDetail.Find "CDCNo = " & TxtCDCID


RsMst_GSL!ClientIDRow = RsTxn_CDC_GSLDetail!ClientIDRowSeller
RsMst_GSL.Update
RsMst_GSL.Close


RsTxn_CDC_GSLDetail!Flag = "F" 'MSHFlexGrid1.TextMatrix(i, 18)
RsTxn_CDC_GSLDetail!ClientIDRow = RsTxn_CDC_GSLDetail!ClientIDRowSeller
RsTxn_CDC_GSLDetail!ClientIDRowSeller = 0
RsTxn_CDC_GSLDetail!PayoffDate = Null
RsTxn_CDC_GSLDetail!StorageChargesPayoffRate = Null
RsTxn_CDC_GSLDetail.Update
RsTxn_CDC_GSLDetail.Requery

If RsTxn_CDC_GSLDetail.RecordCount = 0 Then
   Call AddCmd_Click
Else
   RsTxn_CDC_GSLDetail.MoveLast
   Call Indata
End If
End Sub

Private Sub ExitCmd_Click()
 Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_CDC_GSLDetail.MoveFirst
Call Indata
'TmpRs.MoveFirst
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

mExchangeTypeID = GetExchangeID("SPOT")
Call TableTxn_CDC_GSLDetail("Where Flag = 'P' ")
'Call TableTxn_CDC("Where Flag = 'T' ")

If RsTxn_CDC_GSLDetail.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
Else
   'Call GButtonLock(Me, True)
   Call Indata
  
End If
End Sub



Private Sub LastCmd_Click()
RsTxn_CDC_GSLDetail.MoveLast
Call Indata

PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub SaveCmd_Click()
If TxtCDCID.Text = "" Then
  MsgBox "Blank CDC No Not Allowed"
  TxtCDCID.SetFocus
  Exit Sub
End If
If TxtStorageRate.Text = "" Then
  MsgBox "Blank Storage Rate Not Allowed!!!"
  TxtStorageRate.SetFocus
  Exit Sub
End If
If TxtNewClientID = "" Then
  MsgBox "Blank Buyer ClientId Not Allowed!!!"
  TxtNewClientID.SetFocus
  Exit Sub
End If
'If mClientRowID = mNewClientRowID Then
'  MsgBox "Seller And Buyer Clients Should Be Different!!!"
'  TxtNewClientId.SetFocus
'  Exit Sub
'End If
If ChkOwner() = False Then
  MsgBox "GSL Belongs To Another CDC"
  TxtCDCID.SetFocus
  Exit Sub
End If


Call SaveCDC_GSData
RsTxn_CDC_GSLDetail.Close
Call TableTxn_CDC_GSLDetail("Where Flag = 'P' ")

Call Indata
End Sub
Private Sub SaveCDC_GSData()

RsTxn_CDC_GSLDetail.MoveFirst
RsTxn_CDC_GSLDetail.Find "CDCNo = " & TxtCDCID

'mNewCDCNo = GetCDCNo

RsTxn_CDC_GSLDetail!Flag = "P" 'MSHFlexGrid1.TextMatrix(i, 18)
RsTxn_CDC_GSLDetail!ClientIDRowSeller = mClientRowID 'RsTxn_CDC_GSLDetail!ClientIDRow
RsTxn_CDC_GSLDetail!ClientIDRow = mNewClientRowID
RsTxn_CDC_GSLDetail!PayoffDate = StorageTo.Value
RsTxn_CDC_GSLDetail!StorageChargesPayoffRate = TxtStorageRate
RsTxn_CDC_GSLDetail.Update
'RsTxn_CDC_GSLDetail.Requery

Call TableMst_GSL("Where Sm_prefix = '" & CmbPreFix.Text & "' And  GSLID = " & mGSLID & "")
'RsMst_GSL!ISINID = mNewCDCNo 'RsTxn_CDC_GSLDetail!CDCNo
'RsMst_GSL!Flag = "O" 'Check1.Value
RsMst_GSL!ClientIDRow = mNewClientRowID

RsMst_GSL.Update
RsMst_GSL.Close

End Sub
Private Sub SaveCDC_AssayerData()

'Call TableTxn_CDC_AssayerDetail(" Where CDCID = " & TxtCDCID)
RsTxn_CDC_AssayerDetail.Find " sm_Prefix = '" & CmbPreFix.Text & "' AND CDCID = " & RsTxn_CDC!CDCID
If Not RsTxn_CDC_AssayerDetail.EOF Then
   'RsTxn_CDC_AssayerDetail.MoveFirst
Else
   RsTxn_CDC_AssayerDetail.AddNew
End If
RsTxn_CDC_AssayerDetail!G_UID = GetGUID
RsTxn_CDC_AssayerDetail!CDCID = RsTxn_CDC!CDCID
RsTxn_CDC_AssayerDetail!AssayerReportNo = TxtAssayerNo
RsTxn_CDC_AssayerDetail!AssayerName = TxtAssayerName
RsTxn_CDC_AssayerDetail!AssayerReportDate = ARDate.Value
RsTxn_CDC_AssayerDetail!Moisture = TxtMoisture
RsTxn_CDC_AssayerDetail!Damaged = TxtDamaged
RsTxn_CDC_AssayerDetail!Foreign = TxtForeign
RsTxn_CDC_AssayerDetail!Polarization = TxtPolarization
RsTxn_CDC_AssayerDetail!GrainSize = TxtGrainSize
RsTxn_CDC_AssayerDetail!Brokan = TxtBrokan
RsTxn_CDC_AssayerDetail!OilContain = TxtOil
RsTxn_CDC_AssayerDetail!colour = TxtColour
RsTxn_CDC_AssayerDetail!Ageing = TxtAgeing
RsTxn_CDC_AssayerDetail!StapleLength = TxtStapleLength
RsTxn_CDC_AssayerDetail!OrganicInorganic = TxtOrganic
RsTxn_CDC_AssayerDetail!WaxHoney = TxtWax
RsTxn_CDC_AssayerDetail!Grade = TxtGradeByAssayer
RsTxn_CDC_AssayerDetail!NED = NED.Value
RsTxn_CDC_AssayerDetail!FED = FED.Value
RsTxn_CDC_AssayerDetail!OverAllQuality = TxtOther  'OverAll Quality
RsTxn_CDC_AssayerDetail!AcceptenceFlag = ChkAcceptFlag
RsTxn_CDC_AssayerDetail!Remarks = TxtRemark
RsTxn_CDC_AssayerDetail!Flag = ""
RsTxn_CDC_AssayerDetail.Update

End Sub
Private Sub SaveCDC_InsuranceData()

Dim strSplitString() As String
RsTxn_CDC_InsuranceDetail.AddNew
RsTxn_CDC_InsuranceDetail!G_UID = GetGUID
RsTxn_CDC_InsuranceDetail!CDCID = RsTxn_CDC!CDCID
strSplitString = Split(TxtPolicyNo.Text, "~")
PolicyNo_ = Trim(strSplitString(0))
CNo = Trim(strSplitString(1))
RsTxn_CDC_InsuranceDetail!PolicyNo = PolicyNo_
RsTxn_CDC_InsuranceDetail!CoverNoteNo = CNo
RsTxn_CDC_InsuranceDetail!Flag = ""
RsTxn_CDC_InsuranceDetail.Update

End Sub
Private Function GetCDCID() As Double
Dim Retval As Double
tmp = "Select max(CDCID) from Txn_CDC where sm_Prefix ='" & CmbPreFix.Text & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCDCID = Retval
End Function
Private Function GetCDCNo() As Double
Dim Retval As Double
tmp = "Select max(CDCNo) from Txn_CDC_GSL_Detail where  sm_Prefix ='" & CmbPreFix.Text & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCDCNo = Retval
End Function

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
  Call TableTxn_CDC_GSLDetail("Where Flag = 'P' ")
  If RsTxn_CDC_GSLDetail.RecordCount = 0 Then Exit Sub
  
  'Call GButtonLock(Me, True)
  Call FirstCmd_Click
  EditCmd.Enabled = False
  Exit Sub
End If

End Sub

Private Sub TxtCDCID_LostFocus()
If TxtCDCID = "" Then Exit Sub
If CmbPreFix.Text = "" Then Exit Sub

If SaveCmd.Enabled = False Then Exit Sub

Call TableTxn_CDC_GSLDetail(" Where sm_Prefix ='" & CmbPreFix.Text & "'")
RsTxn_CDC_GSLDetail.Find "CDCNo = " & TxtCDCID
If Not RsTxn_CDC_GSLDetail.EOF Then
   If RsTxn_CDC!PrintFlag <> "P" Then
      MsgBox "CDC Printout is not Generated You can't select this CDC for payout !!!!! "
      Exit Sub
   End If
  mCDCID = RsTxn_CDC_GSLDetail!CDCID
  Call TableTxn_CDC("Where sm_Prefix ='" & CmbPreFix.Text & "' AND CDCID = " & mCDCID & "")
  Call TableTxn_CDC_AssayerDetail("Where  sm_Prefix ='" & CmbPreFix.Text & "' AND CDCID = " & mCDCID & "")
  Call TableTxn_CDC_InsuranceDetail("Where  sm_Prefix ='" & CmbPreFix.Text & "' AND CDCID = " & mCDCID & "")
  
  If RsTxn_CDC_GSLDetail!Flag <> "F" Then
     MsgBox "Payout already done for this CDC, if you want to chage it Delete the same and re-enter !!!!! "
     Exit Sub
  End If
  
  If RsTxn_CDC_AssayerDetail!AcceptenceFlag = 0 Then
     MsgBox "This CDC is Rejected By Assayer you can't perform payout on this lot !!!!! "
     Exit Sub
  End If
  Call PopulateGSLData
  mGSLID = RsTxn_CDC_GSLDetail!GSLID
Else
  MsgBox "No Such CDC No Exist!!!"
  TxtCDCID.SetFocus
  Exit Sub
End If
End Sub
Private Sub PopulateGSLData()


If (RsTxn_CDC_GSLDetail!ClientIDRowSeller) = 0 Then
   mClientRowID = RsTxn_CDC_GSLDetail!ClientIDRow
   TxtNewClientID = ""
   TxtNewClientName = ""
   TxtNewDPID = ""
   TxtNewDPName = ""
Else
   mClientRowID = RsTxn_CDC_GSLDetail!ClientIDRowSeller
   mNewClientRowID = RsTxn_CDC_GSLDetail!ClientIDRow
   ans = GetClientNameByRow(mNewClientRowID)
   TxtNewClientID = ans(0)
   TxtNewClientName = ans(1)
   TxtNewDPID = ans(2)
   TxtNewDPName = ans(3)
   
End If

mCMPID = RsTxn_CDC!CMPID
mCommodityID = RsTxn_CDC!CommodityID
mCDTFNo = RsTxn_CDC_GSLDetail!CdtfNo
mCDTFDate = RsTxn_CDC_GSLDetail!CDTFDate
CDCDate.Value = RsTxn_CDC!CDCDate
ans = GetClientNameByRow(mClientRowID)
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)
TxtCommodityCondition = IIf(IsNull(RsTxn_CDC!CommodityConditon), "", RsTxn_CDC!CommodityConditon)


TxtCMPID = GetCMPName(mCMPID)
TxtCMPAddress = GetCMPAddress(mCMPID, "A")
TxtLocationID = GetLocationFromCMP(mCMPID)
TxtCommodity = GetCommodityName(mCommodityID)
'tmp = " Where  Txn_CxTF_GSL.CxTFNo = " & RsTxn_CDC_GSLDetail!CDTFNo & "  And "
'tmp = tmp & "  Txn_CxTF_GSL.GSLID = " & RsTxn_CDC_GSLDetail!GSLID

'tmp = " Where  Txn_CxTF_GSL.CxTFNo = " & RsTxn_CDC_GSLDetail!CDTFNo & "  And "
tmp = " Where  sm_Prefix ='" & CmbPreFix.Text & "' AND Txn_CxTF_GSL.GSLID = " & RsTxn_CDC_GSLDetail!GSLID
Call TableTxn_CxTF_GSL(tmp)

TxtCDTFNo = RsTxn_CDC_GSLDetail!CdtfNo
TxtCDTFDate = RsTxn_CDC_GSLDetail!CDTFDate

ans = GetGSLByGLSID(RsTxn_CDC_GSLDetail!GSLID, CmbPreFix.Text)


TxtGodownAddress = ans(11)
TxtGodownID = ans(3)
TxtStackId = ans(4)
TxtLotId = ans(5)


For j = 1 To RsTxn_CxTF_GSL.RecordCount
    TxtNoofBags = Val(TxtNoofBags) + RsTxn_CxTF_GSL!NoofBags
    TxtQuantity = Val(TxtQuantity) + RsTxn_CxTF_GSL!AdjNCMSLWeight
    TxtQuantityKG = Val(TxtQuantityKG) + RsTxn_CxTF_GSL!AdjNCMSLWeight * 1000
    
    TxtGunnyWeight = Val(TxtGunnyWeight) + RsTxn_CxTF_GSL!GunnyWeight * 1000
    TxtStdDeduction = Val(TxtStdDeduction) + RsTxn_CxTF_GSL!StdDeduction * 1000 'TmpRs!StdDeduction
    TxtNetWeight = Val(TxtNetWeight) + RsTxn_CxTF_GSL!DematWeight * 1000 '  net weight in KGS
    RsTxn_CxTF_GSL.MoveNext
Next

TxtStorageRate = RsTxn_CDC_GSLDetail!StorageRate
StorageFrom.Value = RsTxn_CDC_GSLDetail!StorageTo + 1
StorageTo.Value = RsTxn_CDC_GSLDetail!StorageTo + Val(TxtNoofDays) ' 7
StorageTo.Enabled = True
LblStorageDays.Caption = (StorageTo.Value - StorageFrom.Value) + 1

'TxtNewClientId = ""

End Sub
Public Sub Indata()
Call ClearData
CmbPreFix.Visible = False
TxtPreFix.Visible = True


TxtPreFix = RsTxn_CDC_GSLDetail!SM_Prefix

Call TableTxn_CDC("Where CDCID = " & RsTxn_CDC_GSLDetail!CDCID & " And sm_prefix = '" & TxtPreFix & "'")
mCDCID = RsTxn_CDC_GSLDetail!CDCID

Call TableMst_Config
TxtNoofDays = IIf(IsNull(RsMst_Config!SportStorageChargesDays), 0, RsMst_Config!SportStorageChargesDays)

Call TableTxn_CDC_AssayerDetail("Where CDCID = " & mCDCID & " And sm_prefix = '" & TxtPreFix & "'")
Call TableTxn_CDC_InsuranceDetail("Where CDCID = " & mCDCID & " And sm_prefix = '" & TxtPreFix & "'")
Call IndataCDCReIssue_GSL
'Call InDataCDCReIssue_Assayer
'Call IndataCDCReIssue_Insurance
Call GButtonLock(Me, True)
EditCmd.Enabled = False
SearchCmd.Enabled = False
End Sub

Private Sub IndataCDCReIssue_GSL()
TxtCDCID = RsTxn_CDC_GSLDetail!CDCNo
TxtCDCID.Locked = True
CDCDate.Value = RsTxn_CDC_GSLDetail!CDCDate

mNewClientRowID = RsTxn_CDC_GSLDetail!ClientIDRow
mClientRowID = RsTxn_CDC_GSLDetail!ClientIDRowSeller

mCMPID = RsTxn_CDC!CMPID
mCommodityID = RsTxn_CDC!CommodityID
mCDTFNo = RsTxn_CDC_GSLDetail!CdtfNo
mCDTFDate = RsTxn_CDC_GSLDetail!CDTFDate

ans = GetClientNameByRow(mNewClientRowID)
TxtNewClientID = ans(0)
TxtNewClientName = ans(1)
TxtNewDPID = ans(2)
TxtNewDPName = ans(3)
TxtCMPID = GetCMPName(mCMPID)
TxtCMPAddress = GetCMPAddress(mCMPID, "A")
TxtLocationID = GetLocationFromCMP(mCMPID)
TxtCommodityCondition = IIf(IsNull(RsTxn_CDC!CommodityConditon), "", RsTxn_CDC!CommodityConditon)
TxtCommodity = GetCommodityName(mCommodityID)


'tmp = " Where  Txn_CxTF_GSL.CxTFNo = " & RsTxn_CDC_GSLDetail!CDTFNo & "  And "
'tmp = tmp & "  Txn_CxTF_GSL.GSLID = " & RsTxn_CDC_GSLDetail!GSLID
tmp = "  Where Txn_CxTF_GSL.GSLID = " & RsTxn_CDC_GSLDetail!GSLID & " And sm_prefix = '" & TxtPreFix & "'"

Call TableTxn_CxTF_GSL(tmp)


ans = GetClientNameByRow(mClientRowID)
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)

TxtCDTFNo = RsTxn_CDC_GSLDetail!CdtfNo
TxtCDTFDate = RsTxn_CDC_GSLDetail!CDTFDate

ans = GetGSLByGLSID(RsTxn_CDC_GSLDetail!GSLID, TxtPreFix)
mGSLID = RsTxn_CDC_GSLDetail!GSLID

TxtGodownAddress = ans(11)
TxtGodownID = ans(3)
TxtStackId = ans(4)
TxtLotId = ans(5)


For j = 1 To RsTxn_CxTF_GSL.RecordCount
    TxtNoofBags = Val(TxtNoofBags) + RsTxn_CxTF_GSL!NoofBags
    TxtQuantity = Val(TxtQuantity) + RsTxn_CxTF_GSL!AdjNCMSLWeight
    TxtQuantityKG = Val(TxtQuantityKG) + RsTxn_CxTF_GSL!AdjNCMSLWeight * 1000
    
    TxtGunnyWeight = Val(TxtGunnyWeight) + RsTxn_CxTF_GSL!GunnyWeight * 1000
    TxtStdDeduction = Val(TxtStdDeduction) + IIf(IsNull(RsTxn_CxTF_GSL!StdDeduction), 0, RsTxn_CxTF_GSL!StdDeduction) * 1000 'TmpRs!StdDeduction
    TxtNetWeight = Val(TxtNetWeight) + RsTxn_CxTF_GSL!DematWeight * 1000 '  net weight in KGS
    RsTxn_CxTF_GSL.MoveNext
Next

If Not IsNull(RsTxn_CDC_GSLDetail!StorageChargesPayoffRate) Then
   TxtStorageRate = RsTxn_CDC_GSLDetail!StorageChargesPayoffRate
End If
TxtStorageRate.Locked = True
StorageFrom.Value = RsTxn_CDC_GSLDetail!StorageTo + 1
If Not IsNull(RsTxn_CDC_GSLDetail!PayoffDate) Then
   StorageTo.Value = RsTxn_CDC_GSLDetail!PayoffDate
End If
StorageTo.Enabled = False
LblStorageDays.Caption = (StorageTo.Value - StorageFrom.Value) + 1
TxtStorageAmount = Val(TxtQuantity) * Val(TxtStorageRate) * Val(LblStorageDays.Caption)
TxtNewClientID.Locked = True


End Sub
Public Sub ClearData()
TxtCDCID.Text = ""
CDCDate.Value = Date
TxtCMPID = ""
TxtCMPAddress = ""
TxtCommodity = ""
TxtLocationID = ""
TxtClientID = ""
TxtClientName = ""
TxtDPID = ""
TxtDPName = ""
TxtCommodityCondition = ""
TxtGodownAddress = ""
TxtGodownID = ""
TxtStackId = ""
TxtLotId = ""
TxtNoofBags = ""
TxtQuantity = ""
TxtGunnyWeight = ""
TxtStdDeduction = ""
TxtNetWeight = ""
TxtStorageAmount = ""
TxtNewClientID = ""
TxtNewClientName = ""
TxtNewDPID = ""
TxtNewDPName = ""
TxtQuantityKG = ""
TxtStorageRate = ""
LblStorageDays.Caption = ""
mClientIDRow = 0
mCMPID = 0
mCommodityID = 0
mGSLID = 0

End Sub
Private Sub TxtNewClientID_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
CmbClientDP.Visible = False
If TxtNewClientID = "" Then Exit Sub
'If CmbExchangeTypeID.Text = "" Then Exit Sub

Dim ans As Variant
mExchangeTypeID = GetExchangeID("SPOT")
ans = GetClientNameDtl(TxtNewClientID, mExchangeTypeID)

If ans(0, 0) = False Then
   MsgBox "Not Found !!!"
   TxtNewClientID.SetFocus
   Exit Sub
End If
If ans(3, 0) > 1 Then
   CmbClientDP.Visible = True
End If
CmbClientDP.Text = ""
CmbClientDP.Clear
tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      "Mst_Client.ClientName,  Mst_Client.ClientIDRow " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtNewClientID & "'));"
      
      '"WHERE (((Mst_Depository.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtNewClientId & "'));"
Call Tmp1Table

For op = 1 To TmpRs1.RecordCount
    If CmbClientDP.Text = "" Then
       CmbClientDP.Text = TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    End If
    CmbClientDP.AddItem TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    TmpRs1.MoveNext
Next
If ans(3, 0) <= 1 Then
   CmbClientDP_LostFocus
Else
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   End If
End If
'Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)

 
End Sub
Private Sub CmbClientDP_LostFocus()

Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If
TmpRs1.MoveFirst


TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtNewClientName = TmpRs1!ClientName
TxtNewDPID = TmpRs1!DepositoryID
TxtNewDPName = TmpRs1!DepositoryName
CmbClientDP.Visible = False
mNewClientRowID = TmpRs1!ClientIDRow
'CmdClientAdd2List.SetFocus
End Sub
Private Sub StorageTo_LostFocus()
'If StorageTo.Value > Date Then
'    MsgBox "Future Date not Allowed !!!!! "
'    StorageTo.SetFocus
'    Exit Sub
'End If

If StorageFrom.Value > StorageTo.Value Then
    MsgBox "From Date must be Less Than To Date Always"
    StorageTo.SetFocus
    Exit Sub
Else
    mStorageDays = StorageTo.Value - StorageFrom + 1
    LblStorageDays = mStorageDays '- Val(TxtNoofDays) '7 'mStorageDays - 6
    TxtStorageAmount = Val(TxtQuantity) * Val(TxtStorageRate) * Val(LblStorageDays.Caption)
End If
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CDC_GSLDetail.MoveNext
If RsTxn_CDC_GSLDetail.EOF Then
   RsTxn_CDC_GSLDetail.MoveLast
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
RsTxn_CDC_GSLDetail.MovePrevious

If RsTxn_CDC_GSLDetail.BOF Then
   RsTxn_CDC_GSLDetail.MoveFirst
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
Function ChkOwner() As Boolean
Dim Retval As Boolean
Retval = True
Call TableMst_GSL("Where ISINID = '" & TxtCDCID & "'")
If RsMst_GSL.RecordCount = 0 Then
  Retval = False
End If
ChkOwner = Retval
End Function
Private Sub TxtStorageRate_LostFocus()
If TxtStorageRate.Text = "" Then Exit Sub
If IsNumeric(TxtStorageRate) = False Then
   MsgBox "Invalid numeric formate"
   TxtStorageRate.SetFocus
   Exit Sub
End If
TxtStorageAmount = Val(TxtQuantity) * Val(TxtStorageRate) * Val(LblStorageDays.Caption)
End Sub
