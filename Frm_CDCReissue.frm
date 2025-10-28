VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_CDCReissue 
   Caption         =   "CDC ReIssue"
   ClientHeight    =   8835
   ClientLeft      =   -3000
   ClientTop       =   555
   ClientWidth     =   14715
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   7860
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   14460
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
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1830
      End
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
         TabIndex        =   141
         Top             =   1980
         Width           =   1830
      End
      Begin VB.TextBox TxtNewCDCNo 
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
         TabIndex        =   137
         Top             =   720
         Width           =   1815
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
         Left            =   13560
         Locked          =   -1  'True
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   495
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
         Height          =   360
         Left            =   13560
         TabIndex        =   133
         Top             =   240
         Visible         =   0   'False
         Width           =   495
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
         Height          =   360
         Left            =   11880
         TabIndex        =   131
         Top             =   240
         Visible         =   0   'False
         Width           =   495
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
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   2835
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
         TabIndex        =   121
         Top             =   2400
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
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   2400
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
         TabIndex        =   2
         Top             =   240
         Width           =   1815
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
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   1560
         Width           =   7335
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
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1140
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
         TabIndex        =   85
         Top             =   1980
         Width           =   7320
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
         TabIndex        =   84
         Top             =   1125
         Width           =   7335
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
         TabIndex        =   83
         Top             =   1560
         Width           =   1830
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4320
         Left            =   105
         TabIndex        =   20
         Top             =   3420
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   7620
         _Version        =   393216
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Godown Stack Detail"
         TabPicture(0)   =   "Frm_CDCReissue.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label11"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label23"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "LblStorageDays"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label27"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label13"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label30"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label31"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label44"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label45"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lblNameDid"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label46"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label48"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label49"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label50"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label52"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label53"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label54"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label55"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label56"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Label59"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "CmbClientDP"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "StorageTo"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "StorageFrom"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "TxtGodownID"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "TxtStackId"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "TxtLotId"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "TxtNoOfBags"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "TxtQuantity"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "TxtNewDPID"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "TxtNewClientName"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "TxtNewClientId"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "TxtNewDPName"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "TxtStorageRate"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "TxtNetWeight"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "TxtStdDeduction"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "TxtGunnyWeight"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "TxtGodownAddress"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "TxtStorageAmount"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "TxtQuantityKG"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).ControlCount=   39
         TabCaption(1)   =   "Assayer Details"
         TabPicture(1)   =   "Frm_CDCReissue.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Insurance Detail"
         TabPicture(2)   =   "Frm_CDCReissue.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "History Detail"
         TabPicture(3)   =   "Frm_CDCReissue.frx":0054
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "Label63"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "Label64"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "TxtCreated"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).Control(3)=   "TxtUpdated"
         Tab(3).Control(3).Enabled=   0   'False
         Tab(3).ControlCount=   4
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
            Left            =   1365
            Locked          =   -1  'True
            TabIndex        =   145
            TabStop         =   0   'False
            Top             =   1485
            Width           =   9315
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
            Left            =   1365
            Locked          =   -1  'True
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   720
            Width           =   9330
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
            Left            =   -68280
            Locked          =   -1  'True
            TabIndex        =   130
            TabStop         =   0   'False
            Top             =   1230
            Width           =   1305
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
            Left            =   -64920
            Locked          =   -1  'True
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   2535
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
            Left            =   -73200
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   480
            Width           =   12255
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
            Left            =   -69240
            Locked          =   -1  'True
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   1665
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
            Left            =   -69240
            Locked          =   -1  'True
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   2100
            Width           =   2265
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
            Left            =   -69240
            Locked          =   -1  'True
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   2535
            Width           =   2265
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
            Left            =   -62280
            TabIndex        =   4
            Top             =   1230
            Width           =   1275
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
            Left            =   -69240
            Locked          =   -1  'True
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   3390
            Width           =   8250
         End
         Begin VB.TextBox TxtNewClientId 
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
            Left            =   -73200
            TabIndex        =   6
            Top             =   2955
            Width           =   2070
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
            Left            =   -69240
            Locked          =   -1  'True
            TabIndex        =   108
            Top             =   2955
            Width           =   8250
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
            Left            =   -73200
            Locked          =   -1  'True
            TabIndex        =   107
            Top             =   3390
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
            Left            =   -69240
            Locked          =   -1  'True
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   1230
            Width           =   825
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
            Left            =   -73200
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   2535
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
            Left            =   -73200
            Locked          =   -1  'True
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   2100
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
            Left            =   -73200
            Locked          =   -1  'True
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   1665
            Width           =   2070
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
            Left            =   -73200
            Locked          =   -1  'True
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   1230
            Width           =   2070
         End
         Begin VB.Frame Frame2 
            Height          =   3945
            Left            =   -74880
            TabIndex        =   34
            Top             =   360
            Width           =   13935
            Begin VB.TextBox TxtGrainSize 
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
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   2700
               Width           =   1575
            End
            Begin VB.TextBox TxtPolarization 
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
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   2280
               Width           =   1575
            End
            Begin VB.CheckBox ChkAcceptFlag 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   4800
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   3120
               Width           =   255
            End
            Begin VB.TextBox TxtRemark 
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
               Height          =   2940
               Left            =   10440
               Locked          =   -1  'True
               MaxLength       =   5000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   600
               Width           =   3375
            End
            Begin VB.TextBox TxtOther 
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
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   3120
               Width           =   1575
            End
            Begin VB.TextBox TxtOrganic 
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
               Left            =   8280
               Locked          =   -1  'True
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   1440
               Width           =   2055
            End
            Begin VB.TextBox TxtColour 
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
               Left            =   4800
               Locked          =   -1  'True
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   1860
               Width           =   1575
            End
            Begin VB.TextBox TxtStapleLength 
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
               Left            =   8280
               Locked          =   -1  'True
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   1020
               Width           =   2055
            End
            Begin VB.TextBox TxtWax 
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
               Left            =   8280
               Locked          =   -1  'True
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   1860
               Width           =   2055
            End
            Begin VB.TextBox TxtAgeing 
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
               Left            =   4800
               Locked          =   -1  'True
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   2280
               Width           =   1575
            End
            Begin VB.TextBox TxtDamaged 
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
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   1440
               Width           =   1575
            End
            Begin VB.TextBox TxtOil 
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
               Left            =   4800
               Locked          =   -1  'True
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   1440
               Width           =   1575
            End
            Begin VB.TextBox TxtForeign 
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
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   1845
               Width           =   1575
            End
            Begin VB.TextBox TxtMoisture 
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
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   1020
               Width           =   1575
            End
            Begin VB.TextBox TxtBrokan 
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
               Left            =   4800
               Locked          =   -1  'True
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   1020
               Width           =   1575
            End
            Begin VB.TextBox TxtAssayerNo 
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
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   195
               Width           =   4455
            End
            Begin VB.TextBox TxtGradeByAssayer 
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
               Left            =   8280
               Locked          =   -1  'True
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   2280
               Width           =   2055
            End
            Begin VB.TextBox TxtAssayerName 
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
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   600
               Width           =   8415
            End
            Begin MSComCtl2.DTPicker ARDate 
               Height          =   375
               Left            =   8265
               TabIndex        =   53
               Top             =   180
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   61407233
               CurrentDate     =   39884
            End
            Begin MSComCtl2.DTPicker NED 
               Height          =   375
               Left            =   4800
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   2685
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   61407233
               CurrentDate     =   39884
            End
            Begin MSComCtl2.DTPicker FED 
               Height          =   375
               Left            =   8280
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   2685
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   61407233
               CurrentDate     =   39884
            End
            Begin VB.Label Label28 
               Caption         =   "Grain Size(mm)"
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
               TabIndex        =   77
               Top             =   2700
               Width           =   1695
            End
            Begin VB.Label Label25 
               Caption         =   "Polarization"
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
               TabIndex        =   76
               Top             =   2340
               Width           =   1335
            End
            Begin VB.Label Label5 
               Caption         =   "Accepted"
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
               Left            =   3720
               TabIndex        =   75
               Top             =   3120
               Width           =   1095
            End
            Begin VB.Label Label47 
               Caption         =   "Remark"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   15
               Left            =   3480
               TabIndex        =   74
               Top             =   2220
               Width           =   15
            End
            Begin VB.Label Label37 
               Caption         =   "Over All Quality"
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
               TabIndex        =   73
               Top             =   3120
               Width           =   1695
            End
            Begin VB.Label Label36 
               Caption         =   "Colour"
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
               Left            =   3720
               TabIndex        =   72
               Top             =   1860
               Width           =   735
            End
            Begin VB.Label Label35 
               Caption         =   "Organic - InOrganic"
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
               Left            =   6480
               TabIndex        =   71
               Top             =   1440
               Width           =   1935
            End
            Begin VB.Label Label34 
               Caption         =   "Staple Length"
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
               Left            =   6480
               TabIndex        =   70
               Top             =   1020
               Width           =   1455
            End
            Begin VB.Label Label33 
               Caption         =   "Wax-Honey"
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
               Left            =   6480
               TabIndex        =   69
               Top             =   1860
               Width           =   1695
            End
            Begin VB.Label Label32 
               Caption         =   "Ageing"
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
               Left            =   3720
               TabIndex        =   68
               Top             =   2280
               Width           =   735
            End
            Begin VB.Label Label24 
               Caption         =   "Damaged"
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
               TabIndex        =   67
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label22 
               Caption         =   "Oil Contain"
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
               Left            =   3720
               TabIndex        =   66
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label21 
               Caption         =   "Foreign"
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
               TabIndex        =   65
               Top             =   1845
               Width           =   1455
            End
            Begin VB.Label Label20 
               Caption         =   "Brokan"
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
               Left            =   3720
               TabIndex        =   64
               Top             =   1020
               Width           =   855
            End
            Begin VB.Label Label17 
               Caption         =   "Moisture"
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
               TabIndex        =   63
               Top             =   1020
               Width           =   975
            End
            Begin VB.Label Label15 
               Caption         =   "AR No"
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
               TabIndex        =   62
               Top             =   195
               Width           =   1335
            End
            Begin VB.Label Label16 
               Caption         =   "Assayer Name"
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
               TabIndex        =   61
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label29 
               Caption         =   "AR Date"
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
               Left            =   6720
               TabIndex        =   60
               Top             =   195
               Width           =   975
            End
            Begin VB.Label Label1 
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
               Height          =   255
               Left            =   11520
               TabIndex        =   59
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label14 
               Caption         =   "Grade"
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
               Left            =   6480
               TabIndex        =   58
               Top             =   2280
               Width           =   855
            End
            Begin VB.Label Label18 
               Caption         =   "NED"
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
               Left            =   3720
               TabIndex        =   57
               Top             =   2700
               Width           =   615
            End
            Begin VB.Label Label19 
               Caption         =   "FED"
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
               Left            =   6495
               TabIndex        =   56
               Top             =   2700
               Width           =   615
            End
         End
         Begin VB.Frame Frame3 
            Height          =   3930
            Left            =   -74880
            TabIndex        =   21
            Top             =   360
            Width           =   13815
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
               Height          =   375
               Left            =   7080
               Locked          =   -1  'True
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   1155
               Width           =   1815
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
               Height          =   2190
               Left            =   2145
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   1605
               Width           =   11415
            End
            Begin VB.TextBox TxtInsuranceCo 
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
               Left            =   2160
               TabIndex        =   23
               Top             =   240
               Width           =   6720
            End
            Begin VB.TextBox TxtPolicyNo 
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
               Left            =   2160
               TabIndex        =   22
               Top             =   675
               Width           =   6720
            End
            Begin MSComCtl2.DTPicker ValidFrom 
               Height          =   375
               Left            =   2160
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   1155
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   61407233
               CurrentDate     =   39884
            End
            Begin MSComCtl2.DTPicker ValidTo 
               Height          =   375
               Left            =   4680
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   1155
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   61407233
               CurrentDate     =   39884
            End
            Begin VB.Label Label43 
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
               Height          =   375
               Left            =   120
               TabIndex        =   33
               Top             =   1635
               Width           =   1095
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
               Left            =   6120
               TabIndex        =   32
               Top             =   1155
               Width           =   975
            End
            Begin VB.Label Label41 
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
               Left            =   105
               TabIndex        =   31
               Top             =   1185
               Width           =   1695
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
               Left            =   3720
               TabIndex        =   30
               Top             =   1155
               Width           =   975
            End
            Begin VB.Label Label39 
               Caption         =   "Policy No /  Cover note."
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
               TabIndex        =   29
               Top             =   645
               Width           =   1815
            End
            Begin VB.Label Label38 
               Caption         =   "Insurance Co."
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
               TabIndex        =   28
               Top             =   285
               Width           =   1455
            End
         End
         Begin MSComCtl2.DTPicker StorageFrom 
            Height          =   375
            Left            =   -64920
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   1665
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   61407233
            CurrentDate     =   39884
         End
         Begin MSComCtl2.DTPicker StorageTo 
            Height          =   375
            Left            =   -62280
            TabIndex        =   5
            Top             =   1665
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   61407233
            CurrentDate     =   39884
         End
         Begin VB.ComboBox CmbClientDP 
            Height          =   315
            Left            =   -69240
            Sorted          =   -1  'True
            TabIndex        =   7
            Top             =   2955
            Visible         =   0   'False
            Width           =   8280
         End
         Begin VB.Label Label64 
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
            TabIndex        =   147
            Top             =   1545
            Width           =   975
         End
         Begin VB.Label Label63 
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
            TabIndex        =   146
            Top             =   780
            Width           =   1695
         End
         Begin VB.Label Label59 
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
            Left            =   -64920
            TabIndex        =   135
            Top             =   1320
            Width           =   3255
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
            Left            =   -66735
            TabIndex        =   129
            Top             =   2625
            Width           =   1575
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
            Left            =   -74880
            TabIndex        =   126
            Top             =   540
            Width           =   1575
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
            Left            =   -70920
            TabIndex        =   119
            Top             =   1725
            Width           =   1455
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
            Left            =   -70920
            TabIndex        =   118
            Top             =   2160
            Width           =   1455
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
            Left            =   -70920
            TabIndex        =   117
            Top             =   2595
            Width           =   1215
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
            Left            =   -74880
            TabIndex        =   113
            Top             =   3015
            Width           =   1605
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
            Left            =   -70920
            TabIndex        =   112
            Top             =   3030
            Width           =   1695
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
            Left            =   -74895
            TabIndex        =   111
            Top             =   3390
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
            Left            =   -70920
            TabIndex        =   110
            Top             =   3360
            Width           =   1575
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
            Left            =   -70920
            TabIndex        =   106
            Top             =   3600
            Visible         =   0   'False
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
            Left            =   -70920
            TabIndex        =   105
            Top             =   1290
            Width           =   1455
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
            Left            =   -74880
            TabIndex        =   103
            Top             =   2595
            Width           =   1215
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
            Left            =   -74880
            TabIndex        =   101
            Top             =   2160
            Width           =   750
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
            Left            =   -74880
            TabIndex        =   99
            Top             =   1725
            Width           =   1455
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
            Left            =   -74880
            TabIndex        =   97
            Top             =   1290
            Width           =   1455
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
            Left            =   -66735
            TabIndex        =   82
            Top             =   2130
            Width           =   1710
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
            Left            =   -64920
            TabIndex        =   81
            Top             =   2160
            Width           =   75
         End
         Begin VB.Label Label23 
            Caption         =   "Storage To"
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
            Left            =   -63480
            TabIndex        =   80
            Top             =   1680
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
            Left            =   -66720
            TabIndex        =   79
            Top             =   1725
            Width           =   1455
         End
      End
      Begin MSComCtl2.DTPicker CDCDate 
         Height          =   375
         Left            =   9240
         TabIndex        =   88
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   61407233
         CurrentDate     =   39884
      End
      Begin MSComCtl2.DTPicker CDCDateNew 
         Height          =   375
         Left            =   6120
         TabIndex        =   3
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61407233
         CurrentDate     =   39884
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
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label62 
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
         TabIndex        =   142
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label61 
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
         Height          =   510
         Left            =   120
         TabIndex        =   140
         Top             =   1905
         Width           =   1710
      End
      Begin VB.Label Label60 
         Caption         =   "New CDC Date"
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
         Left            =   3915
         TabIndex        =   139
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "New CDC NO"
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
         TabIndex        =   138
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label58 
         Caption         =   "CDTF Date"
         Height          =   255
         Left            =   12360
         TabIndex        =   134
         Top             =   300
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label57 
         Caption         =   "CDTF NO"
         Height          =   255
         Left            =   10680
         TabIndex        =   132
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
         TabIndex        =   125
         Top             =   2400
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
         TabIndex        =   124
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "WH Name (CMP)"
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   2445
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
         TabIndex        =   95
         Top             =   1560
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
         TabIndex        =   94
         Top             =   1560
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
         TabIndex        =   93
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
         TabIndex        =   92
         Top             =   1185
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
         Left            =   8160
         TabIndex        =   91
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
         TabIndex        =   90
         Top             =   1980
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
         TabIndex        =   89
         Top             =   1185
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   14430
      TabIndex        =   0
      Top             =   7920
      Width           =   14490
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         Height          =   750
         Left            =   11835
         TabIndex        =   18
         Top             =   120
         Width           =   2175
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   9540
         TabIndex        =   9
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   7245
         TabIndex        =   17
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   4950
         TabIndex        =   16
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   2655
         TabIndex        =   15
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   495
         Width           =   2295
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   9540
         TabIndex        =   13
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   7260
         TabIndex        =   12
         Top             =   120
         Width           =   2280
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   4965
         TabIndex        =   11
         Top             =   120
         Width           =   2280
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   2655
         TabIndex        =   8
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   120
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Frm_CDCReissue"
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
Dim Edit_Flg As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call TableMst_Config
TxtNoofDays = IIf(IsNull(RsMst_Config!SportStorageChargesDays), 0, RsMst_Config!SportStorageChargesDays)
RsMst_Config.Close
TxtNewCDCNo = ""
'Call TableTxn_CDC("Where Flag = 'O'")
Call ClearData
Call GButtonLock(Me, False)
CmdPrint.Enabled = False
TxtCDCID.Locked = False
TxtStorageRate.Locked = False
TxtNewClientID.Locked = False
CmbPreFix.Visible = True
TxtPreFix.Visible = False
CmbPreFix.Text = ""
TxtQuantityKG = ""

TxtCreated = ""
TxtUpdated = ""

'If RsTxn_CDC.RecordCount <> 0 Then
' TxtCDCID.SetFocus
'End If

End Sub

Private Sub CDCDateNew_LostFocus()

If CDCDateNew.Value > Date Then
   MsgBox "Future date is not allowed!!! "
   CDCDateNew.SetFocus
End If

If CDCDateNew.Value < CDCDate.Value Then
   MsgBox "New CDC date must be greater than old CDC date!!! "
   CDCDateNew.SetFocus
End If

End Sub

Private Sub CmbPreFix_GotFocus()
tmp = CmbPreFix.Text
Call TableMst_StoreManLocations("")

CmbPreFix.Clear
If RsMst_StoreManLocations.RecordCount = 0 Then
   MsgBox "No Bank found!!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_StoreManLocations.RecordCount
    CmbPreFix.AddItem RsMst_StoreManLocations!StoreManLocations
    RsMst_StoreManLocations.MoveNext
Next
CmbPreFix.Text = tmp
End Sub

Private Sub CmbPreFix_LostFocus()
If CmbPreFix.Text = "" Then
   Exit Sub
End If
RsMst_StoreManLocations.MoveFirst
RsMst_StoreManLocations.Find "StoreManLocations  = '" & CmbPreFix.Text & "'"

If RsMst_StoreManLocations.EOF Then
   MsgBox "Invalid Location ID selection!!!"
   CmbPreFix.SetFocus
   Exit Sub
End If
mSM_Prefix = RsMst_StoreManLocations!SM_Prefix

End Sub

Private Sub CmdPrint_Click()
RsTxn_CDC!PrintFlag = "P"
RsTxn_CDC.Update

Gen_mTimeStamp = GetTimeStamp
Call Create_Xls(Gen_mTimeStamp & Gen_UserName & "_CDC-" & Trim(TxtPreFix) & "" & Trim(TxtNewCDCNo) & ".xls")

Call Xls_Print_Rpt
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
'Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_UserName & "_CDC-" & Trim(TxtPreFix) & "" & Trim(TxtNewCDCNo) & ".xls")
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & "" & Gen_UserName & "_CDC-" & Trim(TxtPreFix) & "" & Trim(TxtNewCDCNo) & ".xls")
 
 
 
'i = 2
Set oWS = oWB.Worksheets("Sheet" & 1)
'oWS.PageSetup.RightMargin = 0.05
oWS.Name = TxtPreFix & "-" & TxtNewCDCNo   'MSHFlexGrid2.TextMatrix(i, 23)
oWS.ClearArrows
mRow = 1

oWS.Range("a1:d1").Merge
oWS.Range("a:a").ColumnWidth = 21.86
oWS.Range("b:b").ColumnWidth = 21.86
oWS.Range("c:c").ColumnWidth = 21.86
oWS.Range("d:d").ColumnWidth = 21.86

'oWS.Range("A1:E1").Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, c + 1) = "NATIONAL COLLATERAL MANAGEMENT SERVICES LIMITED"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 11

'oWS.Cells(1, 1).Font.ColorIndex = 4
mRow = mRow + 1
'oWS.Range("a3:e3").Merge
tmp = "a" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Commodity Deposit Certificate (CDC)"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).RowHeight = 24

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Warehouse License Number"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Size = 9

oWS.Cells(mRow, c + 2) = GetLicenseFromGSL(RsTxn_CDC_GSLDetail!GSLID)
oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
oWS.Cells(mRow, c + 2).Font.Size = 9

oWS.Cells(mRow, c + 3) = "CDC No"
oWS.Cells(mRow, c + 3).Font.Size = 9
oWS.Cells(mRow, c + 3).Font.Bold = True
oWS.Cells(mRow, c + 4) = TxtPreFix & "-" & TxtNewCDCNo
oWS.Cells(mRow, c + 4).HorizontalAlignment = 2
oWS.Cells(mRow, c + 4).Font.Size = 9

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Location & Address of Warehouse"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).Font.Bold = True
tmp = "b" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = TxtLocationID & "--" & TxtCMPAddress
oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
oWS.Cells(mRow, c + 2).Font.Size = 9
oWS.Cells(mRow, c + 2).WrapText = True

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "CMP ID & Address"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 1).Font.Bold = True

tmp = "b" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge

oWS.Cells(mRow, c + 2) = TxtCMPID
oWS.Cells(mRow, c + 2).HorizontalAlignment = 2
oWS.Cells(mRow, c + 2).Font.Size = 9
oWS.Cells(mRow, c + 2).WrapText = True

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Lot No"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).WrapText = True

tmp = "b" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = TxtLotId
oWS.Cells(mRow, c + 2).Font.Size = 9

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Name of Depositor"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 9

tmp = "b" & mRow & ":" & "d" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = TxtNewClientID & "-" & TxtNewClientName
oWS.Cells(mRow, c + 2).Font.Size = 9
oWS.Cells(mRow, c + 2).WrapText = True

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Commodity"
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 2) = "No of Units"
oWS.Cells(mRow, c + 2).Font.Size = 9
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 3) = "Net Quantity(Kgs)"
oWS.Cells(mRow, c + 3).Font.Size = 9
oWS.Cells(mRow, c + 3).Font.Bold = True
oWS.Cells(mRow, c + 3).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 4) = "Condition of Commodity"
oWS.Cells(mRow, c + 4).Font.Size = 9
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).HorizontalAlignment = -4108

mRow = mRow + 1
oWS.Cells(mRow, c + 1) = TxtCommodity
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 2) = TxtNoOfBags
oWS.Cells(mRow, c + 2).Font.Size = 9
oWS.Cells(mRow, c + 3) = Format(TxtNetWeight, "#######0.000") 'TxtNetWeight
oWS.Cells(mRow, c + 3).Font.Size = 9
oWS.Cells(mRow, c + 4) = TxtCommodityCondition
oWS.Cells(mRow, c + 4).Font.Size = 9

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Name & License No of Assayer"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 2) = "Grade Parameter Values as per Assayer's Certificate"
oWS.Cells(mRow, c + 2).Font.Size = 9
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 2).WrapText = True
oWS.Cells(mRow, c + 3) = "Grade as per Assayer's Certificate"
oWS.Cells(mRow, c + 3).Font.Size = 9
oWS.Cells(mRow, c + 3).Font.Bold = True
oWS.Cells(mRow, c + 3).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 3).WrapText = True
oWS.Cells(mRow, c + 4) = "Grade Valid till"
oWS.Cells(mRow, c + 4).Font.Size = 9
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 4).WrapText = True

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = TxtAssayerNo & "-" & TxtAssayerName
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).WrapText = True
oWS.Cells(mRow, c + 2) = ""
oWS.Cells(mRow, c + 2).Font.Size = 9
oWS.Cells(mRow, c + 2).WrapText = True
oWS.Cells(mRow, c + 3) = TxtGradeByAssayer
oWS.Cells(mRow, c + 3).Font.Size = 9
oWS.Cells(mRow, c + 3).WrapText = True
oWS.Cells(mRow, c + 4) = Format(FED.Value, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 4).Font.Size = 9
oWS.Cells(mRow, c + 4).WrapText = True
mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Insurance Details"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).WrapText = True

tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = "Insurance Company"
oWS.Cells(mRow, c + 2).Font.Size = 9
oWS.Cells(mRow, c + 2).Font.Bold = True
oWS.Cells(mRow, c + 2).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 2).WrapText = True

oWS.Cells(mRow, c + 4) = "Amount Insured"
oWS.Cells(mRow, c + 4).Font.Size = 9
oWS.Cells(mRow, c + 4).Font.Bold = True
oWS.Cells(mRow, c + 4).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 4).WrapText = True

mRow = mRow + 1
    
oWS.Cells(mRow, c + 1) = TxtRisks
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).WrapText = True
tmp = "b" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 2) = TxtInsuranceCo
oWS.Cells(mRow, c + 2).Font.Size = 9
oWS.Cells(mRow, c + 2).WrapText = True
oWS.Cells(mRow, c + 4) = Format(TxtAmount, "#######0.000") 'TxtAmount
oWS.Cells(mRow, c + 4).Font.Size = 9
oWS.Cells(mRow, c + 4).WrapText = True

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Goods Are Accepted for Storage From"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).WrapText = True

oWS.Cells(mRow, c + 2) = Format(StorageFrom.Value, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 2).Font.Size = 9

oWS.Cells(mRow, c + 3) = "Storage To"
oWS.Cells(mRow, c + 3).Font.Size = 9
oWS.Cells(mRow, c + 3).Font.Bold = True
oWS.Cells(mRow, c + 3).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 3).WrapText = True

oWS.Cells(mRow, c + 4) = Format(StorageTo.Value, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 4).Font.Size = 9


mRow = mRow + 1

tmp = "a" & mRow & ":" & "c" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 4) = "Rs." & " " & Format(TxtStorageAmount, "#######0.00") 'TxtStorageAmount
oWS.Cells(mRow, c + 4).Font.Size = 9
oWS.Cells(mRow, c + 4).Font.Bold = True

mRow = mRow + 1
tmp = "a" & mRow & ":" & "b" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 3) = "Hologram No:-"
oWS.Cells(mRow, c + 3).Font.Size = 9
oWS.Cells(mRow, c + 3).Font.Bold = True

oWS.Cells(mRow, c + 4) = ""
oWS.Cells(mRow, c + 4).Font.Size = 9
oWS.Cells(mRow, c + 4).RowHeight = 35

mRow = mRow + 1

oWS.Cells(mRow, c + 1) = "Date"
oWS.Cells(mRow, c + 1).Font.Size = 9
oWS.Cells(mRow, c + 1).Font.Bold = True


oWS.Cells(mRow, c + 2) = Format(CDCDateNew.Value, "dd-mmm-yyyy")
oWS.Cells(mRow, c + 2).Font.Size = 9

oWS.Cells(mRow, c + 3) = "Signature & Stamp of Warehouse Manager"
oWS.Cells(mRow, c + 3).Font.Size = 9
oWS.Cells(mRow, c + 3).Font.Bold = True

oWS.Cells(mRow, c + 4) = ""
oWS.Cells(mRow, c + 4).Font.Size = 9
oWS.Cells(mRow, c + 4).RowHeight = 24
mRow = mRow + 1

k = mRow

For Cell = 1 To k
    tmp = "a" & Cell & ":" & "d" & Cell
    oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
    oWS.Range(tmp).Font.Name = "Zurich BT"
    'oWS.Range(tmp).Font.Size = 9
    oWS.Range(tmp).WrapText = True
Next Cell
 
If Gen_ExcelPassword <> "" Then
   oWS.Protect Gen_ExcelPassword, , , , , , True, True
End If
oWB.Save
'oWB.Close
'oXL.Quit
Call Log_Print_History(TxtPreFix, "CDC", TxtNewCDCNo, Gen_FLName)

MsgBox ("Excel file created !!! ")
oXL.Visible = True

End Sub

Private Sub CmdPrintold_Click()
On Error Resume Next
RsTxn_CDC!PrintFlag = "P"
RsTxn_CDC.Update
'RsTxn_CxTF_Details!Flag = "P"
'RsTxn_CxTF_Details.Update
tmp = "Delete from PrintDetail "
Call TmpTable

tmp = "Delete from PrintMain "
Call TmpTable
tmp = "select * from PrintMain"
Call TmpTable
TmpRs.AddNew
TmpRs!ID = TxtNewCDCNo
TmpRs!F1 = TxtLocationID & "," & TxtCMPAddress
TmpRs!F2 = TxtCMPID
TmpRs!F3 = TxtLotId
TmpRs!F4 = TxtNewClientID & "-" & TxtNewClientName
'TmpRs!F5 = TxtClientName
TmpRs!F6 = TxtCommodity
TmpRs!F7 = TxtNoOfBags
TmpRs!F8 = TxtNetWeight
TmpRs!F9 = TxtAssayerNo & "-" & TxtAssayerName
TmpRs!F10 = TxtGradeByAssayer
TmpRs!F11 = FED.Value
'TmpRs!F12 = TxtRisks
TmpRs!F13 = TxtInsuranceCo
TmpRs!F14 = TxtAmount
TmpRs!F15 = StorageFrom.Value
TmpRs!F16 = StorageTo.Value
TmpRs.Update

TmpRs.Close
cn.Close
If DataEnvironment1.rsCommand1.State <> 0 Then
   DataEnvironment1.rsCommand1.Close
End If

Rpt_CDC.Refresh
Rpt_CDC.Width = 15200
Rpt_CDC.Show

End Sub

Private Sub ExitCmd_Click()
 Unload Me
End Sub

Private Sub FirstCmd_Click()
RsTxn_CDC.MoveFirst
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
   MsgBox "Access denied!!!"
   Exit Sub
End If

mExchangeTypeID = GetExchangeID("SPOT")
Call TableTxn_CDC("Where Flag = 'R' ")

'Call TableTxn_CDC_GSLDetail("Where OldCDCNo <> 0 ")
'Call TableTxn_CDC("Where Flag = 'T' ")

If RsTxn_CDC.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
Else
   Call GButtonLock(Me, True)
   Call Indata
End If
End Sub

Private Sub LastCmd_Click()
RsTxn_CDC.MoveLast
Call Indata

PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub SaveCmd_Click()
If TxtCDCID.Text = "" Then
  MsgBox "Blank CDC Number not allowed!!!"
  TxtCDCID.SetFocus
  Exit Sub
End If
If TxtStorageRate.Text = "" Then
  MsgBox "Blank Storage Rate not allowed!!!"
  TxtStorageRate.SetFocus
  Exit Sub
End If
If TxtNewClientID = "" Then
  MsgBox "Blank Buyer Client-Id not allowed!!!"
  TxtNewClientID.SetFocus
  Exit Sub
End If
'If mClientRowID = mNewClientRowID Then
'  MsgBox "Seller And Buyer Clients Should Be Different!!!"
'  TxtNewClientId.SetFocus
'  Exit Sub
'End If
If ChkOwner() = False Then
  MsgBox "GSL belongs to another CDC!!!"
  TxtCDCID.SetFocus
  Exit Sub
End If

If Edit_Flg = "A" Then
  RsTxn_CDC.AddNew
  RsTxn_CDC!CDCID = GetCDCID
  RsTxn_CDC!G_UID = GetGUID

  'TxtCDCID.Text = RsTxn_CDC!CDCID
Else
  RsTxn_CDC.MoveFirst
  RsTxn_CDC.Find "CDCID= " & TxtCDCID
  If Not RsTxn_CDC.EOF Then
    If RsTxn_CDC!CDCID <> TxtCDCID.Text Then
      MsgBox "Duplicate CDC Number not allowed!!! "
      Exit Sub
    End If
  End If
  RsTxn_CDC.MoveFirst
  RsTxn_CDC.Find " CDCID= " & TxtCDCID.Text
End If

RsTxn_CDC!CDCID = RsTxn_CDC!CDCID
RsTxn_CDC!CDCDate = CDCDateNew.Value
RsTxn_CDC!CMPID = mCMPID
RsTxn_CDC!ClientIDRow = mNewClientRowID
RsTxn_CDC!CommodityID = mCommodityID
RsTxn_CDC!Flag = "R"
RsTxn_CDC!PrintFlag = "0"
RsTxn_CDC!SM_Prefix = CmbPreFix.Text

If IsNull(RsTxn_CDC!CreatedBy) = True Or RsTxn_CDC!CreatedBy = "" Then
   RsTxn_CDC!CreatedBy = Gen_UserLoginID
   RsTxn_CDC!CreatedDate = Now
Else
   RsTxn_CDC!UpdatedBy = Gen_UserLoginID
   RsTxn_CDC!UpdatedDate = Now
End If


RsTxn_CDC.Update
   
Call SaveCDC_GSData
Call SaveCDC_AssayerData
Call SaveCDC_InsuranceData
Call Indata
End Sub
Private Sub SaveCDC_GSData()

RsTxn_CDC_GSLDetail.AddNew
mNewCDCNo = GetCDCNo
RsTxn_CDC_GSLDetail!CDCNo = mNewCDCNo
RsTxn_CDC_GSLDetail!G_UID = GetGUID
RsTxn_CDC_GSLDetail!CDCID = RsTxn_CDC!CDCID
RsTxn_CDC_GSLDetail!GSLID = mGSLID
RsTxn_CDC_GSLDetail!StorageFrom = StorageFrom.Value
RsTxn_CDC_GSLDetail!StorageTo = StorageTo
RsTxn_CDC_GSLDetail!StorageRate = Val(TxtStorageRate)
RsTxn_CDC_GSLDetail!StorageAmount = Val(TxtStorageAmount)
RsTxn_CDC_GSLDetail!CDCDate = Date 'CDCDate.Value 'MSHFlexGrid1.TextMatrix(i, 16)
RsTxn_CDC_GSLDetail!PrintFlag = "0" 'MSHFlexGrid1.TextMatrix(i, 18)
RsTxn_CDC_GSLDetail!Flag = "F" 'MSHFlexGrid1.TextMatrix(i, 18)
RsTxn_CDC_GSLDetail!OldCDCNo = TxtCDCID
RsTxn_CDC_GSLDetail!CdtfNo = TxtCDTFNo
RsTxn_CDC_GSLDetail!CDTFDate = TxtCDTFDate
RsTxn_CDC_GSLDetail!ClientIDRow = mNewClientRowID
RsTxn_CDC_GSLDetail!SM_Prefix = CmbPreFix.Text
RsTxn_CDC_GSLDetail.Update
RsTxn_CDC_GSLDetail.MoveFirst
RsTxn_CDC_GSLDetail.Find "CDCNo = " & TxtCDCID
If Not RsTxn_CDC_GSLDetail.EOF Then
   RsTxn_CDC_GSLDetail!Flag = "T"
   RsTxn_CDC_GSLDetail.Update
End If


'RsTxn_CDC_GSLDetail.Requery
Call TableMst_GSL("Where sm_Prefix ='" & CmbPreFix.Text & "' And GSLID = " & mGSLID & "")
RsMst_GSL!ISINID = mNewCDCNo 'RsTxn_CDC_GSLDetail!CDCNo
RsMst_GSL!Flag = "O" 'Check1.Value
RsMst_GSL.Update
RsMst_GSL.Close

End Sub
Private Sub SaveCDC_AssayerData()

'Call TableTxn_CDC_AssayerDetail(" Where CDCID = " & TxtCDCID)
RsTxn_CDC_AssayerDetail.Find "CDCID = " & RsTxn_CDC!CDCID
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
RsTxn_CDC_AssayerDetail!SM_Prefix = CmbPreFix.Text
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
RsTxn_CDC_InsuranceDetail!SM_Prefix = CmbPreFix.Text
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
tmp = "Select max(CDCNo) from Txn_CDC_GSL_Detail where sm_Prefix ='" & CmbPreFix.Text & "'"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetCDCNo = Retval
End Function

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
  Call TableTxn_CDC_GSLDetail("Where OldCDCNo <> 0 ")
  If RsTxn_CDC_GSLDetail.RecordCount = 0 Then Exit Sub
  Call ClearData
  Call GButtonLock(Me, True)
  Call FirstCmd_Click
  Exit Sub
End If

End Sub

Private Sub TxtCDCID_LostFocus()
If TxtCDCID = "" Then Exit Sub
If CmbPreFix.Text = "" Then Exit Sub

Call TableTxn_CDC_GSLDetail(" Where sm_Prefix ='" & CmbPreFix.Text & "'")
RsTxn_CDC_GSLDetail.Find "CDCNo = " & TxtCDCID
If Not RsTxn_CDC_GSLDetail.EOF Then
  If RsTxn_CDC_GSLDetail!Flag <> "P" Then
     MsgBox "Do Pay Off first!!!"
     Exit Sub
  End If
  mCDCID = RsTxn_CDC_GSLDetail!CDCID
  Call TableTxn_CDC("Where sm_Prefix ='" & CmbPreFix.Text & "' And CDCID = " & mCDCID & "")
  Call TableTxn_CDC_AssayerDetail("Where  sm_Prefix ='" & CmbPreFix.Text & "' And CDCID = " & mCDCID & "")
  Call TableTxn_CDC_InsuranceDetail("Where  sm_Prefix ='" & CmbPreFix.Text & "' And CDCID = " & mCDCID & "")
  Call PopulateGSLData
  Call PopulateAssayerData
  Call PopulateInsuranceData
  mGSLID = RsTxn_CDC_GSLDetail!GSLID
Else
  MsgBox "There is no such CDC No!!!"
  TxtCDCID.SetFocus
  Exit Sub
End If
End Sub
Private Sub PopulateGSLData()

mClientRowID = RsTxn_CDC_GSLDetail!ClientIDRowSeller
mNewClientRowID = RsTxn_CDC_GSLDetail!ClientIDRow
mCMPID = RsTxn_CDC!CMPID
CDCDate.Value = RsTxn_CDC!CDCDate
mCommodityID = RsTxn_CDC!CommodityID
mCDTFNo = RsTxn_CDC_GSLDetail!CdtfNo
mCDTFDate = RsTxn_CDC_GSLDetail!CDTFDate

ans = GetClientNameByRow(mClientRowID)
TxtClientID = ans(0)
TxtClientName = ans(1)
TxtDPID = ans(2)
TxtDPName = ans(3)




ans = GetClientNameByRow(mNewClientRowID)
TxtNewClientID = ans(0)
TxtNewClientName = ans(1)
TxtNewDPID = ans(2)
TxtNewDPName = ans(3)

TxtCMPID = GetCMPName(mCMPID)
TxtCMPAddress = GetCMPAddress(mCMPID, "A")
TxtLocationID = GetLocationFromCMP(mCMPID)
TxtCommodity = GetCommodityName(mCommodityID)
TxtCommodityCondition = IIf(IsNull(RsTxn_CDC!CommodityConditon), "", RsTxn_CDC!CommodityConditon)

'tmp = " Where  Txn_CxTF_GSL.CxTFNo = " & RsTxn_CDC_GSLDetail!CDTFNo & "  And "
'tmp = tmp & "  Txn_CxTF_GSL.GSLID = " & RsTxn_CDC_GSLDetail!GSLID

tmp = "  Where  sm_Prefix ='" & CmbPreFix.Text & "' And Txn_CxTF_GSL.GSLID = " & RsTxn_CDC_GSLDetail!GSLID

Call TableTxn_CxTF_GSL(tmp)


TxtCDTFNo = RsTxn_CDC_GSLDetail!CdtfNo
TxtCDTFDate = RsTxn_CDC_GSLDetail!CDTFDate

ans = GetGSLByGLSID(RsTxn_CDC_GSLDetail!GSLID, CmbPreFix.Text)
TxtStorageRate = RsTxn_CDC_GSLDetail!StorageChargesPayoffRate

TxtGodownAddress = ans(11)
TxtGodownID = ans(3)
TxtStackId = ans(4)
TxtLotId = ans(5)
For j = 1 To RsTxn_CxTF_GSL.RecordCount
    TxtNoOfBags = Val(TxtNoOfBags) + RsTxn_CxTF_GSL!NoofBags
    TxtQuantity = Val(TxtQuantity) + RsTxn_CxTF_GSL!AdjNCMSLWeight
    TxtQuantityKG = Val(TxtQuantityKG) + RsTxn_CxTF_GSL!AdjNCMSLWeight * 1000
    TxtGunnyWeight = Val(TxtGunnyWeight) + RsTxn_CxTF_GSL!GunnyWeight * 1000
    TxtStdDeduction = Val(TxtStdDeduction) + RsTxn_CxTF_GSL!StdDeduction * 1000 'TmpRs!StdDeduction
    TxtNetWeight = Val(TxtNetWeight) + RsTxn_CxTF_GSL!DematWeight * 1000 '  net weight in KGS
    RsTxn_CxTF_GSL.MoveNext
Next
StorageFrom.Value = RsTxn_CDC_GSLDetail!PayoffDate + 1
StorageTo.Value = RsTxn_CDC_GSLDetail!PayoffDate + Val(TxtNoofDays)  ' 7
StorageTo.Enabled = True
LblStorageDays.Caption = (StorageTo.Value - StorageFrom.Value) + 1

'TxtNewClientId = ""

End Sub
Private Sub PopulateAssayerData()

TxtAssayerNo = RsTxn_CDC_AssayerDetail!AssayerReportNo
TxtAssayerName = RsTxn_CDC_AssayerDetail!AssayerName
ARDate.Value = RsTxn_CDC_AssayerDetail!AssayerReportDate
TxtMoisture = RsTxn_CDC_AssayerDetail!Moisture
TxtDamaged = RsTxn_CDC_AssayerDetail!Damaged
TxtForeign = RsTxn_CDC_AssayerDetail!Foreign
TxtPolarization = RsTxn_CDC_AssayerDetail!Polarization
TxtGrainSize = RsTxn_CDC_AssayerDetail!GrainSize
TxtOther = RsTxn_CDC_AssayerDetail!OverAllQuality
TxtBrokan = IIf(IsNull(RsTxn_CDC_AssayerDetail!Brokan), "", RsTxn_CDC_AssayerDetail!Brokan)
TxtOil = IIf(IsNull(RsTxn_CDC_AssayerDetail!OilContain), "", RsTxn_CDC_AssayerDetail!OilContain) '
TxtColour = IIf(IsNull(RsTxn_CDC_AssayerDetail!colour), "", RsTxn_CDC_AssayerDetail!colour) '
TxtAgeing = IIf(IsNull(RsTxn_CDC_AssayerDetail!Ageing), "", RsTxn_CDC_AssayerDetail!Ageing) '
TxtStapleLength = IIf(IsNull(RsTxn_CDC_AssayerDetail!StapleLength), "", RsTxn_CDC_AssayerDetail!StapleLength) '
TxtOrganic = IIf(IsNull(RsTxn_CDC_AssayerDetail!OrganicInorganic), "", RsTxn_CDC_AssayerDetail!OrganicInorganic) '
TxtWax = IIf(IsNull(RsTxn_CDC_AssayerDetail!WaxHoney), "", RsTxn_CDC_AssayerDetail!WaxHoney) '
TxtGradeByAssayer = IIf(IsNull(RsTxn_CDC_AssayerDetail!Grade), "", RsTxn_CDC_AssayerDetail!Grade) '
NED.Value = RsTxn_CDC_AssayerDetail!NED
FED.Value = RsTxn_CDC_AssayerDetail!FED
ChkAcceptFlag = RsTxn_CDC_AssayerDetail!AcceptenceFlag
TxtRemark = RsTxn_CDC_AssayerDetail!Remarks

End Sub
Private Sub PopulateInsuranceData()

If RsTxn_CDC_InsuranceDetail.RecordCount = 0 Then Exit Sub
TxtPolicyNo = RsTxn_CDC_InsuranceDetail!PolicyNo & " ~ " & RsTxn_CDC_InsuranceDetail!CoverNoteNo
Call TableMst_PolicyDetail("where PolicyNo = '" & RsTxn_CDC_InsuranceDetail!PolicyNo & "' And CoverNoteNo = '" & RsTxn_CDC_InsuranceDetail!CoverNoteNo & "'")
TxtInsuranceCo = GetInsuranceCoName(RsMst_PolicyDetail!InsuranceCoID)
ValidFrom.Value = RsMst_PolicyDetail!ValidFrom
ValidTo.Value = RsMst_PolicyDetail!ValidTo
TxtAmount = RsMst_PolicyDetail!Amount
TxtRisks = RsMst_PolicyDetail!Risks
RsMst_PolicyDetail.Close

End Sub
Public Sub Indata()
Call ClearData
CmbPreFix.Visible = False
TxtPreFix.Visible = True
TxtQuantityKG = ""
'RsTxn_CDC_GSLDetail.MoveFirst
Call TableMst_Config
TxtNoofDays = IIf(IsNull(RsMst_Config!SportStorageChargesDays), 0, RsMst_Config!SportStorageChargesDays)
TxtPreFix = RsTxn_CDC!SM_Prefix
RsMst_Config.Close

TxtCreated = IIf(IsNull(RsTxn_CDC!CreatedBy), "", RsTxn_CDC!CreatedBy) & " :- " & IIf(IsNull(RsTxn_CDC!CreatedDate), "", RsTxn_CDC!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_CDC!UpdatedBy), "", RsTxn_CDC!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_CDC!UpdatedDate), "", RsTxn_CDC!UpdatedDate)


Call TableTxn_CDC_GSLDetail("Where  sm_Prefix ='" & TxtPreFix.Text & "' And CDCID = " & RsTxn_CDC!CDCID)
mCDCID = RsTxn_CDC_GSLDetail!CDCID
Call TableTxn_CDC_AssayerDetail("Where  sm_Prefix ='" & TxtPreFix.Text & "' And CDCID = " & mCDCID & "")
Call TableTxn_CDC_InsuranceDetail("Where  sm_Prefix ='" & TxtPreFix.Text & "' And CDCID = " & mCDCID & "")
Call IndataCDCReIssue_GSL
Call InDataCDCReIssue_Assayer
Call IndataCDCReIssue_Insurance
Call GButtonLock(Me, True)
CmdPrint.Enabled = True
EditCmd.Enabled = False
DeleteCmd.Enabled = False
SearchCmd.Enabled = False

End Sub

Private Sub IndataCDCReIssue_GSL()
TxtCDCID = RsTxn_CDC_GSLDetail!OldCDCNo
TxtNewCDCNo = RsTxn_CDC_GSLDetail!CDCNo
 
TxtCDCID.Locked = True
CDCDate.Value = RsTxn_CDC_GSLDetail!CDCDate


mClientRowID = RsTxn_CDC!ClientIDRow
mCMPID = RsTxn_CDC!CMPID
mCommodityID = RsTxn_CDC!CommodityID
mCDTFNo = RsTxn_CDC_GSLDetail!CdtfNo
mCDTFDate = RsTxn_CDC_GSLDetail!CDTFDate

CDCDateNew.Value = RsTxn_CDC!CDCDate

ans = GetClientNameByRow(mClientRowID)
TxtNewClientID = ans(0)
TxtNewClientName = ans(1)
TxtNewDPID = ans(2)
TxtNewDPName = ans(3)
TxtCMPID = GetCMPName(mCMPID)
TxtCMPAddress = GetCMPAddress(mCMPID, "A")
TxtLocationID = GetLocationFromCMP(mCMPID)
TxtCommodity = GetCommodityName(mCommodityID)
TxtCommodityCondition = IIf(IsNull(RsTxn_CDC!CommodityConditon), "", RsTxn_CDC!CommodityConditon)

'tmp = " Where  Txn_CxTF_GSL.CxTFNo = " & RsTxn_CDC_GSLDetail!CDTFNo & "  And "
'tmp = tmp & "  Txn_CxTF_GSL.GSLID = " & RsTxn_CDC_GSLDetail!GSLID

'tmp = " Where  Txn_CxTF_GSL.CxTFNo = " & RsTxn_CDC_GSLDetail!CDTFNo & "  And "
tmp = "  where  sm_Prefix ='" & TxtPreFix.Text & "' And Txn_CxTF_GSL.GSLID = " & RsTxn_CDC_GSLDetail!GSLID

Call TableTxn_CxTF_GSL(tmp)


TxtCDTFNo = RsTxn_CDC_GSLDetail!CdtfNo
TxtCDTFDate = RsTxn_CDC_GSLDetail!CDTFDate

ans = GetGSLByGLSID(RsTxn_CDC_GSLDetail!GSLID, TxtPreFix)

TxtGodownAddress = ans(11)
TxtGodownID = ans(3)
TxtStackId = ans(4)
TxtLotId = ans(5)

For j = 1 To RsTxn_CxTF_GSL.RecordCount
    TxtNoOfBags = Val(TxtNoOfBags) + RsTxn_CxTF_GSL!NoofBags
    TxtQuantity = Val(TxtQuantity) + RsTxn_CxTF_GSL!AdjNCMSLWeight
    TxtQuantityKG = Val(TxtQuantityKG) + RsTxn_CxTF_GSL!AdjNCMSLWeight * 1000
    TxtGunnyWeight = Val(TxtGunnyWeight) + RsTxn_CxTF_GSL!GunnyWeight * 1000
    TxtStdDeduction = Val(TxtStdDeduction) + RsTxn_CxTF_GSL!StdDeduction * 1000 'TmpRs!StdDeduction
    TxtNetWeight = Val(TxtNetWeight) + RsTxn_CxTF_GSL!DematWeight * 1000 '  net weight in KGS
    RsTxn_CxTF_GSL.MoveNext
Next

TxtStorageRate = RsTxn_CDC_GSLDetail!StorageRate
TxtStorageRate.Locked = True
StorageFrom.Value = RsTxn_CDC_GSLDetail!StorageFrom
StorageTo.Value = RsTxn_CDC_GSLDetail!StorageTo
StorageTo.Enabled = False
LblStorageDays.Caption = (StorageTo.Value - StorageFrom.Value) + 1
TxtStorageAmount = Val(RsTxn_CDC_GSLDetail!StorageAmount) 'Val(TxtQuantity) * Val(TxtStorageRate) * Val(LblStorageDays.Caption)
TxtNewClientID.Locked = True


tmp = "select * from Txn_CDC_GSL_Detail where  sm_Prefix ='" & TxtPreFix.Text & "' And CDCNo = " & RsTxn_CDC_GSLDetail!OldCDCNo
Call TmpTable
If TmpRs.RecordCount > 0 Then
   ans = GetClientNameByRow(TmpRs!ClientIDRowSeller)
   TxtClientID = ans(0)
   TxtClientName = ans(1)
   TxtDPID = ans(2)
   TxtDPName = ans(3)
End If



End Sub
Private Sub InDataCDCReIssue_Assayer()

'Call TableTxn_CDC_AssayerDetail("where CDCID = " & RsTxn_CDC!CDCID & "")
If RsTxn_CDC_AssayerDetail.RecordCount = 0 Then Exit Sub

TxtAssayerNo = RsTxn_CDC_AssayerDetail!AssayerReportNo
TxtAssayerName = RsTxn_CDC_AssayerDetail!AssayerName
ARDate.Value = RsTxn_CDC_AssayerDetail!AssayerReportDate
TxtMoisture = RsTxn_CDC_AssayerDetail!Moisture
TxtDamaged = RsTxn_CDC_AssayerDetail!Damaged
TxtForeign = RsTxn_CDC_AssayerDetail!Foreign
TxtPolarization = RsTxn_CDC_AssayerDetail!Polarization
TxtGrainSize = RsTxn_CDC_AssayerDetail!GrainSize
TxtBrokan = RsTxn_CDC_AssayerDetail!Brokan
TxtOil = RsTxn_CDC_AssayerDetail!OilContain
TxtColour = RsTxn_CDC_AssayerDetail!colour
TxtAgeing = RsTxn_CDC_AssayerDetail!Ageing
TxtStapleLength = RsTxn_CDC_AssayerDetail!StapleLength
TxtOrganic = RsTxn_CDC_AssayerDetail!OrganicInorganic
TxtWax = RsTxn_CDC_AssayerDetail!WaxHoney
TxtGradeByAssayer = RsTxn_CDC_AssayerDetail!Grade
NED.Value = RsTxn_CDC_AssayerDetail!NED
FED.Value = RsTxn_CDC_AssayerDetail!FED
TxtOther = RsTxn_CDC_AssayerDetail!OverAllQuality  'OverAll Quality
ChkAcceptFlag.Value = RsTxn_CDC_AssayerDetail!AcceptenceFlag
TxtRemark = RsTxn_CDC_AssayerDetail!Remarks

End Sub
Private Sub IndataCDCReIssue_Insurance()

'Call TableTxn_CDC_InsuranceDetail("Where CDCID = " & RsTxn_CDC!CDCID & "")
If RsTxn_CDC_InsuranceDetail.RecordCount = 0 Then Exit Sub

TxtPolicyNo = RsTxn_CDC_InsuranceDetail!PolicyNo & " ~ " & RsTxn_CDC_InsuranceDetail!CoverNoteNo
Call TableMst_PolicyDetail("where PolicyNo = '" & RsTxn_CDC_InsuranceDetail!PolicyNo & "' And CoverNoteNo = '" & RsTxn_CDC_InsuranceDetail!CoverNoteNo & "'")
TxtInsuranceCo = GetInsuranceCoName(RsMst_PolicyDetail!InsuranceCoID)
ValidFrom.Value = RsMst_PolicyDetail!ValidFrom
ValidTo.Value = RsMst_PolicyDetail!ValidTo
TxtAmount = RsMst_PolicyDetail!Amount
TxtRisks = RsMst_PolicyDetail!Risks
RsMst_PolicyDetail.Close


End Sub
Public Sub ClearData()
TxtCDCID.Text = ""
TxtPreFix = ""
CDCDateNew.Value = Date
TxtCMPID = ""
TxtCMPAddress = ""
TxtCommodity = ""
TxtQuantityKG = ""
TxtCommodityCondition = ""
TxtLocationID = ""
TxtClientID = ""
TxtClientName = ""
TxtDPID = ""
TxtDPName = ""
TxtGodownAddress = ""
TxtGodownID = ""
TxtStackId = ""
TxtLotId = ""
TxtNoOfBags = ""
TxtQuantity = ""
TxtGunnyWeight = ""
TxtStdDeduction = ""
TxtNetWeight = ""
TxtStorageAmount = ""
TxtNewClientID = ""
TxtNewClientName = ""
TxtNewDPID = ""
TxtNewDPName = ""
TxtStorageRate = ""
TxtAssayerNo = ""
TxtAssayerName = ""
ARDate.Value = Date
TxtMoisture = ""
TxtDamaged = ""
TxtForeign = ""
TxtPolarization = ""
TxtGrainSize = ""
TxtBrokan = ""
TxtOil = ""
TxtColour = ""
TxtAgeing = ""
TxtStapleLength = ""
TxtOrganic = ""
TxtWax = ""
TxtGradeByAssayer = ""
NED.Value = Date
FED.Value = Date
TxtOther = ""  'OverAll Quality
ChkAcceptFlag.Value = 0
TxtRemark = ""
ValidFrom.Value = Date
ValidTo.Value = Date
TxtAmount = ""
TxtRisks = ""
TxtInsuranceCo.Text = ""
TxtPolicyNo.Text = ""
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
   MsgBox "Not found!!!"
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
      "WHERE (((Mst_Depository.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtNewClientID & "'));"
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
   MsgBox "Invalid selection!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If
TmpRs1.MoveFirst


TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection!!! "
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
If StorageFrom.Value > StorageTo.Value Then
    MsgBox "From Date must be less than To Date!!!"
    StorageTo.SetFocus
    Exit Sub
Else
    mStorageDays = StorageTo.Value - StorageFrom + 1
    LblStorageDays = mStorageDays '- 7 'mStorageDays - 6
    TxtStorageAmount = Val(TxtQuantity) * Val(TxtStorageRate) * Val(LblStorageDays.Caption)
End If
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_CDC.MoveNext
If RsTxn_CDC.EOF Then
   RsTxn_CDC.MoveLast
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
RsTxn_CDC.MovePrevious

If RsTxn_CDC.BOF Then
   RsTxn_CDC.MoveFirst
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
Call TableMst_GSL("Where sm_Prefix ='" & CmbPreFix.Text & "' And ISINID = '" & TxtCDCID & "'")
If RsMst_GSL.RecordCount = 0 Then
  Retval = False
End If
ChkOwner = Retval
End Function
Private Sub TxtStorageRate_LostFocus()
If TxtStorageRate.Text = "" Then Exit Sub
If IsNumeric(TxtStorageRate) = False Then
   MsgBox "Invalid numeric format!!!"
   TxtStorageRate.SetFocus
   Exit Sub
End If
TxtStorageAmount = Val(TxtQuantity) * Val(TxtStorageRate) * Val(LblStorageDays.Caption)
End Sub
