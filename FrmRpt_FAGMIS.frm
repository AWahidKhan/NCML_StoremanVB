VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_FAGMIS 
   Caption         =   "FAG MIS"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   10890
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   9840
      Width           =   10935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8895
      Left            =   15240
      TabIndex        =   0
      Top             =   -1560
      Width           =   255
   End
   Begin VB.PictureBox Frame4 
      Height          =   9660
      Left            =   0
      ScaleHeight     =   9600
      ScaleWidth      =   14895
      TabIndex        =   7
      Top             =   0
      Width           =   14955
      Begin VB.Frame Frame11 
         Height          =   1020
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   14760
         Begin VB.Frame Frame1 
            Caption         =   "Excel files to be generated"
            Height          =   795
            Left            =   10920
            TabIndex        =   57
            Top             =   105
            Width           =   3780
            Begin VB.CheckBox Check3 
               Caption         =   "Invoices"
               Height          =   255
               Left            =   1200
               TabIndex        =   60
               Top             =   360
               Width           =   975
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Cash Receipt"
               Height          =   255
               Left            =   2340
               TabIndex        =   59
               Top             =   360
               Width           =   1335
            End
            Begin VB.CheckBox Check1 
               Caption         =   "MIS"
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   360
               Value           =   1  'Checked
               Width           =   1215
            End
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
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   405
            Visible         =   0   'False
            Width           =   3750
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
            Left            =   8430
            TabIndex        =   4
            Top             =   390
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
            Left            =   7200
            TabIndex        =   3
            Top             =   390
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
            Left            =   9660
            TabIndex        =   5
            Top             =   390
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker TxtFromDate 
            Height          =   330
            Left            =   120
            TabIndex        =   1
            Top             =   390
            Width           =   1695
            _ExtentX        =   2990
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
            Format          =   74383363
            CurrentDate     =   39850
            MinDate         =   38353
         End
         Begin MSComCtl2.DTPicker TxtToDate 
            Height          =   330
            Left            =   1875
            TabIndex        =   2
            Top             =   405
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   74383363
            CurrentDate     =   39850
            MinDate         =   38353
         End
         Begin VB.Label Label2 
            Caption         =   "To Month Year"
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
            Left            =   1920
            TabIndex        =   17
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "State"
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
            Left            =   4650
            TabIndex        =   11
            Top             =   135
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "From Month Year"
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
            TabIndex        =   10
            Top             =   120
            Width           =   1695
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   8415
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   14760
         _ExtentX        =   26035
         _ExtentY        =   14843
         _Version        =   393216
         Tabs            =   14
         TabsPerRow      =   7
         TabHeight       =   1058
         TabCaption(0)   =   "FAG MIS"
         TabPicture(0)   =   "FrmRpt_FAGMIS.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MSHFlexGrid1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Mapping Pending"
         TabPicture(1)   =   "FrmRpt_FAGMIS.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "MSHFlexGrid2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Outstanding Details"
         TabPicture(2)   =   "FrmRpt_FAGMIS.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "MSHFlexGrid3"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "StoreMan Invoice (with Doc No)"
         TabPicture(3)   =   "FrmRpt_FAGMIS.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "MSHFlexGrid4"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Invoices(without Doc No)"
         TabPicture(4)   =   "FrmRpt_FAGMIS.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "MSHFlexGrid5"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Advance Remaining (Excl. Adv. allocation)"
         TabPicture(5)   =   "FrmRpt_FAGMIS.frx":008C
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "MSHFlexGrid6"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Allocation via Advance Allocation"
         TabPicture(6)   =   "FrmRpt_FAGMIS.frx":00A8
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "MSHFlexGrid7"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "NNNR Invoice Details"
         TabPicture(7)   =   "FrmRpt_FAGMIS.frx":00C4
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "SSTab3"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "Ad-Hoc Invoice Details"
         TabPicture(8)   =   "FrmRpt_FAGMIS.frx":00E0
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "SSTab4"
         Tab(8).ControlCount=   1
         TabCaption(9)   =   "Reservation Invoice Details"
         TabPicture(9)   =   "FrmRpt_FAGMIS.frx":00FC
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "SSTab5"
         Tab(9).ControlCount=   1
         TabCaption(10)  =   "DP Invoice Details"
         TabPicture(10)  =   "FrmRpt_FAGMIS.frx":0118
         Tab(10).ControlEnabled=   0   'False
         Tab(10).Control(0)=   "SSTab6"
         Tab(10).ControlCount=   1
         TabCaption(11)  =   "Cash Receipt Details"
         TabPicture(11)  =   "FrmRpt_FAGMIS.frx":0134
         Tab(11).ControlEnabled=   0   'False
         Tab(11).Control(0)=   "SSTab2"
         Tab(11).ControlCount=   1
         TabCaption(12)  =   "Advance Cash Receipt Details"
         TabPicture(12)  =   "FrmRpt_FAGMIS.frx":0150
         Tab(12).ControlEnabled=   0   'False
         Tab(12).Control(0)=   "SSTab7"
         Tab(12).ControlCount=   1
         TabCaption(13)  =   "Other Adjustment And TDS Details"
         TabPicture(13)  =   "FrmRpt_FAGMIS.frx":016C
         Tab(13).ControlEnabled=   0   'False
         Tab(13).Control(0)=   "SSTab8"
         Tab(13).ControlCount=   1
         Begin TabDlg.SSTab SSTab8 
            Height          =   6975
            Left            =   -74880
            TabIndex        =   52
            Top             =   1320
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   12303
            _Version        =   393216
            Tabs            =   4
            Tab             =   2
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Adjust and TDS Details (Consolidated)"
            TabPicture(0)   =   "FrmRpt_FAGMIS.frx":0188
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "MSHFlexGrid14"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Pending for Approval by FAG"
            TabPicture(1)   =   "FrmRpt_FAGMIS.frx":01A4
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "MSHFlexGrid34"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Batch Pending By FAG"
            TabPicture(2)   =   "FrmRpt_FAGMIS.frx":01C0
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "MSHFlexGrid35"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Pending Import of Doc No by FAG"
            TabPicture(3)   =   "FrmRpt_FAGMIS.frx":01DC
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "MSHFlexGrid36"
            Tab(3).ControlCount=   1
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid14 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   53
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid34 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   54
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid35 
               Height          =   6465
               Left            =   120
               TabIndex        =   55
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid36 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   56
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
         End
         Begin TabDlg.SSTab SSTab7 
            Height          =   6975
            Left            =   -74880
            TabIndex        =   48
            Top             =   1320
            Width           =   14385
            _ExtentX        =   25374
            _ExtentY        =   12303
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "Advance Cash Receipt(Consolidated)"
            TabPicture(0)   =   "FrmRpt_FAGMIS.frx":01F8
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "MSHFlexGrid13"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Pending Import of Doc No by FAG"
            TabPicture(1)   =   "FrmRpt_FAGMIS.frx":0214
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "MSHFlexGrid32"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Pending Invoice Allocation by FAG"
            TabPicture(2)   =   "FrmRpt_FAGMIS.frx":0230
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "MSHFlexGrid33"
            Tab(2).ControlCount=   1
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid13 
               Height          =   6465
               Left            =   120
               TabIndex        =   49
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid32 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   50
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid33 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   51
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
         End
         Begin TabDlg.SSTab SSTab5 
            Height          =   6975
            Left            =   -74880
            TabIndex        =   37
            Top             =   1320
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   12303
            _Version        =   393216
            Tabs            =   4
            Tab             =   3
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Reservation Inv Details(Consolidated)"
            TabPicture(0)   =   "FrmRpt_FAGMIS.frx":024C
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "MSHFlexGrid10"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Pending for Approval by FAG"
            TabPicture(1)   =   "FrmRpt_FAGMIS.frx":0268
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "MSHFlexGrid25"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Batch pending by FAG"
            TabPicture(2)   =   "FrmRpt_FAGMIS.frx":0284
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "MSHFlexGrid26"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Pending import of TI No"
            TabPicture(3)   =   "FrmRpt_FAGMIS.frx":02A0
            Tab(3).ControlEnabled=   -1  'True
            Tab(3).Control(0)=   "MSHFlexGrid27"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).ControlCount=   1
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid10 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   38
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid25 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   39
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid26 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   40
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid27 
               Height          =   6465
               Left            =   120
               TabIndex        =   41
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
         End
         Begin TabDlg.SSTab SSTab4 
            Height          =   6975
            Left            =   -74880
            TabIndex        =   32
            Top             =   1320
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   12303
            _Version        =   393216
            Tabs            =   4
            Tab             =   3
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Ad-Hoc Invoice Details(Consolidated)"
            TabPicture(0)   =   "FrmRpt_FAGMIS.frx":02BC
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "MSHFlexGrid9"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Pending for Approval by FAG"
            TabPicture(1)   =   "FrmRpt_FAGMIS.frx":02D8
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "MSHFlexGrid22"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Batch pending by FAG"
            TabPicture(2)   =   "FrmRpt_FAGMIS.frx":02F4
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "MSHFlexGrid23"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Pending import of TI No"
            TabPicture(3)   =   "FrmRpt_FAGMIS.frx":0310
            Tab(3).ControlEnabled=   -1  'True
            Tab(3).Control(0)=   "MSHFlexGrid24"
            Tab(3).Control(0).Enabled=   0   'False
            Tab(3).ControlCount=   1
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid9 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   33
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid22 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   34
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid23 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   35
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid24 
               Height          =   6465
               Left            =   120
               TabIndex        =   36
               Top             =   360
               Width           =   14160
               _ExtentX        =   24977
               _ExtentY        =   11404
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
         End
         Begin TabDlg.SSTab SSTab3 
            Height          =   6975
            Left            =   -74880
            TabIndex        =   28
            Top             =   1320
            Width           =   14295
            _ExtentX        =   25215
            _ExtentY        =   12303
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "NNNR Invoice Details(Consolidated)"
            TabPicture(0)   =   "FrmRpt_FAGMIS.frx":032C
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "MSHFlexGrid8"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Batch Pending By FAG"
            TabPicture(1)   =   "FrmRpt_FAGMIS.frx":0348
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "MSHFlexGrid20"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Pending Import of TI No"
            TabPicture(2)   =   "FrmRpt_FAGMIS.frx":0364
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "MSHFlexGrid21"
            Tab(2).ControlCount=   1
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid8 
               Height          =   6465
               Left            =   120
               TabIndex        =   29
               Top             =   360
               Width           =   14040
               _ExtentX        =   24765
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid20 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   30
               Top             =   360
               Width           =   14040
               _ExtentX        =   24765
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid21 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   31
               Top             =   360
               Width           =   14040
               _ExtentX        =   24765
               _ExtentY        =   11404
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
         End
         Begin TabDlg.SSTab SSTab2 
            Height          =   6975
            Left            =   -74880
            TabIndex        =   21
            Top             =   1320
            Width           =   14175
            _ExtentX        =   25003
            _ExtentY        =   12303
            _Version        =   393216
            Tabs            =   6
            Tab             =   5
            TabsPerRow      =   6
            TabHeight       =   1058
            TabCaption(0)   =   "Cash Receipt Details(Consoladated)"
            TabPicture(0)   =   "FrmRpt_FAGMIS.frx":0380
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "MSHFlexGrid12"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Cash Receipt Instrument Not Deposited in Bank (By CMPM/CMG)"
            TabPicture(1)   =   "FrmRpt_FAGMIS.frx":039C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "MSHFlexGrid15"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Cash Receipt Pending for Approval By FAG"
            TabPicture(2)   =   "FrmRpt_FAGMIS.frx":03B8
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "MSHFlexGrid17"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Cash Receipt Batch Pending By FAG"
            TabPicture(3)   =   "FrmRpt_FAGMIS.frx":03D4
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "MSHFlexGrid16"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "CR Pending Import of Doc No from Navision (FAG)"
            TabPicture(4)   =   "FrmRpt_FAGMIS.frx":03F0
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "MSHFlexGrid18"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "CR Pending Invoice Allocation By FAG"
            TabPicture(5)   =   "FrmRpt_FAGMIS.frx":040C
            Tab(5).ControlEnabled=   -1  'True
            Tab(5).Control(0)=   "MSHFlexGrid19"
            Tab(5).Control(0).Enabled=   0   'False
            Tab(5).ControlCount=   1
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid12 
               Height          =   6105
               Left            =   -74880
               TabIndex        =   22
               Top             =   720
               Width           =   13920
               _ExtentX        =   24553
               _ExtentY        =   10769
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid15 
               Height          =   6105
               Left            =   -74880
               TabIndex        =   23
               Top             =   720
               Width           =   13920
               _ExtentX        =   24553
               _ExtentY        =   10769
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid17 
               Height          =   6105
               Left            =   -74880
               TabIndex        =   24
               Top             =   720
               Width           =   13920
               _ExtentX        =   24553
               _ExtentY        =   10769
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid16 
               Height          =   6105
               Left            =   -74880
               TabIndex        =   25
               Top             =   720
               Width           =   13920
               _ExtentX        =   24553
               _ExtentY        =   10769
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid18 
               Height          =   6105
               Left            =   -74880
               TabIndex        =   26
               Top             =   720
               Width           =   13920
               _ExtentX        =   24553
               _ExtentY        =   10769
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid19 
               Height          =   6105
               Left            =   120
               TabIndex        =   27
               Top             =   720
               Width           =   13920
               _ExtentX        =   24553
               _ExtentY        =   10769
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
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   6945
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   14400
            _ExtentX        =   25400
            _ExtentY        =   12250
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   6945
            Left            =   -74880
            TabIndex        =   14
            Top             =   1320
            Width           =   14400
            _ExtentX        =   25400
            _ExtentY        =   12250
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   6945
            Left            =   -74880
            TabIndex        =   15
            Top             =   1320
            Width           =   14400
            _ExtentX        =   25400
            _ExtentY        =   12250
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
            Height          =   6945
            Left            =   -74880
            TabIndex        =   16
            Top             =   1320
            Width           =   14400
            _ExtentX        =   25400
            _ExtentY        =   12250
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid5 
            Height          =   6945
            Left            =   -74880
            TabIndex        =   18
            Top             =   1320
            Width           =   14400
            _ExtentX        =   25400
            _ExtentY        =   12250
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
            Height          =   6945
            Left            =   -74880
            TabIndex        =   19
            Top             =   1320
            Width           =   14400
            _ExtentX        =   25400
            _ExtentY        =   12250
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid7 
            Height          =   6945
            Left            =   -74880
            TabIndex        =   20
            Top             =   1320
            Width           =   14400
            _ExtentX        =   25400
            _ExtentY        =   12250
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
         Begin TabDlg.SSTab SSTab6 
            Height          =   6975
            Left            =   -74880
            TabIndex        =   42
            Top             =   1320
            Width           =   14295
            _ExtentX        =   25215
            _ExtentY        =   12303
            _Version        =   393216
            Tab             =   1
            TabHeight       =   520
            TabCaption(0)   =   "DP Invoice Details(Consolidated)"
            TabPicture(0)   =   "FrmRpt_FAGMIS.frx":0428
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "MSHFlexGrid11"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Batch Pending By FAG"
            TabPicture(1)   =   "FrmRpt_FAGMIS.frx":0444
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "MSHFlexGrid28"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Pending Import of TI No"
            TabPicture(2)   =   "FrmRpt_FAGMIS.frx":0460
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "MSHFlexGrid31"
            Tab(2).ControlCount=   1
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid29 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   43
               Top             =   360
               Width           =   14040
               _ExtentX        =   24765
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid30 
               Height          =   6465
               Left            =   -74880
               TabIndex        =   44
               Top             =   360
               Width           =   14040
               _ExtentX        =   24765
               _ExtentY        =   11404
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid11 
               Height          =   6945
               Left            =   -74880
               TabIndex        =   45
               Top             =   360
               Width           =   14040
               _ExtentX        =   24765
               _ExtentY        =   12250
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid28 
               Height          =   6945
               Left            =   120
               TabIndex        =   46
               Top             =   360
               Width           =   14040
               _ExtentX        =   24765
               _ExtentY        =   12250
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
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid31 
               Height          =   6945
               Left            =   -74880
               TabIndex        =   47
               Top             =   360
               Width           =   14040
               _ExtentX        =   24765
               _ExtentY        =   12250
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
         End
      End
   End
End
Attribute VB_Name = "FrmRpt_FAGMIS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mStateID, mTotalDays As Variant

Private Sub Form_Activate()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMaximized Then
   SetScrollBars
End If
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

Private Sub MSHFlexGrid1_Click()
If MSHFlexGrid1.RowSel = 2 And MSHFlexGrid1.ColSel = 2 Then
   SSTab1.Tab = 7
   SSTab3.Tab = 1
   Exit Sub
End If
If MSHFlexGrid1.RowSel = 2 And MSHFlexGrid1.ColSel = 3 Then
   SSTab1.Tab = 7
   SSTab3.Tab = 2
   Exit Sub
End If
If MSHFlexGrid1.RowSel = 3 And MSHFlexGrid1.ColSel = 1 Then
   SSTab1.Tab = 8
   SSTab4.Tab = 1
   Exit Sub
End If
If MSHFlexGrid1.RowSel = 3 And MSHFlexGrid1.ColSel = 2 Then
   SSTab1.Tab = 8
   SSTab4.Tab = 2
   Exit Sub
End If
If MSHFlexGrid1.RowSel = 3 And MSHFlexGrid1.ColSel = 3 Then
   SSTab1.Tab = 8
   SSTab4.Tab = 3
   Exit Sub
End If

If MSHFlexGrid1.RowSel = 4 And MSHFlexGrid1.ColSel = 1 Then
   SSTab1.Tab = 9
   SSTab5.Tab = 1
   Exit Sub
End If
If MSHFlexGrid1.RowSel = 4 And MSHFlexGrid1.ColSel = 2 Then
   SSTab1.Tab = 9
   SSTab5.Tab = 2
   Exit Sub
End If
If MSHFlexGrid1.RowSel = 4 And MSHFlexGrid1.ColSel = 3 Then
   SSTab1.Tab = 9
   SSTab5.Tab = 3
   Exit Sub
End If

If MSHFlexGrid1.RowSel = 5 And MSHFlexGrid1.ColSel = 2 Then
   SSTab1.Tab = 10
   SSTab6.Tab = 1
   Exit Sub
End If
If MSHFlexGrid1.RowSel = 5 And MSHFlexGrid1.ColSel = 3 Then
   SSTab1.Tab = 10
   SSTab6.Tab = 2
   Exit Sub
End If

If MSHFlexGrid1.RowSel = 6 And MSHFlexGrid1.ColSel = 1 Then
   SSTab1.Tab = 11
   SSTab2.Tab = 2
   Exit Sub
End If
If MSHFlexGrid1.RowSel = 6 And MSHFlexGrid1.ColSel = 2 Then
   SSTab1.Tab = 11
   SSTab2.Tab = 3
   Exit Sub
End If
If MSHFlexGrid1.RowSel = 6 And MSHFlexGrid1.ColSel = 3 Then
   SSTab1.Tab = 11
   SSTab2.Tab = 4
   Exit Sub
End If

If MSHFlexGrid1.RowSel = 6 And MSHFlexGrid1.ColSel = 4 Then
   SSTab1.Tab = 11
   SSTab2.Tab = 5
   Exit Sub
End If

If MSHFlexGrid1.RowSel = 6 And MSHFlexGrid1.ColSel = 5 Then
   SSTab1.Tab = 11
   SSTab2.Tab = 1
   Exit Sub
End If

If MSHFlexGrid1.RowSel = 7 And MSHFlexGrid1.ColSel = 3 Then
   SSTab1.Tab = 12
   SSTab7.Tab = 1
   Exit Sub
End If

If MSHFlexGrid1.RowSel = 7 And MSHFlexGrid1.ColSel = 4 Then
   SSTab1.Tab = 12
   SSTab7.Tab = 2
   Exit Sub
End If

If MSHFlexGrid1.RowSel = 8 And MSHFlexGrid1.ColSel = 1 Then
   SSTab1.Tab = 13
   SSTab8.Tab = 1
   Exit Sub
End If

If MSHFlexGrid1.RowSel = 8 And MSHFlexGrid1.ColSel = 2 Then
   SSTab1.Tab = 13
   SSTab8.Tab = 2
   Exit Sub
End If

If MSHFlexGrid1.RowSel = 8 And MSHFlexGrid1.ColSel = 3 Then
   SSTab1.Tab = 13
   SSTab8.Tab = 3
   Exit Sub
End If
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
VScroll1.Max = Me.ScaleHeight - Frame4.Height '4
VScroll1.LargeChange = Frame4.ScaleHeight
VScroll1.SmallChange = Frame4.ScaleHeight / 5
HScroll1.Max = Me.ScaleWidth - Frame4.Width '4
HScroll1.LargeChange = Me.ScaleWidth
HScroll1.SmallChange = Me.ScaleWidth / 5
If Frame4.Height >= Me.Height Then VScroll1.Enabled = True ''4
If Frame4.Height < Me.Height Then VScroll1.Enabled = False '4
If Frame4.Width >= Me.Width Then HScroll1.Enabled = True '4
If Frame4.Width < Me.Width Then HScroll1.Enabled = False '4
End Sub

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text

If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If

CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp

End Sub

Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If
mStateID = RsMst_State!StateID
End Sub

Private Sub CmdClear_Click()
TxtFromDate.Value = Date
TxtToDate.Value = Date
CmbStateID = ""
CmdExcel.Enabled = False
Call Grid_Head
Call Grid_Head1
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
Call Grid_Head15
Call Grid_Head16
Call Grid_Head17
Call Grid_Head18
Call Grid_Head19
Call Grid_Head20
Call Grid_Head21
Call Grid_Head22
Call Grid_Head23
Call Grid_Head24
Call Grid_Head25
Call Grid_Head26
Call Grid_Head27
Call Grid_Head28
Call Grid_Head31
Call Grid_Head32
Call Grid_Head33
Call Grid_Head34
Call Grid_Head35
Call Grid_Head36
Check1.Value = 1
Check2.Value = 0
Check3.Value = 0

SSTab1.Tab = 0
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
If Check1.Value = 1 Then
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_FAGMIS.xls")
   Call Xls_Detail_Rpt
End If
If Check2.Value = 1 Then
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_FAGMIS_Invoices.xls")
   Call Xls_Excel_Invoices
End If
If Check3.Value = 1 Then
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_FAGMIS_CashReceipt.xls")
   Call Xls_Excel_CR
End If

End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_FAGMIS.xls")
mRow = oWB.Worksheets.Count
''If 34 - oWB.Worksheets.Count > 0 Then
''   oWB.Worksheets.Add , oWB.Worksheets(mRow), 34 - oWB.Worksheets.Count
''End If

Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "FAG MIS"
oWS.ClearArrows
mRow = 2
oWS.Range("a:a").ColumnWidth = 24.14
oWS.Range("f:f").ColumnWidth = 20.29
oWS.Range("e:e").ColumnWidth = 14.57
oWS.Range("d:d").ColumnWidth = 16.57
oWS.Range("c:c").ColumnWidth = 10.86
oWS.Range("b:b").ColumnWidth = 10.57

tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Process Status from StoreMan"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &H8080FF

For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = mRow + 1 ' + 6
    
    For c = 0 To MSHFlexGrid1.Cols - 1
        If I = 1 Then
           oWS.Cells(mRow, c + 1).Font.Bold = True
           oWS.Cells(mRow, c + 1).RowHeight = 45.75
           oWS.Cells(mRow, c + 1).WrapText = True
        End If
        If c <> 0 Then
           oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
        Else
           oWS.Cells(mRow, c + 1).Font.Bold = True
        End If
        oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
    Next
Next

tmp = "a" & 2 & ":" & "f" & mRow
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
c = 0
mRow = mRow + 2
tmp1 = mRow
tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = "Client Balance Reconciliation Between Navision and StoreMan"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F

mRow = mRow + 1
tmp = "a" & mRow & ":" & "f" & mRow
oWS.Range(tmp).Merge
oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(2, 0)
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14

For I = 1 To MSHFlexGrid3.Rows - 1
    If I <> 2 Then
       mRow = mRow + 1
       oWS.Cells(mRow, 1).Font.Bold = True
       If I = 1 Then
          oWS.Cells(mRow, 1).Interior.Color = &H55FFFF
          oWS.Cells(mRow, 6).Interior.Color = &H55FFFF
          oWS.Cells(mRow, 6).Font.Bold = True
       End If
       If I = MSHFlexGrid3.Rows - 2 Then
          oWS.Cells(mRow, 1).Interior.Color = &H8080BF
          oWS.Cells(mRow, 6).Interior.Color = &H8080BF
       End If
       tmp = "a" & mRow & ":" & "e" & mRow
       oWS.Range(tmp).Merge
       oWS.Cells(mRow, 1) = MSHFlexGrid3.TextMatrix(I, 0)
       oWS.Cells(mRow, 6) = MSHFlexGrid3.TextMatrix(I, 1)
    End If
Next

tmp = "a" & tmp1 & ":" & "f" & mRow
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)

Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "Mapping Pending"
oWS.ClearArrows
mRow = 1
oWS.Range("a:a").ColumnWidth = 21.29
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1
    For c = 0 To MSHFlexGrid2.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
    Next
Next

''Set oWS = oWB.Worksheets("Sheet3")
''oWS.Name = "Outstanding Details"
''oWS.ClearArrows
''mRow = 1
''oWS.Range("a:a").ColumnWidth = 58.71
''oWS.Range("b:b").ColumnWidth = 11.29
''
''For I = 0 To MSHFlexGrid3.Rows - 1
''    mRow = I + 1 ' + 6
''    For c = 0 To MSHFlexGrid3.Cols - 1
''        oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(I, c)
''    Next
''Next

Set oWS = oWB.Worksheets("Sheet3")
oWS.Name = "StoreMan Invoices(With DocNo)"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid4.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid4.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid4.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet4")
oWS.Name = "Invoices(without DocNo)"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid5.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid5.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid5.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet5")
oWS.Name = "Advance CR"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid6.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid6.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid6.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet6")
oWS.Name = "Allocated Advance"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid7.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid7.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid7.TextMatrix(I, c)
    Next
Next

oWB.Save
MsgBox ("Excel File for MIS Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdGetReport_Click()
Dim mTotalDays  As Variant
Call Grid_Head
Call Grid_Head7
Call Grid_Head8
Call Grid_Head9
Call Grid_Head10
Call Grid_Head11
Call Grid_Head12
Call Grid_Head13

SSTab1.Tab = 0

Dim wc As String
Dim tmp1 As String
wc = ""

TxtFromDate.Value = (Year(TxtFromDate.Value) & "-" & Month(TxtFromDate.Value) & "-01")
mTotalDays = Day(DateSerial(Year(TxtToDate.Value), Month(TxtToDate.Value) + 1, 0))
TxtToDate.Value = (Year(TxtToDate.Value) & "-" & Month(TxtToDate.Value) & "-" & mTotalDays)

If TxtFromDate.Value > TxtToDate.Value Then
   MsgBox "From MonthYear should be less than To Month Year"
   TxtFromDate.SetFocus
   Exit Sub
End If
MSHFlexGrid1.Rows = 10
MSHFlexGrid1.TextMatrix(2, 0) = "NNNR Invoice Pending"
MSHFlexGrid1.TextMatrix(3, 0) = "Ad-Hoc Invoice"
MSHFlexGrid1.TextMatrix(4, 0) = "Reservation Invoice"
MSHFlexGrid1.TextMatrix(5, 0) = "DP Invoicing"
MSHFlexGrid1.TextMatrix(6, 0) = "Cash Receipt"
MSHFlexGrid1.TextMatrix(7, 0) = "Advance Cash Receipt"
MSHFlexGrid1.TextMatrix(8, 0) = "other Adjustment And TDS"

Call AddNNNRPending
Call AddAdHocPending
Call AddReservationPending
Call AddDPPending
Call AddCashReceiptPending
Call AddAdvanceCashReceiptPending
Call AddOtherAdjPending

For I = 2 To MSHFlexGrid1.Rows - 1
    MSHFlexGrid1.TextMatrix(0, 1) = Val(MSHFlexGrid1.TextMatrix(0, 1)) + Val(MSHFlexGrid1.TextMatrix(I, 1))
    MSHFlexGrid1.TextMatrix(0, 2) = Val(MSHFlexGrid1.TextMatrix(0, 2)) + Val(MSHFlexGrid1.TextMatrix(I, 2))
    MSHFlexGrid1.TextMatrix(0, 3) = Val(MSHFlexGrid1.TextMatrix(0, 3)) + Val(MSHFlexGrid1.TextMatrix(I, 3))
    MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(I, 4))
    MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(I, 5))
Next

For I = 2 To MSHFlexGrid1.Rows - 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        If LCase(MSHFlexGrid1.TextMatrix(I, c)) = "na" Then
           MSHFlexGrid1.Row = I
           MSHFlexGrid1.Col = c
           MSHFlexGrid1.CellBackColor = vbBlack
        End If
    Next
Next


Call FillPendingNAVCodes
Call GetOutstandingDetails

CmdExcel.Enabled = True
MsgBox "Done!!!"

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access Denied !!!!!!!!!"
   Frame11.Enabled = False
   Exit Sub
End If

Frame11.Enabled = True
Call CmdClear_Click

End Sub

Private Sub Grid_Head3()
MSHFlexGrid4.Clear
MSHFlexGrid4.Rows = 3
MSHFlexGrid4.FixedRows = 2
MSHFlexGrid4.Cols = 14

MSHFlexGrid4.TextMatrix(1, 0) = "TI No"
MSHFlexGrid4.TextMatrix(1, 1) = "TI Date"
MSHFlexGrid4.TextMatrix(1, 2) = "TI Amount"
MSHFlexGrid4.TextMatrix(1, 3) = "TI ST Amount"
MSHFlexGrid4.TextMatrix(1, 4) = "State Name"
MSHFlexGrid4.TextMatrix(1, 5) = "Location Name"
MSHFlexGrid4.TextMatrix(1, 6) = "Post Invoice No"
MSHFlexGrid4.TextMatrix(1, 7) = "Entry No"
MSHFlexGrid4.TextMatrix(1, 8) = "NAV Code"
MSHFlexGrid4.TextMatrix(1, 9) = "Client Name"
MSHFlexGrid4.TextMatrix(1, 10) = "Rec Amount"
MSHFlexGrid4.TextMatrix(1, 11) = "Rec TDS"
MSHFlexGrid4.TextMatrix(1, 12) = "Rec Other Deduction"
MSHFlexGrid4.TextMatrix(1, 13) = "Balance Amount"

MSHFlexGrid4.ColWidth(0) = 1800
MSHFlexGrid4.ColWidth(1) = 1500
MSHFlexGrid4.ColWidth(2) = 1400
MSHFlexGrid4.ColWidth(3) = 1400
MSHFlexGrid4.ColWidth(4) = 2200
MSHFlexGrid4.ColWidth(5) = 2200
MSHFlexGrid4.ColWidth(6) = 1500
MSHFlexGrid4.ColWidth(7) = 1400
MSHFlexGrid4.ColWidth(8) = 1400
MSHFlexGrid4.ColWidth(9) = 2200
MSHFlexGrid4.ColWidth(10) = 1400
MSHFlexGrid4.ColWidth(11) = 1400
MSHFlexGrid4.ColWidth(12) = 1400
MSHFlexGrid4.ColWidth(13) = 1400
MSHFlexGrid4.RowHeight(1) = 600
MSHFlexGrid4.ColAlignment(9) = vbLeftJustify
MSHFlexGrid4.WordWrap = True
End Sub


Private Sub Grid_Head2()
MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.FixedRows = 2
MSHFlexGrid3.Cols = 3

MSHFlexGrid3.TextMatrix(1, 0) = "Particulars"
MSHFlexGrid3.TextMatrix(1, 1) = "Amount"

MSHFlexGrid3.ColWidth(0) = 6500
MSHFlexGrid3.ColWidth(1) = 1500
End Sub



Private Sub Grid_Head1()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2
MSHFlexGrid2.Cols = 3

MSHFlexGrid2.TextMatrix(1, 0) = ""
MSHFlexGrid2.TextMatrix(1, 1) = "Updated"
MSHFlexGrid2.TextMatrix(1, 2) = "Pending"

MSHFlexGrid2.ColWidth(0) = 4000
MSHFlexGrid2.ColWidth(1) = 1500
MSHFlexGrid2.ColWidth(2) = 1500


End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 3
MSHFlexGrid1.Cols = 6
MSHFlexGrid1.FixedRows = 2

MSHFlexGrid1.TextMatrix(1, 0) = ""
MSHFlexGrid1.TextMatrix(1, 1) = "Pending for Approval By FAG"
MSHFlexGrid1.TextMatrix(1, 2) = "Batch Pending By FAG"
MSHFlexGrid1.TextMatrix(1, 3) = "Pending Import of TI / Doc No from Navision (FAG)"
MSHFlexGrid1.TextMatrix(1, 4) = "Pending Invoice Allocation By FAG"
MSHFlexGrid1.TextMatrix(1, 5) = "Pending instrument Not Deposited in Bank (By CMPM/CMG)"

MSHFlexGrid1.ColWidth(0) = 2500
MSHFlexGrid1.ColWidth(1) = 1500
MSHFlexGrid1.ColWidth(2) = 1500
MSHFlexGrid1.ColWidth(3) = 1500
MSHFlexGrid1.ColWidth(4) = 1500
MSHFlexGrid1.ColWidth(5) = 1500

MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(1) = 1400

End Sub
Private Sub AddNNNRPending()

tmp = " Select ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType,"
tmp = tmp & " TPI.InvoiceDate , TPI.InvoiceNo, TPI.ServiceTaxNAVCode, Amount, BatchID"
tmp = tmp & " From Txn_ProInvoice TPI"
tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TPI.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Client MClient On TPI.ClientIDRow = MClient.ClientIDRow"
tmp = tmp & " Where Me.ExchangeTypeID = 1 and isnull(TPI.BatchID,0) = 0 And TPI.BillingMonth>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TPI.BillingMonth<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "' And TPI.Flag = 'C'"
tmp = tmp & " Order by ML.LocationName, MClient.ClientName, ME.ExchangeType, TPI.InvoiceDate, Amount, BatchID"

Call Tmp1Table

'''''''''''''''''''''''--For TI Upload pending
''tmp = " Select * "
''tmp = tmp & " From Txn_ProInvoice TPI"
''tmp = tmp & " Where TPI.ExchangeTypeID = 1 And TPI.BillingMonth>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TPI.BillingMonth<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
''tmp = tmp & " And TPI.InvoiceNo not in (select PIID from txn_TIDetails Where PIFLag='NNNR') And TPI.BatchID > 0"

tmp = "Select ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType,"
tmp = tmp & " TPI.InvoiceDate , TPI.InvoiceNo, TPI.ServiceTaxNAVCode, Amount, BatchID, TI.TIID"
tmp = tmp & " From Txn_ProInvoice TPI"
tmp = tmp & " left Join Mst_Location ML On TPI.LocationID = ML.LocationID"
tmp = tmp & " left Join Mst_ExchangeType ME On TPI.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " left Join Mst_Client MClient On TPI.ClientIDRow = MClient.ClientIDRow"
tmp = tmp & " left Join txn_TIDetails TI On TPI.InvoiceNo = TI.PIID And TI.PIFLag='NNNR'"
tmp = tmp & " Where Me.ExchangeTypeID = 1 And BatchID>0 And TPI.BillingMonth>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TPI.BillingMonth<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "' And TPI.Flag = 'C'"
tmp = tmp & " And TI.TIID is null"
tmp = tmp & " Order by ML.LocationName, MClient.ClientName, ME.ExchangeType, TPI.InvoiceDate, Amount, BatchID"
Call Tmp2Table

MSHFlexGrid1.TextMatrix(2, 1) = "NA"
'MSHFlexGrid1.Row = 2
'MSHFlexGrid1.Col = 1
'MSHFlexGrid1.CellBackColor = vbBlack
MSHFlexGrid1.TextMatrix(2, 2) = TmpRs1.RecordCount
MSHFlexGrid1.TextMatrix(2, 3) = TmpRs2.RecordCount
MSHFlexGrid1.TextMatrix(2, 4) = "NA"
'MSHFlexGrid1.Col = 4
'MSHFlexGrid1.CellBackColor = vbBlack
MSHFlexGrid1.TextMatrix(2, 5) = "NA"
'MSHFlexGrid1.Col = 5
'MSHFlexGrid1.CellBackColor = vbBlack
Call Grid_Head20
If TmpRs1.RecordCount > 0 Then
    I = MSHFlexGrid8.Rows - 1
    MSHFlexGrid8.Rows = MSHFlexGrid8.Rows + TmpRs1.RecordCount
    MSHFlexGrid20.Rows = MSHFlexGrid20.Rows + TmpRs1.RecordCount
    j = 2
    For k = 1 To TmpRs1.RecordCount
        MSHFlexGrid8.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
        MSHFlexGrid8.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
        MSHFlexGrid8.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!ExchangeType), "", TmpRs1!ExchangeType)
        MSHFlexGrid8.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!InvoiceDate), "", Format(TmpRs1!InvoiceDate, "dd-MMM-yyyy"))
        MSHFlexGrid8.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!InvoiceNo), "", TmpRs1!InvoiceNo)
        MSHFlexGrid8.TextMatrix(I, 5) = IIf(IsNull(TmpRs1!ServiceTaxNAVCode), "", TmpRs1!ServiceTaxNAVCode)
        MSHFlexGrid8.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!Amount), "", TmpRs1!Amount)
        MSHFlexGrid8.TextMatrix(0, 6) = Val(MSHFlexGrid8.TextMatrix(0, 6)) + Val(MSHFlexGrid8.TextMatrix(I, 6))
        MSHFlexGrid8.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!BatchID), "", TmpRs1!BatchID)
        
        MSHFlexGrid20.TextMatrix(j, 0) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
        MSHFlexGrid20.TextMatrix(j, 1) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
        MSHFlexGrid20.TextMatrix(j, 2) = IIf(IsNull(TmpRs1!ExchangeType), "", TmpRs1!ExchangeType)
        MSHFlexGrid20.TextMatrix(j, 3) = IIf(IsNull(TmpRs1!InvoiceDate), "", Format(TmpRs1!InvoiceDate, "dd-MMM-yyyy"))
        MSHFlexGrid20.TextMatrix(j, 4) = IIf(IsNull(TmpRs1!InvoiceNo), "", TmpRs1!InvoiceNo)
        MSHFlexGrid20.TextMatrix(j, 5) = IIf(IsNull(TmpRs1!ServiceTaxNAVCode), "", TmpRs1!ServiceTaxNAVCode)
        MSHFlexGrid20.TextMatrix(j, 6) = IIf(IsNull(TmpRs1!Amount), "", TmpRs1!Amount)
        MSHFlexGrid20.TextMatrix(0, 6) = Val(MSHFlexGrid20.TextMatrix(0, 6)) + Val(MSHFlexGrid20.TextMatrix(j, 6))
        MSHFlexGrid20.TextMatrix(j, 7) = IIf(IsNull(TmpRs1!BatchID), "", TmpRs1!BatchID)
        
        I = I + 1
        j = j + 1
        TmpRs1.MoveNext
    Next
End If
Call Grid_Head21
If TmpRs2.RecordCount > 0 Then
    I = MSHFlexGrid8.Rows - 1
    MSHFlexGrid8.Rows = MSHFlexGrid8.Rows + TmpRs2.RecordCount
    MSHFlexGrid21.Rows = MSHFlexGrid21.Rows + TmpRs2.RecordCount
    j = 2
    For k = 1 To TmpRs2.RecordCount
        MSHFlexGrid8.TextMatrix(I, 0) = IIf(IsNull(TmpRs2!LocationName), "", TmpRs2!LocationName)
        MSHFlexGrid8.TextMatrix(I, 1) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
        MSHFlexGrid8.TextMatrix(I, 2) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
        MSHFlexGrid8.TextMatrix(I, 3) = IIf(IsNull(TmpRs2!InvoiceDate), "", Format(TmpRs2!InvoiceDate, "dd-MMM-yyyy"))
        MSHFlexGrid8.TextMatrix(I, 4) = IIf(IsNull(TmpRs2!InvoiceNo), "", TmpRs2!InvoiceNo)
        MSHFlexGrid8.TextMatrix(I, 5) = IIf(IsNull(TmpRs2!ServiceTaxNAVCode), "", TmpRs2!ServiceTaxNAVCode)
        MSHFlexGrid8.TextMatrix(I, 6) = IIf(IsNull(TmpRs2!Amount), "", TmpRs2!Amount)
        MSHFlexGrid8.TextMatrix(0, 6) = Val(MSHFlexGrid8.TextMatrix(0, 6)) + Val(MSHFlexGrid8.TextMatrix(I, 6))
        MSHFlexGrid8.TextMatrix(I, 7) = IIf(IsNull(TmpRs2!BatchID), "", TmpRs2!BatchID)
        
        MSHFlexGrid21.TextMatrix(j, 0) = IIf(IsNull(TmpRs2!LocationName), "", TmpRs2!LocationName)
        MSHFlexGrid21.TextMatrix(j, 1) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
        MSHFlexGrid21.TextMatrix(j, 2) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
        MSHFlexGrid21.TextMatrix(j, 3) = IIf(IsNull(TmpRs2!InvoiceDate), "", Format(TmpRs2!InvoiceDate, "dd-MMM-yyyy"))
        MSHFlexGrid21.TextMatrix(j, 4) = IIf(IsNull(TmpRs2!InvoiceNo), "", TmpRs2!InvoiceNo)
        MSHFlexGrid21.TextMatrix(j, 5) = IIf(IsNull(TmpRs2!ServiceTaxNAVCode), "", TmpRs2!ServiceTaxNAVCode)
        MSHFlexGrid21.TextMatrix(j, 6) = IIf(IsNull(TmpRs2!Amount), "", TmpRs2!Amount)
        MSHFlexGrid21.TextMatrix(0, 6) = Val(MSHFlexGrid21.TextMatrix(0, 6)) + Val(MSHFlexGrid21.TextMatrix(j, 6))
        MSHFlexGrid21.TextMatrix(j, 7) = IIf(IsNull(TmpRs2!BatchID), "", TmpRs2!BatchID)
        
        I = I + 1
        j = j + 1
        TmpRs2.MoveNext
    Next
End If
End Sub

Private Sub AddDPPending()

'--For BOEs Processed By FAG

tmp = " Select ML.LocationID, ML.LocationName, TPI.DPNavCode,"
tmp = tmp & " TPI.MonthYear , TPI.DPInvoiceID,TPI.Amt,TPI.BatchID,"
tmp = tmp & " (Select top 1 DepositoryName from Mst_Depository Where TPI.DPNAVCode = NAVCode) DepositoryName"
tmp = tmp & " From Mst_DPInvoiceUpload TPI"
tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
tmp = tmp & " Where isnull(TPI.BatchID,0) = 0 And TPI.MonthYear>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TPI.MonthYear<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Order by ML.LocationName, DepositoryName, TPI.MonthYear, TPI.DPNAVCode, TPI.Amt, TPI.BatchID "
Call Tmp1Table

'''''''''''''''''''''''--For TI Uploaded
tmp = " Select ML.LocationID, ML.LocationName, TPI.DPNavCode, "
tmp = tmp & " TPI.MonthYear , TPI.DPInvoiceID, TPI.Amt,TPI.BatchID,"
tmp = tmp & " (Select top 1 DepositoryName from Mst_Depository Where TPI.DPNAVCode = NAVCode) DepositoryName"
tmp = tmp & " From Mst_DPInvoiceUpload TPI "
tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
tmp = tmp & " Where isnull(TPI.BatchID,0) > 0 And TPI.MonthYear>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TPI.MonthYear<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " And TPI.DPInvoiceID not in (select PIID from txn_TIDetails Where PIFLag='DP')"
tmp = tmp & " Order by ML.LocationName, DepositoryName, TPI.MonthYear, TPI.DPNAVCode, TPI.Amt, TPI.BatchID "
Call Tmp2Table

MSHFlexGrid1.TextMatrix(5, 1) = "NA"
MSHFlexGrid1.TextMatrix(5, 2) = TmpRs1.RecordCount
MSHFlexGrid1.TextMatrix(5, 3) = TmpRs2.RecordCount
MSHFlexGrid1.TextMatrix(5, 4) = "NA"
MSHFlexGrid1.TextMatrix(5, 5) = "NA"


'MSHFlexGrid1.Row = 5
'MSHFlexGrid1.Col = 1
'MSHFlexGrid1.CellBackColor = vbBlack
'MSHFlexGrid1.Col = 4
'MSHFlexGrid1.CellBackColor = vbBlack
'MSHFlexGrid1.Col = 5
'MSHFlexGrid1.CellBackColor = vbBlack
Call Grid_Head28
If TmpRs1.RecordCount > 0 Then
    I = MSHFlexGrid11.Rows - 1
    MSHFlexGrid11.Rows = MSHFlexGrid11.Rows + TmpRs1.RecordCount
    MSHFlexGrid28.Rows = MSHFlexGrid28.Rows + TmpRs1.RecordCount
    j = 2
    For k = 1 To TmpRs1.RecordCount
        MSHFlexGrid11.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
        MSHFlexGrid11.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!DepositoryName), "", TmpRs1!DepositoryName)
        MSHFlexGrid11.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!DPNAVCode), "", TmpRs1!DPNAVCode)
        MSHFlexGrid11.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!MonthYear), "", Format(TmpRs1!MonthYear, "MMM-yyyy"))
        MSHFlexGrid11.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!DPInvoiceID), "", TmpRs1!DPInvoiceID)
        MSHFlexGrid11.TextMatrix(I, 5) = IIf(IsNull(TmpRs1!Amt), "", TmpRs1!Amt)
        MSHFlexGrid11.TextMatrix(0, 5) = Val(MSHFlexGrid11.TextMatrix(0, 5)) + Val(MSHFlexGrid11.TextMatrix(I, 5))
        MSHFlexGrid11.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!BatchID), "", TmpRs1!BatchID)
        
        MSHFlexGrid28.TextMatrix(j, 0) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
        MSHFlexGrid28.TextMatrix(j, 1) = IIf(IsNull(TmpRs1!DepositoryName), "", TmpRs1!DepositoryName)
        MSHFlexGrid28.TextMatrix(j, 2) = IIf(IsNull(TmpRs1!DPNAVCode), "", TmpRs1!DPNAVCode)
        MSHFlexGrid28.TextMatrix(j, 3) = IIf(IsNull(TmpRs1!MonthYear), "", Format(TmpRs1!MonthYear, "MMM-yyyy"))
        MSHFlexGrid28.TextMatrix(j, 4) = IIf(IsNull(TmpRs1!DPInvoiceID), "", TmpRs1!DPInvoiceID)
        MSHFlexGrid28.TextMatrix(j, 5) = IIf(IsNull(TmpRs1!Amt), "", TmpRs1!Amt)
        MSHFlexGrid28.TextMatrix(0, 5) = Val(MSHFlexGrid28.TextMatrix(0, 5)) + Val(MSHFlexGrid28.TextMatrix(j, 5))
        MSHFlexGrid28.TextMatrix(j, 6) = IIf(IsNull(TmpRs1!BatchID), "", TmpRs1!BatchID)
        
        I = I + 1
        j = j + 1
        TmpRs1.MoveNext
    Next
End If
Call Grid_Head31
If TmpRs2.RecordCount > 0 Then
    I = MSHFlexGrid11.Rows - 1
    MSHFlexGrid11.Rows = MSHFlexGrid11.Rows + TmpRs2.RecordCount
    MSHFlexGrid31.Rows = MSHFlexGrid31.Rows + TmpRs2.RecordCount
    j = 2
    For k = 1 To TmpRs2.RecordCount
        MSHFlexGrid11.TextMatrix(I, 0) = IIf(IsNull(TmpRs2!LocationName), "", TmpRs2!LocationName)
        MSHFlexGrid11.TextMatrix(I, 1) = IIf(IsNull(TmpRs2!DepositoryName), "", TmpRs2!DepositoryName)
        MSHFlexGrid11.TextMatrix(I, 2) = IIf(IsNull(TmpRs2!DPNAVCode), "", TmpRs2!DPNAVCode)
        MSHFlexGrid11.TextMatrix(I, 3) = IIf(IsNull(TmpRs2!MonthYear), "", Format(TmpRs2!MonthYear, "MMM-yyyy"))
        MSHFlexGrid11.TextMatrix(I, 4) = IIf(IsNull(TmpRs2!DPInvoiceID), "", TmpRs2!DPInvoiceID)
        MSHFlexGrid11.TextMatrix(I, 5) = IIf(IsNull(TmpRs2!Amt), "", TmpRs2!Amt)
        MSHFlexGrid11.TextMatrix(0, 5) = Val(MSHFlexGrid11.TextMatrix(0, 5)) + Val(MSHFlexGrid11.TextMatrix(I, 5))
        MSHFlexGrid11.TextMatrix(I, 6) = IIf(IsNull(TmpRs2!BatchID), "", TmpRs2!BatchID)
        
        MSHFlexGrid31.TextMatrix(j, 0) = IIf(IsNull(TmpRs2!LocationName), "", TmpRs2!LocationName)
        MSHFlexGrid31.TextMatrix(j, 1) = IIf(IsNull(TmpRs2!DepositoryName), "", TmpRs2!DepositoryName)
        MSHFlexGrid31.TextMatrix(j, 2) = IIf(IsNull(TmpRs2!DPNAVCode), "", TmpRs2!DPNAVCode)
        MSHFlexGrid31.TextMatrix(j, 3) = IIf(IsNull(TmpRs2!MonthYear), "", Format(TmpRs2!MonthYear, "MMM-yyyy"))
        MSHFlexGrid31.TextMatrix(j, 4) = IIf(IsNull(TmpRs2!DPInvoiceID), "", TmpRs2!DPInvoiceID)
        MSHFlexGrid31.TextMatrix(j, 5) = IIf(IsNull(TmpRs2!Amt), "", TmpRs2!Amt)
        MSHFlexGrid31.TextMatrix(0, 5) = Val(MSHFlexGrid31.TextMatrix(0, 5)) + Val(MSHFlexGrid31.TextMatrix(j, 5))
        MSHFlexGrid31.TextMatrix(j, 6) = IIf(IsNull(TmpRs2!BatchID), "", TmpRs2!BatchID)
        
        I = I + 1
        j = j + 1
        TmpRs2.MoveNext
    Next
End If
End Sub

Private Sub AddAdHocPending()

'--For Pending By FAG

tmp = "Select ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType, TPI.InvoiceDate , TPI.PINo, ST.NavCode, Amount, BatchID"
tmp = tmp & " From Txn_TempInvoice TPI"
tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TPI.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Client MClient On TPI.ClientIDRow = MClient.ClientIDRow"
tmp = tmp & " Left Join Mst_ServiceTax ST On TPI.ServiceTaxID = ST.ServiceTaxID"
tmp = tmp & " Where  TPI.Flag='P'And TPI.InvoiceDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TPI.InvoiceDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Order by ML.LocationName, MClient.ClientName, ME.ExchangeType, TPI.InvoiceDate, Amount, BatchID"

Call Tmp1Table

'--For Batch Pending By FAG

tmp = "Select ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType, TPI.InvoiceDate , TPI.PINo, ST.NavCode, Amount, BatchID"
tmp = tmp & " From Txn_TempInvoice TPI"
tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TPI.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Client MClient On TPI.ClientIDRow = MClient.ClientIDRow"
tmp = tmp & " Left Join Mst_ServiceTax ST On TPI.ServiceTaxID = ST.ServiceTaxID"
tmp = tmp & " Where  TPI.Flag='A' And TPI.BatchID=0 And TPI.InvoiceDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TPI.InvoiceDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Order by ML.LocationName, MClient.ClientName, ME.ExchangeType, TPI.InvoiceDate, Amount, BatchID"

Call Tmp2Table

'''''''''''''''''''''''--For TI Uploaded

tmp = "Select ML.LocationID, ML.LocationName, MClient.ClientID, MClient.ClientName, Me.ExchangeTypeID, Me.ExchangeType,"
tmp = tmp & " TPI.InvoiceDate , TPI.PINo, ST.NAVCode, Amount, BatchID"
tmp = tmp & " From Txn_TempInvoice TPI"
tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
tmp = tmp & " Inner Join Mst_ExchangeType ME On TPI.ExchangeTypeID = ME.ExchangeTypeID"
tmp = tmp & " Inner Join Mst_Client MClient On TPI.ClientIDRow = MClient.ClientIDRow"
tmp = tmp & " Left Join Mst_ServiceTax ST On TPI.ServiceTaxID = ST.ServiceTaxID"
tmp = tmp & " Where TPI.Flag='A' And TPI.BatchID>0 And TPI.InvoiceDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TPI.InvoiceDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " And TPI.PINo not in (select PIID from txn_TIDetails Where PIFLag='ADHOC')"
tmp = tmp & " Order by ML.LocationName, MClient.ClientName, ME.ExchangeType, TPI.InvoiceDate, Amount, BatchID"
Call Tmp3Table

MSHFlexGrid1.TextMatrix(3, 1) = TmpRs1.RecordCount
MSHFlexGrid1.TextMatrix(3, 2) = TmpRs2.RecordCount
MSHFlexGrid1.TextMatrix(3, 3) = TmpRs3.RecordCount
MSHFlexGrid1.TextMatrix(3, 4) = "NA"
MSHFlexGrid1.TextMatrix(3, 5) = "NA"

'MSHFlexGrid1.Row = 3
'MSHFlexGrid1.Col = 4
'MSHFlexGrid1.CellBackColor = vbBlack
'MSHFlexGrid1.Col = 5
'MSHFlexGrid1.CellBackColor = vbBlack
Call Grid_Head22
If TmpRs1.RecordCount > 0 Then
    I = MSHFlexGrid9.Rows - 1
    MSHFlexGrid9.Rows = MSHFlexGrid9.Rows + TmpRs1.RecordCount
    MSHFlexGrid22.Rows = MSHFlexGrid22.Rows + TmpRs1.RecordCount
    j = 2
    For k = 1 To TmpRs1.RecordCount
        MSHFlexGrid9.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
        MSHFlexGrid9.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
        MSHFlexGrid9.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!ExchangeType), "", TmpRs1!ExchangeType)
        MSHFlexGrid9.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!InvoiceDate), "", Format(TmpRs1!InvoiceDate, "dd-MMM-yyyy"))
        MSHFlexGrid9.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!PINo), "", TmpRs1!PINo)
        MSHFlexGrid9.TextMatrix(I, 5) = IIf(IsNull(TmpRs1!NAVCode), "", TmpRs1!NAVCode)
        MSHFlexGrid9.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!Amount), "", TmpRs1!Amount)
        MSHFlexGrid9.TextMatrix(0, 6) = Val(MSHFlexGrid9.TextMatrix(0, 6)) + Val(MSHFlexGrid9.TextMatrix(I, 6))
        MSHFlexGrid9.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!BatchID), "", TmpRs1!BatchID)
        
        MSHFlexGrid22.TextMatrix(j, 0) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
        MSHFlexGrid22.TextMatrix(j, 1) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
        MSHFlexGrid22.TextMatrix(j, 2) = IIf(IsNull(TmpRs1!ExchangeType), "", TmpRs1!ExchangeType)
        MSHFlexGrid22.TextMatrix(j, 3) = IIf(IsNull(TmpRs1!InvoiceDate), "", Format(TmpRs1!InvoiceDate, "dd-MMM-yyyy"))
        MSHFlexGrid22.TextMatrix(j, 4) = IIf(IsNull(TmpRs1!PINo), "", TmpRs1!PINo)
        MSHFlexGrid22.TextMatrix(j, 5) = IIf(IsNull(TmpRs1!NAVCode), "", TmpRs1!NAVCode)
        MSHFlexGrid22.TextMatrix(j, 6) = IIf(IsNull(TmpRs1!Amount), "", TmpRs1!Amount)
        MSHFlexGrid22.TextMatrix(0, 6) = Val(MSHFlexGrid22.TextMatrix(0, 6)) + Val(MSHFlexGrid22.TextMatrix(j, 6))
        MSHFlexGrid22.TextMatrix(j, 7) = IIf(IsNull(TmpRs1!BatchID), "", TmpRs1!BatchID)
        
        I = I + 1
        j = j + 1
        TmpRs1.MoveNext
    Next
End If
Call Grid_Head23
If TmpRs2.RecordCount > 0 Then
    I = MSHFlexGrid9.Rows - 1
    MSHFlexGrid9.Rows = MSHFlexGrid9.Rows + TmpRs2.RecordCount
    MSHFlexGrid23.Rows = MSHFlexGrid23.Rows + TmpRs2.RecordCount
    j = 2
    For k = 1 To TmpRs2.RecordCount
        MSHFlexGrid9.TextMatrix(I, 0) = IIf(IsNull(TmpRs2!LocationName), "", TmpRs2!LocationName)
        MSHFlexGrid9.TextMatrix(I, 1) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
        MSHFlexGrid9.TextMatrix(I, 2) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
        MSHFlexGrid9.TextMatrix(I, 3) = IIf(IsNull(TmpRs2!InvoiceDate), "", Format(TmpRs2!InvoiceDate, "dd-MMM-yyyy"))
        MSHFlexGrid9.TextMatrix(I, 4) = IIf(IsNull(TmpRs2!PINo), "", TmpRs2!PINo)
        MSHFlexGrid9.TextMatrix(I, 5) = IIf(IsNull(TmpRs2!NAVCode), "", TmpRs2!NAVCode)
        MSHFlexGrid9.TextMatrix(I, 6) = IIf(IsNull(TmpRs2!Amount), "", TmpRs2!Amount)
        MSHFlexGrid9.TextMatrix(0, 6) = Val(MSHFlexGrid9.TextMatrix(0, 6)) + Val(MSHFlexGrid9.TextMatrix(I, 6))
        MSHFlexGrid9.TextMatrix(I, 7) = IIf(IsNull(TmpRs2!BatchID), "", TmpRs2!BatchID)
        
        MSHFlexGrid23.TextMatrix(j, 0) = IIf(IsNull(TmpRs2!LocationName), "", TmpRs2!LocationName)
        MSHFlexGrid23.TextMatrix(j, 1) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
        MSHFlexGrid23.TextMatrix(j, 2) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
        MSHFlexGrid23.TextMatrix(j, 3) = IIf(IsNull(TmpRs2!InvoiceDate), "", Format(TmpRs2!InvoiceDate, "dd-MMM-yyyy"))
        MSHFlexGrid23.TextMatrix(j, 4) = IIf(IsNull(TmpRs2!PINo), "", TmpRs2!PINo)
        MSHFlexGrid23.TextMatrix(j, 5) = IIf(IsNull(TmpRs2!NAVCode), "", TmpRs2!NAVCode)
        MSHFlexGrid23.TextMatrix(j, 6) = IIf(IsNull(TmpRs2!Amount), "", TmpRs2!Amount)
        MSHFlexGrid23.TextMatrix(0, 6) = Val(MSHFlexGrid23.TextMatrix(0, 6)) + Val(MSHFlexGrid23.TextMatrix(j, 6))
        MSHFlexGrid23.TextMatrix(j, 7) = IIf(IsNull(TmpRs2!BatchID), "", TmpRs2!BatchID)
        
        I = I + 1
        j = j + 1
        TmpRs2.MoveNext
    Next
End If
Call Grid_Head24
If TmpRs3.RecordCount > 0 Then
    I = MSHFlexGrid9.Rows - 1
    MSHFlexGrid9.Rows = MSHFlexGrid9.Rows + TmpRs3.RecordCount
    MSHFlexGrid24.Rows = MSHFlexGrid24.Rows + TmpRs3.RecordCount
    j = 2
    For k = 1 To TmpRs3.RecordCount
        MSHFlexGrid9.TextMatrix(I, 0) = IIf(IsNull(TmpRs3!LocationName), "", TmpRs3!LocationName)
        MSHFlexGrid9.TextMatrix(I, 1) = IIf(IsNull(TmpRs3!ClientName), "", TmpRs3!ClientName)
        MSHFlexGrid9.TextMatrix(I, 2) = IIf(IsNull(TmpRs3!ExchangeType), "", TmpRs3!ExchangeType)
        MSHFlexGrid9.TextMatrix(I, 3) = IIf(IsNull(TmpRs3!InvoiceDate), "", Format(TmpRs3!InvoiceDate, "dd-MMM-yyyy"))
        MSHFlexGrid9.TextMatrix(I, 4) = IIf(IsNull(TmpRs3!PINo), "", TmpRs3!PINo)
        MSHFlexGrid9.TextMatrix(I, 5) = IIf(IsNull(TmpRs3!NAVCode), "", TmpRs3!NAVCode)
        MSHFlexGrid9.TextMatrix(I, 6) = IIf(IsNull(TmpRs3!Amount), "", TmpRs3!Amount)
        MSHFlexGrid9.TextMatrix(0, 6) = Val(MSHFlexGrid9.TextMatrix(0, 6)) + Val(MSHFlexGrid9.TextMatrix(I, 6))
        MSHFlexGrid9.TextMatrix(I, 7) = IIf(IsNull(TmpRs3!BatchID), "", TmpRs3!BatchID)
        
        MSHFlexGrid24.TextMatrix(j, 0) = IIf(IsNull(TmpRs3!LocationName), "", TmpRs3!LocationName)
        MSHFlexGrid24.TextMatrix(j, 1) = IIf(IsNull(TmpRs3!ClientName), "", TmpRs3!ClientName)
        MSHFlexGrid24.TextMatrix(j, 2) = IIf(IsNull(TmpRs3!ExchangeType), "", TmpRs3!ExchangeType)
        MSHFlexGrid24.TextMatrix(j, 3) = IIf(IsNull(TmpRs3!InvoiceDate), "", Format(TmpRs3!InvoiceDate, "dd-MMM-yyyy"))
        MSHFlexGrid24.TextMatrix(j, 4) = IIf(IsNull(TmpRs3!PINo), "", TmpRs3!PINo)
        MSHFlexGrid24.TextMatrix(j, 5) = IIf(IsNull(TmpRs3!NAVCode), "", TmpRs3!NAVCode)
        MSHFlexGrid24.TextMatrix(j, 6) = IIf(IsNull(TmpRs3!Amount), "", TmpRs3!Amount)
        MSHFlexGrid24.TextMatrix(0, 6) = Val(MSHFlexGrid24.TextMatrix(0, 6)) + Val(MSHFlexGrid24.TextMatrix(j, 6))
        MSHFlexGrid24.TextMatrix(j, 7) = IIf(IsNull(TmpRs3!BatchID), "", TmpRs3!BatchID)
        
        I = I + 1
        j = j + 1
        TmpRs3.MoveNext
    Next
End If

End Sub

Private Sub AddReservationPending()

'--For Pending By FAG

tmp = "Select ML.LocationName,TRI.InvoiceNo,TRI.ActualAmount,TRI.BillingDate,MSC.ContractName,"
tmp = tmp & " TRI.BillingAmount , TRI.YearMonth, TRI.Remarks, MC.Commodity, MCL.ClientNAme, TRI.Flag,TRI.ServiceTaxNAVCode"
tmp = tmp & " from Txn_ReservationInvoice TRI"
tmp = tmp & " inner join Mst_StorageContract MSC on MSC.ContractID=TRI.ContractID"
tmp = tmp & " Left join MSt_Commodity MC on MC.CommodityID=MSC.CommodityID"
tmp = tmp & " inner join Mst_Client MCL on MCL.ClientIDRow=MSC.ClientIDRow"
tmp = tmp & " inner join mst_Location ML on ML.LocationID=TRI.LocationID"
tmp = tmp & " Where  TRI.Flag='P'And TRI.YearMonth>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TRI.YearMonth<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Order by ML.LocationName, MCl.ClientName, MC.Commodity, TRI.BillingDate, TRI.BillingAmount"
Call Tmp1Table

'--For Batch Pending By FAG

tmp = "Select ML.LocationName,TRI.InvoiceNo,TRI.ActualAmount,TRI.BillingDate,MSC.ContractName,"
tmp = tmp & " TRI.BillingAmount , TRI.YearMonth, TRI.Remarks, MC.Commodity, MCL.ClientNAme, TRI.Flag,TRI.ServiceTaxNAVCode"
tmp = tmp & " from Txn_ReservationInvoice TRI"
tmp = tmp & " inner join Mst_StorageContract MSC on MSC.ContractID=TRI.ContractID"
tmp = tmp & " Left join MSt_Commodity MC on MC.CommodityID=MSC.CommodityID"
tmp = tmp & " inner join Mst_Client MCL on MCL.ClientIDRow=MSC.ClientIDRow"
tmp = tmp & " inner join mst_Location ML on ML.LocationID=TRI.LocationID"
tmp = tmp & " Where  TRI.Flag='A' And TRI.BatchID=0 And TRI.YearMonth>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TRI.YearMonth<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Order by ML.LocationName, MCl.ClientName, MC.Commodity, TRI.BillingDate, TRI.BillingAmount"
Call Tmp2Table

'''''''''''''''''''''''--For TI Uploaded

'tmp = "Select ML.LocationID, ML.LocationName, MCL.ClientID, MCL.ClientName,"
'tmp = tmp & " TPI.YearMonth , TPI.InvoiceNo,BillingAmount, TPI.BatchID"
'tmp = tmp & " From Txn_ReservationInvoice TPI"
'tmp = tmp & " Inner Join Mst_Location ML On TPI.LocationID = ML.LocationID"
'tmp = tmp & " inner join Mst_StorageContract MSC on MSC.ContractID=TPI.ContractID"
'tmp = tmp & " Left join MSt_Commodity MC on MC.CommodityID=MSC.CommodityID"
'tmp = tmp & " inner join Mst_Client MCL on MCL.ClientIDRow=MSC.ClientIDRow"
tmp = "Select ML.LocationName,TRI.InvoiceNo,TRI.ActualAmount,TRI.BillingDate,MSC.ContractName,"
tmp = tmp & " TRI.BillingAmount , TRI.YearMonth, TRI.Remarks, MC.Commodity, MCL.ClientNAme, TRI.Flag,TRI.ServiceTaxNAVCode"
tmp = tmp & " from Txn_ReservationInvoice TRI"
tmp = tmp & " inner join Mst_StorageContract MSC on MSC.ContractID=TRI.ContractID"
tmp = tmp & " Left join MSt_Commodity MC on MC.CommodityID=MSC.CommodityID"
tmp = tmp & " inner join Mst_Client MCL on MCL.ClientIDRow=MSC.ClientIDRow"
tmp = tmp & " inner join mst_Location ML on ML.LocationID=TRI.LocationID"
tmp = tmp & " Where TRI.Flag='A' And TRI.BatchID>0 And  TRI.YearMonth>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TRI.YearMonth<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " And TRI.InvoiceNo not in (select PIID from txn_TIDetails Where PIFLag='Reservation')"
tmp = tmp & " Order by ML.LocationName, MCl.ClientName, MC.Commodity, TRI.BillingDate, TRI.BillingAmount"
Call Tmp3Table

MSHFlexGrid1.TextMatrix(4, 1) = TmpRs1.RecordCount
MSHFlexGrid1.TextMatrix(4, 2) = TmpRs2.RecordCount
MSHFlexGrid1.TextMatrix(4, 3) = TmpRs3.RecordCount
MSHFlexGrid1.TextMatrix(4, 4) = "NA"
MSHFlexGrid1.TextMatrix(4, 5) = "NA"

'MSHFlexGrid1.Row = 4
'MSHFlexGrid1.Col = 4
'MSHFlexGrid1.CellBackColor = vbBlack
'MSHFlexGrid1.Col = 5
'MSHFlexGrid1.CellBackColor = vbBlack

Call Grid_Head25
If TmpRs1.RecordCount > 0 Then
    I = MSHFlexGrid10.Rows - 1
    MSHFlexGrid10.Rows = MSHFlexGrid10.Rows + TmpRs1.RecordCount
    MSHFlexGrid25.Rows = MSHFlexGrid25.Rows + TmpRs1.RecordCount
    j = 2
    For k = 1 To TmpRs1.RecordCount
        MSHFlexGrid10.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
        MSHFlexGrid10.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
        MSHFlexGrid10.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!Commodity), "", TmpRs1!Commodity)
        MSHFlexGrid10.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!BillingDate), "", Format(TmpRs1!BillingDate, "dd-MMM-yyyy"))
        MSHFlexGrid10.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!InvoiceNo), "", TmpRs1!InvoiceNo)
        MSHFlexGrid10.TextMatrix(I, 5) = IIf(IsNull(TmpRs1!ServiceTaxNAVCode), "", TmpRs1!ServiceTaxNAVCode)
        MSHFlexGrid10.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!BillingAmount), "", TmpRs1!BillingAmount)
        MSHFlexGrid10.TextMatrix(0, 6) = Val(MSHFlexGrid10.TextMatrix(0, 6)) + Val(MSHFlexGrid10.TextMatrix(I, 6))
        MSHFlexGrid10.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!ActualAmount), "", TmpRs1!ActualAmount)
        MSHFlexGrid10.TextMatrix(0, 7) = Val(MSHFlexGrid10.TextMatrix(0, 7)) + Val(MSHFlexGrid10.TextMatrix(I, 7))
        MSHFlexGrid10.TextMatrix(I, 8) = IIf(IsNull(TmpRs1!ContractName), "", TmpRs1!ContractName)
        MSHFlexGrid10.TextMatrix(I, 9) = IIf(IsNull(TmpRs1!Remarks), "", TmpRs1!Remarks)
        MSHFlexGrid10.TextMatrix(I, 10) = IIf(IsNull(TmpRs1!Flag), "", TmpRs1!Flag)
        If LCase(MSHFlexGrid10.TextMatrix(I, 10)) = LCase("P") Then
            MSHFlexGrid10.TextMatrix(I, 10) = "Pending"
        ElseIf LCase(MSHFlexGrid10.TextMatrix(I, 10)) = LCase("A") Then
            MSHFlexGrid10.TextMatrix(I, 10) = "Approved"
        End If
        MSHFlexGrid10.TextMatrix(I, 11) = IIf(IsNull(TmpRs1!YearMonth), "", Format(TmpRs1!YearMonth, "MMM-yyyy"))
        
        
        MSHFlexGrid25.TextMatrix(j, 0) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
        MSHFlexGrid25.TextMatrix(j, 1) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
        MSHFlexGrid25.TextMatrix(j, 2) = IIf(IsNull(TmpRs1!Commodity), "", TmpRs1!Commodity)
        MSHFlexGrid25.TextMatrix(j, 3) = IIf(IsNull(TmpRs1!BillingDate), "", Format(TmpRs1!BillingDate, "dd-MMM-yyyy"))
        MSHFlexGrid25.TextMatrix(j, 4) = IIf(IsNull(TmpRs1!InvoiceNo), "", TmpRs1!InvoiceNo)
        MSHFlexGrid25.TextMatrix(j, 5) = IIf(IsNull(TmpRs1!ServiceTaxNAVCode), "", TmpRs1!ServiceTaxNAVCode)
        MSHFlexGrid25.TextMatrix(j, 6) = IIf(IsNull(TmpRs1!BillingAmount), "", TmpRs1!BillingAmount)
        MSHFlexGrid25.TextMatrix(0, 6) = Val(MSHFlexGrid25.TextMatrix(0, 6)) + Val(MSHFlexGrid25.TextMatrix(j, 6))
        MSHFlexGrid25.TextMatrix(j, 7) = IIf(IsNull(TmpRs1!ActualAmount), "", TmpRs1!ActualAmount)
        MSHFlexGrid25.TextMatrix(0, 7) = Val(MSHFlexGrid25.TextMatrix(0, 7)) + Val(MSHFlexGrid25.TextMatrix(j, 7))
        MSHFlexGrid25.TextMatrix(j, 8) = IIf(IsNull(TmpRs1!ContractName), "", TmpRs1!ContractName)
        MSHFlexGrid25.TextMatrix(j, 9) = IIf(IsNull(TmpRs1!Remarks), "", TmpRs1!Remarks)
        MSHFlexGrid25.TextMatrix(j, 10) = IIf(IsNull(TmpRs1!Flag), "", TmpRs1!Flag)
        If LCase(MSHFlexGrid25.TextMatrix(j, 10)) = LCase("P") Then
            MSHFlexGrid25.TextMatrix(j, 10) = "Pending"
        ElseIf LCase(MSHFlexGrid25.TextMatrix(j, 10)) = LCase("A") Then
            MSHFlexGrid25.TextMatrix(j, 10) = "Approved"
        End If
        MSHFlexGrid25.TextMatrix(j, 11) = IIf(IsNull(TmpRs1!YearMonth), "", Format(TmpRs1!YearMonth, "MMM-yyyy"))
        
        I = I + 1
        j = j + 1
        TmpRs1.MoveNext
    Next
End If

Call Grid_Head26
If TmpRs2.RecordCount > 0 Then
    I = MSHFlexGrid10.Rows - 1
    MSHFlexGrid10.Rows = MSHFlexGrid10.Rows + TmpRs2.RecordCount
    MSHFlexGrid26.Rows = MSHFlexGrid26.Rows + TmpRs2.RecordCount
    j = 2
    For k = 1 To TmpRs2.RecordCount
        MSHFlexGrid10.TextMatrix(I, 0) = IIf(IsNull(TmpRs2!LocationName), "", TmpRs2!LocationName)
        MSHFlexGrid10.TextMatrix(I, 1) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
        MSHFlexGrid10.TextMatrix(I, 2) = IIf(IsNull(TmpRs2!Commodity), "", TmpRs2!Commodity)
        MSHFlexGrid10.TextMatrix(I, 3) = IIf(IsNull(TmpRs2!BillingDate), "", Format(TmpRs2!BillingDate, "dd-MMM-yyyy"))
        MSHFlexGrid10.TextMatrix(I, 4) = IIf(IsNull(TmpRs2!InvoiceNo), "", TmpRs2!InvoiceNo)
        MSHFlexGrid10.TextMatrix(I, 5) = IIf(IsNull(TmpRs2!NAVCode), "", TmpRs2!NAVCode)
        MSHFlexGrid10.TextMatrix(I, 6) = IIf(IsNull(TmpRs2!BillingAmount), "", TmpRs2!BillingAmount)
        MSHFlexGrid10.TextMatrix(0, 6) = Val(MSHFlexGrid10.TextMatrix(0, 6)) + Val(MSHFlexGrid10.TextMatrix(I, 6))
        MSHFlexGrid10.TextMatrix(I, 7) = IIf(IsNull(TmpRs2!ActualAmount), "", TmpRs2!ActualAmount)
        MSHFlexGrid10.TextMatrix(0, 7) = Val(MSHFlexGrid10.TextMatrix(0, 7)) + Val(MSHFlexGrid10.TextMatrix(I, 7))
        MSHFlexGrid10.TextMatrix(I, 8) = IIf(IsNull(TmpRs2!ContractName), "", TmpRs2!ContractName)
        MSHFlexGrid10.TextMatrix(I, 9) = IIf(IsNull(TmpRs2!Remarks), "", TmpRs2!Remarks)
        MSHFlexGrid10.TextMatrix(I, 10) = IIf(IsNull(TmpRs2!Flag), "", TmpRs2!Flag)
        If LCase(MSHFlexGrid10.TextMatrix(I, 10)) = LCase("P") Then
            MSHFlexGrid10.TextMatrix(I, 10) = "Pending"
        ElseIf LCase(MSHFlexGrid10.TextMatrix(I, 10)) = LCase("A") Then
            MSHFlexGrid10.TextMatrix(I, 10) = "Approved"
        End If
        MSHFlexGrid10.TextMatrix(I, 11) = IIf(IsNull(TmpRs2!YearMonth), "", Format(TmpRs2!YearMonth, "MMM-yyyy"))
        
        MSHFlexGrid26.TextMatrix(j, 0) = IIf(IsNull(TmpRs2!LocationName), "", TmpRs2!LocationName)
        MSHFlexGrid26.TextMatrix(j, 1) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
        MSHFlexGrid26.TextMatrix(j, 2) = IIf(IsNull(TmpRs2!Commodity), "", TmpRs2!Commodity)
        MSHFlexGrid26.TextMatrix(j, 3) = IIf(IsNull(TmpRs2!BillingDate), "", Format(TmpRs2!BillingDate, "dd-MMM-yyyy"))
        MSHFlexGrid26.TextMatrix(j, 4) = IIf(IsNull(TmpRs2!InvoiceNo), "", TmpRs2!InvoiceNo)
        MSHFlexGrid26.TextMatrix(j, 5) = IIf(IsNull(TmpRs2!NAVCode), "", TmpRs2!NAVCode)
        MSHFlexGrid26.TextMatrix(j, 6) = IIf(IsNull(TmpRs2!BillingAmount), "", TmpRs2!BillingAmount)
        MSHFlexGrid26.TextMatrix(0, 6) = Val(MSHFlexGrid26.TextMatrix(0, 6)) + Val(MSHFlexGrid26.TextMatrix(j, 6))
        MSHFlexGrid26.TextMatrix(j, 7) = IIf(IsNull(TmpRs2!ActualAmount), "", TmpRs2!ActualAmount)
        MSHFlexGrid26.TextMatrix(0, 7) = Val(MSHFlexGrid26.TextMatrix(0, 7)) + Val(MSHFlexGrid26.TextMatrix(j, 7))
        MSHFlexGrid26.TextMatrix(j, 8) = IIf(IsNull(TmpRs2!ContractName), "", TmpRs2!ContractName)
        MSHFlexGrid26.TextMatrix(j, 9) = IIf(IsNull(TmpRs2!Remarks), "", TmpRs2!Remarks)
        MSHFlexGrid26.TextMatrix(j, 10) = IIf(IsNull(TmpRs2!Flag), "", TmpRs2!Flag)
        If LCase(MSHFlexGrid26.TextMatrix(j, 10)) = LCase("P") Then
            MSHFlexGrid26.TextMatrix(j, 10) = "Pending"
        ElseIf LCase(MSHFlexGrid26.TextMatrix(j, 10)) = LCase("A") Then
            MSHFlexGrid26.TextMatrix(j, 10) = "Approved"
        End If
        MSHFlexGrid26.TextMatrix(j, 11) = IIf(IsNull(TmpRs2!YearMonth), "", Format(TmpRs2!YearMonth, "MMM-yyyy"))
        
        I = I + 1
        j = j + 1
        TmpRs2.MoveNext
    Next
End If
Call Grid_Head27
If TmpRs3.RecordCount > 0 Then
    I = MSHFlexGrid10.Rows - 1
    MSHFlexGrid10.Rows = MSHFlexGrid10.Rows + TmpRs3.RecordCount
    MSHFlexGrid27.Rows = MSHFlexGrid27.Rows + TmpRs3.RecordCount
    j = 2
    For k = 1 To TmpRs3.RecordCount
        MSHFlexGrid10.TextMatrix(I, 0) = IIf(IsNull(TmpRs3!LocationName), "", TmpRs3!LocationName)
        MSHFlexGrid10.TextMatrix(I, 1) = IIf(IsNull(TmpRs3!ClientName), "", TmpRs3!ClientName)
        MSHFlexGrid10.TextMatrix(I, 2) = IIf(IsNull(TmpRs3!Commodity), "", TmpRs3!Commodity)
        MSHFlexGrid10.TextMatrix(I, 3) = IIf(IsNull(TmpRs3!BillingDate), "", Format(TmpRs3!BillingDate, "dd-MMM-yyyy"))
        MSHFlexGrid10.TextMatrix(I, 4) = IIf(IsNull(TmpRs3!InvoiceNo), "", TmpRs3!InvoiceNo)
        MSHFlexGrid10.TextMatrix(I, 5) = IIf(IsNull(TmpRs3!ServiceTaxNAVCode), "", TmpRs3!ServiceTaxNAVCode)
        MSHFlexGrid10.TextMatrix(I, 6) = IIf(IsNull(TmpRs3!BillingAmount), "", TmpRs3!BillingAmount)
        MSHFlexGrid10.TextMatrix(0, 6) = Val(MSHFlexGrid10.TextMatrix(0, 6)) + Val(MSHFlexGrid10.TextMatrix(I, 6))
        MSHFlexGrid10.TextMatrix(I, 7) = IIf(IsNull(TmpRs3!ActualAmount), "", TmpRs3!ActualAmount)
        MSHFlexGrid10.TextMatrix(0, 7) = Val(MSHFlexGrid10.TextMatrix(0, 7)) + Val(MSHFlexGrid10.TextMatrix(I, 7))
        MSHFlexGrid10.TextMatrix(I, 8) = IIf(IsNull(TmpRs3!ContractName), "", TmpRs3!ContractName)
        MSHFlexGrid10.TextMatrix(I, 9) = IIf(IsNull(TmpRs3!Remarks), "", TmpRs3!Remarks)
        MSHFlexGrid10.TextMatrix(I, 10) = IIf(IsNull(TmpRs3!Flag), "", TmpRs3!Flag)
        If LCase(MSHFlexGrid10.TextMatrix(I, 10)) = LCase("P") Then
            MSHFlexGrid10.TextMatrix(I, 10) = "Pending"
        ElseIf LCase(MSHFlexGrid10.TextMatrix(I, 10)) = LCase("A") Then
            MSHFlexGrid10.TextMatrix(I, 10) = "Approved"
        End If
        MSHFlexGrid10.TextMatrix(I, 11) = IIf(IsNull(TmpRs3!YearMonth), "", Format(TmpRs3!YearMonth, "MMM-yyyy"))
        
        MSHFlexGrid27.TextMatrix(j, 0) = IIf(IsNull(TmpRs3!LocationName), "", TmpRs3!LocationName)
        MSHFlexGrid27.TextMatrix(j, 1) = IIf(IsNull(TmpRs3!ClientName), "", TmpRs3!ClientName)
        MSHFlexGrid27.TextMatrix(j, 2) = IIf(IsNull(TmpRs3!Commodity), "", TmpRs3!Commodity)
        MSHFlexGrid27.TextMatrix(j, 3) = IIf(IsNull(TmpRs3!BillingDate), "", Format(TmpRs3!BillingDate, "dd-MMM-yyyy"))
        MSHFlexGrid27.TextMatrix(j, 4) = IIf(IsNull(TmpRs3!InvoiceNo), "", TmpRs3!InvoiceNo)
        MSHFlexGrid27.TextMatrix(j, 5) = IIf(IsNull(TmpRs3!ServiceTaxNAVCode), "", TmpRs3!ServiceTaxNAVCode)
        MSHFlexGrid27.TextMatrix(j, 6) = IIf(IsNull(TmpRs3!BillingAmount), "", TmpRs3!BillingAmount)
        MSHFlexGrid27.TextMatrix(0, 6) = Val(MSHFlexGrid27.TextMatrix(0, 6)) + Val(MSHFlexGrid27.TextMatrix(j, 6))
        MSHFlexGrid27.TextMatrix(j, 7) = IIf(IsNull(TmpRs3!ActualAmount), "", TmpRs3!ActualAmount)
        MSHFlexGrid27.TextMatrix(0, 7) = Val(MSHFlexGrid27.TextMatrix(0, 7)) + Val(MSHFlexGrid27.TextMatrix(j, 7))
        MSHFlexGrid27.TextMatrix(j, 8) = IIf(IsNull(TmpRs3!ContractName), "", TmpRs3!ContractName)
        MSHFlexGrid27.TextMatrix(j, 9) = IIf(IsNull(TmpRs3!Remarks), "", TmpRs3!Remarks)
        MSHFlexGrid27.TextMatrix(j, 10) = IIf(IsNull(TmpRs3!Flag), "", TmpRs3!Flag)
        If LCase(MSHFlexGrid27.TextMatrix(j, 10)) = LCase("P") Then
            MSHFlexGrid27.TextMatrix(j, 10) = "Pending"
        ElseIf LCase(MSHFlexGrid27.TextMatrix(j, 10)) = LCase("A") Then
            MSHFlexGrid27.TextMatrix(j, 10) = "Approved"
        End If
        MSHFlexGrid27.TextMatrix(j, 11) = IIf(IsNull(TmpRs3!YearMonth), "", Format(TmpRs3!YearMonth, "MMM-yyyy"))
        
        I = I + 1
        j = j + 1
        TmpRs3.MoveNext
    Next
End If

End Sub


Private Sub AddCashReceiptPending()

'--For Pending By FAG

tmp = "Select TCR.CRID,TCR.SM_Prefix,CRDate,LocationName,InstrumentType,DrawnOn,BranchName,InstrumentNo,InstrumentDate,"
tmp = tmp & " InstrumentAmount , ReceivedFrom, Remark, AdvanceFlag, BatchID,DepositDate,FAGPostingDate"
tmp = tmp & " from Txn_NewCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " Where TCR.ReturnFlag<>'Y' And isnull(TCR.RejectFlag,'')<>'Y' And DepositDate Is Not Null And FAGPostingDate Is Null And TCR.CRDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.CRDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Order by LocationName, CRID, CRDate, BranchName, InstrumentAmount, BatchID "

Call Tmp1Table

'--For Batch Pending By FAG

tmp = "Select TCR.CRID,TCR.SM_Prefix,CRDate,LocationName,InstrumentType,DrawnOn,BranchName,InstrumentNo,InstrumentDate,"
tmp = tmp & " InstrumentAmount , ReceivedFrom, Remark, AdvanceFlag, BatchID,DepositDate,FAGPostingDate"
tmp = tmp & " from Txn_NewCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " Where TCR.ReturnFlag<>'Y' And isnull(TCR.RejectFlag,'')<>'Y' And FAGPostingDate Is not Null And isnull(BatchID,0)=0 And TCR.CRDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.CRDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Order by LocationName, CRID, CRDate, BranchName, InstrumentAmount, BatchID "

Call Tmp2Table

'''''''''''''''''''''''--For TI Uploaded pending
tmp = "Select distinct TCR.CRID,TCR.SM_Prefix,CRDate,LocationName,InstrumentType,DrawnOn,BranchName,InstrumentNo,InstrumentDate,"
tmp = tmp & " InstrumentAmount , ReceivedFrom, TCR.Remark, AdvanceFlag, TCR.batchID,DepositDate,FAGPostingDate"
tmp = tmp & " from Txn_NewCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " inner join Txn_NewCashReceiptTIDetail TCC on TCC.CRID=TCR.CRID And TCC.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y' And isnull(TCR.RejectFlag,'')<>'Y' And batchID > 0 And TCC.EntryNo Is Null And TCR.CRDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.CRDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "' And isnull(TCC.CRAAID,0)=0"
tmp = tmp & " Union"
tmp = tmp & " Select distinct TCR.CRID,TCR.SM_Prefix,CRDate,LocationName,InstrumentType,DrawnOn,BranchName,InstrumentNo,InstrumentDate,"
tmp = tmp & " InstrumentAmount , ReceivedFrom, TCR.Remark, AdvanceFlag, TCR.batchID,DepositDate,FAGPostingDate"
tmp = tmp & " from Txn_NewCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " inner join Txn_NewCashReceiptClientDetail TCC on TCC.CRID=TCR.CRID And TCC.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y' And isnull(TCR.RejectFlag,'')<>'Y' And batchID > 0 And TCC.EntryNo Is Null And TCC.Amount > 0 And TCR.CRDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.CRDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"

Call Tmp3Table

''------for Allocation Pending
'tmp = "Select distinct TCR.CRID,TCR.SM_Prefix,CRDate,LocationName,InstrumentType,DrawnOn,BranchName,InstrumentNo,InstrumentDate,"
'tmp = tmp & " InstrumentAmount , ReceivedFrom, TCR.Remark, AdvanceFlag, TCR.BatchID"
'tmp = tmp & " from Txn_NewCashReceipt TCR"
'tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
'tmp = tmp & " Inner join Txn_NewCashReceiptClientDetail TCRD on TCR.CRID=TCRD.CRID and TCR.SM_Prefix=TCRD.SM_Prefix"
'tmp = tmp & " Where TCR.AdvanceFlag = 1 And TCR.ReturnFlag<>'Y'"
'tmp = tmp & " and (TCR.InstrumentAmount-(Select isnull(sum(RecAmount),0) from Txn_NewCashReceiptTIDetail Where CRID= TCR.CRID and SM_Prefix=TCR.SM_Prefix))>0 And TCR.CRDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.CRDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"

tmp = "Select distinct TCR.CRID,TCR.SM_Prefix,CRDate,LocationName,InstrumentType,DrawnOn,BranchName,InstrumentNo,InstrumentDate,"
tmp = tmp & " InstrumentAmount , ReceivedFrom, TCR.Remark, AdvanceFlag, TCR.BatchID,DepositDate,FAGPostingDate"
tmp = tmp & " from Txn_NewCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " Inner join Txn_NewCashReceiptTIDetail TCRD on TCR.CRID=TCRD.CRID and TCR.SM_Prefix=TCRD.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y'"
tmp = tmp & " And isnull(TCR.RejectFlag,'')<>'Y' And TCR.batchID > 0 And TCRD.EntryNo Is not Null And isnull(InvAllocFlag,0)=0 And TCR.CRDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.CRDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Order by LocationName, TCR.CRID, CRDate, BranchName, InstrumentAmount, BatchID "

Call Tmp4Table

'--for not deposited
tmp = "Select TCR.CRID,TCR.SM_Prefix,CRDate,LocationName,InstrumentType,DrawnOn,BranchName,InstrumentNo,InstrumentDate,"
tmp = tmp & " InstrumentAmount , ReceivedFrom, Remark, AdvanceFlag, BatchID,DepositDate,FAGPostingDate"
tmp = tmp & " from Txn_NewCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " Where TCR.ReturnFlag <> 'Y' And isnull(TCR.RejectFlag,'')<>'Y' And DepositDate Is Null And TCR.CRDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.CRDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Order by LocationName, TCR.CRID, CRDate, BranchName, InstrumentAmount, BatchID "

Call Tmp5Table

MSHFlexGrid1.TextMatrix(6, 1) = TmpRs1.RecordCount
MSHFlexGrid1.TextMatrix(6, 2) = TmpRs2.RecordCount
MSHFlexGrid1.TextMatrix(6, 3) = TmpRs3.RecordCount
MSHFlexGrid1.TextMatrix(6, 4) = TmpRs4.RecordCount
MSHFlexGrid1.TextMatrix(6, 5) = TmpRs5.RecordCount
Call Grid_Head17
If TmpRs1.RecordCount > 0 Then
    I = MSHFlexGrid12.Rows - 1
    MSHFlexGrid12.Rows = MSHFlexGrid12.Rows + TmpRs1.RecordCount
    MSHFlexGrid17.Rows = MSHFlexGrid17.Rows + TmpRs1.RecordCount
    j = 2
    For k = 1 To TmpRs1.RecordCount
        MSHFlexGrid12.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!CRID), "", TmpRs1!CRID)
        MSHFlexGrid12.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!CRDate), "", Format(TmpRs1!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
        MSHFlexGrid12.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!InstrumentType), "", TmpRs1!InstrumentType)
        MSHFlexGrid12.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!DrawnOn), "", TmpRs1!DrawnOn)
        MSHFlexGrid12.TextMatrix(I, 5) = IIf(IsNull(TmpRs1!BranchName), "", TmpRs1!BranchName)
        MSHFlexGrid12.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!InstrumentNo), "", TmpRs1!InstrumentNo)
        MSHFlexGrid12.TextMatrix(I, 7) = IIf(IsNull(TmpRs1!InstrumentDate), "", Format(TmpRs1!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 8) = IIf(IsNull(TmpRs1!InstrumentAmount), "", TmpRs1!InstrumentAmount)
        MSHFlexGrid12.TextMatrix(0, 8) = Val(MSHFlexGrid12.TextMatrix(0, 8)) + Val(MSHFlexGrid12.TextMatrix(I, 8))
        MSHFlexGrid12.TextMatrix(I, 9) = IIf(IsNull(TmpRs1!ReceivedFrom), "", TmpRs1!ReceivedFrom)
        MSHFlexGrid12.TextMatrix(I, 10) = IIf(IsNull(TmpRs1!Remark), "", TmpRs1!Remark)
        MSHFlexGrid12.TextMatrix(I, 11) = IIf(IsNull(TmpRs1!BatchID), "", TmpRs1!BatchID)
        MSHFlexGrid12.TextMatrix(I, 12) = IIf(IsNull(TmpRs1!DepositDate), "", Format(TmpRs1!DepositDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 13) = IIf(IsNull(TmpRs1!FAGPostingDate), "", Format(TmpRs1!FAGPostingDate, "dd-MMM-yyyy"))
        
        MSHFlexGrid17.TextMatrix(j, 0) = IIf(IsNull(TmpRs1!CRID), "", TmpRs1!CRID)
        MSHFlexGrid17.TextMatrix(j, 1) = IIf(IsNull(TmpRs1!CRDate), "", Format(TmpRs1!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid17.TextMatrix(j, 2) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
        MSHFlexGrid17.TextMatrix(j, 3) = IIf(IsNull(TmpRs1!InstrumentType), "", TmpRs1!InstrumentType)
        MSHFlexGrid17.TextMatrix(j, 4) = IIf(IsNull(TmpRs1!DrawnOn), "", TmpRs1!DrawnOn)
        MSHFlexGrid17.TextMatrix(j, 5) = IIf(IsNull(TmpRs1!BranchName), "", TmpRs1!BranchName)
        MSHFlexGrid17.TextMatrix(j, 6) = IIf(IsNull(TmpRs1!InstrumentNo), "", TmpRs1!InstrumentNo)
        MSHFlexGrid17.TextMatrix(j, 7) = IIf(IsNull(TmpRs1!InstrumentDate), "", Format(TmpRs1!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid17.TextMatrix(j, 8) = IIf(IsNull(TmpRs1!InstrumentAmount), "", TmpRs1!InstrumentAmount)
        MSHFlexGrid17.TextMatrix(0, 8) = Val(MSHFlexGrid17.TextMatrix(0, 8)) + Val(MSHFlexGrid17.TextMatrix(j, 8))
        MSHFlexGrid17.TextMatrix(j, 9) = IIf(IsNull(TmpRs1!ReceivedFrom), "", TmpRs1!ReceivedFrom)
        MSHFlexGrid17.TextMatrix(j, 10) = IIf(IsNull(TmpRs1!Remark), "", TmpRs1!Remark)
        MSHFlexGrid17.TextMatrix(j, 11) = IIf(IsNull(TmpRs1!BatchID), "", TmpRs1!BatchID)
        MSHFlexGrid17.TextMatrix(j, 12) = IIf(IsNull(TmpRs1!DepositDate), "", Format(TmpRs1!DepositDate, "dd-MMM-yyyy"))
        MSHFlexGrid17.TextMatrix(j, 13) = IIf(IsNull(TmpRs1!FAGPostingDate), "", Format(TmpRs1!FAGPostingDate, "dd-MMM-yyyy"))
        
        I = I + 1
        j = j + 1
        TmpRs1.MoveNext
    Next
End If
Call Grid_Head16
If TmpRs2.RecordCount > 0 Then
    I = MSHFlexGrid12.Rows - 1
    MSHFlexGrid12.Rows = MSHFlexGrid12.Rows + TmpRs2.RecordCount
    MSHFlexGrid16.Rows = MSHFlexGrid16.Rows + TmpRs2.RecordCount
    j = 2
    For k = 1 To TmpRs2.RecordCount
        MSHFlexGrid12.TextMatrix(I, 0) = IIf(IsNull(TmpRs2!CRID), "", TmpRs2!CRID)
        MSHFlexGrid12.TextMatrix(I, 1) = IIf(IsNull(TmpRs2!CRDate), "", Format(TmpRs2!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 2) = IIf(IsNull(TmpRs2!LocationName), "", TmpRs2!LocationName)
        MSHFlexGrid12.TextMatrix(I, 3) = IIf(IsNull(TmpRs2!InstrumentType), "", TmpRs2!InstrumentType)
        MSHFlexGrid12.TextMatrix(I, 4) = IIf(IsNull(TmpRs2!DrawnOn), "", TmpRs2!DrawnOn)
        MSHFlexGrid12.TextMatrix(I, 5) = IIf(IsNull(TmpRs2!BranchName), "", TmpRs2!BranchName)
        MSHFlexGrid12.TextMatrix(I, 6) = IIf(IsNull(TmpRs2!InstrumentNo), "", TmpRs2!InstrumentNo)
        MSHFlexGrid12.TextMatrix(I, 7) = IIf(IsNull(TmpRs2!InstrumentDate), "", Format(TmpRs2!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 8) = IIf(IsNull(TmpRs2!InstrumentAmount), "", TmpRs2!InstrumentAmount)
        MSHFlexGrid12.TextMatrix(0, 8) = Val(MSHFlexGrid12.TextMatrix(0, 8)) + Val(MSHFlexGrid12.TextMatrix(I, 8))
        MSHFlexGrid12.TextMatrix(I, 9) = IIf(IsNull(TmpRs2!ReceivedFrom), "", TmpRs2!ReceivedFrom)
        MSHFlexGrid12.TextMatrix(I, 10) = IIf(IsNull(TmpRs2!Remark), "", TmpRs2!Remark)
        MSHFlexGrid12.TextMatrix(I, 11) = IIf(IsNull(TmpRs2!BatchID), "", TmpRs2!BatchID)
        MSHFlexGrid12.TextMatrix(I, 12) = IIf(IsNull(TmpRs2!DepositDate), "", Format(TmpRs2!DepositDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 13) = IIf(IsNull(TmpRs2!FAGPostingDate), "", Format(TmpRs2!FAGPostingDate, "dd-MMM-yyyy"))
        
        MSHFlexGrid16.TextMatrix(j, 0) = IIf(IsNull(TmpRs2!CRID), "", TmpRs2!CRID)
        MSHFlexGrid16.TextMatrix(j, 1) = IIf(IsNull(TmpRs2!CRDate), "", Format(TmpRs2!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid16.TextMatrix(j, 2) = IIf(IsNull(TmpRs2!LocationName), "", TmpRs2!LocationName)
        MSHFlexGrid16.TextMatrix(j, 3) = IIf(IsNull(TmpRs2!InstrumentType), "", TmpRs2!InstrumentType)
        MSHFlexGrid16.TextMatrix(j, 4) = IIf(IsNull(TmpRs2!DrawnOn), "", TmpRs2!DrawnOn)
        MSHFlexGrid16.TextMatrix(j, 5) = IIf(IsNull(TmpRs2!BranchName), "", TmpRs2!BranchName)
        MSHFlexGrid16.TextMatrix(j, 6) = IIf(IsNull(TmpRs2!InstrumentNo), "", TmpRs2!InstrumentNo)
        MSHFlexGrid16.TextMatrix(j, 7) = IIf(IsNull(TmpRs2!InstrumentDate), "", Format(TmpRs2!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid16.TextMatrix(j, 8) = IIf(IsNull(TmpRs2!InstrumentAmount), "", TmpRs2!InstrumentAmount)
        MSHFlexGrid16.TextMatrix(0, 8) = Val(MSHFlexGrid16.TextMatrix(0, 8)) + Val(MSHFlexGrid16.TextMatrix(j, 8))
        MSHFlexGrid16.TextMatrix(j, 9) = IIf(IsNull(TmpRs2!ReceivedFrom), "", TmpRs2!ReceivedFrom)
        MSHFlexGrid16.TextMatrix(j, 10) = IIf(IsNull(TmpRs2!Remark), "", TmpRs2!Remark)
        MSHFlexGrid16.TextMatrix(j, 11) = IIf(IsNull(TmpRs2!BatchID), "", TmpRs2!BatchID)
        MSHFlexGrid16.TextMatrix(j, 12) = IIf(IsNull(TmpRs2!DepositDate), "", Format(TmpRs2!DepositDate, "dd-MMM-yyyy"))
        MSHFlexGrid16.TextMatrix(j, 13) = IIf(IsNull(TmpRs2!FAGPostingDate), "", Format(TmpRs2!FAGPostingDate, "dd-MMM-yyyy"))
        
        I = I + 1
        j = j + 1
        TmpRs2.MoveNext
    Next
End If
Call Grid_Head18
If TmpRs3.RecordCount > 0 Then
    I = MSHFlexGrid12.Rows - 1
    MSHFlexGrid12.Rows = MSHFlexGrid12.Rows + TmpRs3.RecordCount
    MSHFlexGrid18.Rows = MSHFlexGrid18.Rows + TmpRs3.RecordCount
    j = 2
    For k = 1 To TmpRs3.RecordCount
        MSHFlexGrid12.TextMatrix(I, 0) = IIf(IsNull(TmpRs3!CRID), "", TmpRs3!CRID)
        MSHFlexGrid12.TextMatrix(I, 1) = IIf(IsNull(TmpRs3!CRDate), "", Format(TmpRs3!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 2) = IIf(IsNull(TmpRs3!LocationName), "", TmpRs3!LocationName)
        MSHFlexGrid12.TextMatrix(I, 3) = IIf(IsNull(TmpRs3!InstrumentType), "", TmpRs3!InstrumentType)
        MSHFlexGrid12.TextMatrix(I, 4) = IIf(IsNull(TmpRs3!DrawnOn), "", TmpRs3!DrawnOn)
        MSHFlexGrid12.TextMatrix(I, 5) = IIf(IsNull(TmpRs3!BranchName), "", TmpRs3!BranchName)
        MSHFlexGrid12.TextMatrix(I, 6) = IIf(IsNull(TmpRs3!InstrumentNo), "", TmpRs3!InstrumentNo)
        MSHFlexGrid12.TextMatrix(I, 7) = IIf(IsNull(TmpRs3!InstrumentDate), "", Format(TmpRs3!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 8) = IIf(IsNull(TmpRs3!InstrumentAmount), "", TmpRs3!InstrumentAmount)
        MSHFlexGrid12.TextMatrix(0, 8) = Val(MSHFlexGrid12.TextMatrix(0, 8)) + Val(MSHFlexGrid12.TextMatrix(I, 8))
        MSHFlexGrid12.TextMatrix(I, 9) = IIf(IsNull(TmpRs3!ReceivedFrom), "", TmpRs3!ReceivedFrom)
        MSHFlexGrid12.TextMatrix(I, 10) = IIf(IsNull(TmpRs3!Remark), "", TmpRs3!Remark)
        MSHFlexGrid12.TextMatrix(I, 11) = IIf(IsNull(TmpRs3!BatchID), "", TmpRs3!BatchID)
        MSHFlexGrid12.TextMatrix(I, 12) = IIf(IsNull(TmpRs3!DepositDate), "", Format(TmpRs3!DepositDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 13) = IIf(IsNull(TmpRs3!FAGPostingDate), "", Format(TmpRs3!FAGPostingDate, "dd-MMM-yyyy"))
        
        MSHFlexGrid18.TextMatrix(j, 0) = IIf(IsNull(TmpRs3!CRID), "", TmpRs3!CRID)
        MSHFlexGrid18.TextMatrix(j, 1) = IIf(IsNull(TmpRs3!CRDate), "", Format(TmpRs3!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid18.TextMatrix(j, 2) = IIf(IsNull(TmpRs3!LocationName), "", TmpRs3!LocationName)
        MSHFlexGrid18.TextMatrix(j, 3) = IIf(IsNull(TmpRs3!InstrumentType), "", TmpRs3!InstrumentType)
        MSHFlexGrid18.TextMatrix(j, 4) = IIf(IsNull(TmpRs3!DrawnOn), "", TmpRs3!DrawnOn)
        MSHFlexGrid18.TextMatrix(j, 5) = IIf(IsNull(TmpRs3!BranchName), "", TmpRs3!BranchName)
        MSHFlexGrid18.TextMatrix(j, 6) = IIf(IsNull(TmpRs3!InstrumentNo), "", TmpRs3!InstrumentNo)
        MSHFlexGrid18.TextMatrix(j, 7) = IIf(IsNull(TmpRs3!InstrumentDate), "", Format(TmpRs3!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid18.TextMatrix(j, 8) = IIf(IsNull(TmpRs3!InstrumentAmount), "", TmpRs3!InstrumentAmount)
        MSHFlexGrid18.TextMatrix(0, 8) = Val(MSHFlexGrid18.TextMatrix(0, 8)) + Val(MSHFlexGrid18.TextMatrix(j, 8))
        MSHFlexGrid18.TextMatrix(j, 9) = IIf(IsNull(TmpRs3!ReceivedFrom), "", TmpRs3!ReceivedFrom)
        MSHFlexGrid18.TextMatrix(j, 10) = IIf(IsNull(TmpRs3!Remark), "", TmpRs3!Remark)
        MSHFlexGrid18.TextMatrix(j, 11) = IIf(IsNull(TmpRs3!BatchID), "", TmpRs3!BatchID)
        MSHFlexGrid18.TextMatrix(j, 12) = IIf(IsNull(TmpRs3!DepositDate), "", Format(TmpRs3!DepositDate, "dd-MMM-yyyy"))
        MSHFlexGrid18.TextMatrix(j, 13) = IIf(IsNull(TmpRs3!FAGPostingDate), "", Format(TmpRs3!FAGPostingDate, "dd-MMM-yyyy"))
        
        I = I + 1
        j = j + 1
        TmpRs3.MoveNext
    Next
End If
Call Grid_Head19
If TmpRs4.RecordCount > 0 Then
    I = MSHFlexGrid12.Rows - 1
    MSHFlexGrid12.Rows = MSHFlexGrid12.Rows + TmpRs4.RecordCount
    MSHFlexGrid19.Rows = MSHFlexGrid19.Rows + TmpRs4.RecordCount
    j = 2
    For k = 1 To TmpRs4.RecordCount
        MSHFlexGrid12.TextMatrix(I, 0) = IIf(IsNull(TmpRs4!CRID), "", TmpRs4!CRID)
        MSHFlexGrid12.TextMatrix(I, 1) = IIf(IsNull(TmpRs4!CRDate), "", Format(TmpRs4!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 2) = IIf(IsNull(TmpRs4!LocationName), "", TmpRs4!LocationName)
        MSHFlexGrid12.TextMatrix(I, 3) = IIf(IsNull(TmpRs4!InstrumentType), "", TmpRs4!InstrumentType)
        MSHFlexGrid12.TextMatrix(I, 4) = IIf(IsNull(TmpRs4!DrawnOn), "", TmpRs4!DrawnOn)
        MSHFlexGrid12.TextMatrix(I, 5) = IIf(IsNull(TmpRs4!BranchName), "", TmpRs4!BranchName)
        MSHFlexGrid12.TextMatrix(I, 6) = IIf(IsNull(TmpRs4!InstrumentNo), "", TmpRs4!InstrumentNo)
        MSHFlexGrid12.TextMatrix(I, 7) = IIf(IsNull(TmpRs4!InstrumentDate), "", Format(TmpRs4!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 8) = IIf(IsNull(TmpRs4!InstrumentAmount), "", TmpRs4!InstrumentAmount)
        MSHFlexGrid12.TextMatrix(0, 8) = Val(MSHFlexGrid12.TextMatrix(0, 8)) + Val(MSHFlexGrid12.TextMatrix(I, 8))
        MSHFlexGrid12.TextMatrix(I, 9) = IIf(IsNull(TmpRs4!ReceivedFrom), "", TmpRs4!ReceivedFrom)
        MSHFlexGrid12.TextMatrix(I, 10) = IIf(IsNull(TmpRs4!Remark), "", TmpRs4!Remark)
        MSHFlexGrid12.TextMatrix(I, 11) = IIf(IsNull(TmpRs4!BatchID), "", TmpRs4!BatchID)
        MSHFlexGrid12.TextMatrix(I, 12) = IIf(IsNull(TmpRs4!DepositDate), "", Format(TmpRs4!DepositDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 13) = IIf(IsNull(TmpRs4!FAGPostingDate), "", Format(TmpRs4!FAGPostingDate, "dd-MMM-yyyy"))
        
        MSHFlexGrid19.TextMatrix(j, 0) = IIf(IsNull(TmpRs4!CRID), "", TmpRs4!CRID)
        MSHFlexGrid19.TextMatrix(j, 1) = IIf(IsNull(TmpRs4!CRDate), "", Format(TmpRs4!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid19.TextMatrix(j, 2) = IIf(IsNull(TmpRs4!LocationName), "", TmpRs4!LocationName)
        MSHFlexGrid19.TextMatrix(j, 3) = IIf(IsNull(TmpRs4!InstrumentType), "", TmpRs4!InstrumentType)
        MSHFlexGrid19.TextMatrix(j, 4) = IIf(IsNull(TmpRs4!DrawnOn), "", TmpRs4!DrawnOn)
        MSHFlexGrid19.TextMatrix(j, 5) = IIf(IsNull(TmpRs4!BranchName), "", TmpRs4!BranchName)
        MSHFlexGrid19.TextMatrix(j, 6) = IIf(IsNull(TmpRs4!InstrumentNo), "", TmpRs4!InstrumentNo)
        MSHFlexGrid19.TextMatrix(j, 7) = IIf(IsNull(TmpRs4!InstrumentDate), "", Format(TmpRs4!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid19.TextMatrix(j, 8) = IIf(IsNull(TmpRs4!InstrumentAmount), "", TmpRs4!InstrumentAmount)
        MSHFlexGrid19.TextMatrix(0, 8) = Val(MSHFlexGrid19.TextMatrix(0, 8)) + Val(MSHFlexGrid19.TextMatrix(j, 8))
        MSHFlexGrid19.TextMatrix(j, 9) = IIf(IsNull(TmpRs4!ReceivedFrom), "", TmpRs4!ReceivedFrom)
        MSHFlexGrid19.TextMatrix(j, 10) = IIf(IsNull(TmpRs4!Remark), "", TmpRs4!Remark)
        MSHFlexGrid19.TextMatrix(j, 11) = IIf(IsNull(TmpRs4!BatchID), "", TmpRs4!BatchID)
        MSHFlexGrid19.TextMatrix(j, 12) = IIf(IsNull(TmpRs4!DepositDate), "", Format(TmpRs4!DepositDate, "dd-MMM-yyyy"))
        MSHFlexGrid19.TextMatrix(j, 13) = IIf(IsNull(TmpRs4!FAGPostingDate), "", Format(TmpRs4!FAGPostingDate, "dd-MMM-yyyy"))
        
        I = I + 1
        j = j + 1
        TmpRs4.MoveNext
    Next
End If

Call Grid_Head15
If TmpRs5.RecordCount > 0 Then
    I = MSHFlexGrid12.Rows - 1
    MSHFlexGrid12.Rows = MSHFlexGrid12.Rows + TmpRs5.RecordCount
    MSHFlexGrid15.Rows = MSHFlexGrid15.Rows + TmpRs5.RecordCount
    j = 2
    For k = 1 To TmpRs5.RecordCount
        MSHFlexGrid12.TextMatrix(I, 0) = IIf(IsNull(TmpRs5!CRID), "", TmpRs5!CRID)
        MSHFlexGrid12.TextMatrix(I, 1) = IIf(IsNull(TmpRs5!CRDate), "", Format(TmpRs5!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 2) = IIf(IsNull(TmpRs5!LocationName), "", TmpRs5!LocationName)
        MSHFlexGrid12.TextMatrix(I, 3) = IIf(IsNull(TmpRs5!InstrumentType), "", TmpRs5!InstrumentType)
        MSHFlexGrid12.TextMatrix(I, 4) = IIf(IsNull(TmpRs5!DrawnOn), "", TmpRs5!DrawnOn)
        MSHFlexGrid12.TextMatrix(I, 5) = IIf(IsNull(TmpRs5!BranchName), "", TmpRs5!BranchName)
        MSHFlexGrid12.TextMatrix(I, 6) = IIf(IsNull(TmpRs5!InstrumentNo), "", TmpRs5!InstrumentNo)
        MSHFlexGrid12.TextMatrix(I, 7) = IIf(IsNull(TmpRs5!InstrumentDate), "", Format(TmpRs5!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 8) = IIf(IsNull(TmpRs5!InstrumentAmount), "", TmpRs5!InstrumentAmount)
        MSHFlexGrid12.TextMatrix(0, 8) = Val(MSHFlexGrid12.TextMatrix(0, 8)) + Val(MSHFlexGrid12.TextMatrix(I, 8))
        MSHFlexGrid12.TextMatrix(I, 9) = IIf(IsNull(TmpRs5!ReceivedFrom), "", TmpRs5!ReceivedFrom)
        MSHFlexGrid12.TextMatrix(I, 10) = IIf(IsNull(TmpRs5!Remark), "", TmpRs5!Remark)
        MSHFlexGrid12.TextMatrix(I, 11) = IIf(IsNull(TmpRs5!BatchID), "", TmpRs5!BatchID)
        MSHFlexGrid12.TextMatrix(I, 12) = IIf(IsNull(TmpRs5!DepositDate), "", Format(TmpRs5!DepositDate, "dd-MMM-yyyy"))
        MSHFlexGrid12.TextMatrix(I, 13) = IIf(IsNull(TmpRs5!FAGPostingDate), "", Format(TmpRs5!FAGPostingDate, "dd-MMM-yyyy"))
        
        MSHFlexGrid15.TextMatrix(j, 0) = IIf(IsNull(TmpRs5!CRID), "", TmpRs5!CRID)
        MSHFlexGrid15.TextMatrix(j, 1) = IIf(IsNull(TmpRs5!CRDate), "", Format(TmpRs5!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid15.TextMatrix(j, 2) = IIf(IsNull(TmpRs5!LocationName), "", TmpRs5!LocationName)
        MSHFlexGrid15.TextMatrix(j, 3) = IIf(IsNull(TmpRs5!InstrumentType), "", TmpRs5!InstrumentType)
        MSHFlexGrid15.TextMatrix(j, 4) = IIf(IsNull(TmpRs5!DrawnOn), "", TmpRs5!DrawnOn)
        MSHFlexGrid15.TextMatrix(j, 5) = IIf(IsNull(TmpRs5!BranchName), "", TmpRs5!BranchName)
        MSHFlexGrid15.TextMatrix(j, 6) = IIf(IsNull(TmpRs5!InstrumentNo), "", TmpRs5!InstrumentNo)
        MSHFlexGrid15.TextMatrix(j, 7) = IIf(IsNull(TmpRs5!InstrumentDate), "", Format(TmpRs5!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid15.TextMatrix(j, 8) = IIf(IsNull(TmpRs5!InstrumentAmount), "", TmpRs5!InstrumentAmount)
        MSHFlexGrid15.TextMatrix(0, 8) = Val(MSHFlexGrid15.TextMatrix(0, 8)) + Val(MSHFlexGrid15.TextMatrix(j, 8))
        MSHFlexGrid15.TextMatrix(j, 9) = IIf(IsNull(TmpRs5!ReceivedFrom), "", TmpRs5!ReceivedFrom)
        MSHFlexGrid15.TextMatrix(j, 10) = IIf(IsNull(TmpRs5!Remark), "", TmpRs5!Remark)
        MSHFlexGrid15.TextMatrix(j, 11) = IIf(IsNull(TmpRs5!BatchID), "", TmpRs5!BatchID)
        MSHFlexGrid15.TextMatrix(j, 12) = IIf(IsNull(TmpRs5!DepositDate), "", Format(TmpRs5!DepositDate, "dd-MMM-yyyy"))
        MSHFlexGrid15.TextMatrix(j, 13) = IIf(IsNull(TmpRs5!FAGPostingDate), "", Format(TmpRs5!FAGPostingDate, "dd-MMM-yyyy"))
        
        I = I + 1
        j = j + 1
        TmpRs5.MoveNext
    Next
End If
End Sub

Private Sub AddAdvanceCashReceiptPending()

'''''''''''''''''''''''--For TI Uploaded pending
tmp = "Select distinct TCR.CRID,TCR.SM_Prefix,CRDate,LocationName,InstrumentType,DrawnOn,BranchName,InstrumentNo,InstrumentDate,"
tmp = tmp & " InstrumentAmount , ReceivedFrom, TCR.Remark, AdvanceFlag, TCR.batchID"
tmp = tmp & " from Txn_NewCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " inner join Txn_NewCashReceiptTIDetail TCC on TCC.CRID=TCR.CRID And TCC.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y' And isnull(TCR.RejectFlag,'')<>'Y' And TCR.AdvanceFlag = 1  And batchID > 0 And TCC.EntryNo Is Null And TCR.CRDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.CRDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "' And isnull(TCC.CRAAID,0)=0"
tmp = tmp & " Union"
tmp = tmp & " Select distinct TCR.CRID,TCR.SM_Prefix,CRDate,LocationName,InstrumentType,DrawnOn,BranchName,InstrumentNo,InstrumentDate,"
tmp = tmp & " InstrumentAmount , ReceivedFrom, TCR.Remark, AdvanceFlag, TCR.batchID"
tmp = tmp & " from Txn_NewCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " inner join Txn_NewCashReceiptClientDetail TCC on TCC.CRID=TCR.CRID And TCC.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Where TCR.ReturnFlag<>'Y' And isnull(TCR.RejectFlag,'')<>'Y' And TCR.AdvanceFlag = 1 And batchID > 0 And TCC.EntryNo Is Null And TCC.Amount > 0 And TCR.CRDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.CRDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"

Call Tmp3Table

'-for allocation pending
''tmp = "Select distinct TCR.CRID,TCR.SM_Prefix,CRDate,LocationName,InstrumentType,DrawnOn,BranchName,InstrumentNo,InstrumentDate,"
''tmp = tmp & " InstrumentAmount , ReceivedFrom, TCR.Remark, AdvanceFlag, TCR.BatchID"
''tmp = tmp & " from Txn_NewCashReceipt TCR"
''tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
''tmp = tmp & " Inner join Txn_NewCashReceiptClientDetail TCRD on TCR.CRID=TCRD.CRID and TCR.SM_Prefix=TCRD.SM_Prefix"
''tmp = tmp & " Where TCR.AdvanceFlag = 1 And TCR.ReturnFlag<>'Y'"
''tmp = tmp & " and (TCR.InstrumentAmount-(Select isnull(sum(RecAmount),0) from Txn_NewCashReceiptTIDetail Where CRID= TCR.CRID and SM_Prefix=TCR.SM_Prefix))>0 And TCR.CRDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.CRDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"

tmp = "Select distinct TCR.CRID,TCR.SM_Prefix,CRDate,LocationName,InstrumentType,DrawnOn,BranchName,InstrumentNo,InstrumentDate,"
tmp = tmp & " InstrumentAmount , ReceivedFrom, TCR.Remark, AdvanceFlag, TCR.BatchID"
tmp = tmp & " from Txn_NewCashReceipt TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " inner join Txn_CRAdvanceAllocation TCA  on TCR.CRID=TCA.CRID and TCR.SM_Prefix=TCA.SM_Prefix"
tmp = tmp & " Where TCR.AdvanceFlag = 1 And isnull(TCR.RejectFlag,'')<>'Y' And TCR.ReturnFlag<>'Y' and isnull(TCA.batchID,0)=0"
tmp = tmp & " And TCR.CRDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.CRDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Order by LocationName, TCR.CRID, CRDate, BranchName, InstrumentAmount, TCR.BatchID "
Call Tmp4Table

MSHFlexGrid1.TextMatrix(7, 1) = "NA"
MSHFlexGrid1.TextMatrix(7, 2) = "NA"
MSHFlexGrid1.TextMatrix(7, 3) = TmpRs3.RecordCount
MSHFlexGrid1.TextMatrix(7, 4) = TmpRs4.RecordCount
MSHFlexGrid1.TextMatrix(7, 5) = "NA"

'MSHFlexGrid1.Row = 7
'MSHFlexGrid1.Col = 1
'MSHFlexGrid1.CellBackColor = vbBlack
'MSHFlexGrid1.Col = 2
'MSHFlexGrid1.CellBackColor = vbBlack
'MSHFlexGrid1.Col = 5
'MSHFlexGrid1.CellBackColor = vbBlack

Call Grid_Head32
If TmpRs3.RecordCount > 0 Then
    I = MSHFlexGrid13.Rows - 1
    MSHFlexGrid13.Rows = MSHFlexGrid13.Rows + TmpRs3.RecordCount
    MSHFlexGrid32.Rows = MSHFlexGrid32.Rows + TmpRs3.RecordCount
    j = 2
    For k = 1 To TmpRs3.RecordCount
        MSHFlexGrid13.TextMatrix(I, 0) = IIf(IsNull(TmpRs3!CRID), "", TmpRs3!CRID)
        MSHFlexGrid13.TextMatrix(I, 1) = IIf(IsNull(TmpRs3!CRDate), "", Format(TmpRs3!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid13.TextMatrix(I, 2) = IIf(IsNull(TmpRs3!LocationName), "", TmpRs3!LocationName)
        MSHFlexGrid13.TextMatrix(I, 3) = IIf(IsNull(TmpRs3!InstrumentType), "", TmpRs3!InstrumentType)
        MSHFlexGrid13.TextMatrix(I, 4) = IIf(IsNull(TmpRs3!DrawnOn), "", TmpRs3!DrawnOn)
        MSHFlexGrid13.TextMatrix(I, 5) = IIf(IsNull(TmpRs3!BranchName), "", TmpRs3!BranchName)
        MSHFlexGrid13.TextMatrix(I, 6) = IIf(IsNull(TmpRs3!InstrumentNo), "", TmpRs3!InstrumentNo)
        MSHFlexGrid13.TextMatrix(I, 7) = IIf(IsNull(TmpRs3!InstrumentDate), "", Format(TmpRs3!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid13.TextMatrix(I, 8) = IIf(IsNull(TmpRs3!InstrumentAmount), "", TmpRs3!InstrumentAmount)
        MSHFlexGrid13.TextMatrix(0, 8) = Val(MSHFlexGrid13.TextMatrix(0, 8)) + Val(MSHFlexGrid13.TextMatrix(I, 8))
        MSHFlexGrid13.TextMatrix(I, 9) = IIf(IsNull(TmpRs3!ReceivedFrom), "", TmpRs3!ReceivedFrom)
        MSHFlexGrid13.TextMatrix(I, 10) = IIf(IsNull(TmpRs3!Remark), "", TmpRs3!Remark)
        MSHFlexGrid13.TextMatrix(I, 11) = IIf(IsNull(TmpRs3!BatchID), "", TmpRs3!BatchID)
        
        MSHFlexGrid32.TextMatrix(j, 0) = IIf(IsNull(TmpRs3!CRID), "", TmpRs3!CRID)
        MSHFlexGrid32.TextMatrix(j, 1) = IIf(IsNull(TmpRs3!CRDate), "", Format(TmpRs3!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid32.TextMatrix(j, 2) = IIf(IsNull(TmpRs3!LocationName), "", TmpRs3!LocationName)
        MSHFlexGrid32.TextMatrix(j, 3) = IIf(IsNull(TmpRs3!InstrumentType), "", TmpRs3!InstrumentType)
        MSHFlexGrid32.TextMatrix(j, 4) = IIf(IsNull(TmpRs3!DrawnOn), "", TmpRs3!DrawnOn)
        MSHFlexGrid32.TextMatrix(j, 5) = IIf(IsNull(TmpRs3!BranchName), "", TmpRs3!BranchName)
        MSHFlexGrid32.TextMatrix(j, 6) = IIf(IsNull(TmpRs3!InstrumentNo), "", TmpRs3!InstrumentNo)
        MSHFlexGrid32.TextMatrix(j, 7) = IIf(IsNull(TmpRs3!InstrumentDate), "", Format(TmpRs3!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid32.TextMatrix(j, 8) = IIf(IsNull(TmpRs3!InstrumentAmount), "", TmpRs3!InstrumentAmount)
        MSHFlexGrid32.TextMatrix(0, 8) = Val(MSHFlexGrid32.TextMatrix(0, 8)) + Val(MSHFlexGrid32.TextMatrix(j, 8))
        MSHFlexGrid32.TextMatrix(j, 9) = IIf(IsNull(TmpRs3!ReceivedFrom), "", TmpRs3!ReceivedFrom)
        MSHFlexGrid32.TextMatrix(j, 10) = IIf(IsNull(TmpRs3!Remark), "", TmpRs3!Remark)
        MSHFlexGrid32.TextMatrix(j, 11) = IIf(IsNull(TmpRs3!BatchID), "", TmpRs3!BatchID)
        
        I = I + 1
        j = j + 1
        TmpRs3.MoveNext
    Next
End If

Call Grid_Head33
If TmpRs4.RecordCount > 0 Then
    I = MSHFlexGrid13.Rows - 1
    MSHFlexGrid13.Rows = MSHFlexGrid13.Rows + TmpRs4.RecordCount
    MSHFlexGrid33.Rows = MSHFlexGrid33.Rows + TmpRs4.RecordCount
    j = 2
    For k = 1 To TmpRs4.RecordCount
        MSHFlexGrid13.TextMatrix(I, 0) = IIf(IsNull(TmpRs4!CRID), "", TmpRs4!CRID)
        MSHFlexGrid13.TextMatrix(I, 1) = IIf(IsNull(TmpRs4!CRDate), "", Format(TmpRs4!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid13.TextMatrix(I, 2) = IIf(IsNull(TmpRs4!LocationName), "", TmpRs4!LocationName)
        MSHFlexGrid13.TextMatrix(I, 3) = IIf(IsNull(TmpRs4!InstrumentType), "", TmpRs4!InstrumentType)
        MSHFlexGrid13.TextMatrix(I, 4) = IIf(IsNull(TmpRs4!DrawnOn), "", TmpRs4!DrawnOn)
        MSHFlexGrid13.TextMatrix(I, 5) = IIf(IsNull(TmpRs4!BranchName), "", TmpRs4!BranchName)
        MSHFlexGrid13.TextMatrix(I, 6) = IIf(IsNull(TmpRs4!InstrumentNo), "", TmpRs4!InstrumentNo)
        MSHFlexGrid13.TextMatrix(I, 7) = IIf(IsNull(TmpRs4!InstrumentDate), "", Format(TmpRs4!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid13.TextMatrix(I, 8) = IIf(IsNull(TmpRs4!InstrumentAmount), "", TmpRs4!InstrumentAmount)
        MSHFlexGrid13.TextMatrix(0, 8) = Val(MSHFlexGrid13.TextMatrix(0, 8)) + Val(MSHFlexGrid13.TextMatrix(I, 8))
        MSHFlexGrid13.TextMatrix(I, 9) = IIf(IsNull(TmpRs4!ReceivedFrom), "", TmpRs4!ReceivedFrom)
        MSHFlexGrid13.TextMatrix(I, 10) = IIf(IsNull(TmpRs4!Remark), "", TmpRs4!Remark)
        MSHFlexGrid13.TextMatrix(I, 11) = IIf(IsNull(TmpRs4!BatchID), "", TmpRs4!BatchID)
        
        MSHFlexGrid33.TextMatrix(j, 0) = IIf(IsNull(TmpRs4!CRID), "", TmpRs4!CRID)
        MSHFlexGrid33.TextMatrix(j, 1) = IIf(IsNull(TmpRs4!CRDate), "", Format(TmpRs4!CRDate, "dd-MMM-yyyy"))
        MSHFlexGrid33.TextMatrix(j, 2) = IIf(IsNull(TmpRs4!LocationName), "", TmpRs4!LocationName)
        MSHFlexGrid33.TextMatrix(j, 3) = IIf(IsNull(TmpRs4!InstrumentType), "", TmpRs4!InstrumentType)
        MSHFlexGrid33.TextMatrix(j, 4) = IIf(IsNull(TmpRs4!DrawnOn), "", TmpRs4!DrawnOn)
        MSHFlexGrid33.TextMatrix(j, 5) = IIf(IsNull(TmpRs4!BranchName), "", TmpRs4!BranchName)
        MSHFlexGrid33.TextMatrix(j, 6) = IIf(IsNull(TmpRs4!InstrumentNo), "", TmpRs4!InstrumentNo)
        MSHFlexGrid33.TextMatrix(j, 7) = IIf(IsNull(TmpRs4!InstrumentDate), "", Format(TmpRs4!InstrumentDate, "dd-MMM-yyyy"))
        MSHFlexGrid33.TextMatrix(j, 8) = IIf(IsNull(TmpRs4!InstrumentAmount), "", TmpRs4!InstrumentAmount)
        MSHFlexGrid33.TextMatrix(0, 8) = Val(MSHFlexGrid33.TextMatrix(0, 8)) + Val(MSHFlexGrid33.TextMatrix(j, 8))
        MSHFlexGrid33.TextMatrix(j, 9) = IIf(IsNull(TmpRs4!ReceivedFrom), "", TmpRs4!ReceivedFrom)
        MSHFlexGrid33.TextMatrix(j, 10) = IIf(IsNull(TmpRs4!Remark), "", TmpRs4!Remark)
        MSHFlexGrid33.TextMatrix(j, 11) = IIf(IsNull(TmpRs4!BatchID), "", TmpRs4!BatchID)
        
        I = I + 1
        j = j + 1
        TmpRs4.MoveNext
    Next
End If
End Sub


Private Sub AddOtherAdjPending()

'--For Pending By FAG
Call Grid_Head13
tmp = "Select TCR.TDSAID,TCR.SM_Prefix,AdjDate,LocationName,Remark,  BatchID, "
tmp = tmp & " Case ShowDP When 0 then MCL.ClientName When 1 Then MD.DepositoryName end ClientName, "
tmp = tmp & " (Select isnull(sum(RecTDS),0) from Txn_CRTDSAdjustmentDetail Where TDSAID = TCR.TDSAID And SM_Prefix = TCR.SM_Prefix) RecTDS,"
tmp = tmp & " (Select isnull(sum(RecOtherDed),0) from Txn_CRTDSAdjustmentDetail Where TDSAID = TCR.TDSAID And SM_Prefix = TCR.SM_Prefix) RecOtherDed"
tmp = tmp & " from Txn_CRTDSAdjustment TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " Left Join Mst_Client MCL on MCL.CLientIDRow = TCR.CLientIDRow"
tmp = tmp & " Left Join Mst_Depository MD on MD.DepositoryID = TCR.DPID"
tmp = tmp & " Where TCR.Flag='P' And TCR.AdjDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.AdjDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Order by LocationName, ClientName, TDSAID, AdjDate, Remark, BatchID "

Call Tmp1Table

'--For Batch Pending By FAG

tmp = "Select TCR.TDSAID,TCR.SM_Prefix,AdjDate,LocationName,Remark,  BatchID, "
tmp = tmp & " Case ShowDP When 0 then MCL.ClientName When 1 Then MD.DepositoryName end ClientName, "
tmp = tmp & " (Select isnull(sum(RecTDS),0) from Txn_CRTDSAdjustmentDetail Where TDSAID = TCR.TDSAID And SM_Prefix = TCR.SM_Prefix) RecTDS,"
tmp = tmp & " (Select isnull(sum(RecOtherDed),0) from Txn_CRTDSAdjustmentDetail Where TDSAID = TCR.TDSAID And SM_Prefix = TCR.SM_Prefix) RecOtherDed"
tmp = tmp & " from Txn_CRTDSAdjustment TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " Left Join Mst_Client MCL on MCL.CLientIDRow = TCR.CLientIDRow"
tmp = tmp & " Left Join Mst_Depository MD on MD.DepositoryID = TCR.DPID"
tmp = tmp & " Where TCR.Flag='A' And isnull(TCR.BatchID,0)=0 And TCR.AdjDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.AdjDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Order by LocationName, TDSAID, AdjDate, Remark, BatchID "
Call Tmp2Table

'''''''''''''''''''''''--For TI Uploaded pendig
tmp = "Select distinct TCR.TDSAID,TCR.SM_Prefix,AdjDate,LocationName,Remark,TCR.batchID, "
tmp = tmp & " Case ShowDP When 0 then MCL.ClientName When 1 Then MD.DepositoryName end ClientName, "
tmp = tmp & " (Select isnull(sum(RecTDS),0) from Txn_CRTDSAdjustmentDetail Where TDSAID = TCR.TDSAID And SM_Prefix = TCR.SM_Prefix) RecTDS,"
tmp = tmp & " (Select isnull(sum(RecOtherDed),0) from Txn_CRTDSAdjustmentDetail Where TDSAID = TCR.TDSAID And SM_Prefix = TCR.SM_Prefix) RecOtherDed"
tmp = tmp & " from Txn_CRTDSAdjustment TCR"
tmp = tmp & " Inner Join Mst_Location ML On TCR.LocationID = ML.LocationID"
tmp = tmp & " inner join Txn_CRTDSAdjustmentDetail TCC on TCC.TDSAID=TCR.TDSAID And TCC.SM_Prefix=TCR.SM_Prefix"
tmp = tmp & " Left Join Mst_Client MCL on MCL.CLientIDRow = TCR.CLientIDRow"
tmp = tmp & " Left Join Mst_Depository MD on MD.DepositoryID = TCR.DPID"
tmp = tmp & " Where batchID > 0 And (TCC.TDSEntryNo Is Null and TCC.OtherDedEntryNo Is Null) And TCR.AdjDate>='" & Format(TxtFromDate.Value, "yyyy-MM-dd") & "' And TCR.AdjDate<='" & Format(TxtToDate.Value, "yyyy-MM-dd") & "'"
tmp = tmp & " Order by LocationName, TCR.TDSAID, AdjDate, Remark, BatchID "
Call Tmp3Table

MSHFlexGrid1.TextMatrix(8, 1) = TmpRs1.RecordCount
MSHFlexGrid1.TextMatrix(8, 2) = TmpRs2.RecordCount
MSHFlexGrid1.TextMatrix(8, 3) = TmpRs3.RecordCount
MSHFlexGrid1.TextMatrix(8, 4) = "NA"
MSHFlexGrid1.TextMatrix(8, 5) = "NA"

'MSHFlexGrid1.Row = 8
'MSHFlexGrid1.Col = 4
'MSHFlexGrid1.CellBackColor = vbBlack
'MSHFlexGrid1.Col = 5
'MSHFlexGrid1.CellBackColor = vbBlack
Call Grid_Head34
If TmpRs1.RecordCount > 0 Then
    I = MSHFlexGrid14.Rows - 1
    MSHFlexGrid14.Rows = MSHFlexGrid14.Rows + TmpRs1.RecordCount
    MSHFlexGrid34.Rows = MSHFlexGrid34.Rows + TmpRs1.RecordCount
    j = 2
    For k = 1 To TmpRs1.RecordCount
        MSHFlexGrid14.TextMatrix(I, 0) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
        MSHFlexGrid14.TextMatrix(I, 1) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
        MSHFlexGrid14.TextMatrix(I, 2) = IIf(IsNull(TmpRs1!AdjDate), "", Format(TmpRs1!AdjDate, "dd-MMM-yyyy"))
        MSHFlexGrid14.TextMatrix(I, 3) = IIf(IsNull(TmpRs1!RecTDS), "", TmpRs1!RecTDS)
        MSHFlexGrid14.TextMatrix(0, 3) = Val(MSHFlexGrid14.TextMatrix(0, 3)) + Val(MSHFlexGrid14.TextMatrix(I, 3))
        MSHFlexGrid14.TextMatrix(I, 4) = IIf(IsNull(TmpRs1!RecOtherDed), "", TmpRs1!RecOtherDed)
        MSHFlexGrid14.TextMatrix(0, 4) = Val(MSHFlexGrid14.TextMatrix(0, 4)) + Val(MSHFlexGrid14.TextMatrix(I, 4))
        MSHFlexGrid14.TextMatrix(I, 5) = IIf(IsNull(TmpRs1!Remark), "", TmpRs1!Remark)
        MSHFlexGrid14.TextMatrix(I, 6) = IIf(IsNull(TmpRs1!BatchID), "", TmpRs1!BatchID)
        
        MSHFlexGrid34.TextMatrix(j, 0) = IIf(IsNull(TmpRs1!LocationName), "", TmpRs1!LocationName)
        MSHFlexGrid34.TextMatrix(j, 1) = IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName)
        MSHFlexGrid34.TextMatrix(j, 2) = IIf(IsNull(TmpRs1!AdjDate), "", Format(TmpRs1!AdjDate, "dd-MMM-yyyy"))
        MSHFlexGrid34.TextMatrix(j, 3) = IIf(IsNull(TmpRs1!RecTDS), "", TmpRs1!RecTDS)
        MSHFlexGrid34.TextMatrix(0, 3) = Val(MSHFlexGrid34.TextMatrix(0, 3)) + Val(MSHFlexGrid34.TextMatrix(j, 3))
        MSHFlexGrid34.TextMatrix(j, 4) = IIf(IsNull(TmpRs1!RecOtherDed), "", TmpRs1!RecOtherDed)
        MSHFlexGrid34.TextMatrix(0, 4) = Val(MSHFlexGrid34.TextMatrix(0, 4)) + Val(MSHFlexGrid34.TextMatrix(j, 4))
        MSHFlexGrid34.TextMatrix(j, 5) = IIf(IsNull(TmpRs1!Remark), "", TmpRs1!Remark)
        MSHFlexGrid34.TextMatrix(j, 6) = IIf(IsNull(TmpRs1!BatchID), "", TmpRs1!BatchID)
        
        I = I + 1
        j = j + 1
        TmpRs1.MoveNext
    Next
End If

Call Grid_Head35
If TmpRs2.RecordCount > 0 Then
    I = MSHFlexGrid14.Rows - 1
    MSHFlexGrid14.Rows = MSHFlexGrid14.Rows + TmpRs2.RecordCount
    MSHFlexGrid35.Rows = MSHFlexGrid35.Rows + TmpRs2.RecordCount
    j = 2
    For k = 1 To TmpRs2.RecordCount
        MSHFlexGrid14.TextMatrix(I, 0) = IIf(IsNull(TmpRs2!LocationName), "", TmpRs2!LocationName)
        MSHFlexGrid14.TextMatrix(I, 1) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
        MSHFlexGrid14.TextMatrix(I, 2) = IIf(IsNull(TmpRs2!AdjDate), "", Format(TmpRs2!AdjDate, "dd-MMM-yyyy"))
        MSHFlexGrid14.TextMatrix(I, 3) = IIf(IsNull(TmpRs2!RecTDS), "", TmpRs2!RecTDS)
        MSHFlexGrid14.TextMatrix(0, 3) = Val(MSHFlexGrid14.TextMatrix(0, 3)) + Val(MSHFlexGrid14.TextMatrix(I, 3))
        MSHFlexGrid14.TextMatrix(I, 4) = IIf(IsNull(TmpRs2!RecOtherDed), "", TmpRs2!RecOtherDed)
        MSHFlexGrid14.TextMatrix(0, 4) = Val(MSHFlexGrid14.TextMatrix(0, 4)) + Val(MSHFlexGrid14.TextMatrix(I, 4))
        MSHFlexGrid14.TextMatrix(I, 5) = IIf(IsNull(TmpRs2!Remark), "", TmpRs2!Remark)
        MSHFlexGrid14.TextMatrix(I, 6) = IIf(IsNull(TmpRs2!BatchID), "", TmpRs2!BatchID)
        
        MSHFlexGrid35.TextMatrix(j, 0) = IIf(IsNull(TmpRs2!LocationName), "", TmpRs2!LocationName)
        MSHFlexGrid35.TextMatrix(j, 1) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
        MSHFlexGrid35.TextMatrix(j, 2) = IIf(IsNull(TmpRs2!AdjDate), "", Format(TmpRs2!AdjDate, "dd-MMM-yyyy"))
        MSHFlexGrid35.TextMatrix(j, 3) = IIf(IsNull(TmpRs2!RecTDS), "", TmpRs2!RecTDS)
        MSHFlexGrid35.TextMatrix(0, 3) = Val(MSHFlexGrid35.TextMatrix(0, 3)) + Val(MSHFlexGrid35.TextMatrix(j, 3))
        MSHFlexGrid35.TextMatrix(j, 4) = IIf(IsNull(TmpRs2!RecOtherDed), "", TmpRs2!RecOtherDed)
        MSHFlexGrid35.TextMatrix(0, 4) = Val(MSHFlexGrid35.TextMatrix(0, 4)) + Val(MSHFlexGrid35.TextMatrix(j, 4))
        MSHFlexGrid35.TextMatrix(j, 5) = IIf(IsNull(TmpRs2!Remark), "", TmpRs2!Remark)
        MSHFlexGrid35.TextMatrix(j, 6) = IIf(IsNull(TmpRs2!BatchID), "", TmpRs2!BatchID)
        
        I = I + 1
        j = j + 1
        TmpRs2.MoveNext
    Next
End If
Call Grid_Head36

If TmpRs3.RecordCount > 0 Then
    I = MSHFlexGrid14.Rows - 1
    MSHFlexGrid14.Rows = MSHFlexGrid14.Rows + TmpRs3.RecordCount
    MSHFlexGrid36.Rows = MSHFlexGrid36.Rows + TmpRs3.RecordCount
    j = 2
    For k = 1 To TmpRs3.RecordCount
        MSHFlexGrid14.TextMatrix(I, 0) = IIf(IsNull(TmpRs3!LocationName), "", TmpRs3!LocationName)
        MSHFlexGrid14.TextMatrix(I, 1) = IIf(IsNull(TmpRs3!ClientName), "", TmpRs3!ClientName)
        MSHFlexGrid14.TextMatrix(I, 2) = IIf(IsNull(TmpRs3!AdjDate), "", Format(TmpRs3!AdjDate, "dd-MMM-yyyy"))
        MSHFlexGrid14.TextMatrix(I, 3) = IIf(IsNull(TmpRs3!RecTDS), "", TmpRs3!RecTDS)
        MSHFlexGrid14.TextMatrix(0, 3) = Val(MSHFlexGrid14.TextMatrix(0, 3)) + Val(MSHFlexGrid14.TextMatrix(I, 3))
        MSHFlexGrid14.TextMatrix(I, 4) = IIf(IsNull(TmpRs3!RecOtherDed), "", TmpRs3!RecOtherDed)
        MSHFlexGrid14.TextMatrix(0, 4) = Val(MSHFlexGrid14.TextMatrix(0, 4)) + Val(MSHFlexGrid14.TextMatrix(I, 4))
        MSHFlexGrid14.TextMatrix(I, 5) = IIf(IsNull(TmpRs3!Remark), "", TmpRs3!Remark)
        MSHFlexGrid14.TextMatrix(I, 6) = IIf(IsNull(TmpRs3!BatchID), "", TmpRs3!BatchID)
        
        MSHFlexGrid36.TextMatrix(j, 0) = IIf(IsNull(TmpRs3!LocationName), "", TmpRs3!LocationName)
        MSHFlexGrid36.TextMatrix(j, 1) = IIf(IsNull(TmpRs3!ClientName), "", TmpRs3!ClientName)
        MSHFlexGrid36.TextMatrix(j, 2) = IIf(IsNull(TmpRs3!AdjDate), "", Format(TmpRs3!AdjDate, "dd-MMM-yyyy"))
        MSHFlexGrid36.TextMatrix(j, 3) = IIf(IsNull(TmpRs3!RecTDS), "", TmpRs3!RecTDS)
        MSHFlexGrid36.TextMatrix(0, 3) = Val(MSHFlexGrid36.TextMatrix(0, 3)) + Val(MSHFlexGrid36.TextMatrix(j, 3))
        MSHFlexGrid36.TextMatrix(j, 4) = IIf(IsNull(TmpRs3!RecOtherDed), "", TmpRs3!RecOtherDed)
        MSHFlexGrid36.TextMatrix(0, 4) = Val(MSHFlexGrid36.TextMatrix(0, 4)) + Val(MSHFlexGrid36.TextMatrix(j, 4))
        MSHFlexGrid36.TextMatrix(j, 5) = IIf(IsNull(TmpRs3!Remark), "", TmpRs3!Remark)
        MSHFlexGrid36.TextMatrix(j, 6) = IIf(IsNull(TmpRs3!BatchID), "", TmpRs3!BatchID)
        
        I = I + 1
        j = j + 1
        TmpRs3.MoveNext
    Next
End If

End Sub


Private Sub FillPendingNAVCodes()
Call Grid_Head1

MSHFlexGrid2.Rows = 6
MSHFlexGrid2.TextMatrix(2, 0) = "Client Code Mapping"
MSHFlexGrid2.TextMatrix(3, 0) = "DP Code Mapping"
MSHFlexGrid2.TextMatrix(4, 0) = "Location Code Mapping"

tmp = "Select * from Mst_Client Where isnull(NAVCode,'')<>''"
Call TmpTable
MSHFlexGrid2.TextMatrix(2, 1) = TmpRs.RecordCount

tmp = "Select * from Mst_Client Where isnull(NAVCode,'')=''"
Call TmpTable
MSHFlexGrid2.TextMatrix(2, 2) = TmpRs.RecordCount

tmp = "Select * from Mst_Depository Where isnull(NAVCode,'')<>''"
Call TmpTable
MSHFlexGrid2.TextMatrix(3, 1) = TmpRs.RecordCount

tmp = "Select * from Mst_Depository Where isnull(NAVCode,'')=''"
Call TmpTable
MSHFlexGrid2.TextMatrix(3, 2) = TmpRs.RecordCount

tmp = "Select * from Mst_Location Where isnull(NAVCode,'')<>''"
Call TmpTable
MSHFlexGrid2.TextMatrix(4, 1) = TmpRs.RecordCount

tmp = "Select * from Mst_Location Where isnull(NAVCode,'')=''"
Call TmpTable
MSHFlexGrid2.TextMatrix(4, 2) = TmpRs.RecordCount

For I = 2 To MSHFlexGrid2.Rows - 1
    MSHFlexGrid2.TextMatrix(0, 1) = Val(MSHFlexGrid2.TextMatrix(0, 1)) + Val(MSHFlexGrid2.TextMatrix(I, 1))
    MSHFlexGrid2.TextMatrix(0, 2) = Val(MSHFlexGrid2.TextMatrix(0, 2)) + Val(MSHFlexGrid2.TextMatrix(I, 2))
Next

End Sub

Private Sub GetOutstandingDetails()

Call Grid_Head2
MSHFlexGrid3.Rows = 10

MSHFlexGrid3.TextMatrix(2, 0) = "Navision Client Ledger Uploaded up to Date"
MSHFlexGrid3.TextMatrix(3, 0) = "Total Outstanding As per StoreMan Invoice (Balance Amount)  (A)"
MSHFlexGrid3.TextMatrix(4, 0) = "Cash Receipt without NAV Doc No without Advance                      (B)"
MSHFlexGrid3.TextMatrix(5, 0) = "Cash Receipt Advance Amount                                                                 (C)"
MSHFlexGrid3.TextMatrix(6, 0) = "Total Outstanding As per StoreMan                                                        (D=A+B-C)"
MSHFlexGrid3.TextMatrix(7, 0) = "Total Outstanding As per Navision                                                          (E)"
MSHFlexGrid3.TextMatrix(8, 0) = "Difference                                                                                                          (F=D-E)"

tmp = "Select Max(ActualDatePosted) as 'LastUpdation' from Txn_ClientLedger Where ActualDatePosted <= GetDate() "
Call TmpTable

MSHFlexGrid3.TextMatrix(2, 0) = "-"
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.TextMatrix(2, 0) = Format(TmpRs!LastUpdation, "dd-mmm-yyyy")
End If

If MSHFlexGrid3.TextMatrix(2, 0) <> "-" Then
   tmp = "Select Max(ActualTimePosted) as 'LastUpdation' from Txn_ClientLedger Where ActualDatePosted = '" & MSHFlexGrid3.TextMatrix(2, 0) & "' "
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      MSHFlexGrid3.TextMatrix(2, 0) = MSHFlexGrid3.TextMatrix(2, 0) & " " & TmpRs!LastUpdation
   End If
End If

MSHFlexGrid3.TextMatrix(2, 0) = "Navision Client Ledger Uploaded up to Date " & MSHFlexGrid3.TextMatrix(2, 0)

Call Grid_Head3


''tmp = "Select TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode ,'' ClientName,"
''tmp = tmp & " (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCRT.EntryNo > 0 )RecAmount,"
''tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCRT.EntryNo > 0) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID And (TCRT.TDSEntryNo > 0 or TCRT.OtherDedEntryNo >0))RecTDS,"
''tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCRT.EntryNo > 0 )+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID and(TCRT.TDSEntryNo > 0 or TCRT.OtherDedEntryNo > 0))RecOtherDed"
''tmp = tmp & " ,MC.ClientID+'~'+MC.ClientName as ClientName ,MD.DepositoryID+'~'+ MD.DepositoryName as DepositoryName"
''tmp = tmp & " From  Txn_TIDetails TI"
''tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
''tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
''tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
''tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
''tmp = tmp & " Order By TI.TIDate,TI.NAVCode"

'--For Invoice balance (with Entry No)AdjAmountFAG


tmp = "Select TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode ,'' ClientName,"
tmp = tmp & " isnull(AdjAmountFAG,0) + (Select isnull(Sum(TCRT.RecAmount),0) as RecAmount From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCRT.EntryNo > 0 And isnull(TCRT.CRAAID,0)=0)RecAmount,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCRT.EntryNo > 0 And isnull(TCRT.CRAAID,0)=0) + (Select isnull(Sum(TCRT.RecTDS),0) as RecTDS From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID And (TCRT.TDSEntryNo > 0 or TCRT.OtherDedEntryNo >0))RecTDS,"
tmp = tmp & " (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_NewCashReceiptTIDetail TCRT Inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.ReturnFlag<>'Y'And TIID=TI.TIID And TCRT.EntryNo > 0 And isnull(TCRT.CRAAID,0)=0)+ (Select isnull(Sum(TCRT.RecOtherDed),0) as RecOtherDed From Txn_CRTDSAdjustmentDetail TCRT Inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix Where TCR.Flag<>'R'And TIID=TI.TIID and(TCRT.TDSEntryNo > 0 or TCRT.OtherDedEntryNo > 0))RecOtherDed"
tmp = tmp & " ,MC.ClientID+'~'+MC.ClientName as ClientName ,MD.DepositoryID+'~'+ MD.DepositoryName as DepositoryName"
tmp = tmp & " From  Txn_TIDetails TI"
tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Order By TI.TIDate,TI.NAVCode"

Call TmpTable

MSHFlexGrid4.Rows = TmpRs.RecordCount + 3
For I = 2 To TmpRs.RecordCount + 1

    For c = 0 To TmpRs.Fields.Count - 3
        MSHFlexGrid4.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    If InStr(1, TmpRs!NAVCode, "CSWD") > 0 Then
       MSHFlexGrid4.TextMatrix(I, 9) = IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName)
    Else
       MSHFlexGrid4.TextMatrix(I, 9) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
    End If

    MSHFlexGrid4.TextMatrix(I, 1) = Format(MSHFlexGrid4.TextMatrix(I, 1), "dd-MMM-yyyy")
    MSHFlexGrid4.TextMatrix(0, 2) = Val(MSHFlexGrid4.TextMatrix(0, 2)) + Val(MSHFlexGrid4.TextMatrix(I, 2))
    MSHFlexGrid4.TextMatrix(0, 3) = Val(MSHFlexGrid4.TextMatrix(0, 3)) + Val(MSHFlexGrid4.TextMatrix(I, 3))
    MSHFlexGrid4.TextMatrix(0, 10) = Val(MSHFlexGrid4.TextMatrix(0, 10)) + Val(MSHFlexGrid4.TextMatrix(I, 10))
    MSHFlexGrid4.TextMatrix(0, 11) = Val(MSHFlexGrid4.TextMatrix(0, 11)) + Val(MSHFlexGrid4.TextMatrix(I, 11))
    MSHFlexGrid4.TextMatrix(0, 12) = Val(MSHFlexGrid4.TextMatrix(0, 12)) + Val(MSHFlexGrid4.TextMatrix(I, 12))
    MSHFlexGrid4.TextMatrix(I, 13) = Format(Val(MSHFlexGrid4.TextMatrix(I, 2)) - (Val(MSHFlexGrid4.TextMatrix(I, 10)) + Val(MSHFlexGrid4.TextMatrix(I, 11)) + Val(MSHFlexGrid4.TextMatrix(I, 12))), "######0.00")
    MSHFlexGrid4.TextMatrix(0, 13) = Val(MSHFlexGrid4.TextMatrix(0, 13)) + Val(MSHFlexGrid4.TextMatrix(I, 13))
    TmpRs.MoveNext
Next

''tmp = "Select TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode ,'' ClientName,"
''tmp = tmp & " isnull(Sum(TCRT.RecAmount),0) as RecAmount,"
''tmp = tmp & " isnull(Sum(TCRT.RecTDS),0) as RecTDS,"
''tmp = tmp & " isnull(Sum(TCRT.RecOtherDed), 0) As RecOtherDed,MC.ClientID+'~'+MC.ClientName as ClientName ,MD.DepositoryID+'~'+ MD.DepositoryName as DepositoryName"
''tmp = tmp & " From  Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
''tmp = tmp & " inner Join Txn_TIDetails TI on TI.TIID=TCRT.TIID"
''tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
''tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
''tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
''tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
''tmp = tmp & " Where isnull(TCRT.EntryNo,0)=0 And TCR.ReturnFlag<>'Y'"
''tmp = tmp & " Group By TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode"
''tmp = tmp & " ,MC.ClientID,MC.ClientName ,MD.DepositoryID,MD.DepositoryName"


''---outstanding from SM
'--- Receved amt in CR without Doc No
 tmp = " Select TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode ,'' ClientName,"
 tmp = tmp & " isnull(AdjAmountFAG,0) + isnull(Sum(TCRT.RecAmount),0) as RecAmount,"
 tmp = tmp & " isnull(Sum(TCRT.RecTDS),0) as RecTDS,"
 tmp = tmp & " isnull(Sum(TCRT.RecOtherDed), 0) As RecOtherDed,MC.ClientID+'~'+MC.ClientName as ClientName ,MD.DepositoryID+'~'+ MD.DepositoryName as DepositoryName"
 tmp = tmp & " From  Txn_NewCashReceiptTIDetail TCRT"
 tmp = tmp & " inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
 tmp = tmp & " inner Join Txn_TIDetails TI on TI.TIID=TCRT.TIID"
 tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
 tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
 tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
 tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
 tmp = tmp & " Where isnull(TCRT.EntryNo,0)=0 And TCR.ReturnFlag<>'Y' And isnull(TCRT.CRAAID,0)=0"
 tmp = tmp & " Group By TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode, AdjAmountFAG"
 tmp = tmp & " ,MC.ClientID,MC.ClientName ,MD.DepositoryID,MD.DepositoryName"

Call TmpTable

Call Grid_Head4
MSHFlexGrid5.Rows = TmpRs.RecordCount + 3
For I = 2 To TmpRs.RecordCount + 1

    For c = 0 To TmpRs.Fields.Count - 3
        MSHFlexGrid5.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    If InStr(1, TmpRs!NAVCode, "CSWD") > 0 Then
       MSHFlexGrid5.TextMatrix(I, 9) = IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName)
    Else
       MSHFlexGrid5.TextMatrix(I, 9) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
    End If
       
    MSHFlexGrid5.TextMatrix(I, 1) = Format(MSHFlexGrid5.TextMatrix(I, 1), "dd-MMM-yyyy")
    MSHFlexGrid5.TextMatrix(0, 2) = Val(MSHFlexGrid5.TextMatrix(0, 2)) + Val(MSHFlexGrid5.TextMatrix(I, 2))
    MSHFlexGrid5.TextMatrix(0, 3) = Val(MSHFlexGrid5.TextMatrix(0, 3)) + Val(MSHFlexGrid5.TextMatrix(I, 3))
    MSHFlexGrid5.TextMatrix(0, 10) = Val(MSHFlexGrid5.TextMatrix(0, 10)) + Val(MSHFlexGrid5.TextMatrix(I, 10))
    MSHFlexGrid5.TextMatrix(0, 11) = Val(MSHFlexGrid5.TextMatrix(0, 11)) + Val(MSHFlexGrid5.TextMatrix(I, 11))
    MSHFlexGrid5.TextMatrix(0, 12) = Val(MSHFlexGrid5.TextMatrix(0, 12)) + Val(MSHFlexGrid5.TextMatrix(I, 12))
    MSHFlexGrid5.TextMatrix(I, 13) = Format(Val(MSHFlexGrid5.TextMatrix(I, 2)) - (Val(MSHFlexGrid5.TextMatrix(I, 10)) + Val(MSHFlexGrid5.TextMatrix(I, 11)) + Val(MSHFlexGrid5.TextMatrix(I, 12))), "######0.00")
    MSHFlexGrid5.TextMatrix(0, 13) = Val(MSHFlexGrid5.TextMatrix(0, 13)) + Val(MSHFlexGrid5.TextMatrix(I, 13))
    TmpRs.MoveNext
Next

'--- Receved amt in TDS Adj without Doc No

tmp = "Select TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode ,'' ClientName,0 as RecAmt,"
tmp = tmp & " isnull(Sum(TCRT.RecTDS),0) as RecTDS,"
tmp = tmp & " isnull(Sum(TCRT.RecOtherDed), 0) As RecOtherDed,MC.ClientID+'~'+MC.ClientName as ClientName ,MD.DepositoryID+'~'+ MD.DepositoryName as DepositoryName"
tmp = tmp & " From  Txn_CRTDSAdjustmentDetail TCRT"
tmp = tmp & " inner Join Txn_CRTDSAdjustment TCR on TCR.TDSAID= TCRT.TDSAID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " inner Join Txn_TIDetails TI on TI.TIID=TCRT.TIID"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Left Join Mst_Client MC on TCR.ClientIDRow=MC.ClientIDRow"
tmp = tmp & " Left Join Mst_Depository MD on TCR.DPID=MD.DepositoryID"
tmp = tmp & " Where isnull(TCRT.OtherDedEntryNo,0)=0 and isnull(TCRT.TDSEntryNo,0)=0 And TCR.Flag<>'R'"
tmp = tmp & " Group By TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode"
tmp = tmp & " ,MC.ClientID,MC.ClientName ,MD.DepositoryID,MD.DepositoryName"

Call TmpTable

For I = 1 To TmpRs.RecordCount
    r = MSHFlexGrid5.Rows - 1
    For c = 0 To TmpRs.Fields.Count - 3
        MSHFlexGrid5.TextMatrix(r, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    If InStr(1, TmpRs!NAVCode, "CSWD") > 0 Then
       MSHFlexGrid5.TextMatrix(r, 9) = IIf(IsNull(TmpRs!DepositoryName), "", TmpRs!DepositoryName)
    Else
       MSHFlexGrid5.TextMatrix(r, 9) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
    End If
       
    MSHFlexGrid5.TextMatrix(r, 1) = Format(MSHFlexGrid5.TextMatrix(r, 1), "dd-MMM-yyyy")
    MSHFlexGrid5.TextMatrix(0, 2) = Val(MSHFlexGrid5.TextMatrix(0, 2)) + Val(MSHFlexGrid5.TextMatrix(r, 2))
    MSHFlexGrid5.TextMatrix(0, 3) = Val(MSHFlexGrid5.TextMatrix(0, 3)) + Val(MSHFlexGrid5.TextMatrix(r, 3))
    MSHFlexGrid5.TextMatrix(0, 10) = Val(MSHFlexGrid5.TextMatrix(0, 10)) + Val(MSHFlexGrid5.TextMatrix(r, 10))
    MSHFlexGrid5.TextMatrix(0, 11) = Val(MSHFlexGrid5.TextMatrix(0, 11)) + Val(MSHFlexGrid5.TextMatrix(r, 11))
    MSHFlexGrid5.TextMatrix(0, 12) = Val(MSHFlexGrid5.TextMatrix(0, 12)) + Val(MSHFlexGrid5.TextMatrix(r, 12))
    MSHFlexGrid5.TextMatrix(r, 13) = Format(Val(MSHFlexGrid5.TextMatrix(r, 2)) - (Val(MSHFlexGrid5.TextMatrix(r, 10)) + Val(MSHFlexGrid5.TextMatrix(r, 11)) + Val(MSHFlexGrid5.TextMatrix(r, 12))), "######0.00")
    MSHFlexGrid5.TextMatrix(0, 13) = Val(MSHFlexGrid5.TextMatrix(0, 13)) + Val(MSHFlexGrid5.TextMatrix(r, 13))
    MSHFlexGrid5.Rows = MSHFlexGrid5.Rows + 1
    TmpRs.MoveNext
Next

'----To Subtract from Outstanding( allocated via allocation & batch Created)
'----Allocated via advance Allocated & batch in Created
tmp = "Select TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode ,"
tmp = tmp & " isnull(AdjAmountFAG,0) + isnull(Sum(TCRT.RecAmount),0) as RecAmount,"
tmp = tmp & " isnull(Sum(TCRT.RecTDS),0) as RecTDS,"
tmp = tmp & " isnull(Sum(TCRT.RecOtherDed), 0) As RecOtherDed,MC.ClientID+'~'+MC.ClientName as ClientName ,MD.DepositoryID+'~'+ MD.DepositoryName as DepositoryName,TCRT.CRID,TCRT.SM_Prefix"
tmp = tmp & " From  Txn_NewCashReceiptTIDetail TCRT"
tmp = tmp & " inner Join Txn_NewCashReceipt TCR on TCR.CRID= TCRT.CRID and TCR.SM_Prefix=TCRT.SM_Prefix"
tmp = tmp & " inner Join Txn_TIDetails TI on TI.TIID=TCRT.TIID"
tmp = tmp & " inner Join Txn_CRAdvanceAllocation CRAA on CRAA.CRAAID=TCRT.CRAAID"
tmp = tmp & " Left Join Mst_Location ML on TI.LocationID=ML.LocationID"
tmp = tmp & " Left Join Mst_State MS on ML.StateID=MS.StateID"
tmp = tmp & " Left Join Mst_Client MC on TI.ClientIDRow=MC.ClientIDRow"
tmp = tmp & " Left Join Mst_Depository MD on TI.DepositoryID=MD.DepositoryID"
tmp = tmp & " Where isnull(TCRT.EntryNo,0)=0 And TCR.ReturnFlag<>'Y' And isnull(CRAA.BatchID,0)>0"
tmp = tmp & " Group By TINo,TIDate,TIAmount,TISTAmount,MS.StateName,ML.LocationName,TI.PostInvoiceNo,TI.EntryNo,TI.NAVCode, AdjAmountFAG"
tmp = tmp & " ,MC.ClientID,MC.ClientName ,MD.DepositoryID,MD.DepositoryName,TCRT.CRID,TCRT.SM_Prefix"

Call Tmp4Table
Call Grid_Head6
Dim mallocAmt As Double
mallocAmt = 0
If TmpRs4.RecordCount > 0 Then
    MSHFlexGrid7.Rows = MSHFlexGrid7.Rows + TmpRs4.RecordCount
    For I = 2 To TmpRs4.RecordCount + 1
        mallocAmt = mallocAmt + (TmpRs4!RecAmount + TmpRs4!RecTDS + TmpRs4!RecOtherDed)
        MSHFlexGrid7.TextMatrix(I, 0) = IIf(IsNull(TmpRs4!StateName), "", TmpRs4!StateName)
        MSHFlexGrid7.TextMatrix(I, 1) = IIf(IsNull(TmpRs4!LocationName), "", TmpRs4!LocationName)
        MSHFlexGrid7.TextMatrix(I, 2) = IIf(IsNull(TmpRs4!TINo), "", TmpRs4!TINo)
        MSHFlexGrid7.TextMatrix(I, 3) = IIf(IsNull(TmpRs4!TIDate), "", Format(TmpRs4!TIDate, "dd-MMM-yyyy"))
        MSHFlexGrid7.TextMatrix(I, 4) = IIf(IsNull(TmpRs4!TIAmount), "", TmpRs4!TIAmount)
        MSHFlexGrid7.TextMatrix(0, 4) = Val(MSHFlexGrid7.TextMatrix(0, 4)) + Val(MSHFlexGrid7.TextMatrix(I, 4))
        MSHFlexGrid7.TextMatrix(I, 5) = IIf(IsNull(TmpRs4!PostInvoiceNo), "", TmpRs4!PostInvoiceNo)
        MSHFlexGrid7.TextMatrix(I, 6) = IIf(IsNull(TmpRs4!EntryNo), "", TmpRs4!EntryNo)
        MSHFlexGrid7.TextMatrix(I, 7) = IIf(IsNull(TmpRs4!NAVCode), "", TmpRs4!NAVCode)
        If IsNull(TmpRs4!ClientName) = True Then
            MSHFlexGrid7.TextMatrix(I, 8) = IIf(IsNull(TmpRs4!DepositoryName), "", TmpRs4!DepositoryName)
        Else
            MSHFlexGrid7.TextMatrix(I, 8) = TmpRs4!ClientName
        End If
        MSHFlexGrid7.TextMatrix(I, 9) = IIf(IsNull(TmpRs4!RecAmount), "", TmpRs4!RecAmount)
        MSHFlexGrid7.TextMatrix(0, 9) = Val(MSHFlexGrid7.TextMatrix(0, 9)) + Val(MSHFlexGrid7.TextMatrix(I, 9))
        MSHFlexGrid7.TextMatrix(I, 10) = IIf(IsNull(TmpRs4!RecTDS), "", TmpRs4!RecTDS)
        MSHFlexGrid7.TextMatrix(0, 10) = Val(MSHFlexGrid7.TextMatrix(0, 10)) + Val(MSHFlexGrid7.TextMatrix(I, 10))
        MSHFlexGrid7.TextMatrix(I, 11) = IIf(IsNull(TmpRs4!RecOtherDed), "", TmpRs4!RecOtherDed)
        MSHFlexGrid7.TextMatrix(0, 11) = Val(MSHFlexGrid7.TextMatrix(0, 11)) + Val(MSHFlexGrid7.TextMatrix(I, 11))
        
        MSHFlexGrid7.TextMatrix(I, 12) = IIf(IsNull(TmpRs4!CRID), "", TmpRs4!CRID)
        MSHFlexGrid7.TextMatrix(I, 13) = IIf(IsNull(TmpRs4!SM_Prefix), "", TmpRs4!SM_Prefix)
        
        TmpRs4.MoveNext
    Next
End If

'' for Advace Remaining

tmp = "Select TCRD.CRID,TCRD.SM_Prefix,ME.ExchangeType,TCRD.ClientIDRow,MC.ClientID,MC.ClientName,isnull(TCRD.AdvNAVCode,MC.NAVCode)as NAVCode,MC.RNAVCode,isnull(TCRD.Amount,0) as ExcessAmount"
tmp = tmp & " From Txn_NewCashReceiptClientDetail TCRD"
tmp = tmp & " Inner join Mst_ExchangeType ME on ME.ExchangeTypeID=TCRD.ExchangeTypeID"
tmp = tmp & " left join Mst_Client MC on MC.ClientIDRow=TCRD.ClientIDRow"
tmp = tmp & " Where TCRD.Amount > 0 And TCRD.EntryNo Is Not Null"
tmp = tmp & " Group by TCRD.CRID,TCRD.SM_Prefix,ExchangeType,ShowDP,TCRD.ClientIDRow,MC.ClientID,MC.ClientName,TCRD.Amount,TCRD.AdvNAVCode,MC.NAVCode,MC.RNAVCode"

''tmp = "Select sum(TCRD.Amount)Amount"
''tmp = tmp & " From Txn_NewCashReceiptClientDetail TCRD"
''tmp = tmp & " Where TCRD.Amount > 0 And TCRD.EntryNo Is Not Null"


Call Tmp2Table

Dim mRemAmt As Double
mRemAmt = 0
If TmpRs2.RecordCount > 0 Then
    MSHFlexGrid6.Rows = MSHFlexGrid6.Rows + TmpRs2.RecordCount
    For I = 2 To TmpRs2.RecordCount + 1
        mRemAmt = mRemAmt + TmpRs2!ExcessAmount
        MSHFlexGrid6.TextMatrix(I, 0) = IIf(IsNull(TmpRs2!CRID), "", TmpRs2!CRID)
        MSHFlexGrid6.TextMatrix(I, 1) = IIf(IsNull(TmpRs2!SM_Prefix), "", TmpRs2!SM_Prefix)
        MSHFlexGrid6.TextMatrix(I, 2) = IIf(IsNull(TmpRs2!ExchangeType), "", TmpRs2!ExchangeType)
        MSHFlexGrid6.TextMatrix(I, 3) = IIf(IsNull(TmpRs2!ClientName), "", TmpRs2!ClientName)
        MSHFlexGrid6.TextMatrix(I, 4) = IIf(IsNull(TmpRs2!NAVCode), "", TmpRs2!NAVCode)
        MSHFlexGrid6.TextMatrix(I, 5) = IIf(IsNull(TmpRs2!RNAVCode), "", TmpRs2!RNAVCode)
        MSHFlexGrid6.TextMatrix(I, 6) = IIf(IsNull(TmpRs2!ExcessAmount), "", TmpRs2!ExcessAmount)
        MSHFlexGrid6.TextMatrix(0, 6) = Val(MSHFlexGrid6.TextMatrix(0, 6)) + Val(MSHFlexGrid6.TextMatrix(I, 6))
        TmpRs2.MoveNext
    Next
End If

'''-Add to advane rem
''
''tmp = "Select TCRT.CRID,TCRT.SM_Prefix,sum(TCRT.RecAmount) RecAmt,sum(TCRT.RecTDS)RecTDS,sum(TCRT.recOtherDed)recOtherDed"
''tmp = tmp & " From  Txn_NewCashReceiptTIDetail TCRT"
''tmp = tmp & " inner join Txn_CRAdvanceAllocation CRAA on CRAA.CRAAID=TCRT.CRAAID"
''tmp = tmp & " Where CRAA.BatchID = 0"
''tmp = tmp & " Group by  TCRT.CRID,TCRT.SM_Prefix"

MSHFlexGrid3.TextMatrix(3, 1) = Format(MSHFlexGrid4.TextMatrix(0, 13) - mallocAmt, "#####0.00")

MSHFlexGrid3.TextMatrix(4, 1) = Format(Val(MSHFlexGrid5.TextMatrix(0, 10)) + Val(MSHFlexGrid5.TextMatrix(0, 11)) + Val(MSHFlexGrid5.TextMatrix(0, 12)), "######0.00") 'Format(MSHFlexGrid5.TextMatrix(0, 13), "#####0.00")

MSHFlexGrid3.TextMatrix(5, 1) = Format(mRemAmt - mallocAmt, "#####0.00")

MSHFlexGrid3.TextMatrix(6, 1) = Format(Val(MSHFlexGrid3.TextMatrix(4, 1)) + Val(MSHFlexGrid3.TextMatrix(3, 1)) - Val(MSHFlexGrid3.TextMatrix(5, 1)), "#####0.00")

''tmp = "Select isnull(sum(InvoiceAmount),0)InvoiceAmount"
''tmp = tmp & " From Txn_ClientLedger Where CustomerNo in (select distinct NAVCode From mst_Client)or"
''tmp = tmp & " CustomerNo in (select distinct RNAVCode From mst_Client) OR CustomerNo in (select distinct NAVCode From mst_Depository)"

tmp = "Select isnull(sum(InvoiceAmount),0)InvoiceAmount"
tmp = tmp & " From Txn_ClientLedger Where CustomerNo Like '%CSWC%' Or"
tmp = tmp & " CustomerNo Like '%CSWD%' OR CustomerNo Like '%CSWO%'"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.TextMatrix(7, 1) = Format(TmpRs.Fields(0), "#####0.00")
End If

MSHFlexGrid3.TextMatrix(8, 1) = Format(Val(MSHFlexGrid3.TextMatrix(6, 1)) - Val(MSHFlexGrid3.TextMatrix(7, 1)), "#####0.00")

End Sub

Private Sub Grid_Head4()
MSHFlexGrid5.Clear
MSHFlexGrid5.Rows = 3
MSHFlexGrid5.FixedRows = 2
MSHFlexGrid5.Cols = 14

MSHFlexGrid5.TextMatrix(1, 0) = "TI No"
MSHFlexGrid5.TextMatrix(1, 1) = "TI Date"
MSHFlexGrid5.TextMatrix(1, 2) = "TI Amount"
MSHFlexGrid5.TextMatrix(1, 3) = "TI ST Amount"
MSHFlexGrid5.TextMatrix(1, 4) = "State Name"
MSHFlexGrid5.TextMatrix(1, 5) = "Location Name"
MSHFlexGrid5.TextMatrix(1, 6) = "Post Invoice No"
MSHFlexGrid5.TextMatrix(1, 7) = "Entry No"
MSHFlexGrid5.TextMatrix(1, 8) = "NAV Code"
MSHFlexGrid5.TextMatrix(1, 9) = "Client Name"
MSHFlexGrid5.TextMatrix(1, 10) = "Rec Amount"
MSHFlexGrid5.TextMatrix(1, 11) = "Rec TDS"
MSHFlexGrid5.TextMatrix(1, 12) = "Rec Other Deduction"
MSHFlexGrid5.TextMatrix(1, 13) = "Balance Amount"

MSHFlexGrid5.ColWidth(0) = 1500
MSHFlexGrid5.ColWidth(1) = 1500
MSHFlexGrid5.ColWidth(2) = 1400
MSHFlexGrid5.ColWidth(3) = 1400
MSHFlexGrid5.ColWidth(4) = 2200
MSHFlexGrid5.ColWidth(5) = 2200
MSHFlexGrid5.ColWidth(6) = 1500
MSHFlexGrid5.ColWidth(7) = 1400
MSHFlexGrid5.ColWidth(8) = 1400
MSHFlexGrid5.ColWidth(9) = 2200
MSHFlexGrid5.ColWidth(10) = 1400
MSHFlexGrid5.ColWidth(11) = 1400
MSHFlexGrid5.ColWidth(12) = 1400
MSHFlexGrid5.ColWidth(13) = 1400
MSHFlexGrid5.ColAlignment(9) = vbLeftJustify

MSHFlexGrid5.RowHeight(1) = 600
MSHFlexGrid5.WordWrap = True
End Sub


Private Sub Grid_Head5()
MSHFlexGrid6.Clear
MSHFlexGrid6.Rows = 3
MSHFlexGrid6.FixedRows = 2
MSHFlexGrid6.Cols = 7

MSHFlexGrid6.TextMatrix(1, 0) = "Cash Receipt ID"
MSHFlexGrid6.TextMatrix(1, 1) = "SM Prefix"
MSHFlexGrid6.TextMatrix(1, 2) = "Exchange Type"
MSHFlexGrid6.TextMatrix(1, 3) = "Client Name"
MSHFlexGrid6.TextMatrix(1, 4) = "NAV Code"
MSHFlexGrid6.TextMatrix(1, 5) = "RNAV Code"
MSHFlexGrid6.TextMatrix(1, 6) = "Excess Amount"

MSHFlexGrid6.ColWidth(0) = 1500
MSHFlexGrid6.ColWidth(1) = 1500
MSHFlexGrid6.ColWidth(2) = 1400
MSHFlexGrid6.ColWidth(3) = 2500
MSHFlexGrid6.ColWidth(4) = 1500
MSHFlexGrid6.ColWidth(5) = 1500
MSHFlexGrid6.ColWidth(6) = 1500

MSHFlexGrid6.RowHeight(1) = 600
MSHFlexGrid6.WordWrap = True
End Sub

Private Sub Grid_Head6()
MSHFlexGrid7.Clear
MSHFlexGrid7.Rows = 3
MSHFlexGrid7.FixedRows = 2
MSHFlexGrid7.Cols = 14

MSHFlexGrid7.TextMatrix(1, 0) = "State Name"
MSHFlexGrid7.TextMatrix(1, 1) = "Location Name"
MSHFlexGrid7.TextMatrix(1, 2) = "TI No"
MSHFlexGrid7.TextMatrix(1, 3) = "TI Date"
MSHFlexGrid7.TextMatrix(1, 4) = "TI Amount"
MSHFlexGrid7.TextMatrix(1, 5) = "Post Invoice No"
MSHFlexGrid7.TextMatrix(1, 6) = "Entry No"
MSHFlexGrid7.TextMatrix(1, 7) = "NAV Code"
MSHFlexGrid7.TextMatrix(1, 8) = "Client Name"
MSHFlexGrid7.TextMatrix(1, 9) = "Rec. Amount"
MSHFlexGrid7.TextMatrix(1, 10) = "Rec. TDS"
MSHFlexGrid7.TextMatrix(1, 11) = "Rec Other Ded"
MSHFlexGrid7.TextMatrix(1, 12) = "CR ID"
MSHFlexGrid7.TextMatrix(1, 13) = "SM_Prefix"

MSHFlexGrid7.ColWidth(0) = 1500
MSHFlexGrid7.ColWidth(1) = 1500
MSHFlexGrid7.ColWidth(2) = 1400
MSHFlexGrid7.ColWidth(3) = 1400
MSHFlexGrid7.ColWidth(4) = 1500
MSHFlexGrid7.ColWidth(5) = 1500
MSHFlexGrid7.ColWidth(6) = 1500
MSHFlexGrid7.ColWidth(7) = 2000
MSHFlexGrid7.ColWidth(8) = 2500
MSHFlexGrid7.ColWidth(9) = 1500
MSHFlexGrid7.ColWidth(10) = 1500
MSHFlexGrid7.ColWidth(11) = 1500

MSHFlexGrid7.RowHeight(1) = 600
MSHFlexGrid7.WordWrap = True
End Sub

Private Sub Grid_Head9()
MSHFlexGrid10.Clear
MSHFlexGrid10.Rows = 3
MSHFlexGrid10.FixedRows = 2
MSHFlexGrid10.Cols = 12

MSHFlexGrid10.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid10.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid10.TextMatrix(1, 2) = "Commodity"
MSHFlexGrid10.TextMatrix(1, 3) = "Billing Date"
MSHFlexGrid10.TextMatrix(1, 4) = "Invoice No."
MSHFlexGrid10.TextMatrix(1, 5) = "NAV Code"
MSHFlexGrid10.TextMatrix(1, 6) = "BillingAmount"
MSHFlexGrid10.TextMatrix(1, 7) = "ActualAmount"
MSHFlexGrid10.TextMatrix(1, 8) = "ContractName"
MSHFlexGrid10.TextMatrix(1, 9) = "remarks"
MSHFlexGrid10.TextMatrix(1, 10) = "Flag"
MSHFlexGrid10.TextMatrix(1, 11) = "Month Year"

MSHFlexGrid10.ColWidth(0) = 1800
MSHFlexGrid10.ColWidth(1) = 3500
MSHFlexGrid10.ColWidth(2) = 1800
MSHFlexGrid10.ColWidth(3) = 1500
MSHFlexGrid10.ColWidth(4) = 1500
MSHFlexGrid10.ColWidth(5) = 1800
MSHFlexGrid10.ColWidth(6) = 1500
MSHFlexGrid10.ColWidth(7) = 1500
MSHFlexGrid10.ColWidth(8) = 2500
MSHFlexGrid10.ColWidth(9) = 2500
MSHFlexGrid10.ColWidth(10) = 1200
MSHFlexGrid10.ColWidth(11) = 1500

MSHFlexGrid10.RowHeight(1) = 600
MSHFlexGrid10.WordWrap = True
End Sub



Private Sub Grid_Head10()
MSHFlexGrid11.Clear
MSHFlexGrid11.Rows = 3
MSHFlexGrid11.FixedRows = 2
MSHFlexGrid11.Cols = 7

MSHFlexGrid11.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid11.TextMatrix(1, 1) = "Depository Name"
MSHFlexGrid11.TextMatrix(1, 2) = "DP NAV Code"
MSHFlexGrid11.TextMatrix(1, 3) = "Month Year"
MSHFlexGrid11.TextMatrix(1, 4) = "DP InoivceID"
MSHFlexGrid11.TextMatrix(1, 5) = "Amount"
MSHFlexGrid11.TextMatrix(1, 6) = "BatchID"

MSHFlexGrid11.ColWidth(0) = 1800
MSHFlexGrid11.ColWidth(1) = 3500
MSHFlexGrid11.ColWidth(2) = 1500
MSHFlexGrid11.ColWidth(3) = 1500
MSHFlexGrid11.ColWidth(4) = 1500
MSHFlexGrid11.ColWidth(5) = 1500
MSHFlexGrid11.ColWidth(6) = 1200

MSHFlexGrid11.RowHeight(1) = 600
MSHFlexGrid11.WordWrap = True
End Sub

Private Sub Grid_Head11()
MSHFlexGrid12.Clear
MSHFlexGrid12.Rows = 3
MSHFlexGrid12.FixedRows = 2
MSHFlexGrid12.Cols = 14

MSHFlexGrid12.TextMatrix(1, 0) = "CR ID"
MSHFlexGrid12.TextMatrix(1, 1) = "CR Date"
MSHFlexGrid12.TextMatrix(1, 2) = "Location Name"
MSHFlexGrid12.TextMatrix(1, 3) = "InstrumentType"
MSHFlexGrid12.TextMatrix(1, 4) = "DrawnOn"
MSHFlexGrid12.TextMatrix(1, 5) = "BranchName"
MSHFlexGrid12.TextMatrix(1, 6) = "Instrument No"
MSHFlexGrid12.TextMatrix(1, 7) = "Instrument Date"
MSHFlexGrid12.TextMatrix(1, 8) = "Instrument Amount"
MSHFlexGrid12.TextMatrix(1, 9) = "Received From"
MSHFlexGrid12.TextMatrix(1, 10) = "Remark"
MSHFlexGrid12.TextMatrix(1, 11) = "BatchID"
MSHFlexGrid12.TextMatrix(1, 12) = "Deposit Date"
MSHFlexGrid12.TextMatrix(1, 13) = "FAG Posting Date"

MSHFlexGrid12.ColWidth(0) = 1200
MSHFlexGrid12.ColWidth(1) = 1500
MSHFlexGrid12.ColWidth(2) = 1800
MSHFlexGrid12.ColWidth(3) = 1800
MSHFlexGrid12.ColWidth(4) = 1500
MSHFlexGrid12.ColWidth(5) = 2500
MSHFlexGrid12.ColWidth(6) = 1500
MSHFlexGrid12.ColWidth(7) = 1500
MSHFlexGrid12.ColWidth(8) = 1500
MSHFlexGrid12.ColWidth(9) = 2500
MSHFlexGrid12.ColWidth(10) = 2500
MSHFlexGrid12.ColWidth(11) = 1200
MSHFlexGrid12.ColWidth(12) = 1400
MSHFlexGrid12.ColWidth(13) = 1400

MSHFlexGrid12.RowHeight(1) = 600
MSHFlexGrid12.WordWrap = True
End Sub

Private Sub Grid_Head12()
MSHFlexGrid13.Clear
MSHFlexGrid13.Rows = 3
MSHFlexGrid13.FixedRows = 2
MSHFlexGrid13.Cols = 12

MSHFlexGrid13.TextMatrix(1, 0) = "CR ID"
MSHFlexGrid13.TextMatrix(1, 1) = "CR Date"
MSHFlexGrid13.TextMatrix(1, 2) = "Location Name"
MSHFlexGrid13.TextMatrix(1, 3) = "InstrumentType"
MSHFlexGrid13.TextMatrix(1, 4) = "DrawnOn"
MSHFlexGrid13.TextMatrix(1, 5) = "BranchName"
MSHFlexGrid13.TextMatrix(1, 6) = "Instrument No"
MSHFlexGrid13.TextMatrix(1, 7) = "Instrument Date"
MSHFlexGrid13.TextMatrix(1, 8) = "Instrument Amount"
MSHFlexGrid13.TextMatrix(1, 9) = "Received From"
MSHFlexGrid13.TextMatrix(1, 10) = "Remark"
MSHFlexGrid13.TextMatrix(1, 11) = "BatchID"

MSHFlexGrid13.ColWidth(0) = 1200
MSHFlexGrid13.ColWidth(1) = 1500
MSHFlexGrid13.ColWidth(2) = 1800
MSHFlexGrid13.ColWidth(3) = 1500
MSHFlexGrid13.ColWidth(4) = 1800
MSHFlexGrid13.ColWidth(5) = 2500
MSHFlexGrid13.ColWidth(6) = 1500
MSHFlexGrid13.ColWidth(7) = 1500
MSHFlexGrid13.ColWidth(8) = 1500
MSHFlexGrid13.ColWidth(9) = 2500
MSHFlexGrid13.ColWidth(10) = 2500
MSHFlexGrid13.ColWidth(11) = 1200

MSHFlexGrid13.RowHeight(1) = 600
MSHFlexGrid13.WordWrap = True
End Sub

Private Sub Grid_Head13()
MSHFlexGrid14.Clear
MSHFlexGrid14.Rows = 3
MSHFlexGrid14.FixedRows = 2
MSHFlexGrid14.Cols = 7

MSHFlexGrid14.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid14.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid14.TextMatrix(1, 2) = "AdjustmentDate"
MSHFlexGrid14.TextMatrix(1, 3) = "Rec TDS"
MSHFlexGrid14.TextMatrix(1, 4) = "Rec Other Ded"
MSHFlexGrid14.TextMatrix(1, 5) = "Remark"
MSHFlexGrid14.TextMatrix(1, 6) = "BatchID"

MSHFlexGrid14.ColWidth(0) = 1800
MSHFlexGrid14.ColWidth(1) = 3500
MSHFlexGrid14.ColWidth(2) = 1500
MSHFlexGrid14.ColWidth(3) = 1500
MSHFlexGrid14.ColWidth(4) = 1500
MSHFlexGrid14.ColWidth(5) = 2500
MSHFlexGrid14.ColWidth(6) = 1500

MSHFlexGrid14.RowHeight(1) = 500
MSHFlexGrid14.WordWrap = True
End Sub

Private Sub Grid_Head8()
MSHFlexGrid9.Clear
MSHFlexGrid9.Rows = 3
MSHFlexGrid9.FixedRows = 2
MSHFlexGrid9.Cols = 8

MSHFlexGrid9.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid9.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid9.TextMatrix(1, 2) = "Exchange Type"
MSHFlexGrid9.TextMatrix(1, 3) = "Invoice Date"
MSHFlexGrid9.TextMatrix(1, 4) = "Invoice No."
MSHFlexGrid9.TextMatrix(1, 5) = "NAV Code"
MSHFlexGrid9.TextMatrix(1, 6) = "Amount"
MSHFlexGrid9.TextMatrix(1, 7) = "BatchID"

MSHFlexGrid9.ColWidth(0) = 1800
MSHFlexGrid9.ColWidth(1) = 3500
MSHFlexGrid9.ColWidth(2) = 1800
MSHFlexGrid9.ColWidth(3) = 1500
MSHFlexGrid9.ColWidth(4) = 1500
MSHFlexGrid9.ColWidth(5) = 1800
MSHFlexGrid9.ColWidth(6) = 1500
MSHFlexGrid9.ColWidth(7) = 1200

MSHFlexGrid9.RowHeight(1) = 600
MSHFlexGrid9.WordWrap = True
End Sub

Private Sub Grid_Head7()
MSHFlexGrid8.Clear
MSHFlexGrid8.Rows = 3
MSHFlexGrid8.FixedRows = 2
MSHFlexGrid8.Cols = 8

MSHFlexGrid8.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid8.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid8.TextMatrix(1, 2) = "Exchange Type"
MSHFlexGrid8.TextMatrix(1, 3) = "Invoice Date"
MSHFlexGrid8.TextMatrix(1, 4) = "Invoice No."
MSHFlexGrid8.TextMatrix(1, 5) = "NAV Code"
MSHFlexGrid8.TextMatrix(1, 6) = "Amount"
MSHFlexGrid8.TextMatrix(1, 7) = "BatchID"

MSHFlexGrid8.ColWidth(0) = 1800
MSHFlexGrid8.ColWidth(1) = 3500
MSHFlexGrid8.ColWidth(2) = 1800
MSHFlexGrid8.ColWidth(3) = 1500
MSHFlexGrid8.ColWidth(4) = 1500
MSHFlexGrid8.ColWidth(5) = 1800
MSHFlexGrid8.ColWidth(6) = 1500
MSHFlexGrid8.ColWidth(7) = 1200

MSHFlexGrid8.RowHeight(1) = 600
MSHFlexGrid8.WordWrap = True
End Sub


Private Sub Grid_Head15()
MSHFlexGrid15.Clear
MSHFlexGrid15.Rows = 3
MSHFlexGrid15.FixedRows = 2
MSHFlexGrid15.Cols = 14

MSHFlexGrid15.TextMatrix(1, 0) = "CR ID"
MSHFlexGrid15.TextMatrix(1, 1) = "CR Date"
MSHFlexGrid15.TextMatrix(1, 2) = "Location Name"
MSHFlexGrid15.TextMatrix(1, 3) = "InstrumentType"
MSHFlexGrid15.TextMatrix(1, 4) = "DrawnOn"
MSHFlexGrid15.TextMatrix(1, 5) = "BranchName"
MSHFlexGrid15.TextMatrix(1, 6) = "Instrument No"
MSHFlexGrid15.TextMatrix(1, 7) = "Instrument Date"
MSHFlexGrid15.TextMatrix(1, 8) = "Instrument Amount"
MSHFlexGrid15.TextMatrix(1, 9) = "Received From"
MSHFlexGrid15.TextMatrix(1, 10) = "Remark"
MSHFlexGrid15.TextMatrix(1, 11) = "BatchID"
MSHFlexGrid15.TextMatrix(1, 12) = "Deposit Date"
MSHFlexGrid15.TextMatrix(1, 13) = "FAG Posting Date"

MSHFlexGrid15.ColWidth(0) = 1200
MSHFlexGrid15.ColWidth(1) = 1500
MSHFlexGrid15.ColWidth(2) = 1800
MSHFlexGrid15.ColWidth(3) = 1800
MSHFlexGrid15.ColWidth(4) = 1500
MSHFlexGrid15.ColWidth(5) = 2500
MSHFlexGrid15.ColWidth(6) = 1500
MSHFlexGrid15.ColWidth(7) = 1500
MSHFlexGrid15.ColWidth(8) = 1500
MSHFlexGrid15.ColWidth(9) = 2500
MSHFlexGrid15.ColWidth(10) = 2500
MSHFlexGrid15.ColWidth(11) = 1200
MSHFlexGrid15.ColWidth(12) = 1400
MSHFlexGrid15.ColWidth(13) = 1400

MSHFlexGrid15.RowHeight(1) = 600
MSHFlexGrid15.WordWrap = True
End Sub

Private Sub Grid_Head16()
MSHFlexGrid16.Clear
MSHFlexGrid16.Rows = 3
MSHFlexGrid16.FixedRows = 2
MSHFlexGrid16.Cols = 14

MSHFlexGrid16.TextMatrix(1, 0) = "CR ID"
MSHFlexGrid16.TextMatrix(1, 1) = "CR Date"
MSHFlexGrid16.TextMatrix(1, 2) = "Location Name"
MSHFlexGrid16.TextMatrix(1, 3) = "InstrumentType"
MSHFlexGrid16.TextMatrix(1, 4) = "DrawnOn"
MSHFlexGrid16.TextMatrix(1, 5) = "BranchName"
MSHFlexGrid16.TextMatrix(1, 6) = "Instrument No"
MSHFlexGrid16.TextMatrix(1, 7) = "Instrument Date"
MSHFlexGrid16.TextMatrix(1, 8) = "Instrument Amount"
MSHFlexGrid16.TextMatrix(1, 9) = "Received From"
MSHFlexGrid16.TextMatrix(1, 10) = "Remark"
MSHFlexGrid16.TextMatrix(1, 11) = "BatchID"
MSHFlexGrid16.TextMatrix(1, 12) = "Deposit Date"
MSHFlexGrid16.TextMatrix(1, 13) = "FAG Posting Date"

MSHFlexGrid16.ColWidth(0) = 1200
MSHFlexGrid16.ColWidth(1) = 1500
MSHFlexGrid16.ColWidth(2) = 1800
MSHFlexGrid16.ColWidth(3) = 1800
MSHFlexGrid16.ColWidth(4) = 1500
MSHFlexGrid16.ColWidth(5) = 2500
MSHFlexGrid16.ColWidth(6) = 1500
MSHFlexGrid16.ColWidth(7) = 1500
MSHFlexGrid16.ColWidth(8) = 1500
MSHFlexGrid16.ColWidth(9) = 2500
MSHFlexGrid16.ColWidth(10) = 2500
MSHFlexGrid16.ColWidth(11) = 1200
MSHFlexGrid16.ColWidth(12) = 1400
MSHFlexGrid16.ColWidth(13) = 1400

MSHFlexGrid16.RowHeight(1) = 600
MSHFlexGrid16.WordWrap = True
End Sub


Private Sub Grid_Head17()
MSHFlexGrid17.Clear
MSHFlexGrid17.Rows = 3
MSHFlexGrid17.FixedRows = 2
MSHFlexGrid17.Cols = 14

MSHFlexGrid17.TextMatrix(1, 0) = "CR ID"
MSHFlexGrid17.TextMatrix(1, 1) = "CR Date"
MSHFlexGrid17.TextMatrix(1, 2) = "Location Name"
MSHFlexGrid17.TextMatrix(1, 3) = "InstrumentType"
MSHFlexGrid17.TextMatrix(1, 4) = "DrawnOn"
MSHFlexGrid17.TextMatrix(1, 5) = "BranchName"
MSHFlexGrid17.TextMatrix(1, 6) = "Instrument No"
MSHFlexGrid17.TextMatrix(1, 7) = "Instrument Date"
MSHFlexGrid17.TextMatrix(1, 8) = "Instrument Amount"
MSHFlexGrid17.TextMatrix(1, 9) = "Received From"
MSHFlexGrid17.TextMatrix(1, 10) = "Remark"
MSHFlexGrid17.TextMatrix(1, 11) = "BatchID"
MSHFlexGrid17.TextMatrix(1, 12) = "Deposit Date"
MSHFlexGrid17.TextMatrix(1, 13) = "FAG Posting Date"

MSHFlexGrid17.ColWidth(0) = 1200
MSHFlexGrid17.ColWidth(1) = 1500
MSHFlexGrid17.ColWidth(2) = 1800
MSHFlexGrid17.ColWidth(3) = 1800
MSHFlexGrid17.ColWidth(4) = 1500
MSHFlexGrid17.ColWidth(5) = 2500
MSHFlexGrid17.ColWidth(6) = 1500
MSHFlexGrid17.ColWidth(7) = 1500
MSHFlexGrid17.ColWidth(8) = 1500
MSHFlexGrid17.ColWidth(9) = 2500
MSHFlexGrid17.ColWidth(10) = 2500
MSHFlexGrid17.ColWidth(11) = 1200
MSHFlexGrid17.ColWidth(12) = 1400
MSHFlexGrid17.ColWidth(13) = 1400

MSHFlexGrid17.RowHeight(1) = 600
MSHFlexGrid17.WordWrap = True
End Sub


Private Sub Grid_Head18()
MSHFlexGrid18.Clear
MSHFlexGrid18.Rows = 3
MSHFlexGrid18.FixedRows = 2
MSHFlexGrid18.Cols = 14

MSHFlexGrid18.TextMatrix(1, 0) = "CR ID"
MSHFlexGrid18.TextMatrix(1, 1) = "CR Date"
MSHFlexGrid18.TextMatrix(1, 2) = "Location Name"
MSHFlexGrid18.TextMatrix(1, 3) = "InstrumentType"
MSHFlexGrid18.TextMatrix(1, 4) = "DrawnOn"
MSHFlexGrid18.TextMatrix(1, 5) = "BranchName"
MSHFlexGrid18.TextMatrix(1, 6) = "Instrument No"
MSHFlexGrid18.TextMatrix(1, 7) = "Instrument Date"
MSHFlexGrid18.TextMatrix(1, 8) = "Instrument Amount"
MSHFlexGrid18.TextMatrix(1, 9) = "Received From"
MSHFlexGrid18.TextMatrix(1, 10) = "Remark"
MSHFlexGrid18.TextMatrix(1, 11) = "BatchID"
MSHFlexGrid18.TextMatrix(1, 12) = "Deposit Date"
MSHFlexGrid18.TextMatrix(1, 13) = "FAG Posting Date"

MSHFlexGrid18.ColWidth(0) = 1200
MSHFlexGrid18.ColWidth(1) = 1500
MSHFlexGrid18.ColWidth(2) = 1800
MSHFlexGrid18.ColWidth(3) = 1800
MSHFlexGrid18.ColWidth(4) = 1500
MSHFlexGrid18.ColWidth(5) = 2500
MSHFlexGrid18.ColWidth(6) = 1500
MSHFlexGrid18.ColWidth(7) = 1500
MSHFlexGrid18.ColWidth(8) = 1500
MSHFlexGrid18.ColWidth(9) = 2500
MSHFlexGrid18.ColWidth(10) = 2500
MSHFlexGrid18.ColWidth(11) = 1200
MSHFlexGrid18.ColWidth(12) = 1400
MSHFlexGrid18.ColWidth(13) = 1400

MSHFlexGrid18.RowHeight(1) = 600
MSHFlexGrid18.WordWrap = True
End Sub


Private Sub Grid_Head19()
MSHFlexGrid19.Clear
MSHFlexGrid19.Rows = 3
MSHFlexGrid19.FixedRows = 2
MSHFlexGrid19.Cols = 14

MSHFlexGrid19.TextMatrix(1, 0) = "CR ID"
MSHFlexGrid19.TextMatrix(1, 1) = "CR Date"
MSHFlexGrid19.TextMatrix(1, 2) = "Location Name"
MSHFlexGrid19.TextMatrix(1, 3) = "InstrumentType"
MSHFlexGrid19.TextMatrix(1, 4) = "DrawnOn"
MSHFlexGrid19.TextMatrix(1, 5) = "BranchName"
MSHFlexGrid19.TextMatrix(1, 6) = "Instrument No"
MSHFlexGrid19.TextMatrix(1, 7) = "Instrument Date"
MSHFlexGrid19.TextMatrix(1, 8) = "Instrument Amount"
MSHFlexGrid19.TextMatrix(1, 9) = "Received From"
MSHFlexGrid19.TextMatrix(1, 10) = "Remark"
MSHFlexGrid19.TextMatrix(1, 11) = "BatchID"
MSHFlexGrid19.TextMatrix(1, 12) = "Deposit Date"
MSHFlexGrid19.TextMatrix(1, 13) = "FAG Posting Date"

MSHFlexGrid19.ColWidth(0) = 1200
MSHFlexGrid19.ColWidth(1) = 1500
MSHFlexGrid19.ColWidth(2) = 1800
MSHFlexGrid19.ColWidth(3) = 1800
MSHFlexGrid19.ColWidth(4) = 1500
MSHFlexGrid19.ColWidth(5) = 2500
MSHFlexGrid19.ColWidth(6) = 1500
MSHFlexGrid19.ColWidth(7) = 1500
MSHFlexGrid19.ColWidth(8) = 1500
MSHFlexGrid19.ColWidth(9) = 2500
MSHFlexGrid19.ColWidth(10) = 2500
MSHFlexGrid19.ColWidth(11) = 1200
MSHFlexGrid19.ColWidth(12) = 1400
MSHFlexGrid19.ColWidth(13) = 1400

MSHFlexGrid19.RowHeight(1) = 600
MSHFlexGrid19.WordWrap = True
End Sub


Private Sub Grid_Head22()
MSHFlexGrid22.Clear
MSHFlexGrid22.Rows = 3
MSHFlexGrid22.FixedRows = 2
MSHFlexGrid22.Cols = 8

MSHFlexGrid22.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid22.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid22.TextMatrix(1, 2) = "Exchange Type"
MSHFlexGrid22.TextMatrix(1, 3) = "Invoice Date"
MSHFlexGrid22.TextMatrix(1, 4) = "Invoice No."
MSHFlexGrid22.TextMatrix(1, 5) = "NAV Code"
MSHFlexGrid22.TextMatrix(1, 6) = "Amount"
MSHFlexGrid22.TextMatrix(1, 7) = "BatchID"

MSHFlexGrid22.ColWidth(0) = 1800
MSHFlexGrid22.ColWidth(1) = 3500
MSHFlexGrid22.ColWidth(2) = 1800
MSHFlexGrid22.ColWidth(3) = 1500
MSHFlexGrid22.ColWidth(4) = 1500
MSHFlexGrid22.ColWidth(5) = 1800
MSHFlexGrid22.ColWidth(6) = 1500
MSHFlexGrid22.ColWidth(7) = 1200

MSHFlexGrid22.RowHeight(1) = 600
MSHFlexGrid22.WordWrap = True
End Sub

Private Sub Grid_Head23()
MSHFlexGrid23.Clear
MSHFlexGrid23.Rows = 3
MSHFlexGrid23.FixedRows = 2
MSHFlexGrid23.Cols = 8

MSHFlexGrid23.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid23.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid23.TextMatrix(1, 2) = "Exchange Type"
MSHFlexGrid23.TextMatrix(1, 3) = "Invoice Date"
MSHFlexGrid23.TextMatrix(1, 4) = "Invoice No."
MSHFlexGrid23.TextMatrix(1, 5) = "NAV Code"
MSHFlexGrid23.TextMatrix(1, 6) = "Amount"
MSHFlexGrid23.TextMatrix(1, 7) = "BatchID"

MSHFlexGrid23.ColWidth(0) = 1800
MSHFlexGrid23.ColWidth(1) = 3500
MSHFlexGrid23.ColWidth(2) = 1800
MSHFlexGrid23.ColWidth(3) = 1500
MSHFlexGrid23.ColWidth(4) = 1500
MSHFlexGrid23.ColWidth(5) = 1800
MSHFlexGrid23.ColWidth(6) = 1500
MSHFlexGrid23.ColWidth(7) = 1200

MSHFlexGrid23.RowHeight(1) = 600
MSHFlexGrid23.WordWrap = True
End Sub

Private Sub Grid_Head24()
MSHFlexGrid24.Clear
MSHFlexGrid24.Rows = 3
MSHFlexGrid24.FixedRows = 2
MSHFlexGrid24.Cols = 8

MSHFlexGrid24.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid24.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid24.TextMatrix(1, 2) = "Exchange Type"
MSHFlexGrid24.TextMatrix(1, 3) = "Invoice Date"
MSHFlexGrid24.TextMatrix(1, 4) = "Invoice No."
MSHFlexGrid24.TextMatrix(1, 5) = "NAV Code"
MSHFlexGrid24.TextMatrix(1, 6) = "Amount"
MSHFlexGrid24.TextMatrix(1, 7) = "BatchID"

MSHFlexGrid24.ColWidth(0) = 1800
MSHFlexGrid24.ColWidth(1) = 3500
MSHFlexGrid24.ColWidth(2) = 1800
MSHFlexGrid24.ColWidth(3) = 1500
MSHFlexGrid24.ColWidth(4) = 1500
MSHFlexGrid24.ColWidth(5) = 1800
MSHFlexGrid24.ColWidth(6) = 1500
MSHFlexGrid24.ColWidth(7) = 1200

MSHFlexGrid24.RowHeight(1) = 600
MSHFlexGrid24.WordWrap = True
End Sub


Private Sub Grid_Head25()
MSHFlexGrid25.Clear
MSHFlexGrid25.Rows = 3
MSHFlexGrid25.FixedRows = 2
MSHFlexGrid25.Cols = 12

MSHFlexGrid25.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid25.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid25.TextMatrix(1, 2) = "Commodity"
MSHFlexGrid25.TextMatrix(1, 3) = "Billing Date"
MSHFlexGrid25.TextMatrix(1, 4) = "Invoice No."
MSHFlexGrid25.TextMatrix(1, 5) = "NAV Code"
MSHFlexGrid25.TextMatrix(1, 6) = "BillingAmount"
MSHFlexGrid25.TextMatrix(1, 7) = "ActualAmount"
MSHFlexGrid25.TextMatrix(1, 8) = "ContractName"
MSHFlexGrid25.TextMatrix(1, 9) = "remarks"
MSHFlexGrid25.TextMatrix(1, 10) = "Flag"
MSHFlexGrid25.TextMatrix(1, 11) = "Month Year"

MSHFlexGrid25.ColWidth(0) = 1800
MSHFlexGrid25.ColWidth(1) = 3500
MSHFlexGrid25.ColWidth(2) = 1800
MSHFlexGrid25.ColWidth(3) = 1500
MSHFlexGrid25.ColWidth(4) = 1500
MSHFlexGrid25.ColWidth(5) = 1800
MSHFlexGrid25.ColWidth(6) = 1500
MSHFlexGrid25.ColWidth(7) = 1500
MSHFlexGrid25.ColWidth(8) = 2500
MSHFlexGrid25.ColWidth(9) = 2500
MSHFlexGrid25.ColWidth(10) = 1200
MSHFlexGrid25.ColWidth(11) = 1500

MSHFlexGrid25.RowHeight(1) = 600
MSHFlexGrid25.WordWrap = True
End Sub


Private Sub Grid_Head26()
MSHFlexGrid26.Clear
MSHFlexGrid26.Rows = 3
MSHFlexGrid26.FixedRows = 2
MSHFlexGrid26.Cols = 12

MSHFlexGrid26.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid26.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid26.TextMatrix(1, 2) = "Commodity"
MSHFlexGrid26.TextMatrix(1, 3) = "Billing Date"
MSHFlexGrid26.TextMatrix(1, 4) = "Invoice No."
MSHFlexGrid26.TextMatrix(1, 5) = "NAV Code"
MSHFlexGrid26.TextMatrix(1, 6) = "BillingAmount"
MSHFlexGrid26.TextMatrix(1, 7) = "ActualAmount"
MSHFlexGrid26.TextMatrix(1, 8) = "ContractName"
MSHFlexGrid26.TextMatrix(1, 9) = "remarks"
MSHFlexGrid26.TextMatrix(1, 10) = "Flag"
MSHFlexGrid26.TextMatrix(1, 11) = "Month Year"

MSHFlexGrid26.ColWidth(0) = 1800
MSHFlexGrid26.ColWidth(1) = 3500
MSHFlexGrid26.ColWidth(2) = 1800
MSHFlexGrid26.ColWidth(3) = 1500
MSHFlexGrid26.ColWidth(4) = 1500
MSHFlexGrid26.ColWidth(5) = 1800
MSHFlexGrid26.ColWidth(6) = 1500
MSHFlexGrid26.ColWidth(7) = 1500
MSHFlexGrid26.ColWidth(8) = 2500
MSHFlexGrid26.ColWidth(9) = 2500
MSHFlexGrid26.ColWidth(10) = 1200
MSHFlexGrid26.ColWidth(11) = 1500

MSHFlexGrid26.RowHeight(1) = 600
MSHFlexGrid26.WordWrap = True
End Sub

Private Sub Grid_Head27()
MSHFlexGrid27.Clear
MSHFlexGrid27.Rows = 3
MSHFlexGrid27.FixedRows = 2
MSHFlexGrid27.Cols = 12

MSHFlexGrid27.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid27.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid27.TextMatrix(1, 2) = "Commodity"
MSHFlexGrid27.TextMatrix(1, 3) = "Billing Date"
MSHFlexGrid27.TextMatrix(1, 4) = "Invoice No."
MSHFlexGrid27.TextMatrix(1, 5) = "NAV Code"
MSHFlexGrid27.TextMatrix(1, 6) = "BillingAmount"
MSHFlexGrid27.TextMatrix(1, 7) = "ActualAmount"
MSHFlexGrid27.TextMatrix(1, 8) = "ContractName"
MSHFlexGrid27.TextMatrix(1, 9) = "remarks"
MSHFlexGrid27.TextMatrix(1, 10) = "Flag"
MSHFlexGrid27.TextMatrix(1, 11) = "Month Year"

MSHFlexGrid27.ColWidth(0) = 1800
MSHFlexGrid27.ColWidth(1) = 3500
MSHFlexGrid27.ColWidth(2) = 1800
MSHFlexGrid27.ColWidth(3) = 1500
MSHFlexGrid27.ColWidth(4) = 1500
MSHFlexGrid27.ColWidth(5) = 1800
MSHFlexGrid27.ColWidth(6) = 1500
MSHFlexGrid27.ColWidth(7) = 1500
MSHFlexGrid27.ColWidth(8) = 2500
MSHFlexGrid27.ColWidth(9) = 2500
MSHFlexGrid27.ColWidth(10) = 1200
MSHFlexGrid27.ColWidth(11) = 1500

MSHFlexGrid27.RowHeight(1) = 600
MSHFlexGrid27.WordWrap = True
End Sub

Private Sub Grid_Head28()
MSHFlexGrid28.Clear
MSHFlexGrid28.Rows = 3
MSHFlexGrid28.FixedRows = 2
MSHFlexGrid28.Cols = 7

MSHFlexGrid28.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid28.TextMatrix(1, 1) = "Depository Name"
MSHFlexGrid28.TextMatrix(1, 2) = "DP NAV Code"
MSHFlexGrid28.TextMatrix(1, 3) = "Month Year"
MSHFlexGrid28.TextMatrix(1, 4) = "DP InoivceID"
MSHFlexGrid28.TextMatrix(1, 5) = "Amount"
MSHFlexGrid28.TextMatrix(1, 6) = "BatchID"

MSHFlexGrid28.ColWidth(0) = 1800
MSHFlexGrid28.ColWidth(1) = 3500
MSHFlexGrid28.ColWidth(2) = 1500
MSHFlexGrid28.ColWidth(3) = 1500
MSHFlexGrid28.ColWidth(4) = 1500
MSHFlexGrid28.ColWidth(5) = 1500
MSHFlexGrid28.ColWidth(6) = 1200

MSHFlexGrid28.RowHeight(1) = 600
MSHFlexGrid28.WordWrap = True
End Sub

Private Sub Grid_Head31()
MSHFlexGrid31.Clear
MSHFlexGrid31.Rows = 3
MSHFlexGrid31.FixedRows = 2
MSHFlexGrid31.Cols = 7

MSHFlexGrid31.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid31.TextMatrix(1, 1) = "Depository Name"
MSHFlexGrid31.TextMatrix(1, 2) = "DP NAV Code"
MSHFlexGrid31.TextMatrix(1, 3) = "Month Year"
MSHFlexGrid31.TextMatrix(1, 4) = "DP InoivceID"
MSHFlexGrid31.TextMatrix(1, 5) = "Amount"
MSHFlexGrid31.TextMatrix(1, 6) = "BatchID"

MSHFlexGrid31.ColWidth(0) = 1800
MSHFlexGrid31.ColWidth(1) = 3500
MSHFlexGrid31.ColWidth(2) = 1500
MSHFlexGrid31.ColWidth(3) = 1500
MSHFlexGrid31.ColWidth(4) = 1500
MSHFlexGrid31.ColWidth(5) = 1500
MSHFlexGrid31.ColWidth(6) = 1200

MSHFlexGrid31.RowHeight(1) = 600
MSHFlexGrid31.WordWrap = True
End Sub


Private Sub Grid_Head32()
MSHFlexGrid32.Clear
MSHFlexGrid32.Rows = 3
MSHFlexGrid32.FixedRows = 2
MSHFlexGrid32.Cols = 12

MSHFlexGrid32.TextMatrix(1, 0) = "CR ID"
MSHFlexGrid32.TextMatrix(1, 1) = "CR Date"
MSHFlexGrid32.TextMatrix(1, 2) = "Location Name"
MSHFlexGrid32.TextMatrix(1, 3) = "InstrumentType"
MSHFlexGrid32.TextMatrix(1, 4) = "DrawnOn"
MSHFlexGrid32.TextMatrix(1, 5) = "BranchName"
MSHFlexGrid32.TextMatrix(1, 6) = "Instrument No"
MSHFlexGrid32.TextMatrix(1, 7) = "Instrument Date"
MSHFlexGrid32.TextMatrix(1, 8) = "Instrument Amount"
MSHFlexGrid32.TextMatrix(1, 9) = "Received From"
MSHFlexGrid32.TextMatrix(1, 10) = "Remark"
MSHFlexGrid32.TextMatrix(1, 11) = "BatchID"

MSHFlexGrid32.ColWidth(0) = 1200
MSHFlexGrid32.ColWidth(1) = 1500
MSHFlexGrid32.ColWidth(2) = 1800
MSHFlexGrid32.ColWidth(3) = 1500
MSHFlexGrid32.ColWidth(4) = 1800
MSHFlexGrid32.ColWidth(5) = 2500
MSHFlexGrid32.ColWidth(6) = 1500
MSHFlexGrid32.ColWidth(7) = 1500
MSHFlexGrid32.ColWidth(8) = 1500
MSHFlexGrid32.ColWidth(9) = 2500
MSHFlexGrid32.ColWidth(10) = 2500
MSHFlexGrid32.ColWidth(11) = 1200

MSHFlexGrid32.RowHeight(1) = 600
MSHFlexGrid32.WordWrap = True
End Sub


Private Sub Grid_Head33()
MSHFlexGrid33.Clear
MSHFlexGrid33.Rows = 3
MSHFlexGrid33.FixedRows = 2
MSHFlexGrid33.Cols = 12

MSHFlexGrid33.TextMatrix(1, 0) = "CR ID"
MSHFlexGrid33.TextMatrix(1, 1) = "CR Date"
MSHFlexGrid33.TextMatrix(1, 2) = "Location Name"
MSHFlexGrid33.TextMatrix(1, 3) = "InstrumentType"
MSHFlexGrid33.TextMatrix(1, 4) = "DrawnOn"
MSHFlexGrid33.TextMatrix(1, 5) = "BranchName"
MSHFlexGrid33.TextMatrix(1, 6) = "Instrument No"
MSHFlexGrid33.TextMatrix(1, 7) = "Instrument Date"
MSHFlexGrid33.TextMatrix(1, 8) = "Instrument Amount"
MSHFlexGrid33.TextMatrix(1, 9) = "Received From"
MSHFlexGrid33.TextMatrix(1, 10) = "Remark"
MSHFlexGrid33.TextMatrix(1, 11) = "BatchID"

MSHFlexGrid33.ColWidth(0) = 1200
MSHFlexGrid33.ColWidth(1) = 1500
MSHFlexGrid33.ColWidth(2) = 1800
MSHFlexGrid33.ColWidth(3) = 1500
MSHFlexGrid33.ColWidth(4) = 1800
MSHFlexGrid33.ColWidth(5) = 2500
MSHFlexGrid33.ColWidth(6) = 1500
MSHFlexGrid33.ColWidth(7) = 1500
MSHFlexGrid33.ColWidth(8) = 1500
MSHFlexGrid33.ColWidth(9) = 2500
MSHFlexGrid33.ColWidth(10) = 2500
MSHFlexGrid33.ColWidth(11) = 1200

MSHFlexGrid33.RowHeight(1) = 600
MSHFlexGrid33.WordWrap = True
End Sub

Private Sub Grid_Head34()
MSHFlexGrid34.Clear
MSHFlexGrid34.Rows = 3
MSHFlexGrid34.FixedRows = 2
MSHFlexGrid34.Cols = 7

MSHFlexGrid34.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid34.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid34.TextMatrix(1, 2) = "AdjustmentDate"
MSHFlexGrid34.TextMatrix(1, 3) = "Rec TDS"
MSHFlexGrid34.TextMatrix(1, 4) = "Rec Other Ded"
MSHFlexGrid34.TextMatrix(1, 5) = "Remark"
MSHFlexGrid34.TextMatrix(1, 6) = "BatchID"

MSHFlexGrid34.ColWidth(0) = 1800
MSHFlexGrid34.ColWidth(1) = 3500
MSHFlexGrid34.ColWidth(2) = 1500
MSHFlexGrid34.ColWidth(3) = 1500
MSHFlexGrid34.ColWidth(4) = 1500
MSHFlexGrid34.ColWidth(5) = 2500
MSHFlexGrid34.ColWidth(6) = 1500

MSHFlexGrid34.RowHeight(1) = 500
MSHFlexGrid34.WordWrap = True
End Sub

Private Sub Grid_Head35()
MSHFlexGrid35.Clear
MSHFlexGrid35.Rows = 3
MSHFlexGrid35.FixedRows = 2
MSHFlexGrid35.Cols = 7

MSHFlexGrid35.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid35.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid35.TextMatrix(1, 2) = "AdjustmentDate"
MSHFlexGrid35.TextMatrix(1, 3) = "Rec TDS"
MSHFlexGrid35.TextMatrix(1, 4) = "Rec Other Ded"
MSHFlexGrid35.TextMatrix(1, 5) = "Remark"
MSHFlexGrid35.TextMatrix(1, 6) = "BatchID"

MSHFlexGrid35.ColWidth(0) = 1800
MSHFlexGrid35.ColWidth(1) = 3500
MSHFlexGrid35.ColWidth(2) = 1500
MSHFlexGrid35.ColWidth(3) = 1500
MSHFlexGrid35.ColWidth(4) = 1500
MSHFlexGrid35.ColWidth(5) = 2500
MSHFlexGrid35.ColWidth(6) = 1500

MSHFlexGrid35.RowHeight(1) = 500
MSHFlexGrid35.WordWrap = True
End Sub

Private Sub Grid_Head36()
MSHFlexGrid36.Clear
MSHFlexGrid36.Rows = 3
MSHFlexGrid36.FixedRows = 2
MSHFlexGrid36.Cols = 7

MSHFlexGrid36.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid36.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid36.TextMatrix(1, 2) = "AdjustmentDate"
MSHFlexGrid36.TextMatrix(1, 3) = "Rec TDS"
MSHFlexGrid36.TextMatrix(1, 4) = "Rec Other Ded"
MSHFlexGrid36.TextMatrix(1, 5) = "Remark"
MSHFlexGrid36.TextMatrix(1, 6) = "BatchID"

MSHFlexGrid36.ColWidth(0) = 1800
MSHFlexGrid36.ColWidth(1) = 3500
MSHFlexGrid36.ColWidth(2) = 1500
MSHFlexGrid36.ColWidth(3) = 1500
MSHFlexGrid36.ColWidth(4) = 1500
MSHFlexGrid36.ColWidth(5) = 2500
MSHFlexGrid36.ColWidth(6) = 1500

MSHFlexGrid36.RowHeight(1) = 500
MSHFlexGrid36.WordWrap = True
End Sub

Private Sub Grid_Head20()
MSHFlexGrid20.Clear
MSHFlexGrid20.Rows = 3
MSHFlexGrid20.FixedRows = 2
MSHFlexGrid20.Cols = 8

MSHFlexGrid20.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid20.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid20.TextMatrix(1, 2) = "Exchange Type"
MSHFlexGrid20.TextMatrix(1, 3) = "Invoice Date"
MSHFlexGrid20.TextMatrix(1, 4) = "Invoice No."
MSHFlexGrid20.TextMatrix(1, 5) = "NAV Code"
MSHFlexGrid20.TextMatrix(1, 6) = "Amount"
MSHFlexGrid20.TextMatrix(1, 7) = "BatchID"

MSHFlexGrid20.ColWidth(0) = 1800
MSHFlexGrid20.ColWidth(1) = 3500
MSHFlexGrid20.ColWidth(2) = 1800
MSHFlexGrid20.ColWidth(3) = 1500
MSHFlexGrid20.ColWidth(4) = 1500
MSHFlexGrid20.ColWidth(5) = 1800
MSHFlexGrid20.ColWidth(6) = 1500
MSHFlexGrid20.ColWidth(7) = 1200

MSHFlexGrid20.RowHeight(1) = 600
MSHFlexGrid20.WordWrap = True
End Sub

Private Sub Grid_Head21()
MSHFlexGrid21.Clear
MSHFlexGrid21.Rows = 3
MSHFlexGrid21.FixedRows = 2
MSHFlexGrid21.Cols = 8

MSHFlexGrid21.TextMatrix(1, 0) = "Location Name"
MSHFlexGrid21.TextMatrix(1, 1) = "Client Name"
MSHFlexGrid21.TextMatrix(1, 2) = "Exchange Type"
MSHFlexGrid21.TextMatrix(1, 3) = "Invoice Date"
MSHFlexGrid21.TextMatrix(1, 4) = "Invoice No."
MSHFlexGrid21.TextMatrix(1, 5) = "NAV Code"
MSHFlexGrid21.TextMatrix(1, 6) = "Amount"
MSHFlexGrid21.TextMatrix(1, 7) = "BatchID"

MSHFlexGrid21.ColWidth(0) = 1800
MSHFlexGrid21.ColWidth(1) = 3500
MSHFlexGrid21.ColWidth(2) = 1800
MSHFlexGrid21.ColWidth(3) = 1500
MSHFlexGrid21.ColWidth(4) = 1500
MSHFlexGrid21.ColWidth(5) = 1800
MSHFlexGrid21.ColWidth(6) = 1500
MSHFlexGrid21.ColWidth(7) = 1200

MSHFlexGrid21.RowHeight(1) = 600
MSHFlexGrid21.WordWrap = True
End Sub


Private Sub Xls_Excel_Invoices()

Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_FAGMIS_Invoices.xls")
mRow = oWB.Worksheets.Count
''If 34 - oWB.Worksheets.Count > 0 Then
''   oWB.Worksheets.Add , oWB.Worksheets(mRow), 34 - oWB.Worksheets.Count
''End If


Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "NNNR Invoices Consolidated"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid8.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid8.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid8.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "NNNR Batch Pending By FAG"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid20.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid20.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid20.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet3")
oWS.Name = "NNNR Pending import of TI No"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid21.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid21.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid21.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet4")
oWS.Name = "AdHoc Invoices Consolidated"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid9.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid9.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid9.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet5")
oWS.Name = "AdHoc Pending for FAG Approval"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid22.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid22.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid22.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet6")
oWS.Name = "AdHoc Batch pending by FAG"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid23.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid23.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid23.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet7")
oWS.Name = "AdHoc pending Import TINo FAG"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid24.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid24.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid24.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet8")
oWS.Name = "Reserv Invoices Consolidated"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid10.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid10.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid10.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet9")
oWS.Name = "Reserv Inv Pending FAG Approval"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid25.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid25.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid25.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet10")
oWS.Name = "Reserv Inv Batch Pending by FAG"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid26.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid26.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid26.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet11")
oWS.Name = "Reserv Inv Pending Import TINo"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid27.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid27.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid27.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet12")
oWS.Name = "DP Invoice Consolidated"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid11.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid11.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid11.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet13")
oWS.Name = "DP Invoice Batch Pending by FAG"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid28.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid28.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid28.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet14")
oWS.Name = "DP Inv Pending import of TI No"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid31.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid31.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid31.TextMatrix(I, c)
    Next
Next

oWB.Save
MsgBox ("Excel File for Invoices Created !!! ")
oXL.Visible = True

End Sub


Private Sub Xls_Excel_CR()

Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_FAGMIS_CashReceipt.xls")
mRow = oWB.Worksheets.Count
''If 34 - oWB.Worksheets.Count > 0 Then
''   oWB.Worksheets.Add , oWB.Worksheets(mRow), 34 - oWB.Worksheets.Count
''End If


Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "Cash Receipt Consolidated"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid12.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid12.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid12.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "CR Not Deposited in Bank "
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid15.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid15.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid15.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet3")
oWS.Name = "CR Pending for Approval By FAG"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid17.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid17.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid17.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet4")
oWS.Name = "CR Batch Pending by FAG"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid16.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid16.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid16.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet5")
oWS.Name = "CR Pending import of DocNo FAG"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid18.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid18.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid18.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet6")
oWS.Name = "CR Pending for Inv alloc by FAG"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid19.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid19.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid19.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet7")
oWS.Name = "Advance CR (Consolidated)"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid13.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid13.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid13.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet8")
oWS.Name = "Adv CR pending import of DocNo"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid32.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid32.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid32.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet9")
oWS.Name = "Adv CR pending Inv alloc by FAG"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid33.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid33.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid33.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet10")
oWS.Name = "Other Deduc & TDS Consolidated"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid14.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid14.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid14.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet11")
oWS.Name = "Other ded TDS Approval Pending"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid34.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid34.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid34.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet12")
oWS.Name = "Other deduc TDS Batch Pending "
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid35.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid35.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid35.TextMatrix(I, c)
    Next
Next

Set oWS = oWB.Worksheets("Sheet13")
oWS.Name = "Other deduc Pending Import TINo"
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid36.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid36.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid36.TextMatrix(I, c)
    Next
Next

oWB.Save
MsgBox ("Excel File for Cash Receipt Created !!! ")
oXL.Visible = True
End Sub

