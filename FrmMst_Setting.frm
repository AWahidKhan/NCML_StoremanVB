VERSION 5.00
Begin VB.Form FrmMst_Setting 
   Caption         =   "Setting"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4065
   ScaleWidth      =   6480
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   2520
      Width           =   6255
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         TabIndex        =   4
         Top             =   30
         Width           =   6015
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.TextBox TxtNoofDays 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Sport Storage Charges No of Days"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmMst_Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdUpdate_Click()
RsMst_Config!SportStorageChargesDays = TxtNoofDays
RsMst_Config.Update

MsgBox "Data updated !!!! "
Unload Me
End Sub

Private Sub Form_Load()
CmdUpdate.Enabled = False
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call TableMst_Config
TxtNoofDays = IIf(IsNull(RsMst_Config!SportStorageChargesDays), 0, RsMst_Config!SportStorageChargesDays)
CmdUpdate.Enabled = True
End Sub
