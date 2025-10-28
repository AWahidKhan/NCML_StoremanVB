VERSION 5.00
Begin VB.Form FrmHelp 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   7320
      Width           =   7335
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Zurich BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Zurich BT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000007&
      Height          =   810
      Left            =   120
      ScaleHeight     =   750
      ScaleWidth      =   7275
      TabIndex        =   2
      Top             =   120
      Width           =   7335
      Begin VB.TextBox TxtFieldName 
         Height          =   615
         Left            =   45
         TabIndex        =   3
         Top             =   30
         Width           =   7140
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      Height          =   6135
      Left            =   120
      ScaleHeight     =   6075
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   1080
      Width           =   7335
      Begin VB.TextBox TxtDescription 
         Height          =   5970
         Left            =   45
         TabIndex        =   1
         Top             =   30
         Width           =   7140
      End
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSave_Click()
Call Help_DetailTable(" Where FormName = '" & Gen_FormName & "' And FieldName = '" & Gen_FieldName & "'")
If RsHelp_Detail.RecordCount = 0 Then
   RsHelp_Detail.AddNew
   RsHelp_Detail!FormName = Gen_FormName
   RsHelp_Detail!FieldName = Gen_FieldName
End If
RsHelp_Detail!EngDescription = TxtDescription
RsHelp_Detail!DisplayFieldName = TxtFieldName
RsHelp_Detail.Update
MsgBox "Saved successfully!!!"
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

Call Help_DetailTable(" Where FormName = '" & Gen_FormName & "' And FieldName = '" & Gen_FieldName & "'")

Frame1.Visible = False
TxtDescription.Locked = True
TxtFieldName.Locked = True
If LCase(UserType) = "power" Then
   Frame1.Visible = True
   TxtDescription.Locked = False
   TxtFieldName.Locked = False
End If

If RsHelp_Detail.RecordCount = 0 Then
   TxtDescription = ""
   TxtFieldName = Gen_FieldName
Else
   TxtDescription = RsHelp_Detail!EngDescription
   TxtFieldName = RsHelp_Detail!DisplayFieldName
End If

End Sub
