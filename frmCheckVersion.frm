VERSION 5.00
Begin VB.Form frmCheckVersion 
   Caption         =   "Check Version Info and Install new Version"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6720
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCheckVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim verEXE As String
Dim verDB As String


Private Sub GetDBVersion()
Call TableMst_Config("")
'ConfigRs.MoveFirst
verDB = IIf(IsNull(RsMst_Config!Version_No), "1.0.0.0.0", RsMst_Config!Version_No)
If Len(verDB) >= 10 Then
   verDB = "0"
End If
End Sub

Private Sub Form_Load()

Call GetVersionNo
verEXE = vVersionNumber
Gen_VersionEXE = verEXE
reverse:
Call GetDBVersion

If (CompareVersion(verEXE, verDB) = False) Then
    GoTo reverse
Else
    'Check for Exe Expire Date !!!! if Falase Change the Flag and Expire
'    If RsMst_Config!EndDate < Date Or ConfigRs!Approved_by = 0 Then
'       MsgBox "Version Expire ; Please get in touch with Sankaranarayanan (02240419145) or Mayur Dhamani (022-40419172) for rectifying this error immediately "
'       ConfigRs!Approved_by = 0
'       ConfigRs.Update
'       End
'    End If
End If

frmCheckVersion.Hide

'End
End Sub

