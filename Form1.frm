VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   720
      TabIndex        =   4
      Top             =   4320
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   4575
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   3840
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function StripPath(T$) As String
Dim x%, ct%
  StripPath$ = T$
  x% = InStr(T$, "\")
  Do While x%
    ct% = x%
    x% = InStr(ct% + 1, T$, "\")
  Loop
  If ct% > 0 Then StripPath$ = Mid$(T$, ct% + 1)
End Function


Sub UpdatePath()
Dim I, D, J, K As Integer
  For D = 0 To List1.ListCount - 1
     List1.RemoveItem "0"
  Next D

  If Not Right(Dir1.List(-1), 1) = "\" Then
     List1.AddItem "[^] .."
  End If

  For I = 0 To Dir1.ListCount - 1
     List1.AddItem "[\] " & StripPath(Dir1.List(I))
  Next I

  For J = 0 To File1.ListCount - 1
     List1.AddItem "[*] " & File1.List(J)
  Next J

  For K = 0 To Drive1.ListCount - 1
     List1.AddItem "[o] " & Drive1.List(K)
  Next K

  Label1.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
  Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
  UpdatePath
End Sub

Private Sub Form_Load()
  Drive1.Visible = True
  File1.Visible = True
  Dir1.Visible = True
  UpdatePath
  Me.Move (Screen.Width - Me.Width) / 2, _
          (Screen.Height - Me.Height) / 2
End Sub

Private Sub List1_DblClick()
On Error GoTo ErrHdlr
  If Right(List1.Text, 2) = ".." Then
     Dir1.Path = Dir1.Path & "\.."
  ElseIf Left(List1.Text, 3) = "[\]" Then
     If Right(Dir1.List(-1), 1) = "\" Then
       Dir1.Path = Dir1.Path & _
                 Right(List1.Text, Len(List1.Text) - 4)
     Else
        Dir1.Path = Dir1.Path & _
                    "\" & Right(List1.Text, _
                                Len(List1.Text) - 4)
     End If
  ElseIf Left(List1.Text, 3) = "[o]" Then
     Drive1.Drive = Right(Left(List1.Text, 6), 2)
  Else
     MsgBox "File " & Chr(34) & _
            Right(List1.Text, Len(List1.Text) - 4) & _
            Chr(34) & " dipilih.", _
            vbInformation, "File Terpilih"
  End If
  Exit Sub
ErrHdlr:
   MsgBox "Drive tidak siap!", vbCritical, "Tidak Siap"
   Exit Sub
End Sub


