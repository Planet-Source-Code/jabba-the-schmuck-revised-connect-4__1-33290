Attribute VB_Name = "Module1"
'This function allows me to play sounds, and is referred to in the
'PlaySound routine below.
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub CloseForm(frmAnyForm As Form)

'Play sound 0 or 1.
Call PlaySound(Int(Rnd * 2))

'These two loops reduce the height and width of a form
'until it gets to a certain size.
Do While frmAnyForm.Height > 410
 DoEvents
 frmAnyForm.Height = frmAnyForm.Height - 4
 frmAnyForm.Top = frmAnyForm.Top + 1
Loop

Do While frmAnyForm.Width > 1700
 DoEvents
 frmAnyForm.Width = frmAnyForm.Width - 4
 frmAnyForm.Left = frmAnyForm.Left + 3
Loop

'Unloads the form.
Unload frmAnyForm
'End the program.
End

End Sub

Public Function iIncrement(ByRef variable As Integer)
'Increases the variable passed to this routine by one.
variable = variable + 1
End Function

Public Function bIncrement(ByRef variable As Byte)
'Increases the variable passed to this routine by one.
variable = variable + 1
End Function

Public Sub PlaySound(wavFile As Byte)

'Don't play sound if the sound option is un-checked.
If frmConnect4.chkSound.Value = 0 Then Exit Sub

'This allows me to play any of the sounds below by simply calling
'this routine, together with the number of the sound I want to play.
Dim sound
Dim sFilename As String

  Select Case wavFile
    Case 0
      sFilename = "e1"
    Case 1
      sFilename = "e2"
    Case 2
      sFilename = "b1"
    Case 3
      sFilename = "b2"
    Case 4
      sFilename = "w1"
    Case 5
      sFilename = "w2"
    Case 6
      sFilename = "w3"
    Case 7
      sFilename = "l1"
    Case 8
      sFilename = "l2"
    Case 9
      sFilename = "l3"
    Case 10
      sFilename = "l4"
    Case 11
      sFilename = "l5"
  End Select

sFilename = sFilename & ".wav"
sound = sndPlaySound(App.Path & "\Sounds\" & sFilename, &H1)

End Sub

