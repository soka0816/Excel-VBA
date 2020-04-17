# Excel-VBA
'コントロール配列

Private Sub CommandButton1_Click()

   Dim i As Integer
   Dim myMsg As String
   Dim myOpt As MSForms.OptionButton

    For i = 1 To 3
       Set myOpt = UserForm2.Controls("OptionButton" & i)
       If myOpt.Value = True Then
           myMsg = myOpt.Caption & " が選択されています"
           Exit For
       End If
   Next i

    If myMsg = "" Then
      myMsg = "いずれも選択されていません"
   End If

    MsgBox myMsg

End Sub
