Option Explicit
Private Sub UserForm_Initialize()
    UserForm1.Caption = "波形選択"
    OptionButton1.Caption = "矩形波"
    OptionButton2.Caption = "三角波"
    OptionButton3.Caption = "半円波"
    CommandButton1.Caption = "決定"
End Sub
Private Sub CommandButton1_Click()
If OptionButton1.Value = True Then
Unload Me
MsgBox ("矩形波の生成を開始します。")
Module3.Fourier_delta
ElseIf OptionButton2.Value = True Then
Unload Me
MsgBox ("三角波の生成を開始します。")
Module4.Fourier_triangle
ElseIf OptionButton3.Value = True Then
Unload Me
MsgBox ("半円波の生成を開始します。")
Module2.Fourier_circle
End If

End Sub
