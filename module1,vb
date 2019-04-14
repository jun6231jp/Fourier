Option Explicit
Sub Fourier_circle() '円波形生成
Dim i As Integer
Dim j As Double
Dim delta As Double
Dim start As Double
Dim last As Double
Dim grid As Double
Dim n As Integer
Dim k As Integer
Dim count As Long
Dim chk As Integer
MsgBox ("波の重ね合わせで円波形を作ります。")

'設定値を入力
    delta = InputBox("円波形の半径" & vbCrLf & "" & vbCrLf & "例：0.8")
    n = InputBox("重ね合わせる波の数" & vbCrLf & "" & vbCrLf & "フーリエ級数展開では対象の波形を無限の" & vbCrLf & "正弦波と余弦波の重ね合わせで記述します。" & vbCrLf & "ここでは有限な数の波を重ね合わせることで" & vbCrLf & "理想波形に近づけます。" & vbCrLf & "例：80")
If n > 254 Then
MsgBox ("重ね合わせの数が多すぎます")
Else
    k = InputBox("超幾何級数の足し合わせ数" & vbCrLf & "" & vbCrLf & "円波形のフーリエ級数展開を行うと" & vbCrLf & "展開された各波の振幅は超幾何級数となります。" & vbCrLf & "超幾何級数は無限項の和ですが、" & vbCrLf & "ここでは有限な項まで求め理想形に近づけます。" & vbCrLf & "例：60")
    start = InputBox("θ最小値" & vbCrLf & "" & vbCrLf & "波形の横軸の左端です。" & vbCrLf & "2π付近の値を推奨します。" & vbCrLf & "例：5.0")
    last = InputBox("θ最大値" & vbCrLf & "" & vbCrLf & "波形の横軸の右端です。" & vbCrLf & "2π付近の値を推奨します。" & vbCrLf & "例：7.5")
    grid = InputBox("グラフ粒度とグリッド" & vbCrLf & "" & vbCrLf & "波形の粒度を設定します。" & vbCrLf & "グラフのグリッドは設定値の10倍となります。" & vbCrLf & "例：0.01")
If grid * 10 > delta Then
MsgBox ("グリッドが大きすぎます。" & vbCrLf & "円の半径の10%未満に設定してください。")
ElseIf grid * 20 > last - start Then
MsgBox ("グリッドが大きすぎます。" & vbCrLf & "グラフ幅の5%未満に設定してください。")
Else
    '現状のシートのクリア
    Worksheets("Bessel").Activate
    Worksheets("Bessel").Cells.Clear
    Worksheets("Fourier").Activate
    Worksheets("Fourier").Cells.Clear
'グラフ作成のための表作成
Sheet2.Cells(2, 1).Value = "θ"
For i = 0 To n
Sheet2.Cells(2, 2 + i).Value = "n=" & i
Next i
j = start
count = 0

While j < last
Sheet2.Cells(3 + count, 1).Value = j
Sheet2.Cells(3 + count, 2).Value = delta * delta / 4
j = j + grid
count = count + 1
Wend
'フーリエ級数展開したそれぞれの周波数の振幅をベッセル関数で割り出す必要がある。
'ベッセル関数の表作成
   write_k (k)
   write_n (n)
   mktbl delta, n, k
   sum n, k
   'ベッセル関数からフーリエ級数展開の表作成
   mktbl2 delta, n, count
For i = 0 To n
   draw_data count, i, start, last, grid
Next i
draw_data2 count, n, start, last, grid
End If
End If

MsgBox ("続けて波形生成可能です。")
End Sub


Sub write_k(k As Integer)  'ベッセル関数のための表の縦軸を作成

Dim i As Integer
Sheet1.Cells(2, 1).Value = "k"
For i = 0 To k
Sheet1.Cells(3 + i, 1).Value = i
If i = 0 Then
Sheet1.Cells(3 + i, 2).Value = 1
Else
Sheet1.Cells(3 + i, 2).Value = 0
End If
Next i

End Sub
Sub write_n(n As Integer)  'ベッセル関数のための表の横軸を作成

Dim i As Integer
Sheet1.Cells(1, 2).Value = "n"
For i = 0 To n
Sheet1.Cells(2, 2 + i).Value = i
Sheet1.Cells(3, 2 + i).Value = 1
Next i

End Sub

Sub mktbl(delta As Double, n As Integer, k As Integer)  'ベッセル関数の表の中身を作成
Dim i As Integer
Dim j As Integer
For i = 1 To n
For j = 1 To k
Sheet1.Cells(3 + j, 2 + i).Value = ((-Sheet1.Cells(2, 2 + i).Value * Sheet1.Cells(2, 2 + i).Value * delta * delta / 4) ^ Sheet1.Cells(3 + j, 1).Value) / (Fact(Sheet1.Cells(3 + j, 1).Value) * Fact(Sheet1.Cells(3 + j, 1).Value + 1))
Next j
Next i

End Sub
Sub sum(n As Integer, k As Integer)  '作成したベッセル関数の級数を実際に計算し、目的の波形をフーリエ級数展開したときの振幅を求める。
Dim Ans As Double
Dim i As Integer
Dim j As Integer
Dim p As Integer

Dim temp As Double
Ans = 0
temp = 2
Sheet2.Cells(1, 1).Value = "An"
For i = 0 To n
 For j = 0 To k
  Ans = Sheet1.Cells(3 + j, 2 + i).Value + Ans
 Next j
 '値が大きすぎるまたは小さすぎる場合オーバーフローにより値がおかしくなることがある。
 'オーバーフローするような段階での振幅はきわめて小さいため、0にする。
 'アルゴリズム改良の余地あり。


'ベッセル関数は減衰振動特性。
'過去の極値を超えた場合オーバーフローしているため0にする。
If Ans > 0 Then
 If temp < Ans Then
 Ans = 0
 End If
Else
 If temp < -Ans Then
 Ans = 0
 End If
End If
   Sheet2.Cells(1, 2 + i).Value = Ans

If i > 2 Then
' ベッセル関数の極値の絶対値をtempに保存しておく
If Sheet2.Cells(1, 1 + i).Value < 0 Then
 If Sheet2.Cells(1, 2 + i).Value > Sheet2.Cells(1, 1 + i).Value Then
  If Sheet2.Cells(1, 1 + i).Value < Sheet2.Cells(1, i).Value Then
    temp = -Sheet2.Cells(1, 1 + i).Value
  End If
 End If
Else
 If Sheet2.Cells(1, 2 + i).Value < Sheet2.Cells(1, 1 + i).Value Then
  If Sheet2.Cells(1, 1 + i).Value > Sheet2.Cells(1, i).Value Then
    temp = Sheet2.Cells(1, 1 + i).Value
  End If
 End If
End If
End If
Ans = 0
Next i

End Sub
Sub mktbl2(delta As Double, n As Integer, k As Long) 'フーリエ級数展開の表の中身を作成
 Dim i As Integer
 Dim j As Integer
 For i = 1 To n
  For j = 1 To k
   Sheet2.Cells(2 + j, 2 + i).Value = Sheet2.Cells(2 + j, 1 + i).Value + Sheet2.Cells(1, 2 + i).Value * (delta * delta / 4) * 2 * Cos(Sheet2.Cells(2 + j, 1).Value * i)
  Next j
 Next i
End Sub

Function Fact(n As Integer) As Double  '計算のための関数。階乗を計算。
Dim i As Integer
 Fact = 1
 For i = 1 To n
    Fact = Fact * i
 Next i
End Function



Sub draw_data(length As Long, n As Integer, start As Double, last As Double, grid As Double)    '外側から引数をFor文で与えることでアニメーションでグラフを表示
    Dim i As Long
    Dim x As Double, dx As Double


    ' アニメーションにするためグラフ削除はせず時間を短縮。
    'If Sheet2.ChartObjects.count > 0 Then
    '    For i = 1 To Sheet2.ChartObjects.count
    '        If Sheet2.ChartObjects(i).Name = "fouriergraph" Then
    '            Sheet2.ChartObjects(i).Delete
    '            Exit For
    '        End If
    '    Next i
    'End If

    ' グラフ記入
    Sheet2.ChartObjects.Add(200, 20, 430, 440).Select  ' 位置 (200,20) に 360x215 のサイズのグラフ追加
    Selection.Name = "fouriergraph" & n ' グラフにfouriergraph(番号)という名前をつける

    ActiveChart.ChartType = xlXYScatterLinesNoMarkers   ' 散布図のデータポイントなし
    ActiveChart.SetSourceData Union(Range(Cells(2, 1), Cells(length + 2, 1)), Range(Cells(2, n + 2), Cells(length + 2, n + 2))), xlColumns ' 離れた列は Union で Range を作成すれば良い
    ActiveChart.Location xlLocationAsObject, "Fourier"

    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Characters.Text = "円波形生成の過程"   ' グラフタイトル
    ActiveChart.Axes(xlCategory, xlPrimary).HasTitle = True
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "θ"    ' x 軸タイトル
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = True
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "P"       ' y 軸タイトル

    ActiveChart.Axes(xlCategory, xlPrimary).MinimumScale = start      ' x 軸の最小値
    ActiveChart.Axes(xlCategory, xlPrimary).MaximumScale = last   ' x 軸の最大値
    ActiveChart.Axes(xlCategory, xlPrimary).MajorUnit = grid * 10      ' x 軸の目盛幅
    ActiveChart.Axes(xlValue, xlPrimary).MinimumScale = -grid * 10    ' y 軸の最小値
    ActiveChart.Axes(xlValue, xlPrimary).MaximumScale = last - start - grid * 10 ' y 軸の最大値
    ActiveChart.Axes(xlValue, xlPrimary).MajorUnit = grid * 10      ' y 軸の目盛幅
'For i = 0 To n
'    ActiveChart.SeriesCollection(1 + i).Border.Weight = xlMedium
'  Next i
     ActiveChart.SeriesCollection(1).Border.Weight = xlMedium

    ActiveChart.HasLegend = False  ' 凡例に関する設定
    'ActiveChart.Legend.Position = xlBottom
'    ActiveChart.SeriesCollection(1 + n).Name = "circle"
 ActiveChart.SeriesCollection(1).Name = "circle"
    ActiveChart.Deselect
    'グラフ表示後Waitを入れる。
    Application.Wait [Now() + "0:00:00.0005"]
End Sub


Sub draw_data2(length As Long, n As Integer, start As Double, last As Double, grid As Double)    '最後に全ての線を重ねて自動グラフ作成
    Dim i As Long
    Dim j As Long
    Dim x As Double, dx As Double
    For j = 0 To n
    ' アニメーションで作成した全てのグラフ削除
    If Sheet2.ChartObjects.count > 0 Then
        For i = 1 To Sheet2.ChartObjects.count
            If Sheet2.ChartObjects(i).Name = "fouriergraph" & j Then   ' "fouriergraph" という名前を指定して削除
                Sheet2.ChartObjects(i).Delete
                Exit For
            End If
        Next i
    End If
    Next j

    ' 過去の全重ねグラフ削除
    If Sheet2.ChartObjects.count > 0 Then
        For i = 1 To Sheet2.ChartObjects.count
            If Sheet2.ChartObjects(i).Name = "fouriergraphx" Then   ' "fouriergraphx" という名前を指定して削除
                Sheet2.ChartObjects(i).Delete
                Exit For
            End If
        Next i
    End If

    ' グラフ記入
    Sheet2.ChartObjects.Add(200, 20, 430, 440).Select  ' 位置 (200,20) に 360x215 のサイズのグラフ追加
    Selection.Name = "fouriergraphx"  ' グラフに名前をつける

    ActiveChart.ChartType = xlXYScatterLinesNoMarkers   ' 散布図のデータポイントなし
    ActiveChart.SetSourceData Range(Cells(2, 1), Cells(length + 2, n + 2)), xlColumns ' 離れた列は Union で Range を作成すれば良い
    ActiveChart.Location xlLocationAsObject, "Fourier"

    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Characters.Text = "円波形生成結果"   ' グラフタイトル
    ActiveChart.Axes(xlCategory, xlPrimary).HasTitle = True
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "θ"    ' x 軸タイトル
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = True
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "P"       ' y 軸タイトル

    ActiveChart.Axes(xlCategory, xlPrimary).MinimumScale = start      ' x 軸の最小値
    ActiveChart.Axes(xlCategory, xlPrimary).MaximumScale = last   ' x 軸の最大値
    ActiveChart.Axes(xlCategory, xlPrimary).MajorUnit = grid * 10      ' x 軸の目盛幅
    ActiveChart.Axes(xlValue, xlPrimary).MinimumScale = -grid * 10       ' y 軸の最小値
    ActiveChart.Axes(xlValue, xlPrimary).MaximumScale = last - start - grid * 10  ' y 軸の最大値
    ActiveChart.Axes(xlValue, xlPrimary).MajorUnit = grid * 10      ' y 軸の目盛幅
  For i = 0 To n
    ActiveChart.SeriesCollection(1 + i).Border.Weight = xlMedium
  Next i


    ActiveChart.HasLegend = False  ' 凡例に関する設定
    'ActiveChart.Legend.Position = xlBottom
    ActiveChart.SeriesCollection(1 + n).Name = "circle"
    ActiveChart.SeriesCollection(1).Name = "circle"
    ActiveChart.Deselect

End Sub 
