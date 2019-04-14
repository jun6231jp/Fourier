Option Explicit
Sub Fourier_delta() 'デルタ関数生成
Dim i As Integer
Dim j As Double
Dim delta As Double
Dim start As Double
Dim last As Double
Dim grid As Double
Dim n As Integer
Dim count As Long
Dim chk As Integer
MsgBox ("波の重ね合わせでデルタ関数を作ります。")

'設定値を入力
    delta = InputBox("デルタ関数の幅" & vbCrLf & "" & vbCrLf & "デルタ関数は幅0/面積1の波形ですが" & vbCrLf & "ここでは正の波形幅を設定し" & vbCrLf & "面積1の波形を生成します。" & vbCrLf & "例：0.3")
    delta = delta / 2
    n = InputBox("重ね合わせる波の数" & vbCrLf & "" & vbCrLf & "フーリエ級数展開では対象の波形を無限の" & vbCrLf & "正弦波と余弦波の重ね合わせで記述します。" & vbCrLf & "ここでは有限な数の波を重ね合わせることで" & vbCrLf & "理想波形に近づけます。" & vbCrLf & "例：150")
If n > 254 Then
MsgBox ("重ね合わせの数が多すぎます")
Else
    start = InputBox("θ最小値" & vbCrLf & "" & vbCrLf & "波形の横軸の左端です。" & vbCrLf & "2π付近の値を推奨します。" & vbCrLf & "例：5.0")
    last = InputBox("θ最大値" & vbCrLf & "" & vbCrLf & "波形の横軸の右端です。" & vbCrLf & "2π付近の値を推奨します。" & vbCrLf & "例：7.5")
    grid = InputBox("グラフ粒度とグリッド" & vbCrLf & "" & vbCrLf & "波形の粒度を設定します。" & vbCrLf & "グラフのグリッドは設定値の10倍となります。" & vbCrLf & "例：0.01")
If grid * 5 > delta Then
MsgBox ("グリッドが大きすぎます。" & vbCrLf & "波形幅の10%未満に設定してください。")
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
Sheet2.Cells(3 + count, 2).Value = 1 / (2 * 3.14)
j = j + grid
count = count + 1
Wend
'フーリエ級数展開の表作成
mktbl3 delta, n, count
For i = 0 To n
   draw_data3 count, i, start, last, grid, delta
Next i
draw_data4 count, n, start, last, grid, delta
End If
End If

MsgBox ("続けて波形生成可能です。")
End Sub

Sub mktbl3(delta As Double, n As Integer, k As Long) 'フーリエ級数展開の表の中身を作成
 Dim i As Integer
 Dim j As Integer
 For i = 1 To n
  For j = 1 To k
   Sheet2.Cells(2 + j, 2 + i).Value = Sheet2.Cells(2 + j, 1 + i).Value + ((Sin(delta * i) / (delta * i)) * Cos(i * Sheet2.Cells(2 + j, 1))) / 3.14
  Next j
 Next i
End Sub

Sub draw_data3(length As Long, n As Integer, start As Double, last As Double, grid As Double, delta As Double)    '外側から引数をFor文で与えることでアニメーションでグラフを表示
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
    ActiveChart.ChartTitle.Characters.Text = "矩形波生成の過程"   ' グラフタイトル
    ActiveChart.Axes(xlCategory, xlPrimary).HasTitle = True
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "θ"    ' x 軸タイトル
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = True
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "P"       ' y 軸タイトル

    ActiveChart.Axes(xlCategory, xlPrimary).MinimumScale = start      ' x 軸の最小値
    ActiveChart.Axes(xlCategory, xlPrimary).MaximumScale = last   ' x 軸の最大値
    ActiveChart.Axes(xlCategory, xlPrimary).MajorUnit = grid * 10      ' x 軸の目盛幅
    ActiveChart.Axes(xlValue, xlPrimary).MinimumScale = -0.25 / delta ' y 軸の最小値
    ActiveChart.Axes(xlValue, xlPrimary).MaximumScale = 1 / delta - grid * 10 ' y 軸の最大値
    ActiveChart.Axes(xlValue, xlPrimary).MajorUnit = 0.25 / delta   ' y 軸の目盛幅
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


Sub draw_data4(length As Long, n As Integer, start As Double, last As Double, grid As Double, delta As Double)    '最後に全ての線を重ねて自動グラフ作成
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
    ActiveChart.ChartTitle.Characters.Text = "矩形波生成結果"   ' グラフタイトル
    ActiveChart.Axes(xlCategory, xlPrimary).HasTitle = True
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "θ"    ' x 軸タイトル
    ActiveChart.Axes(xlValue, xlPrimary).HasTitle = True
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "P"       ' y 軸タイトル

    ActiveChart.Axes(xlCategory, xlPrimary).MinimumScale = start      ' x 軸の最小値
    ActiveChart.Axes(xlCategory, xlPrimary).MaximumScale = last   ' x 軸の最大値
    ActiveChart.Axes(xlCategory, xlPrimary).MajorUnit = grid * 10      ' x 軸の目盛幅
    ActiveChart.Axes(xlValue, xlPrimary).MinimumScale = -0.25 / delta     ' y 軸の最小値
    ActiveChart.Axes(xlValue, xlPrimary).MaximumScale = 1 / delta ' y 軸の最大値
    ActiveChart.Axes(xlValue, xlPrimary).MajorUnit = 0.25 / delta     ' y 軸の目盛幅
  For i = 0 To n
    ActiveChart.SeriesCollection(1 + i).Border.Weight = xlMedium
  Next i


    ActiveChart.HasLegend = False  ' 凡例に関する設定
    'ActiveChart.Legend.Position = xlBottom
    ActiveChart.SeriesCollection(1 + n).Name = "circle"
    ActiveChart.SeriesCollection(1).Name = "circle"
    ActiveChart.Deselect

End Sub
