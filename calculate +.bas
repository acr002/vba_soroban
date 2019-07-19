'-----------------------------------------------------------[date: 2019.05.10]
Attribute VB_Name = "Module1"
Option Explicit

'***********************************************
' 2019.05.10(金).new
'***********************************************
' 下記の定数を設定してください。
Private Const BLOCK_YS  As Long = 4
Private Const BLOCK_XS  As Long = 10
Private Const ELS_COUNT As Long = 10
Private Const ELS_MAX   As Long = 3
'***********************************************
Private Enum numbers
  zero = 0
  N1 = 1
  N2 = 2
  N3 = 3
  N4 = 4
  N5 = 5
  N6 = 6
  N7 = 7
  N8 = 8
  N9 = 9
End Enum
'***********************************************

Public Sub calculate_mitori()
  Dim y_base As Long
  Dim ya     As Long
  Dim y      As Long
  Dim x      As Long
  Dim t      As Long
  Dim sum    As Long
  For y = N1 To BLOCK_YS
    y_base = (y - N1) * (ELS_COUNT + N3) + N1
    For x = N1 To BLOCK_XS
      For ya = N1 To ELS_COUNT
        t = set_num(sum, ELS_MAX)
        Cells((ya - N1) + y_base, x).Value = t
        sum = sum + t
      Next ya
      Cells(y_base + ELS_COUNT, x).Value = sum * -1
      sum = zero
    Next x
    'With Cells(y_base + ELS_COUNT, N1).Resize(N1, BLOCK_XS)
    '  .Interior.Color = RGB(192, 192, 192)
    '  .Borders(xlEdgeBottom).LineStyle = xlContinuous
    'End With
  Next y
  'MsgBox "end of run"
End Sub
'-----------------------------------------------------------------------------

' 2019.05.10(金)
' 第一引数: 現在の合計値
' 第二引数: 桁数(基本的に固定の数値でいいと思います。最大値のみ指定。最小値は指定できません。)
Private Function set_num(ByVal a_sum As Long, a_figure As Long) As Long
  Dim t_range As Long
  Dim t_range_double As Long
  Dim cn As Long
  Dim t  As Long
  t_range = 10 ^ a_figure
  t_range_double = t_range * 2
  Randomize
  'Do
  '  cn = cn + 1
  '  t = Rnd() * (t_range + a_sum) - a_sum
  'Loop Until t
  Do
    cn = cn + 1
    t = Rnd() * t_range_double - t_range
    if t then
      if t + a_sum > 0 then
        exit do
      end if
    end if
  Loop
  set_num = t
  If cn > 1 Then
    Debug.Print cn
  End If
End Function
'-----------------------------------------------------------------------------

