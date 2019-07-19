'-----------------------------------------------------------[date: 2019.07.19]
Attribute VB_Name = "Module1"
Option Explicit

'***********************************************
' 2019.04.17(…).new
' 2019.05.13(ŒŽ)
' 2019.07.19(‹à)
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
  N10 = 10
End Enum
'***********************************************

Public Sub calculate_x()
  Dim div_A As Long
  Dim div_B As Long
  Dim mul_A As Long
  Dim mul_B As Long
  Dim y     As Long
  Dim b     As Long
  Dim a     As Long
  Dim i     As Long
  Dim ws    As Worksheet
  Set ws = ThisWorkbook.Worksheets("cal")
  With ThisWorkbook.Worksheets("info")
    mul_A = .Cells(11, 3).Value
    mul_B = .Cells(12, 3).Value
    div_A = .Cells(13, 3).Value
    div_B = .Cells(14, 3).Value
  End With
  ' ‚©‚¯ŽZ ///////////////////////////////////////////////////////////////////
  For i = N1 To 40
    a = set_num_x(mul_A)
    b = set_num_x(mul_B)
    'Debug.Print i, a, b, a * b
    y = i + N1
    ws.Cells(y, N1).Value = i
    ws.Cells(y, N2).Value = a
    ws.Cells(y, N3).Value = b
    'ws.Cells(y, N4).Value =
    ws.Cells(y, N5).Value = a * b
  Next i
  ' Š„‚èŽZ ///////////////////////////////////////////////////////////////////
  For i = N1 To 40
    a = set_num_x(div_A)
    b = set_num_x(div_B)
    y = i + N1
    ws.Cells(y, N7).Value = i
    ws.Cells(y, N8).Value = a * b
    ws.Cells(y, N9).Value = b
    ws.Cells(y, 11).Value = a
  Next i
  Call put_sheet(ws)
  'Debug.Print String(48, "-")
End Sub
'-----------------------------------------------------------------------------

' Œ…”(a_column)•ª‚Ì®”‚ð•Ô‚µ‚Ü‚·B
Private Function set_num_x(ByVal a_column As Long) As Long
  Dim t_min As Long
  Dim t_max As Long
  Dim t     As Long
  Dim cn    As Long
  Randomize
  t_max = N10 ^ a_column
  t_min = N10 ^ (a_column - N1)
  Do
    cn = cn + N1
    t = Int(Rnd() * t_max)
  Loop Until t > t_min
  If cn > N2 Then
    Debug.Print "loop count: "; cn; "‰ñ"
  End If
  set_num_x = t
End Function
'-----------------------------------------------------------------------------

' 2019.07.19(‹à).new
Private Sub put_sheet(aws As Worksheet)
  Dim path   As String
  Dim fn_out As String
  Dim seq    As Long
  seq = ThisWorkbook.Worksheets("info").Cells(16, 3).Value + N1
  path = ThisWorkbook.path & "\"
  fn_out = path & "calc_" & Format(seq, "000") & " " & Format(Date, "yyyymmdd")
  aws.ExportAsFixedFormat _
    Type:=xlTypePDF, _
    Filename:=fn_out, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=False
  ThisWorkbook.Worksheets("info").Cells(16, 3).Value = seq
End Sub
'-----------------------------------------------------------------------------

