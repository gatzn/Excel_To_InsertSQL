Attribute VB_Name = "Module1"
Option Explicit

Sub INSERT文を生成_Click()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim sql As String
    sql = ""

    Dim head As String
    head = vbLf & "insert into " & Cells(1, 2).Value & "." & Cells(2, 2).Value & " ("

    Dim target As Range
    Set target = ws.UsedRange

    Dim currentCell As Range
    Dim i As Integer

    ' カラム名を列挙
    For i = 1 To target.Columns.Count
        If (i <> 1) Then
            head = head & ","
        End If
        Set currentCell = ws.Cells(3, i)
        head = head & currentCell.Value
    Next
    head = head & ") values("

    ' 値を列挙
    Dim j As Integer
    Dim strtmp As String
    For j = 4 To target.Rows.Count

        ' 非表示の行は出力しない
        If Rows(j).Hidden Then GoTo Next_RowLoop

        sql = sql & head

        For i = 1 To target.Columns.Count
            If (i <> 1) Then
                sql = sql & ","
            End If
            Set currentCell = ws.Cells(j, i)
            If (IsNull(currentCell) Or currentCell.Value = "" Or Trim(currentCell.Value) = "(null)" Or Trim(currentCell.Value) = "null") Then

                ' null
                sql = sql & "null"
            ElseIf IsNumeric(currentCell.Value) Then

                ' 数値
                sql = sql & currentCell.Value
            ElseIf Left(currentCell.Value, 8) = "(SELECT " Then

                ' SQLっぽいものはそのまま
                sql = sql & currentCell.Value
            Else

                ' 文字列はシングルクォーテーションをエスケープする
                strtmp = Replace(currentCell.Value, "'", "''")

                ' セル内の改行を改行コードに変換
                strtmp = Replace(strtmp, vbLf, "' || CHR(13) || CHR(10) ||'")

                sql = sql & "'" & strtmp & "'"
            End If
        Next

        sql = sql & ");"

Next_RowLoop:
    Next

    sql = sql & vbLf & ""

    Dim cb As New DataObject
    With cb
        .SetText sql
        .PutInClipboard
    End With
    MsgBox ("クリップボードにコピーしました")
End Sub

