Attribute VB_Name = "Module1"
Option Explicit

Sub INSERT���𐶐�_Click()
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

    ' �J���������
    For i = 1 To target.Columns.Count
        If (i <> 1) Then
            head = head & ","
        End If
        Set currentCell = ws.Cells(3, i)
        head = head & currentCell.Value
    Next
    head = head & ") values("

    ' �l���
    Dim j As Integer
    Dim strtmp As String
    For j = 4 To target.Rows.Count

        ' ��\���̍s�͏o�͂��Ȃ�
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

                ' ���l
                sql = sql & currentCell.Value
            ElseIf Left(currentCell.Value, 8) = "(SELECT " Then

                ' SQL���ۂ����̂͂��̂܂�
                sql = sql & currentCell.Value
            Else

                ' ������̓V���O���N�H�[�e�[�V�������G�X�P�[�v����
                strtmp = Replace(currentCell.Value, "'", "''")

                ' �Z�����̉��s�����s�R�[�h�ɕϊ�
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
    MsgBox ("�N���b�v�{�[�h�ɃR�s�[���܂���")
End Sub

