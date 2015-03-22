Public Class Datatable2TSV

    ''' <summary>
    ''' タブ区切りファイルをDatatableにロードする
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <param name="readHeader"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function tsv2dt(fileName As String, dataTableName As String, _
                Optional readHeader As Boolean = False) As DataTable

        Dim _dt As DataTable = New System.Data.DataTable(dataTableName)

        If Not New System.IO.FileInfo(fileName).Exists Then Throw New Exception(fileName & " is not found")

        Dim r As System.IO.StreamReader = Nothing

        Try
            r = New System.IO.StreamReader(fileName)

            Dim rowCount = 0
            Do While r.Peek > -1
                Dim line As String() = (r.ReadLine).Split(vbTab)
                rowCount += 1

                If rowCount = 1 Then

                    If readHeader = False Then
                        '   1行目をカラム名として使用
                        For i As Integer = 0 To line.Count - 1
                            _dt.Columns.Add(line(i), GetType(String))
                        Next

                    Else
                        For i As Integer = 0 To line.Count - 1
                            _dt.Columns.Add("column" & line(i), GetType(String))
                        Next

                    End If

                Else
                    Dim row As System.Data.DataRow = _dt.NewRow
                    For i As Integer = 0 To line.Count - 1
                        row(i) = line(i)
                    Next
                    _dt.Rows.Add(row)

                End If

            Loop

        Catch ex As Exception

        Finally
            If Not r Is Nothing Then r.Close()

        End Try

        Return _dt

    End Function

    ''' <summary>
    ''' Datatableの内容をタブ区切りのファイルに出力する
    ''' </summary>
    ''' <param name="dt">出力するDatatable</param>
    ''' <param name="fileName">出力先ファイル名</param>
    ''' <param name="writeHeader">ヘッダー出力有無</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function dt2tsv(dt As DataTable, fileName As String, _
            Optional writeHeader As Boolean = True, Optional existFileAction As FileAction = FileAction.DeleteCreate) As Long

        Dim f As System.IO.FileInfo
        Dim overWrite As Boolean = False

        If existFileAction = FileAction.DeleteCreate Then
            If New System.IO.FileInfo(fileName).Exists Then
                f = New System.IO.FileInfo(fileName)
                f.Delete()
            End If
        ElseIf existFileAction = FileAction.Backup Then
            If New System.IO.FileInfo(fileName).Exists Then
                f = New System.IO.FileInfo(fileName)
                f.CopyTo(fileName & ".backup", True)
                f.Delete()
            End If

        ElseIf existFileAction = FileAction.Overwrite Then
            If New System.IO.FileInfo(fileName).Exists Then
                overWrite = True
            End If
        End If

        Dim w As System.IO.StreamWriter = Nothing

        Try

            w = New System.IO.StreamWriter(fileName, IIf(existFileAction = FileAction.Overwrite, True, False))

            Dim cindex As Integer = dt.Columns.Count

            '   Output Header
            If writeHeader = True And Not existFileAction = FileAction.Overwrite Then
                For h As Integer = 0 To dt.Columns.Count - 1

                    Dim fieldName As String = dt.Columns(h).Caption
                    w.Write(fieldName)
                    If h < dt.Columns.Count - 1 Then w.Write(vbTab)

                Next h
                w.Write(vbCrLf)
            End If

            '   Output Data
            For Each _row As DataRow In dt.Rows
                For h As Integer = 0 To dt.Columns.Count - 1

                    Dim fieldValue As String = _row(h).ToString
                    w.Write(fieldValue)
                    If h < dt.Columns.Count - 1 Then w.Write(vbTab)

                Next h
                w.Write(vbCrLf)
            Next

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            If Not w Is Nothing Then
                w.Close()
            End If
        End Try

        Return 0
    End Function

End Class

Public Enum FileAction As Integer
    DeleteCreate = 1
    Overwrite = 2
    Backup = 3

End Enum