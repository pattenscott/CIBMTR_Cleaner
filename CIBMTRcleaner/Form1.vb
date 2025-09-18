Public Class Form1
    Private strHeadersPath As String
    Private strLoadPath As String
    Private TheHashTable As New Hashtable
    Private TheHeadersHashTable As New Hashtable
    Private strOutputString As String = ""
    Private blnStripOutQuotesAndCommas As Boolean = False
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            If MsgBox("Strip out Quotes and Commas?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                blnStripOutQuotesAndCommas = True
            End If
            GetLoadPath()
            LoadHeadersFile()
            ParseExtractFile()
            ParseHeadersFile()
            BuildColumnsInclsExtractedLine()
            SolveForColumnOrder()
            BuildOutputString()
            SaveOutPut()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub



    Private Sub LoadHeadersFile()
        Dim fd As OpenFileDialog = New OpenFileDialog()

        fd.Title = "Open HEADERS File"
        fd.InitialDirectory = "C:\"
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strHeadersPath = fd.FileName
        Else
            End
        End If
    End Sub
    Private Sub GetLoadPath()
        Dim fd As OpenFileDialog = New OpenFileDialog()

        fd.Title = "Open CIBMTR File"
        fd.InitialDirectory = "C:\"
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strLoadPath = fd.FileName
        Else
            End
        End If

    End Sub
    Private Sub ParseExtractFile()
        Try
            Dim sFULLREAD As String = My.Computer.FileSystem.ReadAllText(strLoadPath)
            Dim sSeperator As String = vbCrLf
            Dim sParsed() As String = sFULLREAD.Split(sSeperator)

            For i As Integer = sParsed.GetLowerBound(0) To sParsed.GetUpperBound(0)
                If sParsed(i).Length > 0 Then

                    Dim el As New clsExtractedLine
                    el.sLine = sParsed(i)
                    TheHashTable.Add(i, el)

                End If
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub ParseHeadersFile()
        Try
            Dim sFULLREAD As String = My.Computer.FileSystem.ReadAllText(strHeadersPath)
            Dim sSeperator As String = vbCrLf
            Dim sParsed() As String = sFULLREAD.Split(sSeperator)
            Dim intOrder As Integer = 1

            For i As Integer = sParsed.GetLowerBound(0) To sParsed.GetUpperBound(0)
                If sParsed(i).Length > 0 Then
                    Dim cls As New clsHeaderColumn
                    cls.iOrder = intOrder
                    cls.strText = sParsed(i)
                    cls.iColumnInCIBMTR = -1

                    TheHeadersHashTable.Add(intOrder, cls) 'Note: .Key starts on 1.
                    intOrder += 1

                End If
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub BuildColumnsInclsExtractedLine()
        For Each de As DictionaryEntry In TheHashTable
            Dim EL As clsExtractedLine = de.Value
            Dim sValue As String = ""
            Dim bWeAreInsideQuotes As Boolean = False
            Dim iHashTableKey As Integer = 1

            For i As Integer = 1 To EL.sLine.Length
                Select Case Mid(EL.sLine, i, 1)
                    Case Chr(34)
                        bWeAreInsideQuotes = Not (bWeAreInsideQuotes)
                        If blnStripOutQuotesAndCommas = False Then sValue += Chr(34)
                    Case ","
                        If bWeAreInsideQuotes Then
                            If blnStripOutQuotesAndCommas = False Then sValue += Mid(EL.sLine, i, 1)
                        Else
                            EL.ColumnsHT.Add(iHashTableKey, sValue)
                            sValue = ""
                            iHashTableKey += 1
                        End If
                    Case Else
                        sValue += Mid(EL.sLine, i, 1)
                End Select
            Next
        Next
    End Sub


    Private Sub SolveForColumnOrder()

        Dim clsEL As clsExtractedLine = TheHashTable(0)

        For Each de As DictionaryEntry In clsEL.ColumnsHT
            For Each Headerde As DictionaryEntry In TheHeadersHashTable
                Dim clsHC As clsHeaderColumn = Headerde.Value

                If de.Value = clsHC.strText Then
                    clsHC.iColumnInCIBMTR = de.Key
                End If
            Next
        Next

        For Each de As DictionaryEntry In TheHeadersHashTable
            Dim clsHC As clsHeaderColumn = de.Value
            If clsHC.iColumnInCIBMTR = -1 Then
                MsgBox(clsHC.strText & " has not ben found in the CIBMTR file.", vbOKOnly, "SHUTTING DOWN!!")
                End
            End If
        Next

    End Sub


    Private Sub BuildOutputString()
        Try
            Dim sb As New System.Text.StringBuilder()

            For iTheHashTable As Integer = 0 To TheHashTable.Count - 1
                Dim cel As clsExtractedLine = TheHashTable(iTheHashTable)

                For i As Integer = 1 To TheHeadersHashTable.Count  'Starts on 1
                    Dim clsHeader As clsHeaderColumn = TheHeadersHashTable(i)
                    If i = TheHeadersHashTable.Count Then
                        sb.Append(cel.ColumnsHT(clsHeader.iColumnInCIBMTR) & vbCrLf)
                    Else
                        sb.Append(cel.ColumnsHT(clsHeader.iColumnInCIBMTR) & ",")
                    End If
                Next
            Next

            strOutputString = sb.ToString
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub SaveOutPut()
        My.Computer.FileSystem.WriteAllText(strLoadPath & ".out.txt", strOutputString, False)
    End Sub
    Private Class clsExtractedLine
        Friend sLine As String
        Friend ColumnsHT As New Hashtable
    End Class
    Private Class clsHeaderColumn
        Friend strText As String
        Friend iOrder As Integer
        Friend iColumnInCIBMTR As Integer
    End Class
End Class

