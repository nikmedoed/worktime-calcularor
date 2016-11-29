Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json

Public Class Form1
    Dim cI, cS, cE As Integer
    Dim Alpha As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim GEx As New Excel.Application
    Dim BBook As _Workbook
    Dim Sheet As _Worksheet

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If OpenFileDialog1.ShowDialog() = vbOK Then
            BBook = GEx.Workbooks.Open(OpenFileDialog1.FileName)
            Sheet = BBook.Worksheets(1)
        Else
            End
        End If

        For i = 1 To Sheet.UsedRange.Columns.Count
            If Sheet.Cells(1, i).value = "Исполнитель" Then Letter_I.Text = iToStr(i)
            If Sheet.Cells(1, i).value = "В работе-начало работа с заявкой" Then Letter_S.Text = iToStr(i)
            If Sheet.Cells(1, i).value = "Реализована-задача закрыта" Then Letter_E.Text = iToStr(i)
        Next

    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        BBook.Close(False)
        Sheet = Nothing
        BBook = Nothing
        GEx.Quit()
        GEx = Nothing
    End Sub

    Function iToStr(i As Integer) As String
        Dim result As String
        result = ""
        If i > 26 Then result += Alpha(((i - 1) \ 26) - 1)
        result += Alpha(i Mod 26 - 1)
        iToStr = result
    End Function

    Function StrToi(s As String) As Integer
        Dim r, k As String
        r = Alpha.IndexOf(s(0)) + 1
        If Len(s) > 1 Then
            r = r * 26 + 1
            r += Alpha.IndexOf(s(1))
        End If
        StrToi = r
    End Function

    Structure Employe
        Dim name As String
        Dim tst As Date
        Dim tend As Date
        Dim wtime As List(Of Date())

        Public Sub add(st As Date, en As Date)
            Dim D(2) As Date
            D = {st, en}
            Dim a, b As Integer

            If wtime.Count = 0 Then
                wtime.Add(D)
            Else

                If en < wtime(1)(1) Then
                    'wtime.Insert()
                End If

                For i = 1 To wtime.Count - 1

                Next
                End If
        End Sub
    End Structure

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        cI = StrToi(Letter_I.Text)
        cS = StrToi(Letter_S.Text)
        cE = StrToi(Letter_E.Text)

        Dim k As New List(Of Integer)
        k.Add(1)
        k.Add(2)
        MsgBox(CStr(cI) + " " + CStr(cS) + " " + CStr(cE))

        MsgBox(k(1))

        For i = 2 To Sheet.UsedRange.Rows.Count
            ' напиши структуру
            ' собери в неёё инфу
            ' сделай эталонное время работы
            ' вырезай из него
            ' выделить для каждого края
            ' выделить общие края
        Next
    End Sub
End Class
