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
            If tst > st Then tst = st
            If tend < en Then tend = en

            If wtime.Count = 0 Then
                wtime.Add(D)
            Else
                If en < wtime.First()(0) Then
                    wtime.Insert(0, D)
                ElseIf st > wtime.Last()(1) Then
                    wtime.Add(D)
                ElseIf en > wtime.Last()(1) And st < wtime.First()(0) Then
                    wtime.Clear()
                    wtime.Add(D)
                Else
                    For i = 1 To wtime.Count - 1
                        If wtime(i)(0) <= st Then ' нашли стартовую точку между стартовыми точками двух промежутков
                            For j = i To wtime.Count - 1
                                If wtime(j)(1) >= en Then ' нашли конечную точку между конечными точками двух промежутков
                                    If i = j Then '---- промежуток полностью поглощён 
                                        GoTo OK
                                    Else
                                        If wtime(i)(1) >= st Then ' стартовая точка лежит внутри
                                            If wtime(j)(0) <= en Then ' конечная точка лежит внутри
                                                wtime(i)(1) = wtime(j)(1)
                                                wtime.RemoveRange(i + 1, j - i) '---- промежуток соединил несколько промежутков 
                                                GoTo OK
                                            Else ' конечная точка лежит снаружи
                                                wtime(i)(1) = en
                                                wtime.RemoveRange(i + 1, j - i - 1) '---- промежуток соединил несколько промежутков и перенес конец
                                                GoTo OK
                                            End If
                                        Else ' стартовая точка лежит снаружи

                                            If wtime(j)(0) <= en Then ' конечная точка лежит внутри
                                                wtime(j)(0) = st
                                                wtime.RemoveRange(i + 1, j - i - 1) '---- промежуток соединил несколько промежутков 
                                                GoTo OK
                                            Else ' конечная точка лежит снаружи
                                                wtime.RemoveRange(i + 1, j - i - 1) '---- удалили покрытые промежутки и добавили новый
                                                wtime.Insert(i, D)
                                                GoTo OK
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            ' конечная точка за пределами
                            If wtime(i)(1) >= st Then ' стартовая точка лежит внутри
                                wtime(i)(1) = en
                                wtime.RemoveRange(i + 1, wtime.Count - 1 - i) '---- промежуток соединил несколько промежутков 
                                GoTo OK
                            Else ' стартовая точка лежит снаружи
                                wtime.RemoveRange(i + 1, wtime.Count - i - 1) '---- удалили покрытые промежутки и добавили новый
                                wtime.Insert(i, D)
                                GoTo OK
                            End If
                        End If
                    Next
                    ' стартовая точка лежит за пределами, но конец точно где-то внутри
                    For i = 1 To wtime.Count - 1
                        If wtime(i)(1) >= en Then
                            If wtime(i)(0) <= en Then ' конечная точка лежит внутри
                                wtime(i)(0) = st
                                wtime.RemoveRange(0, i - 1) '---- промежуток соединил несколько промежутков 
                                GoTo OK
                            Else ' конечная точка лежит снаружи
                                wtime.RemoveRange(0, i - 1) '---- удалили покрытые промежутки и добавили новый
                                wtime.Insert(0, D)
                                GoTo OK
                            End If
                        End If
                    Next
                    Debug.Print("Неожиданно " + st + "  " + en + " " + name)
                End If
            End If
OK:
        End Sub
    End Structure

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        cI = StrToi(Letter_I.Text)
        cS = StrToi(Letter_S.Text)
        cE = StrToi(Letter_E.Text)
        Dim de, ds As Date

        Debug.Print("Выбранные столбцы " + CStr(cI) + " " + CStr(cS) + " " + CStr(cE))

        Dim em As New List(Of Employe)
        Dim t As New Employe
        Dim el As Integer
        ProgressBar1.Maximum = Sheet.UsedRange.Rows.Count - 1

        de = Sheet.Cells(2, cE).Value
        ds = Sheet.Cells(2, cS).Value

        For i = 2 To Sheet.UsedRange.Rows.Count
            ProgressBar1.Value = i - 1
            t.name = Sheet.Cells(i, cI).Value
            el = em.FindIndex(Function(U As Employe)
                                  Return U.name = t.name
                              End Function
                              )

            t.tst = Sheet.Cells(i, cS).Value
            t.tend = Sheet.Cells(i, cE).Value
            If ds > t.tst Then ds = t.tst
            If de < t.tend Then de = t.tend

            If el = -1 Then
                ' создать
                t.wtime = New List(Of Date())
                em.Add(t)
                em.Last().add(t.tst, t.tend)
            Else
                em(el).add(t.tst, t.tend)
            End If

            Debug.Print(ds + " " + de + " " + CStr(em.Count))
            ' сделай эталонное время работы
            ' вырезай из него
            ' выделить общие края
        Next
    End Sub
End Class
