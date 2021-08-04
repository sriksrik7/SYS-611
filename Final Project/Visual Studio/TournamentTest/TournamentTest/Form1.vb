Imports Excel = Microsoft.Office.Interop.Excel


'Run python script
'https://stackoverflow.com/questions/22961625/visual-basic-windows-form-running-python-script-w-button

Public Class Form1
    'C:\Users\Adrian\Desktop\test2.xlsx

    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim tdelay = 250

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Butt_Generate.Click
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("C:\Users\Adrian\Desktop\test2.xlsx")            ' WORKBOOK TO OPEN THE EXCEL FILE.
        xlWorkSheet = xlWorkBook.Worksheets("Sheet1")       ' NAME OF THE WORK SHEET.

        LblP01.Text = xlWorkSheet.Range("A1").Value
        LblP02.Text = xlWorkSheet.Range("B1").Value
        LblP03.Text = xlWorkSheet.Range("A2").Value
        LblP04.Text = xlWorkSheet.Range("B2").Value
        LblP05.Text = xlWorkSheet.Range("A3").Value
        LblP06.Text = xlWorkSheet.Range("B3").Value
        LblP07.Text = xlWorkSheet.Range("A4").Value
        LblP08.Text = xlWorkSheet.Range("B4").Value

        xlWorkBook.Close() : xlApp.Quit()

        ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing
    End Sub




    Private Sub Butt_Run_Click(sender As Object, e As EventArgs) Handles Butt_Run.Click
        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open("C:\Users\Adrian\Desktop\test2.xlsx")            ' WORKBOOK TO OPEN THE EXCEL FILE.
        xlWorkSheet = xlWorkBook.Worksheets("Sheet1")       ' NAME OF THE WORK SHEET.


        'Button1_Click()

        'LblP1.Text = xlWorkSheet.Range("A1").Value

        If xlWorkSheet.Range("C1").Value = 1 Then
            G01a.Image = TournamentTest.My.Resources.green
        Else
            G01a.Image = TournamentTest.My.Resources.red
        End If
        System.Threading.Thread.Sleep(tdelay)
        If xlWorkSheet.Range("D1").Value = 1 Then
            G01b.Image = TournamentTest.My.Resources.green
        Else
            G01b.Image = TournamentTest.My.Resources.red
        End If
        System.Threading.Thread.Sleep(tdelay)
        If xlWorkSheet.Range("E1").Value = 1 Then
            G01c.Image = TournamentTest.My.Resources.green
        Else
            G01c.Image = TournamentTest.My.Resources.red
        End If
        System.Threading.Thread.Sleep(tdelay)
        If xlWorkSheet.Range("F1").Value = 1 Then
            G01d.Image = TournamentTest.My.Resources.green
        Else
            G01d.Image = TournamentTest.My.Resources.red
        End If
        System.Threading.Thread.Sleep(tdelay)

        If xlWorkSheet.Range("G1").Value = 1 Then
            G01e.Image = TournamentTest.My.Resources.green
        Else
            G01e.Image = TournamentTest.My.Resources.red
        End If

        'Write in winner of 1st game
        System.Threading.Thread.Sleep(500)
        LblP9.Text = xlWorkSheet.Range("A3").Value.ToString


        'Second Game
        If xlWorkSheet.Range("C2").Value = 1 Then
            G02a.Image = TournamentTest.My.Resources.green
        Else
            G02a.Image = TournamentTest.My.Resources.red
        End If
        System.Threading.Thread.Sleep(tdelay)
        If xlWorkSheet.Range("D2").Value = 1 Then
            G02b.Image = TournamentTest.My.Resources.green
        Else
            G02b.Image = TournamentTest.My.Resources.red
        End If
        System.Threading.Thread.Sleep(tdelay)
        If xlWorkSheet.Range("E2").Value = 1 Then
            G02c.Image = TournamentTest.My.Resources.green
        Else
            G02c.Image = TournamentTest.My.Resources.red
        End If
        System.Threading.Thread.Sleep(tdelay)
        If xlWorkSheet.Range("F2").Value = 1 Then
            G02d.Image = TournamentTest.My.Resources.green
        Else
            G02d.Image = TournamentTest.My.Resources.red
        End If
        System.Threading.Thread.Sleep(tdelay)
        If xlWorkSheet.Range("G2").Value = 1 Then
            G02e.Image = TournamentTest.My.Resources.green
        Else
            G02e.Image = TournamentTest.My.Resources.red
        End If
        System.Threading.Thread.Sleep(tdelay)

        'Write in winner of 1st game
        LblP10.Text = xlWorkSheet.Range("B3").Value.ToString

        'Finals Game
        If xlWorkSheet.Range("C3").Value = 1 Then
            G03a.Image = TournamentTest.My.Resources.green
        Else
            G03a.Image = TournamentTest.My.Resources.red
        End If
        System.Threading.Thread.Sleep(tdelay)
        If xlWorkSheet.Range("D3").Value = 1 Then
            G03b.Image = TournamentTest.My.Resources.green
        Else
            G03b.Image = TournamentTest.My.Resources.red
        End If
        System.Threading.Thread.Sleep(tdelay)
        If xlWorkSheet.Range("E3").Value = 1 Then
            G03c.Image = TournamentTest.My.Resources.green
        Else
            G03c.Image = TournamentTest.My.Resources.red
        End If
        System.Threading.Thread.Sleep(tdelay)
        If xlWorkSheet.Range("F3").Value = 1 Then
            G03d.Image = TournamentTest.My.Resources.green
        Else
            G03d.Image = TournamentTest.My.Resources.red
        End If
        System.Threading.Thread.Sleep(tdelay)
        If xlWorkSheet.Range("G3").Value = 1 Then
            G03e.Image = TournamentTest.My.Resources.green
        Else
            G03e.Image = TournamentTest.My.Resources.red
        End If

        'Write in champion
        System.Threading.Thread.Sleep(tdelay)
        LblChamp.Text = xlWorkSheet.Range("A4").Value.ToString



        xlWorkBook.Close() : xlApp.Quit()
        ' CLEAN UP. (CLOSE INSTANCES OF EXCEL OBJECTS.)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp) : xlApp = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook) : xlWorkBook = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet) : xlWorkSheet = Nothing
    End Sub

    Dim PictureBoxList As New List(Of PictureBox)
    Dim selcnt As Integer = 0
    Dim blk() As PictureBox
    Dim zz() As Label




    'Control array for Picture Boxes
    Private Sub Butt_Test_Click(sender As Object, e As EventArgs) Handles Butt_Test.Click


        blk = New PictureBox() {G03a, G03b, G03c, G03d, G03e}
        'zz = New Label() {LblP0, LblP6, LblP7, LblP8, LblX}

        'blk(1).Image = TournamentTest.My.Resources.red
        For index As Integer = 1 To 5
            'zz(index - 1).Text = "Text here"
            blk(index - 1).Image = TournamentTest.My.Resources.green
        Next


    End Sub



    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub


    Private Sub Butt_Exit_Click(sender As Object, e As EventArgs) Handles Butt_Exit.Click

        'Me.Close()
        Application.Exit()

    End Sub


End Class
