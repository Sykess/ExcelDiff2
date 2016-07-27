Imports Microsoft.Office.Core
Imports System.Configuration
Imports System.Text.RegularExpressions

Public Class Form1
    Dim myForm2 As Form2

    Public Class WorkerArgs
        Public sheetOne As String
        Public sheetTwo As String
        Public strRange As String
    End Class

    Private Sub FileOneButton_Click(sender As Object, e As EventArgs) Handles FileOneButton.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            FileOneComboBox.Text = OpenFileDialog1.FileName
        End If
        If Not FileOneComboBox.Items.Contains(FileOneComboBox.Text) Then
            FileOneComboBox.Items.Add(FileOneComboBox.Text)
        End If
    End Sub

    Private Sub FileToButton_Click(sender As Object, e As EventArgs) Handles FileToButton.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            FileTwoComboBox.Text = OpenFileDialog1.FileName
        End If
        If Not FileTwoComboBox.Items.Contains(FileTwoComboBox.Text) Then
            FileTwoComboBox.Items.Add(FileTwoComboBox.Text)
        End If
    End Sub

    Private Sub SaveButton_Click(sender As Object, e As EventArgs) Handles SaveButton.Click
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            SaveComboBox.Text = SaveFileDialog1.FileName
        End If
        If Not SaveComboBox.Items.Contains(SaveComboBox.Text) Then
            SaveComboBox.Items.Add(SaveComboBox.Text)
        End If
    End Sub

    Private Sub CompareButton_Click(sender As Object, e As EventArgs) Handles CompareButton.Click
        Dim strRange As String
        Dim myWrapper As WorkerArgs = New WorkerArgs()
        strRange = "A1:N250"

        myWrapper.sheetOne = FileOneComboBox.Text
        myWrapper.sheetTwo = FileTwoComboBox.Text
        myWrapper.strRange = strRange
        CompareButton.Enabled = False
        BackgroundWorker1.RunWorkerAsync(myWrapper)
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim args As WorkerArgs = e.Argument
        compareFiles(args.sheetOne, args.sheetTwo, args.strRange, e)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        CompareButton.Enabled = True
        If DisplayCheckBox.Checked Then
            myForm2 = New Form2
            myForm2.TextBox1.Text = e.Result.ToString()
            myForm2.Show()
        End If
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
    End Sub

    Private Sub compareFiles(sheetOne As String, sheetTwo As String, strRange As String, e As System.ComponentModel.DoWorkEventArgs)
        Dim iRow As Object, iCol As Object
        Dim xl As Microsoft.Office.Interop.Excel.Application
        Dim wb_1 As Microsoft.Office.Interop.Excel.Workbook
        Dim wb_2 As Microsoft.Office.Interop.Excel.Workbook
        Dim ws_1 As Object
        Dim ws_2 As Object
        Dim myList As New ArrayList()

        Try
            xl = GetObject(, "Excel.Application")
        Catch ex As Exception
            'xl = New Excel.Application
            xl = CreateObject("Excel.Application")
        End Try


        xl.DisplayAlerts = False
        wb_1 = xl.Workbooks.Open(sheetOne)
        wb_2 = xl.Workbooks.Open(sheetTwo)
        'ProgressBar1.Maximum = wb_1.Worksheets.Count
        Dim max_sheets As Integer = wb_1.Worksheets.Count
        Dim current_progress As Integer = 0
        Dim temp_progress As Integer
        'ProgressBar1.Value = 0
        BackgroundWorker1.ReportProgress(current_progress)

        For I = 1 To wb_1.Worksheets.Count
            ws_1 = wb_1.Worksheets(I)
            ws_2 = wb_2.Worksheets(I)

            For iCol = 1 To 14
                For iRow = 1 To 250
                    'Dim cellOne As String = Regex.Replace(ws_1.Cells(iRow, iCol).Value, "[^A-Za-z0-9\-/]", "")
                    'Dim cellOne As String = Regex.Replace(ws_1.Cells(iRow, iCol).Value.ToString, "[^\w\\-]", "")
                    'Dim cellTwo As String = Regex.Replace(ws_2.Cells(iRow, iCol).Value.ToString, "[^\w\\-]", "")
                    Dim cellOne As String = ws_1.Cells(iRow, iCol).Value
                    Dim cellTwo As String = ws_2.Cells(iRow, iCol).Value
                    'If CLEAN(ws_1.Cells(iRow, iCol).Value) = Microsoft.VisualBasic.CLEAN(ws_2.Cells(iRow, iCol).Value) Then
                    'they're the same
                    'If Replace(ws_1.Cells(iRow, iCol).ToString, " ", "") = Replace(ws_2.Cells(iRow, iCol).ToString, " ", "") Then
                    'If Replace(cellOne, " ", "") = Replace(cellTwo, " ", "") Then
                    If cellOne Is vbNullString Then
                        Continue For
                    End If
                    If cellTwo Is vbNullString Then
                        Continue For
                    End If
                    If Regex.Replace(cellOne, "[^A-Za-z0-9\-/]", "") = Regex.Replace(cellTwo, "[^A-Za-z0-9\-/]", "") Then
                    Else
                        'MsgBox("Different")
                        Dim difference As DifferenceInfo
                        difference.Row = iRow
                        difference.Column = iCol
                        'difference.Sheet = "Cell one: " & Replace(cellOne, " ", "") & vbCrLf & "Cell two: " & Replace(cellTwo, " ", "") & vbCrLf
                        difference.Sheet = wb_1.Worksheets(I).Name
                        myList.Add(difference)

                        'MsgBox("Different in Sheet" & difference.Sheet & " Row " & iRow & " Col " & iCol)
                    End If
                Next
            Next
            current_progress += 1
            temp_progress = CDbl(current_progress / max_sheets) * 100
            BackgroundWorker1.ReportProgress(temp_progress)
        Next

        If SaveCheckBox.Checked = True Then
            Try
                My.Computer.FileSystem.WriteAllText(SaveComboBox.Text,
                   "Differences between" & vbCrLf & sheetOne & " and" &
                   vbCrLf & sheetTwo & vbCrLf, False)
                My.Computer.FileSystem.WriteAllText(SaveComboBox.Text,
                   formatEntry("Sheet:", 30) & formatEntry("Col", 10) &
                   formatEntry("Row", 10) & vbCrLf, True)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            For Each obj In myList
                Try
                    My.Computer.FileSystem.WriteAllText(SaveComboBox.Text, formatEntry(obj.Sheet, 30), True)
                    My.Computer.FileSystem.WriteAllText(SaveComboBox.Text, formatEntry(obj.Column, 10), True)
                    My.Computer.FileSystem.WriteAllText(SaveComboBox.Text, formatEntry(obj.Row, 10) & vbCrLf, True)
                Catch ex As Exception
                    MsgBox("Exception writing to file " & ex.ToString)
                End Try
            Next
        End If

        'If DisplayCheckBox.Checked = True Then
        'myForm2 = New Form2
        'myForm2.TextBox1.Text 
        Dim return_string As String
        return_string = "Differences between" & vbCrLf & sheetOne & " and" &
                   vbCrLf & sheetTwo & vbCrLf
        return_string += formatEntry("Sheet:", 30) & formatEntry("Col", 10) &
                   formatEntry("Row", 10) & vbCrLf
        For Each obj In myList
            return_string += formatEntry(obj.Sheet, 30)
            return_string += formatEntry(Chr(obj.Column + 64), 10)
            return_string += formatEntry(obj.Row, 10) & vbCrLf
        Next
        e.Result = return_string
        'myForm2.Show()
        'End If

        'For Each obj In myList
        'MsgBox("Sheet " & obj.Sheet & " row " & obj.Row & " col " & obj.Column)
        'Next
        wb_1.Close()
        wb_2.Close()
        xl.Quit()
    End Sub

    Private Structure DifferenceInfo
        Public Row As Integer
        Public Column As Integer
        Public Sheet As String
    End Structure

    Function formatEntry(ByRef Text As String, ByVal Length As Integer) As String
        formatEntry = Microsoft.VisualBasic.Left(Text & Space(Length), Length)
    End Function

    Function convertToLetter(icol As Long) As String
        convertToLetter = ""
        Dim iMod As Long, iCounter As Long
        iMod = icol Mod 26
        iCounter = Int((icol - 1) / 26)

        If iCounter > 0 Then
            convertToLetter = Chr(iCounter + 64)
        End If

        If iMod = 0 Then
            convertToLetter = convertToLetter & "Z"
        Else
            convertToLetter = convertToLetter & Chr(iMod + 64)
        End If
    End Function

    Private Sub SaveCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles SaveCheckBox.CheckedChanged

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            If My.Settings.FileOne IsNot Nothing Then
                For Each i As String In My.Settings.FileOne
                    FileOneComboBox.Items.Add(i)
                Next
            End If
            FileOneComboBox.Text = My.Settings.FileOne.Item(My.Settings.FileOne.Count - 1)
        Catch ex As Exception
            My.Settings.FileOne = New Specialized.StringCollection()
        End Try

        Try
            If My.Settings.FileTwo IsNot Nothing Then
                For Each i As String In My.Settings.FileTwo
                    FileTwoComboBox.Items.Add(i)
                Next
            End If
            FileTwoComboBox.Text = My.Settings.FileTwo.Item(My.Settings.FileTwo.Count - 1)
        Catch ex As Exception
            My.Settings.FileTwo = New Specialized.StringCollection()
        End Try

        Try
            If My.Settings.SaveFile IsNot Nothing Then
                For Each i As String In My.Settings.SaveFile
                    SaveComboBox.Items.Add(i)
                Next
            End If
            SaveComboBox.Text = My.Settings.SaveFile.Item(My.Settings.SaveFile.Count - 1)

            'Dim appSettings = ConfigurationManager.AppSettings
            'Dim fileOne As String = appSettings("FileOne")
            'Dim fileTwo As String = appSettings("FileTwo")
            'Dim saveFile As String = appSettings("SaveFile")
            'Dim saveCheck As Boolean = appSettings("SaveCheck")
            'Dim displayCheck As Boolean = appSettings("DisplayCheck")
            'FileOneComboBox.Text = fileOne
            'FileTwoComboBox.Text = fileTwo
            'SaveComboBox.Text = saveFile
            'SaveCheckBox.Checked = saveCheck
            'DisplayCheckBox.Checked = displayCheck
        Catch ex As Exception
            'MessageBox.Show("Error reading app settings")
            My.Settings.SaveFile = New Specialized.StringCollection()
        End Try

        Try
            SaveCheckBox.Checked = My.Settings.saveCheck
            DisplayCheckBox.Checked = My.Settings.displayCheck
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Form1_Closing(sender As Object, e As FormClosingEventArgs) Handles MyBase.Closing
        Try
            If My.Settings.FileOne IsNot Nothing Then
                My.Settings.FileOne.Clear()
            End If

            If FileOneComboBox.Items IsNot Nothing Then
                For Each i As String In FileOneComboBox.Items
                    My.Settings.FileOne.Add(i)
                Next
            End If
        Catch ex As Exception
        End Try

        Try
            My.Settings.FileTwo.Clear()

            If FileTwoComboBox.Items IsNot Nothing Then
                My.Settings.FileTwo.Clear()
                For Each i As String In FileTwoComboBox.Items
                    My.Settings.FileTwo.Add(i)
                Next
            End If
        Catch ex As Exception
        End Try

        Try
            My.Settings.SaveFile.Clear()

            If SaveComboBox.Items IsNot Nothing Then
                My.Settings.SaveFile.Clear()
                For Each i As String In SaveComboBox.Items
                    My.Settings.SaveFile.Add(i)
                Next
            End If
        Catch ex As Exception
        End Try

        Try
            My.Settings.Save()
            'Dim cAppConfig As Configuration = ConfigurationManager.OpenExeConfiguration(Application.StartupPath & "\ExcelDiff.exe")
            'Dim appSettings As AppSettingsSection = cAppConfig.AppSettings
            'appSettings.Settings.Item("FileOne").Value = FileOneComboBox.Text
            'appSettings.Settings.Item("FileTwo").Value = FileTwoComboBox.Text
            'appSettings.Settings.Item("SaveFile").Value = SaveComboBox.Text
            'appSettings.Settings.Item("SaveCheck").Value = SaveCheckBox.Checked
            'appSettings.Settings.Item("DisplayCheck").Value = DisplayCheckBox.Checked
            'cAppConfig.Save(ConfigurationSaveMode.Modified)
        Catch ex As Exception
            'MessageBox.Show("Error has occurred while saving settings")
        End Try
        'My.Settings.FileOne = FileOneComboBox.Text
        'My.Settings.FileTwo = FileTwoComboBox.Text
        'My.Settings.Save()
        Try
            My.Settings.saveCheck = SaveCheckBox.Checked
            My.Settings.displayCheck = DisplayCheckBox.Checked
        Catch ex As Exception
        End Try
    End Sub

    Private Sub QuitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub TestTextToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TestTextToolStripMenuItem.Click
        myForm2 = New Form2
        myForm2.TextBox1.Text = "Differences between" & vbCrLf & "TestSheet" & " and" &
               vbCrLf & "TestSheet" & vbCrLf
        myForm2.TextBox1.Text += formatEntry("Sheet:", 30) & formatEntry("Col", 10) &
               formatEntry("Row", 10) & vbCrLf
        myForm2.TextBox1.SelectionStart = 0
        myForm2.Show()

    End Sub


End Class
