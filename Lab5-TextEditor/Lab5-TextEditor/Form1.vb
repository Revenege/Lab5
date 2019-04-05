Imports System.IO
Public Class FrmTextEditor

    'FORM LEVEL VARIABLES
    Dim filename As String = "" 'Stores the filename of the current project. Blank is treated as a new file
    Dim txtChanged As Boolean = False 'Text change flag. True if anything has been edited
    Dim formIsClosing As Boolean = False

    'Handles Sav As Function. Public Function so it can be called by both the Save As event handler
    'and the Save event handler(in case there is no file)
    Public Sub SaveAs()
        saveDialog.Filter = "Plain Text Files(*.txt) | *.txt" 'Sets the type of files being searched for
        If saveDialog.ShowDialog() = DialogResult.OK Then 'If the dialog has a valid file path, then save the file
            filename = saveDialog.FileName
            My.Computer.FileSystem.WriteAllText(filename, txtInput.Text, False) 'saves to the file, deleting previous contents
            txtChanged = False
            Me.Text = "Awesome Text Editor - " + filename
        End If
    End Sub

    'Public function for closing the form. In place to insure that the program only asks if your sure if you want o close the program once and not twice when clicking exit.
    Public Function CloseForm() As Boolean
        If txtChanged = True Then
            Dim selection = MsgBox("Are you sure you want to end the program? All unsaved progress will be lost", MsgBoxStyle.YesNo Or MsgBoxStyle.Exclamation, "Warning, Unsaved Changes!")
            If selection = vbYes Then
                Return True
            Else
                Return False
            End If
        Else
            Return True
        End If
    End Function


    'Event Handler for when the user clicks File -> New
    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        If txtChanged = True Then
            Dim selection = MsgBox("Are you sure you want to create a new document? All unsaved progress will be lost", MsgBoxStyle.YesNo Or MsgBoxStyle.Exclamation, "Warning, Unsaved Changes!")
            If selection = vbYes Then
                txtInput.Text = ""
                filename = ""
                Me.Text = "Awesome Text Editor - New File"
            End If
        Else
            txtInput.Text = ""
            filename = ""
            Me.Text = "Awesome Text Editor - New File"
        End If
    End Sub

    'Event Handler for when the user clicks File -> Open
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        If txtChanged = True Then
            Dim selection = MsgBox("Are you sure you want to open a new document? All unsaved progress will be lost", MsgBoxStyle.YesNo Or MsgBoxStyle.Exclamation, "Warning, Unsaved Changes!")
            If selection = vbYes Then
                openDialog.Filter = "Plain Text Files (*.txt) | *.txt" 'Sets the type of files being searched for
                If openDialog.ShowDialog() = DialogResult.OK Then 'if the dialog has valid results, then open the file and set the filename
                    filename = openDialog.FileName
                    txtInput.Text = ""
                    Using reader As StreamReader = New StreamReader(filename)
                        Do While (reader.Peek() > -1)
                            txtInput.AppendText(reader.ReadLine())
                            txtInput.AppendText(vbNewLine)
                        Loop
                    End Using

                    Me.Text = "Awesome Text Editor - New File"
                End If
            End If
        Else
            openDialog.Filter = "Plain Text Files (*.txt) | *.txt" 'Sets the type of files being searched for
            If openDialog.ShowDialog() = DialogResult.OK Then 'if the dialog has valid results, then open the file and set the filename
                filename = openDialog.FileName
                txtInput.Text = My.Computer.FileSystem.ReadAllText(filename) 'reading the file

                Me.Text = "Awesome Text Editor - New File"
            End If
        End If
    End Sub

    'Event Handler for when the user clicks File -> Exit
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        If (CloseForm()) Then
            formIsClosing = True
            Me.Close()
        End If
    End Sub

    'Event Handler for when the user clicks File -> Save
    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If filename = "" Then
            SaveAs()
        Else
            My.Computer.FileSystem.WriteAllText(filename, txtInput.Text, False)
        End If
    End Sub
    'Event Handler for when the user clicks File -> Save As
    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        SaveAs()
    End Sub

    'If Text has been changed, set our text changed variable to true.
    Private Sub TxtInput_TextChanged(sender As Object, e As EventArgs) Handles txtInput.TextChanged
        txtChanged = True
    End Sub

    'Allows program to close gracefullly 
    Private Sub FormClose(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If (formIsClosing = False) Then
            If txtChanged = True Then
                Dim selection = MsgBox("Are you sure you want to end the program? All unsaved progress will be lost", MsgBoxStyle.YesNo Or MsgBoxStyle.Exclamation, "Warning, Unsaved Changes!")
                If selection = vbNo Then
                    e.Cancel = True
                End If
            End If
        End If
    End Sub

    'Copy Contents of Select Location to Clipbaord
    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        txtInput.Copy()
    End Sub

    'Cut contents of selected location to Clipboard
    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        txtInput.Cut()
    End Sub

    'Paste Contents of Clipboard to selected location
    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        txtInput.Paste()
    End Sub

    'Shows the about text
    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        MsgBox("NETD 2202-01" + vbNewLine + "Scott Jenkins" + vbNewLine + "Lab 5")
    End Sub
End Class
