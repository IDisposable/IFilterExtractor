Option Strict Off
Option Explicit On

Imports VB = Microsoft.VisualBasic

Friend Class TextExtract
    Inherits System.Windows.Forms.Form

    Private Const MAX_EXTRACT_TEXT As Integer = 64 * 1024 ' no more than 64K of raw text for a resume!

    Private Sub PickFile_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles PickFile.FileOk
        On Error GoTo ErrH

        Dim sFile As String = PickFile.FileName

        Dim te As ExtractTextLib.TextExtractor = New ExtractTextLib.TextExtractor
        Dim sText As String
        sText = te.ExtractText(sFile, MAX_EXTRACT_TEXT)
        Err.Clear()
        Me.FileLength.Text = Len(sText).ToString()
        Me.FileText.Text = sText

        Exit Sub
ErrH:
        MsgBox("Error extracting text from '" & sFile & "' Err=" & Err.Number & "-" & Err.Description & " in " & Err.Source)
    End Sub

    Private Sub ExitProgram_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitProgram.Click
        Me.Close()
    End Sub

    Private Sub Open_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Open.Click
        PickFile.ShowDialog()
    End Sub
End Class