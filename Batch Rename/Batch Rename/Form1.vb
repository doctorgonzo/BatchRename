Imports System.IO
Public Class Form1
    Dim x As New OpenFileDialog
    Dim bannedChars() As String = {"?", "/", "\", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "=", "+", "`", "~", "[", "]", "{", "}", "|", ".", ",", "<", ">", ":", "'", ";"}
    'Dim counter As Integer = -1
    'Dim numbering As Integer = 1
    Private Sub btnBrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowse.Click
        x.Reset()
        x.Multiselect = True
        x.ShowDialog()
        If DialogResult.OK Then
            Dim selectedFiles() As String = x.FileNames
            For Each item In selectedFiles
                lstFiles.Items.Add(item).Selected = True
            Next
        End If
    End Sub

    Private Sub btnGo_Click(sender As System.Object, e As System.EventArgs) Handles btnGo.Click
        If txtNewName.Text = "" Then
            MessageBox.Show("You must enter a new filename to continue", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf boxOptions.GetItemChecked(0) = False And boxOptions.GetItemChecked(1) = False And boxOptions.GetItemChecked(2) = False And boxOptions.GetItemChecked(3) = False And boxOptions.GetItemChecked(4) = False And boxOptions.GetItemChecked(5) = False Then
            MessageBox.Show("You must choose one of the naming schemes to continue", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf lstFiles.SelectedIndices.Count < 1 Then
            MessageBox.Show("You must choose some files to rename before we can get this party started.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim counter As Integer = -1
            Dim numbering As Integer = 1
            Dim extensionStart As Integer = x.FileName.LastIndexOf(".")
            Dim extensionEnd As Integer = x.FileName.Count
            If extensionEnd - extensionStart <= 5 And boxOptions.GetItemChecked(5) = False Then          
                Dim extension As String = x.FileName.Substring(extensionStart, extensionEnd - extensionStart)
                For Each item In lstFiles.SelectedItems
                    If My.Computer.FileSystem.FileExists(lstFiles.SelectedItems.Item(counter + 1).Text) Then
                        Select Case True
                            Case boxOptions.GetItemChecked(0) = True : My.Computer.FileSystem.RenameFile((lstFiles.SelectedItems.Item(counter + 1).Text), txtNewName.Text & " " & numbering & extension)
                                counter += 1
                                numbering += 1
                            Case boxOptions.GetItemChecked(1) = True : My.Computer.FileSystem.RenameFile((lstFiles.SelectedItems.Item(counter + 1).Text), txtNewName.Text & "(" & numbering & ")" & extension)
                                counter += 1
                                numbering += 1
                            Case boxOptions.GetItemChecked(2) = True : My.Computer.FileSystem.RenameFile((lstFiles.SelectedItems.Item(counter + 1).Text), txtNewName.Text & " - " & numbering & extension)
                                counter += 1
                                numbering += 1
                            Case boxOptions.GetItemChecked(3) = True : If txtPrefix.Text <> Nothing Then My.Computer.FileSystem.RenameFile((lstFiles.SelectedItems.Item(counter + 1).Text), txtPrefix.Text & " " & txtNewName.Text & "Copy " & numbering & extension) Else MessageBox.Show("Try actually entering a prefix.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                counter += 1
                                numbering += 1
                            Case boxOptions.GetItemChecked(4) = True : If txtSuffix.Text <> Nothing Then My.Computer.FileSystem.RenameFile((lstFiles.SelectedItems.Item(counter + 1).Text), txtNewName.Text & " " & txtSuffix.Text & "Copy " & numbering & extension) Else MessageBox.Show("try actually entering a suffix.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                counter += 1
                                numbering += 1
                        End Select
                    ElseIf My.Computer.FileSystem.FileExists(lstFiles.SelectedItems.Item(counter + 1).Text) = False Then
                        lstFiles.SelectedItems.Item(counter + 1).Remove()
                        MessageBox.Show("That filename no longer exists, so I removed it for you.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Next
            ElseIf boxOptions.GetItemChecked(5) = True Then
                findWhichIsChecked()
            ElseIf boxOptions.GetItemChecked(5) = False And extensionEnd - extensionStart > 5 Then
                MessageBox.Show("You seem to have a file (or files) in your list with no extension. You'll have to set an extension manually before continuing.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
        txtNewName.Clear()
        txtPrefix.Clear()
        txtSuffix.Clear()
    End Sub

    Private Sub boxOptions_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles boxOptions.SelectedIndexChanged
        Dim selection = boxOptions.SelectedIndex
        If boxOptions.CheckedIndices.Count > 1 And boxOptions.GetItemChecked(5) = False Then
            For Each item In boxOptions.CheckedIndices
                boxOptions.GetItemChecked(item)
                If boxOptions.GetItemChecked(item) = True Then
                    boxOptions.SetItemCheckState(item, CheckState.Unchecked)
                    boxOptions.SetItemCheckState(boxOptions.SelectedIndex, CheckState.Checked)
                End If
            Next
        ElseIf boxOptions.GetItemChecked(5) = True And boxOptions.SelectedIndex < 5 = True Then
            For Each item In boxOptions.CheckedIndices
                boxOptions.GetItemChecked(item)
                If boxOptions.GetItemChecked(item) = True Then
                    boxOptions.SetItemCheckState(item, CheckState.Unchecked)
                    If selection > -1 Then
                        boxOptions.SetItemCheckState(selection, CheckState.Checked)
                        boxOptions.SetItemCheckState(5, CheckState.Checked)
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Timer1.Start()
        txtPrefix.Visible = False
        txtSuffix.Visible = False
        txtExtension.Visible = False
    End Sub

    Private Sub findWhichIsChecked()
        Dim counter As Integer = -1
        Dim numbering As Integer = 1
        If boxOptions.GetItemChecked(5) = True Then
            For Each item In lstFiles.SelectedItems
                If My.Computer.FileSystem.FileExists(lstFiles.SelectedItems.Item(counter + 1).Text) Then
                    Select Case True
                        Case boxOptions.GetItemChecked(0) = True : If txtExtension.Text <> Nothing Then My.Computer.FileSystem.RenameFile((lstFiles.SelectedItems.Item(counter + 1).Text), txtNewName.Text & " " & numbering & "." & txtExtension.Text) Else MessageBox.Show("Try actually entering a file extension.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            counter += 1
                            numbering += 1
                        Case boxOptions.GetItemChecked(1) = True : If txtExtension.Text <> Nothing Then My.Computer.FileSystem.RenameFile((lstFiles.SelectedItems.Item(counter + 1).Text), txtNewName.Text & "(" & numbering & ")" & "." & txtExtension.Text) Else MessageBox.Show("Try actually entering a file extension.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            counter += 1
                            numbering += 1
                        Case boxOptions.GetItemChecked(2) = True : If txtExtension.Text <> Nothing Then My.Computer.FileSystem.RenameFile((lstFiles.SelectedItems.Item(counter + 1).Text), txtNewName.Text & " - " & numbering & "." & txtExtension.Text) Else MessageBox.Show("Try actually entering a file extension.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            counter += 1
                            numbering += 1
                        Case boxOptions.GetItemChecked(3) = True : If txtPrefix.Text <> Nothing Then My.Computer.FileSystem.RenameFile((lstFiles.SelectedItems.Item(counter + 1).Text), txtPrefix.Text & " " & txtNewName.Text & " (Copy " & numbering & ")" & "." & txtExtension.Text) Else MessageBox.Show("Try actually entering a prefix.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            counter += 1
                            numbering += 1
                        Case boxOptions.GetItemChecked(4) = True : If txtSuffix.Text <> Nothing Then My.Computer.FileSystem.RenameFile((lstFiles.SelectedItems.Item(counter + 1).Text), txtNewName.Text & " " & txtSuffix.Text & " (Copy " & numbering & ")" & "." & txtExtension.Text) Else MessageBox.Show("Try actually entering a suffix.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            counter += 1
                            numbering += 1
                        Case boxOptions.GetItemChecked(5) = True : If txtExtension.Text <> Nothing Then My.Computer.FileSystem.RenameFile((lstFiles.SelectedItems.Item(counter + 1).Text), txtNewName.Text & "." & txtExtension.Text) Else MessageBox.Show("Try actually entering a file extension.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            counter += 1
                    End Select
                End If
            Next
        End If
    End Sub

    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click
        txtExtension.Clear()
        txtNewName.Clear()
        txtPrefix.Clear()
        txtSuffix.Clear()
        lstFiles.Clear()
        lstFiles.SelectedItems.Clear()
        numLeading.Value = 0
        numTrailing.Value = 0
        rdoLeading.Checked = False
        rdoTrailing.Checked = False
        chkRemoveChars.Checked = False
        For i = 0 To boxOptions.Items.Count - 1
            boxOptions.SetItemCheckState(i, CheckState.Unchecked)
        Next
        boxOptions.ClearSelected()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        If boxOptions.GetItemCheckState(5) = CheckState.Checked Then
            txtExtension.Show()
        Else
            txtExtension.Hide()
        End If
        If boxOptions.GetItemChecked(3) = True Then
            txtPrefix.Show()
        Else
            txtPrefix.Hide()
        End If
        If boxOptions.GetItemChecked(4) = True Then
            txtSuffix.Show()
        Else
            txtSuffix.Hide()
        End If
    End Sub

    Private Sub chkRemoveChars_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkRemoveChars.CheckedChanged
        If chkRemoveChars.Checked = True Then
            txtExtension.Enabled = False
            txtSuffix.Enabled = False
            txtPrefix.Enabled = False
            btnChop.Visible = True
            btnGo.Enabled = False
            txtNewName.Enabled = False
            boxOptions.Enabled = False
            rdoLeading.Visible = True
            rdoTrailing.Visible = True
        ElseIf chkRemoveChars.Checked = False Then
            txtExtension.Enabled = True
            txtSuffix.Enabled = True
            txtPrefix.Enabled = True
            btnChop.Visible = False
            btnGo.Enabled = True
            txtNewName.Enabled = True
            boxOptions.Enabled = True
            lblLeading.Visible = False
            lblTrailing.Visible = False
            rdoLeading.Visible = False
            rdoTrailing.Visible = False
            numLeading.Visible = False
            numTrailing.Visible = False
        End If
    End Sub

    Private Sub rdoLeading_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rdoLeading.CheckedChanged
        If rdoLeading.Checked = True Then
            numLeading.Visible = True
            lblLeading.Visible = True
        ElseIf rdoLeading.Checked = False Then
            numLeading.Visible = False
            lblLeading.Visible = False
        End If
    End Sub

    Private Sub rdoTrailing_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rdoTrailing.CheckedChanged
        If rdoTrailing.Checked = True Then
            numTrailing.Visible = True
            lblTrailing.Visible = True
        ElseIf rdoTrailing.Checked = False Then
            numTrailing.Visible = False
            lblTrailing.Visible = False
        End If
    End Sub

    Private Sub btnChop_Click(sender As System.Object, e As System.EventArgs) Handles btnChop.Click
        chopLead()
    End Sub

    Private Sub chopLead()
        Dim counter As Integer = -1
        If rdoLeading.Checked = True Then
            For Each item In lstFiles.SelectedItems
                Dim beginFileName As Integer = lstFiles.SelectedItems.Item(counter + 1).Text.LastIndexOf("\") + 1
                Dim firstCharacter As Integer = Val(lstFiles.SelectedItems.Item(counter + 1).Text.First)
                Dim startName = lstFiles.SelectedItems.Item(counter + 1).Text
                Dim finalName = startName.Remove(firstCharacter, beginFileName)
                Dim extensionStart = finalName.LastIndexOf(".")
                Dim extensionEnd = finalName.Count
                Dim extensionFinal = extensionEnd - extensionStart
                If numLeading.Value <= (finalName.Count - extensionFinal) Then
                    Dim choppedName = finalName.Remove(0, numLeading.Value)
                    My.Computer.FileSystem.RenameFile(lstFiles.SelectedItems.Item(counter + 1).Text, choppedName)
                    counter += 1
                Else
                    MessageBox.Show("Try entering a more realistic number of characters to delete.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    numLeading.Value = 0
                End If
            Next
        ElseIf rdoTrailing.Checked = True Then
            For Each item In lstFiles.SelectedItems
                Dim beginFileName As Integer = lstFiles.SelectedItems.Item(counter + 1).Text.LastIndexOf("\") + 1
                Dim firstCharacter As Integer = Val(lstFiles.SelectedItems.Item(counter + 1).Text.First)
                Dim startName = lstFiles.SelectedItems.Item(counter + 1).Text
                Dim finalName = startName.Remove(firstCharacter, beginFileName)
                Dim extensionStart = finalName.LastIndexOf(".")
                Dim extensionEnd = finalName.Count
                Dim extensionFinal = extensionEnd - extensionStart
                If numTrailing.Value <= (finalName.Count - extensionFinal) Then
                    Dim choppedName = finalName.Remove(extensionStart - numTrailing.Value, numTrailing.Value)
                    My.Computer.FileSystem.RenameFile(lstFiles.SelectedItems.Item(counter + 1).Text, choppedName)
                    counter += 1
                Else
                    MessageBox.Show("Try entering a more realistic number of characters to delete.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    numTrailing.Value = 0
                End If
            Next
        End If
    End Sub
End Class


