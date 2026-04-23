Public Class Form1

    Private Sub btnBrowse_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowse.Click
        Using dlg As New OpenFileDialog()
            dlg.Multiselect = True
            If dlg.ShowDialog() = DialogResult.OK Then
                For Each filePath As String In dlg.FileNames
                    lstFiles.Items.Add(filePath).Selected = True
                Next
            End If
        End Using
    End Sub

    Private Sub btnGo_Click(sender As System.Object, e As System.EventArgs) Handles btnGo.Click
        If txtNewName.Text = "" Then
            MessageBox.Show("You must enter a new filename to continue", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        If boxOptions.CheckedIndices.Count = 0 Then
            MessageBox.Show("You must choose one of the naming schemes to continue", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        If lstFiles.SelectedIndices.Count < 1 Then
            MessageBox.Show("You must choose some files to rename before we can get this party started.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If boxOptions.GetItemChecked(5) Then
            findWhichIsChecked()
        Else
            Dim numbering As Integer = 1
            For Each item As ListViewItem In lstFiles.SelectedItems
                Dim fullPath As String = item.Text
                If Not My.Computer.FileSystem.FileExists(fullPath) Then
                    item.Remove()
                    MessageBox.Show("That filename no longer exists, so I removed it for you.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Continue For
                End If
                Dim extStart As Integer = fullPath.LastIndexOf(".")
                Dim extension As String = If(extStart >= 0 AndAlso fullPath.Length - extStart <= 5, fullPath.Substring(extStart), "")
                Dim newName As String = ""
                Dim valid As Boolean = True
                Select Case True
                    Case boxOptions.GetItemChecked(0) : newName = txtNewName.Text & " " & numbering & extension
                    Case boxOptions.GetItemChecked(1) : newName = txtNewName.Text & "(" & numbering & ")" & extension
                    Case boxOptions.GetItemChecked(2) : newName = txtNewName.Text & " - " & numbering & extension
                    Case boxOptions.GetItemChecked(3)
                        If txtPrefix.Text = "" Then
                            MessageBox.Show("Try actually entering a prefix.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            valid = False
                        Else
                            newName = txtPrefix.Text & " " & txtNewName.Text & "Copy " & numbering & extension
                        End If
                    Case boxOptions.GetItemChecked(4)
                        If txtSuffix.Text = "" Then
                            MessageBox.Show("Try actually entering a suffix.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            valid = False
                        Else
                            newName = txtNewName.Text & " " & txtSuffix.Text & "Copy " & numbering & extension
                        End If
                End Select
                If valid Then
                    Try
                        My.Computer.FileSystem.RenameFile(fullPath, newName)
                    Catch ex As Exception
                        MessageBox.Show("Could not rename """ & fullPath & """: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                    numbering += 1
                End If
            Next
        End If

        txtNewName.Clear()
        txtPrefix.Clear()
        txtSuffix.Clear()
    End Sub

    Private Sub boxOptions_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles boxOptions.SelectedIndexChanged
        Dim selection = boxOptions.SelectedIndex
        If boxOptions.CheckedIndices.Count > 1 And boxOptions.GetItemChecked(5) = False Then
            For Each item In boxOptions.CheckedIndices
                If boxOptions.GetItemChecked(item) = True Then
                    boxOptions.SetItemCheckState(item, CheckState.Unchecked)
                    boxOptions.SetItemCheckState(boxOptions.SelectedIndex, CheckState.Checked)
                End If
            Next
        ElseIf boxOptions.GetItemChecked(5) = True And boxOptions.SelectedIndex < 5 = True Then
            For Each item In boxOptions.CheckedIndices
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

    Private Sub boxOptions_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles boxOptions.ItemCheck
        BeginInvoke(New Action(AddressOf UpdateConditionalFields))
    End Sub

    Private Sub UpdateConditionalFields()
        txtExtension.Visible = boxOptions.GetItemChecked(5)
        txtPrefix.Visible = boxOptions.GetItemChecked(3)
        txtSuffix.Visible = boxOptions.GetItemChecked(4)
    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        txtPrefix.Visible = False
        txtSuffix.Visible = False
        txtExtension.Visible = False
    End Sub

    Private Sub findWhichIsChecked()
        If txtExtension.Text = "" Then
            MessageBox.Show("Try actually entering a file extension.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim numbering As Integer = 1
        For Each item As ListViewItem In lstFiles.SelectedItems
            Dim fullPath As String = item.Text
            If Not My.Computer.FileSystem.FileExists(fullPath) Then
                Continue For
            End If
            Dim newName As String = ""
            Dim valid As Boolean = True
            Select Case True
                Case boxOptions.GetItemChecked(0) : newName = txtNewName.Text & " " & numbering & "." & txtExtension.Text
                Case boxOptions.GetItemChecked(1) : newName = txtNewName.Text & "(" & numbering & ")" & "." & txtExtension.Text
                Case boxOptions.GetItemChecked(2) : newName = txtNewName.Text & " - " & numbering & "." & txtExtension.Text
                Case boxOptions.GetItemChecked(3)
                    If txtPrefix.Text = "" Then
                        MessageBox.Show("Try actually entering a prefix.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        valid = False
                    Else
                        newName = txtPrefix.Text & " " & txtNewName.Text & " (Copy " & numbering & ")" & "." & txtExtension.Text
                    End If
                Case boxOptions.GetItemChecked(4)
                    If txtSuffix.Text = "" Then
                        MessageBox.Show("Try actually entering a suffix.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        valid = False
                    Else
                        newName = txtNewName.Text & " " & txtSuffix.Text & " (Copy " & numbering & ")" & "." & txtExtension.Text
                    End If
                Case boxOptions.GetItemChecked(5) : newName = txtNewName.Text & "." & txtExtension.Text
            End Select
            If valid Then
                Try
                    My.Computer.FileSystem.RenameFile(fullPath, newName)
                Catch ex As Exception
                    MessageBox.Show("Could not rename """ & fullPath & """: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
                numbering += 1
            End If
        Next
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
        If rdoLeading.Checked Then
            For Each item As ListViewItem In lstFiles.SelectedItems
                Dim fullPath As String = item.Text
                Dim beginFileName As Integer = fullPath.LastIndexOf("\") + 1
                Dim fileName As String = fullPath.Substring(beginFileName)
                Dim extensionStart As Integer = fileName.LastIndexOf(".")
                Dim nameLength As Integer = If(extensionStart >= 0, extensionStart, fileName.Length)
                If numLeading.Value <= nameLength Then
                    Dim choppedName As String = fileName.Remove(0, CInt(numLeading.Value))
                    Try
                        My.Computer.FileSystem.RenameFile(fullPath, choppedName)
                    Catch ex As Exception
                        MessageBox.Show("Could not rename """ & fullPath & """: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                Else
                    MessageBox.Show("Try entering a more realistic number of characters to delete.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    numLeading.Value = 0
                End If
            Next
        ElseIf rdoTrailing.Checked Then
            For Each item As ListViewItem In lstFiles.SelectedItems
                Dim fullPath As String = item.Text
                Dim beginFileName As Integer = fullPath.LastIndexOf("\") + 1
                Dim fileName As String = fullPath.Substring(beginFileName)
                Dim extensionStart As Integer = fileName.LastIndexOf(".")
                Dim nameLength As Integer = If(extensionStart >= 0, extensionStart, fileName.Length)
                If numTrailing.Value <= nameLength Then
                    Dim choppedName As String = fileName.Remove(nameLength - CInt(numTrailing.Value), CInt(numTrailing.Value))
                    Try
                        My.Computer.FileSystem.RenameFile(fullPath, choppedName)
                    Catch ex As Exception
                        MessageBox.Show("Could not rename """ & fullPath & """: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                Else
                    MessageBox.Show("Try entering a more realistic number of characters to delete.", "Dumbass", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    numTrailing.Value = 0
                End If
            Next
        End If
    End Sub

End Class
