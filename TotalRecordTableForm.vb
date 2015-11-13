Public Class TotalRecordTableForm
  Private Sub tblRecord_Click(sender As Object, e As EventArgs) Handles tblRecord.Click
    pnlForTable.Focus()
  End Sub

  Private Sub pnlForTable_Click(sender As Object, e As EventArgs) Handles pnlForTable.Click
    pnlForTable.Focus()
  End Sub

  Private Sub lblCol1InItem1_Click(sender As Object, e As EventArgs) Handles lbl1_1.Click
    Dim label As Label = CType(sender, Label)
    Dim col As Integer = tblRecord.GetColumn(label.Parent)

    MessageBox.Show("click col: " & col.ToString)
  End Sub
End Class