Public Class RecordTableForm
  Public Shared TABLE_ROW_COUNT = 32

  Private Sub tblRecord_Click(sender As Object, e As EventArgs) Handles tblRecord.Click
    pnlForTable.Focus()
  End Sub

  Private Sub pnlForTable_Click(sender As Object, e As EventArgs) Handles pnlForTable.Click
    pnlForTable.Focus()
  End Sub
End Class