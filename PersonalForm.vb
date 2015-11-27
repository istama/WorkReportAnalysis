Public Class PersonalForm

  Public TabPageInfoList As List(Of MainForm.TabInfo)
  Public CurrentlyShowedSheetRecordManager As MainForm.SheetRecordManager

  Private Sub PersonalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    ShowTable()
  End Sub

  Private Sub ShowTable()
    If TabPageInfoList.Exists(Function(i) i.Name = tabInPersonalTab.SelectedTab.Text) Then
      Dim info As MainForm.TabInfo = TabPageInfoList.Find(Function(i) i.Name = tabInPersonalTab.SelectedTab.Text)
      Dim m As New MainForm.SheetRecordManager(info.SheetRecordController)
      CurrentlyShowedSheetRecordManager = m.ReadSheetRecord.LoadTable(Nothing)
    End If
  End Sub

  Private Sub tabInPersonalTab_PageChanged(sender As Object, e As EventArgs) Handles tabInPersonalTab.SelectedIndexChanged
    MainForm.ShowPersonalRecord(tabInPersonalTab.SelectedTab)
  End Sub
End Class