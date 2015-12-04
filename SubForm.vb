
Imports MP.Utils.Common

Public Class SubForm
  Public Tab As TabControl
  Public TabPageInfoList As List(Of MainForm.TabInfo)

  Private Loaded As Boolean = False

  Private Sub PersonalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    ShowTable()
    Loaded = True
  End Sub

  Private Sub ShowTable()
    If TabPageInfoList.Exists(Function(i) i.Name = Tab.SelectedTab.Text) Then
      MainForm.LoadRecordInTabPage(TabPageInfoList, Tab.SelectedTab)
    End If
  End Sub

  Private Sub chkExcludeIncompleteRecordFromSum_CheckedChanged(sender As Object, e As EventArgs) Handles chkExcludeIncompleteRecordFromSum.CheckedChanged
    If Loaded Then
      Call ShowTable()
    End If
  End Sub

  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Me.Close()
  End Sub

  Private Sub cmdOutputCSV_Click(sender As Object, e As EventArgs) Handles cmdOutputCSV.Click
    MainForm.SaveCSVFile()
  End Sub

End Class