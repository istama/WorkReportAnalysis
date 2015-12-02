
Imports MP.Utils.Common

Public Class PersonalForm

  Public TabPageInfoList As List(Of MainForm.TabInfo)
  Public CurrentlyShowedSheetRecordManager As MainForm.SheetRecordManager

  Private Sub PersonalForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    ShowTable()
  End Sub

  Private Sub ShowTable()
    If TabPageInfoList.Exists(Function(i) i.Name = tabInPersonalTab.SelectedTab.Text) Then
      LoadRecordInTabPage(tabInPersonalTab.SelectedTab)
    End If
  End Sub


  Private Sub tabInPersonalTab_PageChanged(sender As Object, e As EventArgs) Handles tabInPersonalTab.SelectedIndexChanged
    LoadRecordInTabPage(tabInPersonalTab.SelectedTab)
  End Sub

  Private Sub LoadRecordInTabPage(selectedInnerTab As TabPage)
    Try
      If TabPageInfoList.Exists(Function(e) e.TableRecordController.TableLayout.TabPage.Equals(selectedInnerTab)) Then
        Dim tabInfo As MainForm.TabInfo = TabPageInfoList.Find(Function(e) e.TableRecordController.TableLayout.TabPage.Equals(selectedInnerTab))
        ShowRecord(tabInfo)
      Else
        Throw New Exception("タブ情報が見つかりません")
      End If
    Catch ex As Exception
      MsgBox.ShowError(ex)
    End Try
  End Sub

  Private Sub ShowRecord(tabInfo As MainForm.TabInfo)
    'MessageBox.Show("tabinfo name: " & tabInfo.Name)
    'MessageBox.Show("tabinfo csv: " & tabInfo.TableRecordController.CSV.GetCSVFileNameCallback())
    Dim recManager As MainForm.SheetRecordManager = New MainForm.SheetRecordManager(tabInfo.TableRecordController)
    LoadTableInnerTabPage(recManager.ReadSheetRecord)
  End Sub

  Private Sub LoadTableInnerTabPage(recordManager As MainForm.SheetRecordManager)
    CurrentlyShowedSheetRecordManager = recordManager.LoadTable()
  End Sub

End Class