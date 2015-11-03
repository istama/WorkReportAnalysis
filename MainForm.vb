
Imports MP.Office
Imports AP = MP.Utils.Common.AppProperties
Imports MP.Utils.Common
Imports MP.Utils.Model
Imports WAP = MP.WorkReportAnalysis.App.WorkReportAnalysisProperties
Imports MP.WorkReportAnalysis.App
Imports MP.WorkReportAnalysis.Table
Imports MP.WorkReportAnalysis.Control
Imports MP.WorkReportAnalysis.Excel
Imports MP.WorkReportAnalysis.Model
Imports MP.WorkReportAnalysis.Layout

Public Class MainForm
  Public Structure TabInfo
    Dim Name As String
    Dim CreateTable As Action(Of TabPage, List(Of RowRecord))
  End Structure

  Private AppProps As PropertyManager = AP.MANAGER

  Private UserRecordManager As UserRecordManager
  Private UserInfoList As List(Of ExpandedUserInfo)

  Private ExcelProps As PropertyManager = WAP.MANAGER
  Private ExcelReader As ExcelReader

  Private InnerTabPageInfoListInPersonalTab As List(Of TabInfo)
  Private InnerTabPageInfoListInTotalTab As List(Of TabInfo)

  Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    SettingLog()
    MyLog.Write("Main Formを起動しました。")
    Try
      AutoUpdate()
    Catch ex As Exception
    End Try

    Try
      LoadAllUsers()

      InitInnerTabPageInPersonalTab()
      InitInnerTabPageInTotalTab()
      InitTableTitles()
      InitCBoxUserInfo()

      ExcelReader = New ExcelReader(FileFormat.GetYear)
      ExcelReader.Init()

      UserRecordManager = New UserRecordManager(ExcelReader, UserInfoList)
      LoadAllUserRecord()
    Catch ex As Exception
      MsgBox.ShowError(ex)
    End Try

  End Sub

  Private Sub SettingLog()
    MyLog.Log.DefaultFileLogWriter.Location = Logging.LogFileLocation.ExecutableDirectory
    MyLog.Log.DefaultFileLogWriter.Append = False
    If AppProps.GetValue(AP.KEY_WRITE_LOG) = "True" Then
      MyLog.LogMode = True
    Else
      MyLog.LogMode = False
    End If
  End Sub

  Private Sub AutoUpdate()
    If AppProps.GetValue(AP.KEY_ENABLE_AUTO_UPDATE) = "True" Then
      MyLog.Write("自動アップデートを開始します。")
      Dim updateManager As UpdateManager = New UpdateManager(FilePath.UpdateScriptPath(), FilePath.ReleaseVersionInfoFilePath())

      updateManager.GenerateDefaultUpdateBatchIfEmpty(AppProps.GetValue(AP.KEY_RELEASE_DIR_FOR_UPDATE), FilePath.ExcludeFileForUpdatePath())
      If updateManager.hasUpdated() Then
        MessageBox.Show("最新のバージョンに更新します。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Me.Close()
        System.Diagnostics.Process.Start(FilePath.UpdateScriptPath())
      End If
    Else
      My.Application.Log.WriteEntry("自動アップデートはオフです。")
    End If
  End Sub

  Private Sub LoadAllUsers()
    Dim a As SerializedAccessor = MySerialize.GenerateAccessor()
    Dim path As String = FilePath.UserinfoFilePath()

    UserInfoList = a.GetInfo(Of UserInfo)(path).ConvertAll(Function(user) New ExpandedUserInfo(user))
  End Sub

  Private Sub InitInnerTabPageInPersonalTab()
    InnerTabPageInfoListInPersonalTab = New List(Of TabInfo)()
    For m As Integer = 10 To 12
      Dim tab As TabInfo
      With tab
        .Name = FileFormat.GetSheetName(m)
        .CreateTable = GetActionForCreatingMonthlyTable(FileFormat.GetYear, m)
      End With
      InnerTabPageInfoListInPersonalTab.Add(tab)
    Next

    Dim sumTab As TabInfo
    With sumTab
      .Name = FileFormat.GetSheetName2
      .CreateTable = AddressOf CreateSumTable
    End With
    InnerTabPageInfoListInPersonalTab.Add(sumTab)

    If tabInPersonalTab.TabPages.Count = InnerTabPageInfoListInPersonalTab.Count Then
      For idx As Integer = 0 To InnerTabPageInfoListInPersonalTab.Count - 1
        tabInPersonalTab.TabPages.Item(idx).Text = InnerTabPageInfoListInPersonalTab(idx).Name
      Next
    Else
      Throw New Exception("Excelファイルのシート数とタブページの数が合いません。")
    End If

  End Sub

  Private Sub InitInnerTabPageInTotalTab()
    InnerTabPageInfoListInTotalTab = New List(Of TabInfo)()
    Dim tabDays As TabInfo
    With tabDays
      .Name = "日"
      .CreateTable = GetActionForCreatingTotalDaysTable()
    End With

    Dim tabWeeks As TabInfo
    With tabWeeks
      .Name = "週"
      .CreateTable = GetActionForCreatingTotalDaysTable()
    End With

    Dim tabMonths As TabInfo
    With tabMonths
      .Name = "月"
      .CreateTable = GetActionForCreatingTotalDaysTable()
    End With

    Dim tabYear As TabInfo
    With tabYear
      .Name = "合計"
      .CreateTable = GetActionForCreatingTotalDaysTable()
    End With

    InnerTabPageInfoListInTotalTab.Add(tabDays)
    InnerTabPageInfoListInTotalTab.Add(tabWeeks)
    InnerTabPageInfoListInTotalTab.Add(tabMonths)
    InnerTabPageInfoListInTotalTab.Add(tabYear)

    If tabInTotalTab.TabPages.Count = InnerTabPageInfoListInTotalTab.Count Then
      For idx As Integer = 0 To InnerTabPageInfoListInTotalTab.Count - 1
        tabInTotalTab.TabPages.Item(idx).Text = InnerTabPageInfoListInTotalTab(idx).Name
      Next
    Else
      Throw New Exception("Excelファイルのシート数とタブページの数が合いません。")
    End If
  End Sub

  Private Sub InitTableTitles()
    RecordTableForm.lblItem1.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME1)
    RecordTableForm.lblItem2.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME2)
    RecordTableForm.lblItem3.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME3)
    RecordTableForm.lblItem4.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME4)
    RecordTableForm.lblItem5.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME5)
    RecordTableForm.lblItem6.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME6)
    RecordTableForm.lblItem7.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME7)

    TotalRecordTableForm.lblItem1.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME1)
    TotalRecordTableForm.lblItem2.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME2)
    TotalRecordTableForm.lblItem3.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME3)
    TotalRecordTableForm.lblItem4.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME4)
    TotalRecordTableForm.lblItem5.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME5)
    TotalRecordTableForm.lblItem6.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME6)
    TotalRecordTableForm.lblItem7.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME7)

    ShowPersonalTableTitles(page10Month)
    ShowTotalTableTitles(pageDays)
  End Sub

  Private Sub InitCBoxUserInfo()
    UserInfoList.ForEach(Function(user) cboxUserInfo.Items.Add(user))
  End Sub

  Private Sub LoadAllUserRecord()
    Dim res As DialogResult = MessageBox.Show("全てのExcelファイルを読み込みますか？" & vbCrLf & "読み込みには時間がかかるかもしれません。", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
    If res = DialogResult.OK Then
      ProgressBarForm.UserRecordManager = UserRecordManager
      ProgressBarForm.ShowDialog()
    End If
  End Sub

  Private Sub cmdReadAllFile_Click(sender As Object, e As EventArgs) Handles cmdReadAllFile.Click
    LoadAllUserRecord()
  End Sub

  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    ExcelReader.Quit()
    Me.Close()
  End Sub

  '閉じるボタンを無効にする
  Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
    Const WM_SYSCOMMAND As Integer = &H112
    Const SC_CLOSE As Long = &HF060L

    If m.Msg = WM_SYSCOMMAND AndAlso
        (m.WParam.ToInt64() And &HFFF0L) = SC_CLOSE Then
      Return
    End If

    MyBase.WndProc(m)
  End Sub

  Private Sub tabMaster_PageChanged(sender As Object, e As EventArgs) Handles tabMaster.SelectedIndexChanged

  End Sub

  Private Sub tabInPersonalTab_PageChanged(sender As Object, e As EventArgs) Handles tabInPersonalTab.SelectedIndexChanged
    ShowPersonalTableTitles(tabInPersonalTab.SelectedTab)
    LoadTabPageInPersonalTab(tabInPersonalTab.SelectedTab, cboxUserInfo.SelectedItem)
  End Sub

  Private Sub cboxUserInfo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxUserInfo.SelectedIndexChanged
    LoadTabPageInPersonalTab(tabInPersonalTab.SelectedTab, cboxUserInfo.SelectedItem)
  End Sub

  Private Sub tabInTotalTab_PageChanged(sender As Object, e As EventArgs) Handles tabInTotalTab.SelectedIndexChanged
    ShowTotalTableTitles(tabInTotalTab.SelectedTab)
    LoadTabPageInTotalTab(tabInTotalTab.SelectedTab)
  End Sub

  Private Sub LoadTabPageInPersonalTab(selectedTabPage As TabPage, selectedUser As ExpandedUserInfo)
    If selectedUser IsNot Nothing Then
      Try
        Dim sheetRecord As SheetRecord =
          UserRecordManager.ReadUserRecord(selectedUser).
          GetSheetRecord(selectedTabPage.Text)

        InnerTabPageInfoListInPersonalTab.
          Find(Function(info) info.Name = selectedTabPage.Text).
          CreateTable(selectedTabPage, sheetRecord.GetAll())
      Catch ex As Exception
        MsgBox.ShowError(ex)
      End Try
    End If
  End Sub

  Private Sub LoadTabPageInTotalTab(selectedTabPage As TabPage)
    Try
      InnerTabPageInfoListInTotalTab.
        Find(Function(info) info.Name = selectedTabPage.Text).
        CreateTable(selectedTabPage, Nothing)
    Catch ex As Exception
      MsgBox.ShowError(ex)
    End Try
  End Sub

  Private Sub CreateMonthlyTable(tabPage As TabPage, record As List(Of RowRecord), year As Integer, month As Integer)
    Dim emptyRows As New List(Of Integer)
    For i As Integer = Date.DaysInMonth(year, month) To 30
      emptyRows.Add(i)
    Next

    Dim newRecord As List(Of RowRecord) = RecordArranger.InsertEmpty(record, False, emptyRows.ToArray)
    Dim header As Func(Of Integer, String) = Function(row) If(row < 31, (row + 1).ToString & "日", "合計")
    Dim panelColor As Func(Of Integer, Color) = GetFuncForGettingRowColorInMonthlyTable(year, month)
    Dim tablePanel As Func(Of String, Integer, Integer, Panel) = GetFuncForCreatingTablePanel(panelColor, RecordTableForm.tblRecord.ColumnCount)

    CreateTable(RecordTableForm.tblRecord, tabPage, newRecord, header, panelColor, tablePanel)
    ShowTableRecord(tabPage)
  End Sub

  Private Sub CreateSumTable(tabPage As TabPage, record As List(Of RowRecord))
    Dim tmp As List(Of RowRecord) = RecordArranger.InsertEmpty(record, True, 0, 6, 12)
    Dim newRecord As List(Of RowRecord) = RecordArranger.PadTailWithEmpty(tmp, False, RecordTableForm.tblRecord.RowCount)
    Dim header As Func(Of Integer, String) = AddressOf GetHeaderTextInSumTable
    Dim panelColor As Func(Of Integer, Color) = AddressOf GetRowColorInSumTable
    Dim tablePanel As Func(Of String, Integer, Integer, Panel) = GetFuncForCreatingTablePanel(panelColor, RecordTableForm.tblRecord.ColumnCount)

    CreateTable(RecordTableForm.tblRecord, tabPage, newRecord, header, panelColor, tablePanel)
    ShowTableRecord(tabPage)
  End Sub

  Private Sub CreateTotalTable(tabPage As TabPage, record As List(Of RowRecord))
    Dim newRecord As List(Of RowRecord) = RecordArranger.PadTailWithEmpty(record, False, TotalRecordTableForm.tblRecord.RowCount)
    Dim header As Func(Of Integer, String) = Function(row) (row + 1).ToString.PadLeft(3, "0"c)
    Dim panelColor As Func(Of Integer, Color) = Function(row) Color.Transparent
    Dim tablePanel As Func(Of String, Integer, Integer, Panel) = GetFuncForCreatingTablePanel(panelColor, RecordTableForm.tblRecord.ColumnCount)

    CreateTable(TotalRecordTableForm.tblRecord, tabPage, newRecord, header, panelColor, tablePanel)
    ShowTotalTableRecord(tabPage)
  End Sub

  Private Sub CreateTable(table As TableLayoutPanel, tabPage As TabPage, record As List(Of RowRecord), HeaderText As Func(Of Integer, String), PanelColor As Func(Of Integer, Color), TablePanel As Func(Of String, Integer, Integer, Panel))

    Dim tableRow As Integer = 0

    For rowIdx As Integer = 0 To (record.Count - 1)
      Dim rr As RowRecord = record(rowIdx)

      Dim tableCol As Integer = 0

      For colIdx As Integer = 0 To (table.ColumnCount - 1)
        Dim value As String
        If colIdx >= rr.List.Count Then
          value = ""
        ElseIf rr.List(colIdx) = ExcelAccessor.ROW_RECORD_HEADER Then
          value = HeaderText(rowIdx)
        Else
          value = RoundDecStr(rr.List(colIdx))
        End If

        Dim control As Control = table.GetControlFromPosition(tableCol, tableRow)
        If control Is Nothing Then
          Dim panel As Panel = TablePanel(value, colIdx, rowIdx)
          table.Controls.Add(panel, tableCol, tableRow)
        Else
          If control.Controls.Count = 0 Then
            Dim panel As Panel = TablePanel(value, colIdx, rowIdx)
            table.SetCellPosition(panel, table.GetCellPosition(control))
          Else
            control.BackColor = PanelColor(rowIdx)
            control.Controls.Item(0).Text = value
          End If
        End If

        tableCol += 1
      Next

      tableRow += 1
    Next

  End Sub

  Private Function GetActionForCreatingMonthlyTable(year As Integer, month As Integer) As Action(Of TabPage, List(Of RowRecord))
    Return _
      Sub(tabPage, record)
        CreateMonthlyTable(tabPage, record, year, month)
      End Sub
  End Function

  Private Function GetActionForCreatingTotalDaysTable() As Action(Of TabPage, List(Of RowRecord))
    Return _
      Sub(tabPage, record)
        Dim m As Integer = dateTotal.Value.Month
        Dim d As Integer = dateTotal.Value.Day
        Dim r As List(Of RowRecord) = UserRecordManager.GetTotalRecordAt(m, d)
        CreateTotalTable(tabPage, r)
      End Sub
  End Function

  Private Function GetFuncForCreatingTablePanel(PanelColor As Func(Of Integer, Color), colSize As Integer) As Func(Of String, Integer, Integer, Panel)
    Return _
      Function(value, col, row)
        Dim color As Color = PanelColor(row)
        Return CreateTablePanel(value, color, col, colSize)
      End Function
  End Function

  Private Function GetFuncForGettingRowColorInMonthlyTable(year As Integer, month As Integer) As Func(Of Integer, Color)
    Return _
      Function(row)
        If row = 31 Then
          Return Color.PaleGreen
        Else
          Dim week As Integer = GetWeek(year, month, row + 1)
          Return If(week = 1 OrElse week = 7, Color.Pink, Color.Transparent)
        End If
      End Function
  End Function

  Private Function CreateTablePanel(value As String, color As Color, colIdx As Integer, colSize As Integer) As Panel
    If colIdx = 0 Then
      Return ControlDrawer.CreateTextPanelInTable(value, color)
    ElseIf colIdx = colSize - 1
      Return ControlDrawer.CreateNotePanelInTable(value, color)
    Else
      Return ControlDrawer.CreateNumberPanelInTable(value, color)
    End If
  End Function

  Private Function GetHeaderTextInSumTable(row As Integer) As String
    If row = 21 Then
      Return "合計"
    ElseIf row Mod 7 = 0 Then
      Return (row / 7 + 10).ToString & "月"
    ElseIf row Mod 7 = 6
      Return "月計"
    Else
      Return "第" & (row Mod 7).ToString & "週"
    End If
  End Function

  Private Function GetRowColorInSumTable(row As Integer) As Color
    If row = 21 Then
      Return Color.PaleTurquoise
    ElseIf row Mod 7 = 6 AndAlso row < 21
      Return Color.PaleGreen
    Else
      Return Color.Transparent
    End If
  End Function

  Private Function RoundDecStr(value As String) As String
    If Char.IsDigit(value) Then
      Dim d As Double = Double.Parse(value)
      Return Math.Round(d, 2).ToString
    Else
      Return value
    End If
  End Function

  Private Sub ShowPersonalTableTitles(showedTabPage As TabPage)
    Dim tblTitles As TableLayoutPanel = RecordTableForm.tblTitles
    tblTitles.Location = New Point(3, 6)
    Dim tblSubTitles As TableLayoutPanel = RecordTableForm.tblSubTItles
    tblSubTitles.Location = New Point(3, 32)

    showedTabPage.Controls.Add(tblTitles)
    showedTabPage.Controls.Add(tblSubTitles)
  End Sub

  Private Sub ShowTotalTableTitles(showedTabPage As TabPage)
    Dim tblTitles As TableLayoutPanel = TotalRecordTableForm.tblTitles
    tblTitles.Location = New Point(3, 46)
    Dim tblSubTitles As TableLayoutPanel = TotalRecordTableForm.tblSubTItles
    tblSubTitles.Location = New Point(3, 72)

    showedTabPage.Controls.Add(tblTitles)
    showedTabPage.Controls.Add(tblSubTitles)
  End Sub

  Private Sub ShowTableRecord(showedTabPage As TabPage)
    Dim scrollPanel As Panel = RecordTableForm.pnlForTable
    scrollPanel.Location = New Point(3, 63)
    scrollPanel.AutoScroll = True

    showedTabPage.Controls.Add(scrollPanel)
  End Sub

  Private Sub ShowTotalTableRecord(showedTabPage As TabPage)
    Dim scrollPanel As Panel = TotalRecordTableForm.pnlForTable
    scrollPanel.Location = New Point(3, 103)
    scrollPanel.AutoScroll = True

    showedTabPage.Controls.Add(scrollPanel)
  End Sub

  Private Function GetWeek(year As Integer, month As Integer, day As Integer) As Integer
    If month >= 1 AndAlso month <= 12 AndAlso day <= Date.DaysInMonth(year, month) Then
      Return Weekday(year & "/" & month & "/" & day)
    Else
      Return -1
    End If
  End Function

End Class
