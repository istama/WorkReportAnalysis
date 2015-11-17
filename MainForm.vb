
Imports MP.Office
Imports AP = MP.Utils.Common.AppProperties
Imports MP.Utils.Common
Imports MP.Utils.Model
Imports WAP = MP.WorkReportAnalysis.App.WorkReportAnalysisProperties
Imports MP.WorkReportAnalysis.App
Imports MP.WorkReportAnalysis.Control
Imports MP.WorkReportAnalysis.Excel
Imports MP.WorkReportAnalysis.Model
Imports RowRecord = MP.Utils.MyCollection.Immutable.MyLinkedList(Of String)
Imports SheetRecord = MP.Utils.MyCollection.Immutable.MyLinkedList(Of MP.Utils.MyCollection.Immutable.MyLinkedList(Of String))
Imports MP.WorkReportAnalysis.Layout

Public Class MainForm
  Public Structure TabInfo
    Dim Name As String
    Dim LoadTableCallBack As LoadTableCallBack
  End Structure

  Public Structure SortInfo
    Dim Col As Integer
    Dim Asc As Boolean
  End Structure

  Private AppProps As PropertyManager = AP.MANAGER

  Private UserRecordLoader As UserRecordLoader
  Private UserRecordManager As UserRecordManager
  Private UserInfoList As List(Of ExpandedUserInfo)
  Private ReadRecordTerm As ReadRecordTerm

  Private ExcelProps As PropertyManager = WAP.MANAGER
  Private Excel As ExcelAccessor

  Private InnerTabPageInfoListInPersonalTab As List(Of TabInfo)
  Private InnerTabPageInfoListInTermTab As List(Of TabInfo)
  Private InnerTabPageInfoListInTotalTab As List(Of TabInfo)

  Private CurrentlyShowedSheetRecord As SheetRecord = Nothing
  Private SortProp As SortInfo

  Private Loaded As Boolean = False

  Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    SettingLog()
    MyLog.Write("Main Formを起動しました。")
    Try
      AutoUpdate()
    Catch ex As Exception
    End Try

    Try
      LoadAllUsers()

      ReadRecordTerm = New ReadRecordTerm(10, FileFormat.GetYear, 12, FileFormat.GetYear)
      Excel = New ExcelAccessor(New Excel())
      Excel.Init()
      UserRecordManager = New UserRecordManager(UserInfoList, ReadRecordTerm)
      UserRecordLoader = New UserRecordLoader(Excel, UserRecordManager)

      InitInnerTabPageInPersonalTab()
      InitInnerTabPageInTermTab()
      InitInnerTabPageInTotalTab()
      InitTableTitles()
      InitCBoxUserInfo()
      InitCBoxWeeklyTerm()
      InitCBoxMonthlyTerm()
      InitCBoxDailyTotal()

      LoadAllUserRecord()
    Catch ex As Exception
      MsgBox.ShowError(ex)
    End Try

    Loaded = True
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

    For Each ym As Tuple(Of Integer, Integer) In ReadRecordTerm.GetTermList.ToList
      Dim tab As TabInfo
      With tab
        .Name = UserRecordManager.GetSheetName(ym.Item2)
        .LoadTableCallBack = GetActionForCreatingMonthlyTable(ym.Item1, ym.Item2)
      End With
      InnerTabPageInfoListInPersonalTab.Add(tab)
    Next

    Dim sumTab As TabInfo
    With sumTab
      .Name = UserRecordManager.GetSumSheetName
      .LoadTableCallBack = GetActionForCreatingSumTable()
    End With
    InnerTabPageInfoListInPersonalTab.Add(sumTab)

    SetTabPageName(tabInPersonalTab, InnerTabPageInfoListInPersonalTab)
  End Sub

  Private Sub InitInnerTabPageInTermTab()
    InnerTabPageInfoListInTermTab = New List(Of TabInfo)()
    Dim tabDays As TabInfo
    With tabDays
      .Name = "日"
      .LoadTableCallBack = GetActionForCreatingDailyTermTable()
    End With

    Dim tabWeeks As TabInfo
    With tabWeeks
      .Name = "週"
      .LoadTableCallBack = GetActionForCreatingWeeklyTermTable()
    End With

    Dim tabMonths As TabInfo
    With tabMonths
      .Name = "月"
      .LoadTableCallBack = GetActionForCreatingMonthlyTermTable()
    End With

    Dim tabYear As TabInfo
    With tabYear
      .Name = "合計"
      .LoadTableCallBack = GetActionForCreatingAllTermTable()
    End With

    InnerTabPageInfoListInTermTab.Add(tabDays)
    InnerTabPageInfoListInTermTab.Add(tabWeeks)
    InnerTabPageInfoListInTermTab.Add(tabMonths)
    InnerTabPageInfoListInTermTab.Add(tabYear)

    SetTabPageName(tabInTermTab, InnerTabPageInfoListInTermTab)
  End Sub

  Private Sub InitInnerTabPageInTotalTab()
    InnerTabPageInfoListInTotalTab = New List(Of TabInfo)()
    Dim tabDays As TabInfo
    With tabDays
      .Name = "日"
      .LoadTableCallBack = GetActionForCreatingDailyTotalTable()
    End With

    Dim tabWeeks As TabInfo
    With tabWeeks
      .Name = "週"
      .LoadTableCallBack =
        GetActionForCreatingPlaneTable(
          Function() UserRecordManager.GetWeeklyTotalRecord(FileFormat.GetYear, GetFuncForFilteringImcompleteRecord()))
    End With

    Dim tabMonths As TabInfo
    With tabMonths
      .Name = "月"
      .LoadTableCallBack =
        GetActionForCreatingPlaneTable(
          Function() UserRecordManager.GetWeeklyTotalRecord(FileFormat.GetYear, GetFuncForFilteringImcompleteRecord()))
    End With

    InnerTabPageInfoListInTotalTab.Add(tabDays)
    InnerTabPageInfoListInTotalTab.Add(tabWeeks)
    InnerTabPageInfoListInTotalTab.Add(tabMonths)

    SetTabPageName(tabInTotalTab, InnerTabPageInfoListInTotalTab)
  End Sub

  Private Sub SetTabPageName(tab As TabControl, tabPageInfoList As List(Of TabInfo))
    If tab.TabPages.Count = tabPageInfoList.Count Then
      For idx As Integer = 0 To tabPageInfoList.Count - 1
        tab.TabPages.Item(idx).Text = tabPageInfoList(idx).Name
      Next
    Else
      Throw New Exception(
        "Excelファイルのシート数とタブページの数が合いません。 / tabPageCount: " &
        tab.TabPages.Count.ToString &
        " tabInfoCount: " & tabPageInfoList.Count)
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

    TermRecordTableForm.lblItem1.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME1)
    TermRecordTableForm.lblItem2.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME2)
    TermRecordTableForm.lblItem3.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME3)
    TermRecordTableForm.lblItem4.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME4)
    TermRecordTableForm.lblItem5.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME5)
    TermRecordTableForm.lblItem6.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME6)
    TermRecordTableForm.lblItem7.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME7)

    TotalRecordTableForm.lblItem1.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME1)
    TotalRecordTableForm.lblItem2.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME2)
    TotalRecordTableForm.lblItem3.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME3)
    TotalRecordTableForm.lblItem4.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME4)
    TotalRecordTableForm.lblItem5.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME5)
    TotalRecordTableForm.lblItem6.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME6)
    TotalRecordTableForm.lblItem7.Text = ExcelProps.GetValue(WAP.KEY_ITEM_NAME7)

    ShowPersonalTableTitles(page10Month)
    ShowTermTableTitles(pageDays)
  End Sub

  Private Sub InitCBoxUserInfo()
    UserInfoList.ForEach(Function(user) cboxUserInfo.Items.Add(user))
  End Sub

  Private Sub InitCBoxWeeklyTerm()
    Dim items As New List(Of DateItem)
    ReadRecordTerm.DateList.ForEach(
      Sub(t)
        For i As Integer = 1 To 5
          items.Add(New WeeklyItem(i, t.Item2, t.Item1))
        Next
      End Sub)

    items.ForEach(Sub(i) cboxWeeklyTerm.Items.Add(i))

    'Dim current As WeeklyItem = items.Find(Function(i) i.Agree(Date.Today))
    'If current IsNot Nothing Then
    '  Dim idx As Integer = cboxWeeklyTerm.Items.IndexOf(current)
    '  cboxWeeklyTerm.SelectedIndex = idx
    'End If
  End Sub

  Private Sub InitCBoxMonthlyTerm()
    Dim items As New List(Of MonthlyItem)
    ReadRecordTerm.DateList.ForEach(
      Sub(t) items.Add(New MonthlyItem(t.Item2, t.Item1)))

    items.ForEach(Sub(i) cboxMonthlyTerm.Items.Add(i))

    'Dim current As MonthlyItem = items.Find(Function(i) i.Agree(Date.Today))
    'If current IsNot Nothing Then
    '  Dim idx As Integer = cboxMonthlyTerm.Items.IndexOf(current)
    '  cboxMonthlyTerm.SelectedIndex = idx
    'End If
    'ReadRecordTerm.DateList.ForEach(
    '  Sub(t) cboxMonthlyTerm.Items.Add(New MonthlyItem(t.Item2, t.Item1)))
  End Sub

  Private Sub InitCBoxDailyTotal()
    Dim items As New List(Of MonthlyItem)
    ReadRecordTerm.DateList.ForEach(
      Sub(t) items.Add(New MonthlyItem(t.Item2, t.Item1)))

    items.ForEach(Sub(i) cboxDailyTotal.Items.Add(i))
  End Sub

  Private Sub InitDPicDailyTerm()
    Dim min As Tuple(Of Integer, Integer) = ReadRecordTerm.DateList.First
    dPicDailyTerm.MinDate = New DateTime(min.Item1, min.Item2, 1, 0, 0, 0)
    Dim max As Tuple(Of Integer, Integer) = ReadRecordTerm.DateList.ToList.Last
    dPicDailyTerm.MaxDate = New DateTime(max.Item1, max.Item2, Date.DaysInMonth(max.Item1, max.Item2), 0, 0, 0)
  End Sub

  Private Sub LoadAllUserRecord()
    Dim res As DialogResult = MessageBox.Show("全てのExcelファイルを読み込みますか？" & vbCrLf & "読み込みには時間がかかるかもしれません。", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
    If res = DialogResult.OK Then
      ProgressBarForm.UserRecordLoader = UserRecordLoader
      ProgressBarForm.ShowDialog()
    End If
  End Sub

  Private Sub cmdReadAllFile_Click(sender As Object, e As EventArgs) Handles cmdReadAllFile.Click
    LoadAllUserRecord()
  End Sub

  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Excel.Quit()
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
    Dim idx As Integer = tabMaster.SelectedIndex
    If idx = 0 Then
      ShowPersonalRecord()
    ElseIf idx = 1
      ShowTermRecord()
    Else
      ShowTotalRecord()
    End If
  End Sub

  Private Sub tabInPersonalTab_PageChanged(sender As Object, e As EventArgs) Handles tabInPersonalTab.SelectedIndexChanged
    ShowPersonalRecord()
  End Sub

  Private Sub cboxUserInfo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxUserInfo.SelectedIndexChanged
    LoadInnerTabPage(tabInPersonalTab.SelectedTab, InnerTabPageInfoListInPersonalTab)
  End Sub

  Private Sub tabInTermTab_PageChanged(sender As Object, e As EventArgs) Handles tabInTermTab.SelectedIndexChanged
    ShowTermRecord()
  End Sub

  Private Sub tabInTotalTab_PageChanged(sender As Object, e As EventArgs) Handles tabInTotalTab.SelectedIndexChanged
    ShowTotalRecord()
  End Sub

  Private Sub datePickerDailyTerm_DateChanged(sender As Object, e As EventArgs) Handles dPicDailyTerm.ValueChanged
    If ReadRecordTerm.InTerm(dPicDailyTerm.Value.Month, dPicDailyTerm.Value.Year) Then
      ShowTermRecord()
    End If
  End Sub

  Private Sub cboxWeeklyTerm_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxWeeklyTerm.SelectedIndexChanged
    If cboxWeeklyTerm.SelectedItem IsNot Nothing Then
      ShowTermRecord()
    End If
  End Sub

  Private Sub cboxMonthlyTerm_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxMonthlyTerm.SelectedIndexChanged
    If cboxMonthlyTerm.SelectedItem IsNot Nothing Then
      ShowTermRecord()
    End If
  End Sub

  Private Sub cboxDailyTotal_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxDailyTotal.SelectedIndexChanged
    If cboxDailyTotal.SelectedItem IsNot Nothing Then
      ShowTotalRecord()
    End If
  End Sub

  Private Sub chkExcludeIncompleteRecordFromSum_CheckedChanged(sender As Object, e As EventArgs) Handles chkExcludeIncompleteRecordFromSum.CheckedChanged
    If Loaded Then
      Call tabMaster_PageChanged(sender, e)
    End If
  End Sub

  Private Sub ShowPersonalRecord()
    ShowPersonalTableTitles(tabInPersonalTab.SelectedTab)
    LoadInnerTabPage(tabInPersonalTab.SelectedTab, InnerTabPageInfoListInPersonalTab)
  End Sub

  Private Sub ShowTermRecord()
    ShowTermTableTitles(tabInTermTab.SelectedTab)
    LoadInnerTabPage(tabInTermTab.SelectedTab, InnerTabPageInfoListInTermTab)
  End Sub

  Private Sub ShowTotalRecord()
    ShowTotalTableTitles(tabInTotalTab.SelectedTab)
    LoadInnerTabPage(tabInTotalTab.SelectedTab, InnerTabPageInfoListInTotalTab)
  End Sub

  Private Sub LoadInnerTabPage(selectedTabPage As TabPage, tabInfoList As List(Of TabInfo))
    Try
      Dim sortInfo As SortInfo
      With sortInfo
        .Col = -1
        .Asc = True
      End With

      tabInfoList.
        Find(Function(info) info.Name = selectedTabPage.Text).
        LoadTableCallBack(selectedTabPage, sortInfo)
    Catch ex As Exception
      MsgBox.ShowError(ex)
    End Try
  End Sub

  Private Sub CreateDailyTableInMonth(tabPage As TabPage, record As SheetRecord, year As Integer, month As Integer)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(RecordTableForm.pnlForTable).
        SetTextCols(0).
        SetNoteCols(RecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(GetFuncForGettingRowColorInMonthlyTable(year, month))

    Dim sumrow As RowRecord = CreateSumRowRecord(record, "合計")
    Dim nRecord As SheetRecord = record.AddLast(sumrow)
    CreateTable(RecordTableForm.tblRecord, tabPage, nRecord, tableDrawer)
    ShowTableRecord(tabPage)
    CurrentlyShowedSheetRecord = nRecord
  End Sub

  Private Sub CreateSumTable(tabPage As TabPage, record As SheetRecord)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(RecordTableForm.pnlForTable).
        SetTextCols(0).
        SetNoteCols(RecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(AddressOf GetRowColorInSumTable)

    'Dim sumrow As RowRecord = CreateSumRowRecord(record, "合計")
    'Dim nRecord As SheetRecord = record.AddLast(sumrow)
    CreateTable(RecordTableForm.tblRecord, tabPage, record, tableDrawer)
    ShowTableRecord(tabPage)
    CurrentlyShowedSheetRecord = record
  End Sub

  Private Sub CreateTermTable(tabPage As TabPage, record As SheetRecord)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(TermRecordTableForm.pnlForTable).
        SetTextCols(0, 1).
        SetNoteCols(TermRecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(Function(row) If(row = record.Count, Color.PaleGreen, Color.Transparent))

    Dim sumrow As RowRecord = CreateSumRowRecord(record, "", "合計")
    Dim nRecord As SheetRecord = record.AddLast(sumrow)
    CreateTable(TermRecordTableForm.tblRecord, tabPage, nRecord, tableDrawer)
    ShowTermTableRecord(tabPage)
    CurrentlyShowedSheetRecord = nRecord
  End Sub

  Private Sub CreateMonthlyTotalTable(tabPage As TabPage, record As SheetRecord, year As Integer, month As Integer)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(TotalRecordTableForm.pnlForTable).
      SetTextCols(0, 1).
      SetNoteCols(TermRecordTableForm.tblRecord.ColumnCount - 1).
      SetFuncBackColor(GetFuncForGettingRowColorInMonthlyTable(year, month))

    Dim sumrow As RowRecord = CreateSumRowRecord(record, "", "合計")
    Dim nRecord As SheetRecord = record.AddLast(sumrow)
    CreateTable(TotalRecordTableForm.tblRecord, tabPage, nRecord, tableDrawer)
    ShowTotalTableRecord(tabPage)
    CurrentlyShowedSheetRecord = nRecord
  End Sub

  Private Sub CreateTotalTable(tabPage As TabPage, record As SheetRecord)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(TotalRecordTableForm.pnlForTable).
        SetTextCols(0, 1).
        SetNoteCols(RecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(Function(row) If(row = TermRecordTableForm.tblRecord.RowCount - 1, Color.Green, Color.Transparent))

    Dim sumrow As RowRecord = CreateSumRowRecord(record, "", "合計")
    Dim nRecord As SheetRecord = record.AddLast(sumrow)
    CreateTable(TotalRecordTableForm.tblRecord, tabPage, record, tableDrawer)
    ShowTotalTableRecord(tabPage)
    CurrentlyShowedSheetRecord = nRecord
  End Sub

  Private Sub CreateTable(table As TableLayoutPanel, tabPage As TabPage, record As SheetRecord, drawer As TableDrawer)

    Dim loopRow As Action(Of RowRecord, Integer, Integer) =
      Sub(rec, insertCol, insertRow)
        If insertCol >= table.ColumnCount Then
          Return
        Else
          Dim value As String
          Dim nextRec As RowRecord

          If rec Is Nothing Then
            MessageBox.Show("レコードがもうありません。 col: " & insertCol.ToString & " row:" & insertRow)
          End If

          If rec.Empty Then
            value = ""
            nextRec = rec
          Else
            value = RoundDecStr(rec.First)
            nextRec = rec.Rest
          End If

          Dim control As Control = table.GetControlFromPosition(insertCol, insertRow)
          If control Is Nothing Then
            Dim panel As Panel = drawer.CreateCell(value, insertCol, insertRow)
            table.Controls.Add(panel, insertCol, insertRow)
          ElseIf control.Controls.Count = 0
            Dim panel As Panel = drawer.CreateCell(value, insertCol, insertRow)
            table.SetCellPosition(panel, table.GetCellPosition(control))
          Else
            control.BackColor = drawer.GetColor(insertRow)
            control.Controls.Item(0).Text = value
          End If

          loopRow(nextRec, insertCol + 1, insertRow)
        End If
      End Sub

    Dim loopSheet As Action(Of SheetRecord, Integer) =
      Sub(rec, insertRow)
        If insertRow >= table.RowCount Then
          Return
        Else
          Dim rr As RowRecord
          Dim nextSheet As SheetRecord
          If rec.Empty Then
            rr = RowRecord.Nil
            nextSheet = rec
          Else
            rr = rec.First
            nextSheet = rec.Rest
          End If
          loopRow(rr, 0, insertRow)
          loopSheet(nextSheet, insertRow + 1)
        End If

      End Sub

    loopSheet(record, 0)
  End Sub

  Private Function CreateSumRowRecord(record As SheetRecord, ParamArray headerCellValues As String()) As RowRecord
    Dim isFiltering As Boolean = chkExcludeIncompleteRecordFromSum.Checked
    Return _
      RecordConverter.CreateSumRowRecord(
        record,
        headerCellValues.Count,
        GetFuncForFilteringImcompleteRecord()
        ).AddRangeToHead(headerCellValues)
    'Dim go As Func(Of Integer, RowRecord) =
    '  Function(idx)
    '    If idx = 7 Then
    '      Return RowRecord.Nil
    '    Else
    '      Dim offset = idx * 3 + headerCellValues.Count
    '      Dim filter As Func(Of RowRecord, Boolean) = GetFuncForFilteringRecord(isFiltering, offset + 1)
    '      Dim sum1 As Double = RecordConverter.sum(record.Filtering(filter), offset)
    '      Dim sum2 As Double = RecordConverter.sum(record, offset + 1)
    '      Dim sum3 As Double = sum1 / sum2
    '      Return go(idx + 1).AddFirst(RoundDecStr(sum3)).AddFirst(RoundDecStr(sum2)).AddFirst(sum1)
    '    End If
    '  End Function

    'Return SheetRecord.Nil.AddFirst(go(0).AddRangeToHead(headerCellValues))
  End Function

  Private Function GetFuncForFilteringImcompleteRecord() As Func(Of RowRecord, Integer, Boolean)
    Dim isFiltering As Boolean = chkExcludeIncompleteRecordFromSum.Checked
    Return _
      Function(rr, colIdx)
        If Not isFiltering Then
          Return True
        ElseIf rr.Count <= colIdx + 1 Then
          Return False
        Else
          Return _
              rr.Count > colIdx + 1 AndAlso
              RecordConverter.ToDouble(rr.GetItem(colIdx)) > 0.0 AndAlso
              RecordConverter.ToDouble(rr.GetItem(colIdx + 1)) > 0.0
        End If
      End Function
  End Function

  Private Function GetFuncForFilteringRecord(isFiltering As Boolean, col As Integer) As Func(Of RowRecord, Boolean)
    Return If(isFiltering,
      Function(row)
        If col < row.Count Then
          Return RecordConverter.ToDouble(row.GetItem(col)) > 0.0
        Else
          Return False
        End If
      End Function,
      Function(row) True)
  End Function

  Delegate Sub LoadTableCallBack(tabPage As TabPage, sortInfo As SortInfo)

  Private Function GetActionForCreatingMonthlyTable(year As Integer, month As Integer) As LoadTableCallBack
    Return _
      Sub(tabPage, sortInfo)
        Dim r As SheetRecord = GetPersonalRecord(tabPage.Text)
        If r IsNot Nothing Then
          CreateDailyTableInMonth(tabPage, Sort(r, sortInfo), year, month)
        End If
      End Sub
  End Function

  Private Function GetActionForCreatingSumTable() As LoadTableCallBack
    Return _
      Sub(tabPage, sortInfo)
        Dim r As SheetRecord = GetPersonalRecord(tabPage.Text)
        If r IsNot Nothing Then
          CreateSumTable(tabPage, Sort(r, sortInfo))
        End If
      End Sub
  End Function

  Private Function GetPersonalRecord(sheetName As String) As SheetRecord
    Dim selectedUser As ExpandedUserInfo = cboxUserInfo.SelectedItem

    If selectedUser IsNot Nothing Then
      If Not UserRecordManager.ContainsRecord(selectedUser.GetIdNum) Then
        UserRecordLoader.Load(selectedUser)
      End If

      Return UserRecordManager.GetSheetRecord(selectedUser.GetIdNum, sheetName, GetFuncForFilteringImcompleteRecord())
    Else
      Return Nothing
    End If
  End Function

  Private Function GetActionForCreatingDailyTermTable() As LoadTableCallBack
    Return _
      Sub(tabPage, sortInfo)
        Dim y As Integer = dPicDailyTerm.Value.Year
        Dim m As Integer = dPicDailyTerm.Value.Month
        Dim d As Integer = dPicDailyTerm.Value.Day
        Dim r As SheetRecord = UserRecordManager.GetDailyTermRecord(d, m, y)
        CreateTermTable(tabPage, Sort(r, sortInfo))
      End Sub
  End Function

  Private Function GetActionForCreatingWeeklyTermTable() As LoadTableCallBack
    Return _
      Sub(tabPage, sortInfo)
        Dim item As WeeklyItem = cboxWeeklyTerm.SelectedItem
        If item IsNot Nothing Then
          Dim w As Integer = item.Week
          Dim m As Integer = item.Month
          Dim y As Integer = item.Year
          Dim r As SheetRecord = UserRecordManager.GetWeeklyTermRecord(w, m, y, GetFuncForFilteringImcompleteRecord())
          CreateTermTable(tabPage, Sort(r, sortInfo))
        End If
      End Sub
  End Function

  Private Function GetActionForCreatingMonthlyTermTable() As LoadTableCallBack
    Return _
      Sub(tabPage, sortInfo)
        Dim item As MonthlyItem = cboxMonthlyTerm.SelectedItem
        If item IsNot Nothing Then
          Dim m As Integer = item.Month
          Dim y As Integer = item.Year
          Dim r As SheetRecord = UserRecordManager.GetMonthlyTermRecord(m, y, GetFuncForFilteringImcompleteRecord())
          CreateTermTable(tabPage, Sort(r, sortInfo))
        End If
      End Sub
  End Function

  Private Function GetActionForCreatingAllTermTable() As LoadTableCallBack
    Return _
      Sub(tabPage, sortInfo)
        Dim r As SheetRecord = UserRecordManager.GetAllTermRecord()
        CreateTermTable(tabPage, Sort(r, sortInfo))
      End Sub
  End Function

  Private Function GetActionForCreatingDailyTotalTable() As LoadTableCallBack
    Return _
      Sub(tabPage, sortInfo)
        Dim item As MonthlyItem = cboxDailyTotal.SelectedItem
        If item IsNot Nothing Then
          Dim m As Integer = item.Month
          Dim y As Integer = item.Year
          Dim r As SheetRecord = UserRecordManager.GetDailyTotalRecord(m, y, GetFuncForFilteringImcompleteRecord())
          CreateMonthlyTotalTable(tabPage, Sort(r, sortInfo), y, m)
        End If
      End Sub
  End Function

  Private Function GetActionForCreatingPlaneTable(f As Func(Of SheetRecord)) As LoadTableCallBack
    Return _
      Sub(tabPage, sortInfo)
        'Dim y As Integer = FileFormat.GetYear()
        'Dim r As SheetRecord = UserRecordManager.GetWeeklyTotalRecord(y, Function(e) True)
        CreateTotalTable(tabPage, Sort(f(), sortInfo))
      End Sub
  End Function

  Private Function Sort(record As SheetRecord, sortInfo As SortInfo) As SheetRecord
    Dim f As Func(Of RowRecord, Double) = Function(row) RecordConverter.ToDouble(row.GetItem(sortInfo.Col))
    If sortInfo.Col = -1 Then
      Return record
    ElseIf sortInfo.Asc Then
      Return record.OrderBy(Function(row) f(row))
    Else
      Return record.OrderByDescending(Function(row) f(row))
    End If
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

  Private Function GetRowColorInSumTable(row As Integer) As Color
    If row = 21 Then
      Return Color.PaleTurquoise
      'Return Color.Transparent
    ElseIf row Mod 7 = 6 AndAlso row < 21
      Return Color.PaleGreen
    Else
      Return Color.Transparent
    End If
  End Function

  Private Function RoundDecStr(value As String) As String
    If MP.Utils.General.MyChar.IsDouble(value) Then
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

  Private Sub ShowTermTableTitles(showedTabPage As TabPage)
    Dim tblTitles As TableLayoutPanel = TermRecordTableForm.tblTitles
    tblTitles.Location = New Point(3, 46)
    Dim tblSubTitles As TableLayoutPanel = TermRecordTableForm.tblSubTItles
    tblSubTitles.Location = New Point(3, 72)

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

    'Dim tblSumOfRecord As TableLayoutPanel = RecordTableForm.tblSumOfRecord
    'tblSumOfRecord.Location = New Point(3, 558)
    'showedTabPage.Controls.Add(tblSumOfRecord)
  End Sub

  Private Sub ShowTermTableRecord(showedTabPage As TabPage)
    Dim scrollPanel As Panel = TermRecordTableForm.pnlForTable
    scrollPanel.Location = New Point(3, 103)
    scrollPanel.AutoScroll = True
    showedTabPage.Controls.Add(scrollPanel)

    'Dim tblSumOfRecord As TableLayoutPanel = TermRecordTableForm.tblSumOfRecord
    'tblSumOfRecord.Location = New Point(3, 558)
    'showedTabPage.Controls.Add(tblSumOfRecord)
  End Sub

  Private Sub ShowTotalTableRecord(showedTabPage As TabPage)
    Dim scrollPanel As Panel = TotalRecordTableForm.pnlForTable
    scrollPanel.Location = New Point(3, 103)
    scrollPanel.AutoScroll = True
    showedTabPage.Controls.Add(scrollPanel)

    'Dim tblSumOfRecord As TableLayoutPanel = TotalRecordTableForm.tblSumOfRecord
    'tblSumOfRecord.Location = New Point(3, 558)
    'showedTabPage.Controls.Add(tblSumOfRecord)
  End Sub

  Private Function GetWeek(year As Integer, month As Integer, day As Integer) As Integer
    If month >= 1 AndAlso month <= 12 AndAlso day <= Date.DaysInMonth(year, month) Then
      Return Weekday(year & "/" & month & "/" & day)
    Else
      Return -1
    End If
  End Function

  Private Sub cmdOutputCSV_Click(sender As Object, e As EventArgs) Handles cmdOutputCSV.Click
    If CurrentlyShowedSheetRecord Is Nothing Then
      MessageBox.Show("データが表示されていません。")
    Else
      'SaveFileDialogクラスのインスタンスを作成
      Dim sfd As New SaveFileDialog()

      'はじめのファイル名を指定する
      sfd.FileName = "新しいファイル.txt"
      'はじめに表示されるフォルダを指定する
      sfd.InitialDirectory = MP.Details.Sys.App.GetCurrentDirectory
      '[ファイルの種類]に表示される選択肢を指定する
      sfd.Filter = "textファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*"
      sfd.FilterIndex = 0
      'タイトルを設定する
      sfd.Title = "保存先のファイルを選択してください"
      'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
      sfd.RestoreDirectory = True
      '既に存在するファイル名を指定したとき警告する
      'デフォルトでTrueなので指定する必要はない
      sfd.OverwritePrompt = True
      '存在しないパスが指定されたとき警告を表示する
      'デフォルトでTrueなので指定する必要はない
      sfd.CheckPathExists = True

      'ダイアログを表示する
      If sfd.ShowDialog() = DialogResult.OK Then
        Dim stream As System.IO.Stream = Nothing
        Dim sw As System.IO.StreamWriter = Nothing
        Try
          stream = sfd.OpenFile()
          If Not (stream Is Nothing) Then
            'ファイルに書き込む
            sw = New System.IO.StreamWriter(stream)
            sw.Write(RecordConverter.ToCSV(CurrentlyShowedSheetRecord))
          End If
        Catch ex As Exception
          MsgBox.ShowError(ex)
        Finally
          If sw IsNot Nothing Then
            sw.Close()
          End If
          If stream IsNot Nothing Then
            stream.Close()
          End If
        End Try
      End If
    End If
  End Sub

  Public Class WeeklyItem
    Inherits DateItem

    Private _Week As Integer
    Public ReadOnly Property Week() As Integer
      Get
        Return _Week
      End Get
    End Property

    Public Shared Function GetWeekNumInMonth(d As Integer, m As Integer, y As Integer) As Integer
      Return GetWeekNumInMonth(New Date(y, m, d))
    End Function

    Public Shared Function GetWeekNumInMonth(day As Date) As Integer
      Dim first As Date = MonthlyItem.GetFirstDateInMonth(day.Month, day.Year)
      Return DatePart("WW", day) - DatePart("ww", first) + 1
    End Function

    Public Sub New(w As Integer, m As Integer, y As Integer)
      MyBase.New(w, m, y)
      _Week = MyBase.Value
    End Sub

    Public Overrides Function Agree(day As Date) As Boolean
      If day.Year = Year() AndAlso day.Month = Month() Then
        Dim w As Integer = GetWeekNumInMonth(day)
        Return w = Week()
      Else
        Return False
      End If
    End Function

    Public Overrides Function ToString() As String
      Return Month & "月 第" & _Week & "週"
    End Function
  End Class

  Public Class MonthlyItem
    Inherits DateItem

    Public Shared Function GetFirstDateInMonth(m As Integer, y As Integer) As Date
      Return New Date(y, m, 1)
    End Function

    Public Sub New(m As Integer, y As Integer)
      MyBase.New(-1, m, y)
    End Sub

    Public Overrides Function Agree(day As Date) As Boolean
      Return day.Year = Year() AndAlso day.Month = Month()
    End Function

    Public Overrides Function ToString() As String
      Return Month & "月"
    End Function
  End Class

  Public MustInherit Class DateItem
    Private _Value As Integer
    Protected ReadOnly Property Value() As Integer
      Get
        Return _Value
      End Get
    End Property

    Private _Month As Integer
    Public ReadOnly Property Month() As Integer
      Get
        Return _Month
      End Get
    End Property

    Private _Year As Integer
    Public ReadOnly Property Year() As Integer
      Get
        Return _Year
      End Get
    End Property

    Public Sub New(value As Integer, m As Integer, y As Integer)
      _Value = value
      _Month = m
      _Year = y
    End Sub

    Public MustOverride Function Agree(day As Date) As Boolean
  End Class


End Class

