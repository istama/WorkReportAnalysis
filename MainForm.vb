
Imports MP.Office
Imports AP = MP.Utils.Common.AppProperties
Imports MP.Utils.Common
Imports MP.Utils.Model
Imports MP.Utils.MyDate
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
    Dim SheetRecordController As SheetRecordController
  End Structure

  Public Structure SortProperties
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

  Private SaveFileDialog As SaveFileDialog
  Private CurrentlyShowedSheetRecordManager As SheetRecordManager = Nothing

  Private Loaded As Boolean = False

  Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    'SettingLog()
    'MyLog.Write("Main Formを起動しました。")
    'Try
    '  AutoUpdate()
    'Catch ex As Exception
    'End Try

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

      InitSvaeFileDialog()

      LoadAllUserRecord()
    Catch ex As Exception
      MsgBox.ShowError(ex)
    End Try

    Loaded = True
  End Sub

  'Private Sub SettingLog()
  '  MyLog.Log.DefaultFileLogWriter.Location = Logging.LogFileLocation.ExecutableDirectory
  '  MyLog.Log.DefaultFileLogWriter.Append = False
  '  If AppProps.GetValue(AP.KEY_WRITE_LOG) = "True" Then
  '    MyLog.LogMode = True
  '  Else
  '    MyLog.LogMode = False
  '  End If
  'End Sub

  'Private Sub AutoUpdate()
  '  If AppProps.GetValue(AP.KEY_ENABLE_AUTO_UPDATE) = "True" Then
  '    MyLog.Write("自動アップデートを開始します。")
  '    Dim updateManager As UpdateManager = New UpdateManager(FilePath.UpdateScriptPath(), FilePath.ReleaseVersionInfoFilePath())

  '    updateManager.GenerateDefaultUpdateBatchIfEmpty(AppProps.GetValue(AP.KEY_RELEASE_DIR_FOR_UPDATE), FilePath.ExcludeFileForUpdatePath())
  '    If updateManager.hasUpdated() Then
  '      MessageBox.Show("最新のバージョンに更新します。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information)
  '      Me.Close()
  '      System.Diagnostics.Process.Start(FilePath.UpdateScriptPath())
  '    End If
  '  Else
  '    My.Application.Log.WriteEntry("自動アップデートはオフです。")
  '  End If
  'End Sub

  Private Sub LoadAllUsers()
    Dim a As SerializedAccessor = MySerialize.GenerateAccessor()
    Dim path As String = FilePath.UserinfoFilePath()

    UserInfoList = a.GetInfo(Of UserInfo)(path).ConvertAll(Function(user) New ExpandedUserInfo(user))
  End Sub

  Private Sub InitInnerTabPageInPersonalTab()
    Dim infoList = New List(Of TabInfo)()

    For Each ym As MonthlyItem In ReadRecordTerm.MonthlyItemList.ToList
      Dim tab As TabInfo
      With tab
        .Name = UserRecordManager.GetSheetName(ym.Month)
        Dim c As SheetRecordController
        With c
          .TabPage = Nothing
          .ReadCallback = AddressOf GetPersonalRecord
          .AddSumOfRecordCallback = AddressOf AddSumOfPersonalRecord
          .LoadTableCallback = GetActionForCreatingPersonalDailyTableInMonth(ym.Year, ym.Month)
          .CanSort = True
          .GetCSVFileNameCallback = GetFuncForGettingCSVFileNameOfPersonal(ym)
          .GetCSVTitlesCallback = AddressOf GetCSVTitlesOfPersonalRecord
          .GetCSVSubTitlesCallback = AddressOf GetCSVSubTitlesOfPersonalRecord
        End With
        .SheetRecordController = c
      End With
      infoList.Add(tab)
    Next

    Dim sumTab As TabInfo
    With sumTab
      .Name = UserRecordManager.GetSumSheetName
      Dim c As SheetRecordController
      With c
        .TabPage = Nothing
        .ReadCallback = AddressOf GetPersonalRecord
        .AddSumOfRecordCallback = Function(roc) roc
        .LoadTableCallback = AddressOf CreateSumTable
        .CanSort = False
        .GetCSVFileNameCallback = AddressOf GetCSVFileNameOfSumOfPersonal
        .GetCSVTitlesCallback = AddressOf GetCSVTitlesOfTotalRecord
        .GetCSVSubTitlesCallback = AddressOf GetCSVSubTitlesOfTotalRecord
      End With
      .SheetRecordController = c
    End With
    infoList.Add(sumTab)

    InnerTabPageInfoListInPersonalTab = connectTabPageWithInfo(tabInPersonalTab, infoList)
  End Sub

  Private Sub InitInnerTabPageInTermTab()
    Dim infoList = New List(Of TabInfo)()
    Dim tabDays As TabInfo
    With tabDays
      .Name = "日"
      Dim c As SheetRecordController
      With c
        .TabPage = Nothing
        .ReadCallback = AddressOf GetDailyTermRecord
        .AddSumOfRecordCallback = AddressOf AddSumOfTermRecord
        .LoadTableCallback = AddressOf CreateTermTable
        .CanSort = True
        .GetCSVFileNameCallback = AddressOf GetCSVFileNameOfDailyTerm
        .GetCSVTitlesCallback = AddressOf GetCSVTitlesOfTotalRecord
        .GetCSVSubTitlesCallback = AddressOf GetCSVSubTitlesOfTermRecord
      End With
      .SheetRecordController = c
    End With

    Dim tabWeeks As TabInfo
    With tabWeeks
      .Name = "週"
      Dim c As SheetRecordController
      With c
        .TabPage = Nothing
        .ReadCallback = AddressOf GetWeeklyTermRecord
        .AddSumOfRecordCallback = AddressOf AddSumOfTermRecord
        .LoadTableCallback = AddressOf CreateTermTable
        .CanSort = True
        .GetCSVFileNameCallback = AddressOf GetCSVFileNameOfWeeklyTerm
        .GetCSVTitlesCallback = AddressOf GetCSVTitlesOfTotalRecord
        .GetCSVSubTitlesCallback = AddressOf GetCSVSubTitlesOfTermRecord
      End With
      .SheetRecordController = c
    End With

    Dim tabMonths As TabInfo
    With tabMonths
      .Name = "月"
      Dim c As SheetRecordController
      With c
        .TabPage = Nothing
        .ReadCallback = AddressOf GetMonthlyTermRecord
        .AddSumOfRecordCallback = AddressOf AddSumOfTermRecord
        .LoadTableCallback = AddressOf CreateTermTable
        .CanSort = True
        .GetCSVFileNameCallback = AddressOf GetCSVFileNameOfMonthlyTerm
        .GetCSVTitlesCallback = AddressOf GetCSVTitlesOfTotalRecord
        .GetCSVSubTitlesCallback = AddressOf GetCSVSubTitlesOfTermRecord
      End With
      .SheetRecordController = c
    End With

    Dim tabYear As TabInfo
    With tabYear
      .Name = "合計"
      Dim c As SheetRecordController
      With c
        .TabPage = Nothing
        .ReadCallback = AddressOf GetAllTermRecord
        .AddSumOfRecordCallback = AddressOf AddSumOfTermRecord
        .LoadTableCallback = AddressOf CreateTermTable
        .CanSort = True
        .GetCSVFileNameCallback = Function() String.Format("日付データ_{0}年.csv", FileFormat.GetYear())
        .GetCSVTitlesCallback = AddressOf GetCSVTitlesOfTotalRecord
        .GetCSVSubTitlesCallback = AddressOf GetCSVSubTitlesOfTermRecord
      End With
      .SheetRecordController = c
    End With

    infoList.Add(tabDays)
    infoList.Add(tabWeeks)
    infoList.Add(tabMonths)
    infoList.Add(tabYear)

    InnerTabPageInfoListInTermTab = connectTabPageWithInfo(tabInTermTab, infoList)
  End Sub

  Private Sub InitInnerTabPageInTotalTab()
    Dim infoList = New List(Of TabInfo)()
    Dim tabDays As TabInfo
    With tabDays
      .Name = "日"
      Dim c As SheetRecordController
      With c
        .TabPage = Nothing
        .ReadCallback = AddressOf GetDailyTotalRecord
        .AddSumOfRecordCallback = AddressOf AddSumOfTermRecord
        .LoadTableCallback = GetActionForCreatingDailyTotalTableInMonth()
        .CanSort = True
        .GetCSVFileNameCallback = AddressOf GetCSVFileNameOfDailyTotal
        .GetCSVTitlesCallback = AddressOf GetCSVTitlesOfTotalRecord
        .GetCSVSubTitlesCallback = AddressOf GetCSVSubTitlesOfTotalRecord
      End With
      .SheetRecordController = c
    End With

    Dim tabWeeks As TabInfo
    With tabWeeks
      .Name = "週"
      Dim c As SheetRecordController
      With c
        .TabPage = Nothing
        .ReadCallback = AddressOf GetWeeklyTotalRecord
        .AddSumOfRecordCallback = AddressOf AddSumOfTermRecord
        .LoadTableCallback = AddressOf CreateTotalTable
        .CanSort = True
        .GetCSVFileNameCallback = Function() String.Format("集計データ_週単位_{0}年", FileFormat.GetYear())
        .GetCSVTitlesCallback = AddressOf GetCSVTitlesOfTotalRecord
        .GetCSVSubTitlesCallback = AddressOf GetCSVSubTitlesOfTotalRecord
      End With
      .SheetRecordController = c
    End With

    Dim tabMonths As TabInfo
    With tabMonths
      .Name = "月"
      Dim c As SheetRecordController
      With c
        .TabPage = Nothing
        .ReadCallback = AddressOf GetMonthlyTotalRecord
        .AddSumOfRecordCallback = AddressOf AddSumOfTermRecord
        .LoadTableCallback = AddressOf CreateTotalTable
        .CanSort = True
        .GetCSVFileNameCallback = Function() String.Format("集計データ_月単位_{0}年", FileFormat.GetYear())
        .GetCSVTitlesCallback = AddressOf GetCSVTitlesOfTotalRecord
        .GetCSVSubTitlesCallback = AddressOf GetCSVSubTitlesOfTotalRecord
      End With
      .SheetRecordController = c
    End With

    infoList.Add(tabDays)
    infoList.Add(tabWeeks)
    infoList.Add(tabMonths)

    InnerTabPageInfoListInTotalTab = connectTabPageWithInfo(tabInTotalTab, infoList)
  End Sub

  Private Function connectTabPageWithInfo(tab As TabControl, tabPageInfoList As List(Of TabInfo)) As List(Of TabInfo)
    Dim l As New List(Of TabInfo)

    If tab.TabPages.Count = tabPageInfoList.Count Then
      For idx As Integer = 0 To tabPageInfoList.Count - 1
        Dim info As TabInfo = tabPageInfoList(idx)
        tab.TabPages.Item(idx).Text = info.Name
        info.SheetRecordController.TabPage = tab.TabPages.Item(idx)
        l.Add(info)
      Next
      Return l
    Else
      Throw New Exception(
        "Excelファイルのシート数とタブページの数が合いません。 / tabPageCount: " &
        tab.TabPages.Count.ToString &
        " tabInfoCount: " & tabPageInfoList.Count)
    End If
  End Function

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

    For i As Integer = 0 To RecordTableForm.tblSubTitles.ColumnCount - 2
      AddHandler RecordTableForm.tblSubTitles.GetControlFromPosition(i, 0).Controls.Item(0).Click, AddressOf SortPersonalRecordCallback
    Next
    For i As Integer = 0 To TermRecordTableForm.tblSubTItles.ColumnCount - 2
      AddHandler TermRecordTableForm.tblSubTItles.GetControlFromPosition(i, 0).Controls.Item(0).Click, AddressOf SortTermRecordCallback
    Next
    For i As Integer = 0 To TotalRecordTableForm.tblSubTItles.ColumnCount - 2
      AddHandler TotalRecordTableForm.tblSubTItles.GetControlFromPosition(i, 0).Controls.Item(0).Click, AddressOf SortTotalRecordCallback
    Next

    ShowPersonalTableTitles(page10Month, GetOffsetLocationInPersonalTab)
    ShowTotalTableTitles(pageSum, GetOffsetLocationInPersonalTab)
    ShowTermTableTitles(pageDays, GetOffsetLocationInTermTab)
    ShowTotalTableTitles(pageDailyTotal, GetOffsetLocationInTotalTab)
  End Sub

  Private Sub SortPersonalRecordCallback(sender As Object, e As EventArgs)
    SortRecordCallback(sender, RecordTableForm.tblSubTitles, tabInPersonalTab.SelectedTab, InnerTabPageInfoListInPersonalTab)
  End Sub

  Private Sub SortTermRecordCallback(sender As Object, e As EventArgs)
    SortRecordCallback(sender, TermRecordTableForm.tblSubTItles, tabInTermTab.SelectedTab, InnerTabPageInfoListInTermTab)
  End Sub

  Private Sub SortTotalRecordCallback(sender As Object, e As EventArgs)
    SortRecordCallback(sender, TotalRecordTableForm.tblSubTItles, tabInTotalTab.SelectedTab, InnerTabPageInfoListInTotalTab)
  End Sub

  Private Sub SortRecordCallback(sender As Object, table As TableLayoutPanel, selectedTab As TabPage, tabInfo As List(Of TabInfo))
    Dim label As Label = CType(sender, Label)
    Dim col As Integer = table.GetColumn(label.Parent)

    If CurrentlyShowedSheetRecordManager IsNot Nothing AndAlso CurrentlyShowedSheetRecordManager.AllowSort Then
      Dim manager As SheetRecordManager = CurrentlyShowedSheetRecordManager.Sort(col)
      LoadTableInnerTabPage(manager)
    End If
  End Sub

  Private Sub InitCBoxUserInfo()
    UserInfoList.ForEach(Function(user) cboxUserInfo.Items.Add(user))
  End Sub

  Private Sub InitCBoxWeeklyTerm()
    Dim items As New List(Of DateItem)
    ReadRecordTerm.WeeklyItemList.ForEach(Sub(w) items.Add(w))

    items.ForEach(Sub(i) cboxWeeklyTerm.Items.Add(i))

    'Dim current As WeeklyItem = items.Find(Function(i) i.Agree(Date.Today))
    'If current IsNot Nothing Then
    '  Dim idx As Integer = cboxWeeklyTerm.Items.IndexOf(current)
    '  cboxWeeklyTerm.SelectedIndex = idx
    'End If
  End Sub

  Private Sub InitCBoxMonthlyTerm()
    Dim items As New List(Of MonthlyItem)
    ReadRecordTerm.MonthlyItemList.ForEach(Sub(m) items.Add(m))

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
    ReadRecordTerm.MonthlyItemList.ForEach(Sub(m) items.Add(m))

    items.ForEach(Sub(i) cboxDailyTotal.Items.Add(i))
  End Sub

  Private Sub InitDPicDailyTerm()
    Dim min As MonthlyItem = ReadRecordTerm.MonthlyItemList.First
    dPicDailyTerm.MinDate = New DateTime(min.Year, min.Month, 1, 0, 0, 0)
    Dim max As MonthlyItem = ReadRecordTerm.MonthlyItemList.ToList.Last
    dPicDailyTerm.MaxDate = New DateTime(max.Year, max.Month, Date.DaysInMonth(max.Year, max.Month), 0, 0, 0)
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
      ShowPersonalRecord(tabInPersonalTab.SelectedTab)
    ElseIf idx = 1
      ShowTermRecord()
    Else
      ShowTotalRecord()
    End If
  End Sub

  Private Sub tabInPersonalTab_PageChanged(sender As Object, e As EventArgs) Handles tabInPersonalTab.SelectedIndexChanged
    ShowPersonalRecord(tabInPersonalTab.SelectedTab)
  End Sub

  Private Sub cboxUserInfo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxUserInfo.SelectedIndexChanged
    ReadSheetRecord(tabInPersonalTab.SelectedTab, InnerTabPageInfoListInPersonalTab)
  End Sub

  Private Sub tabInTermTab_PageChanged(sender As Object, e As EventArgs) Handles tabInTermTab.SelectedIndexChanged
    ShowTermRecord()
  End Sub

  Private Sub tabInTotalTab_PageChanged(sender As Object, e As EventArgs) Handles tabInTotalTab.SelectedIndexChanged
    ShowTotalRecord()
  End Sub

  Private Sub datePickerDailyTerm_DateChanged(sender As Object, e As EventArgs) Handles dPicDailyTerm.ValueChanged
    If ReadRecordTerm.InMonthlyTerm(dPicDailyTerm.Value.Month, dPicDailyTerm.Value.Year) Then
      ShowTermRecord()
    End If
  End Sub

  Private Sub cboxWeeklyTerm_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxWeeklyTerm.SelectedIndexChanged
    If cboxWeeklyTerm.SelectedItem IsNot Nothing Then
      ShowTermRecord()
    Else
      CurrentlyShowedSheetRecordManager = Nothing
    End If
  End Sub

  Private Sub cboxMonthlyTerm_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxMonthlyTerm.SelectedIndexChanged
    If cboxMonthlyTerm.SelectedItem IsNot Nothing Then
      ShowTermRecord()
    Else
      CurrentlyShowedSheetRecordManager = Nothing
    End If
  End Sub

  Private Sub cboxDailyTotal_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxDailyTotal.SelectedIndexChanged
    If cboxDailyTotal.SelectedItem IsNot Nothing Then
      ShowTotalRecord()
    Else
      CurrentlyShowedSheetRecordManager = Nothing
    End If
  End Sub

  Private Sub chkExcludeIncompleteRecordFromSum_CheckedChanged(sender As Object, e As EventArgs) Handles chkExcludeIncompleteRecordFromSum.CheckedChanged
    If Loaded Then
      Call tabMaster_PageChanged(sender, e)
    End If
  End Sub

  Public Sub ShowPersonalRecord(selectedTab As TabPage)
    If tabInPersonalTab.SelectedIndex = tabInPersonalTab.TabCount - 1 Then
      'ShowTotalTableTitles(tabInPersonalTab.SelectedTab, GetOffsetLocationInPersonalTab)
      ShowTotalTableTitles(selectedTab, GetOffsetLocationInPersonalTab)
    Else
      'ShowPersonalTableTitles(tabInPersonalTab.SelectedTab, GetOffsetLocationInPersonalTab)
      ShowPersonalTableTitles(selectedTab, GetOffsetLocationInPersonalTab)
    End If

    'ReadSheetRecord(tabInPersonalTab.SelectedTab, InnerTabPageInfoListInPersonalTab)
    ReadSheetRecord(selectedTab, InnerTabPageInfoListInPersonalTab)
  End Sub

  Private Sub ShowTermRecord()
    ShowTermTableTitles(tabInTermTab.SelectedTab, GetOffsetLocationInTermTab)
    ReadSheetRecord(tabInTermTab.SelectedTab, InnerTabPageInfoListInTermTab)
  End Sub

  Private Sub ShowTotalRecord()
    ShowTotalTableTitles(tabInTotalTab.SelectedTab, GetOffsetLocationInTotalTab)
    ReadSheetRecord(tabInTotalTab.SelectedTab, InnerTabPageInfoListInTotalTab)
  End Sub

  Private Sub ReadSheetRecord(selectedTabPage As TabPage, tabInfoList As List(Of TabInfo))
    Try
      Dim controller As SheetRecordController =
        tabInfoList.
          Find(Function(info) info.Name = selectedTabPage.Text).
          SheetRecordController

      Dim recManager As SheetRecordManager = New SheetRecordManager(controller)
      LoadTableInnerTabPage(recManager.ReadSheetRecord)
    Catch ex As Exception
      MsgBox.ShowError(ex)
    End Try
  End Sub

  Private Sub LoadTableInnerTabPage(recordManager As SheetRecordManager)
    CurrentlyShowedSheetRecordManager = recordManager.LoadTable(Nothing)
  End Sub

  Private Sub CreateDailyTableInMonth(tabPage As TabPage, record As SheetRecord, year As Integer, month As Integer)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(RecordTableForm.pnlForTable).
        SetTextCols(0).
        SetNoteCols(RecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(GetFuncForGettingRowColorInMonthlyTable(year, month))

    CreateTable(RecordTableForm.tblRecord, record, tableDrawer, GetFuncForEmptyHandler)
    ShowPersonalTableRecord(tabPage, GetOffsetLocationInPersonalTab)
  End Sub

  Private Sub CreateSumTable(tabPage As TabPage, record As SheetRecord)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(TotalRecordTableForm.pnlForTable).
        SetTextCols(0).
        SetNoteCols(TotalRecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(Function(row) If(row = record.Count - 1, Color.PaleGreen, Color.Transparent))

    CreateTable(TotalRecordTableForm.tblRecord, record, tableDrawer, GetFuncForEmptyHandler())
    ShowTotalTableRecord(tabPage, GetOffsetLocationInPersonalTab)
  End Sub

  Private Sub CreateTermTable(tabPage As TabPage, record As SheetRecord)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(TermRecordTableForm.pnlForTable).
        SetTextCols(0, 1).
        SetNoteCols(TermRecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(Function(row) If(row = record.Count - 1, Color.PaleGreen, Color.Transparent))

    CreateTable(TermRecordTableForm.tblRecord, record, tableDrawer, GetFuncForCreatingEventHandlerInTermRecord)
    ShowTermTableRecord(tabPage, GetOffsetLocationInTermTab)
  End Sub

  Private Sub CreateDailyTotalTableInMonth(tabPage As TabPage, record As SheetRecord, year As Integer, month As Integer)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(TotalRecordTableForm.pnlForTable).
      SetTextCols(0, 1).
      SetNoteCols(TermRecordTableForm.tblRecord.ColumnCount - 1).
      SetFuncBackColor(GetFuncForGettingRowColorInMonthlyTable(year, month))

    CreateTable(TotalRecordTableForm.tblRecord, record, tableDrawer, GetFuncForEmptyHandler())
    ShowTotalTableRecord(tabPage, GetOffsetLocationInTotalTab)
  End Sub

  Private Sub CreateTotalTable(tabPage As TabPage, record As SheetRecord)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(TotalRecordTableForm.pnlForTable).
        SetTextCols(0, 1).
        SetNoteCols(RecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(Function(row) If(row = record.Count - 1, Color.PaleGreen, Color.Transparent))

    CreateTable(TotalRecordTableForm.tblRecord, record, tableDrawer, GetFuncForEmptyHandler())
    ShowTotalTableRecord(tabPage, GetOffsetLocationInTotalTab)
  End Sub

  Private Sub CreateTable(table As TableLayoutPanel, record As SheetRecord, drawer As TableDrawer, callback As GetHandlerCallback)

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
            'TODO セルをダブルクリックしたときのイベントハンドラを設定する
            Dim panel As Panel = drawer.CreateCell(value, insertCol, insertRow, callback(insertCol, insertRow, rec.First))
            table.Controls.Add(panel, insertCol, insertRow)
          ElseIf control.Controls.Count = 0
            Dim panel As Panel = drawer.CreateCell(value, insertCol, insertRow, callback(insertCol, insertRow, rec.First))
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

  Private Sub Click_IdLabel(sender As Object, e As EventArgs)
    Dim label As Label = CType(sender, Label)

    MessageBox.Show(label.Text)
  End Sub

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

  Delegate Function ReadRecordCallback() As SheetRecord

  Private Function GetPersonalRecord() As SheetRecord
    Dim page As TabPage = tabInPersonalTab.SelectedTab
    Dim selectedUser As ExpandedUserInfo = cboxUserInfo.SelectedItem

    If page IsNot Nothing AndAlso selectedUser IsNot Nothing Then
      If Not UserRecordManager.ContainsRecord(selectedUser.GetIdNum) Then
        UserRecordLoader.Load(selectedUser)
      End If

      Return UserRecordManager.GetSheetRecord(selectedUser.GetIdNum, page.Text, GetFuncForFilteringImcompleteRecord())
    Else
      Return Nothing
    End If
  End Function

  Private Function GetDailyTermRecord() As SheetRecord
    Dim y As Integer = dPicDailyTerm.Value.Year
    Dim m As Integer = dPicDailyTerm.Value.Month
    Dim d As Integer = dPicDailyTerm.Value.Day
    Return UserRecordManager.GetDailyTermRecord(d, m, y)
  End Function

  Private Function GetWeeklyTermRecord() As SheetRecord
    Dim item As WeeklyItem = cboxWeeklyTerm.SelectedItem
    If item IsNot Nothing Then
      Dim w As Integer = item.Week
      Dim m As Integer = item.Month
      Dim y As Integer = item.Year
      Return UserRecordManager.GetWeeklyTermRecord(w, m, y, GetFuncForFilteringImcompleteRecord())
    Else
      Return Nothing
    End If
  End Function

  Private Function GetMonthlyTermRecord() As SheetRecord
    Dim item As MonthlyItem = cboxMonthlyTerm.SelectedItem
    If item IsNot Nothing Then
      Dim m As Integer = item.Month
      Dim y As Integer = item.Year
      Return UserRecordManager.GetMonthlyTermRecord(m, y, GetFuncForFilteringImcompleteRecord())
    Else
      Return Nothing
    End If
  End Function

  Private Function GetAllTermRecord() As SheetRecord
    Return UserRecordManager.GetAllTermRecord(FileFormat.GetYear(), GetFuncForFilteringImcompleteRecord())
  End Function

  Private Function GetDailyTotalRecord() As SheetRecord
    Dim item As MonthlyItem = cboxDailyTotal.SelectedItem
    If item IsNot Nothing Then
      Dim m As Integer = item.Month
      Dim y As Integer = item.Year
      Return UserRecordManager.GetDailyTotalRecord(m, y, GetFuncForFilteringImcompleteRecord())
    Else
      Return Nothing
    End If
  End Function

  Private Function GetWeeklyTotalRecord() As SheetRecord
    Return UserRecordManager.GetWeeklyTotalRecord(FileFormat.GetYear, GetFuncForFilteringImcompleteRecord())
  End Function

  Private Function GetMonthlyTotalRecord() As SheetRecord
    Return UserRecordManager.GetMonthlyTotalRecord(FileFormat.GetYear, GetFuncForFilteringImcompleteRecord())
  End Function

  Delegate Function AddSumOfRowRecordCallback(record As SheetRecord) As SheetRecord

  Private Function AddSumOfPersonalRecord(record As SheetRecord) As SheetRecord
    Dim sumrow As RowRecord = CreateSumRowRecord(record, "合計")
    Return record.AddLast(sumrow)
  End Function

  Private Function AddSumOfTermRecord(record As SheetRecord) As SheetRecord
    Dim sumrow As RowRecord = CreateSumRowRecord(record, "", "合計")
    Return record.AddLast(sumrow)
  End Function

  Private Function CreateSumRowRecord(record As SheetRecord, ParamArray headerCellValues As String()) As RowRecord
    Dim isFiltering As Boolean = chkExcludeIncompleteRecordFromSum.Checked
    Return _
      RecordConverter.CreateSumRowRecord(
        record,
        headerCellValues.Count,
        GetFuncForFilteringImcompleteRecord()
        ).AddRangeToHead(headerCellValues)
  End Function

  Delegate Sub LoadTableCallBack(tabPage As TabPage, record As SheetRecord)

  Private Function GetActionForCreatingPersonalDailyTableInMonth(year As Integer, month As Integer) As LoadTableCallBack
    Return _
      Sub(tabPage, record)
        CreateDailyTableInMonth(tabPage, record, year, month)
      End Sub
  End Function

  Private Function GetActionForCreatingDailyTotalTableInMonth() As LoadTableCallBack
    Return _
      Sub(tabPage, record)
        Dim item As MonthlyItem = cboxDailyTotal.SelectedItem
        If item IsNot Nothing Then
          CreateDailyTotalTableInMonth(tabPage, record, item.Year, item.Month)
        End If
      End Sub
  End Function

  Delegate Function GetHandlerCallback(col As Integer, row As Integer, value As String) As Handler

  Private Function GetFuncForCreatingEventHandlerInTermRecord() As GetHandlerCallback
    Return _
      Function(col, row, value)
        If col = 0 Then
          Dim h As New Handler
          With h
            .DoubleClickCallBack =
              Sub(sender, e)
                Dim l As New List(Of TabInfo)
                For idx As Integer = 0 To ReadRecordTerm.MonthlyItemList.Count
                  Dim tabInfo As TabInfo = InnerTabPageInfoListInPersonalTab(idx)
                  Dim c As SheetRecordController = tabInfo.SheetRecordController
                  c.ReadCallback =
                    Function()
                      Return _
                        UserRecordManager.GetSheetRecord(
                          value, PersonalForm.tabInPersonalTab.SelectedTab.Text, GetFuncForFilteringImcompleteRecord())
                    End Function

                  Dim nTabInfo As TabInfo
                  With nTabInfo
                    .Name = tabInfo.Name
                    .SheetRecordController = c
                  End With

                  l.Add(nTabInfo)
                Next

                Dim tabInfoList As List(Of TabInfo) = connectTabPageWithInfo(PersonalForm.tabInPersonalTab, l)

                PersonalForm.Text = UserInfoList.Find(Function(info) info.GetIdNum = value).GetName
                PersonalForm.TabPageInfoList = tabInfoList
                PersonalForm.ShowDialog()
              End Sub
          End With
          Return h
        Else
          Return Nothing
        End If
      End Function
  End Function

  Private Function GetFuncForEmptyHandler() As GetHandlerCallback
    Return _
      Function(col, row, value)
        Return Nothing
      End Function
  End Function

  Private Function GetFuncForGettingCSVFileNameOfPersonal(ym As MonthlyItem) As Func(Of String)
    Return _
      Function()
        Dim user As ExpandedUserInfo = CType(cboxUserInfo.SelectedItem, ExpandedUserInfo)
        Return String.Format("個人データ{0}月分_{1}_{2}.csv", ym.Month, user.GetIdNum, user.GetName)
      End Function
  End Function

  Private Function GetCSVFileNameOfSumOfPersonal() As String
    Dim user As ExpandedUserInfo = CType(cboxUserInfo.SelectedItem, ExpandedUserInfo)
    Return String.Format("個人データ集計_{0}_{1}.csv", user.GetIdNum, user.GetName)
  End Function

  Private Function GetCSVFileNameOfDailyTerm() As String
    Dim y As Integer = dPicDailyTerm.Value.Year
    Dim m As Integer = dPicDailyTerm.Value.Month
    Dim d As Integer = dPicDailyTerm.Value.Day
    Return String.Format("日付データ_{0}年{1}月{2}日", y, m, d)
  End Function

  Private Function GetCSVFileNameOfWeeklyTerm() As String
    Dim item As WeeklyItem = cboxWeeklyTerm.SelectedItem
    If item IsNot Nothing Then
      Dim t As String = item.ToString
      Dim str As String = ""
      t.Split(" ").ToList.ForEach(Sub(Text) str = str & Text)
      Dim y As Integer = item.Year
      Return String.Format("日付データ_{0}年{1}", y, str)
    Else
      Return Nothing
    End If
  End Function

  Private Function GetCSVFileNameOfMonthlyTerm() As String
    Dim item As MonthlyItem = cboxMonthlyTerm.SelectedItem
    If item IsNot Nothing Then
      Dim m As Integer = item.Month
      Dim y As Integer = item.Year
      Return String.Format("日付データ_{0}年{1}月", y, m)
    Else
      Return Nothing
    End If
  End Function

  Private Function GetCSVFileNameOfDailyTotal() As String
    Dim item As MonthlyItem = cboxDailyTotal.SelectedItem
    If item IsNot Nothing Then
      Dim m As Integer = item.Month
      Dim y As Integer = item.Year
      Return String.Format("集計データ_日単位_{0}年{1}月", y, m)
    Else
      Return Nothing
    End If
  End Function

  Private Function Sort(record As SheetRecord, sortInfo As SortProperties) As SheetRecord
    Dim NONE As Double = If(sortInfo.Asc, Integer.MaxValue, Integer.MinValue)

    Dim f As Func(Of RowRecord, Double) =
      Function(row)
        If row.Count > sortInfo.Col Then
          Dim value As String = row.GetItem(sortInfo.Col)
          Return If(value = "", NONE, RecordConverter.ToDouble(value))
        Else
          Return NONE
        End If
      End Function

    If sortInfo.Col = -1 Then
      Return record
    ElseIf sortInfo.Asc Then
      Return record.OrderBy(f)
    Else
      Return record.OrderByDescending(f)
    End If
  End Function

  Private Function GetFuncForGettingRowColorInMonthlyTable(year As Integer, month As Integer) As Func(Of Integer, Color)
    Return _
      Function(row)
        If row = 31 Then
          Return Color.PaleGreen
        Else
          Dim week As Integer = MyCalendar.GetWeek(year, month, row + 1)
          Return If(week = 1 OrElse week = 7, Color.Pink, Color.Transparent)
        End If
      End Function
  End Function

  Private Function GetCSVTitlesOfPersonalRecord() As String
    Dim item1 As String = WAP.MANAGER.GetValue(WAP.KEY_ITEM_NAME1)
    Dim item2 As String = WAP.MANAGER.GetValue(WAP.KEY_ITEM_NAME2)
    Dim item3 As String = WAP.MANAGER.GetValue(WAP.KEY_ITEM_NAME3)
    Dim item4 As String = WAP.MANAGER.GetValue(WAP.KEY_ITEM_NAME4)
    Dim item5 As String = WAP.MANAGER.GetValue(WAP.KEY_ITEM_NAME5)
    Dim item6 As String = WAP.MANAGER.GetValue(WAP.KEY_ITEM_NAME6)
    Dim item7 As String = WAP.MANAGER.GetValue(WAP.KEY_ITEM_NAME7)
    Return String.Format(",,{0},,,{1},,,{2},,,{3},,,{4},,,{5},,,{6},,", item1, item2, item3, item4, item5, item6, item7)
  End Function

  Private Function GetCSVTitlesOfTotalRecord() As String
    Return "," & GetCSVTitlesOfPersonalRecord()
  End Function

  Private Function GetCSVSubTitlesOfPersonalRecord() As String
    Return "日にち," & GetCSVItems() & ",備考"
  End Function

  Private Function GetCSVSubTitlesOfTermRecord() As String
    Return "ID,名前," & GetCSVItems() & ","
  End Function

  Private Function GetCSVSubTitlesOfTotalRecord() As String
    Return ",日にち," & GetCSVItems() & ","
  End Function

  Private Function GetCSVItems() As String
    Dim csv As String = ""
    For idx As Integer = 0 To 6
      csv = csv & ",件数,時間,生産性"
    Next
    Return csv.Substring(1)
  End Function

  'Private Function GetRowColorInSumTable(row As Integer) As Color
  '  If row = 21 Then
  '    Return Color.PaleTurquoise
  '    'Return Color.Transparent
  '  ElseIf row Mod 7 = 6 AndAlso row < 21
  '    Return Color.PaleGreen
  '  Else
  '    Return Color.Transparent
  '  End If
  'End Function

  Private Function RoundDecStr(value As String) As String
    If MP.Utils.General.MyChar.IsDouble(value) Then
      Dim d As Double = Double.Parse(value)
      Return Math.Round(d, 2).ToString
    Else
      Return value
    End If
  End Function

  Private Function GetOffsetLocationInPersonalTab() As Point
    Return New Point(3, 6)
  End Function

  Private Function GetOffsetLocationInTermTab() As Point
    Return New Point(3, 46)
  End Function

  Private Function GetOffsetLocationInTotalTab() As Point
    Return New Point(3, 46)
  End Function

  Private Sub ShowPersonalTableTitles(showedTabPage As TabPage, offsetLocation As Point)
    ShowedTableTitles(showedTabPage, offsetLocation, RecordTableForm.tblTitles, RecordTableForm.tblSubTitles)
  End Sub

  Private Sub ShowTermTableTitles(showedTabPage As TabPage, offsetLocation As Point)
    ShowedTableTitles(showedTabPage, offsetLocation, TermRecordTableForm.tblTitles, TermRecordTableForm.tblSubTItles)
  End Sub

  Private Sub ShowTotalTableTitles(showedTabPage As TabPage, offsetLocation As Point)
    ShowedTableTitles(showedTabPage, offsetLocation, TotalRecordTableForm.tblTitles, TotalRecordTableForm.tblSubTItles)
  End Sub

  Private Sub ShowedTableTitles(showedTabPage As TabPage, offsetLocation As Point, titles As TableLayoutPanel, subTitles As TableLayoutPanel)
    titles.Location = offsetLocation
    subTitles.Location = New Point(offsetLocation.X, offsetLocation.Y + 26)

    showedTabPage.Controls.Add(titles)
    showedTabPage.Controls.Add(subTitles)
  End Sub

  Private Sub ShowPersonalTableRecord(showedTabPage As TabPage, offsetLocation As Point)
    ShowTableRecord(showedTabPage, offsetLocation, RecordTableForm.pnlForTable)
  End Sub

  Private Sub ShowTermTableRecord(showedTabPage As TabPage, offsetLocation As Point)
    ShowTableRecord(showedTabPage, offsetLocation, TermRecordTableForm.pnlForTable)
  End Sub

  Private Sub ShowTotalTableRecord(showedTabPage As TabPage, offsetLocation As Point)
    ShowTableRecord(showedTabPage, offsetLocation, TotalRecordTableForm.pnlForTable)
  End Sub

  Private Sub ShowTableRecord(showedTabPage As TabPage, offsetLocation As Point, scrollPanel As Panel)
    scrollPanel.Location = New Point(offsetLocation.X, offsetLocation.Y + 57)
    scrollPanel.AutoScroll = True
    showedTabPage.Controls.Add(scrollPanel)
  End Sub

  Private Sub InitSvaeFileDialog()
    SaveFileDialog = New SaveFileDialog()

    'はじめのファイル名を指定する
    SaveFileDialog.FileName = "新しいファイル.csv"
    'はじめに表示されるフォルダを指定する
    'SaveFileDialog.InitialDirectory = MP.Details.Sys.App.GetCurrentDirectory
    '[ファイルの種類]に表示される選択肢を指定する
    SaveFileDialog.Filter = "csvファイル(*.csv)|*.csv|textファイル(*.txt)|*.txt|すべてのファイル(*.*)|*.*"
    SaveFileDialog.FilterIndex = 0
    'タイトルを設定する
    SaveFileDialog.Title = "保存先のファイルを選択してください"
    'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
    SaveFileDialog.RestoreDirectory = True
    '既に存在するファイル名を指定したとき警告する
    'デフォルトでTrueなので指定する必要はない
    SaveFileDialog.OverwritePrompt = True
    '存在しないパスが指定されたとき警告を表示する
    'デフォルトでTrueなので指定する必要はない
    SaveFileDialog.CheckPathExists = True
  End Sub

  Private Sub cmdOutputCSV_Click(sender As Object, e As EventArgs) Handles cmdOutputCSV.Click
    If CurrentlyShowedSheetRecordManager Is Nothing Then
      MessageBox.Show("データが表示されていません。")
    Else
      SaveFileDialog.FileName = CurrentlyShowedSheetRecordManager.GetCSVFileName()

      'ダイアログを表示する
      If SaveFileDialog.ShowDialog() = DialogResult.OK Then
        Dim stream As System.IO.Stream = Nothing
        Dim sw As System.IO.StreamWriter = Nothing
        Try
          stream = SaveFileDialog.OpenFile()
          If Not (stream Is Nothing) Then
            'ファイルに書き込む
            sw = New System.IO.StreamWriter(stream, System.Text.Encoding.GetEncoding("Shift_JIS"))
            sw.Write(CurrentlyShowedSheetRecordManager.ToCSV)
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

  Public Structure SheetRecordController
    Dim TabPage As TabPage
    Dim ReadCallback As ReadRecordCallback
    Dim AddSumOfRecordCallback As AddSumOfRowRecordCallback
    Dim LoadTableCallback As LoadTableCallBack
    Dim CanSort As Boolean
    Dim GetCSVFileNameCallback As Func(Of String)
    Dim GetCSVTitlesCallback As Func(Of String)
    Dim GetCSVSubTitlesCallback As Func(Of String)
  End Structure

  Public Class SheetRecordManager
    Private Controller As SheetRecordController
    Private SheetRecord As SheetRecord
    Private SortProps As SortProperties

    Public Sub New(controller As SheetRecordController)
      Me.Controller = controller
      Me.SheetRecord = Nothing
      Dim p As SortProperties
      With p
        .Col = -1
        .Asc = True
      End With
      SortProps = p
    End Sub

    Public Function GetTabPage() As TabPage
      Return Controller.TabPage
    End Function

    Public Function AllowSort() As Boolean
      Return Controller.CanSort
    End Function

    Public Function ReadSheetRecord() As SheetRecordManager
      SheetRecord = Controller.ReadCallback()
      Return Me
    End Function

    Public Function LoadTable(tabPage As TabPage) As SheetRecordManager
      Dim rec As SheetRecord = GetRecordThatAddSumRow()
      If rec IsNot Nothing Then
        Controller.LoadTableCallback(Controller.TabPage, GetRecordThatAddSumRow())
      End If
      Return Me
    End Function

    Private Function GetRecordThatAddSumRow() As SheetRecord
      If SheetRecord IsNot Nothing Then
        Return Controller.AddSumOfRecordCallback(SheetRecord)
      Else
        Return Nothing
      End If
    End Function

    Public Function Sort(col As Integer) As SheetRecordManager
      If Controller.CanSort Then
        Dim p As SortProperties
        With p
          Dim idx As Integer = If(col = 0, -1, col)
          .Col = idx
          .Asc = If(SortProps.Col = idx, Not SortProps.Asc, False)
        End With

        SheetRecord = SortedRecord(p)
        SortProps = p
        Return Me
      Else
        Return Me
      End If
    End Function

    Private Function SortedRecord(prop As SortProperties) As SheetRecord
      If SheetRecord IsNot Nothing Then
        Dim NONE As Double = If(prop.Asc, Integer.MaxValue, Integer.MinValue)

        Dim f As Func(Of RowRecord, Double) =
          Function(row)
            If row.Count > prop.Col Then
              Dim value As String = row.GetItem(prop.Col)
              Return If(value = "", NONE, RecordConverter.ToDouble(value))
            Else
              Return NONE
            End If
          End Function

        If prop.Col = -1 Then
          Return Controller.ReadCallback()
        ElseIf prop.Asc = True
          Return SheetRecord.OrderBy(f)
        Else
          Return SheetRecord.OrderByDescending(f)
        End If
      Else
        Return Nothing
      End If
    End Function

    Public Function GetCSVFileName() As String
      Return Controller.GetCSVFileNameCallback()
    End Function

    Public Function ToCSV() As String
      Dim rec As SheetRecord = GetRecordThatAddSumRow()
      If rec IsNot Nothing Then
        Return _
          Controller.GetCSVTitlesCallback() & vbCrLf &
          Controller.GetCSVSubTitlesCallback() & vbCrLf &
          RecordConverter.ToCSV(rec)
      Else
        Return ""
      End If
    End Function

    Public Function Clone(tabPage As TabPage) As SheetRecordManager
      Dim c As SheetRecordController = Controller
      c.TabPage = tabPage
      Return New SheetRecordManager(c)
    End Function
  End Class

End Class

