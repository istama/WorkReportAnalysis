
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
  Private Shared TAB_GROUNP_NAME_OF_PERSONAL_RECORD_TAB = "PersonalRecordTab"
  Private Shared TAB_GROUNP_NAME_OF_TERM_RECORD_TAB = "TermRecordTab"
  Private Shared TAB_GROUNP_NAME_OF_TOTAL_RECORD_TAB = "TotalRecordTab"

  Public Structure TabInfo
    Dim TabGroupName As String
    Dim Name As String
    Dim TableRecordController As TableRecordController
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

  Private InnerTabPageInfoList As List(Of TabInfo)
  'Private InnerTabPageInfoListInPersonalTab As List(Of TabInfo)
  'Private InnerTabPageInfoListInTermTab As List(Of TabInfo)
  'Private InnerTabPageInfoListInTotalTab As List(Of TabInfo)

  Private SaveFileDialog As SaveFileDialog
  Private CurrentlyShowedSheetRecordManager As SheetRecordManager = Nothing

  Private Loaded As Boolean = False

  Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Try
      LoadAllUsers()

      ReadRecordTerm = New ReadRecordTerm(10, FileFormat.GetYear, 12, FileFormat.GetYear)

      Excel = New ExcelAccessor(New Excel())
      Excel.Init()
      UserRecordManager = New UserRecordManager(UserInfoList, ReadRecordTerm)
      UserRecordLoader = New UserRecordLoader(Excel, UserRecordManager)

      InnerTabPageInfoList = New List(Of TabInfo)
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

      LoadPersonalRecord()
    Catch ex As Exception
      MsgBox.ShowError(ex)
    End Try

    Loaded = True
  End Sub

  Private Sub LoadAllUsers()
    Dim a As SerializedAccessor = MySerialize.GenerateAccessor()
    Dim path As String = FilePath.UserinfoFilePath()

    UserInfoList = a.GetInfo(Of UserInfo)(path).ConvertAll(Function(user) New ExpandedUserInfo(user))
  End Sub

  Private Sub InitInnerTabPageInPersonalTab()
    Dim infoList = New List(Of TabInfo)()

    Dim layouts As TableLayoutObjects
    With layouts
      .TabPage = Nothing
      .GetTitlesTablePanelCallback = Function() RecordTableForm.tblTitles
      .GetSubTitlesTablePanelCallback = Function() RecordTableForm.tblSubTitles
      .GetRecordTablePanelCallback = Function() RecordTableForm.tblRecord
      .ShowedTableOffsetLocation = GetOffsetLocationInPersonalTab()
    End With

    Dim recCtrl As SheetRecordController
    With recCtrl
      .ReadCallback = GetFuncForGettingPersonalRecord(tabInPersonalTab)
      .AddSumOfRecordCallback = AddressOf AddSumOfPersonalRecord
      .CanSort = True
    End With

    For Each ym As MonthlyItem In ReadRecordTerm.MonthlyItemList.ToList
      Dim tab As TabInfo
      With tab
        .TabGroupName = TAB_GROUNP_NAME_OF_PERSONAL_RECORD_TAB
        .Name = UserRecordManager.GetSheetName(ym.Month)
        Dim c As TableRecordController
        With c
          .TableLayout = layouts
          .Record = recCtrl
          Dim csv As CSVController
          With csv
            .GetCSVFileNameCallback = GetFuncForGettingCSVFileNameOfPersonal(ym)
            .GetCSVTitlesCallback = AddressOf GetCSVTitlesOfPersonalRecord
            .GetCSVSubTitlesCallback = AddressOf GetCSVSubTitlesOfPersonalRecord
          End With
          .CSV = csv
          .LoadTableCallback = GetActionForCreatingPersonalDailyTableInMonth(ym.Year, ym.Month)
        End With
        .TableRecordController = c
      End With
      infoList.Add(tab)
    Next

    Dim sumTab As TabInfo
    With sumTab
      .TabGroupName = TAB_GROUNP_NAME_OF_PERSONAL_RECORD_TAB
      .Name = UserRecordManager.GetSumSheetName
      Dim c As TableRecordController
      With c
        Dim l As TableLayoutObjects
        With l
          .TabPage = Nothing
          .GetTitlesTablePanelCallback = Function() TotalRecordTableForm.tblTitles
          .GetSubTitlesTablePanelCallback = Function() TotalRecordTableForm.tblSubTitles
          .GetRecordTablePanelCallback = Function() TotalRecordTableForm.tblRecord
          .ShowedTableOffsetLocation = GetOffsetLocationInPersonalTab()
        End With
        .TableLayout = l

        Dim r As SheetRecordController
        With r
          .ReadCallback = GetFuncForGettingPersonalRecord(tabInPersonalTab)
          .AddSumOfRecordCallback = Function(sheetRecord) sheetRecord
          .CanSort = False
        End With
        .Record = r

        Dim cc As CSVController
        With cc
          .GetCSVFileNameCallback = AddressOf GetCSVFileNameOfSumOfPersonal
          .GetCSVTitlesCallback = AddressOf GetCSVTitlesOfTotalRecord
          .GetCSVSubTitlesCallback = AddressOf GetCSVSubTitlesOfTotalRecord
        End With
        .CSV = cc

        .LoadTableCallback = AddressOf CreateSumTable
      End With
      .TableRecordController = c
    End With
    infoList.Add(sumTab)

    InnerTabPageInfoList.AddRange(connectTabPageWithInfo(tabInPersonalTab, infoList))
  End Sub

  Private Sub InitInnerTabPageInTermTab()
    Dim infoList = New List(Of TabInfo)()

    Dim layouts As TableLayoutObjects
    With layouts
      .TabPage = Nothing
      .GetTitlesTablePanelCallback = Function() TermRecordTableForm.tblTitles
      .GetSubTitlesTablePanelCallback = Function() TermRecordTableForm.tblSubTitles
      .GetRecordTablePanelCallback = Function() TermRecordTableForm.tblRecord
      .ShowedTableOffsetLocation = GetOffsetLocationInTermTab()
    End With

    Dim rec As SheetRecordController
    With rec
      .ReadCallback = AddressOf GetDailyTermRecord
      .AddSumOfRecordCallback = AddressOf AddSumOfTermRecord
      .CanSort = True
    End With

    Dim csv As CSVController
    With csv
      .GetCSVFileNameCallback = AddressOf GetCSVFileNameOfDailyTerm
      .GetCSVTitlesCallback = AddressOf GetCSVTitlesOfTotalRecord
      .GetCSVSubTitlesCallback = AddressOf GetCSVSubTitlesOfTermRecord
    End With

    Dim tabDays As TabInfo
    With tabDays
      .TabGroupName = TAB_GROUNP_NAME_OF_TERM_RECORD_TAB
      .Name = "日"
      Dim c As TableRecordController
      With c
        .TableLayout = layouts
        .Record = rec
        .CSV = csv
        .LoadTableCallback = AddressOf CreateTermTable
      End With
      .TableRecordController = c
    End With

    Dim tabWeeks As TabInfo
    With tabWeeks
      .TabGroupName = TAB_GROUNP_NAME_OF_TERM_RECORD_TAB
      .Name = "週"
      Dim c As TableRecordController
      With c
        .TableLayout = layouts

        Dim recWeek = rec
        recWeek.ReadCallback = AddressOf GetWeeklyTermRecord
        .Record = recWeek

        Dim cWeek = csv
        cWeek.GetCSVFileNameCallback = AddressOf GetCSVFileNameOfWeeklyTerm
        .CSV = cWeek

        .LoadTableCallback = AddressOf CreateTermTable
      End With
      .TableRecordController = c
    End With

    Dim tabMonths As TabInfo
    With tabMonths
      .TabGroupName = TAB_GROUNP_NAME_OF_TERM_RECORD_TAB
      .Name = "月"
      Dim c As TableRecordController
      With c
        .TableLayout = layouts

        Dim recMonth = rec
        recMonth.ReadCallback = AddressOf GetMonthlyTermRecord
        .Record = recMonth

        Dim cMonth = csv
        cMonth.GetCSVFileNameCallback = AddressOf GetCSVFileNameOfMonthlyTerm
        .CSV = cMonth

        .LoadTableCallback = AddressOf CreateTermTable
      End With
      .TableRecordController = c
    End With

    Dim tabYear As TabInfo
    With tabYear
      .TabGroupName = TAB_GROUNP_NAME_OF_TERM_RECORD_TAB
      .Name = "合計"
      Dim c As TableRecordController
      With c
        .TableLayout = layouts

        Dim recAll = rec
        recAll.ReadCallback = AddressOf GetAllTermRecord
        .Record = recAll

        Dim cYear = csv
        cYear.GetCSVFileNameCallback = Function() String.Format("日付データ_{0}年.csv", FileFormat.GetYear())
        .CSV = cYear

        .LoadTableCallback = AddressOf CreateTermTable
      End With
      .TableRecordController = c
    End With

    infoList.Add(tabDays)
    infoList.Add(tabWeeks)
    infoList.Add(tabMonths)
    infoList.Add(tabYear)

    InnerTabPageInfoList.AddRange(connectTabPageWithInfo(tabInTermTab, infoList))
  End Sub

  Private Sub InitInnerTabPageInTotalTab()
    Dim infoList = New List(Of TabInfo)()

    Dim layouts As TableLayoutObjects
    With layouts
      .TabPage = Nothing
      .GetTitlesTablePanelCallback = Function() TotalRecordTableForm.tblTitles
      .GetSubTitlesTablePanelCallback = Function() TotalRecordTableForm.tblSubTitles
      .GetRecordTablePanelCallback = Function() TotalRecordTableForm.tblRecord
      .ShowedTableOffsetLocation = GetOffsetLocationInTotalTab()
    End With

    Dim rec As SheetRecordController
    With rec
      .ReadCallback = AddressOf GetDailyTotalRecord
      .AddSumOfRecordCallback = AddressOf AddSumOfTermRecord
      .CanSort = True
    End With

    Dim csv As CSVController
    With csv
      .GetCSVFileNameCallback = AddressOf GetCSVFileNameOfDailyTotal
      .GetCSVTitlesCallback = AddressOf GetCSVTitlesOfTotalRecord
      .GetCSVSubTitlesCallback = AddressOf GetCSVSubTitlesOfTotalRecord
    End With

    Dim tabDays As TabInfo
    With tabDays
      .TabGroupName = TAB_GROUNP_NAME_OF_TOTAL_RECORD_TAB
      .Name = "日"
      Dim c As TableRecordController
      With c
        .TableLayout = layouts
        .Record = rec
        .CSV = csv
        .LoadTableCallback = GetActionForCreatingDailyTotalTableInMonth()
      End With
      .TableRecordController = c
    End With

    Dim tabWeeks As TabInfo
    With tabWeeks
      .TabGroupName = TAB_GROUNP_NAME_OF_TOTAL_RECORD_TAB
      .Name = "週"
      Dim c As TableRecordController
      With c
        .TableLayout = layouts

        Dim recWeek As SheetRecordController = rec
        recWeek.ReadCallback = AddressOf GetWeeklyTotalRecord
        .Record = recWeek

        Dim cWeek As CSVController = csv
        cWeek.GetCSVFileNameCallback = Function() String.Format("集計データ_週単位_{0}年", FileFormat.GetYear())
        .CSV = cWeek

        .LoadTableCallback = GetActionForCreatingWeeklyTotalTable()
      End With
      .TableRecordController = c
    End With

    Dim tabMonths As TabInfo
    With tabMonths
      .TabGroupName = TAB_GROUNP_NAME_OF_TOTAL_RECORD_TAB
      .Name = "月"
      Dim c As TableRecordController
      With c
        .TableLayout = layouts

        Dim recMonth As SheetRecordController = rec
        recMonth.ReadCallback = AddressOf GetMonthlyTotalRecord
        .Record = recMonth

        Dim cMonth As CSVController = csv
        cMonth.GetCSVFileNameCallback = Function() String.Format("集計データ_月単位_{0}年", FileFormat.GetYear())
        .CSV = cMonth

        .LoadTableCallback = GetActionForCreatingMonthlyTotalTable()
      End With
      .TableRecordController = c
    End With

    infoList.Add(tabDays)
    infoList.Add(tabWeeks)
    infoList.Add(tabMonths)

    InnerTabPageInfoList.AddRange(connectTabPageWithInfo(tabInTotalTab, infoList))
  End Sub

  Private Function connectTabPageWithInfo(tab As TabControl, tabPageInfoList As List(Of TabInfo)) As List(Of TabInfo)
    Dim l As New List(Of TabInfo)

    If tab.TabPages.Count = tabPageInfoList.Count Then
      For idx As Integer = 0 To tabPageInfoList.Count - 1
        Dim info As TabInfo = tabPageInfoList(idx)
        tab.TabPages.Item(idx).Text = info.Name
        info.TableRecordController.TableLayout.TabPage = tab.TabPages.Item(idx)
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

    SetSortCallback(RecordTableForm.tblSubTitles, 0, RecordTableForm.tblSubTitles.ColumnCount - 2)
    SetSortCallback(TermRecordTableForm.tblSubTitles, 0, TermRecordTableForm.tblSubTitles.ColumnCount - 2)
    SetSortCallback(TotalRecordTableForm.tblSubTitles, 0, TotalRecordTableForm.tblSubTitles.ColumnCount - 2)
  End Sub

  Private Sub SetSortCallback(table As TableLayoutPanel, startIdx As Integer, endIdx As Integer)
    For i As Integer = startIdx To endIdx
      AddHandler table.GetControlFromPosition(i, 0).Controls.Item(0).Click, AddressOf SortRecordCallback
    Next
  End Sub

  Private Sub SortRecordCallback(sender As Object, e As EventArgs)
    If CurrentlyShowedSheetRecordManager IsNot Nothing AndAlso CurrentlyShowedSheetRecordManager.AllowSort Then
      Dim label As Label = CType(sender, Label)
      Dim col As Integer = CurrentlyShowedSheetRecordManager.GetColPos(label.Parent)
      LoadTableInnerTabPage(CurrentlyShowedSheetRecordManager.Sort(col))
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
      LoadPersonalRecord()
    ElseIf idx = 1
      LoadTermRecord()
    Else
      LoadTotalRecord()
    End If
  End Sub

  Private Sub tabInPersonalTab_PageChanged(sender As Object, e As EventArgs) Handles tabInPersonalTab.SelectedIndexChanged
    LoadPersonalRecord()
  End Sub

  Private Sub tabInTermTab_PageChanged(sender As Object, e As EventArgs) Handles tabInTermTab.SelectedIndexChanged
    LoadTermRecord()
  End Sub

  Private Sub tabInTotalTab_PageChanged(sender As Object, e As EventArgs) Handles tabInTotalTab.SelectedIndexChanged
    LoadTotalRecord()
  End Sub

  Private Sub LoadPersonalRecord()
    LoadRecordInTabPage(InnerTabPageInfoList, tabInPersonalTab.SelectedTab)
  End Sub

  Private Sub LoadTermRecord()
    LoadRecordInTabPage(InnerTabPageInfoList, tabInTermTab.SelectedTab)
  End Sub

  Private Sub LoadTotalRecord()
    LoadRecordInTabPage(InnerTabPageInfoList, tabInTotalTab.SelectedTab)
  End Sub

  Public Sub LoadRecordInTabPage(tabPageInfoList As List(Of TabInfo), selectedInnerTab As TabPage)
    Try
      If tabPageInfoList.Exists(Function(e) e.TableRecordController.TableLayout.TabPage.Equals(selectedInnerTab)) Then
        Dim tabInfo As TabInfo = tabPageInfoList.Find(Function(e) e.TableRecordController.TableLayout.TabPage.Equals(selectedInnerTab))
        'MessageBox.Show(tabInfo.Name)

        ShowRecord(tabInfo)
      Else
        Throw New Exception("タブ情報が見つかりません")
      End If
    Catch ex As Exception
      MsgBox.ShowError(ex)
    End Try
  End Sub

  Private Sub ShowRecord(tabInfo As TabInfo)
    Dim recManager As SheetRecordManager = New SheetRecordManager(tabInfo.TableRecordController)
    LoadTableInnerTabPage(recManager.ReadSheetRecord)
  End Sub

  Private Sub LoadTableInnerTabPage(recordManager As SheetRecordManager)
    CurrentlyShowedSheetRecordManager = recordManager.LoadTable()
  End Sub

  Private Sub cboxUserInfo_SelectedIndexChanged(sender As Object, e As EventArgs) _
    Handles _
      cboxUserInfo.SelectedIndexChanged,
      dPicDailyTerm.ValueChanged,
      cboxWeeklyTerm.SelectedIndexChanged,
      cboxMonthlyTerm.SelectedIndexChanged,
      cboxDailyTotal.SelectedIndexChanged
    LoadTableInnerTabPage(CurrentlyShowedSheetRecordManager.ReadSheetRecord)
  End Sub

  Private Sub chkExcludeIncompleteRecordFromSum_CheckedChanged(sender As Object, e As EventArgs) Handles chkExcludeIncompleteRecordFromSum.CheckedChanged
    If Loaded Then
      Call tabMaster_PageChanged(sender, e)
    End If
  End Sub

  Private Sub CreateDailyTableInMonth(table As TableLayoutPanel, record As SheetRecord, year As Integer, month As Integer)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(RecordTableForm.pnlForTable).
        SetTextCols(0).
        SetNoteCols(RecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(GetFuncForGettingRowColorInMonthlyTable(year, month))

    CreateTable(table, record, tableDrawer, GetFuncForEmptyHandler)
    'ShowPersonalTableRecord(tabPage, GetOffsetLocationInPersonalTab)
  End Sub

  Private Sub CreateSumTable(table As TableLayoutPanel, record As SheetRecord)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(TotalRecordTableForm.pnlForTable).
        SetTextCols(0, 1).
        SetNoteCols(TotalRecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(Function(row) If(row = record.Count - 1, Color.PaleGreen, Color.Transparent))

    CreateTable(table, record, tableDrawer, GetFuncForCreatingEventHandlerInTotalRecord)
    'ShowTotalTableRecord(tabPage, GetOffsetLocationInPersonalTab)
  End Sub

  Private Sub CreateTermTable(table As TableLayoutPanel, record As SheetRecord)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(TermRecordTableForm.pnlForTable).
        SetTextCols(0, 1).
        SetNoteCols(TermRecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(Function(row) If(row = record.Count - 1, Color.PaleGreen, Color.Transparent))

    CreateTable(table, record, tableDrawer, GetFuncForCreatingEventHandlerInTermRecord)
    'ShowTermTableRecord(tabPage, GetOffsetLocationInTermTab)
  End Sub

  Private Sub CreateDailyTotalTableInMonth(table As TableLayoutPanel, record As SheetRecord, year As Integer, month As Integer)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(TotalRecordTableForm.pnlForTable).
      SetTextCols(0, 1).
      SetNoteCols(TermRecordTableForm.tblRecord.ColumnCount - 1).
      SetFuncBackColor(GetFuncForGettingRowColorInMonthlyTable(year, month))

    CreateTable(table, record, tableDrawer, GetFuncForCreatingEventHandlerInTotalRecord())
    'ShowTotalTableRecord(tabPage, GetOffsetLocationInTotalTab)
  End Sub

  Private Sub CreateTotalTable(table As TableLayoutPanel, record As SheetRecord, callback As GetHandlerCallback)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(TotalRecordTableForm.pnlForTable).
        SetTextCols(0, 1).
        SetNoteCols(RecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(Function(row) If(row = record.Count - 1, Color.PaleGreen, Color.Transparent))

    CreateTable(table, record, tableDrawer, callback)
    'ShowTotalTableRecord(tabPage, GetOffsetLocationInTotalTab)
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

  Private Function GetFuncForFilteringImcompleteRecord(checkBox As CheckBox) As Func(Of RowRecord, Integer, Boolean)
    Dim isFiltering As Boolean = checkBox.Checked
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

  Private Function GetFuncForGettingPersonalRecord(tab As TabControl) As ReadRecordCallback
    Return _
      Function()
        Dim page As TabPage = tab.SelectedTab
        Dim selectedUser As ExpandedUserInfo = cboxUserInfo.SelectedItem

        If page IsNot Nothing AndAlso selectedUser IsNot Nothing Then
          If Not UserRecordManager.ContainsRecord(selectedUser.GetIdNum) Then
            UserRecordLoader.Load(selectedUser)
          End If

          Return UserRecordManager.GetSheetRecord(selectedUser.GetIdNum, page.Text, GetFuncForFilteringImcompleteRecord(chkExcludeIncompleteRecordFromSum))
        Else
          Return Nothing
        End If
      End Function
  End Function

  Private Function GetDailyTermRecord() As SheetRecord
    If ReadRecordTerm.InMonthlyTerm(dPicDailyTerm.Value.Month, dPicDailyTerm.Value.Year) Then
      Dim y As Integer = dPicDailyTerm.Value.Year
      Dim m As Integer = dPicDailyTerm.Value.Month
      Dim d As Integer = dPicDailyTerm.Value.Day
      Return UserRecordManager.GetDailyTermRecord(d, m, y)
    Else
      Return Nothing
    End If
  End Function

  Private Function GetWeeklyTermRecord() As SheetRecord
    Dim item As WeeklyItem = cboxWeeklyTerm.SelectedItem
    If item IsNot Nothing Then
      Dim w As Integer = item.Week
      Dim m As Integer = item.Month
      Dim y As Integer = item.Year
      Return UserRecordManager.GetWeeklyTermRecord(w, m, y, GetFuncForFilteringImcompleteRecord(chkExcludeIncompleteRecordFromSum))
    Else
      Return Nothing
    End If
  End Function

  Private Function GetMonthlyTermRecord() As SheetRecord
    Dim item As MonthlyItem = cboxMonthlyTerm.SelectedItem
    If item IsNot Nothing Then
      Dim m As Integer = item.Month
      Dim y As Integer = item.Year
      Return UserRecordManager.GetMonthlyTermRecord(m, y, GetFuncForFilteringImcompleteRecord(chkExcludeIncompleteRecordFromSum))
    Else
      Return Nothing
    End If
  End Function

  Private Function GetAllTermRecord() As SheetRecord
    Return UserRecordManager.GetAllTermRecord(FileFormat.GetYear(), GetFuncForFilteringImcompleteRecord(chkExcludeIncompleteRecordFromSum))
  End Function

  Private Function GetDailyTotalRecord() As SheetRecord
    Dim item As MonthlyItem = cboxDailyTotal.SelectedItem
    If item IsNot Nothing Then
      Dim m As Integer = item.Month
      Dim y As Integer = item.Year
      Return UserRecordManager.GetDailyTotalRecord(m, y, GetFuncForFilteringImcompleteRecord(chkExcludeIncompleteRecordFromSum))
    Else
      Return Nothing
    End If
  End Function

  Private Function GetWeeklyTotalRecord() As SheetRecord
    Return UserRecordManager.GetWeeklyTotalRecord(FileFormat.GetYear, GetFuncForFilteringImcompleteRecord(chkExcludeIncompleteRecordFromSum))
  End Function

  Private Function GetMonthlyTotalRecord() As SheetRecord
    Return UserRecordManager.GetMonthlyTotalRecord(FileFormat.GetYear, GetFuncForFilteringImcompleteRecord(chkExcludeIncompleteRecordFromSum))
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
        GetFuncForFilteringImcompleteRecord(chkExcludeIncompleteRecordFromSum)
        ).AddRangeToHead(headerCellValues)
  End Function

  Delegate Sub LoadTableCallBack(table As TableLayoutPanel, record As SheetRecord)

  Private Function GetActionForCreatingPersonalDailyTableInMonth(year As Integer, month As Integer) As LoadTableCallBack
    Return _
      Sub(table, record)
        CreateDailyTableInMonth(table, record, year, month)
      End Sub
  End Function

  Private Function GetActionForCreatingDailyTotalTableInMonth() As LoadTableCallBack
    Return _
      Sub(table, record)
        Dim item As MonthlyItem = cboxDailyTotal.SelectedItem
        If item IsNot Nothing Then
          CreateDailyTotalTableInMonth(table, record, item.Year, item.Month)
        End If
      End Sub
  End Function

  Private Function GetActionForCreatingWeeklyTotalTable() As LoadTableCallBack
    Return _
      Sub(table, record)
        CreateTotalTable(table, record, GetFuncForCreatingEventHandlerInTotalRecord)
      End Sub
  End Function

  Private Function GetActionForCreatingMonthlyTotalTable() As LoadTableCallBack
    Return _
      Sub(table, record)
        CreateTotalTable(table, record, GetFuncForCreatingEventHandlerInTotalRecord)
      End Sub
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

  Delegate Function GetHandlerCallback(col As Integer, row As Integer, value As String) As Handler

  Private Function GetFuncForCreatingEventHandlerInTermRecord() As GetHandlerCallback
    Return _
      Function(col, row, value)
        If col = 0 Then
          Dim h As New Handler
          With h
            .DoubleClickCallBack =
              Sub(sender, e)
                Dim id As String = CType(sender, Label).Text.PadLeft(3, "0"c)

                Dim userInfo As ExpandedUserInfo = UserInfoList.Find(Function(info) info.GetIdNum = id)

                If userInfo IsNot Nothing Then
                  Dim tabInfoList As New List(Of TabInfo)
                  InnerTabPageInfoList.
                    FindAll(Function(tabInfo) tabInfo.TabGroupName = TAB_GROUNP_NAME_OF_PERSONAL_RECORD_TAB).
                    ForEach(
                      Sub(t)
                        Dim tabInfo As TabInfo = t
                        Dim n As String = tabInfo.Name
                        tabInfo.TableRecordController.Record.ReadCallback =
                          Function()
                            Dim page As TabPage = tabInPersonalTab.SelectedTab
                            Return If(
                              page IsNot Nothing,
                              UserRecordManager.GetSheetRecord(
                                id,
                                page.Text,
                                GetFuncForFilteringImcompleteRecord(SubForm.chkExcludeIncompleteRecordFromSum)),
                              Nothing)
                          End Function

                        tabInfo.TableRecordController.CSV.GetCSVFileNameCallback =
                          Function() String.Format("個人データ{0}_{1}_{2}.csv", n, userInfo.GetIdNum, userInfo.GetName)

                        tabInfoList.Add(tabInfo)
                      End Sub)

                  ShowSubForm(userInfo.GetName, tabInPersonalTab, PersonalTab, tabInfoList)
                End If
              End Sub
          End With
          Return h
        Else
          Return Nothing
        End If
      End Function
  End Function

  Private Function GetFuncForCreatingEventHandlerInTotalRecord() As GetHandlerCallback
    Return _
      Function(col, row, value)
        If col = 1 Then
          Dim h As New Handler
          With h
            .DoubleClickCallBack =
              Sub(sender, e)
                Dim isHit As Boolean = False
                Dim label As Label = CType(sender, Label)

                Dim wIdx As Integer = GetIndexOfControlItems(cboxWeeklyTerm.Items, label.Text)
                If wIdx >= 0 Then
                  cboxWeeklyTerm.SelectedIndex = wIdx
                  tabInTermTab.SelectedIndex = 1
                  isHit = True
                Else
                  Dim mIdx As Integer = GetIndexOfControlItems(cboxMonthlyTerm.Items, label.Text)
                  If mIdx >= 0 Then
                    cboxMonthlyTerm.SelectedIndex = mIdx
                    tabInTermTab.SelectedIndex = 2
                    isHit = True
                  Else
                    Dim selectedIdx As Integer = tabInTotalTab.SelectedIndex
                    If selectedIdx = 0 Then
                      If label.Text.Length > 1 Then
                        Dim day As String = label.Text.Substring(0, label.Text.Length - 1)
                        If Utils.General.MyChar.IsInteger(day) AndAlso day >= 1 AndAlso day < 32 Then
                          Dim dItem = cboxDailyTotal.SelectedItem
                          If dItem IsNot Nothing Then
                            Dim month As Integer = CType(dItem, MonthlyItem).Month
                            dPicDailyTerm.Value = New DateTime(FileFormat.GetYear, month, day)
                            tabInTermTab.SelectedIndex = 0
                            isHit = True
                          End If
                        End If
                      End If
                    Else
                    End If
                  End If
                End If

                If isHit Then
                  Dim tabInfoList As List(Of TabInfo) =
                    InnerTabPageInfoList.
                      FindAll(Function(info) info.TabGroupName = TAB_GROUNP_NAME_OF_TERM_RECORD_TAB)

                  ShowSubForm("日付データ", tabInTermTab, TermTab, tabInfoList)
                End If
              End Sub
          End With
          Return h
        Else
          Return Nothing
        End If
      End Function
  End Function

  Private Sub ShowSubForm(formName As String, showedTab As TabControl, parentPageOfShowedTab As TabPage, tabInfoListInSubForm As List(Of TabInfo))
    Dim tmpManager As SheetRecordManager = CurrentlyShowedSheetRecordManager
    Dim tmpPoint As Point = showedTab.Location
    Dim tmpTabPageInfoList = InnerTabPageInfoList

    SubForm.Text = formName

    showedTab.Location = New Point(0, 0)
    SubForm.Tab = showedTab
    SubForm.Controls.Add(showedTab)

    Dim subList = connectTabPageWithInfo(showedTab, tabInfoListInSubForm)
    InnerTabPageInfoList = subList
    SubForm.TabPageInfoList = subList

    SubForm.ShowDialog()

    CurrentlyShowedSheetRecordManager = tmpManager
    showedTab.Location = tmpPoint
    parentPageOfShowedTab.Controls.Add(showedTab)
    InnerTabPageInfoList = tmpTabPageInfoList
  End Sub

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

  Private Function RoundDecStr(value As String) As String
    If MP.Utils.General.MyChar.IsDouble(value) Then
      Dim d As Double = Double.Parse(value)
      Return Math.Round(d, 2).ToString
    Else
      Return value
    End If
  End Function

  Private Function GetIndexOfControlItems(items As IList, text As String) As Integer
    For i As Integer = 0 To items.Count - 1
      If items(i).ToString = text Then
        Return i
      End If
    Next
    Return -1
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
    SaveCSVFile()
  End Sub

  Public Sub SaveCSVFile()
    If CurrentlyShowedSheetRecordManager Is Nothing OrElse Not CurrentlyShowedSheetRecordManager.ExistsSheetRecord Then
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

  Public Structure TableLayoutObjects
    Dim TabPage As TabPage
    Dim GetTitlesTablePanelCallback As Func(Of TableLayoutPanel)
    Dim GetSubTitlesTablePanelCallback As Func(Of TableLayoutPanel)
    Dim GetRecordTablePanelCallback As Func(Of TableLayoutPanel)
    Dim ShowedTableOffsetLocation As Point
  End Structure

  Public Structure SheetRecordController
    Dim ReadCallback As ReadRecordCallback
    Dim AddSumOfRecordCallback As AddSumOfRowRecordCallback
    Dim CanSort As Boolean
  End Structure

  Public Structure CSVController
    Dim GetCSVFileNameCallback As Func(Of String)
    Dim GetCSVTitlesCallback As Func(Of String)
    Dim GetCSVSubTitlesCallback As Func(Of String)
  End Structure

  Public Structure TableRecordController
    Dim TableLayout As TableLayoutObjects
    Dim Record As SheetRecordController
    Dim CSV As CSVController
    Dim LoadTableCallback As LoadTableCallBack
  End Structure

  Public Class SheetRecordManager
    Private Controller As TableRecordController
    Private SheetRecord As SheetRecord
    Private SortProps As SortProperties

    Public Sub New(controller As TableRecordController)
      Me.Controller = controller
      Me.SheetRecord = Nothing
      Dim p As SortProperties
      With p
        .Col = -1
        .Asc = True
      End With
      SortProps = p
    End Sub

    Public Function ExistsSheetRecord() As Boolean
      Return SheetRecord IsNot Nothing
    End Function

    Public Function GetTabPage() As TabPage
      Return Controller.TableLayout.TabPage
    End Function

    Public Function GetColPos(control As Control) As Integer
      Dim table As TableLayoutPanel = Controller.TableLayout.GetRecordTablePanelCallback()
      Return table.GetColumn(control)
    End Function

    Public Function AllowSort() As Boolean
      Return Controller.Record.CanSort
    End Function

    Public Function ReadSheetRecord() As SheetRecordManager
      SheetRecord = Controller.Record.ReadCallback()
      Return Me
    End Function

    Public Function LoadTable() As SheetRecordManager
      ShowTableTitles()

      Dim rec As SheetRecord = GetRecordThatAddSumRow()
      If rec IsNot Nothing Then
        Controller.LoadTableCallback(Controller.TableLayout.GetRecordTablePanelCallback(), GetRecordThatAddSumRow())
        ShowTableRecord()
      End If
      Return Me
    End Function

    Private Function GetRecordThatAddSumRow() As SheetRecord
      If SheetRecord IsNot Nothing Then
        Return Controller.Record.AddSumOfRecordCallback(SheetRecord)
      Else
        Return Nothing
      End If
    End Function

    Private Sub ShowTableTitles()
      Dim titles As TableLayoutPanel = Controller.TableLayout.GetTitlesTablePanelCallback()
      Dim offset As Point = Controller.TableLayout.ShowedTableOffsetLocation
      titles.Location = offset
      Dim subTitles As TableLayoutPanel = Controller.TableLayout.GetSubTitlesTablePanelCallback()
      subTitles.Location = New Point(offset.X, offset.Y + 26)

      Dim tabPage As TabPage = Controller.TableLayout.TabPage
      tabPage.Controls.Add(titles)
      tabPage.Controls.Add(subTitles)
    End Sub

    Private Sub ShowTableRecord()
      Dim offset As Point = Controller.TableLayout.ShowedTableOffsetLocation
      Dim scrollPanel As Panel = Controller.TableLayout.GetRecordTablePanelCallback().Parent
      scrollPanel.Location = New Point(offset.X, offset.Y + 57)
      scrollPanel.AutoScroll = True

      Dim tabPage As TabPage = Controller.TableLayout.TabPage
      tabPage.Controls.Add(scrollPanel)
    End Sub

    Public Function Sort(col As Integer) As SheetRecordManager
      If Controller.Record.CanSort Then
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
          Return Controller.Record.ReadCallback()
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
      Return Controller.CSV.GetCSVFileNameCallback()
    End Function

    Public Function ToCSV() As String
      Dim rec As SheetRecord = GetRecordThatAddSumRow()
      If rec IsNot Nothing Then
        Return _
          Controller.CSV.GetCSVTitlesCallback() & vbCrLf &
          Controller.CSV.GetCSVSubTitlesCallback() & vbCrLf &
          RecordConverter.ToCSV(rec)
      Else
        Return ""
      End If
    End Function

    Public Function Clone(tabPage As TabPage) As SheetRecordManager
      Dim c As TableRecordController = Controller
      c.TableLayout.TabPage = tabPage
      Return New SheetRecordManager(c)
    End Function
  End Class

End Class

