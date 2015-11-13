
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
    Dim FuncCreateTable As Action(Of TabPage, SheetRecord)
  End Structure

  Public Structure SortInfo
    Dim col As Integer
    Dim asc As Boolean
  End Structure

  Private AppProps As PropertyManager = AP.MANAGER

  Private UserRecordLoader As UserRecordLoader
  Private UserRecordManager As UserRecordManager
  Private UserInfoList As List(Of ExpandedUserInfo)
  Private ReadRecordTerm As ReadRecordTerm

  Private ExcelProps As PropertyManager = WAP.MANAGER
  Private Excel As ExcelAccessor

  Private InnerTabPageInfoListInPersonalTab As List(Of TabInfo)
  Private InnerTabPageInfoListInTotalTab As List(Of TabInfo)

  Private SortProp As SortInfo

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
      'Excel.Init()
      UserRecordManager = New UserRecordManager(UserInfoList, ReadRecordTerm)
      UserRecordLoader = New UserRecordLoader(Excel, UserRecordManager)

      InitInnerTabPageInPersonalTab()
      InitInnerTabPageInTotalTab()
      InitTableTitles()
      InitCBoxUserInfo()
      InitCBoxWeeklyTotal()
      InitCBoxMonthlyTotal()

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

    For Each ym As Tuple(Of Integer, Integer) In ReadRecordTerm.GetTermList.ToList
      Dim tab As TabInfo
      With tab
        .Name = UserRecordManager.GetSheetName(ym.Item2)
        .FuncCreateTable = GetActionForCreatingMonthlyTable(ym.Item1, ym.Item2)
      End With
      InnerTabPageInfoListInPersonalTab.Add(tab)
    Next

    Dim sumTab As TabInfo
    With sumTab
      .Name = UserRecordManager.GetSumSheetName
      .FuncCreateTable = AddressOf CreateSumTable
    End With
    InnerTabPageInfoListInPersonalTab.Add(sumTab)

    If tabInPersonalTab.TabPages.Count = InnerTabPageInfoListInPersonalTab.Count Then
      For idx As Integer = 0 To InnerTabPageInfoListInPersonalTab.Count - 1
        tabInPersonalTab.TabPages.Item(idx).Text = InnerTabPageInfoListInPersonalTab(idx).Name
      Next
    Else
      Throw New Exception(
        "Excelファイルのシート数とタブページの数が合いません。 / tabPageCount: " &
        tabInPersonalTab.TabPages.Count.ToString &
        " tabInfoCount: " & InnerTabPageInfoListInPersonalTab.Count)
    End If

  End Sub

  Private Sub InitInnerTabPageInTotalTab()
    InnerTabPageInfoListInTotalTab = New List(Of TabInfo)()
    Dim tabDays As TabInfo
    With tabDays
      .Name = "日"
      .FuncCreateTable = GetActionForCreatingDailyTotalTable()
    End With

    Dim tabWeeks As TabInfo
    With tabWeeks
      .Name = "週"
      .FuncCreateTable = GetActionForCreatingWeeklyTotalTable()
    End With

    Dim tabMonths As TabInfo
    With tabMonths
      .Name = "月"
      .FuncCreateTable = GetActionForCreatingMonthlyTotalTable()
    End With

    Dim tabYear As TabInfo
    With tabYear
      .Name = "合計"
      .FuncCreateTable = GetActionForCreatingAllTotalTable()
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

  Private Sub InitCBoxWeeklyTotal()
    Dim items As New List(Of DateItem)
    ReadRecordTerm.DateList.ForEach(
      Sub(t)
        For i As Integer = 1 To 5
          items.Add(New WeeklyItem(i, t.Item2, t.Item1))
        Next
      End Sub)

    items.ForEach(Sub(i) cboxWeeklyTotal.Items.Add(i))

    'Dim current As WeeklyItem = items.Find(Function(i) i.Agree(Date.Today))
    'If current IsNot Nothing Then
    '  Dim idx As Integer = cboxWeeklyTotal.Items.IndexOf(current)
    '  cboxWeeklyTotal.SelectedIndex = idx
    'End If
  End Sub

  Private Sub InitCBoxMonthlyTotal()
    Dim items As New List(Of MonthlyItem)
    ReadRecordTerm.DateList.ForEach(
      Sub(t) items.Add(New MonthlyItem(t.Item2, t.Item1)))

    items.ForEach(Sub(i) cboxMonthlyTotal.Items.Add(i))

    'Dim current As MonthlyItem = items.Find(Function(i) i.Agree(Date.Today))
    'If current IsNot Nothing Then
    '  Dim idx As Integer = cboxMonthlyTotal.Items.IndexOf(current)
    '  cboxMonthlyTotal.SelectedIndex = idx
    'End If
    'ReadRecordTerm.DateList.ForEach(
    '  Sub(t) cboxMonthlyTotal.Items.Add(New MonthlyItem(t.Item2, t.Item1)))
  End Sub

  Private Sub InitDPicDailyTotal()
    Dim min As Tuple(Of Integer, Integer) = ReadRecordTerm.DateList.First
    dPicDailyTotal.MinDate = New DateTime(min.Item1, min.Item2, 1, 0, 0, 0)
    Dim max As Tuple(Of Integer, Integer) = ReadRecordTerm.DateList.ToList.Last
    dPicDailyTotal.MaxDate = New DateTime(max.Item1, max.Item2, Date.DaysInMonth(max.Item1, max.Item2), 0, 0, 0)
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
    Else
      ShowTotalRecord()
    End If
  End Sub

  Private Sub tabInPersonalTab_PageChanged(sender As Object, e As EventArgs) Handles tabInPersonalTab.SelectedIndexChanged
    ShowPersonalRecord()
  End Sub

  Private Sub cboxUserInfo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxUserInfo.SelectedIndexChanged
    LoadTabPageInPersonalTab(tabInPersonalTab.SelectedTab, cboxUserInfo.SelectedItem)
  End Sub

  Private Sub tabInTotalTab_PageChanged(sender As Object, e As EventArgs) Handles tabInTotalTab.SelectedIndexChanged
    ShowTotalRecord()
  End Sub

  Private Sub datePickerDailyTotal_DateChanged(sender As Object, e As EventArgs) Handles dPicDailyTotal.ValueChanged
    If ReadRecordTerm.InTerm(dPicDailyTotal.Value.Month, dPicDailyTotal.Value.Year) Then
      ShowTotalRecord()
    End If
  End Sub

  Private Sub cboxWeeklyTotal_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxWeeklyTotal.SelectedIndexChanged
    If cboxWeeklyTotal.SelectedItem IsNot Nothing Then
      ShowTotalRecord()
    End If
  End Sub

  Private Sub cboxMonthlyTotal_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboxMonthlyTotal.SelectedIndexChanged
    If cboxMonthlyTotal.SelectedItem IsNot Nothing Then
      ShowTotalRecord()
    End If
  End Sub

  Private Sub ShowPersonalRecord()
    ShowPersonalTableTitles(tabInPersonalTab.SelectedTab)
    LoadTabPageInPersonalTab(tabInPersonalTab.SelectedTab, cboxUserInfo.SelectedItem)
  End Sub

  Private Sub ShowTotalRecord()
    ShowTotalTableTitles(tabInTotalTab.SelectedTab)
    LoadTabPageInTotalTab(tabInTotalTab.SelectedTab)
  End Sub

  Private Sub LoadTabPageInPersonalTab(selectedTabPage As TabPage, selectedUser As ExpandedUserInfo)
    If selectedUser IsNot Nothing Then
      Try
        If Not UserRecordManager.ContainsRecord(selectedUser.GetIdNum) Then
          UserRecordLoader.Load(selectedUser)
        End If
        Dim sheetRecord As SheetRecord = UserRecordManager.GetSheetRecord(selectedUser.GetIdNum, selectedTabPage.Text)

        InnerTabPageInfoListInPersonalTab.
          Find(Function(info) info.Name = selectedTabPage.Text).
          FuncCreateTable(selectedTabPage, sheetRecord)
      Catch ex As Exception
        MsgBox.ShowError(ex)
      End Try
    End If
  End Sub

  Private Sub LoadTabPageInTotalTab(selectedTabPage As TabPage)
    Try
      InnerTabPageInfoListInTotalTab.
        Find(Function(info)
               'MessageBox.Show(info.Name & " " & selectedTabPage.Text)
               Return info.Name = selectedTabPage.Text
             End Function).
        FuncCreateTable(selectedTabPage, Nothing)
    Catch ex As Exception
      MsgBox.ShowError(ex)
    End Try
  End Sub

  Private Sub CreateMonthlyTable(tabPage As TabPage, record As SheetRecord, year As Integer, month As Integer)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(RecordTableForm.pnlForTable).
        SetTextCols(0).
        SetNoteCols(RecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(GetFuncForGettingRowColorInMonthlyTable(year, month))

    CreateTable(RecordTableForm.tblRecord, tabPage, record, tableDrawer)
    ShowTableRecord(tabPage)
  End Sub

  Private Sub CreateSumTable(tabPage As TabPage, record As SheetRecord)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(RecordTableForm.pnlForTable).
        SetTextCols(0).
        SetNoteCols(RecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(AddressOf GetRowColorInSumTable)

    CreateTable(RecordTableForm.tblRecord, tabPage, record, tableDrawer)
    ShowTableRecord(tabPage)
  End Sub

  Private Sub CreateTotalTable(tabPage As TabPage, record As SheetRecord)
    Dim tableDrawer As TableDrawer =
      New TableDrawer(TotalRecordTableForm.pnlForTable).
        SetTextCols(0, 1).
        SetNoteCols(TotalRecordTableForm.tblRecord.ColumnCount - 1).
        SetFuncBackColor(Function(row) If(row = record.Count - 1, Color.PaleTurquoise, Color.Transparent))

    CreateTable(TotalRecordTableForm.tblRecord, tabPage, record, tableDrawer)
    ShowTotalTableRecord(tabPage)
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
            value = rec.First
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

  Private Function GetActionForCreatingMonthlyTable(year As Integer, month As Integer) As Action(Of TabPage, SheetRecord)
    Return _
      Sub(tabPage, record)
        CreateMonthlyTable(tabPage, record, year, month)
      End Sub
  End Function

  Private Function GetActionForCreatingDailyTotalTable() As Action(Of TabPage, SheetRecord)
    Return _
      Sub(tabPage, record)
        Dim y As Integer = dPicDailyTotal.Value.Year
        Dim m As Integer = dPicDailyTotal.Value.Month
        Dim d As Integer = dPicDailyTotal.Value.Day
        Dim r As SheetRecord = UserRecordManager.GetDailyTotalRecord(d, m, y)
        CreateTotalTable(tabPage, r)
      End Sub
  End Function

  Private Function GetActionForCreatingWeeklyTotalTable() As Action(Of TabPage, SheetRecord)
    Return _
      Sub(tabPage, record)
        Dim item As WeeklyItem = cboxWeeklyTotal.SelectedItem
        If item IsNot Nothing Then
          Dim w As Integer = item.Week
          Dim m As Integer = item.Month
          Dim y As Integer = item.Year
          Dim r As SheetRecord = UserRecordManager.GetWeeklyTotalRecord(w, m, y)
          CreateTotalTable(tabPage, r)
        End If
      End Sub
  End Function

  Private Function GetActionForCreatingMonthlyTotalTable() As Action(Of TabPage, SheetRecord)
    Return _
      Sub(tabPage, record)
        Dim item As MonthlyItem = cboxMonthlyTotal.SelectedItem
        If item IsNot Nothing Then
          Dim m As Integer = item.Month
          Dim y As Integer = item.Year
          Dim r As SheetRecord = UserRecordManager.GetMonthlyTotalRecord(m, y)
          CreateTotalTable(tabPage, r)
        End If
      End Sub
  End Function

  Private Function GetActionForCreatingAllTotalTable() As Action(Of TabPage, SheetRecord)
    Return _
      Sub(tabPage, record)
        Dim r As SheetRecord = UserRecordManager.GetAllTotalRecord()
        CreateTotalTable(tabPage, r)
      End Sub
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

