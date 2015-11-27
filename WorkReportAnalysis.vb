
'Option Strict On

Imports MP.Utils.Model

Imports MP.Utils.MyDate
Imports MP.Utils.MyCollection.Immutable
Imports RowRecord = MP.Utils.MyCollection.Immutable.MyLinkedList(Of String)
Imports SheetRecord = MP.Utils.MyCollection.Immutable.MyLinkedList(Of MP.Utils.MyCollection.Immutable.MyLinkedList(Of String))

Namespace WorkReportAnalysis
  Namespace App
    'Public Structure ItemTree
    '  Dim Key As String
    '  Dim DependedKeys As MyLinkedList(Of ItemTree)
    'End Structure

    Public Class WorkReportAnalysisProperties
      Private Shared SETTING_FILE_NAME As String = "excel.properties"

      Public Shared KEY_YEAR As String = "Year"
      Public Shared KEY_MAIN_FORM_NAME As String = "MainFormName"
      Public Shared KEY_EXCEL_FILE_DIR As String = "ExcelFileDir"
      Public Shared KEY_EXCEL_FILE_NAME_FORMAT As String = "ExcelFileNameFormat"
      Public Shared KEY_SHEET_NAME_FORMAT As String = "SheetNameFormat"
      Public Shared KEY_SUM_SHEET_NAME_FORMAT As String = "SheetNameFormat2"
      Public Shared KEY_FIRST_DAY_OF_A_MONTH_ROW As String = "FirstDayOfAMonthRow"
      Public Shared KEY_FIRST_MONTH_OF_SUM_SHEET_ROW As String = "FirstRow2"


      Public Shared ITEM_COUNT As Integer = 7
      Public Shared ITEM_DETAIL_COUNT As Integer = 3

      Public Shared KEY_ITEM_NAME1 As String = "ItemName1"
      Public Shared KEY_COL1_OF_ITEM1 As String = "Col1OfItem1"
      Public Shared KEY_COL2_OF_ITEM1 As String = "Col2OfItem1"
      Public Shared KEY_COL3_OF_ITEM1 As String = "Col3OfItem1"
      Public Shared KEY_ITEM_NAME2 As String = "ItemName2"
      Public Shared KEY_COL1_OF_ITEM2 As String = "Col1OfItem2"
      Public Shared KEY_COL2_OF_ITEM2 As String = "Col2OfItem2"
      Public Shared KEY_COL3_OF_ITEM2 As String = "Col3OfItem2"
      Public Shared KEY_ITEM_NAME3 As String = "ItemName3"
      Public Shared KEY_COL1_OF_ITEM3 As String = "Col1OfItem3"
      Public Shared KEY_COL2_OF_ITEM3 As String = "Col2OfItem3"
      Public Shared KEY_COL3_OF_ITEM3 As String = "Col3OfItem3"
      Public Shared KEY_ITEM_NAME4 As String = "ItemName4"
      Public Shared KEY_COL1_OF_ITEM4 As String = "Col1OfItem4"
      Public Shared KEY_COL2_OF_ITEM4 As String = "Col2OfItem4"
      Public Shared KEY_COL3_OF_ITEM4 As String = "Col3OfItem4"
      Public Shared KEY_ITEM_NAME5 As String = "ItemName5"
      Public Shared KEY_COL1_OF_ITEM5 As String = "Col1OfItem5"
      Public Shared KEY_COL2_OF_ITEM5 As String = "Col2OfItem5"
      Public Shared KEY_COL3_OF_ITEM5 As String = "Col3OfItem5"
      Public Shared KEY_ITEM_NAME6 As String = "ItemName6"
      Public Shared KEY_COL1_OF_ITEM6 As String = "Col1OfItem6"
      Public Shared KEY_COL2_OF_ITEM6 As String = "Col2OfItem6"
      Public Shared KEY_COL3_OF_ITEM6 As String = "Col3OfItem6"
      Public Shared KEY_ITEM_NAME7 As String = "ItemName7"
      Public Shared KEY_COL1_OF_ITEM7 As String = "Col1OfItem7"
      Public Shared KEY_COL2_OF_ITEM7 As String = "Col2OfItem7"
      Public Shared KEY_COL3_OF_ITEM7 As String = "Col3OfItem7"

      Public Shared KEY_NOTE_COL As String = "NoteCol"
      Public Shared KEY_WORKDAY_COL As String = "WorkdayCol"

      Public Shared MANAGER As New MP.Utils.Common.PropertyManager(SETTING_FILE_NAME, DefaultSettingProperties(), True)

      Public Shared Function AllItemKeys() As String()
        Dim keys As String() = New String() {
          KEY_COL1_OF_ITEM1,
          KEY_COL2_OF_ITEM1,
          KEY_COL3_OF_ITEM1,
          KEY_COL1_OF_ITEM2,
          KEY_COL2_OF_ITEM2,
          KEY_COL3_OF_ITEM2,
          KEY_COL1_OF_ITEM3,
          KEY_COL2_OF_ITEM3,
          KEY_COL3_OF_ITEM3,
          KEY_COL1_OF_ITEM4,
          KEY_COL2_OF_ITEM4,
          KEY_COL3_OF_ITEM4,
          KEY_COL1_OF_ITEM5,
          KEY_COL2_OF_ITEM5,
          KEY_COL3_OF_ITEM5,
          KEY_COL1_OF_ITEM6,
          KEY_COL2_OF_ITEM6,
          KEY_COL3_OF_ITEM6,
          KEY_COL1_OF_ITEM7,
          KEY_COL2_OF_ITEM7,
          KEY_COL3_OF_ITEM7,
          KEY_NOTE_COL
        }
        Return keys
      End Function

      Public Shared Function ItemKeys(itemNum As Integer) As String()
        If itemNum < 0 OrElse itemNum >= ITEM_COUNT Then
          Throw New Exception("指定された番号がアイテムの範囲を越えています。 itemNum: " & itemNum)
        End If

        Return AllItemKeys.Skip(itemNum * ITEM_DETAIL_COUNT).Take(ITEM_DETAIL_COUNT).ToArray
      End Function

      Public Shared Function ItemKeysList() As String()()
        Dim l As New List(Of String())
        For num As Integer = 0 To ITEM_COUNT - 1
          Dim i As String() = ItemKeys(num)
          l.Add(ItemKeys(num))
        Next
        Return l.ToArray
      End Function

      'Public Shared Function ItemTree() As ItemTree
      '  Return CreateLinkedKeys(KEY_WORKDAY_COL, ItemKeysList())
      'End Function

      'Private Shared Function CreateLinkedKeys(parentKey As String, childs As String()()) As ItemTree
      '  Dim f As Func(Of String()(), MyLinkedList(Of ItemTree)) =
      '    Function(c)
      '      If c.Count = 0 Then
      '        Return MyLinkedList(Of ItemTree).Nil
      '      Else
      '        Return f(c.Skip(1)).AddFirst(CreateChainKeys(c.First))
      '      End If
      '    End Function

      '  Dim tree As ItemTree
      '  With tree
      '    .Key = parentKey
      '    .DependedKeys = f(childs)
      '  End With
      '  Return tree
      'End Function

      'Private Shared Function CreateChainKeys(ParamArray keys As String()) As ItemTree
      '  Dim f As Func(Of String(), ItemTree) =
      '    Function(k)
      '      Dim n As MyLinkedList(Of ItemTree) = MyLinkedList(Of ItemTree).Nil
      '      If k.Count > 0 Then
      '        n = n.AddFirst(f(k.Skip(1)))
      '      End If

      '      Dim t As ItemTree
      '      With t
      '        .Key = k.First
      '        .DependedKeys = n
      '      End With

      '      Return t
      '    End Function
      '  Return f(keys)
      'End Function

      Private Shared Function DefaultSettingProperties() As IDictionary(Of String, String)
        Dim dict As IDictionary(Of String, String) = New Dictionary(Of String, String)
        dict(KEY_YEAR) = "2015"
        dict(KEY_MAIN_FORM_NAME) = "作業件数集計"
        dict(KEY_EXCEL_FILE_DIR) = MP.Details.Sys.App.GetCurrentDirectory()
        dict(KEY_EXCEL_FILE_NAME_FORMAT) = "件数報告書-{0}.xls"
        dict(KEY_SHEET_NAME_FORMAT) = "{0}月分"
        dict(KEY_SUM_SHEET_NAME_FORMAT) = "集計"

        dict(KEY_FIRST_DAY_OF_A_MONTH_ROW) = "7"
        dict(KEY_FIRST_MONTH_OF_SUM_SHEET_ROW) = "12"

        dict(KEY_ITEM_NAME1) = "郵政"
        dict(KEY_COL1_OF_ITEM1) = "C"
        dict(KEY_COL2_OF_ITEM1) = "D"
        dict(KEY_COL3_OF_ITEM1) = "Q"
        dict(KEY_ITEM_NAME2) = "NB"
        dict(KEY_COL1_OF_ITEM2) = "E"
        dict(KEY_COL2_OF_ITEM2) = "F"
        dict(KEY_COL3_OF_ITEM2) = "R"
        dict(KEY_ITEM_NAME3) = "新規入力"
        dict(KEY_COL1_OF_ITEM3) = "G"
        dict(KEY_COL2_OF_ITEM3) = "H"
        dict(KEY_COL3_OF_ITEM3) = "S"
        dict(KEY_ITEM_NAME4) = "郵政写真"
        dict(KEY_COL1_OF_ITEM4) = "I"
        dict(KEY_COL2_OF_ITEM4) = "J"
        dict(KEY_COL3_OF_ITEM4) = "T"
        dict(KEY_ITEM_NAME5) = "NB写真"
        dict(KEY_COL1_OF_ITEM5) = "K"
        dict(KEY_COL2_OF_ITEM5) = "L"
        dict(KEY_COL3_OF_ITEM5) = "U"
        dict(KEY_ITEM_NAME6) = "校正"
        dict(KEY_COL1_OF_ITEM6) = "M"
        dict(KEY_COL2_OF_ITEM6) = "N"
        dict(KEY_COL3_OF_ITEM6) = "V"
        dict(KEY_ITEM_NAME7) = "作字"
        dict(KEY_COL1_OF_ITEM7) = "O"
        dict(KEY_COL2_OF_ITEM7) = "P"
        dict(KEY_COL3_OF_ITEM7) = "W"

        dict(KEY_NOTE_COL) = "X"
        dict(KEY_WORKDAY_COL) = "Y"

        Return dict
      End Function

    End Class

    Public Class FileFormat
      Private Shared m As Utils.Common.PropertyManager = App.WorkReportAnalysisProperties.MANAGER

      Public Shared Function GetFilePath(userIdNum As String) As String
        Return GetFileDir() & "\" & GetFileName(userIdNum)
      End Function

      Public Shared Function GetFileDir() As String
        Return m.GetValue(App.WorkReportAnalysisProperties.KEY_EXCEL_FILE_DIR)
      End Function

      Public Shared Function GetFileName(userIdNum As String) As String
        Dim format As String = m.GetValue(App.WorkReportAnalysisProperties.KEY_EXCEL_FILE_NAME_FORMAT)
        Return String.Format(format, userIdNum)
      End Function

      Public Shared Function GetSheetName(month As Integer) As String
        Dim format As String = m.GetValue(App.WorkReportAnalysisProperties.KEY_SHEET_NAME_FORMAT)
        Return String.Format(format, month)
      End Function

      Public Shared Function GetSheetName2() As String
        Return m.GetValue(App.WorkReportAnalysisProperties.KEY_SUM_SHEET_NAME_FORMAT)
      End Function

      Public Shared Function GetFirstDayOfAMonthRow() As Integer
        Dim row As String = m.GetValue(App.WorkReportAnalysisProperties.KEY_FIRST_DAY_OF_A_MONTH_ROW)
        If MP.Utils.General.MyChar.IsInteger(row) Then
          Return Integer.Parse(row)
        Else
          Throw New Exception("プロパティ<" & App.WorkReportAnalysisProperties.KEY_FIRST_DAY_OF_A_MONTH_ROW & ">の値が不正です。")
        End If
      End Function

      Public Shared Function GetFirstRow() As Integer
        Return GetIntger(App.WorkReportAnalysisProperties.KEY_FIRST_DAY_OF_A_MONTH_ROW)
      End Function

      Public Shared Function GetFirstRow2() As Integer
        Return GetIntger(App.WorkReportAnalysisProperties.KEY_FIRST_MONTH_OF_SUM_SHEET_ROW)
      End Function

      Private Shared Function GetIntger(key As String) As Integer
        Dim value As String = m.GetValue(key)
        If MP.Utils.General.MyChar.IsInteger(value) Then
          Return Integer.Parse(value)
        Else
          Throw New Exception("プロパティ<" & key & ">の値が不正です。")
        End If
      End Function

      Public Shared Function GetItemCols(ParamArray excludedItemKeys As String()) As List(Of String)
        Return App.WorkReportAnalysisProperties.AllItemKeys().ToList.
          FindAll(Function(k) Not excludedItemKeys.Contains(k)).
          ConvertAll(
            Function(k)
              Dim col As String = m.GetValue(k)
              Return If(Char.IsLetter(CType(col, Char)), col, "")
            End Function)
      End Function

      Public Shared Function GetItemColsList(ParamArray excludedItemKeys As String()) As List(Of List(Of String))
        Return _
          App.WorkReportAnalysisProperties.ItemKeysList.ToList.
            ConvertAll(
              Function(keys)
                Dim l As New List(Of String)
                For Each k As String In keys
                  Dim col As String = m.GetValue(k)
                  l.Add(If(Char.IsLetter(CType(col, Char)) AndAlso Not excludedItemKeys.Contains(col), col, ""))
                Next
                Return l
              End Function)
      End Function

      Public Shared Function GetYear() As Integer
        Return GetIntger(App.WorkReportAnalysisProperties.KEY_YEAR)
      End Function
    End Class
  End Namespace

  Namespace Control
    Public Class UserRecordLoader
      Private _UserRecordManager As UserRecordManager
      Public ReadOnly Property UserRecordManager As UserRecordManager
        Get
          Return _UserRecordManager
        End Get
      End Property

      Private Excel As Excel.ExcelAccessor
      Private AccessPropList As List(Of Excel.AccessProperties)
      Private AccessPropListWithTree As List(Of Excel.AccessPropertiesWithTree)

      Private Shared m As Utils.Common.PropertyManager = App.WorkReportAnalysisProperties.MANAGER

      Public Sub New(excel As Excel.ExcelAccessor, manager As UserRecordManager)
        _UserRecordManager = manager
        Me.Excel = excel
        AccessPropList = New List(Of Excel.AccessProperties)
        AccessPropListWithTree = New List(Of Excel.AccessPropertiesWithTree)

        Init()
      End Sub

      Private Sub Init()
        AccessPropList.Clear()

        Dim term As MyLinkedList(Of MonthlyItem) =
          _UserRecordManager.GetReadRecordTerm.MonthlyItemList

        term.ForEach(
          Sub(e)
            Dim p As Excel.AccessProperties
            With p
              .RecordKey = UserRecordManager.GetSheetName(e.Month)
              .SheetName = UserRecordManager.GetSheetName(e.Month)
              .Cols = App.FileFormat.GetItemCols()
              .FirstRow = App.FileFormat.GetFirstDayOfAMonthRow()
              .RowSize = Date.DaysInMonth(e.Year, e.Month) '+ 1
            End With
            AccessPropList.Add(p)
          End Sub)

        Dim colTree As Excel.ColTree =
          WorkReportAnalysis.Excel.ExcelAccessor.CreateTreeList(
            m.GetValue(App.WorkReportAnalysisProperties.KEY_WORKDAY_COL),
            False,
            App.FileFormat.GetItemColsList)

        term.ForEach(
          Sub(e)
            Dim format As Excel.SheetFormat
            With format
              .OffsetRow = App.FileFormat.GetFirstDayOfAMonthRow
              .RowSize = Date.DaysInMonth(e.Year, e.Month)
              .ColTree = colTree
            End With

            Dim p As Excel.AccessPropertiesWithTree
            With p
              .RecordKey = UserRecordManager.GetSheetName(e.Month)
              .SheetName = UserRecordManager.GetSheetName(e.Month)
              .SheetFormat = format
            End With
            AccessPropListWithTree.Add(p)
          End Sub)

        'Dim sum As Excel.AccessProperties
        'With sum
        '  .RecordKey = UserRecordManager.GetSumSheetName
        '  .SheetName = UserRecordManager.GetSumSheetName
        '  .Cols = App.FileFormat.GetItemCols(App.WorkReportAnalysisProperties.KEY_NOTE_COL)
        '  .FirstRow = App.FileFormat.GetFirstRow2()
        '  .RowSize = 6 * term.Count '+ 1
        'End With
        'AccessPropList.Add(sum)
      End Sub

      Public Sub LoadUserRecord()

        For Each info As Model.ExpandedUserInfo In _UserRecordManager.GetUserInfoList
          Dim fileName As String = ""
          Try
            If Not _UserRecordManager.ContainsRecord(info.GetIdNum) Then
              fileName = App.FileFormat.GetFilePath(info.GetIdNum())
              Load(info)
            End If
          Catch ex As Exception
            Dim res As DialogResult =
              MessageBox.Show(
              "ファイル:" & fileName & "の読み込みに失敗しました。 / id: " & vbCrLf &
              "読み込みを続けますか？" & vbCrLf & vbCrLf & ex.Message,
              "Error!", MessageBoxButtons.YesNo, MessageBoxIcon.Error)
            If res = DialogResult.No Then
              Exit For
            End If
          End Try
        Next

      End Sub

      Public Sub Load(userInfo As Model.ExpandedUserInfo)
        If Not _UserRecordManager.ContainsRecord(userInfo.GetIdNum) Then
          Dim fileName As String = App.FileFormat.GetFilePath(userInfo.GetIdNum())
          Dim userRecord As New Model.UserRecord(userInfo)

          'Excel.Read(fileName, AccessPropList).
          'ForEach(Sub(res) userRecord.Add(res.SheetName, res))
          Excel.ReadWithTree(fileName, AccessPropListWithTree).
            ForEach(Sub(res) userRecord.Add(res.SheetName, res))
          SyncLock _UserRecordManager
            _UserRecordManager.Add(userInfo.GetIdNum, userRecord)
          End SyncLock
        End If
      End Sub

    End Class

    Public Class UserRecordManager
      Private UserInfoList As List(Of Model.ExpandedUserInfo)
      Private UserRecordMap As IDictionary(Of String, Model.UserRecord)
      Private RecordTerm As ReadRecordTerm

      Public Shared Function GetSheetName(month As Integer) As String
        Return App.FileFormat.GetSheetName(month)
      End Function

      Public Shared Function GetSumSheetName() As String
        Return App.FileFormat.GetSheetName2
      End Function

      Public Sub New(userInfoList As List(Of Model.ExpandedUserInfo), term As ReadRecordTerm)
        Me.UserInfoList = userInfoList
        UserRecordMap = New Dictionary(Of String, Model.UserRecord)
        RecordTerm = term
      End Sub

      Public Function GetReadRecordTerm() As ReadRecordTerm
        Return RecordTerm
      End Function

      Public Function GetUserInfoList() As List(Of Model.ExpandedUserInfo)
        Return UserInfoList
      End Function

      Public Sub Add(id As String, record As Model.UserRecord)
        If UserInfoList.Find(Function(info) info.GetIdNum = id) IsNot Nothing Then
          UserRecordMap.Add(id, record)
        End If
      End Sub

      Public Function ContainsRecord(id As String) As Boolean
        Return UserRecordMap.ContainsKey(id)
      End Function

      Public Function GetUserRecord(id As String) As Model.UserRecord
        If UserRecordMap.ContainsKey(id) Then
          Return UserRecordMap(id)
        Else
          Throw New Exception("指定したIDのデータはありません。 id: " & id)
        End If
      End Function

      Public Function GetSheetRecord(id As String, key As String, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        If key = GetSumSheetName() Then
          Return GetUserSumRecord(id, filter)
        Else
          Dim m As MonthlyItem =
            RecordTerm.MonthlyItemList.Find(Function(e) key = GetSheetName(e.Month))
          If m IsNot Nothing Then
            Return GetUserDailyRecordInMonth(id, m.Month, m.Year)
          Else
            Throw New Exception("指定したキーのデータはありません。 key: " & key)
          End If
        End If
      End Function

      Public Function GetUserDailyRecordInMonth(id As String, month As Integer, year As Integer) As SheetRecord
        Dim record As SheetRecord = GetUserDailyRecord(id, 1, Date.DaysInMonth(year, month), month, year)
        Return PadEmptyRowRecord(record, RecordTableForm.tblRecord.RowCount - 1)
      End Function

      Public Function GetUserDailyRecordInWeek(id As String, week As Integer, month As Integer, year As Integer) As SheetRecord
        If Not CheckValidDate(1, week, month, year) Then
          Return SheetRecord.Nil
        End If

        Dim days As List(Of Integer) = MyCalendar.GetDaysInWeek(year, month, week)
        Dim weekRec As SheetRecord = GetUserDailyRecord(id, days.First, days.Count, month, year)

        If days.Count < 7 Then
          Dim restDays As Integer = 7 - days.Count
          If week = 1 Then
            Dim lastMonth As Integer = If(month = 1, 12, month - 1)
            Dim lastYear As Integer = If(month = 1, year - 1, year)
            Dim dayOfSunday As Integer = Date.DaysInMonth(lastYear, lastMonth) - (restDays - 1)
            Dim lastRec As SheetRecord = GetUserDailyRecord(id, dayOfSunday, restDays, lastMonth, lastYear)
            weekRec = weekRec.AddRangeToHead(lastRec)
          Else
            Dim nextMonth As Integer = If(month = 12, 1, month + 1)
            Dim nextYear As Integer = If(month = 12, year + 1, year)
            Dim nextRec As SheetRecord = GetUserDailyRecord(id, 1, restDays, nextMonth, nextYear)
            weekRec = nextRec.AddRangeToHead(weekRec)
          End If
        End If

        Return weekRec
      End Function

      Private Function GetUserDailyRecord(id As String, startDay As Integer, count As Integer, month As Integer, year As Integer) As SheetRecord
        If Not CheckValidDate(1, 1, month, year) Then
          Return SheetRecord.Nil()
        End If

        Dim endDay As Integer = startDay + count - 1
        Dim sheetName As String = GetSheetName(month)
        Dim record As SheetRecord = GetUserRecord(id).GetSheetRecord(sheetName)
        Dim go As Func(Of Integer, SheetRecord, SheetRecord) =
          Function(idx, rec)
            If idx < endDay Then
              Dim rr As RowRecord = rec.First.AddFirst((idx + 1).ToString & "日")
              Return go(idx + 1, rec.Rest).AddFirst(rr)
            Else
              Return SheetRecord.Nil
            End If
          End Function
        Return go(startDay - 1, record.Skip(startDay - 1))
      End Function

      Public Function GetUserWeeklyRecordAll(id As String, year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        Dim go As Func(Of MyLinkedList(Of WeeklyItem), SheetRecord) =
          Function(term)
            If term.Empty Then
              Return SheetRecord.Nil
            Else
              Dim w As WeeklyItem = term.First
              Dim rr As RowRecord = GetUserWeeklyRecord(id, w, filter)
              Return go(term.Rest).AddFirst(rr)
            End If
          End Function
        Return go(RecordTerm.WeeklyItemList)
      End Function

      Public Function GetUserWeeklyRecordInMonth(id As String, month As Integer, year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        Dim weekCnt As Integer = MyCalendar.GetLastWeekNum(year, month)
        Dim go As Func(Of MyLinkedList(Of WeeklyItem), SheetRecord) =
          Function(week)
            If week.Empty OrElse week.First.Year < year OrElse week.First.Month < month Then
              Return SheetRecord.Nil
            Else
              Dim w As WeeklyItem = week.First
              Dim rr As RowRecord = GetUserWeeklyRecord(id, w, filter)
              Return go(week.Rest).AddFirst(rr)
            End If
          End Function
        Return go(RecordTerm.WeeklyItemList)
      End Function

      Private Function GetUserWeeklyRecord(id As String, w As WeeklyItem, filter As Func(Of RowRecord, Integer, Boolean)) As RowRecord
        Dim rec As SheetRecord = GetUserDailyRecordInWeek(id, w.Week, w.Month, w.Year)
        Return RecordConverter.CreateSumRowRecord(rec, 1, filter).AddFirst(w.ToString)
      End Function

      Public Function GetUserMonthlyRecordAll(id As String, year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        Dim go As Func(Of MyLinkedList(Of MonthlyItem), SheetRecord) =
          Function(term)
            If term.Empty Then
              Return SheetRecord.Nil
            Else
              Dim ym As MonthlyItem = term.First
              Dim rr As RowRecord = GetUserMonthlyRecord(id, ym, filter)
              Return go(term.Rest).AddFirst(rr)
            End If
          End Function
        Return go(RecordTerm.MonthlyItemList)
      End Function

      Private Function GetUserMonthlyRecord(id As String, m As MonthlyItem, filter As Func(Of RowRecord, Integer, Boolean)) As RowRecord
        Dim rec As SheetRecord = GetUserDailyRecordInMonth(id, m.Month, m.Year)
        Return RecordConverter.CreateSumRowRecord(rec, 1, filter).AddFirst(m.ToString)
      End Function

      Public Function GetUserTotalRecord(id As String, year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As RowRecord
        Dim rec As SheetRecord = GetUserMonthlyRecordAll(id, year, filter)
        Return RecordConverter.CreateSumRowRecord(rec, 1, filter).AddFirst(year & "年")
      End Function

      Public Function GetUserSumRecord(id As String, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        Dim weekly As SheetRecord =
          GetUserWeeklyRecordAll(id, RecordTerm.StartYear, filter).
            ConvertAll(Function(rr) rr.AddFirst("")).
            AddFirst(RowRecord.Nil.AddFirst("週計"))
        Dim monthly As SheetRecord =
          GetUserMonthlyRecordAll(id, RecordTerm.StartYear, filter).
            ConvertAll(Function(rr) rr.AddFirst("")).
            AddFirst(RowRecord.Nil.AddFirst("月計"))
        Dim total As SheetRecord =
          SheetRecord.Nil.
          AddFirst(GetUserTotalRecord(id, RecordTerm.StartYear, filter).AddFirst("")).
          AddFirst(RowRecord.Nil.AddFirst("合計"))

        Return total.AddRangeToHead(monthly).AddRangeToHead(weekly)
      End Function

      Public Function GetDailyTermRecord(day As Integer, month As Integer, year As Integer) As SheetRecord
        If Not CheckValidDate(day, 1, month, year) Then
          Return SheetRecord.Nil
        End If

        Return GetAllUserRecord(Function(id) GetUserDailyRecord(id, day, 1, month, year).First.Rest)
      End Function

      Public Function GetWeeklyTermRecord(week As Integer, month As Integer, year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        If Not CheckValidDate(1, week, month, year) Then
          Return SheetRecord.Nil
        End If

        Return GetAllUserRecord(Function(id) GetUserWeeklyRecord(id, New WeeklyItem(week, month, year, RecordTerm), filter).Rest)
      End Function

      Public Function GetMonthlyTermRecord(month As Integer, year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        If Not CheckValidDate(1, 1, month, year) Then
          Return SheetRecord.Nil
        End If

        Return GetAllUserRecord(Function(id) GetUserMonthlyRecord(id, New MonthlyItem(month, year), filter).Rest)
      End Function

      Public Function GetAllTermRecord(year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        Return GetAllUserRecord(Function(id) GetUserTotalRecord(id, year, filter).Rest)
      End Function

      Private Function GetAllUserRecord(getRowRec As Func(Of String, RowRecord)) As SheetRecord
        Dim record As List(Of RowRecord) =
          UserInfoList.
            ConvertAll(
              Function(info)
                If UserRecordMap.ContainsKey(info.GetIdNum) Then
                  Dim rr As RowRecord = getRowRec(info.GetIdNum)
                  Return rr.AddFirst(info.GetName).AddFirst(info.GetIdNum)
                Else
                  Return RowRecord.Nil
                End If
              End Function)

        Return SheetRecord.Nil.AddRangeToHead(record)
      End Function

      'Public Function GetAllUserRecord() As MyLinkedList(Of Model.UserRecord)
      '  Dim go As Func(Of List(Of Model.UserRecord), MyLinkedList(Of Model.UserRecord)) =
      '    Function(values)
      '      If values.Count = 0 Then
      '        Return MyLinkedList(Of Model.UserRecord).Nil()
      '      Else
      '        Return go(values.Skip(1)).AddFirst(values.First)
      '      End If
      '    End Function
      '  Return go(UserRecordMap.Values)
      'End Function

      Public Function GetDailyTotalRecord(month As Integer, year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        If Not CheckValidDate(1, 1, month, year) Then
          Return SheetRecord.Nil
        End If

        Dim sheet As New List(Of RowRecord)
        For d As Integer = 1 To Date.DaysInMonth(year, month)
          Dim rec As SheetRecord = GetDailyTermRecord(d, month, year)
          Dim rr As RowRecord = RecordConverter.CreateSumRowRecord(rec, 2, filter).AddFirst(d.ToString & "日").AddFirst("")
          sheet.Add(rr)
        Next
        Return PadEmptyRowRecord(SheetRecord.Nil.AddRangeToHead(sheet), TotalRecordTableForm.tblRecord.RowCount - 1)
      End Function

      Public Function GetWeeklyTotalRecord(year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        Dim sheet As New List(Of RowRecord)
        RecordTerm.WeeklyItemList.
          ForEach(
            Sub(w)
              Dim rec As SheetRecord = GetWeeklyTermRecord(w.Week, w.Month, w.Year, filter)
              Dim rr As RowRecord = RecordConverter.CreateSumRowRecord(rec, 2, filter).AddFirst(w.ToString).AddFirst("")
              sheet.Add(rr)
            End Sub)
        Return SheetRecord.Nil.AddRangeToHead(sheet)
      End Function

      Public Function GetMonthlyTotalRecord(year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        Dim sheet As New List(Of RowRecord)
        RecordTerm.MonthlyItemList.
          ForEach(
            Sub(m)
              Dim rec As SheetRecord = GetMonthlyTermRecord(m.Month, m.Year, filter)
              Dim rr As RowRecord = RecordConverter.CreateSumRowRecord(rec, 2, filter).AddFirst(m.ToString).AddFirst("")
              sheet.Add(rr)
            End Sub)

        Return SheetRecord.Nil.AddRangeToHead(sheet)
      End Function

      Private Function PadEmptyRowRecord(record As SheetRecord, allcnt As Integer) As SheetRecord
        If record.Count < allcnt Then
          Return PadEmptyRowRecord(record.AddLast(RowRecord.Nil), allcnt)
        Else
          Return record
        End If
      End Function

      Private Function CheckValidDate(day As Integer, week As Integer, month As Integer, year As Integer) As Boolean
        If InvalidMonth(month) Then
          Throw New Exception("月の値が不正です。 month:  " & month)
        ElseIf InvalidWeek(week, month, year) Then
          Throw New Exception("週の値が不正です。 week: " & week)
        ElseIf InvalidDay(day, month, year) Then
          Throw New Exception("日の値が不正です。 day: " & day & " month: " & month & " year: " & year)
        ElseIf OutOfTerm(month, year) Then
          Return False
        Else
          Return True
        End If
      End Function

      Private Function OutOfTerm(month As Integer, year As Integer) As Boolean
        Return Not RecordTerm.MonthlyItemList.Exists(Function(t) t.Month = month AndAlso t.Year = year)
      End Function

      Private Function InvalidMonth(month As Integer) As Boolean
        Return month < 1 OrElse month > 12
      End Function

      Private Function InvalidWeek(week As Integer, month As Integer, year As Integer) As Boolean
        Return week < 1 OrElse week > Utils.MyDate.MyCalendar.GetLastWeekNum(year, month)
      End Function

      Private Function InvalidDay(day As Integer, month As Integer, year As Integer) As Boolean
        Return day < 1 OrElse day > Date.DaysInMonth(year, month)
      End Function

      Private Function ToImmutableListFrom(Of T)(list As List(Of T)) As MyLinkedList(Of T)
        Dim il As MyLinkedList(Of T) = MyLinkedList(Of T).Nil
        Dim go As Func(Of Integer, MyLinkedList(Of T)) =
          Function(idx)
            If idx >= list.Count Then
              Return MyLinkedList(Of T).Nil
            Else
              Return go(idx + 1).AddFirst(list(idx))
            End If
          End Function
        Return go(0)
      End Function

    End Class

    Public Class ReadRecordTerm
      Private _StartMonth As Integer
      Public ReadOnly Property StartMonth() As Integer
        Get
          Return _StartMonth
        End Get
      End Property

      Private _StartYear As Integer
      Public ReadOnly Property StartYear() As Integer
        Get
          Return _StartYear
        End Get
      End Property

      Private _EndMonth As Integer
      Public ReadOnly Property EndMonth() As Integer
        Get
          Return _EndMonth
        End Get
      End Property

      Private _EndYear As Integer
      Public ReadOnly Property EndYear() As Integer
        Get
          Return _EndYear
        End Get
      End Property

      Private _MonthlyItemList As MyLinkedList(Of MonthlyItem)
      Public ReadOnly Property MonthlyItemList As MyLinkedList(Of MonthlyItem)
        Get
          Return _MonthlyItemList
        End Get
      End Property

      Private _WeeklyItemList As MyLinkedList(Of WeeklyItem)
      Public ReadOnly Property WeeklyItemList As MyLinkedList(Of WeeklyItem)
        Get
          Return _WeeklyItemList
        End Get
      End Property

      Public Sub New(startMonth As Integer, startYear As Integer, endMonth As Integer, endYear As Integer)
        If IsValidTerm(startMonth, startYear, endMonth, endYear) Then
          Me._StartMonth = startMonth
          Me._StartYear = startYear
          Me._EndMonth = endMonth
          Me._EndYear = endYear
          _MonthlyItemList = GetMonthlyTerm(startMonth, startYear, endMonth, endYear)
          _WeeklyItemList = GetWeeklyTerm(1, startMonth, startYear, endMonth, endYear)
        Else
          Throw New Exception("期間の値が不正です。")
        End If
      End Sub

      Public Function InMonthlyTerm(month As Integer, year As Integer) As Boolean
        Return AfterStartDate(month, year) AndAlso BeforeEndDate(month, year)
      End Function

      Private Function AfterStartDate(month As Integer, year As Integer) As Boolean
        Return _
          year = StartYear AndAlso month >= StartMonth OrElse
          year > StartYear
      End Function

      Private Function BeforeEndDate(month As Integer, year As Integer) As Boolean
        Return _
          year = EndYear AndAlso month <= EndMonth OrElse
          year < EndYear
      End Function

      Private Function GetMonthlyTerm(startMonth As Integer, startYear As Integer, endMonth As Integer, endYear As Integer) As MyLinkedList(Of MonthlyItem)
        If startYear = endYear AndAlso startMonth = endMonth Then
          Return New MyLinkedList(Of MonthlyItem)(New MonthlyItem(startMonth, startYear))
        Else
          Dim l As MyLinkedList(Of MonthlyItem) =
            If(startMonth < 12,
               GetMonthlyTerm(startMonth + 1, startYear, endMonth, endYear),
               GetMonthlyTerm(1, startYear + 1, endMonth, endYear))
          Return l.AddFirst(New MonthlyItem(startMonth, startYear))
        End If
      End Function

      Private Function GetWeeklyTerm(week As Integer, startMonth As Integer, startYear As Integer, endMonth As Integer, endYear As Integer) As MyLinkedList(Of WeeklyItem)
        If startYear = endYear AndAlso startMonth = endMonth AndAlso week = MyCalendar.GetLastWeekNum(startYear, startMonth) Then
          Return New MyLinkedList(Of WeeklyItem)(New WeeklyItem(week, startMonth, startYear, Me))
        Else
          Dim l As MyLinkedList(Of WeeklyItem) =
            If(week < MyCalendar.GetLastWeekNum(startYear, startMonth),
              GetWeeklyTerm(week + 1, startMonth, startYear, endMonth, endYear),
              If(startMonth < 12,
                GetWeeklyTerm(1, startMonth + 1, startYear, endMonth, endYear),
                GetWeeklyTerm(1, 1, startYear + 1, endMonth, endYear)))
          If week = 1 AndAlso (startMonth <> Me.StartMonth OrElse startYear <> Me.StartYear) AndAlso MyCalendar.GetWeek(startYear, startMonth, 1) <> 1 Then
            Return l
          Else
            Return l.AddFirst(New WeeklyItem(week, startMonth, startYear, Me))
          End If
        End If
      End Function

      Private Function IsValidTerm(startMonth As Integer, startYear As Integer, endMonth As Integer, endYear As Integer) As Boolean
        Return _
          (startMonth >= 1 AndAlso startMonth <= 12 AndAlso endMonth >= 1 AndAlso endMonth <= 12) AndAlso
          startYear < endYear OrElse (startYear = endYear AndAlso startMonth <= endMonth)
      End Function

    End Class

    Public Class WeeklyItem
      Inherits DateItem

      Private _Week As Integer
      Public ReadOnly Property Week() As Integer
        Get
          Return _Week
        End Get
      End Property

      Private IsWeekUnder7Days As Boolean
      Private RecordTerm As ReadRecordTerm

      'Public Shared Function GetWeekNumInMonth(d As Integer, m As Integer, y As Integer) As Integer
      '  Return GetWeekNumInMonth(New Date(y, m, d))
      'End Function

      'Public Shared Function GetWeekNumInMonth(day As Date) As Integer
      '  Dim first As Date = MonthlyItem.GetFirstDateInMonth(day.Month, day.Year)
      '  Return DatePart("WW", day) - DatePart("ww", first) + 1
      'End Function

      Public Sub New(w As Integer, m As Integer, y As Integer, recordTerm As ReadRecordTerm)
        MyBase.New(w, m, y)
        _Week = MyBase.Value
        IsWeekUnder7Days =
        (w = 1 AndAlso MyCalendar.GetWeek(y, m, 1) <> 1) OrElse
        (w = MyCalendar.GetLastWeekNum(y, m) AndAlso MyCalendar.GetWeek(y, m, Date.DaysInMonth(y, m)) <> 7)
        Me.RecordTerm = recordTerm
      End Sub

      Public Overrides Function Agree(day As Date) As Boolean
        If day.Year = Year() AndAlso day.Month = Month() Then
          Dim w As Integer = MyCalendar.GetWeekNumInMonth(day)
          Return w = Week()
        Else
          Return False
        End If
      End Function

      Public Function IsWeekUnderSevenDays() As Boolean
        Return IsWeekUnder7Days
      End Function

      Public Overrides Function ToString() As String
        If IsWeekUnder7Days Then
          If Week() = 1 Then
            If (Month <> RecordTerm.StartMonth OrElse Year <> RecordTerm.StartYear) Then
              Dim m As Integer = If(Month = 1, 12, Month - 1)
              Dim y As Integer = If(m = 12, Year - 1, Year)
              Dim lastWeek As Integer = MyCalendar.GetLastWeekNum(y, m)
              Return CreateString(lastWeek, m) & "/" & CreateString(Week, Month)
            End If
          ElseIf Month <> RecordTerm.EndMonth OrElse Year <> RecordTerm.EndYear Then
            Dim m As Integer = If(Month = 12, 1, Month + 1)
            Dim y As Integer = If(m = 1, Year + 1, Year)
            Return CreateString(Week, Month) & "/" & CreateString(1, m)
          End If
        End If

        Return CreateString(Week, Month)
      End Function

      Private Function CreateString(w As Integer, m As Integer) As String
        Return m & "月 第" & w & "週"
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

    Public Class RecordConverter
      Public Shared Function filter(record As SheetRecord, f As Func(Of RowRecord, Boolean)) As SheetRecord
        Return record.Filtering(f)
      End Function

      Public Shared Function CreateSumRowRecord(record As SheetRecord, startColIdx As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As RowRecord
        Dim go As Func(Of Integer, RowRecord) =
          Function(idx)
            If idx = 7 Then
              Return RowRecord.Nil
            Else
              Dim offset = idx * 3 + startColIdx
              Dim sum1 As Double = RecordConverter.sum(record, offset, filter)
              Dim sum2 As Double = RecordConverter.sum(record, offset + 1, filter)

              Dim sum3 As Double = 0.0
              If sum2 > 0 Then
                sum3 = sum1 / sum2
              End If
              Return go(idx + 1).AddFirst(sum3.ToString).AddFirst(sum2.ToString).AddFirst(sum1.ToString)
            End If
          End Function

        Return go(0)
      End Function

      Public Shared Function sum(record As SheetRecord, col As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As Double
        Return _
          record.FoldLeft(
            0.0,
            Function(value, rec)
              Dim d As Double = 0
              If col < rec.Count AndAlso filter(rec, col) Then
                d = ToDouble(rec.GetItem(col))
              End If
              Return value + d
            End Function)
      End Function

      Public Shared Function ToInt(r As String) As Integer
        If Utils.General.MyChar.IsInteger(r) Then
          Return Integer.Parse(r)
        Else
          Return 0
        End If
      End Function

      Public Shared Function ToDouble(r As String) As Double
        If Utils.General.MyChar.IsDouble(r) Then
          Return Double.Parse(r)
        Else
          Return 0.0
        End If
      End Function

      Public Shared Function ToCSV(record As SheetRecord) As String
        Dim rec As SheetRecord =
          record.ConvertAll(
            Function(row)
              Return _
                row.ConvertAll(
                  Function(e)
                    If Utils.General.MyChar.IsDouble(e) Then
                      Dim d As Double = Double.Parse(e)
                      Return Math.Round(d, 2).ToString
                    Else
                      Return e
                    End If
                  End Function)
            End Function)

        Return _
          rec.
            ConvertAll(Function(row) row.MkString(",")).
            MkString(vbCrLf)
      End Function
    End Class

  End Namespace

  Namespace Excel
    Public Class ColTree
      Public Col As String
      Public IsColThatReturnValue As Boolean
      Public DependedCols As MyLinkedList(Of ColTree)
      Public AllCount As Integer
    End Class

    Public Structure AccessProperties
      Dim RecordKey As String
      Dim SheetName As String
      Dim Cols As List(Of String)
      Dim FirstRow As Integer
      Dim RowSize As Integer
    End Structure

    Public Structure AccessPropertiesWithTree
      Dim RecordKey As String
      Dim SheetName As String
      Dim SheetFormat As SheetFormat
    End Structure

    Public Structure SheetFormat
      Dim ColTree As ColTree
      Dim OffsetRow As Integer
      Dim RowSize As Integer
    End Structure

    Public Structure ReadRecord
      Dim SheetName As String
      Dim SheetRecord As SheetRecord
      'Dim AccessProperties As AccessProperties
    End Structure

    Public Class ExcelAccessor
      Private Excel As Office.Excel

      Public Sub New(excel As Office.Excel)
        Me.Excel = excel
      End Sub

      Public Sub Init()
        Excel.Init()
      End Sub

      Public Sub Quit()
        Excel.Quit()
      End Sub

      Public Sub Open(fileName As String)
        Excel.Open(fileName, True)
      End Sub

      Public Sub Close()
        Excel.Close()
      End Sub

      Public Function Read(fileName As String, props As List(Of AccessProperties)) As List(Of ReadRecord)
        Dim list As New List(Of ReadRecord)

        Try
          Open(fileName)
          props.ForEach(
            Sub(p)
              Dim rp As ReadRecord
              With rp
                .SheetRecord = ReadSheetRecord(p)
                .SheetName = p.SheetName
              End With

              list.Add(rp)
            End Sub)

        Finally
          Close()
        End Try

        Return list
      End Function

      Public Function ReadWithTree(fileName As String, props As List(Of AccessPropertiesWithTree)) As List(Of ReadRecord)
        Dim list As New List(Of ReadRecord)

        Try
          Open(fileName)
          props.ForEach(
            Sub(p)
              Dim rp As ReadRecord
              With rp
                .SheetRecord = ReadSheetRecordWithTree(p)
                .SheetName = p.SheetName
              End With

              list.Add(rp)
            End Sub)

        Finally
          Close()
        End Try

        Return list
      End Function

      Private Function ReadSheetRecord(prop As AccessProperties) As SheetRecord
        Dim sRecord As SheetRecord = SheetRecord.Nil()

        Dim cells As List(Of Office.Cell) = CreateCellList(prop.Cols, prop.FirstRow, prop.RowSize)
        Dim tmp As List(Of String) = Excel.Read(prop.SheetName, ExtractValidCells(cells))
        Dim values As List(Of String) = MakeRecordList(cells, tmp)

        Dim go As Func(Of Integer, Integer, RowRecord) =
          Function(idx, max)
            If idx > max Then
              Return RowRecord.Nil()
            Else
              Return go(idx + 1, max).AddFirst(values(idx))
            End If
          End Function

        For row As Integer = 0 To (prop.RowSize - 1)
          Dim offset As Integer = row * prop.Cols.Count()
          Dim rec As RowRecord = go(offset, offset + prop.Cols.Count() - 1)

          sRecord = sRecord.AddFirst(rec)
        Next

        Return sRecord.reverse
      End Function

      Private Function ReadSheetRecordWithTree(prop As AccessPropertiesWithTree) As SheetRecord
        Dim sRecord As SheetRecord = SheetRecord.Nil()

        Dim cells As List(Of Office.CellTree) = CreateCellTree(prop.SheetFormat.ColTree, prop.SheetFormat.OffsetRow, prop.SheetFormat.RowSize)
        Dim values As List(Of String) = Excel.Read(prop.SheetName, cells)

        Dim createRowRecord As Func(Of Integer, Integer, RowRecord) =
          Function(idx, max)
            Return If(idx < max,
              createRowRecord(idx + 1, max).AddFirst(values(idx)),
              RowRecord.Nil())
          End Function

        Dim createSheetRecord As Func(Of Integer, Integer, SheetRecord) =
          Function(row, max)
            If row < max Then
              Dim offset As Integer = row * prop.SheetFormat.ColTree.AllCount
              Dim rec As RowRecord = createRowRecord(offset, offset + prop.SheetFormat.ColTree.AllCount)
              Return createSheetRecord(row + 1, max).AddFirst(rec)
            Else
              Return SheetRecord.Nil
            End If
          End Function

        'For row As Integer = 0 To (prop.RowSize - 1)
        '  Dim offset As Integer = row * prop.ColTree.AllCount
        '  Dim rec As RowRecord = createRowRecord(offset, offset + prop.ColTree.AllCount)

        '  sRecord = sRecord.AddFirst(rec)
        'Next

        Return createSheetRecord(0, prop.SheetFormat.RowSize)
      End Function

      Private Function CreateCellTree(cols As ColTree, offsetRow As Integer, rowCnt As Integer) As List(Of Office.CellTree)
        Dim f As Func(Of ColTree, Integer, Office.CellTree) =
          Function(tree, row)
            Dim cell As Office.Cell
            With cell
              .Col = If(Char.IsLetter(CType(tree.Col, Char)), Office.Alph.ToInt(tree.Col), -1)
              .Row = row
              .WrittenText = ""
            End With

            Dim childs As MyLinkedList(Of Office.CellTree) =
              tree.DependedCols.ConvertAll(Function(t) f(t, row))

            Dim cellTree As New Office.CellTree
            With cellTree
              .Cell = cell
              .IsCellThatReturnValue = tree.IsColThatReturnValue
              .NextCell = childs
            End With

            Return cellTree
          End Function

        Dim l As New List(Of Office.CellTree)
        For row As Integer = 0 To rowCnt - 1
          l.Add(f(cols, offsetRow + row))
        Next

        Return l
      End Function

      Private Function CreateCellList(cols As List(Of String), offsetRow As Integer, rowCnt As Integer) As List(Of Office.Cell)
        Dim l As New List(Of Office.Cell)
        For idx As Integer = 0 To (rowCnt - 1)
          Dim row As Integer = offsetRow + idx
          Dim cells As List(Of Office.Cell) =
            cols.ConvertAll(
              Function(k)
                Dim col As Integer = If(Char.IsLetter(CType(k, Char)), Office.Alph.ToInt(k), -1)
                Dim cell As Office.Cell
                With cell
                  .Row = row
                  .Col = col
                  .WrittenText = ""
                End With
                Return cell
              End Function)
          l.AddRange(cells)
        Next
        Return l
      End Function

      Private Function ExtractValidCells(cells As List(Of Office.Cell)) As List(Of Office.Cell)
        Return cells.FindAll(Function(cell) IsValidCell(cell))
      End Function

      Private Function IsValidCell(cell As Office.Cell) As Boolean
        Return cell.Row > 0 AndAlso cell.Col > 0
      End Function

      Private Function MakeRecordList(cells As List(Of Office.Cell), record As List(Of String)) As List(Of String)
        Dim idx As Integer = 0
        Return cells.ConvertAll(
          Function(cell)
            If IsValidCell(cell) Then
              idx += 1
              Return record(idx - 1)
            Else
              Return ""
            End If
          End Function)
      End Function

      Public Shared Function CreateTreeList(parent As String, isParentColThatReturnValue As Boolean, childs As List(Of List(Of String))) As ColTree
        Dim cList As MyLinkedList(Of ColTree) = CreateSequentialTreeList(childs)

        Dim tree As New ColTree
        With tree
          .Col = parent
          .IsColThatReturnValue = isParentColThatReturnValue
          .DependedCols = cList
          .AllCount = If(isParentColThatReturnValue, 1, 0) + cList.FoldLeft(0, Function(sum, e) e.AllCount + sum)
        End With
        Return tree
      End Function

      Public Shared Function CreateSequentialTreeList(colsArray As List(Of List(Of String))) As MyLinkedList(Of ColTree)
        Dim f As Func(Of List(Of List( OF String)), MyLinkedList(Of ColTree)) =
          Function(c)
            If c.Count = 0 Then
              Return MyLinkedList(Of ColTree).Nil
            Else
              Return f(c.Skip(1).ToList).AddFirst(CreateSequentialTree(c.First.ToArray))
            End If
          End Function

        Return f(colsArray)
      End Function

      Public Shared Function CreateSequentialTree(ParamArray cols As String()) As ColTree
        Dim f As Func(Of String(), ColTree) =
          Function(k)
            Dim n As MyLinkedList(Of ColTree) = MyLinkedList(Of ColTree).Nil
            If k.Count > 1 Then
              n = n.AddFirst(f(k.Skip(1).ToArray))
            End If

            Dim t As New ColTree
            With t
              .Col = k.First
              .IsColThatReturnValue = True
              .DependedCols = n
              .AllCount = 1 + n.FoldLeft(0, Function(sum, e) e.AllCount + sum)
            End With

            Return t
          End Function
        Return f(cols)
      End Function
    End Class

  End Namespace

  Namespace Model
    Public Class ExpandedUserInfo
      Public UserInfo As UserInfo

      Public Sub New(userInfo As UserInfo)
        Me.UserInfo = userInfo
      End Sub

      Public Function GetIdNum() As String
        Return UserInfo.GetIdNum()
      End Function

      Public Function GetName() As String
        Return UserInfo.Name
      End Function

      Public Overrides Function ToString() As String
        Return GetIdNum() & " - " & GetName()
      End Function
    End Class

    Public Class UserRecord
      Private UserInfo As ExpandedUserInfo
      Private RecordAndProperties As IDictionary(Of String, Excel.ReadRecord)

      Public Sub New(userInfo As ExpandedUserInfo)
        Me.UserInfo = userInfo
        Me.RecordAndProperties = New Dictionary(Of String, Excel.ReadRecord)
      End Sub

      Public Function GetIdNum() As String
        Return UserInfo.GetIdNum()
      End Function

      Public Function GetName() As String
        Return UserInfo.GetName()
      End Function

      Public Sub Add(key As String, recordAndProperty As Excel.ReadRecord)
        Me.RecordAndProperties.Add(key, recordAndProperty)
      End Sub

      Public Function ContainsKey(key As String) As Boolean
        Return RecordAndProperties.ContainsKey(key)
      End Function

      Public Function GetSheetRecord(key As String) As SheetRecord
        Return RecordAndProperties(key).SheetRecord
      End Function

    End Class

  End Namespace

  Namespace Layout
    Public Class Handler
      Public DoubleClickCallBack As Action(Of Object, EventArgs)

      Sub DoubleClick(sender As Object, e As EventArgs)
        If DoubleClickCallBack IsNot Nothing Then
          DoubleClickCallBack(sender, e)
        End If
      End Sub
    End Class

    Public Class TableDrawer
      Private ScrolledPanel As Panel

      Private TextCols As List(Of Integer)
      Private NoteCols As List(Of Integer)

      Private FuncBackColor As Func(Of Integer, Color)

      Public Sub New(scrolledPanel As Panel)
        Me.ScrolledPanel = scrolledPanel
        TextCols = New List(Of Integer)
        NoteCols = New List(Of Integer)
        FuncBackColor = Function(row) Color.Transparent
      End Sub

      Private Sub New(scrolledPanel As Panel, textCols As List(Of Integer), noteCols As List(Of Integer))
        Me.ScrolledPanel = scrolledPanel
        Me.TextCols = textCols
        Me.NoteCols = noteCols
        FuncBackColor = Function(row) Color.Transparent
      End Sub

      Public Function SetTextCols(ParamArray cols As Integer()) As TableDrawer
        TextCols.AddRange(cols)
        Return Me
      End Function

      Public Function SetNoteCols(ParamArray cols As Integer()) As TableDrawer
        NoteCols.AddRange(cols)
        Return Me
      End Function

      Public Function SetFuncBackColor(f As Func(Of Integer, Color)) As TableDrawer
        FuncBackColor = f
        Return Me
      End Function

      Public Function CreateCell(text As String, insertCol As Integer, insertRow As Integer, handler As Handler) As Panel
        Dim panel As Panel = CreatePanel(FuncBackColor(insertRow), handler)
        Dim label As Label = CreateLabel(text, insertCol, handler)
        panel.Controls.Add(label)
        Return panel
      End Function

      Public Function CreatePanel(backColor As Color, handler As Handler) As Panel
        Dim panel As Panel = ControlDrawer.CreatePanelInTable(backColor)
        AddHandler panel.Click, AddressOf ClickEvent
        If handler IsNot Nothing Then
          AddHandler panel.DoubleClick, AddressOf handler.DoubleClick
        End If
        Return panel
      End Function

      Public Function CreateLabel(text As String, insertCol As Integer, handler As Handler) As Label
        Dim label As Label
        If TextCols.Contains(insertCol) Then
          label = ControlDrawer.CreateTextLabelInTable(text)
        ElseIf NoteCols.Contains(insertCol)
          label = ControlDrawer.CreateNoteLabelInTable(text)
        Else
          label = ControlDrawer.CreateNumberLabelInTable(text)
        End If

        AddHandler label.Click, AddressOf ClickEvent
        If handler IsNot Nothing Then
          AddHandler label.DoubleClick, AddressOf handler.DoubleClick
        End If
        Return label
      End Function

      Private Sub ClickEvent(sender As Object, e As EventArgs)
        ScrolledPanel.Focus()
      End Sub

      Public Function GetColor(insertRow As Integer) As Color
        Return FuncBackColor(insertRow)
      End Function

    End Class

    Public Class ControlDrawer
      Public Shared Function CreateTextPanelInTable(text As String, backColor As Color) As Panel
        Return Create(text, DockStyle.Left, backColor, False)
      End Function

      Public Shared Function CreateNumberPanelInTable(numText As String, backColor As Color) As Panel
        Return Create(numText, DockStyle.Right, backColor, False)
      End Function

      Public Shared Function CreateNotePanelInTable(text As String, backColor As Color) As Panel
        Return Create(text, DockStyle.Left, backColor, True)
      End Function

      Private Shared Function Create(text As String, dock As DockStyle, backColor As Color, useToolTip As Boolean) As Panel
        Dim panel As Panel = CreatePanelInTable(backColor)
        Dim label As Label = CreateLabelInTable(text, dock, useToolTip)
        panel.Controls.Add(label)
        Return panel
      End Function

      Public Shared Function CreatePanelInTable(backColor As Color) As Panel
        Dim panel As Panel = New Panel()
        panel.Margin = New Padding(1, 1, 1, 1)
        panel.Dock = DockStyle.Fill
        panel.BackColor = backColor
        Return panel
      End Function

      Public Shared Function CreateTextLabelInTable(text As String) As Label
        Return CreateLabelInTable(text, DockStyle.Left, False)
      End Function

      Public Shared Function CreateNumberLabelInTable(numText As String) As Label
        Return CreateLabelInTable(numText, DockStyle.Right, False)
      End Function

      Public Shared Function CreateNoteLabelInTable(text As String) As Label
        Return CreateLabelInTable(text, DockStyle.Left, True)
      End Function

      Public Shared Function CreateLabelInTable(text As String, dock As DockStyle, useToolTip As Boolean) As Label
        Dim label As Label = New Label()
        label.Text = text
        label.AutoSize = True
        label.Dock = dock
        label.TextAlign = ContentAlignment.MiddleCenter
        AddHandler label.Click, AddressOf ClickEvent
        If useToolTip Then
          Dim tip As ToolTip = New ToolTip()
          tip.SetToolTip(label, text)
        End If
        Return label
      End Function

      Private Shared Sub ClickEvent(sender As Object, e As EventArgs)
        RecordTableForm.pnlForTable.Focus()
      End Sub
    End Class
  End Namespace
End Namespace