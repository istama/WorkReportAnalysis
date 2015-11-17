
Imports MP.Utils.Model

Imports MP.Utils.MyCollection.Immutable
Imports RowRecord = MP.Utils.MyCollection.Immutable.MyLinkedList(Of String)
Imports SheetRecord = MP.Utils.MyCollection.Immutable.MyLinkedList(Of MP.Utils.MyCollection.Immutable.MyLinkedList(Of String))

Namespace WorkReportAnalysis
  Namespace App
    Public Class WorkReportAnalysisProperties
      Private Shared SETTING_FILE_NAME = "excel.properties"

      Public Shared KEY_YEAR = "Year"
      Public Shared KEY_MAIN_FORM_NAME = "MainFormName"
      Public Shared KEY_EXCEL_FILE_DIR = "ExcelFileDir"
      Public Shared KEY_EXCEL_FILE_NAME_FORMAT = "ExcelFileNameFormat"
      Public Shared KEY_SHEET_NAME_FORMAT = "SheetNameFormat"
      Public Shared KEY_SUM_SHEET_NAME_FORMAT = "SheetNameFormat2"
      Public Shared KEY_FIRST_DAY_OF_A_MONTH_ROW = "FirstDayOfAMonthRow"
      Public Shared KEY_FIRST_MONTH_OF_SUM_SHEET_ROW = "FirstRow2"

      Public Shared KEY_ITEM_NAME1 = "ItemName1"
      Public Shared KEY_COL1_OF_ITEM1 = "Col1OfItem1"
      Public Shared KEY_COL2_OF_ITEM1 = "Col2OfItem1"
      Public Shared KEY_COL3_OF_ITEM1 = "Col3OfItem1"
      Public Shared KEY_ITEM_NAME2 = "ItemName2"
      Public Shared KEY_COL1_OF_ITEM2 = "Col1OfItem2"
      Public Shared KEY_COL2_OF_ITEM2 = "Col2OfItem2"
      Public Shared KEY_COL3_OF_ITEM2 = "Col3OfItem2"
      Public Shared KEY_ITEM_NAME3 = "ItemName3"
      Public Shared KEY_COL1_OF_ITEM3 = "Col1OfItem3"
      Public Shared KEY_COL2_OF_ITEM3 = "Col2OfItem3"
      Public Shared KEY_COL3_OF_ITEM3 = "Col3OfItem3"
      Public Shared KEY_ITEM_NAME4 = "ItemName4"
      Public Shared KEY_COL1_OF_ITEM4 = "Col1OfItem4"
      Public Shared KEY_COL2_OF_ITEM4 = "Col2OfItem4"
      Public Shared KEY_COL3_OF_ITEM4 = "Col3OfItem4"
      Public Shared KEY_ITEM_NAME5 = "ItemName5"
      Public Shared KEY_COL1_OF_ITEM5 = "Col1OfItem5"
      Public Shared KEY_COL2_OF_ITEM5 = "Col2OfItem5"
      Public Shared KEY_COL3_OF_ITEM5 = "Col3OfItem5"
      Public Shared KEY_ITEM_NAME6 = "ItemName6"
      Public Shared KEY_COL1_OF_ITEM6 = "Col1OfItem6"
      Public Shared KEY_COL2_OF_ITEM6 = "Col2OfItem6"
      Public Shared KEY_COL3_OF_ITEM6 = "Col3OfItem6"
      Public Shared KEY_ITEM_NAME7 = "ItemName7"
      Public Shared KEY_COL1_OF_ITEM7 = "Col1OfItem7"
      Public Shared KEY_COL2_OF_ITEM7 = "Col2OfItem7"
      Public Shared KEY_COL3_OF_ITEM7 = "Col3OfItem7"

      Public Shared KEY_NOTE_COL = "NoteCol"
      Public Shared KEY_WORKDAY_COL = "WorkdayCol"

      Public Shared MANAGER = New MP.Utils.Common.PropertyManager(SETTING_FILE_NAME, DefaultSettingProperties(), True)

      Public Shared Function ItemKeys() As String()
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


      Private Shared Function DefaultSettingProperties() As IDictionary(Of String, String)
        Dim dict As IDictionary(Of String, String) = New Dictionary(Of String, String)
        dict(KEY_YEAR) = "2015"
        dict(KEY_MAIN_FORM_NAME) = "作業件数集計"
        dict(KEY_EXCEL_FILE_DIR) = MP.Details.Sys.App.GetCurrentDirectory()
        dict(KEY_EXCEL_FILE_NAME_FORMAT) = "件数報告書-{0}.xls"
        dict(KEY_SHEET_NAME_FORMAT) = "{0}月分"
        dict(KEY_SUM_SHEET_NAME_FORMAT) = "集計"

        dict(KEY_FIRST_DAY_OF_A_MONTH_ROW) = 7
        dict(KEY_FIRST_MONTH_OF_SUM_SHEET_ROW) = 12

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
        Return App.WorkReportAnalysisProperties.ItemKeys().ToList.
          FindAll(Function(k) Not excludedItemKeys.Contains(k)).
          ConvertAll(
            Function(k)
              Dim col As String = m.GetValue(k)
              Return If(Char.IsLetter(col), col, "")
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

      Public Sub New(excel As Excel.ExcelAccessor, manager As UserRecordManager)
        _UserRecordManager = manager
        Me.Excel = excel
        AccessPropList = New List(Of Excel.AccessProperties)

        Init()
      End Sub

      Private Sub Init()
        AccessPropList.Clear()

        Dim term As MyLinkedList(Of Tuple(Of Integer, Integer)) =
          _UserRecordManager.GetReadRecordTerm.GetTermList()

        term.ForEach(
          Sub(e)
            Dim p As Excel.AccessProperties
            With p
              .RecordKey = UserRecordManager.GetSheetName(e.Item2)
              .SheetName = UserRecordManager.GetSheetName(e.Item2)
              .Cols = App.FileFormat.GetItemCols()
              .FirstRow = App.FileFormat.GetFirstDayOfAMonthRow()
              .RowSize = Date.DaysInMonth(e.Item1, e.Item2) '+ 1
            End With
            AccessPropList.Add(p)
          End Sub)

        Dim sum As Excel.AccessProperties
        With sum
          .RecordKey = UserRecordManager.GetSumSheetName
          .SheetName = UserRecordManager.GetSumSheetName
          .Cols = App.FileFormat.GetItemCols(App.WorkReportAnalysisProperties.KEY_NOTE_COL)
          .FirstRow = App.FileFormat.GetFirstRow2()
          .RowSize = 6 * term.Count '+ 1
        End With
        AccessPropList.Add(sum)
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

          Excel.Read(fileName, AccessPropList).
              ForEach(Sub(res) userRecord.Add(res.AccessProperties.SheetName, res))
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

      Public Shared Function GetSheetName(month) As String
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
          Dim t As Tuple(Of Integer, Integer) =
            RecordTerm.DateList.Find(Function(e) key = GetSheetName(e.Item2))
          If t IsNot Nothing Then
            Return GetUsersDailyRecordInMonth(id, t.Item2, t.Item1)
          Else
            Throw New Exception("指定したキーのデータはありません。 key: " & key)
          End If
        End If
      End Function

      Public Function GetUsersDailyRecordInMonth(id As String, month As Integer, year As Integer) As SheetRecord
        CheckValidDate(1, 1, month, year)

        Dim sheetName As String = GetSheetName(month)
        Dim record As SheetRecord = GetUserRecord(id).GetSheetRecord(sheetName)
        Dim go As Func(Of Integer, SheetRecord, SheetRecord) =
          Function(idx, rec)
            If idx < Date.DaysInMonth(year, month) Then
              Dim rr As RowRecord = rec.First.AddFirst((idx + 1).ToString & "日")
              Return go(idx + 1, rec.Rest).AddFirst(rr)
            ElseIf idx < RecordTableForm.TABLE_ROW_COUNT - 1
              Return go(idx + 1, rec).AddFirst(RowRecord.Nil)
              'ElseIf idx = RecordTableForm.TABLE_ROW_COUNT - 1
              'Dim rr As RowRecord = rec.First.AddFirst("合計")
              'Return go(idx + 1, rec.Rest).AddFirst(rr)
            Else
              Return SheetRecord.Nil
            End If
          End Function
        Return go(0, record)
      End Function

      Public Function GetUserDailyRecordInWeek(id As String, week As Integer, month As Integer, year As Integer) As SheetRecord
        CheckValidDate(1, week, month, year)

        Dim record As SheetRecord = GetUsersDailyRecordInMonth(id, month, year)
        Dim days As List(Of Integer) = Utils.MyDate.MyCalendar.GetDaysInWeek(year, month, week)
        'MessageBox.Show("w: " & week & " m: " & month & " y: " & year & " cnt: " & days.Count & " days: " & days.First)
        'MessageBox.Show(record.Skip(days.First - 1).First.First)
        'MessageBox.Show(record.Skip(days.First - 1).Take(days.Count).First.First)

        Return record.Skip(days.First - 1).Take(days.Count)
      End Function

      Public Function GetUserWeeklyRecordInMonth(id As String, month As Integer, year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        Dim go As Func(Of Integer, SheetRecord) =
          Function(week)
            If week > 5 Then
              Return SheetRecord.Nil
            Else
              Dim rec As SheetRecord = GetUserDailyRecordInWeek(id, week, month, year)
              'MessageBox.Show("w: " & week & " m: " & month & " y: " & year & " cnt: " & rec.Count & " first: " & rec.First.First)

              Dim rr As RowRecord = RecordConverter.CreateSumRowRecord(rec, 1, filter)
              Return go(week + 1).AddFirst(rr)
            End If
          End Function
        Return go(1)
      End Function

      Public Function GetUserMonthlyRecord(id As String, year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        Dim go As Func(Of MyLinkedList(Of Tuple(Of Integer, Integer)), SheetRecord) =
          Function(term)
            If term.Empty Then
              Return SheetRecord.Nil
            Else
              Dim ym As Tuple(Of Integer, Integer) = term.First
              Dim rec As SheetRecord = GetUsersDailyRecordInMonth(id, ym.Item2, ym.Item1)
              Dim rr As RowRecord = RecordConverter.CreateSumRowRecord(rec, 1, filter)
              Return go(term.Rest).AddFirst(rr)
            End If
          End Function
        Return go(RecordTerm.DateList)
      End Function

      Public Function GetUserSumRecord(id As String, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        Dim monthlyList As New List(Of RowRecord)
        Dim go As Func(Of MyLinkedList(Of Tuple(Of Integer, Integer)), SheetRecord) =
          Function(term)
            If term.Empty Then
              Return SheetRecord.Nil
            Else
              Dim ym As Tuple(Of Integer, Integer) = term.First
              Dim title As RowRecord = RowRecord.Nil.AddFirst(ym.Item2 & "月")
              Dim weekly As SheetRecord =
                GetUserWeeklyRecordInMonth(id, ym.Item2, ym.Item1, filter).
                  ZipWithIndex.
                  ConvertAll(Of RowRecord)(
                    Function(t)
                      Dim rr As RowRecord = t.Item1
                      Return rr.AddFirst("第" & (t.Item2 + 1) & "週")
                    End Function)
              Dim sum As RowRecord = RecordConverter.CreateSumRowRecord(weekly, 1, Function(row, col) True).AddFirst("月計")
              monthlyList.Add(sum)

              Dim res As SheetRecord = weekly.AddFirst(title).AddLast(sum)

              Return go(term.Rest).AddRangeToHead(res)
            End If
          End Function

        Dim sRec As SheetRecord = go(RecordTerm.DateList)
        Dim total As RowRecord = RecordConverter.CreateSumRowRecord(SheetRecord.Nil.AddRangeToHead(monthlyList), 1, Function(row, col) True).AddFirst("合計")
        Return sRec.AddLast(total)
      End Function

      'Public Shared Function GetWeekNumInMonth(day As Date) As Integer
      '  Dim first As Date = MonthlyItem.GetFirstDateInMonth(day.Month, day.Year)
      '  Return DatePart("WW", day) - DatePart("ww", first) + 1
      'End Function

      'Public Function GetUsersSumRecord(id As String) As SheetRecord
      '  Dim record As SheetRecord = GetUserRecord(id).GetSheetRecord(GetSumSheetName)
      '  Dim dateList As MyLinkedList(Of Tuple(Of Integer, Integer)) = RecordTerm.DateList
      '  Dim go As Func(Of Integer, SheetRecord, MyLinkedList(Of Tuple(Of Integer, Integer)), SheetRecord) =
      '    Function(rowIdx, rec, term)
      '      If rec.Empty Then
      '        Return SheetRecord.Nil
      '      ElseIf rec.IsLast Then
      '        Return SheetRecord.Nil
      '        'Dim rr As RowRecord = rec.First.AddFirst("合計")
      '        'Return go(rowIdx + 1, rec.Rest, term).AddFirst(rr)
      '      Else
      '        Dim rr As RowRecord = rec.First
      '        Dim idx As Integer = rowIdx Mod 6

      '        If idx < 5 Then
      '          Dim nrr As RowRecord = rr.AddFirst("第" & (idx + 1).ToString & "週")
      '          If idx = 0 Then
      '            Dim empty As RowRecord = RowRecord.Nil
      '            Dim header = If(
      '              Not term.Empty,
      '              term.First.Item2.ToString & "月",
      '              "")
      '            Return go(rowIdx + 1, rec.Rest, term.Rest).AddFirst(nrr).AddFirst(empty.AddFirst(header))
      '          Else
      '            Return go(rowIdx + 1, rec.Rest, term).AddFirst(nrr)
      '          End If
      '        Else
      '          Dim nrr As RowRecord = rr.AddFirst("月計")
      '          Return go(rowIdx + 1, rec.Rest, term).AddFirst(nrr)
      '        End If
      '      End If
      '    End Function
      '  Return go(0, record, dateList)
      'End Function

      Public Function GetDailyTermRecord(day As Integer, month As Integer, year As Integer) As SheetRecord
        CheckValidDate(day, 1, month, year)

        Dim sheetName As String = GetSheetName(month)
        Return _
            GetTermRecord(Function(userRecord) userRecord.GetSheetRecord(sheetName).GetItem(day - 1))
      End Function

      Public Function GetWeeklyTermRecord(week As Integer, month As Integer, year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        CheckValidDate(1, week, month, year)

        'Dim offset As Integer =
        '  RecordTerm.DateList.IndexWhere(Function(t) t.Item2 = month AndAlso t.Item1 = year) * 6
        'Return _
        '      GetTermRecord(Function(userRecord) userRecord.GetSheetRecord(GetSumSheetName).GetItem(offset + week - 1))
        Return GetTermRecord2(
          Function(id)
            'Dim wRec As SheetRecord = GetUserDailyRecordInWeek(id, week, month, year)
            Return GetUserWeeklyRecordInMonth(id, month, year, filter).GetItem(week - 1)
            'Return RecordConverter.CreateSumRowRecord(wRec, 2, filter)
          End Function)
      End Function

      Public Function GetMonthlyTermRecord(month As Integer, year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        CheckValidDate(1, 1, month, year)

        Return _
          GetTermRecord2(
            Function(id)
              Dim idx As Integer = RecordTerm.DateList.IndexWhere(Function(t) month = t.Item2 AndAlso year = t.Item1)
              Return GetUserMonthlyRecord(id, year, filter).GetItem(idx)
            End Function)

        'Dim idx As Integer =
        '    RecordTerm.DateList.IndexWhere(Function(t) t.Item2 = month AndAlso t.Item1 = year) * 6 + 5
        'Return _
        '      GetTermRecord(Function(userRecord) userRecord.GetSheetRecord(GetSumSheetName).GetItem(idx))
      End Function

      Public Function GetAllTermRecord() As SheetRecord
        Dim idx As Integer =
          RecordTerm.DateList.Count * 6
        Return _
            GetTermRecord(Function(userRecord) userRecord.GetSheetRecord(GetSumSheetName).GetItem(idx))
      End Function

      Private Function GetTermRecord(getRowRecord As Func(Of Model.UserRecord, RowRecord)) As SheetRecord
        Dim record As List(Of RowRecord) =
          UserRecordMap.ToList.
            FindAll(Function(kv) UserInfoList.Exists(Function(info) info.GetIdNum = kv.Key)).
            ConvertAll(Of RowRecord)(
              Function(kv)
                Dim key As String = kv.Key.PadLeft(3, "0"c)
                Dim rec As Model.UserRecord = kv.Value
                Dim userName As String = UserInfoList.Find(Function(info) info.GetIdNum = key).GetName
                Dim rr As RowRecord = getRowRecord(rec)
                Return rr.AddFirst(userName).AddFirst(key)
              End Function)

        Return ListToImmutableList(record)
      End Function

      Private Function GetTermRecord2(getRowRec As Func(Of String, RowRecord)) As SheetRecord
        Dim record As List(Of RowRecord) =
          UserInfoList.
            ConvertAll(
              Function(info)
                Dim rr As RowRecord = getRowRec(info.GetIdNum)
                Return rr.AddFirst(info.GetName).AddFirst(info.GetIdNum)
              End Function)
        Return ListToImmutableList(record)
      End Function

      Public Function GetAllUserRecord() As MyLinkedList(Of Model.UserRecord)
        Dim go As Func(Of List(Of Model.UserRecord), MyLinkedList(Of Model.UserRecord)) =
          Function(values)
            If values.Count = 0 Then
              Return MyLinkedList(Of Model.UserRecord).Nil()
            Else
              Return go(values.Skip(1)).AddFirst(values.First)
            End If
          End Function
        Return go(UserRecordMap.Values)
      End Function

      Public Function GetDailyTotalRecord(month As Integer, year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        Dim sheet As New List(Of RowRecord)
        For d As Integer = 1 To Date.DaysInMonth(year, month)
          Dim rec As SheetRecord = GetDailyTermRecord(d, month, year)
          Dim rr As RowRecord = RecordConverter.CreateSumRowRecord(rec, 2, filter).AddFirst(d.ToString & "日").AddFirst("")
          sheet.Add(rr)
        Next
        Return SheetRecord.Nil.AddRangeToHead(sheet)
      End Function

      Public Function GetWeeklyTotalRecord(year As Integer, filter As Func(Of RowRecord, Integer, Boolean)) As SheetRecord
        Dim sheet As New List(Of RowRecord)
        RecordTerm.GetTermList.
          ForEach(
            Sub(t)
              For w As Integer = 1 To 5
                Dim rec As SheetRecord = GetWeeklyTermRecord(w, t.Item2, t.Item1, filter)
                Dim rr As RowRecord = RecordConverter.CreateSumRowRecord(rec, 2, filter).AddFirst(t.Item2.ToString & "月第" & w.ToString & "週").AddFirst("")
                sheet.Add(rr)
              Next
            End Sub)
        Return SheetRecord.Nil.AddRangeToHead(sheet)
      End Function

      Private Sub CheckValidDate(day As Integer, week As Integer, month As Integer, year As Integer)
        If InvalidMonth(month) Then
          Throw New Exception("月の値が不正です。 month:  " & month)
        ElseIf InvalidWeek(week) Then
          Throw New Exception("週の値が不正です。 week: " & week)
        ElseIf InvalidDay(day, month, year) Then
          Throw New Exception("日の値が不正です。 day: " & day & " month: " & month & " year: " & year)
        ElseIf OutOfTerm(month, year) Then
          Throw New Exception("指定した年月のデータはありません。 month: " & month & " year: " & year)
        End If
      End Sub

      Private Function OutOfTerm(month As Integer, year As Integer) As Boolean
        Return Not RecordTerm.DateList.Exists(Function(t) t.Item2 = month AndAlso t.Item1 = year)
      End Function

      Private Function InvalidMonth(month As Integer) As Boolean
        Return month < 1 OrElse month > 12
      End Function

      Private Function InvalidWeek(week As Integer) As Boolean
        Return week < 1 OrElse week > 5
      End Function

      Private Function InvalidDay(day As Integer, month As Integer, year As Integer) As Boolean
        Return day < 1 OrElse day > Date.DaysInMonth(year, month)
      End Function

      Private Function ListToImmutableList(Of T)(list As List(Of T)) As MyLinkedList(Of T)
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

      Private _DateList As MyLinkedList(Of Tuple(Of Integer, Integer))
      Public ReadOnly Property DateList As MyLinkedList(Of Tuple(Of Integer, Integer))
        Get
          Return _DateList
        End Get
      End Property

      Public Sub New(startMonth As Integer, startYear As Integer, endMonth As Integer, endYear As Integer)
        If ValidTerm(startMonth, startYear, endMonth, endYear) Then
          Me._StartMonth = startMonth
          Me._StartYear = startYear
          Me._EndMonth = endMonth
          Me._EndYear = endYear
          _DateList = GetTermList()
          '_DateList.ForEach(Function(t) MessageBox.Show(t.Item1 & " " & t.Item2))
        Else
          Throw New Exception("期間の値が不正です。")
        End If
      End Sub

      Public Function GetTermList() As MyLinkedList(Of Tuple(Of Integer, Integer))
        Return GetTerm(_StartMonth, _StartYear, _EndMonth, _EndYear)
      End Function

      Public Function InTerm(month As Integer, year As Integer) As Boolean
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

      Private Function GetTerm(startMonth As Integer, startYear As Integer, endMonth As Integer, endYear As Integer) As MyLinkedList(Of Tuple(Of Integer, Integer))
        If startYear = endYear AndAlso startMonth = endMonth Then
          Return New MyLinkedList(Of Tuple(Of Integer, Integer))(Tuple.Create(startYear, startMonth))
        Else
          Dim l As MyLinkedList(Of Tuple(Of Integer, Integer)) =
            If(startMonth < 12,
               GetTerm(startMonth + 1, startYear, endMonth, endYear),
               GetTerm(1, startYear + 1, endMonth, endYear))
          Return l.AddFirst(Tuple.Create(startYear, startMonth))
        End If
      End Function

      Private Function ValidTerm(startMonth As Integer, startYear As Integer, endMonth As Integer, endYear As Integer) As Boolean
        Return _
          (startMonth >= 1 AndAlso startMonth <= 12 AndAlso endMonth >= 1 AndAlso endMonth <= 12) AndAlso
          startYear < endYear OrElse (startYear = endYear AndAlso startMonth <= endMonth)
      End Function
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
              Dim sum3 As Double = sum1 / sum2
              Return go(idx + 1).AddFirst(sum3).AddFirst(sum2).AddFirst(sum1)
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
        Return _
          record.
            ConvertAll(Function(row) row.MkString(",")).
            MkString(vbCrLf)
      End Function
    End Class

  End Namespace

  Namespace Excel
    Public Structure AccessProperties
      Dim RecordKey As String
      Dim SheetName As String
      Dim Cols As List(Of String)
      Dim FirstRow As Integer
      Dim RowSize As Integer
    End Structure

    Public Structure RecordAndProperty
      Dim SheetRecord As SheetRecord
      Dim AccessProperties As AccessProperties
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

      Public Function Read(fileName As String, props As List(Of AccessProperties)) As List(Of RecordAndProperty)
        Dim list As New List(Of RecordAndProperty)

        Try
          ' Fix 本番ではコメントアウトしない
          Open(fileName)
          props.ForEach(
            Sub(p)
              Dim rp As RecordAndProperty
              With rp
                .SheetRecord = ReadSheetRecord(p)
                .AccessProperties = p
              End With

              list.Add(rp)
            End Sub)

        Finally
          ' Fix 本番ではコメントアウトしない
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

      Private Function CreateCellList(cols As List(Of String), offsetRow As Integer, rowCnt As Integer) As List(Of Office.Cell)
        Dim l As New List(Of Office.Cell)
        For idx As Integer = 0 To (rowCnt - 1)
          Dim row As Integer = offsetRow + idx
          Dim cells As List(Of Office.Cell) =
            cols.ConvertAll(
              Function(k)
                Dim col As Integer = If(Char.IsLetter(k), Office.Alph.ToInt(k), -1)
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
      Private RecordAndProperties As IDictionary(Of String, Excel.RecordAndProperty)

      Public Sub New(userInfo As ExpandedUserInfo)
        Me.UserInfo = userInfo
        Me.RecordAndProperties = New Dictionary(Of String, Excel.RecordAndProperty)
      End Sub

      Public Function GetIdNum() As String
        Return UserInfo.GetIdNum()
      End Function

      Public Function GetName() As String
        Return UserInfo.GetName()
      End Function

      Public Sub Add(key As String, recordAndProperty As Excel.RecordAndProperty)
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

      Public Function CreateCell(text As String, insertCol As Integer, insertRow As Integer) As Panel
        Dim panel As Panel = CreatePanel(FuncBackColor(insertRow))
        Dim label As Label = CreateLabel(text, insertCol)
        panel.Controls.Add(label)
        Return panel
      End Function

      Public Function CreatePanel(backColor) As Panel
        Dim panel As Panel = ControlDrawer.CreatePanelInTable(backColor)
        AddHandler panel.Click, AddressOf ClickEvent
        Return panel
      End Function

      Public Function CreateLabel(text As String, insertCol As Integer) As Label
        Dim label As Label
        If TextCols.Contains(insertCol) Then
          label = ControlDrawer.CreateTextLabelInTable(text)
        ElseIf NoteCols.Contains(insertCol)
          label = ControlDrawer.CreateNoteLabelInTable(text)
        Else
          label = ControlDrawer.CreateNumberLabelInTable(text)
        End If

        AddHandler label.Click, AddressOf ClickEvent
        Return label
      End Function

      Private Sub ClickEvent(sender As Object, e As MouseEventArgs)
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

      Private Shared Sub ClickEvent(sender As Object, e As MouseEventArgs)
        RecordTableForm.pnlForTable.Focus()
      End Sub
    End Class
  End Namespace
End Namespace