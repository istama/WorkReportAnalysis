
Imports MP.Utils.Model

Namespace WorkReportAnalysis
  Namespace App
    Public Class WorkReportAnalysisProperties
      Private Shared SETTING_FILE_NAME = "excel.properties"

      Public Shared KEY_YEAR = "Year"
      Public Shared KEY_MAIN_FORM_NAME = "MainFormName"
      Public Shared KEY_EXCEL_FILE_DIR = "ExcelFileDir"
      Public Shared KEY_EXCEL_FILE_NAME_FORMAT = "ExcelFileNameFormat"
      Public Shared KEY_SHEET_NAME_FORMAT = "SheetNameFormat"
      Public Shared KEY_SHEET_NAME_FORMAT2 = "SheetNameFormat2"
      Public Shared KEY_FIRST_DAY_OF_A_MONTH_ROW = "FirstDayOfAMonthRow"
      Public Shared KEY_FIRST_ROW2 = "FirstRow2"

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
        dict(KEY_SHEET_NAME_FORMAT2) = "集計"

        dict(KEY_FIRST_DAY_OF_A_MONTH_ROW) = 7
        dict(KEY_FIRST_ROW2) = 12

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
        Return m.GetValue(App.WorkReportAnalysisProperties.KEY_SHEET_NAME_FORMAT2)
      End Function

      Public Shared Function GetFirstDayOfAMonthRow() As Integer
        Dim row As String = m.GetValue(App.WorkReportAnalysisProperties.KEY_FIRST_DAY_OF_A_MONTH_ROW)
        If Char.IsDigit(row) Then
          Return Integer.Parse(row)
        Else
          Throw New Exception("プロパティ<" & App.WorkReportAnalysisProperties.KEY_FIRST_DAY_OF_A_MONTH_ROW & ">の値が不正です。")
        End If
      End Function

      Public Shared Function GetFirstRow() As Integer
        Return GetIntger(App.WorkReportAnalysisProperties.KEY_FIRST_DAY_OF_A_MONTH_ROW)
      End Function

      Public Shared Function GetFirstRow2() As Integer
        Return GetIntger(App.WorkReportAnalysisProperties.KEY_FIRST_ROW2)
      End Function

      Private Shared Function GetIntger(key As String) As Integer
        Dim value As String = m.GetValue(key)
        If Char.IsDigit(value) Then
          Return Integer.Parse(value)
        Else
          Throw New Exception("プロパティ<" & key & ">の値が不正です。")
        End If
      End Function

      Public Shared Function GetItemCols() As List(Of String)
        Return App.WorkReportAnalysisProperties.ItemKeys().ToList.
          ConvertAll(Function(k) If(Char.IsLetter(k), m.GetValue(k), ""))
      End Function

      Public Shared Function GetYear() As Integer
        Return GetIntger(App.WorkReportAnalysisProperties.KEY_YEAR)
      End Function
    End Class
  End Namespace

  Namespace Table
    Public Class RecordArranger

      Public Shared Function InsertEmpty(record As List(Of Model.RowRecord), setHeader As Boolean, ParamArray rows As Integer()) As List(Of Model.RowRecord)
        Dim newRecord As New List(Of Model.RowRecord)

        For idx As Integer = 0 To (record.Count - 1)
          If rows.Contains(idx) Then
            newRecord.Add(EmptyRowRecord(record.First.List.Count, setHeader))
          End If
          newRecord.Add(record(idx))
        Next

        Return newRecord
      End Function

      Public Shared Function PadTailWithEmpty(record As List(Of Model.RowRecord), setHeader As Boolean, size As Integer) As List(Of Model.RowRecord)
        Dim newRecord As New List(Of Model.RowRecord)

        For idx As Integer = 0 To size - 1
          If idx < record.Count Then
            newRecord.Add(record(idx))
          Else
            newRecord.Add(EmptyRowRecord(record.First.List.Count, setHeader))
          End If
        Next

        Return newRecord
      End Function

      Private Shared Function EmptyRowRecord(cnt As Integer, setHeader As Boolean) As Model.RowRecord
        Dim r As Model.RowRecord
        With r
          Dim l As New List(Of String)
          l.Add(If(setHeader, Excel.ExcelAccessor.ROW_RECORD_HEADER, ""))
          For i As Integer = 2 To cnt
            l.Add("")
          Next
          .List = l
        End With
        Return r
      End Function
    End Class
  End Namespace

  Namespace Control
    Public Class UserRecordManager
      Private UserInfoList As List(Of Model.ExpandedUserInfo)
      Private UserRecordMap As IDictionary(Of String, Model.UserRecord)
      Private ExcelReader As Excel.ExcelReader

      Public Sub New(reader As Excel.ExcelReader, userInfoList As List(Of Model.ExpandedUserInfo))
        Me.UserInfoList = userInfoList
        UserRecordMap = New Dictionary(Of String, Model.UserRecord)
        ExcelReader = reader
      End Sub

      Public Function GetUserInfoList() As List(Of Model.ExpandedUserInfo)
        Return UserInfoList
      End Function

      Public Function GetUserRecord(id As String) As Model.UserRecord
        Return UserRecordMap(id)
      End Function

      Public Function GetAllUserRecord() As List(Of Model.UserRecord)
        Return UserRecordMap.Values
      End Function

      Public Function GetTotalRecordAt(month As Integer, day As Integer) As List(Of Model.RowRecord)
        Dim key As String = App.FileFormat.GetSheetName(month)
        Return _
          UserRecordMap.Values.ToList.
            ConvertAll(Function(r) r.GetSheetRecord(key)).
            ConvertAll(Function(s) s.GetAll()(day - 1))
      End Function

      Public Function ReadUserRecord(userInfo As Model.ExpandedUserInfo) As Model.UserRecord
        Dim userRecord As Model.UserRecord

        SyncLock Me
          If UserRecordMap Is Nothing OrElse Not UserRecordMap.ContainsKey(userInfo.GetIdNum()) Then
            userRecord = New Model.UserRecord(userInfo)

            Dim fileName As String = App.FileFormat.GetFileName(userInfo.GetIdNum())
            ExcelReader.read(fileName).ToList.
            ForEach(Sub(kv) userRecord.Add(kv.Key, kv.Value))

            UserRecordMap.Add(userInfo.GetIdNum(), userRecord)
          Else
            'MessageBox.Show("file that already has been read.")
            userRecord = UserRecordMap(userInfo.GetIdNum)
          End If
        End SyncLock

        Return userRecord
      End Function
    End Class
  End Namespace

  Namespace Excel
    Public Structure AccessProperties
      Dim RecordKey As String
      Dim SheetName As String
      Dim Cols As List(Of String)
      Dim FirstRow As Integer
      Dim RowCnt As Integer
      Dim UseRowRecordHeader As Boolean
    End Structure

    Public Class ExcelReader
      Public KEY_TOTAL_SHEET_RECORD As String = "Total"

      Private Excel As Office.Excel
      Private AccessPropList As List(Of AccessProperties)

      Public Sub New(year As Integer)
        Me.Excel = New Office.Excel()
        AccessPropList = New List(Of AccessProperties)

        For month As Integer = 10 To 12
          Dim p As AccessProperties
          With p
            .RecordKey = App.FileFormat.GetSheetName(month)
            .SheetName = App.FileFormat.GetSheetName(month)
            .Cols = App.FileFormat.GetItemCols()
            .FirstRow = App.FileFormat.GetFirstDayOfAMonthRow()
            .RowCnt = Date.DaysInMonth(year, month) + 1
            .UseRowRecordHeader = True
          End With
          AccessPropList.Add(p)
        Next

        Dim t As AccessProperties
        With t
          .RecordKey = App.FileFormat.GetSheetName2
          .SheetName = App.FileFormat.GetSheetName2()
          .Cols = App.FileFormat.GetItemCols()
          .FirstRow = App.FileFormat.GetFirstRow2()
          .RowCnt = 6 * 3 + 1
          .UseRowRecordHeader = True
        End With
        AccessPropList.Add(t)
      End Sub

      Public Sub Init()
        ' Fix 本番ではコメントアウトしない
        'Excel.Init()
      End Sub

      Public Sub Quit()
        ' Fix 本番ではコメントアウトしない
        'Excel.Quit()
      End Sub

      Public Function read(fileName As String) As IDictionary(Of String, Model.SheetRecord)
        Dim a As ExcelAccessor = New ExcelAccessor(Excel)
        Dim dict As New Dictionary(Of String, Model.SheetRecord)
        Try
          ' Fix 本番ではコメントアウトしない
          'a.Open(fileName)

          AccessPropList.
            ForEach(Sub(p) dict.Add(p.RecordKey, a.ReadSheetRecord(p)))
        Catch ex As Exception
          Throw ex
        Finally
          ' Fix 本番ではコメントアウトしない
          'a.Close()
        End Try

        Return dict
      End Function
    End Class

    Public Class ExcelAccessor
      Public Shared ROW_RECORD_HEADER = "<HEAD>"

      Private Shared MAX_DAYS_IN_A_MONTH As Integer = 31
      Private Excel As Office.Excel

      Public Sub New(excel As Office.Excel)
        Me.Excel = excel
      End Sub

      Public Sub Open(fileName As String)
        Excel.Open(fileName, True)
      End Sub

      Public Sub Close()
        Excel.Close()
      End Sub

      Public Function ReadSheetRecord(prop As AccessProperties) As Model.SheetRecord
        Dim sheetRecord As Model.SheetRecord = New Model.SheetRecord()

        Dim cells As List(Of Office.Cell) = CreateCellList(prop.Cols, prop.FirstRow, prop.RowCnt)
        Dim values As List(Of String) = Excel.Read(prop.SheetName, ExtractValidCells(cells))
        Dim record As List(Of String) = MakeRecordList(cells, values)

        For row As Integer = 0 To (prop.RowCnt - 1)
          Dim l As New List(Of String)

          If prop.UseRowRecordHeader Then
            l.Add(ROW_RECORD_HEADER)
          End If

          Dim offset As Integer = row * prop.Cols.Count()
          For idx As Integer = 0 To (prop.Cols.Count() - 1)
            l.Add(record(offset + idx))
          Next

          Dim rowRecord As Model.RowRecord
          With rowRecord
            .List = l
          End With

          sheetRecord.Add(rowRecord)
        Next

        Return sheetRecord
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
        'Dim l As New List(Of Office.Cell)
        'For Each cell As Office.Cell In cells
        '  If IsValidCell(cell) Then
        '    l.Add(cell)
        '  End If
        'Next
        'Return l
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
        'Dim l As New List(Of String)
        'Dim idx As Integer = 0
        'For Each cell As Office.Cell In cells
        '  If IsValidCell(cell) Then
        '    l.Add(record(idx))
        '    idx += 1
        '  Else
        '    l.Add("")
        '  End If
        'Next
        'Return l
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
      Private Record As IDictionary(Of String, SheetRecord)

      Public Sub New(userInfo As ExpandedUserInfo)
        Me.UserInfo = userInfo
        Me.Record = New Dictionary(Of String, SheetRecord)
      End Sub

      Public Function GetIdNum() As String
        Return UserInfo.GetIdNum()
      End Function

      Public Function GetName() As String
        Return UserInfo.GetName()
      End Function

      Public Sub Add(key As String, record As SheetRecord)
        Me.Record.Add(key, record)
      End Sub

      Public Function ContainsKey(key As String) As Boolean
        Return Record.ContainsKey(key)
      End Function

      Public Function GetSheetRecord(key As String) As SheetRecord
        Return Record(key)
      End Function
    End Class

    Public Class SheetRecord
      Private Record As List(Of RowRecord)

      Public Sub New()
        Record = New List(Of RowRecord)()
      End Sub

      Public Sub Add(ByVal record As RowRecord)
        Me.Record.Add(record)
      End Sub

      Public Function GetAll() As List(Of RowRecord)
        Return Record.ToList
      End Function

      Public Function GetFilteringRecord(filter As Filter) As List(Of RowRecord)
        Return Record.FindAll(Function(r) filter(r))
      End Function

      Public Delegate Function Filter(r As RowRecord) As Boolean

    End Class

    Public Module RecordFileters
      Public All As SheetRecord.Filter = AddressOf _All

      Private Function _All(r As RowRecord) As Boolean
        Return True
      End Function
    End Module

    Public Structure RowRecord
      Dim List As List(Of String)
    End Structure

    Module RecordConverter
      Public Function ToInt(r As String) As Integer
        If Char.IsDigit(r) Then
          Return Integer.Parse(r)
        Else
          Return 0
        End If
      End Function

      Public Function ToDouble(r As String) As Double
        If Char.IsDigit(r) Then
          Return Double.Parse(r)
        Else
          Return 0.0
        End If
      End Function
    End Module
  End Namespace

  Namespace Layout
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
        AddHandler panel.Click, AddressOf ClickEvent
        Return panel
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