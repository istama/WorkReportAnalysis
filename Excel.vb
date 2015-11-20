
Option Strict Off

Imports System.Runtime.InteropServices.Marshal
Imports MP.Utils.Common

Namespace Office
  Public Structure Cell
    Dim Row As Integer
    Dim Col As Integer
    Dim WrittenText As String
  End Structure

  Public Delegate Function UpdateF(value As String) As String

  Public Class Excel
    Private Enum OpMode
      READ
      WRITE
      UPDATE
    End Enum

    Private XLS As Object
    Private WorkBooks As Object
    Private Book As Object

    Private isInit As Boolean
    Private isOpened As Boolean

    Dim r As New System.Random(100)

    Sub New()
      isInit = False
    End Sub

    Public Sub Init()
      '本番はコメントアウト外す
      'SyncLock Me
      '  XLS = CreateObject("Excel.Application")
      '  WorkBooks = XLS.WorkBooks

      '  isInit = True
      'End SyncLock
    End Sub

    Public Sub Quit()
      '本番はコメントアウト外す
      'SyncLock Me
      '  Close()

      '  If isInit = True AndAlso Not isOpened Then
      '    ReleaseComObj(WorkBooks)
      '    XLS.Quit()
      '    ReleaseComObj(XLS)
      '    isInit = False
      '  End If
      'End SyncLock
    End Sub

    Public Sub Open(filePath As String, readMode As Boolean)
      '本番はコメントアウト外す
      'SyncLock Me
      '  Book = WorkBooks.Open(filePath, Nothing, readMode)
      '  isOpened = True
      'End SyncLock
    End Sub

    Public Sub Close()
      '本番はコメントアウト外す
      'SyncLock Me
      '  If isOpened AndAlso Book IsNot Nothing Then
      '    Book.Close(False)
      '    ReleaseComObj(Book)
      '    isOpened = False
      '  End If
      'End SyncLock
    End Sub

    Public Function Read(sheetName As String, cells As List(Of Cell)) As List(Of String)
      Return Access2(Function(sheet, cell)
                       'MessageBox.Show("row: " & cell.Row & " col " & cell.Col)

                       Return ((cell.Row + 1) * cell.Col + r.Next(100)).ToString
                       'Return GetTextFromExcel(sheet, cell),
                     End Function,
                     sheetName,
                     cells)
    End Function

    Public Function Update(sheetName As String, cells As List(Of Cell), f As UpdateF) As List(Of String)
      Dim res As List(Of String) =
        Access2(Function(sheet, cell)
                  Dim tmp As Object = GetTextFromExcel(sheet, cell)
                  If tmp IsNot Nothing Then
                    Dim val = CType(tmp, String)
                    If f IsNot Nothing Then
                      val = f(val)
                      SetTextToExcel(val, sheet, cell)
                    End If
                    Return val
                  Else
                    Return Nothing
                  End If
                End Function,
                sheetName, cells)
      Book.Save()
      Return res
    End Function

    Public Sub Write(sheetName As String, cells As List(Of Cell))
      Access2(Function(sheet, cell)
                SetTextToExcel(cell.WrittenText, sheet, cell)
                Return ""
              End Function,
              sheetName, cells)
      Book.Save()
    End Sub

    Private Function Access2(access As Func(Of Object, Cell, Object), sheetName As String, cells As List(Of Cell)) As List(Of String)
      '本番はコメントアウトをはずす
      'If isInit = False Then
      '  Throw New Exception("初期処理が実行されていません。")
      'End If

      Dim worksheets As Object = Nothing
      Dim sheet As Object = Nothing
      Dim values As New List(Of Object)

      Try
        '本番はコメントアウトをはずす
        'worksheets = Book.Worksheets
        'sheet = GetSheet(sheetName, worksheets)

        'cells.ForEach(Sub(cell) values.Add(access(sheet, cell)))
        cells.ForEach(Sub(cell) values.Add(access(Nothing, cell)))

        '本番はコメントアウトする
        Return ToStringList(values)
      Catch ex As Exception
        Throw New Exception(ex.Message & vbCrLf & ex.StackTrace)
      Finally
        ReleaseComObj(sheet)
        ReleaseComObj(worksheets)
      End Try

      Return ToStringList(values)
    End Function

    'Private Function Access(opMode As OpMode, sheetName As String, cells As List(Of Cell), f As UpdateF) As List(Of String)
    '  'If isInit = False Then
    '  '  Throw New Exception("初期処理が実行されていません。")
    '  'End If

    '  Dim worksheets As Object = Nothing
    '  Dim sheet As Object = Nothing
    '  Dim values As New List(Of Object)

    '  Try
    '    'worksheets = Book.Worksheets
    '    'sheet = GetSheet(sheetName, worksheets)

    '    If opMode = OpMode.READ Then
    '      For Each cell As Cell In cells
    '        values.Add(((cell.Row + 1) * cell.Col).ToString)
    '        'values.Add(GetTextFromExcel(sheet, cell))
    '      Next
    '    ElseIf opMode = OpMode.WRITE Then
    '      For Each cell As Cell In cells
    '        SetTextToExcel(cell.WrittenText, sheet, cell)
    '      Next
    '    ElseIf OpMode.UPDATE
    '      For Each cell As Cell In cells
    '        Dim tmp As Object = GetTextFromExcel(sheet, cell)
    '        If tmp IsNot Nothing Then
    '          Dim val = CType(tmp, String)
    '          If f IsNot Nothing Then
    '            val = f(val)
    '            SetTextToExcel(val, sheet, cell)
    '          End If
    '          values.Add(val)
    '        Else
    '          values.Add(Nothing)
    '        End If
    '      Next
    '    End If

    '    Return ToStringList(values)
    '    'If opMode <> OpMode.READ Then
    '    '  Book.Save()
    '    'End If
    '  Catch ex As Exception
    '    Throw New Exception(ex.Message & vbCrLf & ex.StackTrace)
    '  Finally
    '    ReleaseComObj(sheet)
    '    ReleaseComObj(worksheets)
    '    'If book IsNot Nothing Then
    '    '  book.Close(False)
    '    '  ReleaseComObj(book)
    '    'End If
    '  End Try

    '  Return ToStringList(values)
    'End Function

    Private Function ToStringList(l As List(Of Object)) As List(Of String)
      Dim texts As New List(Of String)
      For Each t As String In l
        texts.Add(CType(t, String))
      Next
      Return texts
    End Function

    Private Function GetSheet(sheetName As String, sheets As Object) As Object
      Dim idx As Integer = GetSheetIndex(sheetName, sheets)
      If idx > 0 Then
        Return sheets.Item(idx)
      Else
        Throw New Exception("存在しないワークシートです: " & sheetName)
      End If
    End Function

    ' 指定されたワークシート名のインデックスを返すメソッド ワークシートのインデックスは1から始まるので注意
    Private Function GetSheetIndex(sheetName As String, sheets As Object) As Integer
      Dim i As Integer = 1
      For Each sh As Object In sheets
        If sheetName = sh.Name Then
          Return i
        End If
        i += 1
      Next
      Return 0
    End Function

    Private Function GetTextFromExcel(sheet As Object, cell As Cell) As Object
      Dim rng As Object = GetRange(sheet, cell)
      If rng IsNot Nothing Then
        Dim res As Object = rng.Value
        ReleaseComObj(rng)
        Return If(res IsNot Nothing, res, "")
        'If res IsNot Nothing Then
        '  Return res
        'Else
        '  Return ""
        'End If
      Else
        Return Nothing
      End If
    End Function

    Private Sub SetTextToExcel(text As String, sheet As Object, cell As Cell)
      Dim rng As Object = GetRange(sheet, cell)
      If rng IsNot Nothing Then
        rng.Value = text
        ReleaseComObj(rng)
      End If
    End Sub

    Private Function GetRange(sheet As Object, cell As Cell) As Object
      Dim strCell As String = ToCellName(cell.Row, cell.Col)
      Return If(strCell <> "", sheet.Range(strCell), Nothing)
      'If strCell <> "" Then
      '  Return sheet.Range(strCell)
      'Else
      '  Return Nothing
      'End If
    End Function

    Private Sub ReleaseComObj(ByRef com As Object)
      Try
        If com IsNot Nothing Then
          FinalReleaseComObject(com)
        End If
      Finally
        com = Nothing
      End Try
    End Sub

    Private Function ToCellName(row As Integer, col As Integer) As String
      If row > 0 AndAlso col > 0 Then
        Return Alph.ToWord(col) & row.ToString()
      Else
        Return ""
      End If
    End Function

  End Class

  Public Class Alph
    Shared aInt As Integer = Asc("A")

    Shared Function ToInt(s As String) As Integer
      If Not Char.IsLetter(s) Then
        Throw New Exception("文字列がアルファベットではありません。" + s)
      End If

      Dim ca = s.ToCharArray()
      Dim sisu = ca.Length - 1

      Dim num = 0
      For Each c As Char In ca
        If (sisu = 0) Then
          num += ToInt(c)
        Else
          num += ToInt(c) * System.Math.Pow(26, sisu)
        End If
        sisu -= 1
      Next

      Return num
    End Function

    Shared Function ToInt(c As Char) As Integer
      Dim cc = UCase(c)
      Return Asc(cc) - aInt + 1
    End Function

    Public Shared Function ToWord(value As Integer) As String
      Const BASE_NUM As Integer = 26

      If value <= BASE_NUM Then
        Return ToChar(value)
      Else
        Dim left As Integer = (value - 1) \ BASE_NUM
        Return ToWord(left) & ToWord(value - (BASE_NUM * left))
      End If
    End Function

    Shared Function ToChar(offset As Integer) As Char
      If offset < 1 OrElse offset > 26 Then
        Throw New Exception("数値が範囲の外です")
      End If

      Dim a As Char = "A"
      Dim aCode As Integer = Asc(a)
      Return Convert.ToChar(offset + aCode - 1)
    End Function
  End Class
End Namespace