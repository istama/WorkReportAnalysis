'
' 日付: 2016/09/25
' 
Imports System.IO

Imports Common.IO
Imports Common.COM

''' <summary>
''' Excelファイルとアクセスするためのプロパティ。
''' </summary>
Public Class ExcelProperties
  Inherits AppProperties
  
  ' １つの項目が保持している情報
  Public Structure WorkItemParams
    Dim Id As Integer
    Dim Name As String
    Dim WorkCountCol As String
    Dim WorkTimeCol As String
  End Structure
  
  Public Const KEY_EXCEL_FILEDIR   = "ExcelFileDir"
  Public Const KEY_EXCEL_FILENAME  = "ExcelFileName"
  Public Const KEY_EXCEL_SHEETNAME = "ExcelSheetName"
  Public Const KEY_FIRST_ROW       = "FirstRow"
  
  Public Const KEY_BEGIN_DATE      = "BeginDate"
  Public Const KEY_END_DATE        = "EndDate"
  
  Public Const KEY_ITEM_NAME       = "ItemName"
  Public Const KEY_WORKCOUNT_COL   = "WorkCountCol"
  Public COnst KEY_WORKTIME_COL    = "WorkTimeCol"
  
  Public Const KEY_NOTE_NAME       = "NoteName"
  Public Const KEY_NOTE_COL        = "NoteCol"
  
  Public Const KEY_WORKDAY_COL     = "WorkDay"
  
  ''' <summary>
  ''' コンストラクタ。
  ''' </summary>
  ''' <param name="filePath">プロパティファイルのパス</param>
  Public Sub New(filePath As String)
    MyBase.New(filePath)
  End Sub
  
  ''' <summary>
  ''' プロパティファイルのデフォルト値を返す。
  ''' </summary>
  ''' <returns></returns>
  Protected Overrides Function DefaultProperties() As IDictionary(Of String, String)
    Dim p As New Dictionary(Of String, String)
    
    p.Add(KEY_EXCEL_FILEDIR,   ".")
    p.Add(KEY_EXCEL_FILENAME,  "件数記録{0}.xls")
    p.Add(KEY_EXCEL_SHEETNAME, "{0}月分")
    p.Add(KEY_FIRST_ROW,       "7")
    
    p.Add(KEY_BEGIN_DATE, "2016/09/01")
    p.Add(KEY_END_DATE,   "2016/12/31")
    
    p.Add(KEY_ITEM_NAME     & "1", "Item1")
    p.Add(KEY_WORKCOUNT_COL & "1", "C")
    p.Add(KEY_WORKTIME_COL  & "1", "D")
    
    p.Add(KEY_NOTE_NAME, "Note")
    p.Add(KEY_NOTE_COL,  "X")
    
    p.Add(KEY_WORKDAY_COL, "Y")
    
    Return p
  End Function
  
  ''' <summary>
  ''' デフォルトに設定されてないプロパティがプロパティファイルにあることを認めるかどうか。
  ''' 認める場合はTrueを返す。
  ''' </summary>
  ''' <returns></returns>
  Protected Overrides Function AllowNonDefaultProperty() As Boolean
    Return True
  End Function
  
  ''' <summary>
  ''' 作業項目に関するプロパティ値を取得する。
  ''' 取得する作業項目をインデックスで指定する。
  ''' 存在しないインデックスを渡した場合はすべて空文字がセットされた構造体を返す。
  ''' </summary>
  ''' <param name="index"></param>
  ''' <returns></returns>
  Public Function GetWorkItemParams(index As Integer) As WorkItemParams
    If index < 1 Then Throw New ArgumentException("インデックスは１以上を指定してください。 / " + index.ToString)
    
    Dim params As WorkItemParams
    With params
      .Id           = index
      .Name         = GetValue(KEY_ITEM_NAME     + index.ToString).GetOrDefault(String.Empty)
      .WorkCountCol = GetAndCheckCol(KEY_WORKCOUNT_COL + index.ToString)
      .WorkTimeCol  = GetAndCheckCol(KEY_WORKTIME_COL  + index.ToString)
    End With
    
    Return params
  End Function
  
  ''' <summary>
  ''' Excelのファイルパスを返す。
  ''' </summary>
  ''' <returns></returns>
  Public Function ExcelFilePath() As String
    Dim dir As String  = GetValue(KEY_EXCEL_FILEDIR).GetOrDefault(String.Empty)
    Dim file As String = GetValue(KEY_EXCEL_FILENAME).GetOrDefault(String.Empty)
    Return Path.Combine(dir, file)
  End Function
  
  ''' <summary>
  ''' Excelのシート名を返す。
  ''' </summary>
  ''' <returns></returns>
  Public Function SheetName(month As Integer) As String
    Dim sheet As String = GetValue(KEY_EXCEL_SHEETNAME).GetOrDefault(String.Empty)
    Return String.Format(sheet, month.ToString)
  End Function
  
  ''' <summary>
  ''' データアクセスされるExcelの１行目を返す。
  ''' </summary>
  ''' <returns></returns>
  Public Function FirstRow() As Integer
    Dim v As String = GetValue(KEY_FIRST_ROW).GetOrDefault(String.Empty)
    Dim row As Integer
    If Not Integer.TryParse(v, row) Then
      Throw New InvalidDataException("プロパティ" & KEY_FIRST_ROW & "の値が不正です。 / " & v)
    End If
    
    Return row
  End Function
  
  ''' <summary>
  ''' データを記録する開始日を返す。
  ''' </summary>
  ''' <returns></returns>
  Public Function BeginDate() As DateTime
    Dim v As String = GetValue(KEY_BEGIN_DATE).GetOrDefault(String.Empty)
    Dim d As DateTime
    If Not DateTime.TryParse(v, d) Then
      Throw New InvalidDataException("プロパティ" & KEY_BEGIN_DATE & "の値が不正です。 / " & v)
    End If
    
    Return d
  End Function
  
  ''' <summary>
  ''' データを記録する終了日を返す。
  ''' </summary>
  ''' <returns></returns>
  Public Function EndDate() As DateTime
    Dim v As String = GetValue(KEY_END_DATE).GetOrDefault(String.Empty)
    Dim d As DateTime
    If Not DateTime.TryParse(v, d) Then
      Throw New InvalidDataException("プロパティ" & KEY_END_DATE & "の値が不正です。 / " & v)
    End If
    
    Return d
  End Function
  
  ''' <summary>
  ''' 作業項目ノートの名前を返す。
  ''' </summary>
  ''' <returns></returns>
  Public Function NoteName() As String
    Return GetValue(KEY_NOTE_NAME).GetOrDefault(String.Empty)
  End Function
  
  ''' <summary>
  ''' 作業項目ノートの列を返す。
  ''' </summary>
  ''' <returns></returns>
  Public Function NoteCol() As String
    Return GetAndCheckCol(KEY_NOTE_COL)
  End Function
  
  ''' <summary>
  ''' 出勤日の列を返す。
  ''' </summary>
  Public Function WorkDayCol() As String
    Return GetAndCheckCol(KEY_WORKDAY_COL)
  End Function
  
  ''' <summary>
  ''' 指定したプロパティの値を取得する。
  ''' その値がExcelの列の要件を満たした値ならそのまま返す。
  ''' そうでない場合は空文字を返す。
  ''' </summary>
  ''' <param name="key"></param>
  ''' <returns></returns>
  Private Function GetAndCheckCol(key As String) As String
    Dim col As String = GetValue(key).GetOrDefault(String.Empty)
    If Not Cell.ValidColumn(col) Then
      col = String.Empty
    End If
    
    Return col
  End Function
End Class
