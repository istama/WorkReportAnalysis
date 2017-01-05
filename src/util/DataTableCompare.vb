'
' 日付: 2017/01/05
'
Imports System.Data
Imports System.Linq

Imports Common.Extensions

''' <summary>
''' DataTableの行同士を指定した列で比較できるようにする。
''' </summary>
Public Class DataTableCompare
  ''' ソートしたテーブルの名前
  Private sortedTableName As String  = Nothing  
  ''' 比較した列
  Private comparedColIdx  As Integer = -1
  ''' 昇順でソートしたかどうか
  Private isAsc           As Boolean = True
  
  Public Sub New()
  End Sub
  
  ''' <summary>
  ''' 直近のソートした内容をリセットする。
  ''' </summary>
  Public Sub Reset()
    Me.sortedTableName = Nothing
    Me.comparedColIdx  = -1
    Me.isAsc           = True
  End Sub
  
  ''' <summary>
  ''' 比較用のオブジェクトを取得する。
  ''' </summary>
  Public Function GetDataRowCompare(sortedTableName As String, table As DataTable, comparedColIdx As Integer) As IComparer(Of DataRow)
    If sortedTableName Is Nothing Then Throw New ArgumentNullException("sortedTableName is null")
    If table           Is Nothing Then Throw New ArgumentNullException("table is null")
    
    If comparedColIdx <  0 OrElse
       comparedColIdx >= table.Columns.Count Then 
      Throw New IndexOutOfRangeException("comparedColIdx is range out / " & comparedColIdx.ToString)
    End If
    
    ' 前回、同じ列で昇順にソートしたかどうか
    If sortedTableName = Me.sortedTableName AndAlso comparedColIdx = Me.comparedColIdx AndAlso Me.isAsc Then
      ' した場合は降順でソートする
      Me.isAsc          = False
    Else
      ' していない場合は、今回のソート内容で上書きする
      Me.sortedTableName = sortedTableName
      Me.comparedColIdx  = comparedColIdx
      Me.isAsc           = True
    End If
    
    Return New DataRowCompare(Me.comparedColIdx, table.Columns(Me.comparedColIdx).DataType, Me.isAsc)
  End Function
  
  ''' <summary>
  ''' 行と行を指定した列で比較するためのクラス。
  ''' </summary>
  Private Class DataRowCompare
    Implements IComparer(Of DataRow)
    
    ''' 比較する列のインデックス。
    Private ReadOnly index As Integer
    ''' 比較する列の型。
    Private ReadOnly type  As Type
    ''' 昇順か降順か。 trueなら昇順。
    Private ReadOnly isAsc As Boolean
    
    Public Sub New(index As Integer, type As Type, isAsc As Boolean)
      Me.index = index
      Me.type  = type
      Me.isAsc = isAsc
    End Sub
    
    ''' <summary>
    ''' ２つの行を比較する
    ''' </summary>
    Public Function Compare(x As DataRow, y As DataRow) As Integer Implements IComparer(Of DataRow).Compare
      If Me.type = GetType(String) Then
        If Me.index = 0 Then
          ' 日付やIDを数値で比較する
          Return CmpRowByStrNum(x, y)          
        Else
          Return CmpRow(Of String)(x, y, String.Empty)        
        End If
      ElseIf type = GetType(Integer)
        Return CmpRow(Of Integer)(x, y, -1)      
      ElseIf type = GetType(Double)
        Return CmpRow(Of Double)(x, y, 1.0)      
      Else
        Throw New InvalidCastException("DataTableに無効な型の値があります。")
      End If
    End Function
    
    ''' <summary>
    ''' ２つの行を指定した型で比較する。
    ''' </summary>
    Private Function CmpRow(Of T As IComparable)(xRow As DataRow, yRow As DataRow, def As T) As Integer
      Dim xv As T = xRow.GetOrDefault(Of T)(Me.index, def)
      Dim yv As T = yRow.GetOrDefault(Of T)(Me.index, def)
      
      Return Cmp(xv, yv, def)
    End Function
    
    ''' <summary>
    ''' ２つの行を数字で比較する。
    ''' </summary>
    Private Function CmpRowByStrNum(xRow As DataRow, yRow As DataRow) As Integer
      Dim xv As Integer = HeadNumber(xRow.GetOrDefault(Me.index, String.Empty))
      Dim yv As Integer = HeadNumber(yRow.GetOrDefault(Me.index, String.Empty))
      
      Return Cmp(xv, yv, -1)
    End Function
    
    ''' <summary>
    ''' ２つの値を比較する。
    ''' </summary>
    Private Function Cmp(Of T As IComparable)(x As T, y As T, def As T) As Integer
      If x.Equals(y) Then
        Return 0
      ElseIf x.Equals(def)
        ' xがデフォルト値なら、昇順か降順かに関わらずxを後方にもっていく
        Return 1
      ElseIf y.Equals(def)
        ' yがデフォルト値なら、昇順か降順かに関わらずxを前方にもっていく
        Return -1
      End If
      
      Return DirectCast(IIf(Me.isAsc, x.CompareTo(y), x.CompareTo(y) * -1), Integer)      
    End Function
    
    ''' <summary>
    ''' 文字列の先頭に数字があれば、それを数値に変換して返す。
    ''' </summary>
    Private Function HeadNumber(text As String) As Integer
      If String.IsNullOrWhiteSpace(text) Then 
        Return -1
      End If
      
      ' 先頭から数字の文字列のみを取り出す
      Dim chars As IEnumerable(Of Char) =
        text.TakeWhile(Function(c) Asc(c) >= Asc("0"c) AndAlso Asc(c) <= Asc("9"c))
      
      If chars.Count = 0 Then 
        Return -1
      End If
      
      ' 数値に変換し、桁数をかける
      Dim nums As IEnumerable(Of Integer) =
        chars.Select(Function(c, idx) Integer.Parse(c) * CType(Math.Pow(10, chars.Count - idx - 1), Integer))
      
      ' 合計を返す
      Return nums.Sum
    End Function
  End Class
End Class

