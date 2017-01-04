'
' 日付: 2017/01/04
'
Imports Common.Util

''' <summary>
''' DateTermの拡張メソッド群
''' </summary>
Public Module DateTermExtensions
  
  ''' <summary>
  ''' 期間を１週間単位で区切って返す。
  ''' 区切った期間にはラベルを付与する。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function LabelingWeeklyTerms(term As DateTerm) As IEnumerable(Of DateTerm)
    ' 週末の曜日
    Dim dayOfWeekEnd As DayOfWeek = DayOfWeek.Saturday
    
	  ' 週はじめと週末の日付を取得して、月の第何週かを表す文字列を返す関数
	  Dim funcToGetWeekCount As Func(Of DateTime, DateTime, String) =
	    Function(weekStart, weekEnd)
	      Dim cnt As Integer = DateUtils.GetWeekCountInMonth(weekStart, dayOfWeekEnd)
	      Dim str As String  = String.Format("{0}月第{1}週", weekStart.Month, cnt)
	      
        If weekStart.Month <> weekEnd.Month Then
          str = str & String.Format("/{0}月第1週", weekEnd.Month)
        End If
        
        Return str
      End Function
	  
	  ' 期間を週単位で区切ったリストを取得する
    Return term.WeeklyTerms(dayOfWeekEnd, funcToGetWeekCount)    
  End Function
  
  ''' <summary>
  ''' 期間を１ヶ月単位で区切って返す。
  ''' 区切った期間にはラベルを付与する。
  ''' </summary>
  <System.Runtime.CompilerServices.ExtensionAttribute()>
  Public Function LabelingMonthlyTerms(term As DateTerm) As IEnumerable(Of DateTerm)
    Return term.MonthlyTerms(Function(begin, _end) begin.Month.ToString & "月")
  End Function
End Module
