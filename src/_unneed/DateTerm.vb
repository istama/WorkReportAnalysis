''
'' 日付: 2016/07/15
''
'Imports Common.Util
'
'Public Structure DateTerm
'	Private ReadOnly _min As DateTime
'	Public ReadOnly Property Min As DateTime
'		Get
'			Return _min
'		End Get
'	End Property
'	
'	Private ReadOnly _max As DateTime
'	Public ReadOnly Property Max As DateTime
'		Get
'			Return _max
'		End Get
'	End Property
'	
'	Private _dailyLabelAndTermList As List(Of LabelAndDateTerm)
'	Public ReadOnly Property DailyLabelAndTermList As List(Of LabelAndDateTerm)
'		Get
'		If _dailyLabelAndTermList Is Nothing Then
'				InitDailyLabelAndTermList()
'			End If
'			Return copy(_dailyLabelAndTermList)
'		End Get
'	End Property
'	
'	Public ReadOnly Property DailyLabelList As List(Of String)
'		Get
'			Return zipList(Me.DailyLabelAndTermList).Item1		
'		End Get
'	End Property
'	
'	Public ReadOnly Property DailyTermList As List(Of DateTerm)
'		Get
'			Return zipList(Me.DailyLabelAndTermList).Item2
'		End Get
'	End Property
'		
'	Private _weeklyLabelAndTermList As List(Of LabelAndDateTerm)
'	Public ReadOnly Property WeeklyLabelAndTermList As List(Of LabelAndDateTerm)
'		Get
'			If _weeklyLabelAndTermList Is Nothing Then
'				InitWeeklyLabelAndTermList()
'			End If
'			Return copy(_weeklyLabelAndTermList)
'		End Get
'	End Property
'	
'	Public ReadOnly Property WeeklyLabelList As List(Of String)
'		Get
'			Return zipList(Me.WeeklyLabelAndTermList).Item1		
'		End Get
'	End Property
'	
'	Public ReadOnly Property WeeklyTermList As List(Of DateTerm)
'		Get
'			Return zipList(Me.WeeklyLabelAndTermList).Item2
'		End Get
'	End Property
'	
'	Private _monthlyLabelAndTermList As List(Of LabelAndDateTerm)
'	Public ReadOnly Property MonthlyLabelAndTermList As List(Of LabelAndDateTerm)
'		Get
'			If _monthlyLabelAndTermList Is Nothing Then
'				InitMonthlyLabelAndTermList()
'			End If
'			Return copy(_monthlyLabelAndTermList)
'		End Get
'	End Property
'	
'	Public ReadOnly Property MonthlyLabelList As List(Of String)
'		Get
'			Return zipList(Me.MonthlyLabelAndTermList).Item1
'		End Get
'	End Property
'	
'	Public ReadOnly Property MonthlyTermList As List(Of DateTerm)
'		Get
'			Return zipList(Me.MonthlyLabelAndTermList).Item2
'		End Get
'	End Property
'	
'	Private ReadOnly _dayOfWeekend As DayOfWeek
'	
'	Public Sub New(min As DateTime, max As DateTime, Optional dayOfWeekend As DayOfWeek = DayOfWeek.Sunday)
'		Me._min = min
'		Me._max = max
'		Me._dailyLabelAndTermList   = Nothing
'		Me._weeklyLabelAndTermList  = Nothing
'		Me._monthlyLabelAndTermList = Nothing
'		Me._dayOfWeekend = dayOfWeekend
'	End Sub
'	
'	''' <summary>
'	''' 期間中の日付のリストを返す。
'	''' </summary>
'	Public Function DateListInTerm() As List(Of DateTime)
'		Dim l As New List(Of DateTime)
'		Dim d As DateTime = Me._min
'		While d <= Me._max
'			l.Add(d)
'			d = d.AddDays(1)
'		End While
'		Return l
'	End Function
'	
'	''' <summary>
'	''' 期間とそれを表す文字列のリストを初期化する。
'	''' 期間（DateTerm）は１日単位で分割されるので、DateTerm.MinとDateTerm.Maxは同じ値がセットされる。
'	''' </summary>
'	Private Sub InitDailyLabelAndTermList()
'		Me._dailyLabelAndTermList = New List(Of LabelAndDateTerm)
'		
'		Dim d As DateTime = Me._min
'		While d <= Me._max
'			Dim label As String = String.Format("{0:00}日", d.Day)
'			Dim term As DateTerm = New DateTerm(d, d, Me._dayOfWeekend)
'			Me._dailyLabelAndTermList.Add(New LabelAndDateTerm(label, term))
'			d = d.AddDays(1)
'		End While
'	End Sub
'	
'	''' <summary>
'	''' 週の期間とそれを表す文字列のリストを初期化する。
'	''' </summary>
'	Private Sub InitWeeklyLabelAndTermList()
'		Me._weeklyLabelAndTermList = New List(Of LabelAndDateTerm)
'		
'		Dim beginDate As DateTime = Me._min   ' 週の開始日
'		Dim weekCntInMonth As Integer = 1  		' １ヶ月の中での週をカウント
'		
'		While beginDate < Me._max
'			' 週の開始日と終了日のセットを生成
'			Dim endDate As DateTime = DateUtils.GetDateOfNextWeekDay(beginDate, Me._dayOfWeekend)
'			If endDate > Me._max Then
'				endDate = Me._max
'			End If
'			Dim term As DateTerm = New DateTerm(beginDate, endDate, Me._dayOfWeekend)	' 選択した要素から取得できる値
'			
'			' この週を表す文字列を生成
'			Dim label As String
'			' 週の終了日が翌月にまたがない場合
'			If endDate.Month = beginDate.Month Then
'				label = String.Format("{0:00}月第{1}週", beginDate.Month, weekCntInMonth.ToString)
'				' 終了日が月末日の場合
'				If endDate.Day = DateTime.DaysInMonth(endDate.Year, endDate.Month) Then
'					weekCntInMonth = 1 ' 翌月の第１週からカウントを開始
'				Else
'					weekCntInMonth += 1 ' 週のカウントを加算
'				End If
'			Else
'				label = String.Format("{0:00}月第{1}週/{2:00}月第1週", beginDate.Month, weekCntInMonth.ToString, endDate.Month)
'				weekCntInMonth = 2	' 翌月の第２週からカウントを開始
'			End If
'			
'			Me._weeklyLabelAndTermList.Add(New LabelAndDateTerm(label, term))
'			
'			' 週の開始日を更新
'			beginDate = endDate.AddDays(1)
'		End While		
'	End Sub
'	
'	''' <summary>
'	''' 月の期間とそれを表すリストを初期化する。
'	''' </summary>
'	Private Sub InitMonthlyLabelAndTermList()
'		Me._monthlyLabelAndTermList = New List(Of LabelAndDateTerm)
'		
'		For Each d In DateUtils.GetDateListOfEveryMonth(Me._min, Me._max)
'			' 月の開始日と月末日のセットを作成
'			Dim beginDate As New DateTime(d.Year, d.Month, 1)
'			If d.Year = Me._min.Year AndAlso d.Month = Me._min.Month Then
'				beginDate = Me._min
'			End If
'			
'			Dim endDate As New DateTime(d.Year, d.Month, DateTime.DaysInMonth(d.Year, d.Month))
'			If d.Year = Me._max.Year AndAlso d.Month = Me._max.Month Then
'				endDate = Me._max
'			End If
'			
'			Dim term As DateTerm = New DateTerm(beginDate, endDate, Me._dayOfWeekend)
'			Dim label As String = String.Format("{0:00}月", d.Month)
'			
'			Me._monthlyLabelAndTermList.Add(New LabelAndDateTerm(label, term))
'		Next		
'	End Sub
'	
'	''' <summary>
'	''' リストのコピーを生成して返す。
'	''' 防御的コピー用のメソッド。
'	''' </summary>
'	Private Function copy(list As List(Of LabelAndDateTerm)) As List(Of LabelAndDateTerm)
'		Dim l As New List(Of LabelAndDateTerm)
'		list.ForEach(Sub(t) l.Add(t))
'		Return l
'	End Function
'	
'	''' <summary>
'	''' ラベルと期間のリストをそれぞれ別のリストに分割し、タプルにセットして返す。
'	''' </summary>
'	Private Function zipList(list As List(Of LabelAndDateTerm)) As Tuple(Of List(Of String), List(Of DateTerm))
'		Dim strList As New List(Of String)
'		Dim termList As New List(Of DateTerm)
'		list.ForEach(
'			Sub(t)
'				strList.Add(t.Label)
'				termList.Add(t.Term)
'			End Sub)
'		
'		Return New Tuple(Of List(Of String), List(Of DateTerm))(strList, termList)
'	End Function
'	
'	Public Overrides Function ToString As String
'		Return String.Format("{0} - {1}", Me._min, Me._max)
'	End Function
'End Structure
'
'Public Structure LabelAndDateTerm
'	Private _label As String
'	Public ReadOnly Property Label As String
'		Get
'			Return _label
'		End Get
'	End Property
'	
'	Private _dateTerm As DateTerm
'	Public ReadOnly Property Term As DateTerm
'		Get
'			Return _dateTerm
'		End Get
'	End Property
'	
'	Public Sub New(label As String, term As DateTerm)
'		Me._label = label
'		Me._dateTerm = term
'	End Sub
'End Structure