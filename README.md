# Stock Analysis
I downloaded day-end stock data for every company listed on the NYSE from 1982 through last week. I wanted to analyze simple technical trading strategies (eg buy when price is below lower bollinger band, sell after it goes up), but so far, I have only had time to graph year-end results and calculate the simple moving average. Nonetheless, I've learned a lot about VBA and SQL Server.

### VBA
First, I used VBA to download the data as .csv files from Yahoo! finance, one symbol at a time. I got the symbols from a text file I downloaded separately.
```
Sub download_data_for(mySymbol As String)
    Dim filePath As String: filePath = "D:\data\"
    Let baseURL = "http://ichart.finance.yahoo.com/table.csv?d=4&e=27&f=2016&a=0&b=1&c=1982&s="
    
    Dim myURL As String
    Dim otherURL As String
    Dim otherSymbol As String
    
    'if there's a period in symbol, try it with a forward slash as well
    If InStr(mySymbol, ".") Then
        Let otherSymbol = Replace(mySymbol, ".", "/")
    End If
    If Len(otherSymbol) > 0 Then
        Let otherURL = baseURL & otherSymbol
    End If
        
    Let myURL = baseURL & mySymbol

    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.send
    WinHttpReq.WaitForResponse
    DoEvents 'Wish it could be async, but this keeps it from locking
    
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile filePath & mySymbol & ".csv", 2
        oStream.Close
    
    ElseIf Len(otherURL) > 0 Then
        WinHttpReq.Open "GET", otherURL, False
        WinHttpReq.send
        WinHttpReq.WaitForResponse

        If WinHttpReq.Status = 200 Then
            Set oStream = CreateObject("ADODB.Stream")
            oStream.Open
            oStream.Type = 1
            oStream.Write WinHttpReq.responseBody
            oStream.SaveToFile filePath & mySymbol & ".csv", 2
            oStream.Close
        End If
    
    End If
    
End Sub

Sub download_all_data()
    Dim hf As Integer: hf = FreeFile
    Dim lines() As String, i As Long
    Dim symbols() As String
    Dim tempArray() As String
    
    Open "D:\New Downloads\NYSE.txt" For Input As #hf
        lines = Split(Input$(LOF(hf), #hf), vbNewLine)
    Close #hf
    
    'start at one to skip header
    ReDim symbols(1 To UBound(lines))
    For i = 1 To UBound(lines)
        tempArray = Split(lines(i), vbTab)
        If Len(Join(tempArray)) > 0 Then
            symbols(i) = tempArray(0)
        End If
    Next
    
    For i = 1 To UBound(symbols)
        download_data_for (symbols(i))
    Next
    
End Sub
```
The data wasn't very useful in all those separate files, so I decided to download the data directly into an Excel workbook. There were so many rows that I had to split it into 11 worksheets.
```
Sub insert_data_for(mySymbol As String)
    Let baseURL = "http://ichart.finance.yahoo.com/table.csv?d=4&e=27&f=2016&a=0&b=1&c=1982&s="
    
    Dim myURL As String
    Dim otherURL As String
    Dim otherSymbol As String
    Dim lines() As String
    Dim rows() As String
    Dim currentRow As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    currentRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    
    'Ran out of space, create new sheet
    If currentRow > 950000 Then
        ws.Name = ws.Cells(2, 1) & "-" & ws.Cells(currentRow, 1)
        Set ws = ThisWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count))
        ws.Name = "Sheet1"
        Let header = "Symbol,Date,Open,High,Low,Close,Volume,Adj Close"
        ws.Range(Cells(1, 1), Cells(1, 8)).Value = Split(header, ",")
        currentRow = 1
    End If
    
        
    
    
    
    'if there's a period in symbol, try it with a forward slash as well
    If InStr(mySymbol, ".") Then
        Let otherSymbol = Replace(mySymbol, ".", "/")
    End If
    If Len(otherSymbol) > 0 Then
        Let otherURL = baseURL & otherSymbol
    End If
        
    Let myURL = baseURL & mySymbol

    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.send
    WinHttpReq.WaitForResponse
    DoEvents 'Wish it could be async, but this keeps it from locking
    
    If WinHttpReq.Status = 200 Then
        'parse WinHttpReq.responseBody
        lines = Split(WinHttpReq.responseText, vbLf)
        For i = 1 To UBound(lines) - 1 'Last line is empty
                ws.Cells(i + currentRow, 1).Value = mySymbol
                'Breaks here if sheet one is not selected
                ws.Range(Cells(i + currentRow, 2), Cells(i + currentRow, 8)).Value = Split(lines(i), ",")
        Next
                    
            
    
    ElseIf Len(otherURL) > 0 Then
        WinHttpReq.Open "GET", otherURL, False
        WinHttpReq.send
        WinHttpReq.WaitForResponse

        If WinHttpReq.Status = 200 Then
            lines = Split(WinHttpReq.responseText, vbLf)
            For i = 1 To UBound(lines) - 1 'Last line is empty
                    ws.Cells(i + currentRow, 1).Value = mySymbol
                    'Breaks here if sheet one is not selected
                    ws.Range(Cells(i + currentRow, 2), Cells(i + currentRow, 8)).Value = Split(lines(i), ",")
            Next
        End If
    
    End If
End Sub


Sub insert_all_data()
    Dim hf As Integer: hf = FreeFile
    Dim lines() As String, i As Long
    Dim symbols() As String
    Dim tempArray() As String
    Dim header As String: header = "Symbol,Date,Open,High,Low,Close,Volume,Adj Close"
    
    Open "D:\New Downloads\NYSE.txt" For Input As #hf
        lines = Split(Input$(LOF(hf), #hf), vbNewLine)
    Close #hf
    
    'start at one to skip header
    ReDim symbols(1 To UBound(lines))
    For i = 1 To UBound(lines)
        tempArray = Split(lines(i), vbTab)
        If Len(Join(tempArray)) > 0 Then
            symbols(i) = tempArray(0)
        End If
    Next
    
    ThisWorkbook.Sheets("Sheet1").Range(Cells(1, 1), Cells(1, 8)).Value = Split(header, ",")
    For i = 1 To UBound(symbols)
        insert_data_for symbols(i)
    Next
  
End Sub
```
With the data finally in the spreadsheet, it was time to do some analysis! I created an "Analysis" sheet with a line graph chart to show the year-end close for an inputted symbol. The chart updates for a new symbol when you click a button. In order to get reasonable performance searching through such a large workbook, I had to use a binary search.
```
Sub update_year_end()
    Dim mySymbol As String: mySymbol = ThisWorkbook.Sheets("Analysis").Range("$E$1").Value
    Dim currentSymbol As String
    Dim yearEndAdjClose(34, 0) As Double
    Dim year(34, 0) As String
    Dim currentYear As String: currentYear = "2016"
    Dim currentDate() As String
    Dim yearCount As Integer: yearCount = 0
    Dim done As Boolean: done = False
    
    ThisWorkbook.Sheets("Analysis").Range("$A$2:$B$36").Clear
    
    If Len(mySymbol) Then
    
    For i = 2 To ThisWorkbook.Sheets.Count                                                      'for each workbook besides first
        Dim lastRow As Long: lastRow = ThisWorkbook.Sheets(i).UsedRange.rows.Count
        Dim lastSymbol As String: lastSymbol = ThisWorkbook.Sheets(i).Cells(lastRow, 1).Value
        If StrComp(lastSymbol, mySymbol) > -1 Then                                              'look in sheet if mySymbol might be there
            Dim lower As Long: lower = 1
            Dim upper As Long: upper = lastRow
            Dim middle As Long: middle = upper / 2
            While lower < upper                                                                     'binary search
                'DoEvents
                Dim currentCheck As String: currentCheck = ThisWorkbook.Sheets(i).Cells(middle, 1).Value
                If StrComp(currentCheck, mySymbol) = -1 Then                                        'mySymbol > middle
                    lower = middle + 1
                    middle = (upper + lower) / 2
                ElseIf StrComp(currentCheck, mySymbol) = 1 Then                                     'mySymbol < middle
                    upper = middle - 1
                    middle = (upper + lower) / 2
                Else                                                                                'mySymbol = middle
                    lower = middle
                    upper = middle
                    While ThisWorkbook.Sheets(i).Cells(lower, 1).Value = mySymbol
                        lower = lower - 1
                    Wend
                    While ThisWorkbook.Sheets(i).Cells(upper, 1).Value = mySymbol
                        upper = upper + 1
                    Wend
                    lower = lower + 1
                    upper = upper - 1
                
                    For j = lower To upper
                        currentDate = Split(ThisWorkbook.Sheets(i).Cells(j, 2).Value, "-")
                        If Not currentYear = currentDate(0) Then                                        'if it's a new year
                            yearEndAdjClose(yearCount, 0) = CDbl(ThisWorkbook.Sheets(i).Cells(j, 8))    'get the adj close as a double
                            year(yearCount, 0) = ThisWorkbook.Sheets(i).Cells(j, 2)
                            currentYear = currentDate(0)
                            yearCount = yearCount + 1
                        End If
                    Next
                    GoTo finish                                                                     'To break loop, I guess
                End If
            Wend
            ThisWorkbook.Sheets("Analysis").Range("$A$2").Value = "Not"
            ThisWorkbook.Sheets("Analysis").Range("$A$3").Value = "Found"
            Exit For
        End If
    Next
Exit Sub
finish:
    Dim dest As Range: Set dest = ThisWorkbook.Sheets("Analysis").Range("$A$2:$A$" & CStr(yearCount + 1))
    dest.Value = year
    Set dest = ThisWorkbook.Sheets("Analysis").Range("$B$2:$B$" & CStr(yearCount + 1))
    dest.Value = yearEndAdjClose
    
    End If
   
    
End Sub
```
### SQL Server
Obviously, Excel wasn't very happy with nearly 11 million rows of data in a single workbook, so I imported the data to SQL server using the SSIS Import wizard. I could have written a VBA script to do it (or, better yet, skip the spreadsheet and just download it directly into the database), but I figured I'd see if the wizard was useful. It worked OK, but it forced me to import one sheet at a time, since no two sheets could target the same table in the database on a single import.

Here's the T-SQL for my table showing the data types I used. I added the primary key later using Visual Studio 2015 to guard against duplicate entries.
```
CREATE TABLE [dbo].[HData] (
    [Symbol]    NVARCHAR (6)    NOT NULL,
    [Date]      SMALLDATETIME   NOT NULL,
    [Open]      DECIMAL (12, 6) NOT NULL,
    [High]      DECIMAL (12, 6) NOT NULL,
    [Low]       DECIMAL (12, 6) NOT NULL,
    [Close]     DECIMAL (12, 6) NOT NULL,
    [Volume]    BIGINT          NOT NULL,
    [Adj Close] DECIMAL (12, 6) NOT NULL,
    CONSTRAINT [PK_HData] PRIMARY KEY CLUSTERED ([Date] ASC, [Symbol] ASC)
);
```

With the data in SQL Server, it was quick and easy to look up data such as the highest volume trading days.
```
SELECT Top 10 
Date, SUM(Volume) AS 'Total Volume'
FROM HData
GROUP BY (Date)
ORDER BY Sum(Volume) Desc 
;
```
Results :
```
Date	Total Volume
2008-10-10 00:00:00	9391007200
2011-08-08 00:00:00	8939502600
2011-08-09 00:00:00	8370107400
2010-05-06 00:00:00	8007088000
2008-09-18 00:00:00	8001677500
2011-08-05 00:00:00	7782210300
2011-08-10 00:00:00	7649403200
2010-05-07 00:00:00	7622576700
2008-11-21 00:00:00	7461952900
2008-10-08 00:00:00	7376337400
```
But I wanted more in-depth information like technical trading indicators, and after a little research, I decided that calculated columns would be a good way to do that. However, to do any type of advanced calculations with calculated columns you need to use a user-defined function. I used this one to help define a simple moving average.
```
CREATE FUNCTION dbo.Last14Avg(@curdate datetime, @cursymbol nvarchar(6))
RETURNS decimal(12, 6)
BEGIN
	DECLARE @avg AS decimal(12, 6)
	
	SELECT @avg = AVG([Adj Close]) 
	FROM (
		SELECT Top 14 [Adj Close]
		FROM dbo.HData
		WHERE Symbol = @cursymbol
		AND Date <= @curdate
		ORDER BY Date Desc
	) recentDays

	RETURN @avg
END
; 
```
Then I added the calculated column with the following T-SQL. I wanted to make it persistent (to improve performance), but SQL Server wouldn't let me because it thinks the data is calculated dynamically.
```
ALTER TABLE [dbo].[HData]
ADD SMA14 AS dbo.Last14Avg(Date, Symbol) ;
```
Now I can easily find the simple moving average for any row in the database.
```
SELECT Top 25 Symbol, Date, [Adj Close], SMA14
FROM HData
WHERE Symbol = 'BRK.A' 
ORDER BY Date Desc ;
```
Results :
```
Symbol	Date	Adj Close	SMA14
BRK.A	2016-05-27 00:00:00	214303.000000	213630.642857
BRK.A	2016-05-26 00:00:00	214600.000000	213639.714285
BRK.A	2016-05-25 00:00:00	216185.000000	213811.142857
BRK.A	2016-05-24 00:00:00	215650.000000	213789.357142
BRK.A	2016-05-23 00:00:00	212210.000000	213851.142857
BRK.A	2016-05-20 00:00:00	212833.000000	214218.571428
BRK.A	2016-05-19 00:00:00	211180.000000	214748.500000
BRK.A	2016-05-18 00:00:00	212200.000000	215307.071428
BRK.A	2016-05-17 00:00:00	210695.000000	215801.357142
BRK.A	2016-05-16 00:00:00	212530.000000	216568.142857
BRK.A	2016-05-13 00:00:00	212140.000000	217126.000000
BRK.A	2016-05-12 00:00:00	214860.000000	217635.857142
BRK.A	2016-05-11 00:00:00	214468.000000	217934.285714
BRK.A	2016-05-10 00:00:00	216975.000000	218148.000000
BRK.A	2016-05-09 00:00:00	214430.000000	218287.500000
```
### Power Query, Power Pivot, and Power BI
These look like great tools for visualizing the data. Unfortunately, I have not had the time to learn DAX, yet.
