Attribute VB_Name = "Module1"
Sub Main()
Dim w As Integer
Dim ws_num As Integer

Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
ws_num = ThisWorkbook.Worksheets.Count

For w = 1 To ws_num
   ThisWorkbook.Worksheets(w).Activate
   Call Calc
   Call ConFormatt
Next w

starting_ws.Activate 'activate the worksheet that was originally active
End Sub


Sub Calc()

    Dim ticker As String
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Dim volume As Double
    volume = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    'declaring opening price and setting to first value
    Dim OpenPrice As Double
    OpenPrice = Range("C2").Value
    Dim ClosePrice As Double
    Dim Math1 As Double
    Dim YearChange As Double
    Dim HighPrice As Double
    HighPrice = Range("D2").Value
    Dim LowPrice As Double
    Dim PercentChange As Double

'Setup the sheet
Range("J1").Value = "Ticker"
Range("L1").Value = "Percent Change"
Range("K1").Value = "Yearly Change"
Range("M1").Value = "Stock Total Value"

  For i = 2 To LastRow
    
    ' Check if we are still within the same credit card brand, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      ticker = Cells(i, 1).Value
      
     
      'Calc the percent change
      ClosePrice = Cells(i, 6).Value
      Math1 = OpenPrice - ClosePrice
      PercentChange = Math1 / OpenPrice
      
      'Calc the yearly change
      YearChange = Math1
            
    
      ' Add to the Brand Total
      volume = volume + Cells(i, 7).Value

      ' Print the Credit Card Brand in the Summary Table
      Range("J" & Summary_Table_Row).Value = ticker

      ' Print the Brand Amount to the Summary Table
      Range("M" & Summary_Table_Row).Value = volume

      ' Print the Percent Change to the Summary Table
      Range("K" & Summary_Table_Row).Value = YearChange
      
      ' Print the Year Change to the Summary Table
      Range("L" & Summary_Table_Row).Value = PercentChange
      Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
      
      'Yes, I am formatting the column width becuase I that much of an type A personality
      Range("J" & Summary_Table_Row).ColumnWidth = 15
      Range("K" & Summary_Table_Row).ColumnWidth = 15
      Range("L" & Summary_Table_Row).ColumnWidth = 15
      Range("M" & Summary_Table_Row).ColumnWidth = 15
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      volume = volume + Cells(i, 7).Value

    End If

  Next i

    
    
End Sub

Sub ConFormatt()
Dim c As Double
Dim LastL As Double

LastL = Range("L" & Rows.Count).End(xlUp).Row
For c = 2 To LastL
    If Cells(c, 11).Value > 0 Then
        Cells(c, 11).Interior.ColorIndex = 4
        Cells(c, 11).NumberFormat = "0.000000000"
    ElseIf Cells(c, 11).Value < 0 Then
        Cells(c, 11).Interior.ColorIndex = 3
        Cells(c, 11).NumberFormat = "0.000000000"
    ElseIf IsEmpty(Cells(c, 11)) Then
        Cells(c, 11).Interior.ColorIndex = -4142
        Cells(c, 11).NumberFormat = "0.000000000"
    End If
Next c
End Sub

'So I am so glad were moving on past VBA.  I wish I had learned VBA before learning Python.
'It would have made me a stronger coder.  Learning Python first just made this painful.
'If this wouldn't have been so painful I would have done the hard part of the lesson.
'At some point I just came to a point where I was over VBA.  Sorry. #quiting
