# VBA-Challenge
homework 2







Sub StockData()

Dim ws As Worksheet
Dim ticker As String
Dim vol As Integer
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer

'Ticker Symbol

 Dim Ticker_Name As String
        Ticker_Name = " "

 Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        
     ' Set CurrentWs as a worksheet object
    
    Dim CurrentWs As Worksheet
    Dim Need_Summary_Table_Header As Boolean
    Dim COMMAND_SPREADSHEET As Boolean
    
    Need_Summary_Table_Header = False
    COMMAND_SPREADSHEET = True
    
    ' Loop through all of the worksheets
    
    For Each CurrentWs In Worksheets
    

        ' Set an initial variable for holding the total per ticker name
       
        Total_Ticker_Volume = 0
        
        ' new variables
        
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Delta_Price As Double
        Delta_Price = 0
        Dim Delta_Percent As Double
        Delta_Percent = 0
         
        ' Keep track of the location for each ticker name in summary
        
        Summary_Table_Row = 2
        
        ' Set initial row count for the current worksheet
        
        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

        ' For all worksheet except the first one, the Results
        
        If Need_Summary_Table_Header Then
        
        
            ' Set Titles for the Summary Table for current worksheet
            
            CurrentWs.Range("I1").Value = "Ticker"
            CurrentWs.Range("J1").Value = "Yearly Change"
            CurrentWs.Range("K1").Value = "Percent Change"
            CurrentWs.Range("L1").Value = "Total Stock Volume"
            
            ' Set Titles for new Summary Table on the right for current worksheet
            
            CurrentWs.Range("O2").Value = "Greatest % Increase"
            CurrentWs.Range("O3").Value = "Greatest % Decrease"
            CurrentWs.Range("O4").Value = "Greatest Total Volume"
            CurrentWs.Range("P1").Value = "Ticker"
            CurrentWs.Range("Q1").Value = "Value"
        Else
        
            'reset flag for the rest of worksheets
            
            Need_Summary_Table_Header = True
        End If
        
        ' The rest ticker's open price will be initialized within the for loop below
        
        Open_Price = CurrentWs.Cells(2, 3).Value
        
        ' Loop from the beginning of the current worksheet(Row2) till its last row
        
        For i = 2 To Lastrow
        
      Next i
      
            ' Check if within the same ticker name,
            ' if not - write results to summary table
            
            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
            
            
                ' Set the ticker name
                
                Ticker_Name = CurrentWs.Cells(i, 1).Value
               
                Close_Price = CurrentWs.Cells(i, 6).Value
                Delta_Price = Close_Price - Open_Price
        
                If Open_Price <> 0 Then
                    Delta_Percent = (Delta_Price / Open_Price) * 100
                Else
                
                    ' Check
                    MsgBox ("For " & Ticker_Name & ", Row " & CStr(i) & ": Open Price =" & Open_Price & ". Fix <open> field manually and save the spreadsheet.")
                End If
                
                ' Add to the Ticker name total volume
                
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
                
                ' Print the Ticker Name in the Summary Table
                
                CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                CurrentWs.Range("J" & Summary_Table_Row).Value = Delta_Price
                
                ' Fill "Yearly Change"
    
                
                If (Delta_Price > 0) Then
                
                    'Fill Colors red and green in Ticker column
                    
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Delta_Price <= 0) Then
                   
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                End If
                
        
End Sub
