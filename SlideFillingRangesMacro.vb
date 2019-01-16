git remote add origin git@github.com:Betanya/RangesForSlideFilingLab.git
git push -u origin master
Sub Drwer()
Dim MaxRowsPerDrawer As Integer
 
Dim StartRow As Integer
 
Dim TotalRows As Integer
 
Dim CurrentRow As Integer

Dim StrN As String

Dim StrNpp As String

Dim IDN As String

Dim IDNpp As String

Dim StrPrev As String
Dim IDNPrev As String

Dim PrintRow As Integer
Dim RowsLeft As Integer


MaxRowsPerDrawer = 330

StartRow = 2

PrintRow = 2

TotalRows = ActiveSheet.UsedRange.Rows.Count
RowsLeft = TotalRows
CurrentRow = StartRow
'Basic Operation. Function begins at StartRow'
'   Function then jumps 320 rows forward'
'   Function then checks if the current row contains the same ID as the next row'
'       if No, the function ends the Drawer and starts again using the next row as a starting position'
'       if Yes, the function moves backwards by one row and repeats the check'
'           this repeats until the answer is No'
' '
'By this method we ensure that the maximum possible number of items go into one drawer before moving onto the next'
'while also preventing more than the maximum number of items from ending up in a drawer'
 'Also it will not make the starting or ending value of the drawer if there is not at least 330 slides left'
While CurrentRow < TotalRows And RowsLeft > MaxRowsPerDrawer
    'Mark the current row as the beginning of a drawer'
   Cells(PrintRow, 10) = "StartHalfDrwer"
   Cells(PrintRow, 11) = Range("A" & (CurrentRow)).Value
   PrintRow = PrintRow + 1
    'Jump forward to check if we have a perfect drawer'
    'Add a number to the starting Row so it will be a new column'
   CurrentRow = CurrentRow + MaxRowsPerDrawer
    'Isolate the String values of the nth and n+1th cell'
    RowsLeft = TotalRows - CurrentRow
   StrN = Range("A" & (CurrentRow)).Value
    StrNpp = Range("A" & (CurrentRow + 1)).Value
    'Strip out the numbers that we care about. These are still a string. VJ00041-A2 would turn into 000041'
   IDN = Mid(StrN, 3, 5)
    IDNpp = Mid(StrNpp, 3, 5)
    'Comparison is simpler now that we have these named variables'
 
    If IDN <> IDNpp Then
        'Since the Beginning+320th row does not match the Beginning+321st, we have a perfect drawer and can move on'
       Cells(PrintRow, 10) = "EndHalfDrawer"
       Cells(PrintRow, 11) = Range("A" & (CurrentRow)).Value
         PrintRow = PrintRow + 1
        CurrentRow = CurrentRow + 1
 
    Else
        'Same operations as before, only this time we are moving up the rows until we find a mismatch'
       Do While True
            'Grab the number out of the previous row.'
           CurrentRow = CurrentRow - 1
            StrPrev = Range("A" & (CurrentRow)).Value
            IDNPrev = Mid(StrPrev, 3, 5)
 
            'Continue until we find a mismatch'
           If IDN <> IDNPrev Then
                Exit Do
            End If
 
        Loop
        'Exit Do brings us here, where we mark the end of the drawer and continue back to the start of the big loop'
       Cells(PrintRow, 10) = "EndHalfDrwer"
       Cells(PrintRow, 11) = Range("A" & (CurrentRow)).Value
       PrintRow = PrintRow + 1
        CurrentRow = CurrentRow + 1
 
    End If
 
Wend
End Sub

