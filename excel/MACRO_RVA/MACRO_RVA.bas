Attribute VB_Name = "ModuleRVA"
' MIT License
'
' Copyright (c) 2023 Giuseppe Ferri <jfinfoit@gmail.com>
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.


'
' Note:
' Tested on Excel 2016
'


' InsertAlternatingEmptyRows()
' Starting from the initialRow row, alternate this and subsequent numRows rows
' with blank rows using the same format as the previous rows.
' Empty rows are inserted from bottom to top, while rows below them are shifted down.
'
' @param initalRow As Integer The starting row number
' @param numRows As Integer The number of rows to alternate
' @version 1.0
' @author Joe Ferri
'
Sub InsertAlternatingEmptyRows(initalRow As Integer, numRows As Integer)
  Dim i As Integer
  
  ' disable automatic calculation to improve performance
  Application.Calculation = xlCalculationManual
  
  ' inserts alternating blank rows
  For i = initalRow + numRows - 1 To initalRow + 1 Step -1
      Rows(i & ":" & i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
  Next i
  
  ' copies the entire row, including formats, and inserts after the last row of data
  ' TODO: is it possible to obtain the same result by modifying the "for"?
  Rows((initalRow + numRows - 1) * 2 - 2).Copy
  Rows((initalRow + numRows - 1) * 2).Insert Shift:=xlDown
  
  ' delete temporary selection (necessary to avoid continuous copy error)
  Application.CutCopyMode = False
  
  ' enable automatic calculation again
  Application.Calculation = xlCalculationAutomatic
End Sub


' MACRO_RVA()
' Starting from the initialRow row, alternate this and subsequent numRows rows
' with blank rows using the same format as the previous rows.
' Empty rows are inserted from bottom to top, while rows below them are shifted down.
' The initialRow value is taken from the "OPTIONS" sheet in cell "D4".
' The numRows value is taken from the "OPTIONS" sheet in cell "D6".
'
' @version 1.0
' @author Joe Ferri
'
Sub MACRO_RVA()
  ' gets the starting row from cell D4 in the OPTIONS sheet
  Dim initalRow As Integer
  initalRow = Worksheets("OPTIONS").Range("D4").Value
  
  ' gets the number of rows to alternate from cell D6 in the OPTIONS sheet
  Dim numRows As Integer
  numRows = Worksheets("OPTIONS").Range("D6").Value
  
  ' calls the subroutine and passes the desired number of rows
  Call InsertAlternatingEmptyRows(initalRow, numRows)
End Sub

