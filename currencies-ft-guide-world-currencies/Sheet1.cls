VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Code to run when the button has been clicked
Private Sub CommandButton1_Click()

    Dim formattedWorksheet As String
    outputWorksheetName = "formatted"
    Call addHeader(outputWorksheetName)
    
    ' Regex for lines that have no currency data
    Set objSkipLineRegExp = CreateObject("vbscript.regexp")
    objSkipLineRegExp.Pattern = "Change|Week|GUIDE|:|\d/\d"

    ' Regex for Currency name
    Set objRegExp = CreateObject("vbscript.regexp")
    objRegExp.Pattern = "[a-zA-Z\(\)&]"
    
    ' Select the row we want to start on
    Range("A2").Select
    ' Set Do loop to stop when an empty cell is reached.
    Do Until IsEmpty(ActiveCell)
       ' Insert your code here.
       Call ProcessRow(outputWorksheetName, ActiveCell.row, ActiveCell.Column, objSkipLineRegExp, objRegExp)
       
       ' Step down 1 row from present location.
       ActiveCell.Offset(1, 0).Select
    Loop
    
End Sub

' Add the header to the formatted worksheet
Private Sub addHeader(worksheetName)
    Dim header(9) As String
    header(0) = "Country"
    header(1) = "Currency"
    header(2) = "£ STG"
    header(3) = "£ Week Change"
    header(4) = "US$"
    header(5) = "US$ Change"
    header(6) = "Euro"
    header(7) = "Euro Change"
    header(8) = "Yen"
    header(9) = "Yen Change"
    
    Dim headerCount As Integer
    headerCount = 0
    For Each element In header
        Sheets(worksheetName).Cells.Item(1, headerCount + 1).Value = header(headerCount)
        headerCount = headerCount + 1
    Next element
    
End Sub


' Process the currency data
Private Sub ProcessRowData(worksheetName, row, strRow, objRegExp)
    Dim rowArray() As String
    rowArray() = Split(strRow)
    
    Dim title As String
    Dim arrContent(7) As String
    Dim arrCount As Integer
    arrCount = 0
    
    For Each element In rowArray
        Set regExpMatches = objRegExp.Execute(element)
        
        ' check if cell is currency name or a value
        If regExpMatches.Count = 1 Then
            title = title & " " & element
        Else
            arrContent(arrCount) = element
            arrCount = arrCount + 1
        End If
    Next element

    Dim strCountry As String
    Dim strCurrency As String
    strCountry = title
    strCurrency = ""
    
    Dim arrCurrency() As String
    arrCurrency() = Split(title, "(")
    
    If UBound(arrCurrency) > 0 Then
        strCountry = arrCurrency(0)
        strCurrency = "(" & arrCurrency(1)
    End If
    
    ' Set the currency name
    Sheets(worksheetName).Cells.Item(row, 1).Value = strCountry
    Sheets(worksheetName).Cells.Item(row, 2).Value = strCurrency
    
    Dim cellCount As Integer
    cellCount = 3
    
    ' Stuff the values into the cells
    For Each element In arrContent
        Sheets(worksheetName).Cells.Item(row, cellCount).Value = element
        cellCount = cellCount + 1
    Next element
End Sub


' Process the row only if it has the values we want
Private Sub ProcessRow(outputWorksheetName, row, col, objSkipLineRegExp, objRegExp)

    strRow = Trim(Cells.Item(row, col))
    Set skipLineRegExpMatches = objSkipLineRegExp.Execute(strRow)
    
    If skipLineRegExpMatches.Count = 1 Then
        'MsgBox "SkipLine!"
        'Call doSomething
    Else
        Call ProcessRowData(outputWorksheetName, row, strRow, objRegExp)
    End If
End Sub


