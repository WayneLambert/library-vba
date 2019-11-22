Source: https://www.excelcampus.com/vba/clear-filters-showalldata/

'A list of different filtering alternatives
'Also note how the iCol number argument is being calculated
Sub Filter_Top_5()

Dim lo As ListObject
Dim iCol As Long

Set lo = ActiveSheet.ListObjects(1)

'Set filter field
iCol = lo.ListColumns("Revenue").Index

'Apply filters to a column (field)
lo.Range.AutoFilter Field:=iCol, Criteria1:="5", Operator:=xlTop10Items
  
End Sub

'******************************************************************************************

Sub Filter_Bottom_5()

Dim lo As ListObject
Dim iCol As Long

Set lo = ActiveSheet.ListObjects(1)

'Clear filters
'lo.AutoFilter.ShowAllData

'Set filter field
iCol = lo.ListColumns("Revenue").Index

'Apply filters to a column (field)
lo.Range.AutoFilter Field:=iCol, Criteria1:="5", Operator:=xlBottom10Items

End Sub

'******************************************************************************************

Sub Filter_Above_Average()

Dim lo As ListObject
Dim iCol As Long

Set lo = ActiveSheet.ListObjects(1)

'Set filter field
iCol = lo.ListColumns("Revenue").Index

'Apply filters to a column (field)
lo.Range.AutoFilter Field:=5, Criteria1:=xlFilterAboveAverage, Operator:=xlFilterDynamic

End Sub

'******************************************************************************************

Sub Filter_Below_Average()

Dim lo As ListObject
Dim iCol As Long

Set lo = ActiveSheet.ListObjects(1)

'Set filter field
iCol = lo.ListColumns("Revenue").Index

'Apply filters to a column (field)
lo.Range.AutoFilter Field:=5, Criteria1:=xlFilterBelowAverage, Operator:=xlFilterDynamic

End Sub

'******************************************************************************************

Sub Filter_Greater_Than_100()

Dim lo As ListObject
Dim iCol As Long

Set lo = ActiveSheet.ListObjects(1)

'Set filter field
iCol = lo.ListColumns("Revenue").Index

'Apply filters to a column (field)
lo.Range.AutoFilter Field:=5, Criteria1:=">100", Operator:=xlFilterValues

End Sub

'******************************************************************************************

Sub Clear_All_Table_Filters()

Dim lo As ListObject

Set lo = ActiveSheet.ListObjects(1)

'Clear all Table filters
lo.AutoFilter.ShowAllData
  
End Sub

'******************************************************************************************

Sub Clear_Field_Filter_Revenue()

Dim lo As ListObject
Dim iCol As Long

Set lo = ActiveSheet.ListObjects(1)

'Set filter field
iCol = lo.ListColumns("Revenue").Index

'Clear filter from field (column)
lo.Range.AutoFilter Field:=iCol

End Sub

'******************************************************************************************

Sub AutoFilter_Table()
'AutoFilters on Tables work the same way.

Dim lo As ListObject 'Excel Table

'Set the ListObject (Table) variable
Set lo = Sheet1.ListObjects(1)

'AutoFilter is member of Range object
'The parent of the Range object is the List Object
lo.Range.AutoFilter

End Sub

'******************************************************************************************

Sub Clear_All_Filters_Range()

'To Clear All Fitlers use the ShowAllData method for
'for the sheet.  Add error handling to bypass error if
'no filters are applied.  Does not work for Tables.
On Error Resume Next
Sheet1.ShowAllData
On Error GoTo 0

End Sub

'******************************************************************************************

Sub Clear_Column_Filter_Range()

'To clear the filter from a Single Column, specify the
'Field number only and no other parameters
Sheet1.Range("B3:G1000").AutoFilter Field:=4

End Sub

'******************************************************************************************

Sub Clear_All_Filters_Table()

Dim lo As ListObject

'Set reference to the first Table on the sheet
Set lo = Sheet1.ListObjects(1)

'Clear All Filters for entire Table
lo.AutoFilter.ShowAllData

End Sub

'******************************************************************************************

Sub Clear_All_Table_Filters_On_Sheet()

Dim lo As ListObject

'Loop through all Tables on the sheet
For Each lo In Sheet1.ListObjects

'Clear All Filters for entire Table
lo.AutoFilter.ShowAllData

Next lo

End Sub

'******************************************************************************************

Sub Clear_Column_Filter_Table()

Dim lo As ListObject

'Set reference to the first Table on the sheet
Set lo = Sheet1.ListObjects(1)

'Clear filter on Single Table Column
'by specifying the Field parameter only
lo.Range.AutoFilter Field:=4

End Sub

'******************************************************************************************

Sub Dynamic_Field_Number()
'Techniques to find and set the Field based on the column name.

Dim lo As ListObject
Dim iCol As Long

'Set reference to the first Table on the sheet
Set lo = Sheet1.ListObjects(1)

'Set filter field column number
iCol = lo.ListColumns("Product").Index

'Use Match function for regular ranges
'iCol = WorksheetFunction.Match("Product", Sheet1.Range("B3:F3"), 0)

'Use the variable for the Field parameter
lo.Range.AutoFilter Field:=iCol, Criteria1:="Product 3"

End Sub

'******************************************************************************************

Sub Blank_NonBlank_Cells_Filter()

Dim lo As ListObject
Dim iCol As Long

'Set reference to the first Table on the sheet
Set lo = Sheet1.ListObjects(1)

'Set filter field
iCol = lo.ListColumns("Product").Index

'Blank cells - set equal to nothing
lo.Range.AutoFilter Field:=iCol, Criteria1:="="

'Non-blank cells - use NOT operator <>
lo.Range.AutoFilter Field:=iCol, Criteria1:="<>"

End Sub

'******************************************************************************************

Sub AutoFilter_Text_Examples()
'Examples for filtering columns with TEXT

Dim lo As ListObject
Dim iCol As Long

'Set reference to the first Table on the sheet
Set lo = Sheet1.ListObjects(1)

'Set filter field
iCol = lo.ListColumns("Product").Index

'Clear Filters
lo.AutoFilter.ShowAllData

'All lines starting with .AutoFilter are a continuation
'of the with statement.
With lo.Range

'Single Item
.AutoFilter Field:=iCol, Criteria1:="Product 2"

'2 Criteria using Operator:=xlOr
.AutoFilter Field:=iCol, _
            Criteria1:="Product 3", _
            Operator:=xlOr, _
            Criteria2:="Product 4"

'More than 2 Criteria (list of items in an Array function)
.AutoFilter Field:=iCol, _
            Criteria1:=Array("Product 4", "Product 5", "Product 7"), _
            Operator:=xlFilterValues
                    
'Begins With - use asterisk as wildcard character at end of string
.AutoFilter Field:=iCol, Criteria1:="Product*"

'Ends With - use asterisk as wildcard character at beginning
'of string
.AutoFilter Field:=iCol, Criteria1:="*2"

'Contains - wrap search text in asterisks
.AutoFilter Field:=iCol, Criteria1:="*uct*"

'Does not contain text
'Start with Not operator <> and wrap search text in asterisks
.AutoFilter Field:=iCol, Criteria1:="<>*8*"

'Contains a wildcard character * or ?
'Use a tilde ~ before the character to search for values with
'wildcards
.AutoFilter Field:=iCol, Criteria1:="Product 1~*"

End With

End Sub

'******************************************************************************************

Sub AutoFilter_Number_Examples()
'Examples for filtering columns with NUMBERS

Dim lo As ListObject
Dim iCol As Long

'Set reference to the first Table on the sheet
Set lo = Sheet1.ListObjects(1)

'Set filter field
iCol = lo.ListColumns("Revenue").Index

'Clear Filters
lo.AutoFilter.ShowAllData

With lo.Range

'Single number - Use formatting that is visible in
'filter drop-down menu
.AutoFilter Field:=iCol, Criteria1:="$2,955.25"

'Not equal to - Does not require number formatting to match
.AutoFilter Field:=iCol, Criteria1:="<>2955.25"

'Greater than or less than a number
'(comparison operator < > = before number in Criteria1)
.AutoFilter Field:=iCol, Criteria1:="<4000"

'Between 2 numbers
'(greater than or equal to 100 and less than 4000)
.AutoFilter Field:=iCol, _
            Criteria1:=">=100", _
            Operator:=xlAnd, _
            Criteria2:="<4000"

'Outside range (less than 100 OR greater than 4000)
.AutoFilter Field:=iCol, _
            Criteria1:="<100", _
            Operator:=xlOr, _
            Criteria2:=">4000"

'Top 10 items (Criteria1 is number of items)
.AutoFilter Field:=iCol, _
            Criteria1:="10", _
            Operator:=xlTop10Items

'Bottom 5 items (Criteria1 is number of items)
.AutoFilter Field:=iCol, _
            Criteria1:="5", _
            Operator:=xlBottom10Items

'Top 10 Percent (Criteria1 is number of items)
.AutoFilter Field:=iCol, _
            Criteria1:="10", _
            Operator:=xlTop10Percent

'Bottom 7 Percent
.AutoFilter Field:=iCol, _
            Criteria1:="7", _
            Operator:=xlBottom10Percent

'Above Average - Operator:=xlFilterDynamic
.AutoFilter Field:=iCol, _
            Criteria1:=xlFilterAboveAverage, _
            Operator:=xlFilterDynamic

'Below Average
.AutoFilter Field:=iCol, _
            Criteria1:=xlFilterBelowAverage, _
            Operator:=xlFilterDynamic

End With

End Sub

'******************************************************************************************

Sub AutoFilter_Date_Examples()
'Examples for filtering columns with DATES

Dim lo As ListObject
Dim iCol As Long

'Set reference to the first Table on the sheet
Set lo = Sheet1.ListObjects(1)

'Set filter field
iCol = lo.ListColumns("Date").Index

'Clear Filters
lo.AutoFilter.ShowAllData

With lo.Range
    
'Single Date - Use same date format that is applied to column
.AutoFilter Field:=iCol, Criteria1:="=1/2/2014"

'Before Date
.AutoFilter Field:=iCol, Criteria1:="<1/31/2014"

'After or equal to Date
.AutoFilter Field:=iCol, Criteria1:=">=1/31/2014"

'Date Range (between dates)
.AutoFilter Field:=iCol, _
                    Criteria1:=">=1/1/2014", _
                    Operator:=xlAnd, _
                    Criteria2:="<=12/31/2015"
                    
End Sub

'******************************************************************************************

Sub AutoFilter_Multiple_Dates_Examples()
'Examples for filtering columns for multiple DATE TIME PERIODS

Dim lo As ListObject
Dim iCol As Long

'Set reference to the first Table on the sheet
Set lo = Sheet1.ListObjects(1)

'Set filter field
iCol = lo.ListColumns("Date").Index

'Clear Filters
lo.AutoFilter.ShowAllData

With lo.Range

'When filtering for multiple periods that are selected from
'filter drop-down menu,use Operator:=xlFilterValues and
'Criteria2 with a patterned Array.  The first number is the
'time period.  Second number is the last date in the period.

'First dimension of array is the time period group
    '0-Years
    '1-Months
    '2-Days
    '3-Hours
    '4-Minutes
    '5-Seconds


'Multiple Years (2014 and 2016) use last day of the time
'period for each array item
.AutoFilter Field:=iCol, _
            Operator:=xlFilterValues, _
            Criteria2:=Array(0, "12/31/2014", 0, "12/31/2016")

'Multiple Months (Jan, Apr, Jul, Oct in 2015)
.AutoFilter Field:=iCol, _
            Operator:=xlFilterValues, _
            Criteria2:=Array(1, "1/31/2015", 1, "4/30/2015", 1, "7/31/2015", 1, "10/31/2015")

'Multiple Days
'Last day of each month: Jan, Apr, Jul, Oct in 2015)
.AutoFilter Field:=iCol, _
            Operator:=xlFilterValues, _
            Criteria2:=Array(2, "1/31/2015", 2, "4/30/2015", 2, "7/31/2015", 2, "10/31/2015")

'Set filter field
    iCol = lo.ListColumns("Date Time").Index
    
'Clear Filters
lo.AutoFilter.ShowAllData

'Multiple Hours (All dates in the 11am hour on 1/10/2018
'and 11pm hour on 1/20/2018)
.AutoFilter Field:=iCol, _
            Operator:=xlFilterValues, _
            Criteria2:=Array(3, "1/10/2018 13:59:59", 3, "1/20/2018 23:59:59")

End With

End Sub

'******************************************************************************************

Sub AutoFilter_Dates_in_Period_Examples()
'Examples for filtering columns for DATES IN PERIOD
'Date filters presets found in the Date Filters sub menu

Dim lo As ListObject
Dim iCol As Long

'Set reference to the first Table on the sheet
Set lo = Sheet1.ListObjects(1)

'Set filter field
iCol = lo.ListColumns("Date").Index

'Clear Filters
lo.AutoFilter.ShowAllData

'Operator:=xlFilterDynamic
'Criteria1:= one of the following enumerations

' Value Constant
' 1     xlFilterToday
' 2     xlFilterYesterday
' 3     xlFilterTomorrow
' 4     xlFilterThisWeek
' 5     xlFilterLastWeek
' 6     xlFilterNextWeek
' 7     xlFilterThisMonth
' 8     xlFilterLastMonth
' 9     xlFilterNextMonth
' 10    xlFilterThisQuarter
' 11    xlFilterLastQuarter
' 12    xlFilterNextQuarter
' 13    xlFilterThisYear
' 14    xlFilterLastYear
' 15    xlFilterNextYear
' 16    xlFilterYearToDate
' 17    xlFilterAllDatesInPeriodQuarter1
' 18    xlFilterAllDatesInPeriodQuarter2
' 19    xlFilterAllDatesInPeriodQuarter3
' 20    xlFilterAllDatesInPeriodQuarter4
' 21    xlFilterAllDatesInPeriodJanuary
' 22    xlFilterAllDatesInPeriodFebruray <-February is misspelled in Constant
' 23    xlFilterAllDatesInPeriodMarch
' 24    xlFilterAllDatesInPeriodApril
' 25    xlFilterAllDatesInPeriodMay
' 26    xlFilterAllDatesInPeriodJune
' 27    xlFilterAllDatesInPeriodJuly
' 28    xlFilterAllDatesInPeriodAugust
' 29    xlFilterAllDatesInPeriodSeptember
' 30    xlFilterAllDatesInPeriodOctober
' 31    xlFilterAllDatesInPeriodNovember
' 32    xlFilterAllDatesInPeriodDecember
    
With lo.Range

'All dates in January (across all years)
.AutoFilter Field:=iCol, _
            Operator:=xlFilterDynamic, _
            Criteria1:=xlFilterAllDatesInPeriodJanuary

'All dates in Q2 (across all years)
.AutoFilter Field:=iCol, _
            Operator:=xlFilterDynamic, _
            Criteria1:=xlFilterAllDatesInPeriodQuarter2

End With

End Sub

'******************************************************************************************

Sub AutoFilter_Color_Icon_Examples()
'Examples for filtering columns with COLORS and ICONS

Dim lo As ListObject
Dim iCol As Long

'Set reference to the first Table on the sheet
Set lo = Sheet1.ListObjects(1)

'Set filter field
iCol = lo.ListColumns("Product").Index

'Clear Filters
lo.AutoFilter.ShowAllData

With lo.Range

'Colors

'Font and fill colors are set in Criteria 1.
'The macro recorder gives us the RGB value.  RGB can also
'be found in the More Colors menu on the Custom tab.

'Filter for dark red cell fill color
.AutoFilter Field:=iCol, _
            Criteria1:=RGB(192, 0, 0), _
            Operator:=xlFilterCellColor
    
'Font Color for dark green
.AutoFilter Field:=iCol, _
            Criteria1:=RGB(0, 97, 0), _
            Operator:=xlFilterFontColor


'Icons

'Set filter field
iCol = lo.ListColumns("Icon").Index
        
'Clear Filters
lo.AutoFilter.ShowAllData

'Filter for Icon (conditional formatting)
.AutoFilter Field:=iCol, _
            Criteria1:=ThisWorkbook.IconSets(xl4CRV).Item(4), _
            Operator:=xlFilterIcon
    
End With
    
End Sub

'******************************************************************************************