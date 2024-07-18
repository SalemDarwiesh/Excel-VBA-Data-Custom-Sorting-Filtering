# Excel Data Custome Sorting and Filtering with VBA

Optimize data management in Excel using VBA with this script that sorts and filters data based on specified criteria.

## Description

This repository contains a VBA script (`SortAndFilterData`) designed to automate sorting and filtering operations in Excel. It applies filters to columns A:C, sorts data in column B in ascending order, and sorts data in column A with a custom order.

## Usage

To use this script:

1. **Download the VBA Script**: Copy the `SortAndFilterData` subroutine code from this repository.

2. **Open Your Excel Workbook**:
   - Press `Alt + F11` to open the Visual Basic for Applications (VBA) editor.
   - Import the VBA script into your workbook by copying and pasting the code into a new module.

3. **Run the Script**:
   - Once imported, execute the `SortAndFilterData` subroutine to apply sorting and filtering to your data.

## Example

```vba
Sub SortAndFilterData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("SHEETNAME")
    
    ' Apply AutoFilter to columns A:C
    With ws
        .Columns("A:C").AutoFilter
        ' Clear existing sorting fields
        .AutoFilter.Sort.SortFields.Clear
        
        ' Sort by column B ascending
        .AutoFilter.Sort.SortFields.Add Key:=Range("B:B"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        
        ' Sort by column A ascending, with custom order
        .AutoFilter.Sort.SortFields.Add Key:=Range("A:A"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, _
            CustomOrder:="ITEMONE,ITEMTWO,ITEMFOUR,ITEMTHREE,*ITEMCONTAINS*", _
            DataOption:=xlSortNormal
        
        ' Apply the sorting
        With .AutoFilter.Sort
            .Header = xlYes ' Assuming your data has headers
            .MatchCase = False
            .Apply
        End With
    End With
End Sub
