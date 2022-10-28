' Steps:
' ----------------------------------------------------------------------------

' Part I:

' 1. Extract the number before the phrase "_census_data" to figure out the year.
' 2. Add the year to the first column of each spreadsheet.
' 3. Split the "Place" column into "County" and "State".
' 4. Convert the household and per capita income columns to currency values for all cells.


Sub Census_pt1():

        ' --------------------------------------------
        ' INSERT THE YEAR
        ' --------------------------------------------

        ' Create a Variable to Hold File Name, Last Row, and Year
        Dim WorksheetName As String
        Dim CensusYear() As String
        Dim LastRow As Integer
        
        
        ' Split the WorksheetName
        WorksheetName = ActiveSheet.Name
        CensusYear = Split(WorksheetName, "_")
        ActiveSheet.Range("A1").EntireColumn.Insert
        Range("A1").Value = "Year"
        
        
        ' Determine the Last Row
        LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row


        ' Grabbed the WorksheetName


        


        ' Add a Column for the Year


        ' Add the word Year to the First Column Header


        ' Add the Year to all rows


        ' --------------------------------------------
        ' SPLIT COUNTY AND STATE
        ' --------------------------------------------

        ' Add the State Column after County

        
        ' Rename Place to County

        
        ' Label State Column


        ' Split County and State and store values in appropriate
        ' column by looping through and renaming each
        'For i = 2 To LastRow



        'Next i

        ' --------------------------------------------
        ' CORRECT THE CURRENCY FORMAT
        ' --------------------------------------------

        ' Add the currency





End Sub

