Public Class frmFunctions
    'Programmed by: Lukas Siemers, ITE 285, tuesdays/thursdays (11:00 - 12:15) The purpose of this Program is to get more comfortable with Functions and Sub procedures
    Structure Inventory
        Dim Productnumber As String                 'Variables created for the Data input at Load Time
        Dim distributor As String                   'Variables created for the Data input at Load Time
        Dim itemNumber As String                    'Variables created for the Data input at Load Time
        Dim ProductDescription As String            'Variables created for the Data input at Load Time
        Dim receivedDate As String                  'Variables created for the Data input at Load Time
        Dim availableQuantity As Integer            'Variables created for the Data input at Load Time
        Dim reorderQuantity As Double               'Variables created for the Data input at Load Time
        Dim wholesalePrice As Decimal                'Variables created for the Data input at Load Time
        Dim retailPrice As Decimal                   'Variables created for the Data input at Load Time
    End Structure
    Dim item() As Inventory         'Class level structed Array (available to all modules)
    Private Sub frmFunctions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim strItem() As String = IO.File.ReadAllLines("Inventory.txt")
        Dim record, data() As String
        ReDim item(strItem.Count - 1)       'Redimming
        For i As Integer = 0 To strItem.Count - 1
            record = strItem(i)         'record to split
            data = record.Split(","c)   'Splitting 
            item(i).Productnumber = data(0)         'Assigning value
            item(i).ProductDescription = data(1)    'Assigning Value
            item(i).receivedDate = data(2)          'Assigning Value
            item(i).availableQuantity = data(3)     'Assigning Value
            item(i).reorderQuantity = data(4)       'Assigning Value
            item(i).wholesalePrice = data(5)        'Assigning Value
            item(i).retailPrice = data(6)           'Assigning Value
        Next
    End Sub
    Private Sub btnMaster_Click(sender As Object, e As EventArgs) Handles btnMaster.Click
        Dim itemQuery = From thing In item  'seperates each line
                        Order By thing.ProductDescription Ascending     'Ordering by ascending 
                        Let results = thing.ProductDescription & ": " & thing.availableQuantity & " @ " & thing.retailPrice.ToString("c2") 'creating Format
                        Select results  'selecting the results

        Dim itemArray() As String = itemQuery.ToArray   'filling the Array with the Information above
        DisplayMaster(itemArray)    'Displaying Sub
    End Sub
    Sub DisplayMaster(itemArray() As String)    'Sub for displaying
        lstDisplay.Items.Clear()
        lstDisplay.Items.Add("Master Inventory Report") 'Title
        lstDisplay.Items.Add(" ")
        For Each piece In itemArray
            lstDisplay.Items.Add(piece)     'Dispalys each line with seperated information
        Next
    End Sub
    Private Sub btnProfit_Click(sender As Object, e As EventArgs) Handles btnProfit.Click
        Dim high, low, avg As Double    'Variables for Sub
        Dim Profit = From thing In item
                     Order By thing.Productnumber Ascending             ' Determines the Format order
                     Let estProfit = ItemProfit(thing.availableQuantity, thing.retailPrice, thing.wholesalePrice)   'selecting the variables for the format desing
                     Let Format = thing.Productnumber & ": " & thing.ProductDescription & " - " & thing.availableQuantity & ": Estimated Profit " & estProfit.ToString("c2") 'Format
                     Where thing.availableQuantity > 0  'If the Value is greater then 0
                     Select Format  'selecting the format

        Dim highlowavg = From thing In item
                         Let estProfit = ItemProfit(thing.availableQuantity, thing.retailPrice, thing.wholesalePrice)   'selecting the variables for the format desing              
                         Where thing.availableQuantity > 0  'If the Value is greater then 0
                         Select estProfit

        Dim profitArray() As String = Profit.ToArray    'reading the fromat into an array
        Dim ItemArray() As Double = highlowavg.ToArray  'reading the format into an array

        DisplayProfits(profitArray, ItemArray, high, low, avg) 'Display Sub
    End Sub
    Sub DisplayProfits(ProfitArray() As String, ItemArray() As Double, high As Double, low As Double, avg As Double) 'Display Sub
        lstDisplay.Items.Clear()
        lstDisplay.Items.Add("Inventory Profit Report") 'Title
        lstDisplay.Items.Add(" ")
        If ProfitArray.Count > 0 Then   'if the Profit is above 0 then Display
            For Each piece In ProfitArray   'selecting Line by Line
                lstDisplay.Items.Add(piece) 'Displaying Line by Line
            Next
        Else
            lstDisplay.Items.Add("None Found")      'else if none was found
        End If
        lstDisplay.Items.Add(" ")
        DetermineStats(ItemArray, high, low, avg)   'calling the high,low,avg sub
        lstDisplay.Items.Add("Highest Projected Profit: " & high.ToString("c2"))    'highest profit
        lstDisplay.Items.Add("Lowest Projected Profit: " & low.ToString("c2"))      'lowest profit 
        lstDisplay.Items.Add("Average Projected Profit: " & avg.ToString("c2"))     'average profit
    End Sub
    Sub DetermineStats(ItemArray() As Double, ByRef high As Double, ByRef low As Double, ByRef avg As Double)
        high = ItemArray.Max    'maximum profit
        low = ItemArray.Min     'minimum profit
        avg = ItemArray.Average 'average profit
    End Sub
    Function ItemProfit(availableQuantity As Integer, retailPrice As Decimal, WholesalePrice As Decimal) As Double
        Return (availableQuantity * (retailPrice - WholesalePrice))  'Equation for the Profit
    End Function
    Private Sub btnPurge_Click(sender As Object, e As EventArgs) Handles btnPurge.Click

        Dim input As String
        GetDate(input)  'Input from sub

        If DataOk(input) = False Then   'Data Validation
            Return
        End If

        Dim purgeQuery = From stuff In item
                         Let purgedate = stuff.receivedDate
                         Where CDate(purgedate) <= CDate(input)    'If date is equal to input or greater then 
                         Let Format = NewRecord(stuff.Productnumber, stuff.ProductDescription, stuff.receivedDate, stuff.availableQuantity, stuff.reorderQuantity, stuff.wholesalePrice, stuff.retailPrice) 'format
                         Select Format          'selecting

        Dim currentInventory = From stuff In item
                               Let inventory = stuff.receivedDate
                               Where CDate(inventory) > CDate(input)    'If inventory is greater then input
                               Let Format = NewRecord(stuff.Productnumber, stuff.ProductDescription, stuff.receivedDate, stuff.availableQuantity, stuff.reorderQuantity, stuff.wholesalePrice, stuff.retailPrice)   'format
                               Select Format        'selecting 

        Dim DateArray() As String = purgeQuery.ToArray      'reading the query to the Array
        Dim NewDateArray() As String = currentInventory.ToArray 'reading the query to the Array

        DisplayCurrent(NewDateArray)    'calling the display sub
        DisplayPurged(DateArray)        'calling the display sub
    End Sub
    Sub GetDate(ByRef input As String)
        input = InputBox("Please enter a Purge Date in Date Format (mm/dd/yyyy)")   'Input Sub
    End Sub
    Function DataOk(input As String) As Boolean         'Data Validation Function
        If IsDate(input) = False Then
            MessageBox.Show("Please enter in Date Format (mm/dd/yyyy)") 'messagebox
            Return False    'If format is wrong
        ElseIf input = "" Then
            MessageBox.Show("Please enter a Date")  'If the inputbox is empty
            Return False
        Else
            Return True     'If format is right 
        End If
    End Function
    Function NewRecord(Productnumber As String, ProductDescription As String, receivedDate As String, availableQuantity As Double, reorderQuantity As Double, wholesalePrice As Decimal, retailPrice As Decimal) As String
        Dim FormatInventory As String
        FormatInventory = Productnumber & "," & ProductDescription & "," & receivedDate & "," & availableQuantity & "," & reorderQuantity & "," & wholesalePrice & "," & retailPrice    'Creating the Format
        Return FormatInventory  'Returning Format
    End Function
    Sub DisplayCurrent(NewDateArray() As String)      'Displaying Sub
        lstDisplay.Items.Clear()
        lstDisplay.Items.Add("Active Records")  'Title
        If NewDateArray.Count > 0 Then
            For Each inventory In NewDateArray                                  'Displays each line
                lstDisplay.Items.Add(inventory)
                IO.File.WriteAllLines("CurrentInventory.txt", NewDateArray)     'New Textfile
            Next
        Else
            lstDisplay.Items.Add("None Found")  'else message
        End If
    End Sub
    Sub DisplayPurged(Datearray() As String)
        lstDisplay.Items.Add(" ")
        lstDisplay.Items.Add("Purged Record")   'Title
        If Datearray.Count > 0 Then
            For Each inputDate In Datearray                                  'Display each Line
                lstDisplay.Items.Add(inputDate)
                IO.File.WriteAllLines("PurgedInventory.txt", Datearray)      'New Textfile
            Next
        Else
            lstDisplay.Items.Add("None Found")  'else message
        End If
    End Sub
    Private Sub btnOoutputQuery_Click(sender As Object, e As EventArgs) Handles btnOoutputQuery.Click
        Dim myQuery = From thing In item
                      Let expirationDate = CDate(thing.receivedDate)    'equal
                      Let Format = ExpDate(thing.receivedDate, thing.ProductDescription, thing.receivedDate)
                      Select Format     'selecting

        Dim DateArray() As String = myQuery.ToArray 'Addint the query to the Array
        DisplayDates(DateArray)     'display Sub
    End Sub
    Function ExpDate(receivedDate As Date, ProductDescription As String, expirationDate As Date) As String
        Dim today As New DateTime(2022, 3, 15)      'assigning a Date for variable today
        Dim dt1 As DateTime = CDate(expirationDate)     'converting the expiration to date
        Dim TTF As New TimeSpan         'creating a new timeSpan
        TTF = today.Subtract(dt1)       'Subtracting today from the expiration date to determing the leftover days
        Dim Format As String    'dimming format
        Format = receivedDate & ", " & ProductDescription & ", " & "Estimated days till Experation Date: " & TTF.TotalDays  'creating format and adding the days
        Return Format   'returning format
    End Function
    Sub DisplayDates(DateArray() As String)     'Display Sub
        lstDisplay.Items.Clear()
        lstDisplay.Items.Add("Days till experation as of Today (3/15/2022)")    'Title
        lstDisplay.Items.Add(" ")
        For Each thing In DateArray
            lstDisplay.Items.Add(thing) 'Displaying each line
        Next
    End Sub
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        lstDisplay.Items.Clear()        'Clearing Display
    End Sub
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()      'Closing the Program 
    End Sub
End Class
