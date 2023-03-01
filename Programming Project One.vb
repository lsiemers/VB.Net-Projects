Public Class frmProjectOne
    'Programmed by: Lukas Siemers, ISC/ITE 285, Tuesday/Thursday 11:00 - 12:15, The purpose of this Program is to Read in inventory, -
    'look at profits, reorder needed items and generate a list of distributors that will then be transfered to a text file
    Dim Inventory() As String           'Class array
    Private Sub frmProjectOne_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Inventory = IO.File.ReadAllLines("Inventory.txt")   'Reading in the File
        For Each item In Inventory
            lstDisplay.Items.Add(item)                      'Displaying it as part of the load event
        Next
    End Sub
    Private Sub btnGenerateMasterRepor_Click(sender As Object, e As EventArgs) Handles btnGenerateMasterRepor.Click
        lstDisplay.Items.Clear()
        Dim sortedQuery = From item In Inventory
                          Let data = item.Split(","c)                       'Splitting by the "," to Idententify each record
                          Let ProductDescription = data(1)                  'Gathering The Product Description
                          Let AvailableQuantity = CDbl(data(3))             'Gathering the Available Quantity
                          Let RetailPrice = CDec(data(6))                   'Gathering the retail Price
                          Order By ProductDescription Ascending             'Ordering by the Product Description
                          Select ProductDescription, AvailableQuantity, RetailPrice     'Selecting the Variables

        lstDisplay.Items.Add("Master Inventory List")       'Title
        lstDisplay.Items.Add(" ")
        For Each item In sortedQuery
            lstDisplay.Items.Add(item.ProductDescription & " - " & item.AvailableQuantity & " @ " & item.RetailPrice.ToString("c2"))    'Displaying in the Format
        Next
    End Sub
    Private Sub btnGenerateProfit_Click(sender As Object, e As EventArgs) Handles btnGenerateProfit.Click
        Dim totalProfit As Double
        lstDisplay.Items.Clear()
        Dim sortedQuery = From item In Inventory
                          Let data = item.Split(","c)                           'Splitting by the "," to Idententify each record
                          Let distributor = data(0)                             'Gathering the Distrubutor Number
                          Let ProductDescription = data(1)                      'Gathering the Product Description
                          Let AvailableQuantity = CDbl(data(3))                 'Gathering the Available Quantity
                          Let WholesalePrice = CDec(data(5))                    'Gathering the Wholesale Price
                          Let RetailPrice = CDec(data(6))                       'Gathering the Retail Price
                          Let format = CDec(AvailableQuantity * RetailPrice) - (AvailableQuantity * WholesalePrice)     'Creating a Format and math equation
                          Order By format Descending                            'Ordering by the math equation results above 
                          Select distributor, ProductDescription, AvailableQuantity, WholesalePrice, RetailPrice, format    'Selecting the varaibles
                          Where format > 0  'Only selecting where format(Profit) is above $0

        lstDisplay.Items.Add("Projected Profits Report")                                'Title
        lstDisplay.Items.Add(" ")
        If sortedQuery.Count > 0 Then                                                   'If the format(profit) is above $0
            For Each item In sortedQuery
                lstDisplay.Items.Add(item.distributor & " - " & item.ProductDescription)
                lstDisplay.Items.Add("Qty: " & item.AvailableQuantity & " - " & "Price: " & item.RetailPrice.ToString("c2"))
                lstDisplay.Items.Add("Estimated Profits: " & item.format.ToString("c2"))
                lstDisplay.Items.Add(" ")
                totalProfit += item.format                                              'adding all of the profits together for displaying outside the loop
            Next
        Else
            lstDisplay.Items.Add("None Found")                                          'Message if no Profits are provided
        End If
        lstDisplay.Items.Add(" ")
        lstDisplay.Items.Add("Total Projected Profits: " & totalProfit.ToString("c2"))  'The total that is added above 
    End Sub
    Private Sub btnReorderList_Click(sender As Object, e As EventArgs) Handles btnReorderList.Click
        Dim symbol As String                                                'Dimmed a string variable to use for the ***
        lstDisplay.Items.Clear()
        Dim sortedQuery = From item In Inventory
                          Let data = item.Split(","c)                       'Splitting the Record
                          Let distributor = data(0)                         'Gathering the Distributor
                          Let ProductDescription = data(1)                  'Gathering the Product Description
                          Let Productdate = data(2)                         'Gathering the Product date
                          Let AvailableQuantity = CDbl(data(3))             'Gathering the Available Quantity
                          Let ReorderQuantity = CDbl(data(4))               'Gathering Reorder Quantity
                          Let WholesalePrice = CDec(data(5))                'Gathering Wholesale Price
                          Let RetailPrice = CDec(data(6))                   'Gathering Retail Price
                          Let Qtybelow = ReorderQuantity - AvailableQuantity    'Gathering the Quantity that is below reordering by subtracting the two variables
                          Order By AvailableQuantity Ascending              'Ordering by the Available Quantity Ascending 
                          Select distributor, ProductDescription, AvailableQuantity, Productdate, Qtybelow, ReorderQuantity     'selecting all variables

        lstDisplay.Items.Add("ReOrder Report")          'Title
        lstDisplay.Items.Add(" ")

        For Each item In sortedQuery
            If item.AvailableQuantity <= 0 Then     'If available quantity is less or equal to 0 
                symbol = "***"                                  'It will add the *** if the statement is true
            Else
                symbol = " "                        'If its not true it will just be empty space
            End If
            If item.AvailableQuantity < item.ReorderQuantity Then       'If the Available qty is less then Reorder qty
                lstDisplay.Items.Add(item.distributor & " - " & item.ProductDescription)    'Displaying
                lstDisplay.Items.Add("Received Date: " & item.Productdate)                  'Displaying
                lstDisplay.Items.Add("Available Qty: " & item.AvailableQuantity & " - " & "Qty Below Reorder Amt: " & item.Qtybelow & symbol)   'Displaying including the symbol
                lstDisplay.Items.Add(" ")
            End If
        Next
        If sortedQuery.Count = 0 Then               'if statement if no reordering is needed
            lstDisplay.Items.Add("No Reordering needed!")
        End If
    End Sub
    Private Sub btnDistrubotorsList_Click(sender As Object, e As EventArgs) Handles btnDistrubotorsList.Click
        lstDisplay.Items.Clear()
        Dim sortedQuery = From item In Inventory
                          Let data = item.Split(","c)           'Splitting by the "," for each record
                          Let distributor = data(0)             'Gathering the distributor
                          Let Code = distributor.Substring(0, 3)    'Substringing out the code to identify the code
                          Order By Code Ascending                       'Ordering by the code
                          Select Code
                          Distinct                                      'Deleting doubles

        lstDisplay.Items.Add("Current Distributors List")               'Titles
        lstDisplay.Items.Add(" ")
        For Each item In sortedQuery
            lstDisplay.Items.Add(item)                                  'Displaying
        Next
    End Sub

    Private Sub btnProductFile_Click(sender As Object, e As EventArgs) Handles btnProductFile.Click
        Dim filename As String              'Dimming a Variable to input the distributor code and _textfile.txt
        lstDisplay.Items.Clear()
        Dim search As String = InputBox("Please Enter Distributurs #")          'User input to search for the code
        Dim sortedQuery = From item In Inventory
                          Let data = item.Split(","c)                               'Splitting
                          Let distributor = data(0)                                 'Gathering the Distributor
                          Let ProductDescription = data(1)                          'Gathering the Product Description
                          Let Productdate = data(2)                                 'Gathering the Product Date
                          Let AvailableQuantity = CDbl(data(3))                     'Gathering the Available Quantity
                          Let ReorderQuantity = CDbl(data(4))                       'Gathering the Reorder Quantity
                          Let WholesalePrice = CDec(data(5))                        'Gathering the Wholesale Price
                          Let RetailPrice = CDec(data(6))                           'Gathering the Retail Price
                          Let Code = distributor.Substring(0, 3)                    'Substring the distributor code
                          Let PartNumber = distributor.Substring(3, 3)              'Substring the the Second code
                          Let Formattype = distributor & "," & ProductDescription & "," & Productdate & "," & AvailableQuantity & "," & ReorderQuantity & "," & WholesalePrice & "," & RetailPrice
                          Where Code = search       'if user input is equal to the code
                          Order By PartNumber Ascending     'Ordering by substring ascendin
                          Select Formattype         'selecting the format 


        lstDisplay.Items.Add("Given Distributor's Records for " & search)   'Title
        lstDisplay.Items.Add(" ")
        If sortedQuery.Count > 0 Then       'only does this part if the matches are greater then 0
            For Each line In sortedQuery
                lstDisplay.Items.Add(line)      'reading in information
                filename = search & "_Inventory.txt"    'Creating the Format to write in the textfile
                IO.File.WriteAllLines(filename, sortedQuery)    'Reading in the Information with the variable i created above
            Next
        Else
            lstDisplay.Items.Add("No matches found for Distributor Code!")      'Displaying if no code matches the user input
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()  'Close Event
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        lstDisplay.Items.Clear()    'Clearing the Display
    End Sub
End Class
