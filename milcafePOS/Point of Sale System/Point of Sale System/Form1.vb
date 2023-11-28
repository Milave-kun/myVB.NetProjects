Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Drawing.Printing

Public Class Form1
    ' Define a global connection string variable
    Dim connectionString As String = "Data Source=DESKTOP-808QF8L\SQLEXPRESS;Initial Catalog=salesInfo;Integrated Security=True"

    Public SidePrice As Double
    Dim Total As Double
    Dim Qty As Double = 0
    Dim Discount As Double
    Dim CustomerName As String ' To store the customer's name
    Dim invoiceNumber As String ' To store the generated invoice number
    Dim bs As New BindingSource
    Dim isCalculated As Boolean = False ' Boolean flag to track whether the calculation has been done

    'Form 1 Load
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Font = DataGridView2.Font
        BindData()
        BindData2()
        roundCorners(Me)

    End Sub

    'BindData1
    Private Sub BindData()
        Dim query As String = "SELECT * FROM salesRecord"
        Using con As SqlConnection = New SqlConnection(connectionString)
            Using cmd As SqlCommand = New SqlCommand(query, con)
                Using da As New SqlDataAdapter()
                    da.SelectCommand = cmd
                    Using dt As New DataTable()
                        da.Fill(dt)
                        DataGridView1.DataSource = dt
                    End Using
                End Using
            End Using
        End Using
    End Sub

    'BindData2
    Private Sub BindData2()
        Dim query As String = "SELECT * FROM manageOrders"
        Using con As SqlConnection = New SqlConnection(connectionString)
            Using cmd As SqlCommand = New SqlCommand(query, con)
                Using da As New SqlDataAdapter()
                    da.SelectCommand = cmd
                    Using dt As New DataTable()
                        da.Fill(dt)
                        DataGridView2.DataSource = dt ' Set the DataSource of DataGridView2 to the DataTable
                    End Using
                End Using
            End Using
        End Using

        ' Check if the "Cancel" button column already exists
        If DataGridView2.Columns("CancelColumn") Is Nothing Then
            ' Create a DataGridViewButtonColumn for the "Cancel" button
            Dim cancelColumn As New DataGridViewButtonColumn()
            cancelColumn.Name = "CancelColumn"
            cancelColumn.HeaderText = "Cancel"
            cancelColumn.Text = "Cancel"
            cancelColumn.UseColumnTextForButtonValue = True
            DataGridView2.Columns.Add(cancelColumn)
        End If

        ' Check if the "Confirm" button column already exists
        If DataGridView2.Columns("ConfirmColumn") Is Nothing Then
            ' Create a DataGridViewButtonColumn for the "Confirm" button
            Dim confirmColumn As New DataGridViewButtonColumn()
            confirmColumn.Name = "ConfirmColumn"
            confirmColumn.HeaderText = "Confirm"
            confirmColumn.Text = "Confirm"
            confirmColumn.UseColumnTextForButtonValue = True
            DataGridView2.Columns.Add(confirmColumn)
        End If

        ' Set the button's style for the "Cancel" and "Confirm" button columns
        If DataGridView2.Columns("CancelColumn") IsNot Nothing Then
            Dim cancelColumnIndex As Integer = DataGridView2.Columns("CancelColumn").Index
            For Each row As DataGridViewRow In DataGridView2.Rows
                row.Cells(cancelColumnIndex).Style.ForeColor = Color.White ' Text color
                row.Cells(cancelColumnIndex).Style.BackColor = Color.Red ' Background color
            Next
        End If

        If DataGridView2.Columns("ConfirmColumn") IsNot Nothing Then
            Dim confirmColumnIndex As Integer = DataGridView2.Columns("ConfirmColumn").Index
            For Each row As DataGridViewRow In DataGridView2.Rows
                row.Cells(confirmColumnIndex).Style.ForeColor = Color.White ' Text color
                row.Cells(confirmColumnIndex).Style.BackColor = Color.Green ' Background color
            Next
        End If
    End Sub

    'Private Function ExtracInvoiceNumber
    Private Function ExtractInvoiceNumber(orderText As String) As String
        Dim match = Regex.Match(orderText, "Invoice Number: (INV-\d+)")
        If match.Success Then
            Return match.Groups(1).Value
        Else
            Return String.Empty
        End If
    End Function

    'Calculate & Place Button
    Private Sub btnCalculate_Click(sender As Object, e As EventArgs) Handles btnCalculate.Click
        ' Declare internal variables for holding taxed value, customer name, table number, and discount
        Dim Tax As Double = 0.115
        Dim DiscountTotal As Double
        Dim price As Double ' You need to declare the price
        Dim quantity As Double ' You need to declare the quantity

        If ListBoxContainsCalculatedData() Then
            MessageBox.Show("This order has already been calculated. Please start a new order.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If ListBox1.Items.Count = 0 Then
            ' Display a warning message if the ListBox is empty
            MessageBox.Show("No items in the order. Please add items before calculating.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            ' You can use an input dialog or another method to obtain this information
            Dim customerName As String = InputBox("Enter Customer Name:", "Customer Information")


            ' Generate an invoice number based on the current date and time
            Dim invoiceNumber As String = "INV-" & DateTime.Now.ToString("yyyyMMddHHmmss")

            ' Create a dynamic description based on selected items
            Dim descriptionValue As String = "" & vbCrLf
            Dim itemsOrdered As New List(Of String) ' Create a list to store the items ordered

            ' Calculate the price and quantity based on your data
            price = Total ' You need to calculate or set the price based on your data
            quantity = Qty ' You need to calculate or set the quantity based on your data

            ' Display Subtotal, calculate Total, and Tax
            ListBox1.Items.Add("-----------------------------------------------------------")
            ListBox1.Items.Add("Subtotal: " + CStr(Total.ToString("C2")))
            ListBox1.Items.Add("Total items ordered: " + CStr(Qty))
            DiscountTotal = Total * Discount
            Total = Total - DiscountTotal
            Tax = Total * Tax
            Total += Tax

            ' Display quantity of dishes ordered, discount, tax, and total
            ListBox1.Items.Add("")
            ListBox1.Items.Add("Discount: " + CStr(Discount * 100) + "%")
            ListBox1.Items.Add("Tax: " + CStr(Tax.ToString("C2")))
            ListBox1.Items.Add("Total Due: " + CStr(Total.ToString("C2")))

            ' Calls the custom dialog box, "CustomerInfo," for inputting customer name and table number, then displays that info and also the current date and time
            ListBox1.Items.Add("-----------------------------------------------------------")
            ListBox1.Items.Add("Invoice No.: " + invoiceNumber)
            If Not String.IsNullOrEmpty(customerName) Then
                ' Add the customer's name to the ListBox
                ListBox1.Items.Add("Customer Name: " + customerName)
            End If
            ListBox1.Items.Add("Today's Date: " + CStr(Today) + "   " + CStr(TimeOfDay))

            ' Displays a thank-you message
            ListBox1.Items.Add("")
            ListBox1.Items.Add("Thanks for stopping by Milcafe")
            ListBox1.Items.Add("")

            ' Insert data into the database, including the generated invoice number and customer name
            InsertInvoiceData(invoiceNumber, price, quantity, Total, Date.Now, customerName)
        End If

        ' Set the flag to indicate that the calculation has been done
        isCalculated = True

        BindData()
        BindData2()

    End Sub

    'Private Function ListBoxContainsCalculatedData
    Private Function ListBoxContainsCalculatedData() As Boolean
        ' Check if the ListBox contains a specific line that indicates it has been calculated
        Dim calculatedLine As String = "Total Due:"

        For Each item As Object In ListBox1.Items
            If item.ToString().Contains(calculatedLine) Then
                Return True
            End If
        Next

        Return False
    End Function

    'InsertInvoiceData
    Private Sub InsertInvoiceData(invoiceNumber As String, price As Double, quantity As Double, totalValue As Double, salesDate As Date, customerNames As String)
        Dim sqlInsert As String = "INSERT INTO manageOrders (InvoiceNo, Price, Quantity, Total, SalesDate, CustomerName) " &
             "VALUES (@InvoiceNo, @Price, @Quantity, @Total, @SalesDate, @CustomerName)"

        Using connection As New SqlConnection(connectionString)
            connection.Open()
            Using command As New SqlCommand(sqlInsert, connection)
                ' Set the parameter values, including the invoice number, customer name, and other values
                command.Parameters.AddWithValue("@InvoiceNo", invoiceNumber)
                command.Parameters.AddWithValue("@Price", price)
                command.Parameters.AddWithValue("@Quantity", quantity)
                command.Parameters.AddWithValue("@Total", totalValue)
                command.Parameters.AddWithValue("@SalesDate", salesDate)
                command.Parameters.AddWithValue("@CustomerName", customerNames)

                ' Execute the SQL command
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    'BtnHotCappucino
    Private Sub btnHotCappucino_Click_1(sender As Object, e As EventArgs) Handles btnHotCappucino.Click
        If ListBoxContainsCalculatedData() Then
            MessageBox.Show("This order has already been calculated. Please start a new order.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim ItemTotal As String = CStr("₱165") ' Coffee price
        Total = Total + 165
        Qty = Qty + 1
        ListBox1.Items.Add("1 Cappucino " + "   " + CStr(ItemTotal) + Space(12) + Total.ToString("C2"))

    End Sub


    Private Sub btnIcedCappucino_Click_1(sender As Object, e As EventArgs) Handles btnIcedCappucino.Click
        If ListBoxContainsCalculatedData() Then
            MessageBox.Show("This order has already been calculated. Please start a new order.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim ItemTotal As String = CStr("₱165") ' Coffee price
        Total = Total + 165
        Qty = Qty + 1
        ListBox1.Items.Add("1 Iced Cappucino " + "   " + CStr(ItemTotal) + Space(12) + Total.ToString("C2"))
    End Sub

    Private Sub btnHotBlackCoffee_Click(sender As Object, e As EventArgs) Handles btnHotBlackCoffee.Click
        If ListBoxContainsCalculatedData() Then
            MessageBox.Show("This order has already been calculated. Please start a new order.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim ItemTotal As String = CStr("₱140") ' Coffee price
        Total = Total + 140
        Qty = Qty + 1
        ListBox1.Items.Add("1 Black Coffee " + "   " + CStr(ItemTotal) + Space(12) + Total.ToString("C2"))
    End Sub

    Private Sub btnIcedBlackCoffee_Click(sender As Object, e As EventArgs) Handles btnIcedBlackCoffee.Click
        If ListBoxContainsCalculatedData() Then
            MessageBox.Show("This order has already been calculated. Please start a new order.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim ItemTotal As String = CStr("₱145") ' Coffee price
        Total = Total + 145
        Qty = Qty + 1
        ListBox1.Items.Add("1 Iced Black Coffee " + "   " + CStr(ItemTotal) + Space(12) + Total.ToString("C2"))
    End Sub

    Private Sub btnHotSpanishLatte_Click(sender As Object, e As EventArgs) Handles btnHotSpanishLatte.Click
        If ListBoxContainsCalculatedData() Then
            MessageBox.Show("This order has already been calculated. Please start a new order.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim ItemTotal As String = CStr("₱150") ' Coffee price
        Total = Total + 150
        Qty = Qty + 1
        ListBox1.Items.Add("1 Spanish Latte " + "   " + CStr(ItemTotal) + Space(12) + Total.ToString("C2"))
    End Sub

    Private Sub GunaImageButton7_Click(sender As Object, e As EventArgs) Handles GunaImageButton7.Click
        If ListBoxContainsCalculatedData() Then
            MessageBox.Show("This order has already been calculated. Please start a new order.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim ItemTotal As String = CStr("₱150") ' Coffee price
        Total = Total + 150
        Qty = Qty + 1
        ListBox1.Items.Add("1 Iced Spanish Latte " + "   " + CStr(ItemTotal) + Space(12) + Total.ToString("C2"))
    End Sub

    Private Sub GunaImageButton3_Click(sender As Object, e As EventArgs) Handles GunaImageButton3.Click
        If ListBoxContainsCalculatedData() Then
            MessageBox.Show("This order has already been calculated. Please start a new order.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim ItemTotal As String = CStr("₱170") ' Coffee price
        Total = Total + 170
        Qty = Qty + 1
        ListBox1.Items.Add("1 Matcha Latte " + "   " + CStr(ItemTotal) + Space(12) + Total.ToString("C2"))
    End Sub

    Private Sub GunaImageButton6_Click(sender As Object, e As EventArgs) Handles GunaImageButton6.Click
        If ListBoxContainsCalculatedData() Then
            MessageBox.Show("This order has already been calculated. Please start a new order.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim ItemTotal As String = CStr("₱175") ' Coffee price
        Total = Total + 175
        Qty = Qty + 1
        ListBox1.Items.Add("1 Iced Matcha Latte " + "   " + CStr(ItemTotal) + Space(12) + Total.ToString("C2"))
    End Sub

    Private Sub GunaImageButton4_Click(sender As Object, e As EventArgs) Handles GunaImageButton4.Click
        Dim ItemTotal As String = CStr("₱140") ' Coffee price
        If ListBoxContainsCalculatedData() Then
            MessageBox.Show("This order has already been calculated. Please start a new order.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Total = Total + 140
        Qty = Qty + 1
        ListBox1.Items.Add("1 Mocha " + "   " + CStr(ItemTotal) + Space(12) + Total.ToString("C2"))
    End Sub

    Private Sub GunaImageButton5_Click(sender As Object, e As EventArgs) Handles GunaImageButton5.Click
        If ListBoxContainsCalculatedData() Then
            MessageBox.Show("This order has already been calculated. Please start a new order.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        Dim ItemTotal As String = CStr("₱140") ' Coffee price
        Total = Total + 140
        Qty = Qty + 1
        ListBox1.Items.Add("1 Iced Mocha " + "   " + CStr(ItemTotal) + Space(12) + Total.ToString("C2"))
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        BindData()
    End Sub

    'Form Round Corners
    Private Sub roundCorners(obj As Form)
        obj.FormBorderStyle = FormBorderStyle.None
        obj.BackColor = Color.White ' Set the background color to white

        Dim DGP As New Drawing2D.GraphicsPath
        DGP.StartFigure()

        ' Top left corner
        DGP.AddArc(New Rectangle(0, 0, 16, 16), 180, 90)
        DGP.AddLine(8, 0, obj.Width - 8, 0)

        ' Top right corner
        DGP.AddArc(New Rectangle(obj.Width - 16, 0, 16, 16), -90, 90)
        DGP.AddLine(obj.Width, 8, obj.Width, obj.Height - 8)

        ' Bottom right corner
        DGP.AddArc(New Rectangle(obj.Width - 16, obj.Height - 16, 16, 16), 0, 90)
        DGP.AddLine(obj.Width - 8, obj.Height, 8, obj.Height)

        ' Bottom left corner
        DGP.AddArc(New Rectangle(0, obj.Height - 16, 16, 16), 90, 90)
        DGP.CloseFigure()

        obj.Region = New Region(DGP)
    End Sub

    Private Sub btnRemoveItem_Click(sender As Object, e As EventArgs) Handles btnRemoveItem.Click
        ' Check if the calculation has already been done
        If isCalculated Then
            ' Display an error message if the calculation is done
            MessageBox.Show("The order has already been calculated. You can't remove items now.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Check if an item is selected in ListBox1
        If ListBox1.SelectedIndex = -1 Then
            MessageBox.Show("Please select an item to remove.")
            Return
        End If


        ' Remove the selected item from the ListBox
        ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
    End Sub

    Private Sub btnNewOrder_Click(sender As Object, e As EventArgs) Handles btnNewOrder.Click
        ' Clear the entire ListBox
        ListBox1.Items.Clear()

        ' Reset the order-related variables to their initial values
        Total = 0
        Qty = 0
        Discount = 0

        ' Reset the isCalculated flag to False
        isCalculated = False
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Using con As SqlConnection = New SqlConnection(connectionString)
            ' Get the search text from the TextBox (txtSearch)
            Dim searchText As String = txtSearch.Text.Trim()

            If Not String.IsNullOrEmpty(searchText) Then
                ' Use a parameterized query to filter by invoice number or customer name
                Dim query As String = "SELECT * FROM manageOrders WHERE InvoiceNo LIKE @SearchText OR CustomerName LIKE @SearchText"

                Using cmd As SqlCommand = New SqlCommand(query, con)
                    ' Use the '%' wildcard for partial matching
                    cmd.Parameters.AddWithValue("@SearchText", "%" & searchText & "%")

                    Using da As New SqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        da.Fill(dt)

                        ' Set the DataGridView's DataSource to the filtered data
                        DataGridView2.DataSource = dt
                    End Using
                End Using
            Else
                ' If the search text is empty, retrieve all records from the database
                Dim query As String = "SELECT * FROM manageOrders"

                Using cmd As SqlCommand = New SqlCommand(query, con)
                    Using da As New SqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        da.Fill(dt)

                        ' Set the DataGridView's DataSource to all records
                        DataGridView2.DataSource = dt
                    End Using
                End Using
            End If
        End Using
    End Sub


    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then ' Check if a valid row was clicked

            If e.ColumnIndex = DataGridView2.Columns("CancelColumn").Index AndAlso e.RowIndex >= 0 Then
                ' The "Cancel" button was clicked, handle the cancellation logic here
                Dim invoiceCell As DataGridViewCell = DataGridView2.Rows(e.RowIndex).Cells("InvoiceNo")
                If invoiceCell IsNot Nothing AndAlso invoiceCell.Value IsNot Nothing Then
                    Dim invoiceNumber As String = invoiceCell.Value.ToString()
                    CancelOrder(invoiceNumber)
                End If
            ElseIf e.ColumnIndex = DataGridView2.Columns("ConfirmColumn").Index AndAlso e.RowIndex >= 0 Then
                ' The "Confirm" button was clicked, handle the confirmation logic here
                Dim invoiceCell As DataGridViewCell = DataGridView2.Rows(e.RowIndex).Cells("InvoiceNo")
                If invoiceCell IsNot Nothing AndAlso invoiceCell.Value IsNot Nothing Then
                    Dim invoiceNumber As String = invoiceCell.Value.ToString()
                    ConfirmOrder(invoiceNumber)
                End If
            End If
        End If

        BindData()
        BindData2()
    End Sub

    Private Sub ConfirmOrder(invoiceNumber As String)

        Dim confirmationMessage As String = "Are you sure you want to confirm this order?"
        Dim confirmationTitle As String = "Confirm Order"

        ' Display a confirmation message box
        Dim confirmResult As DialogResult = MessageBox.Show(confirmationMessage, confirmationTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If confirmResult = DialogResult.Yes Then
            Dim transferQuery As String = "INSERT INTO salesRecord (InvoiceNo, Price, Quantity, Total, SalesDate, CustomerName) " &
        "SELECT InvoiceNo, Price, Quantity, Total, SalesDate, CustomerName " &
        "FROM manageOrders WHERE InvoiceNo = @InvoiceNo"

            ' Delete the record from the manageOrder table
            Dim deleteQuery As String = "DELETE FROM manageOrders WHERE InvoiceNo = @InvoiceNo"

            Using connection As New SqlConnection(connectionString)
                connection.Open()

                Using transferCommand As New SqlCommand(transferQuery, connection)
                    transferCommand.Parameters.AddWithValue("@InvoiceNo", invoiceNumber)
                    transferCommand.ExecuteNonQuery()
                End Using

                Using deleteCommand As New SqlCommand(deleteQuery, connection)
                    deleteCommand.Parameters.AddWithValue("@InvoiceNo", invoiceNumber)
                    deleteCommand.ExecuteNonQuery()
                End Using
            End Using

            BindData()
            BindData2()
        Else
            ' User chose not to confirm the order, no action needed.
        End If
    End Sub

    Private Sub CancelOrder(invoiceNumber As String)
        Dim confirmationMessage As String = "Are you sure you want to cancel this order?"
        Dim confirmationTitle As String = "Cancel Order"

        Dim confirmResult As DialogResult = MessageBox.Show(confirmationMessage, confirmationTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If confirmResult = DialogResult.Yes Then
            ' Delete the data from the database using the extracted invoiceNumber
            Dim sqlDelete As String = "DELETE FROM manageOrders WHERE InvoiceNo = @InvoiceNo"

            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Using command As New SqlCommand(sqlDelete, connection)
                    ' Set the parameter for InvoiceNo
                    command.Parameters.AddWithValue("@InvoiceNo", invoiceNumber)

                    ' Execute the DELETE statement
                    command.ExecuteNonQuery()
                End Using
            End Using

            ' You can also remove the canceled order from DataGridView2, but this depends on your specific requirements
            ' DataGridView2.Rows.RemoveAt(DataGridView2.CurrentRow.Index)

            ' Refresh the data in DataGridView1 (if needed)
            BindData()
            BindData2()
        Else
            ' User chose not to cancel the order, no action needed.
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnFilter_Click(sender As Object, e As EventArgs) Handles btnFilter.Click
        'Retrieve the selected start And end dates from DateTimePickers
        Dim startDate As DateTime = DateTimePickerStart.Value
        Dim endDate As DateTime = DateTimePickerEnd.Value

        ' Filter the data in the DataGridView using the date range
        FilterSalesRecordByDateRange(startDate, endDate)
    End Sub

    Private Sub FilterSalesRecordByDateRange(startDate As DateTime, endDate As DateTime)
        ' Set the start date to the beginning of the day
        startDate = startDate.Date

        Using con As SqlConnection = New SqlConnection(connectionString)
            ' Use a parameterized query to filter by date range
            Dim query As String = "SELECT * FROM salesRecord WHERE SalesDate >= @StartDate AND SalesDate <= @EndDate"
            Using cmd As SqlCommand = New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@StartDate", startDate)
                cmd.Parameters.AddWithValue("@EndDate", endDate.Date.AddDays(1).AddSeconds(-1))

                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)

                    ' Set the DataGridView's DataSource to the filtered data
                    DataGridView1.DataSource = dt
                End Using
            End Using
        End Using
    End Sub


    Private Sub btnPrintReceipt_Click(sender As Object, e As EventArgs) Handles btnPrintReceipt.Click
        ' Check if the ListBox is empty
        If ListBox1.Items.Count = 0 Then
            ' Display an error message if there are no items
            MessageBox.Show("No items to print. Please add items to the order first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ' Exit the function
        End If

        ' Proceed with printing the receipt

        'button calls "Private Sub PrintDocument1_PrintPage" and prints the receipt on default printer
        PrintDocument1.PrinterSettings.Copies = 1
        PrintDocument1.Print()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim printFont As New Font("Arial", 12)
        Dim lineHeight As Single = printFont.GetHeight(e.Graphics)
        Dim yPos As Single = 100
        Dim count As Integer = 0

        For Each item As Object In ListBox1.Items
            e.Graphics.DrawString(item.ToString(), printFont, Brushes.Black, 100, yPos)
            yPos += lineHeight
            count += 1

            ' You can adjust the yPos and count to control the spacing between items if needed.
            ' For example, yPos += 20 for more space between items.
        Next

        ' If there are more items to print, set e.HasMorePages to True to continue to the next page.
        If count < ListBox1.Items.Count Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
    End Sub

    'Button Print DataGridView
    Private Sub btnPrintDataGridView_Click(sender As Object, e As EventArgs) Handles btnPrintDataGridView.Click
        PrintPreviewDialog1.Document = PrintDocument2
        PrintPreviewDialog1.WindowState = FormWindowState.Maximized
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private mRow As Integer = 0
    Private newPage As Boolean = True

    'PrintDocument2
    Private Sub PrintDocument2_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument2.PrintPage
        Dim format As New StringFormat
        format.Alignment = StringAlignment.Center
        e.Graphics.DrawString("Sales Record", New Font("Century Gothic", 20, FontStyle.Bold),
                              Brushes.Black, New Point(415, 40), format)

        Dim fmt As StringFormat = New StringFormat(StringFormatFlags.LineLimit)
        fmt.LineAlignment = StringAlignment.Center
        fmt.Trimming = StringTrimming.EllipsisCharacter
        fmt.Alignment = StringAlignment.Center

        Dim y As Integer = 100
        Dim x As Integer = 32  ' Adjust this value to make the table wider
        Dim h As Integer = 0
        Dim rc As Rectangle
        Dim row As DataGridViewRow

        ' Print headers on each page
        row = DataGridView1.Rows(0)
        x = 32  ' Adjust this value to make the table wider
        For Each cell As DataGridViewCell In row.Cells
            If cell.Visible Then
                rc = New Rectangle(x, y, 130, cell.Size.Height)  ' Adjust the cell width here
                e.Graphics.FillRectangle(Brushes.LightGray, rc)
                e.Graphics.DrawRectangle(Pens.Black, rc)

                e.Graphics.DrawString(DataGridView1.Columns(cell.ColumnIndex).HeaderText,
                                      DataGridView1.Font, Brushes.Black, rc, fmt)

                x += rc.Width
                h = Math.Max(h, rc.Height)
            End If
        Next
        y += h

        ' Print rows
        newPage = False
        Dim displayNow As Integer
        For displayNow = mRow To DataGridView1.RowCount - 1
            row = DataGridView1.Rows(displayNow)
            x = 32  ' Adjust this value to make the table wider
            h = 0

            For Each cell As DataGridViewCell In row.Cells
                If cell.Visible Then
                    rc = New Rectangle(x, y, 130, cell.Size.Height)  ' Adjust the cell width here
                    e.Graphics.DrawRectangle(Pens.Black, rc)

                    fmt.Alignment = StringAlignment.Near
                    rc.Offset(10, 0)

                    e.Graphics.DrawString(cell.FormattedValue.ToString(),
                                          DataGridView1.Font, Brushes.Black, rc, fmt)

                    x += rc.Width
                    h = Math.Max(h, rc.Height)
                End If
            Next
            y += h
        Next

        mRow += 1

        If y + h > e.MarginBounds.Bottom Then
            e.HasMorePages = True
            newPage = True
        Else
            e.HasMorePages = False
            mRow = 0
        End If
    End Sub

End Class