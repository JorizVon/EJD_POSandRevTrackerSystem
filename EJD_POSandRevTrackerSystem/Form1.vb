Imports System.Diagnostics.Contracts
Imports System.Transactions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports K4os.Compression.LZ4.Streams
Imports MySql.Data.MySqlClient
Imports Mysqlx.Notice
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    'LOG IN SET LOCATION CENTER FUNCTION'
    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        CenterControlInForm(LogIngrp)
    End Sub
    Private Sub CenterControlInForm(ctrl As Control)
        Dim formCenterX As Integer = Me.Width \ 2
        Dim formCenterY As Integer = Me.Height \ 2

        Dim controlCenterX As Integer = ctrl.Width \ 2
        Dim controlCenterY As Integer = ctrl.Height \ 2

        ctrl.Location = New Point(formCenterX - controlCenterX, formCenterY - controlCenterY)
    End Sub

    'ADMIN/EMPLOYEE LOG IN FUNCTION'

    'ADMIN/EMPLOYEE LOG IN SHOW PASS OR NOT FUNCTION'
    Private Sub eyeshow_Click(sender As Object, e As EventArgs) Handles eyeshow.Click
        If showPass.Checked Then
            eyeshow.BackgroundImage = Image.FromFile("C:\Users\Gutie\OneDrive\Desktop\eyeclosedicon.png")
            showPass.Checked = False
            passwordtxt.UseSystemPasswordChar = True
        Else
            eyeshow.BackgroundImage = Image.FromFile("C:\Users\Gutie\OneDrive\Desktop\eyeopenedicon.png")
            showPass.Checked = True
            passwordtxt.UseSystemPasswordChar = False
        End If
    End Sub
    Private Sub passwordtxt_TextChanged(sender As Object, e As EventArgs) Handles passwordtxt.TextChanged
        If showPass.Checked Then Return
        passwordtxt.UseSystemPasswordChar = True
    End Sub
    'ADMIN/EMPLOYEE LOG IN BUTTON FUNCTION'
    Private Function Login(ByVal username As String, ByVal password As String) As Integer
        Dim employeeID As Integer = -1
        Dim query As String = "SELECT Employee_ID FROM employee_account WHERE Username = @username AND Password = @password"

        Try
            ' Open the database connection
            conn.Open()

            ' Set up the command
            cmd.Connection = conn
            cmd.CommandText = query
            cmd.Parameters.AddWithValue("@username", username)
            cmd.Parameters.AddWithValue("@password", password)

            ' Execute the query and retrieve the result
            Dim result As Object = cmd.ExecuteScalar()

            ' Check if a result was returned
            If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                employeeID = Convert.ToInt32(result)
            End If
        Catch ex As Exception
            ' Handle any exceptions
            MessageBox.Show("An error occurred while logging in: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' Close the database connection
            conn.Close()
        End Try

        ' Return the employee ID
        Return employeeID
    End Function

    Private Sub loginbtn_Click(sender As Object, e As EventArgs) Handles loginbtn.Click
        ' Check if the entered username and password are the default values
        If usernametxt.Text = "Username" AndAlso passwordtxt.Text = "Password" Then
            ' Clear the default values and perform necessary actions
            usernametxt.Clear()
            passwordtxt.Clear()
            showPass.Checked = False
            LogIngrp.Visible = False
            Adminctrlgrp.Visible = True
            AdminTabs.Visible = True
            LoadDataIntoTransactionTable()
            RevenueTracker()
            LoadSearchTransactionHistory()
        Else
            ' Call the Login function and pass the username and password
            Dim employeeID As Integer = Login(usernametxt.Text, passwordtxt.Text)

            ' Check if the login attempt was successful
            If employeeID <> -1 Then
                ' Display a success message
                MessageBox.Show("Login successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                ' Display an error message
                MessageBox.Show("Incorrect username or password.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub




    'ADMIN NAVIGATION BUTTONS FUNCTION'
    'DASHBOARD TAB BUTTON FUNCTION'
    Private Sub dashboardbtn_Click(sender As Object, e As EventArgs) Handles dashboardbtn.Click
        AdminTabs.SelectedTab = dashboardtab
        LoadDataIntoTransactionTable()
        LoadSearchTransactionHistory()
        TransactionHTable.ClearSelection()
    End Sub

    'VPSTOCKS TAB BUTTON FUNCTION'
    Private Sub VPstocksbtn_Click(sender As Object, e As EventArgs) Handles VPstocksbtn.Click
        AdminTabs.SelectedTab = VPstockstab
        LoadDataIntoVehiclePartTable()
        LoadSearchVparts()
        VehiclePartTable.ClearSelection()
        VPaddoption.Visible = True
        VPaddgrp.Visible = False
    End Sub

    'SERVICE LIST TAB BUTTON FUNCTION'
    Private Sub Servicelistbtn_Click(sender As Object, e As EventArgs) Handles Servicelistbtn.Click
        AdminTabs.SelectedTab = Servicelisttab
        LoadDataIntoServiceListTable()
        LoadSearchServices()
        ServiceListTable.ClearSelection()
        Saddoption.Visible = True
        Saddgrp.Visible = False
    End Sub

    'EMPLOYEE ACCOUNT TAB BUTTON FUNCTION'
    Private Sub EmployeeAccbtn_Click(sender As Object, e As EventArgs) Handles EmployeeaAccbtn.Click
        AdminTabs.SelectedTab = EmployeeAcctab
        LoadDataIntoEmployeeTable()
        LoadSearchEmployees()
        EmployeeTable.ClearSelection()
        Eaddoption.Visible = True
        Eaddgrp.Visible = False
    End Sub

    Private Sub Adminoutbtn_Click(sender As Object, e As EventArgs) Handles Adminoutbtn.Click
        If MsgBox("Are you sure you want to log out?", MsgBoxStyle.YesNo) = vbYes Then
            AdminTabs.Visible = False
            Adminctrlgrp.Visible = False
            LogIngrp.Visible = True
        Else

        End If
    End Sub

    'DESELECTOR FUNCTIONS FOR TABLE IN TABS'
    Private Sub Dashboardtab_MouseDown(sender As Object, e As MouseEventArgs) Handles dashboardtab.MouseDown
        TransactionHTable.ClearSelection()
    End Sub
    Private Sub VPstockstab_MouseDown(sender As Object, e As MouseEventArgs) Handles VPstockstab.MouseDown
        VehiclePartTable.ClearSelection()
    End Sub
    Private Sub Servicelisttab_MouseDown(sender As Object, e As MouseEventArgs) Handles EmployeeAcctab.MouseDown
        ServiceListTable.ClearSelection()
    End Sub
    Private Sub EmployeeAcctab_MouseDown(sender As Object, e As MouseEventArgs) Handles Servicelisttab.MouseDown
        EmployeeTable.ClearSelection()
    End Sub


    '=================================================================================================================================================='

    'DAHBOARD FUNCTIONS'
    'DASHBOARD FUNCTION  |  LOADING TRANSACTION HISTORY TABLE'
    Private Sub LoadDataIntoTransactionTable()
        Dim str As String = "SELECT CONCAT(employee.Fname, ' ', employee.Lname) AS Employee_Name, costumer.Full_Name, transaction.Tdate, transaction.Transaction_ID, transaction.Amount " &
                        "FROM `transaction` " &
                        "LEFT JOIN `costumer` ON transaction.Customer_ID = costumer.Costumer_ID " &
                        "LEFT JOIN `vehicle_ownership` ON costumer.Costumer_ID = vehicle_ownership.Costumer_ID " &
                        "LEFT JOIN `vehicle` ON vehicle_ownership.Vehicle_ID = vehicle.Vehicle_ID " &
                        "LEFT JOIN `service_history` ON vehicle_ownership.Vehicle_ID = service_history.Vehicle_ID " &
                        "LEFT JOIN `service` ON service_history.Service_ID = service.Service_ID " &
                        "LEFT JOIN `service_part_requirements` ON service.Service_ID = service_part_requirements.Service_ID " &
                        "LEFT JOIN `vehicle_part` ON service_part_requirements.VPart_ID = vehicle_part.VPart_ID " &
                        "LEFT JOIN `employee_service_assignment` ON service_history.Service_ID = employee_service_assignment.Service_ID " &
                        "LEFT JOIN `employee` ON employee_service_assignment.Employee_ID = employee.Employee_ID
                         WHERE transaction.Archived = 'Active'"
        Try
            TransactionHTable.Rows.Clear()
            readquery(str)
            With cmdread
                While .Read
                    TransactionHTable.Rows.Add(.GetValue(0), .GetValue(1), .GetValue(2), .GetValue(3), .GetValue($"₱{4}.00"))
                End While
            End With
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
    'DASHBOARD FUNCTION  |  SEARCH FUNCTION'
    Private Sub THsearch_TextChanged(sender As Object, e As EventArgs) Handles searchTH.TextChanged
        LoadTransactionHistoryToSearch(searchTH.Text)
    End Sub
    Private Sub LoadSearchTransactionHistory()
        Dim query As String = "SELECT DISTINCT employee.Fname, employee.Lname, costumer.Full_Name, transaction.Tdate, transaction.Transaction_ID, transaction.Amount " &
                          "FROM `transaction` " &
                          "LEFT JOIN `costumer` ON transaction.Customer_ID = costumer.Costumer_ID " &
                          "LEFT JOIN `vehicle_ownership` ON costumer.Costumer_ID = vehicle_ownership.Costumer_ID " &
                          "LEFT JOIN `vehicle` ON vehicle_ownership.Vehicle_ID = vehicle.Vehicle_ID " &
                          "LEFT JOIN `service_history` ON vehicle_ownership.Vehicle_ID = service_history.Vehicle_ID " &
                          "LEFT JOIN `service` ON service_history.Service_ID = service.Service_ID " &
                          "LEFT JOIN `service_part_requirements` ON service.Service_ID = service_part_requirements.Service_ID " &
                          "LEFT JOIN `vehicle_part` ON service_part_requirements.VPart_ID = vehicle_part.VPart_ID " &
                          "LEFT JOIN `employee_service_assignment` ON service_history.Service_ID = employee_service_assignment.Service_ID " &
                          "LEFT JOIN `employee` ON employee_service_assignment.Employee_ID = employee.Employee_ID " &
                          "WHERE transaction.Archived = 'Active'"
        Try
            searchTH.Items.Clear()
            readquery(query)
            With cmdread
                While .Read
                    searchTH.Items.Add(.GetValue(0))
                    searchTH.Items.Add(.GetValue(1))
                    searchTH.Items.Add(.GetValue(2))
                    searchTH.Items.Add(.GetValue(3))
                    searchTH.Items.Add(.GetValue(4))
                    searchTH.Items.Add(.GetValue(5))
                End While
            End With
        Catch ex As Exception
            MsgBox("Error loading search ComboBoxes: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Function LoadTransactionHistoryToSearch(ByVal keyword As String) As Boolean
        Dim success As Boolean = False
        Dim query As String = "SELECT CONCAT(employee.Fname, ' ', employee.Lname) AS Employee_Name, 
                           costumer.Full_Name, 
                           transaction.Tdate, 
                           transaction.Transaction_ID, 
                           transaction.Amount 
                           FROM `transaction` 
                           LEFT JOIN `costumer` ON transaction.Customer_ID = costumer.Costumer_ID 
                           LEFT JOIN `vehicle_ownership` ON costumer.Costumer_ID = vehicle_ownership.Costumer_ID 
                           LEFT JOIN `vehicle` ON vehicle_ownership.Vehicle_ID = vehicle.Vehicle_ID 
                           LEFT JOIN `service_history` ON vehicle_ownership.Vehicle_ID = service_history.Vehicle_ID 
                           LEFT JOIN `service` ON service_history.Service_ID = service.Service_ID 
                           LEFT JOIN `service_part_requirements` ON service.Service_ID = service_part_requirements.Service_ID 
                           LEFT JOIN `vehicle_part` ON service_part_requirements.VPart_ID = vehicle_part.VPart_ID 
                           LEFT JOIN `employee_service_assignment` ON service_history.Service_ID = employee_service_assignment.Service_ID 
                           LEFT JOIN `employee` ON employee_service_assignment.Employee_ID = employee.Employee_ID 
                           WHERE transaction.Archived = 'Active'"

        ' Append search condition if a keyword is provided
        If Not String.IsNullOrWhiteSpace(keyword) Then
            query &= $" AND (employee.Fname LIKE '%{keyword}%' OR employee.Lname LIKE '%{keyword}%' OR costumer.Full_Name LIKE '%{keyword}%' OR transaction.Tdate LIKE '%{keyword}%' OR transaction.Transaction_ID LIKE '%{keyword}%' OR transaction.Amount LIKE '%{keyword}%')" 
        End If

        Try
            TransactionHTable.Rows.Clear()
            readquery(query)

            With cmdread
                While .Read
                    TransactionHTable.Rows.Add(.GetValue(0), .GetValue(1), .GetValue(2), .GetValue(3), .GetValue(4))
                End While
                success = True
            End With
        Catch ex As Exception
            success = False
            MsgBox("Error loading transaction history data: " & ex.Message, MsgBoxStyle.Critical)
        End Try

        Return success
    End Function

    'DASHBOARD FUNCTION  |  MORE INFO BTN AND ARCHIVE FUNCTION'
    Private Sub TransactionHTable_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles TransactionHTable.CellContentClick
        ' Check if the clicked cell is in the THdetails column
        If e.ColumnIndex = TransactionHTable.Columns("THdetails").Index AndAlso e.RowIndex >= 0 Then
            THdetailsForm.Visible = True
            ' Get the Transaction_ID from the clicked row
            Dim transactionID As String = TransactionHTable.Rows(e.RowIndex).Cells("THtransactionID").Value.ToString()

            ' Construct the SQL query to fetch data based on the Transaction_ID
            Dim query As String = $"SELECT 
                                    transaction.Tdate,
                                    CONCAT(employee.Fname, ' ', employee.Lname) AS Employee_Name, 
                                    costumer.Full_Name, 
                                    vehicle_category.CatName, 
                                    vehicle.Plate_No, 
                                    service.Sname, 
                                    service.Scost, 
                                    vehicle_part.VPname, 
                                    vehicle_part.VPcost, 
                                    vehicle_part.VPquantity, 
                                    transaction.Amount 
                                FROM 
                                    `transaction` 
                                LEFT JOIN 
                                    `costumer` ON transaction.Customer_ID = costumer.Costumer_ID 
                                LEFT JOIN 
                                    `vehicle_ownership` ON costumer.Costumer_ID = vehicle_ownership.Costumer_ID 
                                LEFT JOIN 
                                    `vehicle` ON vehicle_ownership.Vehicle_ID = vehicle.Vehicle_ID 
                                LEFT JOIN 
                                    `service_history` ON vehicle_ownership.Vehicle_ID = service_history.Vehicle_ID 
                                LEFT JOIN 
                                    `service` ON service_history.Service_ID = service.Service_ID 
                                LEFT JOIN 
                                    `service_part_requirements` ON service.Service_ID = service_part_requirements.Service_ID 
                                LEFT JOIN 
                                    `vehicle_part` ON service_part_requirements.VPart_ID = vehicle_part.VPart_ID 
                                LEFT JOIN 
                                    `employee_service_assignment` ON service_history.Service_ID = employee_service_assignment.Service_ID 
                                LEFT JOIN 
                                    `employee` ON employee_service_assignment.Employee_ID = employee.Employee_ID 
                                LEFT JOIN 
                                    `vehicle_category` ON vehicle.Vehicle_ID = vehicle_ownership.Vehicle_ID 
                                WHERE 
                                    transaction.Transaction_ID = {transactionID};"

            Try
                ' Execute the SQL query
                readquery(query)

                ' Read the data from the query result
                If cmdread.Read() Then
                    ' Extract values from the query result
                    Dim TransactionDate As String = cmdread("Tdate")
                    Dim employeeName As String = cmdread("Employee_Name").ToString()
                    Dim fullName As String = cmdread("Full_Name").ToString()
                    Dim vCategoryID As String = cmdread("CatName").ToString()
                    Dim plateNo As String = cmdread("Plate_No").ToString()
                    Dim sName As String = cmdread("Sname").ToString()
                    Dim sCost As String = cmdread("Scost").ToString()
                    Dim vpName As String = cmdread("VPname").ToString()
                    Dim vpCost As String = cmdread("VPcost").ToString()
                    Dim vpQuantity As String = cmdread("VPquantity").ToString()
                    Dim amount As String = cmdread("Amount").ToString()

                    ' Display values in text boxes
                    TextBox1.Text = transactionID
                    TextBox2.Text = employeeName
                    TextBox3.Text = fullName
                    TextBox4.Text = vCategoryID
                    TextBox5.Text = plateNo
                    TextBox6.Text = sName
                    TextBox7.Text = $"₱{sCost}.00"
                    TextBox8.Text = vpName
                    TextBox9.Text = $"₱{vpCost}.00"
                    TextBox10.Text = vpQuantity
                    TextBox11.Text = $"₱{amount}.00"
                    TextBox12.Text = TransactionDate
                End If
            Catch ex As Exception
                MsgBox("Error fetching data: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        ElseIf e.ColumnIndex = TransactionHTable.Columns("THarchive").Index AndAlso e.RowIndex >= 0 Then
            ' Prompt confirmation from the user
            Dim result As DialogResult = MessageBox.Show("Are you sure you want to archive this transaction?", "Confirm Archive", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                ' Get the Transaction_ID from the clicked row
                Dim transactionID As String = TransactionHTable.Rows(e.RowIndex).Cells("THtransactionID").Value.ToString()

                ' Construct the SQL query to update the Archive status based on the Transaction_ID
                Dim query As String = $"UPDATE `transaction` SET Archive = 'Archived' WHERE Transaction_ID = {transactionID};"

                Try
                    ' Execute the SQL query to update the Archive status
                    readquery(query)

                    ' Display a message indicating success
                    MsgBox("Transaction archived successfully.")
                Catch ex As Exception
                    MsgBox("Error archiving transaction: " & ex.Message, MsgBoxStyle.Critical)
                End Try
            End If
        End If
    End Sub

    'DASHBOARD FUNCTIONS | EXIT DETAILS GROUP'
    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox11.Clear()
        THdetailsForm.Visible = False
    End Sub

    'DASHBOARD FUNCTIONS | REVENUE TRACKER'
    Private Sub RevenueTracker()
        Dim i As Integer = 0
        Dim currentDate As DateTime = DateTime.Now
        Dim currentYear As Integer = currentDate.Year
        Dim previousYear As Integer = currentYear - 1
        Dim currentMonth As Integer = currentDate.Month
        Dim previousMonth As Integer = currentMonth - 1


        Dim curryearRev As String = $"SELECT 
                                    SUM(`Amount`) AS `Curryearrev` 
                                FROM 
                                    `transaction` 
                                WHERE 
                                    YEAR(`Tdate`) = {currentYear};"

        Dim prevyearrev As String = $"SELECT 
                                    SUM(`Amount`) AS `prevyearrev` 
                                FROM 
                                    `transaction`  
                                WHERE 
                                    YEAR(`Tdate`) = {previousYear};"

        Dim currmonthrev As String = $"SELECT 
                                    SUM(`Amount`) AS `currmonthrev` 
                                FROM 
                                    `transaction` 
                                WHERE 
                                    YEAR(`Tdate`) = {currentYear} AND MONTH(`Tdate`) = {currentMonth};"

        Dim prevmonthrev As String = $"SELECT 
                                    SUM(`Amount`) AS `prevmonthrev` 
                                FROM 
                                    `transaction` 
                                WHERE 
                                    YEAR(`Tdate`) = {currentYear} AND MONTH(`Tdate`) = {previousMonth};"

        Try
            ' Execute the query for the current year
            readquery(curryearRev)
            If cmdread.Read() Then
                Dim currentYearRevenue As String = cmdread("Curryearrev").ToString()
                latestyearrev.Text = $"₱{currentYearRevenue}.00"
            End If

            ' Execute the query for the previous year
            readquery(prevyearrev)
            If cmdread.Read() Then
                Dim previousYearRevenue As String = cmdread("prevyearrev").ToString()
                lastyearrev.Text = $"₱{previousYearRevenue}.00"
            End If

            ' Execute the query for the current month
            readquery(currmonthrev)
            If cmdread.Read() Then
                Dim currentMonthRevenue As String = cmdread("currmonthrev").ToString()
                latestmonthrev.Text = $"₱{currentMonthRevenue}.00"
            End If

            ' Execute the query for the previous month
            readquery(prevmonthrev)
            If cmdread.Read() Then
                Dim previousMonthRevenue As String = cmdread("prevmonthrev").ToString()
                Lastmonthrev.Text = $"₱{previousMonthRevenue}.00"
            End If
        Catch ex As Exception
            ' Handle exceptions
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub


    '=================================================================================================================================================='

    'VEHICLE PART STOCKS FUNCTIONS'

    'VEHICLE PART STOCKS FUNCTION  |  LOADING VEHICLE PART TABLE'
    Private Sub LoadDataIntoVehiclePartTable()
        Dim str As String = "SELECT VPart_ID, VPname, VPcost, VPquantity FROM vehicle_part"
        Try
            VehiclePartTable.Rows.Clear()
            readquery(str)
            With cmdread
                While .Read
                    VehiclePartTable.Rows.Add(.GetValue(0), .GetValue(1), .GetValue(2), .GetValue(3))
                End While
                LoadSearchVparts()
            End With
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    'VEHICLE PART STOCKS FUNCTION  |  SEARCH FUNCTION'
    Private Sub searchVP_TextChanged(sender As Object, e As EventArgs) Handles VPsearch.TextChanged
        LoadVehicleTableToVPsearch(VPsearch.Text)
    End Sub
    'VP SEARCH FUNCTION LOAD SEARCH COMBO BOX'
    Private Sub LoadSearchVparts()
        Dim s As String = "SELECT VPart_ID, VPname, VPcost, VPquantity FROM vehicle_part"
        Try
            VPsearch.Items.Clear()
            readquery(s)
            With cmdread
                While .Read
                    VPsearch.Items.Add(.GetValue(0))
                    VPsearch.Items.Add(.GetValue(1))
                    VPsearch.Items.Add(.GetValue(2))
                    VPsearch.Items.Add(.GetValue(3))
                End While
            End With
        Catch ex As Exception
            MsgBox("Error loading search ComboBoxes: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
    'VP SEARCH FUNCTION LOAD VP TABLE TO COMBO BOX FOR SEARCH'
    Private Function LoadVehicleTableToVPsearch(ByVal VPsearch As String) As Boolean
        Dim success As Boolean = False
        Dim s As String
        Dim i As Integer = 0
        ' IF VPSEARCH HAS DATA
        If Not String.IsNullOrWhiteSpace(VPsearch) Then
            s = "SELECT VPart_ID, VPname, VPcost, VPquantity FROM vehicle_part WHERE VPart_ID LIKE '%" & VPsearch & "%' OR VPname LIKE '%" & VPsearch & "%' OR VPcost LIKE '%" & VPsearch & "%' OR VPquantity LIKE '%" & VPsearch & "%'"
        Else
            s = "SELECT VPart_ID, VPname, VPcost, VPquantity FROM vehicle_part"
        End If

        Try
            VehiclePartTable.Rows.Clear()
            readquery(s)
            With cmdread
                While .Read
                    VehiclePartTable.Rows.Add(.GetValue(0), .GetValue(1), .GetValue(2), .GetValue(3))
                    i = i + 1
                End While
                success = True
            End With
        Catch ex As Exception
            success = False
        End Try
        Return success
    End Function

    'VEHICLE PART STOCKS FUNCTION  | ADD FUNCTION'
    Private Sub VPadd()
        Dim sql As String

        ' IF ALL TEXTBOXES ARE FILLED
        If String.IsNullOrEmpty(VPnametxt.Text) OrElse String.IsNullOrEmpty(VPcosttxt.Text) OrElse String.IsNullOrEmpty(VPquantitytxt.Text) Then
            MessageBox.Show("Please fill in all fields.", "NOTICE!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' IF COST ONLY HAS DIGITS
        If Not IsNumeric(VPcosttxt.Text) Then
            MessageBox.Show("Please enter only digits for VPcost.", "NOTICE!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            VPcosttxt.Clear()
            VPcosttxt.Focus()
            Return
        End If

        ' IF QUANTITY ONLY HAS DIGITS
        If Not IsNumeric(VPquantitytxt.Text) Then
            MessageBox.Show("Please enter only digits for VPquantity.", "NOTICE!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            VPquantitytxt.Clear()
            VPquantitytxt.Focus()
            Return
        End If

        sql = "INSERT INTO vehicle_part (VPname, VPcost, VPquantity) VALUES ('" & VPnametxt.Text & "', '" & VPcosttxt.Text & "', '" & VPquantitytxt.Text & "');"

        Try
            ' CONFIRMATION TO ADDING
            If MsgBox("Are you certain you want to add these data?", MsgBoxStyle.YesNo) = vbYes Then
                VPnametxt.Clear()
                VPcosttxt.Clear()
                VPquantitytxt.Clear()
                readquery(sql)
                LoadDataIntoVehiclePartTable()
                If MsgBox("Do you want to continue adding?", MsgBoxStyle.YesNo) = vbNo Then
                    VPaddgrp.Visible = False
                    VPaddoption.Visible = True
                Else
                    VPnametxt.Focus()
                End If
            Else
                MsgBox("Data Not Saved")
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    'VEHICLE PART STOCKS FUNCTION  | ADD BUTTON FUNCTION'
    Private Sub VPaddoption_Click(sender As Object, e As EventArgs) Handles VPaddoption.Click
        VPaddgrp.Visible = True
        VPupdatebtn.Visible = False
        VPaddoption.Visible = False
        VPaddbtn.Visible = True
    End Sub
    Private Sub VPaddbtn_Click(sender As Object, e As EventArgs) Handles VPaddbtn.Click
        VPadd()
        VPupdatebtn.Visible = False
        VPaddbtn.Visible = True
    End Sub

    'VEHICLE PART STOCKS FUNCTION  | UPDATE BUTTON FUNCTION'
    Private Sub VPupdatebtn_Click(sender As Object, e As EventArgs) Handles VPupdatebtn.Click
        Dim rowIndex = VehiclePartTable.CurrentCell.RowIndex
        'RETRIEVING ORIGINAL DATA'
        Dim originalName = VehiclePartTable.Rows(rowIndex).Cells("VPname").Value.ToString
        Dim originalCost = VehiclePartTable.Rows(rowIndex).Cells("VPcost").Value.ToString
        Dim originalQuantity = VehiclePartTable.Rows(rowIndex).Cells("VPquantity").Value.ToString

        ' IF NO CHANGES WERE MADE
        If VPnametxt.Text = originalName AndAlso VPcosttxt.Text = originalCost AndAlso VPquantitytxt.Text = originalQuantity Then
            MsgBox("No changes were made to the data.")
            Return
        End If

        ' IF ALL TEXTBOX HAVE DATA
        If String.IsNullOrEmpty(VPnametxt.Text) OrElse String.IsNullOrEmpty(VPcosttxt.Text) OrElse String.IsNullOrEmpty(VPquantitytxt.Text) Then
            MessageBox.Show("Please fill in all fields.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' IF VPCOST AND VPQUANTITY ONLY HAS DIGITS INPUTTED
        If Not IsNumeric(VPcosttxt.Text) OrElse Not IsNumeric(VPquantitytxt.Text) Then
            MessageBox.Show("Please enter only digits for VPcost and VPquantity.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        'UPDATING QUERY
        Dim idValue = VehiclePartTable.Rows(rowIndex).Cells("VPart_ID").Value.ToString
        Dim query = $"UPDATE vehicle_part SET VPname = '{VPnametxt.Text}', VPcost = {VPcosttxt.Text}, VPquantity = {VPquantitytxt.Text} WHERE VPart_ID = {idValue}"

        Try
            readquery(query)
            MsgBox("Record updated successfully.")
            LoadDataIntoVehiclePartTable()
            VPnametxt.Clear()
            VPcosttxt.Clear()
            VPquantitytxt.Clear()

            VPaddbtn.Visible = True
            VPupdatebtn.Visible = False
            VPcancelUpbtn.Visible = False
            VPaddgrp.Visible = False
        Catch ex As Exception
            MsgBox("Error updating record: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub VPcancelUpbtn_Click(sender As Object, e As EventArgs) Handles VPcancelUpbtn.Click
        VPnametxt.Clear()
        VPcosttxt.Clear()
        VPquantitytxt.Clear()
        VPaddgrp.Visible = False
        VPaddoption.Visible = True
    End Sub

    'VEHICLE PART STOCKS FUNCTION  | UPDATE AND DELETE FUNCTION'
    Private Sub VehiclePartTable_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles VehiclePartTable.CellContentClick
        ' Check if the clicked cell is in the "VPedit" column
        If e.ColumnIndex = VehiclePartTable.Columns("VPedit").Index AndAlso e.RowIndex >= 0 Then
            Dim rowIndex = e.RowIndex

            VPnametxt.Text = VehiclePartTable.Rows(rowIndex).Cells("VPname").Value.ToString
            VPcosttxt.Text = VehiclePartTable.Rows(rowIndex).Cells("VPcost").Value.ToString
            VPquantitytxt.Text = VehiclePartTable.Rows(rowIndex).Cells("VPquantity").Value.ToString

            VPaddbtn.Visible = False
            VPupdatebtn.Visible = True
            VPcancelUpbtn.Visible = True
            VPaddgrp.Visible = True

            VehiclePartTable.ClearSelection()

        ElseIf e.ColumnIndex = VehiclePartTable.Columns("VPdelete").Index AndAlso e.RowIndex >= 0 Then
            Dim result = MsgBox("Are you sure you want to delete this record?", MsgBoxStyle.YesNo, "Delete Record Confirmation")

            If result = MsgBoxResult.Yes Then
                Dim idValue = VehiclePartTable.Rows(e.RowIndex).Cells("VPart_ID").Value.ToString
                Dim query = $"DELETE FROM vehicle_part WHERE VPart_ID = {idValue}"

                Try
                    readquery(query)
                    MsgBox("Record deleted successfully.")
                    LoadDataIntoVehiclePartTable()
                Catch ex As Exception
                    MsgBox("Error deleting record: " & ex.Message, MsgBoxStyle.Critical)
                End Try
            Else
                VehiclePartTable.ClearSelection()
            End If
        End If
    End Sub


    '=================================================================================================================================================='


    'SERVICE LIST FUNCTION  |  LOADING SERVICE LIST TABLE'
    Private Sub LoadDataIntoServiceListTable()
        Dim str As String = "SELECT Service_ID, Sname, Scost FROM service"
        Try
            ServiceListTable.Rows.Clear()
            readquery(str)
            With cmdread
                While .Read
                    ServiceListTable.Rows.Add(.GetValue(0), .GetValue(1), .GetValue(2))
                End While
                LoadSearchServices()
            End With
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
    'SERVICE LIST FUNCTION  |  SEARCH FUNCTION'
    Private Sub searchS_TextChanged(sender As Object, e As EventArgs) Handles Ssearch.TextChanged
        LoadServiceTableToSsearch(Ssearch.Text)
    End Sub
    'SERVICE SEARCH FUNCTION LOAD SEARCH COMBO BOX'
    Private Sub LoadSearchServices()
        Dim s As String = "SELECT Service_ID, Sname, Scost FROM service"
        Try
            Ssearch.Items.Clear()
            readquery(s)
            With cmdread
                While .Read
                    Ssearch.Items.Add(.GetValue(0))
                    Ssearch.Items.Add(.GetValue(1))
                    Ssearch.Items.Add(.GetValue(2))
                End While
            End With
        Catch ex As Exception
            MsgBox("Error loading search ComboBoxes: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    'SERVICE SEARCH FUNCTION LOAD SERVICE TABLE TO COMBO BOX FOR SEARCH'
    Private Function LoadServiceTableToSsearch(ByVal Ssearch As String) As Boolean
        Dim success As Boolean = False
        Dim s As String
        Dim i As Integer = 0
        ' IF VPSEARCH HAS DATA
        If Not String.IsNullOrWhiteSpace(Ssearch) Then
            s = "SELECT Service_ID, Sname, Scost FROM service WHERE Service_ID LIKE '%" & Ssearch & "%' OR Sname LIKE '%" & Ssearch & "%' OR Scost LIKE '%" & Ssearch & "%'"
        Else
            s = "SELECT Service_ID, Sname, Scost FROM service"
        End If

        Try
            ServiceListTable.Rows.Clear()
            readquery(s)
            With cmdread
                While .Read
                    ServiceListTable.Rows.Add(.GetValue(0), .GetValue(1), .GetValue(2))
                    i = i + 1
                End While
                success = True
            End With
        Catch ex As Exception
            success = False
        End Try
        Return success
    End Function

    ' SERVICE FUNCTION | ADD FUNCTION
    Private Sub AddService()
        Dim sql As String

        ' IF ALL TEXTBOXES ARE FILLED
        If String.IsNullOrEmpty(Snametxt.Text) OrElse String.IsNullOrEmpty(Scosttxt.Text) Then
            MessageBox.Show("Please fill in all fields.", "NOTICE!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' IF COST ONLY HAS DIGITS
        If Not IsNumeric(Scosttxt.Text) Then
            MessageBox.Show("Please enter only digits for Scost.", "NOTICE!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Scosttxt.Clear()
            Scosttxt.Focus()
            Return
        End If

        ' SQL query for inserting service data
        sql = "INSERT INTO service (Sname, Scost) VALUES ('" & Snametxt.Text & "', '" & Scosttxt.Text & "');"

        Try
            ' Confirmation before adding data
            If MsgBox("Are you certain you want to add this data?", MsgBoxStyle.YesNo) = vbYes Then
                Snametxt.Clear()
                Scosttxt.Clear()
                readquery(sql)
                LoadDataIntoServiceListTable()
                If MsgBox("Do you want to continue adding?", MsgBoxStyle.YesNo) = vbNo Then
                    Saddgrp.Visible = False
                    Saddoption.Visible = True
                Else
                    Snametxt.Focus()
                End If
            Else
                MsgBox("Data Not Saved")
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    'VEHICLE PART STOCKS FUNCTION  | ADD BUTTON FUNCTION'
    Private Sub Saddoption_Click(sender As Object, e As EventArgs) Handles Saddoption.Click
        Saddgrp.Visible = True
        Supdatebtn.Visible = False
        Saddoption.Visible = False
        Saddbtn.Visible = True
        Snametxt.Focus()
    End Sub
    Private Sub Saddbtn_Click(sender As Object, e As EventArgs) Handles Saddbtn.Click
        AddService()
        Supdatebtn.Visible = False
        Saddbtn.Visible = True
    End Sub
    ' SERVICE FUNCTION | UPDATE FUNCTION
    Private Sub Supdatebtn_Click(sender As Object, e As EventArgs) Handles Supdatebtn.Click
        Dim rowIndex As Integer = ServiceListTable.CurrentCell.RowIndex
        ' Retrieve original data
        Dim originalName As String = ServiceListTable.Rows(rowIndex).Cells("Sname").Value.ToString()
        Dim originalCost As String = ServiceListTable.Rows(rowIndex).Cells("Scost").Value.ToString()

        ' Check if no changes were made
        If Snametxt.Text = originalName AndAlso Scosttxt.Text = originalCost Then
            MsgBox("No changes were made to the data.")
            Return
        End If

        ' Check if all textboxes have data
        If String.IsNullOrEmpty(Snametxt.Text) OrElse String.IsNullOrEmpty(Scosttxt.Text) Then
            MessageBox.Show("Please fill in all fields.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' Check if Scosttxt only has digits inputted
        If Not IsNumeric(Scosttxt.Text) Then
            MessageBox.Show("Please enter only digits for Scost.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Scosttxt.Clear()
            Scosttxt.Focus()
            Return
        End If

        ' Updating query
        Dim idValue As String = ServiceListTable.Rows(rowIndex).Cells("Service_ID").Value.ToString()
        Dim query As String = $"UPDATE service SET Sname = '{Snametxt.Text}', Scost = {Scosttxt.Text} WHERE Service_ID = {idValue}"

        Try
            readquery(query)
            MsgBox("Record updated successfully.")
            LoadDataIntoServiceListTable()
            ServiceListTable.ClearSelection()
            Snametxt.Clear()
            Scosttxt.Clear()
            Saddbtn.Visible = True
            Supdatebtn.Visible = False
            Saddgrp.Visible = False
        Catch ex As Exception
            MsgBox("Error updating record: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Sub ScancelUpbtn_Click(sender As Object, e As EventArgs) Handles ScancelUpbtn.Click
        ServiceListTable.ClearSelection()
        Snametxt.Clear()
        Scosttxt.Clear()
        Saddgrp.Visible = False
        Saddoption.Visible = True
    End Sub
    'SERVICE FUNCTION  | UPDATE AND DELETE FUNCTION'
    Private Sub ServiceListTable_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles ServiceListTable.CellContentClick
        ' Check if the clicked cell is in the "VPedit" column
        If e.ColumnIndex = ServiceListTable.Columns("Sedit").Index AndAlso e.RowIndex >= 0 Then
            Dim rowIndex As Integer = e.RowIndex

            Snametxt.Text = ServiceListTable.Rows(rowIndex).Cells("Sname").Value.ToString()
            Scosttxt.Text = ServiceListTable.Rows(rowIndex).Cells("Scost").Value.ToString()

            Saddbtn.Visible = False
            Supdatebtn.Visible = True
            ScancelUpbtn.Visible = True
            Saddgrp.Visible = True

            ServiceListTable.ClearSelection()

        ElseIf e.ColumnIndex = ServiceListTable.Columns("Sdelete").Index AndAlso e.RowIndex >= 0 Then
            Dim result As MsgBoxResult = MsgBox("Are you sure you want to delete this record?", MsgBoxStyle.YesNo, "Delete Record Confirmation")

            If result = MsgBoxResult.Yes Then
                Dim idValue As String = ServiceListTable.Rows(e.RowIndex).Cells("Service_ID").Value.ToString()
                Dim query As String = $"DELETE FROM service WHERE Service_ID = {idValue}"

                Try
                    readquery(query)
                    MsgBox("Record deleted successfully.")
                    LoadDataIntoServiceListTable()
                Catch ex As Exception
                    MsgBox("Error deleting record: " & ex.Message, MsgBoxStyle.Critical)
                End Try
            Else
                ServiceListTable.ClearSelection()
            End If
        End If
    End Sub


    '============================================================================================================================================'


    'EMPLOYEE FUNCTIONS'

    'EMPLOYEE FUNCTION  |  LOADING EMPLOYEE TABLE'
    Private Sub LoadDataIntoEmployeeTable()
        Dim str As String = "SELECT employee.Employee_ID, CONCAT(employee.Fname, ' ', employee.Lname) AS FullName, employee.Contact_Info, employee.Job_Title, employee_account.Username, employee_account.Password " &
                    "FROM employee_account, employee " &
                    "WHERE employee_account.Employee_ID = employee.Employee_ID and employee.Status = 'Active'"
        Try
            EmployeeTable.Rows.Clear()
            readquery(str)
            With cmdread
                While .Read
                    EmployeeTable.Rows.Add(.GetValue(0), .GetValue(1), .GetValue(2), .GetValue(3), .GetValue(4), .GetValue(5))
                End While
                LoadSearchEmployees()
            End With
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    'EMPLOYEE FUNCTION  |  SEARCH FUNCTION'
    Private Sub Esearch_TextChanged(sender As Object, e As EventArgs) Handles Esearch.TextChanged
        LoadEmployeeTableToSearch(Esearch.Text)
    End Sub

    'EMPLOYEE SEARCH FUNCTION LOAD SEARCH COMBO BOX'
    Private Sub LoadSearchEmployees()
        Dim s As String = "SELECT employee.Employee_ID, employee.Fname, employee.Lname, employee.Contact_Info, employee.Job_Title, employee_account.Username, employee_account.Password " &
                        "FROM employee_account, employee " &
                        "WHERE employee_account.Employee_ID = employee.Employee_ID and employee.Status = 'Active'"
        Try
            Esearch.Items.Clear()
            readquery(s)
            With cmdread
                While .Read
                    Esearch.Items.Add(.GetValue(0))
                    Esearch.Items.Add(.GetValue(1))
                    Esearch.Items.Add(.GetValue(2))
                    Esearch.Items.Add(.GetValue(3))
                    Esearch.Items.Add(.GetValue(4))
                    Esearch.Items.Add(.GetValue(5))
                    Esearch.Items.Add(.GetValue(6))
                End While
            End With
        Catch ex As Exception
            MsgBox("Error loading search ComboBoxes: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    'EMPLOYEE SEARCH FUNCTION LOAD EMPLOYEE TABLE TO COMBO BOX FOR SEARCH'
    Private Function LoadEmployeeTableToSearch(ByVal EmployeeSearch As String) As Boolean
        Dim success As Boolean = False
        Dim s As String
        Dim i As Integer = 0
        ' IF EMPLOYEESEARCH HAS DATA
        If Not String.IsNullOrWhiteSpace(EmployeeSearch) Then
            s = "SELECT employee.Employee_ID, CONCAT(employee.Fname, ' ', employee.Lname) AS FullName, employee.Contact_Info, employee.Job_Title, employee_account.Username, employee_account.Password " &
                "FROM employee_account, employee " &
                "WHERE employee_account.Employee_ID = employee.Employee_ID and employee.Status = 'Active' AND (employee.Employee_ID LIKE '%" & EmployeeSearch & "%' OR CONCAT(employee.Fname, ' ', employee.Lname) LIKE '%" & EmployeeSearch & "%' OR employee.Contact_Info LIKE '%" & EmployeeSearch & "%' OR employee.Job_Title LIKE '%" & EmployeeSearch & "%' OR employee_account.Username LIKE '%" & EmployeeSearch & "%' OR employee_account.Password LIKE '%" & EmployeeSearch & "%')"
        Else
            s = "SELECT employee.Employee_ID, CONCAT(employee.Fname, ' ', employee.Lname) AS FullName, employee.Contact_Info, employee.Job_Title, employee_account.Username, employee_account.Password " &
                "FROM employee_account, employee " &
                "WHERE employee_account.Employee_ID = employee.Employee_ID and employee.Status = 'Active'"
        End If

        Try
            EmployeeTable.Rows.Clear()
            readquery(s)
            With cmdread
                While .Read
                    EmployeeTable.Rows.Add(.GetValue(0), .GetValue(1), .GetValue(2), .GetValue(3), .GetValue(4), .GetValue(5))
                    i = i + 1
                End While
                success = True
            End With
        Catch ex As Exception
            success = False
        End Try
        Return success
    End Function

    'EMPLOYEE ACCOUNT FUNCTION  | ADD FUNCTION'
    Private Sub Employeeadd()
        Dim sql As String

        ' IF ANY TEXTBOX IS EMPTY
        If String.IsNullOrEmpty(Fnametxt.Text) OrElse String.IsNullOrEmpty(Lnametxt.Text) OrElse String.IsNullOrEmpty(Econtacttxt.Text) OrElse String.IsNullOrEmpty(JobTitletxt.Text) OrElse String.IsNullOrEmpty(Eusernametxt.Text) OrElse String.IsNullOrEmpty(Epasswordtxt.Text) Then
            MessageBox.Show("Please fill in all fields.", "NOTICE!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' QUERY INSERT
        sql = $"INSERT INTO employee (Fname, Lname, Contact_Info, Job_Title, Status) VALUES ('{Fnametxt.Text}', '{Lnametxt.Text}', '{Econtacttxt.Text}', '{JobTitletxt.Text}', 'Active');"

        Dim sqlAccount As String = $"INSERT INTO employee_account (Employee_ID, Username, Password) VALUES (LAST_INSERT_ID(), '{Eusernametxt.Text}', '{Epasswordtxt.Text}');"

        Try
            ' CONFIRMATION TO ADDING
            If MsgBox("Are you certain you want to add these data?", MsgBoxStyle.YesNo) = vbYes Then
                readquery(sql)
                readquery(sqlAccount)
                MsgBox("Record added successfully.")
                LoadDataIntoEmployeeTable()

                ' Clear textboxes and reset visibility
                Fnametxt.Clear()
                Lnametxt.Clear()
                Econtacttxt.Clear()
                JobTitletxt.Clear()
                Eusernametxt.Clear()
                Epasswordtxt.Clear()

                Eaddgrp.Visible = False
                Eaddoption.Visible = True
            Else
                MsgBox("Data Not Saved")
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    'EMPLOYEE ACCOUNT FUNCTION  | ADD BUTTON FUNCTION'
    Private Sub Eaddoption_Click(sender As Object, e As EventArgs) Handles Eaddoption.Click
        Eaddgrp.Visible = True
        Eupdatebtn.Visible = False
        Eaddoption.Visible = False
        Eaddbtn.Visible = True
    End Sub
    Private Sub Eaddbtn_Click(sender As Object, e As EventArgs) Handles Eaddbtn.Click
        Employeeadd()
        Eupdatebtn.Visible = False
        Eaddbtn.Visible = True
    End Sub

    'EMPLOYEE ACCOUNT FUNCTION  | UPDATE BUTTON FUNCTION'
    Private Sub Eupdatebtn_Click(sender As Object, e As EventArgs) Handles Eupdatebtn.Click
        Dim rowIndex = EmployeeTable.CurrentCell.RowIndex

        ' RETRIEVING ORIGINAL DATA
        Dim originalName As String = EmployeeTable.Rows(rowIndex).Cells("Fullname").Value.ToString()
        Dim originalContact As String = EmployeeTable.Rows(rowIndex).Cells("Econtact").Value.ToString()
        Dim originalJobTitle As String = EmployeeTable.Rows(rowIndex).Cells("JobTitle").Value.ToString()
        Dim originalUsername As String = EmployeeTable.Rows(rowIndex).Cells("Eusername").Value.ToString()
        Dim originalPassword As String = EmployeeTable.Rows(rowIndex).Cells("Epassword").Value.ToString()

        'SPLITTING FNAME AND LNAME TOGETHER
        Dim originalNameParts() As String = originalName.Split(" ")
        Dim originalFname As String = originalNameParts(0)
        Dim originalLname As String = originalNameParts(1)

        ' IF NO CHANGES WERE MADE
        If Fnametxt.Text = originalFname AndAlso Lnametxt.Text = originalLname AndAlso Econtacttxt.Text = originalContact AndAlso JobTitletxt.Text = originalJobTitle AndAlso Eusernametxt.Text = originalUsername AndAlso Epasswordtxt.Text = originalPassword Then
            MsgBox("No changes were made to the data.")
            Return
        End If

        ' IF ALL TEXTBOXES HAVE DATA
        If String.IsNullOrEmpty(Fnametxt.Text) OrElse String.IsNullOrEmpty(Lnametxt.Text) OrElse String.IsNullOrEmpty(Econtacttxt.Text) OrElse String.IsNullOrEmpty(JobTitletxt.Text) OrElse String.IsNullOrEmpty(Eusernametxt.Text) OrElse String.IsNullOrEmpty(Epasswordtxt.Text) Then
            MessageBox.Show("Please fill in all fields.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' UPDATING QUERY
        Dim idValue As String = EmployeeTable.Rows(rowIndex).Cells("Employee_ID").Value.ToString()

        Dim query As String = $"UPDATE employee SET Fname = '{Fnametxt.Text}', Lname = '{Lnametxt.Text}', Contact_Info = '{Econtacttxt.Text}', Job_Title = '{JobTitletxt.Text}' WHERE Employee_ID = {idValue}"
        Dim queryAccount As String = $"UPDATE employee_account SET Username = '{Eusernametxt.Text}', Password = '{Epasswordtxt.Text}' WHERE Employee_ID = {idValue}"

        Try
            readquery(query)
            readquery(queryAccount)

            MsgBox("Record updated successfully.")
            LoadDataIntoEmployeeTable()

            Fnametxt.Clear()
            Lnametxt.Clear()
            Econtacttxt.Clear()
            JobTitletxt.Clear()
            Eusernametxt.Clear()
            Epasswordtxt.Clear()

            Eaddgrp.Visible = False
            Eaddoption.Visible = True
        Catch ex As Exception
            MsgBox("Error updating record: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Sub EcancelUpbtn_Click(sender As Object, e As EventArgs) Handles EcancelUpbtn.Click
        Fnametxt.Clear()
        Lnametxt.Clear()
        Econtacttxt.Clear()
        JobTitletxt.Clear()
        Eusernametxt.Clear()
        Epasswordtxt.Clear()
        Eaddgrp.Visible = False
        Eaddoption.Visible = True
    End Sub

    'EMPLOYEE ACCOUNT FUNCTION  | UPDATE AND DELETE FUNCTION'
    Private Sub EmployeeTable_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles EmployeeTable.CellContentClick
        'IF CELL CLICKED IN Eedit COLUMN
        If e.ColumnIndex = EmployeeTable.Columns("Eedit").Index AndAlso e.RowIndex >= 0 Then
            Dim rowIndex = e.RowIndex

            Dim fullName As String = EmployeeTable.Rows(rowIndex).Cells("Fullname").Value.ToString()
            Dim lastName As String = fullName.Split(" ").Last() ' Extracting the last name from the full name

            Fnametxt.Text = fullName.Replace(lastName, "").Trim() ' Getting the first name
            Lnametxt.Text = lastName
            Econtacttxt.Text = EmployeeTable.Rows(rowIndex).Cells("Econtact").Value.ToString()
            JobTitletxt.Text = EmployeeTable.Rows(rowIndex).Cells("JobTitle").Value.ToString()
            Eusernametxt.Text = EmployeeTable.Rows(rowIndex).Cells("Eusername").Value.ToString()
            Epasswordtxt.Text = EmployeeTable.Rows(rowIndex).Cells("Epassword").Value.ToString()

            Eaddbtn.Visible = False
            Eupdatebtn.Visible = True
            EcancelUpbtn.Visible = True
            Eaddgrp.Visible = True

            EmployeeTable.ClearSelection()
            'IF CELL CLICKED IN Edelete COLUMN
        ElseIf e.ColumnIndex = EmployeeTable.Columns("Edelete").Index AndAlso e.RowIndex >= 0 Then
            Dim result = MsgBox("Are you sure you want to delete this record?", MsgBoxStyle.YesNo, "Delete Record Confirmation")

            If result = MsgBoxResult.Yes Then
                Dim idValue = EmployeeTable.Rows(e.RowIndex).Cells("Employee_ID").Value.ToString()
                Dim query = $"DELETE FROM employee WHERE Employee_ID = {idValue}"

                Try
                    readquery(query)
                    MsgBox("Record deleted successfully.")
                    LoadDataIntoEmployeeTable()
                Catch ex As Exception
                    MsgBox("Error deleting record: " & ex.Message, MsgBoxStyle.Critical)
                End Try
            Else
                EmployeeTable.ClearSelection()
            End If
            'IF CELL CLICKED IN Earchive COLUMN
        ElseIf e.ColumnIndex = EmployeeTable.Columns("Earchive").Index AndAlso e.RowIndex >= 0 Then
            Dim result = MsgBox("Are you sure you want to archive this record?", MsgBoxStyle.YesNo, "Archive Record Confirmation")

            If result = MsgBoxResult.Yes Then
                Dim idValue = EmployeeTable.Rows(e.RowIndex).Cells("Employee_ID").Value.ToString()
                Dim query = $"UPDATE employee SET Status = 'Archived' WHERE Employee_ID = {idValue}"
                Try
                    readquery(query)
                    MsgBox("Record archived successfully.")
                    LoadDataIntoEmployeeTable()
                Catch ex As Exception
                    MsgBox("Error archiving record: " & ex.Message, MsgBoxStyle.Critical)
                End Try
            Else
                EmployeeTable.ClearSelection()
            End If
        End If
    End Sub

End Class
