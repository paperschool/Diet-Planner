Imports System.Data.OleDb
Imports System.ComponentModel

Module Ingredients

    Dim da As OleDbDataAdapter
    Dim dt As DataTable
    Dim ds As DataSet = New DataSet

    'a stock id variable used to store both sql statements and stock id information
    Dim Ingredient_ID As String

    Sub populate_ingredients()

        Dim SearchQ As String = Main.txtSearchIngredient.Text
        Dim query As Boolean = False

        If SearchQ = "" Then
            query = False
        Else
            query = True
        End If

        Dim ingda As OleDbDataAdapter
        Dim ingdt As DataTable
        Dim ingds As DataSet = New DataSet



        Dim ingredient_list As String

        If query = False Then
            ingredient_list = "SELECT * FROM tblIngredients"
        Else
            ingredient_list = "SELECT * FROM tblIngredients WHERE IngredientName like '%" & SearchQ & "%'"
        End If

        'this defines uses the sql query and connection string
        ingda = New OleDbDataAdapter(ingredient_list, conn)
        'this datadapter is filled with the dataset declared earlier
        ingda.Fill(ingds, "ing")
        'this datatable is filled with data pulled from the dataset
        ingdt = ingds.Tables("ing")

        'this block is used to fill the navigation datagrid view with the data in the dataset
        With Main.dgvIngredients
            .AutoGenerateColumns = True
            .DataSource = ingds
            .DataMember = "ing"

        End With



        Main.dgvIngredients.RowTemplate.Height = 30
        Main.dgvIngredients.RowTemplate.MinimumHeight = 30
        Main.dgvIngredients.AutoResizeRows()
        Main.dgvIngredients.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Main.dgvIngredients.Columns(0).Visible = False
        Main.dgvIngredients.Columns(1).Width = 300


    End Sub

    Sub ingredient_selection(i)

        Dim ingda As OleDbDataAdapter
        Dim ingdt As DataTable
        Dim ingds As DataSet = New DataSet

        Dim ingredientID As String = Main.dgvIngredients.Item(0, i).Value()

        Dim ingredient_content As String = "SELECT * FROM tblIngredients WHERE IngredientID='" & ingredientID & "'"
        'this defines uses the sql query and connection string
        ingda = New OleDbDataAdapter(ingredient_content, conn)
        'this datadapter is filled with the dataset declared earlier
        ingda.Fill(ingds, "ing_con")
        'this datatable is filled with data pulled from the dataset
        ingdt = ingds.Tables("ing_con")

        Main.txtInID.Text = ingds.Tables("ing_con").Rows(0).Item(0).ToString
        Main.txtIngName.Text = ingds.Tables("ing_con").Rows(0).Item(1).ToString
        Main.txtIngQuantity.Text = ingds.Tables("ing_con").Rows(0).Item(2).ToString
        Main.txtIngCalories.Text = ingds.Tables("ing_con").Rows(0).Item(3).ToString
        Main.txtIngProtein.Text = ingds.Tables("ing_con").Rows(0).Item(4).ToString
        Main.txtIngCarbohydrates.Text = ingds.Tables("ing_con").Rows(0).Item(5).ToString
        Main.txtIngFats.Text = ingds.Tables("ing_con").Rows(0).Item(6).ToString


    End Sub

    Sub clear_ingredient_form()

        Dim result As Integer = MessageBox.Show("Are you sure you want to empty the Ingredient profile?", "Warning", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Exit Sub
        ElseIf result = DialogResult.Yes Then
        End If

        For Each i In Main.gbIngredients.Controls
            If (TypeOf i Is NSTextBox) Then
                i.text = ""
            End If
        Next

        Main.txtIngQuantity.Text = 100

        Main.dgvIngredients.Sort(Main.dgvIngredients.Columns(1), ListSortDirection.Ascending)

    End Sub

    Sub Save_Ingredient()

        If Main.txtInID.Text <> "" Then
            Edit_Data()
        Else
            Generate_Ingredient_ID()
            Insert_Data()
        End If

        populate_ingredients_Meal()

        Main.dgvIngredients.Sort(Main.dgvIngredients.Columns(1), ListSortDirection.Ascending)

    End Sub

    Sub Generate_Ingredient_ID()

        'SQL Select Statement Searching for the highest UserID assciated with that user type'
        Ingredient_ID = "SELECT MAX(IngredientID) FROM tblIngredients WHERE IngredientID LIKE '" & "IN" & "%%%%%' "
        'Data adapter defined'
        da = New OleDbDataAdapter(Ingredient_ID, conn)
        'Data adapter told to fill the da   taset
        da.Fill(ds, "tblIngredients")
        'Defining the datatable
        dt = ds.Tables("tblIngredients")
        'Associating User ID to the first item on the first row of the datatabl whilst converting the data to a string'
        Ingredient_ID = dt.Rows(0).Item(0).ToString

        'For defining the first student in the class;
        If Ingredient_ID = "" Then
            'this is some "Fake Data" that is used to calculate the first User ID'
            Ingredient_ID = "IN00000"
            Exit Sub
        End If

        'This strips off the User Type Identified'
        Ingredient_ID = Ingredient_ID.Substring(2, 5)
        'This removes the leading zero's from the stripped string'
        Ingredient_ID = Ingredient_ID.TrimStart("0"c)

        'This catches the User ID string if it is something like "S0000", after all the stripping and formating, would leave nothing, this code repairs it'
        If Ingredient_ID = "" Then
            'Re-Defines the Ingredient_ID to 0'
            Ingredient_ID = 0
        End If

        'Converts the split string into an integer'
        Ingredient_ID = CType(Ingredient_ID, Integer) + 1

        'This block is responsible for adding the appropriate value onto the end of the user ID'
        'This converts the User ID to a string'
        Ingredient_ID = CType(Ingredient_ID, String)
        'Length check's used to identify the different tiers of the user id'
        If Ingredient_ID.Length = 1 Then
            'This combines the User ID's componants into a full id'
            Ingredient_ID = "IN" & "0000" & Ingredient_ID
            'Length check's used to identify the different tiers of the user id'
        ElseIf Ingredient_ID.Length = 2 Then
            'This combines the User ID's componants into a full id'
            Ingredient_ID = "IN" & "000" & Ingredient_ID
            'Length check's used to identify the different tiers of the user id'
        ElseIf Ingredient_ID.Length = 3 Then
            'This combines the User ID's componants into a full id'
            Ingredient_ID = "IN" & "00" & Ingredient_ID
        ElseIf Ingredient_ID.Length = 4 Then
            Ingredient_ID = "IN" & "0" & Ingredient_ID
        Else
            'This combines the User ID's componants into a full id'
            Ingredient_ID = "IN" & Ingredient_ID
        End If

        ds.Clear()

    End Sub

    Sub Insert_Data()

        Dim Stockinsert As String = "INSERT INTO tblIngredients ([IngredientID],[IngredientName],[Quantity],[Calories],[Protein],[Carbohydrates],[Fats]) VALUES (@IngredientID,@IngredientName,@Quantity,@Calories,@Protein,@Carbohydrates,@Fats)"
        Dim StockinsertCommand As New OleDbCommand
        With StockinsertCommand
            .CommandText = Stockinsert
            'this block of parameters matches the stock id, main information and size information with the variables in the sql statement
            .Parameters.AddWithValue("@IngredientID", Ingredient_ID)
            .Parameters.AddWithValue("@IngredientName", Main.txtIngName.Text)
            .Parameters.AddWithValue("@Quantity", Main.txtIngQuantity.Text)
            .Parameters.AddWithValue("@Calories", Main.txtIngCalories.Text)
            .Parameters.AddWithValue("@Protein", Main.txtIngProtein.Text)
            .Parameters.AddWithValue("@Carbohydrates", Main.txtIngCarbohydrates.Text)
            .Parameters.AddWithValue("@Fats", Main.txtIngFats.Text)
            .Connection = conn
            .ExecuteNonQuery()
        End With

        For Each i In Main.gbIngredients.Controls
            If (TypeOf i Is NSTextBox) Then
                i.text = ""
            End If
        Next

        populate_ingredients()

    End Sub

    Sub Edit_Data()


        Dim ingredientupdate As String = "UPDATE tblIngredients SET IngredientName = @IngredientName, Quantity = @Quantity, Calories = @Calories, Protein = @Protein, Carbohydrates = @Carbohydrates, Fats = @Fats WHERE IngredientID = @IngredientID"
        Dim ingredientupdateCommand As New OleDbCommand
        With ingredientupdateCommand
            .CommandText = ingredientupdate
            .Parameters.AddWithValue("@IngredientName", Main.txtIngName.Text)
            .Parameters.AddWithValue("@Quantity", Main.txtIngQuantity.Text)
            .Parameters.AddWithValue("@Calories", Main.txtIngCalories.Text)
            .Parameters.AddWithValue("@Protein", Main.txtIngProtein.Text)
            .Parameters.AddWithValue("@Carbohydrates", Main.txtIngCarbohydrates.Text)
            .Parameters.AddWithValue("@Fats", Main.txtIngFats.Text)
            .Parameters.AddWithValue("@IngredientID", Main.txtInID.Text)
            .Connection = conn
            .ExecuteNonQuery()
        End With


        populate_ingredients()

        Main.dgvIngredients.Sort(Main.dgvIngredients.Columns(1), ListSortDirection.Ascending)

    End Sub

    Sub delete_ingredient_data()

        Dim result As Integer = MessageBox.Show("Are you sure you want to DELETE the Ingredient profile, this may modify other meals?", "Warning", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Exit Sub
        ElseIf result = DialogResult.Yes Then
        End If

        Dim ingredientdelete As String = "DELETE FROM tblIngredients WHERE IngredientID = @IngredientID"
        Dim ingredientdeleteCommand As New OleDbCommand
        With ingredientdeleteCommand
            .CommandText = ingredientdelete
            .Parameters.AddWithValue("@IngredientID", Main.txtInID.Text)
            .Connection = conn
            .ExecuteNonQuery()
        End With

        Main.Reload()

    End Sub

    Sub Ingredient_modifier()

        Dim mod_amount As Integer = Main.cbQuantityModifier.Text

        If mod_amount = Nothing Then
            populate_ingredients()
            Exit Sub
        End If

        Dim ingda As OleDbDataAdapter
        Dim ingdt As DataTable
        Dim ingds As DataSet = New DataSet


        Dim ingredient_list As String = "SELECT * FROM tblIngredients"

        'this defines uses the sql query and connection string
        ingda = New OleDbDataAdapter(ingredient_list, conn)
        'this datadapter is filled with the dataset declared earlier
        ingda.Fill(ingds, "ing")
        'this datatable is filled with data pulled from the dataset
        ingdt = ingds.Tables("ing")

        'this block is used to fill the navigation datagrid view with the data in the dataset
        With Main.dgvIngredients
            .AutoGenerateColumns = True
            .DataSource = ingds
            .DataMember = "ing"
        End With


        For Each row As DataRow In ingdt.Rows

            row.Item(2) = CInt(row.Item(2)) * (mod_amount * 0.01)
            row.Item(3) = CInt(row.Item(3)) * (mod_amount * 0.01)
            row.Item(4) = CInt(row.Item(4)) * (mod_amount * 0.01)
            row.Item(5) = CInt(row.Item(5)) * (mod_amount * 0.01)
            row.Item(6) = CInt(row.Item(6)) * (mod_amount * 0.01)

        Next

        Main.dgvIngredients.AutoResizeRows()
        Main.dgvIngredients.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Main.dgvIngredients.Columns(0).Visible = False

    End Sub

    Sub Search_ingredients()

        populate_ingredients()

        Main.dgvIngredients.Sort(Main.dgvIngredients.Columns(1), ListSortDirection.Ascending)

    End Sub

End Module
