Imports System.Data.OleDb

Module Meal_Designer


    Dim da As OleDbDataAdapter
    Dim dt As DataTable
    Dim ds As DataSet = New DataSet

    Dim originalProtein As String
    Dim originalCarbohydrates As String
    Dim originalFats As String
    Dim originalCalories As String

    Sub Clear_meal()

        Main.txtMealBuilderID.Text = ""
        Main.txtMealName.Text = ""
        Main.cbMealType.Text = Nothing
        Main.txtMealCalories.Text = ""
        Main.txtMealProtein.Text = ""
        Main.txtMealCarbohydrates.Text = ""
        Main.txtMealFats.Text = ""


        empty_meal_table()


    End Sub

    Sub meal_main_info()


        If Main.txtMealName.Text = "" Then
            Main.gbMealBuilder.Enabled = False
            Main.gbMealBuilder.BackColor = Color.DimGray
        Else
            Main.gbMealBuilder.BackColor = Color.FromArgb(47, 47, 47)
            Main.gbMealBuilder.Enabled = True
        End If


    End Sub

    Sub populate_ingredients_Meal()


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


        With Main.dgvMealIngredients
            .AutoGenerateColumns = True
            .DataSource = ingds
            .DataMember = "ing"
        End With

        Main.dgvMealIngredients.Columns(0).Visible = False
        Main.dgvMealIngredients.Columns(2).Visible = False
        Main.dgvMealIngredients.Columns(3).Visible = False
        Main.dgvMealIngredients.Columns(4).Visible = False
        Main.dgvMealIngredients.Columns(5).Visible = False
        Main.dgvMealIngredients.Columns(6).Visible = False

        Main.dgvMealIngredients.AutoResizeRows()
        Main.dgvMealIngredients.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    Sub Build_Meal(i)

        Dim ingredientID As String = Main.dgvMealIngredients.Item(0, i).Value()

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

        For Each row As DataRow In ingdt.Rows
            If row.Item(0) = ingredientID Then
                Main.txtMealIngID.Text = row.Item(0)
                Main.txtMealIngName.Text = row.Item(1)
                Main.txtMealIngCalories.Text = row.Item(3)
                originalCalories = row.Item(3)
                Main.txtMealIngProtein.Text = row.Item(4)
                originalProtein = row.Item(4)
                Main.txtMealIngCarbohydrates.Text = row.Item(5)
                originalCarbohydrates = row.Item(5)
                Main.txtMealIngFats.Text = row.Item(6)
                originalFats = row.Item(6)
                Exit For
            End If
        Next

    End Sub

    Sub Gram_Modify()

        Dim gram_mod As String = Main.txtMealIngGrams.Text

        If gram_mod = "" Then
            gram_mod = "100"
        End If

        Main.txtMealIngProtein.Text = CDbl(originalProtein) * (CDbl(gram_mod) * 0.01)
        Main.txtMealIngCarbohydrates.Text = CDbl(originalCarbohydrates) * (CDbl(gram_mod) * 0.01)
        Main.txtMealIngFats.Text = CDbl(originalFats) * (CDbl(gram_mod) * 0.01)
        Main.txtMealIngCalories.Text = CDbl(originalCalories) * (CDbl(gram_mod) * 0.01)

    End Sub

    Dim mealdt As New DataTable

    Sub meal_table_creation()

        mealdt.Columns.Add("IngredientID", Type.GetType("System.String"))
        mealdt.Columns.Add("Ingredient Name", Type.GetType("System.String"))
        mealdt.Columns.Add("Quantity", Type.GetType("System.String"))
        mealdt.Columns.Add("Calories", Type.GetType("System.String"))
        mealdt.Columns.Add("Protein", Type.GetType("System.String"))
        mealdt.Columns.Add("Carbohydrates", Type.GetType("System.String"))
        mealdt.Columns.Add("Fats", Type.GetType("System.String"))

    End Sub

    Sub empty_meal_table()

        For x As Integer = mealdt.Rows.Count - 1 To 0 Step -1
            mealdt.Rows(x).Delete()

        Next

    End Sub

    Sub Add_ingredient_Meal()

        For i As Integer = 0 To Main.dgvMealBuilder.Rows.Count - 1
            If Main.dgvMealBuilder.Item(0, i).Value = Main.txtMealIngID.Text Then
                MsgBox("Duplicate Ingredient " & Main.dgvMealBuilder.Item(1, i).Value & " , Please Select one of each ingredient")
                Exit Sub
            End If
        Next

        If Main.txtMealIngGrams.Text = "" Then
            MsgBox("Ingredient Cannot Have 0 Grams")
            Exit Sub
        End If

        Dim mealdr As DataRow = mealdt.NewRow

        mealdr("IngredientID") = Main.txtMealIngID.Text
        mealdr("Ingredient Name") = Main.txtMealIngName.Text
        mealdr("Quantity") = Main.txtMealIngGrams.Text
        mealdr("Calories") = Main.txtMealIngCalories.Text
        mealdr("Protein") = Main.txtMealIngProtein.Text
        mealdr("Carbohydrates") = Main.txtMealIngCarbohydrates.Text
        mealdr("Fats") = Main.txtMealIngFats.Text

        mealdt.Rows.Add(mealdr)

        With Main.dgvMealBuilder
            .AutoGenerateColumns = True
            .DataSource = mealdt
        End With


        Main.dgvMealBuilder.AutoResizeRows()
        Main.dgvMealBuilder.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Main.dgvMealBuilder.Columns(0).Visible = False

        macro_counter()

    End Sub

    Sub remove_ingredient_meal(i)

        Dim ingredient As String = Main.dgvMealBuilder.Item(0, i).Value()

        For x As Integer = mealdt.Rows.Count - 1 To 0 Step -1
            If mealdt.Rows(x).Item(0) = ingredient  Then
                mealdt.Rows(x).Delete()
            End If
        Next

        macro_counter()

    End Sub

    Sub macro_counter()

        Dim num_ingredients As Integer = Main.dgvMealBuilder.Rows.Count


        Dim Calories As Double = 0
        Dim Protein As Double = 0
        Dim Carbohydrates As Double = 0
        Dim Fats As Double = 0


        For i As Integer = 0 To num_ingredients - 1
            Calories = Calories + Main.dgvMealBuilder.Item(3, i).Value()
            Protein = Protein + Main.dgvMealBuilder.Item(4, i).Value()
            Carbohydrates = Carbohydrates + Main.dgvMealBuilder.Item(5, i).Value()
            Fats = Fats + Main.dgvMealBuilder.Item(6, i).Value()
        Next


        Main.txtMealCalories.Text = Calories
        Main.txtMealProtein.Text = Protein
        Main.txtMealCarbohydrates.Text = Carbohydrates
        Main.txtMealFats.Text = Fats


    End Sub

    Sub delete_built_meal()

        Dim result As Integer = MessageBox.Show("Are you sure you want to DELETE the Meal profile?", "Warning", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Exit Sub
        ElseIf result = DialogResult.Yes Then
        End If

        Dim ingredientdelete As String = "DELETE FROM tblMeal WHERE MealID = @MealID"
        Dim ingredientdeleteCommand As New OleDbCommand
        With ingredientdeleteCommand
            .CommandText = ingredientdelete
            .Parameters.AddWithValue("@Meal", Main.txtMealBuilderID.Text)
            .Connection = conn
            .ExecuteNonQuery()
        End With

        MsgBox("Meal Deleted")

        Main.Reload()

    End Sub

    Sub save_meal_validation()

        If Main.txtMealName.Text = "" Then
            MsgBox("Provide Meal Name")
            val = True
        ElseIf Main.cbMealType.Text = "" Then
            MsgBox("Provide Meal Type")
            val = True
        ElseIf mealdt.Rows.Count = 0 Then
            MsgBox("Please select at least one ingredient")
            val = True
        Else
            For Each row In mealdt.Rows
                If row.item(2) = "" Then
                    MsgBox("Meal Cannot Have 0 Grams")
                    val = True
                End If
            Next
        End If

    End Sub

    Dim val As Boolean = False

    Sub save_built_meal()

        val = False

        save_meal_validation()

        If val = True Then
            Exit Sub
        End If

        Dim Meal_Name As String = Main.txtMealIngName.Text
        Dim Meal_Type As String = Main.cbMealType.Text

        Dim num_ingredients As Integer = Main.dgvMealBuilder.Rows.Count

        Dim ingredients(num_ingredients - 1) As String

        Dim modifiers(num_ingredients - 1) As Integer

        For i As Integer = 0 To num_ingredients - 1
            ingredients(i) = Main.dgvMealBuilder.Item(0, i).Value()
            modifiers(i) = Main.dgvMealBuilder.Item(2, i).Value()
        Next

        If Main.txtMealBuilderID.Text = "" Then

            Generate_Meal_ID()

            For x As Integer = 0 To num_ingredients - 1

                Dim InsertMeal As String = "INSERT INTO tblMeal ([MealID],[IngredientID],[MealName],[Modifier],[MealType]) VALUES (@MealID,@IngredientID,@MealName,@Modifier,@MealType)"
                Dim InsertMealCommand As New OleDbCommand
                With InsertMealCommand
                    .CommandText = InsertMeal
                    'this block of parameters matches the stock id, main information and size information with the variables in the sql statement
                    .Parameters.AddWithValue("@MealID", MealID)
                    .Parameters.AddWithValue("@IngredientID", ingredients(x))
                    .Parameters.AddWithValue("@MealName", Main.txtMealName.Text)
                    .Parameters.AddWithValue("@Modifier", modifiers(x))
                    .Parameters.AddWithValue("@MealType", Main.cbMealType.Text)
                    .Connection = conn
                    .ExecuteNonQuery()
                End With
            Next

            Main.Reload()

            MsgBox("Meal Saved")

        Else

            Dim result As Integer = MessageBox.Show("Are you sure you want to Edit the Meal profile?", "Warning", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                Exit Sub
            ElseIf result = DialogResult.Yes Then
            End If

            For x As Integer = 0 To num_ingredients - 1

                Try
                    Dim ingredientdelete As String = "DELETE FROM tblMeal WHERE MealID = @MealID"
                    Dim ingredientdeleteCommand As New OleDbCommand
                    With ingredientdeleteCommand
                        .CommandText = ingredientdelete
                        .Parameters.AddWithValue("@Meal", Main.txtMealBuilderID.Text)
                        .Connection = conn
                        .ExecuteNonQuery()
                    End With
                Catch ex As Exception

                End Try

            Next

            For x As Integer = 0 To num_ingredients - 1

                Dim InsertMeal As String = "INSERT INTO tblMeal ([MealID],[IngredientID],[MealName],[Modifier],[MealType]) VALUES (@MealID,@IngredientID,@MealName,@Modifier,@MealType)"
                Dim InsertMealCommand As New OleDbCommand
                With InsertMealCommand
                    .CommandText = InsertMeal
                    'this block of parameters matches the stock id, main information and size information with the variables in the sql statement
                    .Parameters.AddWithValue("@MealID", Main.txtMealBuilderID.Text)
                    .Parameters.AddWithValue("@IngredientID", ingredients(x))
                    .Parameters.AddWithValue("@MealName", Main.txtMealName.Text)
                    .Parameters.AddWithValue("@Modifier", modifiers(x))
                    .Parameters.AddWithValue("@MealType", Main.cbMealType.Text)
                    .Connection = conn
                    .ExecuteNonQuery()
                End With

            Next

            Main.Reload()

        End If

        Meal_selection_populate()


    End Sub

    Dim MealID As String

    Sub Generate_Meal_ID()

        'SQL Select Statement Searching for the highest UserID assciated with that user type'
        MealID = "SELECT MAX(MealID) FROM tblMeal WHERE MealID LIKE '" & "ML" & "%%%%%' "
        'Data adapter defined'
        da = New OleDbDataAdapter(MealID, conn)
        'Data adapter told to fill the da   taset
        da.Fill(ds, "tblIngredients")
        'Defining the datatable
        dt = ds.Tables("tblIngredients")
        'Associating User ID to the first item on the first row of the datatabl whilst converting the data to a string'
        MealID = dt.Rows(0).Item(0).ToString

        'For defining the first student in the class;
        If MealID = "" Then
            'this is some "Fake Data" that is used to calculate the first User ID'
            MealID = "ML00000"
            Exit Sub
        End If

        'This strips off the User Type Identified'
        MealID = MealID.Substring(2, 5)
        'This removes the leading zero's from the stripped string'
        MealID = MealID.TrimStart("0"c)

        'This catches the User ID string if it is something like "S0000", after all the stripping and formating, would leave nothing, this code repairs it'
        If MealID = "" Then
            'Re-Defines the MealID to 0'
            MealID = 0
        End If

        'Converts the split string into an integer'
        MealID = CType(MealID, Integer) + 1

        'This block is responsible for adding the appropriate value onto the end of the user ID'
        'This converts the User ID to a string'
        MealID = CType(MealID, String)
        'Length check's used to identify the different tiers of the user id'
        If MealID.Length = 1 Then
            'This combines the User ID's componants into a full id'
            MealID = "ML" & "0000" & MealID
            'Length check's used to identify the different tiers of the user id'
        ElseIf MealID.Length = 2 Then
            'This combines the User ID's componants into a full id'
            MealID = "ML" & "000" & MealID
            'Length check's used to identify the different tiers of the user id'
        ElseIf MealID.Length = 3 Then
            'This combines the User ID's componants into a full id'
            MealID = "ML" & "00" & MealID
        ElseIf MealID.Length = 4 Then
            MealID = "ML" & "0" & MealID
        Else
            'This combines the User ID's componants into a full id'
            MealID = "ML" & MealID
        End If

        ds.Clear()

    End Sub

    Sub Meal_selection_populate()

        Dim selda As OleDbDataAdapter
        Dim seldt As DataTable
        Dim selds As DataSet = New DataSet


        Dim meal_selection_list As String = "SELECT * FROM tblMeal"

        'this defines uses the sql query and connection string
        selda = New OleDbDataAdapter(meal_selection_list, conn)
        'this datadapter is filled with the dataset declared earlier
        selda.Fill(selds, "Meal")
        'this datatable is filled with data pulled from the dataset
        seldt = selds.Tables("Meal")

        Dim repetitions(seldt.Rows.Count - 1) As String
        Dim selection(seldt.Rows.Count - 1) As String

        For i As Integer = 0 To seldt.Rows.Count - 1
            selection(i) = seldt.Rows(i).Item(0)
        Next

        repetitions(0) = seldt.Rows(0).Item(0)

        'Unique Record Algortithm

        Dim rep As Boolean = False
        Dim currentID As String = ""
        Dim counter As Integer = 0

        For Each row In seldt.Rows
            rep = False
            For j As Integer = 0 To repetitions.Length - 1
                If counter = 0 Then
                Else
                    If seldt.Rows(counter).Item(0).ToString = repetitions(j) Then
                        rep = True
                        row.delete()
                        Exit For
                    End If
                End If
            Next
            If rep = False Then
                repetitions(counter) = row.item(0)
            End If

            counter += 1
        Next

        selds.AcceptChanges()

        With Main.dgvMealSelection
            .AutoGenerateColumns = True
            .DataSource = seldt
        End With



        For i As Integer = 0 To 4
            If i = 2 Then
            Else
                Main.dgvMealSelection.Columns(i).Visible = False
            End If
        Next

        Main.dgvMealBuilder.AutoResizeRows()
        Main.dgvMealBuilder.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill


    End Sub

    Sub select_meal(i)

        Main.txtMealBuilderID.Text = Main.dgvMealSelection.Item(0, i).Value()
        Main.txtMealName.Text = Main.dgvMealSelection.Item(2, i).Value()
        Main.cbMealType.Text = Main.dgvMealSelection.Item(4, i).Value()

        Dim selectedmeal As String = Main.dgvMealSelection.Item(0, i).Value()

        Dim selda As OleDbDataAdapter
        Dim seldt As DataTable
        Dim selds As DataSet = New DataSet


        Dim meal_selection_list As String = "SELECT * FROM tblMeal"

        'this defines uses the sql query and connection string
        selda = New OleDbDataAdapter(meal_selection_list, conn)
        'this datadapter is filled with the dataset declared earlier
        selda.Fill(selds, "Meal")
        'this datatable is filled with data pulled from the dataset
        seldt = selds.Tables("Meal")

        Dim selmealing(seldt.Rows.Count) As String
        Dim selmealmod(seldt.Rows.Count) As String

        Dim counter As Integer = 0

        For Each row In seldt.Rows

            If row.item(0) = selectedmeal Then
                selmealing(counter) = row.item(1)
                selmealmod(counter) = row.item(3)
            End If
            counter += 1
        Next

        seldt.Clear()
        selds.Clear()

        Dim meal_ingredient_list As String = "SELECT * FROM tblIngredients"

        'this defines uses the sql query and connection string
        selda = New OleDbDataAdapter(meal_ingredient_list, conn)
        'this datadapter is filled with the dataset declared earlier
        selda.Fill(selds, "ing")
        'this datatable is filled with data pulled from the dataset
        seldt = selds.Tables("ing")

        empty_meal_table()

        Dim c As Integer = 0

        For x As Integer = 0 To selmealing.Length - 1

            If selmealing(x) = Nothing Then
                Continue For
            End If

            For Each row In seldt.Rows

                If row.item(0) = selmealing(x) Then

                    Dim mealingdr As DataRow = mealdt.NewRow

                    mealingdr("IngredientID") = selmealing(x)
                    mealingdr("Ingredient Name") = row.item(1)
                    mealingdr("Quantity") = CDbl(row.item(2)) * (CDbl(selmealmod(x)) * 0.01)
                    mealingdr("Calories") = CDbl(row.item(3)) * (CDbl(selmealmod(x)) * 0.01)
                    mealingdr("Protein") = CDbl(row.item(4)) * (CDbl(selmealmod(x)) * 0.01)
                    mealingdr("Carbohydrates") = CDbl(row.item(5)) * (CDbl(selmealmod(x)) * 0.01)
                    mealingdr("Fats") = CDbl(row.item(6)) * (CDbl(selmealmod(x)) * 0.01)

                    mealdt.Rows.Add(mealingdr)

                End If

                c += 1

            Next
        Next

        With Main.dgvMealBuilder
            .AutoGenerateColumns = True
            .DataSource = mealdt
        End With


        Main.dgvMealBuilder.AutoResizeRows()
        Main.dgvMealBuilder.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Main.dgvMealBuilder.Columns(0).Visible = False

        macro_counter()


    End Sub

    Sub built_ingredient_selection(i)

        Main.txtMealIngID.Text = Main.dgvMealBuilder.Item(0, i).Value
        Main.txtMealIngName.Text = Main.dgvMealBuilder.Item(1, i).Value
        Main.txtMealIngGrams.Text = Main.dgvMealBuilder.Item(2, i).Value
        Main.txtMealIngCalories.Text = Main.dgvMealBuilder.Item(3, i).Value
        Main.txtMealIngProtein.Text = Main.dgvMealBuilder.Item(4, i).Value
        Main.txtMealIngCarbohydrates.Text = Main.dgvMealBuilder.Item(5, i).Value
        Main.txtMealIngFats.Text = Main.dgvMealBuilder.Item(6, i).Value


    End Sub

End Module
