Imports System.Data.OleDb

Module Meal_Planner


    Dim da As OleDbDataAdapter
    Dim dt As DataTable
    Dim ds As DataSet = New DataSet

    Sub Meal_Plan_selection_populate()

        Dim mealplanselda As OleDbDataAdapter
        Dim mealplanseldt As DataTable
        Dim mealplanselds As DataSet = New DataSet

        Dim meal_selection_list As String = "SELECT * FROM tblMeal"

        'this defines uses the sql query and connection string
        mealplanselda = New OleDbDataAdapter(meal_selection_list, conn)
        'this datadapter is filled with the dataset declared earlier
        mealplanselda.Fill(mealplanselds, "Meal")
        'this datatable is filled with data pulled from the dataset
        mealplanseldt = mealplanselds.Tables("Meal")

        Dim repetitions(mealplanseldt.Rows.Count - 1) As String
        Dim selection(mealplanseldt.Rows.Count - 1) As String

        For i As Integer = 0 To mealplanseldt.Rows.Count - 1
            selection(i) = mealplanseldt.Rows(i).Item(0)
        Next

        repetitions(0) = mealplanseldt.Rows(0).Item(0)

        'Unique Record Algortithm

        Dim rep As Boolean = False
        Dim currentID As String = ""
        Dim counter As Integer = 0

        For Each row In mealplanseldt.Rows
            rep = False
            For j As Integer = 0 To repetitions.Length - 1
                If counter = 0 Then
                Else
                    If mealplanseldt.Rows(counter).Item(0).ToString = repetitions(j) Then
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

        mealplanselds.AcceptChanges()

        With Main.dgvPlanMealSelection
            .AutoGenerateColumns = True
            .DataSource = mealplanseldt
        End With

        For i As Integer = 0 To 4
            If i = 2 Then
            Else
                Main.dgvPlanMealSelection.Columns(i).Visible = False
            End If
        Next

        Main.dgvPlanMealSelection.AutoResizeRows()
        Main.dgvPlanMealSelection.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    Sub populate_meal_plan_macros(i)

        Dim selectedmeal As String = Main.dgvPlanMealSelection.Item(0, i).Value()

        Main.txtPlanMealID.Text = Main.dgvPlanMealSelection.Item(0, i).Value()
        Main.txtPlanMealName.Text = Main.dgvPlanMealSelection.Item(2, i).Value()
        Main.txtPlanMealType.Text = Main.dgvPlanMealSelection.Item(4, i).Value()

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

        Dim mealgram As Double
        Dim mealcal As Double
        Dim mealprot As Double
        Dim mealcarb As Double
        Dim mealfats As Double

        For j As Integer = 0 To selmealing.Count - 1
            For Each row In seldt.Rows
                If row.item(0) = selmealing(j) Then

                    mealgram = mealgram + selmealmod(j)
                    mealcal = mealcal + (CDbl(row.item(3)) * (CDbl(selmealmod(j)) * 0.01))
                    mealprot = mealprot + (CDbl(row.item(4)) * (CDbl(selmealmod(j)) * 0.01))
                    mealcarb = mealcarb + (CDbl(row.item(5)) * (CDbl(selmealmod(j)) * 0.01))
                    mealfats = mealfats + (CDbl(row.item(6)) * (CDbl(selmealmod(j)) * 0.01))

                End If
            Next
        Next


        Main.txtPlanMealGrams.Text = mealgram
        Main.txtPlanMealCalories.Text = mealcal
        Main.txtPlanMealProtein.Text = mealprot
        Main.txtPlanMealCarbohydrates.Text = mealcarb
        Main.txtPlanMealFats.Text = mealfats

    End Sub

    Sub add_meal_to_planner()

        Dim planmealdr As DataRow = planmealdt.NewRow

        planmealdr("MealID") = Main.txtPlanMealID.Text
        planmealdr("Meal Name") = Main.txtPlanMealName.Text
        planmealdr("Quantity") = Main.txtPlanMealGrams.Text
        planmealdr("Calories") = Main.txtPlanMealCalories.Text
        planmealdr("Protein") = Main.txtPlanMealProtein.Text
        planmealdr("Carbohydrates") = Main.txtPlanMealCarbohydrates.Text
        planmealdr("Fats") = Main.txtPlanMealFats.Text

        planmealdt.Rows.Add(planmealdr)

        With Main.dgvMealPlanGroup
            .AutoGenerateColumns = True
            .DataSource = planmealdt
        End With


        Main.dgvMealPlanGroup.AutoResizeRows()
        Main.dgvMealPlanGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Main.dgvMealPlanGroup.Columns(0).Visible = False

        macro_counter_plan()

    End Sub

    Sub macro_counter_plan()

        Dim Num_ingredients As Integer = Main.dgvMealPlanGroup.Rows.Count

        Dim totGrams As Double
        Dim totCals As Double
        Dim totProt As Double
        Dim totCarbs As Double
        Dim totfats As Double

        For i As Integer = 0 To Num_ingredients - 1
            totGrams = totGrams + Main.dgvMealPlanGroup.Item(2, i).Value()
            totCals = totCals + Main.dgvMealPlanGroup.Item(3, i).Value()
            totProt = totProt + Main.dgvMealPlanGroup.Item(4, i).Value()
            totCarbs = totCarbs + Main.dgvMealPlanGroup.Item(5, i).Value()
            totfats = totfats + Main.dgvMealPlanGroup.Item(6, i).Value()
        Next

        Main.txtPlanTotalGrams.Text = totGrams
        Main.txtPlanTotalCal.Text = totCals
        Main.txtPlanTotalProt.Text = totProt
        Main.txtPlanTotalCarbs.Text = totCarbs
        Main.txtPlanTotalFats.Text = totfats



    End Sub

    Dim planmealdt As New DataTable

    Sub meal_plan_table_creation()

        planmealdt.Columns.Add("MealID", Type.GetType("System.String"))
        planmealdt.Columns.Add("Meal Name", Type.GetType("System.String"))
        planmealdt.Columns.Add("Quantity", Type.GetType("System.String"))
        planmealdt.Columns.Add("Calories", Type.GetType("System.String"))
        planmealdt.Columns.Add("Protein", Type.GetType("System.String"))
        planmealdt.Columns.Add("Carbohydrates", Type.GetType("System.String"))
        planmealdt.Columns.Add("Fats", Type.GetType("System.String"))

    End Sub

    Sub empty_plan_meal_table()

        For x As Integer = planmealdt.Rows.Count - 1 To 0 Step -1
            planmealdt.Rows(x).Delete()

        Next

    End Sub

    Sub remove_meal_plan(i)


        Dim ingredient As String = Main.dgvMealPlanGroup.Item(0, i).Value()

        For x As Integer = planmealdt.Rows.Count - 1 To 0 Step -1
            If planmealdt.Rows(x).Item(0) = ingredient Then
                planmealdt.Rows(x).Delete()
            End If
        Next

        macro_counter()


    End Sub

    Sub save_meal_plan()

        If Main.txtPlanName.Text = "" Then
            MsgBox("No Meal Plan Entered")
            Exit Sub
        End If

        If planmealdt.Rows.Count = 0 Then
            MsgBox("No Meals selected")
            Exit Sub
        End If

        Generate_Plan_ID()

        Dim planname As String = Main.txtPlanName.Text

        Dim planmeals(planmealdt.Rows.Count - 1)

        Dim mealdist(planmealdt.Rows.Count - 1)

        Dim mealdistrep(planmealdt.Rows.Count - 1)

        Dim counter As Integer = 0

        For Each row In planmealdt.Rows
            planmeals(counter) = row.item(0)
            counter += 1
        Next

        counter = 0

        For Each item In planmeals.Distinct
            mealdist(counter) = item
            counter += 1
        Next

        counter = 0

        For Each distitem In mealdist
            For Each item In planmeals
                If item = distitem Then
                    mealdistrep(counter) += 1
                End If
            Next
            counter += 1
        Next

        For x As Integer = 0 To planmealdt.Rows.Count - 1

            If mealdist(x) = Nothing Then
                Continue For
            End If

            Dim InsertMeal As String = "INSERT INTO tblMealPlan ([PlanID],[MealID],[Repetitions],[PlanName]) VALUES (@PlanID,@MealID,@Repetitions,@PlanName)"
            Dim InsertMealCommand As New OleDbCommand
            With InsertMealCommand
                .CommandText = InsertMeal
                'this block of parameters matches the stock id, main information and size information with the variables in the sql statement
                .Parameters.AddWithValue("@PlanID", PlanID)
                .Parameters.AddWithValue("@MealID", mealdist(x))
                .Parameters.AddWithValue("@Repetitions", mealdistrep(x))
                .Parameters.AddWithValue("@PlanName", planname)
                .Connection = conn
                .ExecuteNonQuery()
            End With
        Next

        Main.Reload()

        MsgBox("Plan Saved")



    End Sub

    Dim PlanID As String

    Sub Generate_Plan_ID()

        'SQL Select Statement Searching for the highest UserID assciated with that user type'
        PlanID = "SELECT MAX(PlanID) FROM tblMealPlan WHERE PlanID LIKE '" & "PL" & "%%%%%' "
        'Data adapter defined'
        da = New OleDbDataAdapter(PlanID, conn)
        'Data adapter told to fill the da   taset
        da.Fill(ds, "tblPlan")
        'Defining the datatable
        dt = ds.Tables("tblPlan")
        'Associating User ID to the first item on the first row of the datatabl whilst converting the data to a string'
        PlanID = dt.Rows(0).Item(0).ToString

        'For defining the first student in the class;
        If PlanID = "" Then
            'this is some "Fake Data" that is used to calculate the first User ID'
            PlanID = "PL00000"
            Exit Sub
        End If

        'This strips off the User Type Identified'
        PlanID = PlanID.Substring(2, 5)
        'This removes the leading zero's from the stripped string'
        PlanID = PlanID.TrimStart("0"c)

        'This catches the User ID string if it is something like "S0000", after all the stripping and formating, would leave nothing, this code repairs it'
        If PlanID = "" Then
            'Re-Defines the PlanID to 0'
            PlanID = 0
        End If

        'Converts the split string into an integer'
        PlanID = CType(PlanID, Integer) + 1

        'This block is responsible for adding the appropriate value onto the end of the user ID'
        'This converts the User ID to a string'
        PlanID = CType(PlanID, String)
        'Length check's used to identify the different tiers of the user id'
        If PlanID.Length = 1 Then
            'This combines the User ID's componants into a full id'
            PlanID = "PL" & "0000" & PlanID
            'Length check's used to identify the different tiers of the user id'
        ElseIf PlanID.Length = 2 Then
            'This combines the User ID's componants into a full id'
            PlanID = "PL" & "000" & PlanID
            'Length check's used to identify the different tiers of the user id'
        ElseIf PlanID.Length = 3 Then
            'This combines the User ID's componants into a full id'
            PlanID = "PL" & "00" & PlanID
        ElseIf PlanID.Length = 4 Then
            PlanID = "PL" & "0" & PlanID
        Else
            'This combines the User ID's componants into a full id'
            PlanID = "PL" & PlanID
        End If

        ds.Clear()

    End Sub

    Sub Existing_Plan_selection_populate()

        Dim selda As OleDbDataAdapter
        Dim seldt As DataTable
        Dim selds As DataSet = New DataSet


        Dim meal_Plan_selection_list As String = "SELECT * FROM tblMealPlan"

        'this defines uses the sql query and connection string
        selda = New OleDbDataAdapter(meal_Plan_selection_list, conn)
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

        With Main.dgvPlanSelection
            .AutoGenerateColumns = True
            .DataSource = seldt
        End With

        For i As Integer = 0 To 3
            If i = 3 Then
            Else
                Main.dgvPlanSelection.Columns(i).Visible = False
            End If
        Next

        Main.dgvPlanSelection.AutoResizeRows()
        Main.dgvPlanSelection.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    End Sub

    'Sub populate_plan_information(i)

    '    Dim PlanID As String = Main.dgvPlanSelection.Item(0, i).Value()

    '    Dim selda As OleDbDataAdapter
    '    Dim seldt As DataTable
    '    Dim selds As DataSet = New DataSet


    '    Dim plan_selection_list As String = "SELECT * FROM tblMealPlan"

    '    'this defines uses the sql query and connection string
    '    selda = New OleDbDataAdapter(plan_selection_list, conn)
    '    'this datadapter is filled with the dataset declared earlier
    '    selda.Fill(selds, "plan")
    '    'this datatable is filled with data pulled from the dataset
    '    seldt = selds.Tables("plan")

    '    Dim selplan(seldt.Rows.Count) As String
    '    Dim selplanrep(seldt.Rows.Count) As String

    '    Dim counter As Integer = 0

    '    For Each row In seldt.Rows

    '        If row.item(0) = PlanID Then
    '            selplan(counter) = row.item(1)
    '            selplanrep(counter) = row.item(2)
    '        End If
    '        counter += 1
    '    Next

    '    seldt.Clear()
    '    selds.Clear()

    '    Dim meal_list As String = "SELECT * FROM tblMeal"

    '    'this defines uses the sql query and connection string
    '    selda = New OleDbDataAdapter(meal_list, conn)
    '    'this datadapter is filled with the dataset declared earlier
    '    selda.Fill(selds, "meal")
    '    'this datatable is filled with data pulled from the dataset
    '    seldt = selds.Tables("meal")

    '    Dim numbermeals As Integer = seldt.Rows.Count

    '    counter = 0

    '    Dim selmeal(numbermeals) As String
    '    Dim selmealmod(numbermeals) As String

    '    For Each item In selplan
    '        For Each row In seldt.Rows
    '            If row.item(0) = item Then
    '                selmeal(counter) = row.item(1)
    '                selmealmod(counter) = row.item(3)
    '                counter += 1
    '            End If
    '        Next
    '    Next


    '    seldt.Clear()
    '    selds.Clear()


    '    Dim meal_ingredient_list As String = "SELECT * FROM tblIngredients"

    '    'this defines uses the sql query and connection string
    '    selda = New OleDbDataAdapter(meal_ingredient_list, conn)
    '    'this datadapter is filled with the dataset declared earlier
    '    selda.Fill(selds, "ing")
    '    'this datatable is filled with data pulled from the dataset
    '    seldt = selds.Tables("ing")

    '    Dim mealgram As Double
    '    Dim mealcal As Double
    '    Dim mealprot As Double
    '    Dim mealcarb As Double
    '    Dim mealfats As Double

    '    For j As Integer = 0 To selmeal.Count - 1
    '        For Each row In seldt.Rows
    '            If row.item(0) = selmeal(j) Then

    '                mealgram = mealgram + selmealmod(j)
    '                mealcal = mealcal + (CDbl(row.item(3)) * (CDbl(selmealmod(j)) * 0.01))
    '                mealprot = mealprot + (CDbl(row.item(4)) * (CDbl(selmealmod(j)) * 0.01))
    '                mealcarb = mealcarb + (CDbl(row.item(5)) * (CDbl(selmealmod(j)) * 0.01))
    '                mealfats = mealfats + (CDbl(row.item(6)) * (CDbl(selmealmod(j)) * 0.01))

    '            End If
    '        Next
    '    Next


    '    Main.txtPlanMealGrams.Text = mealgram
    '    Main.txtPlanMealCalories.Text = mealcal
    '    Main.txtPlanMealProtein.Text = mealprot
    '    Main.txtPlanMealCarbohydrates.Text = mealcarb
    '    Main.txtPlanMealFats.Text = mealfats


    '    Dim planmealdr As DataRow = planmealdt.NewRow

    '    counter = 0

    '    For Each item In selplan

    '        If item = Nothing Then
    '            Continue For
    '        End If

    '        For j As Integer = 0 To selmeal.Count - 1

    '            For x As Integer = 0 To CInt(selplanrep(counter) - 1)

    '                mealgram = 0
    '                mealcal = 0
    '                mealprot = 0
    '                mealcarb = 0
    '                mealfats = 0

    '                For Each row In seldt.Rows

    '                    If row.item(0) = selmeal(j) Then

    '                        mealgram = mealgram + selmealmod(j)
    '                        mealcal = mealcal + (CDbl(row.item(3)) * (CDbl(selmealmod(j)) * 0.01))
    '                        mealprot = mealprot + (CDbl(row.item(4)) * (CDbl(selmealmod(j)) * 0.01))
    '                        mealcarb = mealcarb + (CDbl(row.item(5)) * (CDbl(selmealmod(j)) * 0.01))
    '                        mealfats = mealfats + (CDbl(row.item(6)) * (CDbl(selmealmod(j)) * 0.01))

    '                    End If

    '                Next

    '            Next

    '            counter += 1

    '        Next

    '        planmealdr("MealID") = Main.txtPlanMealID.Text
    '        planmealdr("Meal Name") = Main.txtPlanMealName.Text
    '        planmealdr("Quantity") = Main.txtPlanMealGrams.Text
    '        planmealdr("Calories") = Main.txtPlanMealCalories.Text
    '        planmealdr("Protein") = Main.txtPlanMealProtein.Text
    '        planmealdr("Carbohydrates") = Main.txtPlanMealCarbohydrates.Text
    '        planmealdr("Fats") = Main.txtPlanMealFats.Text

    '        planmealdt.Rows.Add(planmealdr)

    '        counter += 1

    '    Next

    '    With Main.dgvMealPlanGroup
    '        .AutoGenerateColumns = True
    '        .DataSource = planmealdt
    '    End With


    '    Main.dgvMealPlanGroup.AutoResizeRows()
    '    Main.dgvMealPlanGroup.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

    '    Main.dgvMealPlanGroup.Columns(0).Visible = False

    '    macro_counter_plan()


    'End Sub

    Dim planmealgroupdt As New DataTable

    Sub meal_plan_group_table_creation()

        planmealgroupdt.Columns.Add("MealID", Type.GetType("System.String"))
        planmealgroupdt.Columns.Add("Meal Name", Type.GetType("System.String"))
        planmealgroupdt.Columns.Add("Quantity", Type.GetType("System.String"))
        planmealgroupdt.Columns.Add("Calories", Type.GetType("System.String"))
        planmealgroupdt.Columns.Add("Protein", Type.GetType("System.String"))
        planmealgroupdt.Columns.Add("Carbohydrates", Type.GetType("System.String"))
        planmealgroupdt.Columns.Add("Fats", Type.GetType("System.String"))

    End Sub

    Sub empty_plan_meal_group_table()

        For x As Integer = planmealgroupdt.Rows.Count - 1 To 0 Step -1
            planmealgroupdt.Rows(x).Delete()

        Next

    End Sub

    

End Module
