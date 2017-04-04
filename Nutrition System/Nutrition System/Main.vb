Imports System.ComponentModel

Public Class Main

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        connect()

        populate_ingredients()
        populate_ingredients_Meal()

        meal_table_creation()
        meal_plan_table_creation()
        meal_plan_group_table_creation()

        gbMealBuilder.Enabled = False

        btnAddMealPlan.Enabled = False

        dgvIngredients.Sort(dgvIngredients.Columns(1), ListSortDirection.Ascending)

        dgvMealIngredients.Sort(dgvMealIngredients.Columns(1), ListSortDirection.Ascending)

        txtIngQuantity.Text = 100

        Meal_selection_populate()
        Meal_Plan_selection_populate()
        sslabelbottom.Text = Now.Date

        Existing_Plan_selection_populate()

        macro_counter_plan()
        macro_counter()




    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles btnSaveIng.Click

        Save_Ingredient()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnClearIng.Click

        clear_ingredient_form()

    End Sub

    Private Sub dgvIngredients_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvIngredients.CellClick

        Dim i As Integer
        Dim a As Integer

        'this sets i to the be the current row number and a to be the current cell selected number
        i = dgvIngredients.CurrentRow.Index
        a = dgvIngredients.CurrentCell.ColumnIndex

        ingredient_selection(i)

    End Sub

    Private Sub btnDeleteIng_Click(sender As Object, e As EventArgs) Handles btnDeleteIng.Click

        delete_ingredient_data()

    End Sub

    Private Sub cbQuantityModifier_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbQuantityModifier.SelectedIndexChanged

        Ingredient_modifier()

    End Sub

    Private Sub dgvMealIngredients_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvMealIngredients.CellClick

        Dim i As Integer

        'this sets i to the be the current row number and a to be the current cell selected number
        i = dgvMealIngredients.CurrentRow.Index

        Build_Meal(i)

    End Sub

    Private Sub txtMealGrams_TextChanged(sender As Object, e As EventArgs) Handles txtMealIngGrams.TextChanged

        Gram_Modify()

    End Sub

    Private Sub btnMealAdd_Click(sender As Object, e As EventArgs) Handles btnMealAdd.Click

        Add_ingredient_Meal()

    End Sub

    Private Sub btnMealRemove_Click(sender As Object, e As EventArgs) Handles btnMealRemove.Click

        Dim i As Integer

        Try
            i = dgvMealBuilder.CurrentRow.Index
        Catch ex As Exception
            MsgBox("Please Select Item to be removed")
            Exit Sub
        End Try

        remove_ingredient_meal(i)

    End Sub

    Private Sub NsButton4_Click(sender As Object, e As EventArgs) Handles NsButton4.Click

        empty_meal_table()

    End Sub

    Private Sub NsTextBox6_TextChanged(sender As Object, e As EventArgs) Handles txtMealName.TextChanged

        meal_main_info()

    End Sub

    Private Sub NsButton2_Click(sender As Object, e As EventArgs) Handles btnSaveMeal.Click

        save_built_meal()

    End Sub

    Private Sub btnClearMeal_Click(sender As Object, e As EventArgs) Handles btnClearMeal.Click

        Clear_meal()

    End Sub

    Private Sub NsTabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles NsTabControl1.SelectedIndexChanged

        Meal_selection_populate()

    End Sub

    Private Sub dgvMealSelection_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvMealSelection.CellClick

        Dim i As Integer

        'this sets i to the be the current row number and a to be the current cell selected number
        i = dgvMealSelection.CurrentRow.Index

        select_meal(i)

    End Sub

    Private Sub dgvMealBuilder_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvMealBuilder.CellClick

        Dim i As Integer

        'this sets i to the be the current row number and a to be the current cell selected number
        i = dgvMealBuilder.CurrentRow.Index

        built_ingredient_selection(i)

    End Sub

    Private Sub btnDeleteMeal_Click(sender As Object, e As EventArgs) Handles btnDeleteMeal.Click

        delete_built_meal()

    End Sub

    Private Sub dgvPlanMealSelection_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPlanMealSelection.CellClick

        Dim i As Integer

        'this sets i to the be the current row number and a to be the current cell selected number
        i = dgvPlanMealSelection.CurrentRow.Index

        populate_meal_plan_macros(i)

    End Sub

    Private Sub txtPlanName_TextChanged(sender As Object, e As EventArgs) Handles txtPlanName.TextChanged

        If txtPlanName.Text = "" Then
            btnAddMealPlan.Enabled = False
        Else
            btnAddMealPlan.Enabled = True
        End If

    End Sub

    Private Sub btnAddMealPlan_Click(sender As Object, e As EventArgs) Handles btnAddMealPlan.Click

        add_meal_to_planner()

    End Sub

    Private Sub btnNewMealPlan_Click(sender As Object, e As EventArgs) Handles btnNewMealPlan.Click

        empty_plan_meal_table()
        macro_counter_plan()


    End Sub

    Private Sub btnRemoveMealPlan_Click(sender As Object, e As EventArgs) Handles btnRemoveMealPlan.Click

        Dim i As Integer

        Try
            i = dgvMealPlanGroup.CurrentRow.Index
        Catch ex As Exception
            MsgBox("Please Select Item to be removed")
            Exit Sub
        End Try



        'this sets i to the be the current row number and a to be the current cell selected number


        remove_meal_plan(i)

    End Sub

    Sub Reload()

        Meal_Plan_selection_populate()
        populate_ingredients()
        populate_ingredients_Meal()
        Meal_selection_populate()
        Existing_Plan_selection_populate()
        Clear_meal()

    End Sub

    Private Sub txtSearchIngredient_TextChanged(sender As Object, e As EventArgs) Handles txtSearchIngredient.TextChanged

        Search_ingredients()

    End Sub

    Private Sub NsButton5_Click(sender As Object, e As EventArgs) Handles btnSavePlan.Click

        save_meal_plan()

    End Sub

    Private Sub dgvPlanSelection_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPlanSelection.CellClick

        Dim i As Integer

        'this sets i to the be the current row number and a to be the current cell selected number
        i = dgvPlanSelection.CurrentRow.Index

        'populate_plan_information(i)

    End Sub
End Class
