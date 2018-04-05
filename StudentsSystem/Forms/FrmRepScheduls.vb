Public Class FrmRepScheduls
    Dim SqlWhere, SqlStr As String
    Dim TxtDscrp As String = ""
    Dim Str As String
    Dim TypeStuday As String = ""
    Dim Tb As New DataTable
    Private Sub BtnRefrech_Click(sender As System.Object, e As System.EventArgs) Handles BtnRefrech.Click

        ChkCurse.Checked = False
        CmbCourses.Enabled = ChkCurse.Checked

        ChkFindMoney.Checked = False
        NumRec.Value = 1
        CmbOpratorFind.SelectedIndex = 0


        CmbOpratorFind.Enabled = ChkFindMoney.Checked
        NumRec.Enabled = ChkFindMoney.Checked

        CmbYears.SelectedIndex = 0
        Sql = "SELECT        Group_ID, Group_Name FROM            tbl_Groups "
        FILLCOMBOBOXITEMS2(Sql, CmbGroups)
        If CmbGroups.Items.Count > 0 Then CmbGroups.SelectedIndex = 0

        DGVItems.DataSource = Nothing
        TxtNoItems.Text = 0
    End Sub

    Private Sub FrmRepScheduls_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Me.KeyPreview = True
        ConvertTabKeyToEnterKey(e)
        If e.KeyCode = Keys.F2 Then BtnRefrech_Click(sender, e)
        If e.KeyCode = Keys.F5 Then BtnFind_Click(sender, e)
        If e.KeyCode = Keys.F3 Then BtnPrint_Click(sender, e)

        If e.KeyCode = Keys.Escape Or e.KeyCode = Keys.F12 Then BtnExit_Click(sender, e)
    End Sub

    Private Sub FrmRepScheduls_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        BtnRefrech_Click(sender, e)
    End Sub

    Private Sub ChkFindMoney_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkFindMoney.CheckedChanged
        CmbOpratorFind.Enabled = ChkFindMoney.Checked
        NumRec.Enabled = ChkFindMoney.Checked
        ' If ChkFindMoney.Checked = False Then
        NumRec.Value = 1

        'End If
        CmbOpratorFind.SelectedIndex = 0
        NumRec.Focus()

    End Sub

    Private Sub CmbOpratorFind_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbOpratorFind.SelectedIndexChanged
        NumRec.Focus()
        '  NumRec.SelectAll()
    End Sub



    Private Sub BtnPrint_Click(sender As System.Object, e As System.EventArgs) Handles BtnPrint.Click
        If DGVItems.Rows.Count = 0 Then Exit Sub
        strReport = "RepBalItems"
        SqlRep = Sql '"SELECT * FROM  TbViewInvoiceDetials where NoIdRec='" + TxtNoRecInv.ToString + "'"
        ''1 New - 2 Insert - 3  Edit  - 4 Del - 5 Print - 6 Find - 7 Login - 8 Logout
        'Dim TxtDescrp As String = TxtNoInv.Text + " " + CmbTypeInv.Text
        'MyClss.UserMovement(UserID, "5", TxtDescrp, Me.Text)
        '  FrmRepPreview.Show()
    End Sub

    Sub DisplayDataGrid()
        ' DataViewName = MyDs.Tables(0).DefaultView

        'For i = 0 To DGViewEmp.ColumnCount - 1
        '    DGViewEmp.Columns(i).ReadOnly = True
        'Next

        'ItemId, ItemName, StockID, StockName, UnitId, UnitName, Q, ItemGrpID, ItemGrpName, YearReg

        With DGVItems
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
            '.Columns("Chk").HeaderText = "اختيار"
            '.Columns("Chk").Width = 90
            '.Columns("Chk").ReadOnly = False
            Sql = "SELECT  Week_Period,Period_ID, Real_ID, Real_Name, Acadmic_Year, Course_ID, Course_Name, Year_No,"
            ' Course_Units_Theo, Course_Units_Prac,Location_iD,Group_ID, Group_Name, Section_ID, Section_Name, Course_Part_ID, Part_Name  "

            .Columns("Period_ID").Visible = False
            .Columns("Real_ID").Visible = False
            .Columns("Course_Units_Theo").Visible = False
            .Columns("Course_Units_Prac").Visible = False
            .Columns("Section_ID").Visible = False
            .Columns("Group_ID").Visible = False
            .Columns("Course_Part_ID").Visible = False
            .Columns("Acadmic_Year").Visible = False
            .Columns("Year_No").Visible = False

            .Columns("Course_ID").HeaderText = "رقم المادة"
            .Columns("Week_Period").HeaderText = "رقم الحصة"
            .Columns("Real_Name").HeaderText = "الفصل الدراسي"
            .Columns("Course_Name").HeaderText = "اسم المادة"
            .Columns("Location_iD").HeaderText = "القاعة"
            .Columns("Group_Name").HeaderText = "المجموعة"
            .Columns("Section_Name").HeaderText = "طبيعة المادة"
            .Columns("Part_Name").HeaderText = "نوع المادة"
            '.Columns("ItemGrpName").HeaderText = "المجموعة"
            '.Columns("YearReg").HeaderText = "سنة الرصيد"

            '.Columns("PurchPrice1").DefaultCellStyle.Format = "N3"
            '.Columns("RetailPrice").DefaultCellStyle.Format = "N3"
            '.Columns("Q").DefaultCellStyle.Format = "N3"

            '.Columns("ItemCode").Width = 100
            .Columns("Course_Name").Width = 170
            TxtNoItems.Text = .RowCount
        End With
        SetRowNumber(DGVItems)
    End Sub

    Private Sub BtnFind_Click(sender As System.Object, e As System.EventArgs) Handles BtnFind.Click
        If CmbYears.Text = Nothing Then
            MessageBox.Show("من فضلك قم باختيار السنة الدراسية من مربع الادخال", "خطاء في الاختيار", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            CmbYears.Focus()
            Exit Sub
        End If

       
        Sql = "SELECT  Week_Period,Period_ID, Real_ID, Real_Name, Acadmic_Year, Course_ID, Course_Name, Year_No, Course_Units_Theo, Course_Units_Prac,Location_iD,Group_ID, Group_Name, Section_ID, Section_Name, Course_Part_ID, Part_Name FROM  TbViewPeriods "

        Sql += " Where Acadmic_Year =" + (CmbYears.SelectedIndex + 1).ToString ' + " And YearReg = '" + NumYears.Value.ToString + "'"

        SqlStr = "SELECT  Week_Period,Period_ID, Real_ID, Real_Name, Acadmic_Year, Course_ID, Course_Name, Year_No, Course_Units_Theo, Course_Units_Prac,Location_iD,Group_ID, Group_Name, Section_ID, Section_Name, Course_Part_ID, Part_Name FROM  TbViewPeriods "
        SqlWhere = " Where Acadmic_Year =" + (CmbYears.SelectedIndex + 1).ToString ' + " And YearReg = '" + NumYears.Value.ToString + "'"

        Dim K As Int16 = 0
        '  For i As Int16 = 0 To Me.ListGrpsItems
        For Each Item As DataRowView In ListRealItems.SelectedItems
            ' TextBox1.Text &= Item.ToString() & " "

            'display the listbox value
            '  NoEmpl = ListGrpsItems.SelectedItems(i)
            K += 1
            If K = 1 Then
                Sql += " And   Real_ID=" + Item.Item(0).ToString() + " "
                SqlWhere += " And   Real_ID=" + Item.Item(0).ToString() + " "
            Else
                Sql += " Or   Real_ID=" + Item.Item(0).ToString + " "
                SqlWhere += " Or   Real_ID=" + Item.Item(0).ToString + " "
            End If
        Next
        If K = 0 Then
            MessageBox.Show("الرجاء اختيار فصل دراسي واحد على الاقل ليتم عرضه", "خطاء في الاختيار", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Exit Sub
        End If

        If ChkCurse.Checked Then
            If CmbCourses.Text = Nothing Or CmbCourses.Items.Count = 0 Then
                MessageBox.Show("من فضلك قم باختيار المادة الدراسية من مربع الادخال", "خطاء في الاختيار", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                CmbCourses.Focus()
                Exit Sub
            End If
            Sql += " And Course_ID= " + CmbCourses.SelectedValue.ToString + " "
            SqlWhere += " And Course_ID= " + CmbCourses.SelectedValue.ToString + " "
        End If

        If ChkFindMoney.Checked = True Then

            If NumRec.Value = 0 Or Information.IsNumeric(NumRec.Value) = False Then
                MessageBox.Show("من فضلك .. قم بادخال قيمة رقمية في مربع الادخال", "خطاء في البحث", MessageBoxButtons.OK, MessageBoxIcon.Information)
                NumRec.Focus()
                Exit Sub
            End If
            ' Sql += " And Q " + CmbOpratorFind.Text.Trim + " " + NumRec.Value.ToString + " "
            '  SqlRep += " And Q " + CmbOpratorFind.Text.Trim + " " + NumRec.Value.ToString + " "

            ds.Clear()
            For I = 1 To NumRec.Value
                'Sql = "SELECT YearNow From TbEissal WHERE (NoCode)='" & TxtNoCode.ToString & "'  AND (YearNow)='" + I.ToString + "' ORDER BY YearNow"
                'Dim Ds2 As New DataSet
                'DAdpt1 = New System.Data.OleDb.OleDbDataAdapter(Sql, Cn)
                'Ds2.Clear()
                'DAdpt1.Fill(Ds2, "Tb")
                'DAdpt1.Dispose()

                'If Ds2.Tables("Tb").Rows.Count = 0 Then GoTo 10
                'Sql = "SELECT   Max(YearNow),Sum(Car),Sum(Efaa),Sum(Moamlat) FROM TbEissal WHERE (NoCode)='" & TxtNoCode.ToString & "'  AND (YearNow)='" + I.ToString + "'"
                'DAdpt2 = New System.Data.OleDb.OleDbDataAdapter(Sql, Cn)

                'DAdpt2.Fill(ds, "Tb3")
                'DAdpt2.Dispose()
                'DGVItems.DataSource = ds.Tables(2)
10:
            Next I




        End If
        '====================================================================================================================
        Sql += " Order by Period_ID,Week_Period "

        Tb = MyClss.GetRecords(Sql)
        DGVItems.DataSource = Nothing
        DGVItems.ColumnHeadersVisible = True
        Me.Cursor = Cursors.AppStarting

        DGVItems.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing
        DGVItems.RowHeadersVisible = False

        DGVItems.DataSource = Tb
        DGVItems.RowHeadersVisible = True

        Me.Cursor = Cursors.Default
        TxtNoItems.Text = DGVItems.RowCount

        DisplayDataGrid()
    End Sub

    Private Sub CmbYears_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles CmbYears.SelectedIndexChanged
        On Error Resume Next
        'Dim objSelectedValue As Object = CmbYears.SelectedValue
        'If TypeOf (objSelectedValue) Is DataRowView Or CmbYears.Items.Count = 0 Or CmbYears.SelectedValue.ToString = Nothing Then Exit Sub

        Sql = "SELECT Real_ID, Real_Name  FROM   tbl_Real Where Acadmic_Year=" + (CmbYears.SelectedIndex + 1).ToString
        FILLCOMBOBOXITEMS2(Sql, ListRealItems)
        ListRealItems.Text = Nothing
        If ListRealItems.Items.Count > 0 Then ListRealItems.SelectedIndex = 0

        Sql = "SELECT  Course_ID, Course_Name  FROM  tbl_Courses Where Year_No=" + (CmbYears.SelectedIndex + 1).ToString
        FILLCOMBOBOXITEMS2(Sql, CmbCourses)
        If CmbCourses.Items.Count > 0 Then CmbCourses.SelectedIndex = 0


    End Sub

    Private Sub BtnExit_Click(sender As System.Object, e As System.EventArgs) Handles BtnExit.Click
        Me.Close()
    End Sub

    Private Sub ChkCurse_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles ChkCurse.CheckedChanged
        CmbCourses.Enabled = ChkCurse.Checked
    End Sub
End Class
