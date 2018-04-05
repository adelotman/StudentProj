Public Class FrmItemBalances
    Dim SqlWhere, SqlStr As String
    Dim TxtDscrp As String = ""
    Dim Str As String
    Dim TypeStuday As String = ""
    Dim Tb As New DataTable
    Private Sub FrmItemBalances_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Me.KeyPreview = True
        ConvertTabKeyToEnterKey(e)
        If e.KeyCode = Keys.F2 Then BtnRefrech_Click(sender, e)
        If e.KeyCode = Keys.F5 Then BtnFind_Click(sender, e)
        If e.KeyCode = Keys.F3 Then BtnPrint_Click(sender, e)

        If e.KeyCode = Keys.Escape Or e.KeyCode = Keys.F12 Then BtnExit_Click(sender, e)
    End Sub

    Private Sub FrmItemBalances_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        BtnRefrech_Click(sender, e)
    End Sub

    Private Sub BtnRefrech_Click(sender As System.Object, e As System.EventArgs) Handles BtnRefrech.Click
        'تحميل بيانات المخزن
        TxtTotals.Text = 0
        TxtTotalTxt.Text = Nothing
        TxtTotals.Text = FormatNumber(TxtTotals.Text, 3, , , TriState.True)

        CmbStocks.Text = Nothing
        Sql = "Select StockID,StockName from TbStock WHERE ([IsActive] ='True')"
        FILLCOMBOBOXITEMS2(Sql, CmbStocks)
        FILLCOMBOBOXITEMS2(Sql, CmbStocksCnv)
        If CmbStocks.Items.Count > 0 Then CmbStocks.SelectedIndex = 1 : CmbStocksCnv.SelectedIndex = 1

        CmbTypePrice.SelectedIndex = 1

        ChkFindMoney.Checked = False
        CmbOpratorFind.Enabled = ChkFindMoney.Checked
        TxtMonyFind.Enabled = ChkFindMoney.Checked
        ' If ChkFindMoney.Checked = False Then
        TxtMonyFind.Text = 0

        'End If
        CmbOpratorFind.SelectedIndex = 0
        TxtMonyFind.Text = FormatNumber(TxtMonyFind.Text, 3, , , TriState.True)
        TxtMonyFind.Focus()
        TxtMonyFind.SelectAll()

        'تحميل مراكز الكلفة
        CmbCenterCost.Text = Nothing
        Sql = "Select NumCost,CostName from TbCostCenter WHERE [IsActive] ='True' Order by NumCost"
        FILLCOMBOBOXITEMS2(Sql, CmbCenterCost)
        If CmbCenterCost.Items.Count > 1 Then CmbCenterCost.SelectedIndex = 1

        Sql = "Select ItemGrpID,ItemGrpName from TbItemGroup Where ItemGrpID>'0' and IsActive='True'"
        FILLCOMBOBOXITEMS2(Sql, ListGrpsItems)

        ListGrpsItems.Text = Nothing
        If ListGrpsItems.Items.Count > 0 Then ListGrpsItems.SelectedIndex = 0
        Sql = "Select ItemGrpName from TbItemGroup Where  IsActive='True'"
        '   AutocompleteSql(Sql, ListGrpsItems)

        NumYears.Minimum = 2014 ' MyClss.ExecuteScalar("Selcet Min(YearReg) From TbViewBalanceItems")
        NumYears.Maximum = Date.Now.Year
        NumYears.Value = Date.Now.Year

        NumYearsCnv.Minimum = 2014 ' MyClss.ExecuteScalar("Selcet Min(YearReg) From TbViewBalanceItems")
        NumYearsCnv.Maximum = Date.Now.Year
        NumYearsCnv.Value = Date.Now.Year

        DGVItems.DataSource = Nothing
        TxtNoItems.Text = 0
        ListGrpsItems.ClearSelected()
        ProgBar.Value = 0
        With Me.DGVItems
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
    End Sub

    Private Sub BtnExit_Click(sender As System.Object, e As System.EventArgs) Handles BtnExit.Click
        Me.Close()
    End Sub

    Private Sub BtnPrint_Click(sender As System.Object, e As System.EventArgs) Handles BtnPrint.Click
        If DGVItems.Rows.Count = 0 Then Exit Sub
        strReport = "RepBalItems"
        SqlRep = Sql '"SELECT * FROM  TbViewInvoiceDetials where NoIdRec='" + TxtNoRecInv.ToString + "'"
        ''1 New - 2 Insert - 3  Edit  - 4 Del - 5 Print - 6 Find - 7 Login - 8 Logout
        'Dim TxtDescrp As String = TxtNoInv.Text + " " + CmbTypeInv.Text
        'MyClss.UserMovement(UserID, "5", TxtDescrp, Me.Text)
        FrmRepPreview.Show()
    End Sub

    Private Sub BtnFind_Click(sender As System.Object, e As System.EventArgs) Handles BtnFind.Click

        If CmbStocks.Text = Nothing Then
            MessageBox.Show("من فضلك قم باختيار اسم المخزن من مربع الادخال", "خطاء في الاختيار", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            CmbStocks.Focus()
            Exit Sub
        End If

        TxtTotals.Text = 0
        TxtTotalTxt.Text = Nothing

        Sql = "SELECT ItemId,ItemCode, ItemName, StockID,ItemGrpName , UnitId, UnitName, Q,PurchPrice1, RetailPrice, ItemGrpID,StockName , YearReg FROM  TbViewBalanceItems"
        Sql += " Where StockID ='" + CmbStocks.SelectedValue.ToString + "' And YearReg = '" + NumYears.Value.ToString + "'"
        Dim K As Int16 = 0
        '  For i As Int16 = 0 To Me.ListGrpsItems
        For Each Item As DataRowView In ListGrpsItems.SelectedItems
            ' TextBox1.Text &= Item.ToString() & " "

            'display the listbox value
            '  NoEmpl = ListGrpsItems.SelectedItems(i)
            K += 1
            If K = 1 Then
                Sql += " And ItemGrpID='" + Item.Item(0).ToString() + "' "
            Else
                Sql += " Or ItemGrpID='" + Item.Item(0).ToString + "' "
            End If
        Next
        'If K = 0 Then
        '    MessageBox.Show("الرجاء اختيار موظف واحد على الاقل ليتم طباعته", "خطاء في الاختيار", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        '    Exit Sub
        'End If
        If ChkFindMoney.Checked = True Then

            If TxtMonyFind.Text = Nothing Or Information.IsNumeric(TxtMonyFind.Text) = False Then
                MessageBox.Show("من فضلك .. قم بادخال قيمة رقمية في مربع الادخال", "خطاء في البحث", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TxtMonyFind.Focus()
                Exit Sub
            End If
            Sql += " And Q " + CmbOpratorFind.Text.Trim + "'" + TxtMonyFind.Text.ToString + "'"
            SqlRep += " And Q " + CmbOpratorFind.Text.Trim + "'" + TxtMonyFind.Text.ToString + "'"
        End If
        '====================================================================================================================
        Sql += " Order by ItemCode,ItemGrpID,ItemId"

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

            .Columns("ItemId").Visible = False
            .Columns("UnitId").Visible = False
            .Columns("ItemGrpID").Visible = False

            'عرض الاعمدة السعر حسب الاسعار
            Select Case CmbTypePrice.SelectedIndex
                Case Is = 0
                    .Columns("PurchPrice1").Visible = False
                    .Columns("RetailPrice").Visible = False
                Case Is = 1
                    .Columns("PurchPrice1").Visible = True
                    .Columns("RetailPrice").Visible = False
                Case Is = 2
                    .Columns("PurchPrice1").Visible = False
                    .Columns("RetailPrice").Visible = True
                Case Else
                    .Columns("PurchPrice1").Visible = False
                    .Columns("RetailPrice").Visible = False
            End Select

            Dim Sum As Double = 0
            Select Case CmbTypePrice.SelectedIndex
                Case Is = 0
                    .Columns("PurchPrice1").Visible = False
                    .Columns("RetailPrice").Visible = False
                    TxtTotals.Text = 0
                    TxtTotalTxt.Text = Nothing

                Case Is = 1
                    For i As Integer = 0 To DGVItems.RowCount - 1
                        Sum += DGVItems.Rows(i).Cells("PurchPrice1").Value * DGVItems.Rows(i).Cells("Q").Value
                    Next
                    TxtTotals.Text = Sum
                    TxtTotalTxt.Text = HANY(Sum, "Libya")

                Case Is = 2
                    For i As Integer = 0 To DGVItems.RowCount - 1
                        Sum += DGVItems.Rows(i).Cells("RetailPrice").Value * DGVItems.Rows(i).Cells("Q").Value
                    Next
                    TxtTotals.Text = Sum
                    TxtTotalTxt.Text = HANY(Sum, "Libya")

                Case Else
                    TxtTotals.Text = 0
                    TxtTotalTxt.Text = Nothing

            End Select
            TxtTotals.Text = FormatNumber(TxtTotals.Text, 3, , , TriState.True)


            .Columns("StockID").Visible = False

            .Columns("ItemCode").HeaderText = "رقم الصنف"
            .Columns("ItemName").HeaderText = "اسم الصنف"
            .Columns("StockName").HeaderText = "المخزن"
            .Columns("UnitName").HeaderText = "الوحدة"
            .Columns("Q").HeaderText = "الكمية"
            .Columns("PurchPrice1").HeaderText = "سعر الشراء"
            .Columns("RetailPrice").HeaderText = "سعر البيع"
            .Columns("ItemGrpName").HeaderText = "المجموعة"
            .Columns("YearReg").HeaderText = "سنة الرصيد"

            .Columns("PurchPrice1").DefaultCellStyle.Format = "N3"
            .Columns("RetailPrice").DefaultCellStyle.Format = "N3"
            .Columns("Q").DefaultCellStyle.Format = "N3"

            .Columns("ItemCode").Width = 100
            .Columns("ItemName").Width = 170

        End With
        SetRowNumber(DGVItems)
    End Sub

    Private Sub BtnCal_Click(sender As System.Object, e As System.EventArgs) Handles BtnCal.Click
        System.Diagnostics.Process.Start("calc")
    End Sub

    Private Sub ChkFindMoney_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkFindMoney.CheckedChanged
        CmbOpratorFind.Enabled = ChkFindMoney.Checked
        TxtMonyFind.Enabled = ChkFindMoney.Checked
        ' If ChkFindMoney.Checked = False Then
        TxtMonyFind.Text = 0

        'End If
        CmbOpratorFind.SelectedIndex = 0
        TxtMonyFind.Text = FormatNumber(TxtMonyFind.Text, 3, , , TriState.True)
        TxtMonyFind.Focus()
        TxtMonyFind.SelectAll()
    End Sub

    Private Sub CmbOpratorFind_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbOpratorFind.SelectedIndexChanged
        TxtMonyFind.Focus()
        TxtMonyFind.SelectAll()

    End Sub

    Private Sub DisplayPay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DisplayPay.Click
        FrmInvTaswiaItems.Show()
    End Sub

    Private Sub BtnCnv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCnv.Click
        ProgBar.Value = 0
        If DGVItems.Rows.Count = 0 Then Exit Sub



        If CmbStocksCnv.Text = Nothing Then
            MessageBox.Show("اسف .. من فضلك اختر اسم المخزن من مربع الادخال", "خطاء في الادخال", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            CmbStocksCnv.Focus()
            Exit Sub
        End If
        If NumYearsCnv.Value < Date.Now.Year Then
            MessageBox.Show("اسف ..لا يمكن تحويل الرصيد الى سنة سابفة !!!", "خطاء في الادخال", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            NumYearsCnv.Focus()
            Exit Sub
        End If

        Dim resault = MessageBox.Show("هل انت متأكد من عملية اضافة رصيد أول المدة الى هذه السنة التي تم اختيارها ؟", "توحيل رصيد اول المدة ", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
        If resault = vbNo Then Exit Sub

        IDInvoiceType = 7
        NoTypePayment = 3
        Sql = "SELECT Max(bInOut) From TbBillType Where IDInvoiceType  ='" + IDInvoiceType.ToString + "'"
        TypeAdd = MyClss.ExecuteScalar(Sql)

        Dim TxtNoInv, TxtNoRecInv As Integer
        Dim DateRegInv As Date = DateSerial(NumYearsCnv.Value, 1, 1)

        'رقم الفاتورة حسب نوع حركة الفاتورة - مشتريات أو استلامات او مرتجع مشتريات 
        Sql = "SELECT Max(IDRecInv) From TbMainInvoice  WHERE ([IDInvoiceType] ='" + IDInvoiceType.ToString + "' And Year(DateRegInv)='" + NumYearsCnv.Value.ToString + "')"
        TxtNoInv = MyClss.ExecuteScalar(Sql) + 1

        Sql = "SELECT Max(NoIdRec) From TbMainInvoice"
        TxtNoRecInv = MyClss.ExecuteScalar(Sql) + 1


        ''النحقق من عدم وجود فواتير اخرى قبل فاتورة رصيد أول المدة
        'Sql = "SELECT Count(*) From TbMainInvoice WHERE ([IDInvoiceType] <>'" + CmbTypeInv.SelectedValue.ToString + "' And DateRegInv <'" + Format(DateRegInv.Value, "yyyy-MM-dd").ToString + "'And Year(DateRegInv)='" + DateRegInv.Value.Year.ToString + "' )"
        'Dim n = MyClss.ExecuteScalar(Sql)
        'If n > 0 Then
        '    MessageBox.Show("اسف .. هناك فواتير محفوظة قبل فواتير رصيد أول المدة الرجاء التحقق من الفواتير المدخلة قبل عملية الحفظ", "خطاء في الادخال", MessageBoxButtons.OK, MessageBoxIcon.Stop)

        '    Exit Sub
        'End If

        Sql = " INSERT INTO TbMainInvoice(IDRecInv,TypeMove, NoTypePayment, SuppID,  DateRegInv"
        Sql += ", StockID,IDInvoiceType, Notes)"
        Sql += " VALUES ('" + TxtNoInv.ToString + "','" + IDInvoiceType.ToString + "','" + NoTypePayment.ToString + "','0','" + Format(DateRegInv, "yyyy-MM-dd").ToString + "','"
        Sql += CmbStocks.SelectedValue.ToString + "','" + IDInvoiceType.ToString + "',N'')"
        Try
            MyClss.RunsInsertDeleteUpdateQry(Sql)
            If TestRunCode = True Then


                '     1 New - 2 Insert - 3  Edit  - 4 Del - 5 Print - 6 Find - 7 Login - 8 Logout
                MyClss.UserMovement(UserID, "2", "رصيد أول المدة", Me.Text)

                Sql = "SELECT Max(NoIdRec) From TbMainInvoice"
                TxtNoRecInv = MyClss.ExecuteScalar(Sql)

                '---------------------------------------------------------------------
                Dim N As Int32 = DGVItems.Rows.Count
                ProgBar.Visible = True
                ProgBar.Minimum = 0
                ProgBar.Maximum = 100
                Dim k1 = 0
                Dim kcount As Int16 = Math.Truncate(N / ProgBar.Maximum)
                Dim m As Int16 = 0
                '------------------------------------------------------------------------------

                For i = 0 To DGVItems.Rows.Count - 1
                    Dim Price As Double = 0

                    With DGVItems.Rows(i)

                        Dim TbPriceCost As New DataTable
                        Sql = "SELECT        SUM(Price) AS Price, COUNT(*) AS Count1  FROM TbViewInvoiceDetials "
                        Sql += " WHERE   (UnitId ='" + .Cells("UnitId").Value.ToString + "') AND (ItemId ='" + .Cells("ItemId").Value.ToString + "') AND year(datereginv)='" + NumYears.Value.ToString + "' And IDInvoiceType='13'"
                        TbPriceCost = MyClss.GetRecords(Sql)
                        Dim AvgPrice, SumPrice, CountPrice As Single
                        SumPrice = 0
                        If Information.IsDBNull(TbPriceCost.Rows(0).Item("Price")) <> True Then
                            SumPrice = TbPriceCost.Rows(0).Item("Price")
                            CountPrice = TbPriceCost.Rows(0).Item("Count1")
                            AvgPrice = SumPrice / CountPrice
                            Price = AvgPrice
                        End If

                        'التحقق من وجود رصيد الصنف مسبقا خلال هذه السنة
                        Sql = "SELECT Max(NoMoveItem) From TbMovementItems"
                        Sql += " WHERE ([IDInvoiceType] ='" + IDInvoiceType.ToString + "'And Year(DateRegInv)='" + NumYearsCnv.Value.ToString + "' And ItemId='" + .Cells("ItemId").Value.ToString + "')"
                        Dim NoMoveItem = MyClss.ExecuteScalar(Sql)
                        If NoMoveItem = 0 Then
                            'Sql = "Select Count(*) From TbMovementItems Where  ItemCode ='" + CmbItems.SelectedValue.ToString + "' And Year(DateRegInv)='" + DateRegInv.Value.Year.ToString + "' And  DateRegInv <'" + Format(DateRegInv.Value, "yyyy-MM-dd").ToString + "'"
                            'Dim n = MyClss.ExecuteScalar(Sql)
                            'If n > 0 Then
                            '    If MessageBox.Show("تحذير .. هذا الصنف تم ادخاله مسبقاً قبل رصيد أول المدة .. هل ترغب في الاستمرار في الحفظ ؟", "تحذير ادخال الرصيد", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then Exit Sub
                            'End If
                            Sql = "INSERT INTO TbMovementItems ( NoIdRec, StockID, ItemCode, UnitId, QCredits, Price,TypeAdd,DateRegInv,IDInvoiceType,ItemId)"
                            Sql += " VALUES  ('" + TxtNoRecInv.ToString + "','" + CmbStocks.SelectedValue.ToString + "','" + .Cells("ItemCode").Value.ToString + "','" + .Cells("UnitId").Value.ToString + "','" + .Cells("Q").Value.ToString + "','" + Price.ToString + "','" + TypeAdd.ToString + "','" + Format(DateRegInv, "yyyy-MM-dd").ToString + "','" + IDInvoiceType.ToString + "','" + .Cells("ItemId").Value.ToString + "')"
                        Else
                            Sql = "UPDATE    TbMovementItems SET   StockID ='" + CmbStocks.SelectedValue.ToString + "', ItemCode ='" + .Cells("ItemCode").Value.ToString + "', UnitId ='" + .Cells("UnitId").Value.ToString + "', QCredits ='" + .Cells("Q").Value.ToString + "', Price ='" + Price.ToString + "',TypeAdd='" + TypeAdd.ToString + "'"
                            Sql += ", DateRegInv ='" + Format(DateRegInv, "yyyy-MM-dd").ToString + "',IDInvoiceType='" + IDInvoiceType.ToString + "'"
                            Sql += " Where NoMoveItem='" + NoMoveItem.ToString + "'"
                        End If
                        MyClss.RunsInsertDeleteUpdateQry(Sql)

                    End With

                    '=============================================================================
                    If DGVItems.Rows.Count >= 100 Then
                        k1 += 1
                        If k1 >= kcount Then
                            m += 1
                            If m >= 100 Then
                                ProgBar.Value = 100
                            Else
                                ProgBar.Value = m
                            End If
                            k1 = 0
                        End If
                        '=============================================================================
                    Else
                        ProgBar.Value = 100
                    End If
                Next
                If TestRunCode = True Then
                    MessageBox.Show("تم تحويل بيانات رصيد اول المدة بنجاح", "حفظ", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As SqlClient.SqlException
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub BtnFirstBal_Click(sender As System.Object, e As System.EventArgs) Handles BtnFirstBal.Click
        IDInvoiceType = 7
        FrmInv.Show()
    End Sub

    Private Sub DGVItems_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVItems.CellContentClick

    End Sub

    Private Sub DGVItems_Resize(sender As Object, e As System.EventArgs) Handles DGVItems.Resize
        SetRowNumber(DGVItems)
    End Sub

    Private Sub DGVItems_Sorted(sender As Object, e As System.EventArgs) Handles DGVItems.Sorted
        SetRowNumber(DGVItems)
    End Sub

    Private Sub BtnFindInvf_Click(sender As System.Object, e As System.EventArgs) Handles BtnFindInvf.Click
        Dim Frm As New FrmFindInvoices
        Frm.Show()
    End Sub

    Private Sub CmbTypePrice_Click(sender As System.Object, e As System.EventArgs) Handles CmbTypePrice.Click

    End Sub

    Private Sub CmbStocks_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles CmbStocks.SelectedIndexChanged

    End Sub
End Class