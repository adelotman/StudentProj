Public Class FrmTypeCards
    Dim TestNewSave As Boolean, Sql As String
    Dim Tb As New DataTable
    Private Sub BtnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnNew.Click
        With Me.DGVItems
            .RowsDefaultCellStyle.BackColor = Color.Bisque
            .AlternatingRowsDefaultCellStyle.BackColor = Color.Beige
        End With
        TestNewSave = True

        'تحميل   الخدمات
        Sql = "Select srvid,ServecName from TbServices Where  IsActive=True Order By srvid"
        FILLCOMBOBOXITEMS2(Sql, CmbSrv)
        If CmbSrv.Items.Count > 0 Then CmbSrv.SelectedIndex = 0
        Sql = "Select ServecName from TbServices Where  IsActive=True"
        AutocompleteSql(Sql, CmbSrv)

        Sql = "Select CardName from TbTypeCards Where  IsActive=True"
        AutocompleteSql(Sql, TxtCardName)

        Sql = "Select * from TbViewCards"
        Tb.Clear()
        Tb = MyClss.GetRecords(Sql)
        DGVItems.DataSource = Tb
        With DGVItems
            .Columns("IDCard").Visible = False
            .Columns("srvid").Visible = False
            .Columns("srvid").HeaderText = "كود الخدمة"
            .Columns("srvid").Width = 100
            .Columns("ServecName").Width = 170
            .Columns("ServecName").HeaderText = "اسم الخدمة"
            .Columns("CardName").Width = 260
            .Columns("CardName").HeaderText = "اسم فئة الكرت"
            .Columns("IsActive").HeaderText = "التفعيل"
            .Columns("PriceCd").HeaderText = "سعر الخدمة"
            .Columns("PriceCd").Width = 120
            .Columns("IsActive").Width = 110
            .Columns("PriceCd").DefaultCellStyle.Format = "N3"
        End With
        TxtNoItems.Text = DGVItems.RowCount
        TestNewSave = True
        BtnSave.Enabled = True
        BtnDel.Enabled = False
        TxtPrice.TextAlign = HorizontalAlignment.Center
        TxtPrice.Text = 0
        TxtPrice.Text = FormatNumber(TxtPrice.Text, 3, , , TriState.True)
        ChkIsActive.Checked = True
        TxtFind2.Text = Nothing
        TxtNo.Text = 0
        TxtCardName.Clear()
        CmbSrv.Focus()
    End Sub




    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        'التحقق من البيانات قبل الحفظ

        If CmbSrv.Text = Nothing Then
            MessageBox.Show("اسف .. من فضلك قم اختيار اسم الخدمة في مربع الادخال ...", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            CmbSrv.Focus()
            Exit Sub
            'ElseIf Information.IsNumeric(TxtCodeUnit.Text) = False Then
            '    MessageBox.Show("اسف .. من فضلك قم بكتابة قيمة رقمية لرمز الخدمة في مربع الادخال ...", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            '    TxtCodeUnit.Focus()
            '    Exit Sub
        End If



        If TxtCardName.Text = Nothing Then
            MessageBox.Show("اسف .. من فضلك قم بكتابة إسم قئة البطاقة في مربع الادخال ...", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TxtCardName.Focus()
            Exit Sub
        End If

        If Information.IsNumeric(TxtPrice.Text) = False Or TxtPrice.Text = Nothing Then
            MessageBox.Show("اسف .. من فضلك قم بكتابة قيمة رقمية في مربع الادخال ...", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TxtPrice.Focus()
            Exit Sub
        End If

        ''حفظ البيانات المدخلة

        'If ChkIsDefult.Checked = True Then
        '    'الغاء قيمة الحقل الافتراضي في جميع بيانات الجدول
        '    Sql = "UPDATE [dbo].[TbCurrency] SET [IsDefult] = 'False'"
        '    MyClss.RunsInsertDeleteUpdateQry(Sql)
        'End If


      
        ''" + Format(CDate(DateTimePicker1.Value.ToString), "yyyy/MM/dd") + "'
        If TestNewSave = True Then
            Sql = "INSERT INTO  TbTypeCards(srvid, CardName, IsActive,PriceCd)"
            Sql += "VALUES (" + CmbSrv.SelectedValue.ToString + ",'" + TxtCardName.Text.ToString + "'," + ChkIsActive.Checked.ToString + "," + TxtPrice.Text.ToString + ")"
        Else
            Sql = "UPDATE TbTypeCards SET [srvid] = " + CmbSrv.SelectedValue.ToString + ", [IsActive] = " + ChkIsActive.Checked.ToString + ", [CardName] = '" + TxtCardName.Text.ToString + "',[PriceCd] = " + TxtPrice.Text + ""
            Sql += " WHERE ([IDCard] =" + DGVItems.CurrentRow.Cells("IDCard").Value.ToString + ")"
        End If
        MyClss.RunsInsertDeleteUpdateQry(Sql)
        If TestRunCode = True Then
            If TestNewSave = True Then
                '1 New - 2 Insert - 3  Edit  - 4 Del - 5 Print - 6 Find - 7 Login - 8 Logout
                MyClss.UserMovement(UserID, "2", TxtCardName.Text, Me.Text)
            Else
                MyClss.UserMovement(UserID, "3", TxtCardName.Text, Me.Text)
            End If
            BtnNew_Click(sender, e)
            MessageBox.Show("تم حفظ البيــــانات بنجاح", "حفظ", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("لم يتم حفظ البيــــانات ", "خطاء في عملية الحفظ", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        End If
    End Sub

    Private Sub BtnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnFind.Click

    End Sub

    Private Sub BtnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click

    End Sub

    Private Sub BtnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDel.Click
        Dim n As Int32 = MyClss.ExecuteScalar("SELECT Count(*) From TbDataCllinets Where srvid='" + DGVItems.CurrentRow.Cells("srvid").Value.ToString + "'")
        'TbInvoiceDetalis
        If n > 0 Then
            MessageBox.Show("أسف لا يمكن حذف البيانات  .. هناك بيانات في جدول اخر مرتبطة بها  ", "فشل عملية الحذف", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
            Exit Sub
        End If


        Dim resault As Integer
        Dim sql As String = "DELETE FROM TbTypeCards   WHERE ([IDCard] ='" + DGVItems.CurrentRow.Cells("IDCard").Value.ToString + "')"
        resault = MessageBox.Show("هل انت متأكد من عملية حذف البيانات ؟", "رسالة حذف ", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
        If resault = vbYes Then
            MyClss.RunsInsertDeleteUpdateQry(sql)
            If TestRunCode = True Then
                '1 New - 2 Insert - 3  Edit  - 4 Del - 5 Print - 6 Find - 7 Login - 8 Logout
                MyClss.UserMovement(UserID, "4", TxtCardName.Text, Me.Text)

                BtnNew_Click(sender, e)
                MessageBox.Show("لقد تم حذف البيانات ", "نجاح عملية الحذف", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
            Else
                MessageBox.Show("لم يتم حذف البيــــانات ", "خطاء في عملية الحذف", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                Exit Sub
            End If
        Else
            MessageBox.Show("تم ايقاف عملية الحذف", "حذف سجل", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
        End If
    End Sub

    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        Me.Close()
    End Sub

    Private Sub DGVItems_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVItems.CellClick
        If DGVItems.RowCount = 0 Then Exit Sub
        With DGVItems.CurrentRow
            CmbSrv.SelectedValue = .Cells("srvid").Value
            TxtCardName.Text = .Cells("CardName").Value
            TxtNo.Text = .Cells("IDCard").Value
            ChkIsActive.Checked = .Cells("IsActive").Value
            TxtPrice.Text = .Cells("PriceCd").Value
            TxtPrice.Text = FormatNumber(TxtPrice.Text, 3, , , TriState.True)
        End With
        BtnDel.Enabled = True
        TestNewSave = False
    End Sub

    Private Sub DGVItems_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DGVItems.CellContentClick

    End Sub

    Private Sub DGVItems_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DGVItems.KeyUp
        If DGVItems.RowCount = 0 Then Exit Sub
        If e.KeyCode = Keys.Up Or e.KeyCode = Keys.Down Then
            With DGVItems.CurrentRow
                CmbSrv.SelectedValue = .Cells("srvid").Value
                TxtCardName.Text = .Cells("CardName").Value
                TxtNo.Text = .Cells("IDCard").Value
                ChkIsActive.Checked = .Cells("IsActive").Value
                TxtPrice.Text = .Cells("PriceCd").Value
            End With
            BtnDel.Enabled = True
            TestNewSave = False
            TxtPrice.Text = FormatNumber(TxtPrice.Text, 3, , , TriState.True)
        End If
    End Sub

    Private Sub FrmTypeCards_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        ConvertTabKeyToEnterKey(e)
        If e.KeyCode = Keys.F2 Then BtnNew_Click(sender, e)
        If e.KeyCode = Keys.F3 Then BtnSave_Click(sender, e)
        If e.KeyCode = Keys.F4 And BtnDel.Enabled = True Then BtnDel_Click(sender, e)
        If e.KeyCode = Keys.Escape Or e.KeyCode = Keys.F12 Then BtnExit_Click(sender, e)
    End Sub

    Private Sub FrmTypeCards_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        BtnNew_Click(sender, e)
    End Sub
End Class