<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmTypeCards
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmTypeCards))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TxtNo = New System.Windows.Forms.Label()
        Me.BtnPrint = New System.Windows.Forms.ToolStripButton()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ToolStripLabel3 = New System.Windows.Forms.ToolStripLabel()
        Me.BtnFind = New System.Windows.Forms.ToolStripButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.DGVItems = New System.Windows.Forms.DataGridView()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.TxtNoItems = New System.Windows.Forms.ToolStripStatusLabel()
        Me.BtnExit = New System.Windows.Forms.ToolStripButton()
        Me.TxtCardName = New System.Windows.Forms.TextBox()
        Me.ToolStripSeparator6 = New System.Windows.Forms.ToolStripSeparator()
        Me.BtnDel = New System.Windows.Forms.ToolStripButton()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ChkIsActive = New System.Windows.Forms.CheckBox()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BtnSave = New System.Windows.Forms.ToolStripButton()
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.TxtPrice = New System.Windows.Forms.TextBox()
        Me.CmbSrv = New System.Windows.Forms.ComboBox()
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.BtnNew = New System.Windows.Forms.ToolStripButton()
        Me.TxtFind2 = New System.Windows.Forms.ToolStripTextBox()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        CType(Me.DGVItems, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.White
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label9.Location = New System.Drawing.Point(0, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(732, 26)
        Me.Label9.TabIndex = 29
        '
        'TxtNo
        '
        Me.TxtNo.AutoSize = True
        Me.TxtNo.BackColor = System.Drawing.SystemColors.Window
        Me.TxtNo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TxtNo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TxtNo.Location = New System.Drawing.Point(655, 30)
        Me.TxtNo.Name = "TxtNo"
        Me.TxtNo.Size = New System.Drawing.Size(65, 28)
        Me.TxtNo.TabIndex = 177
        Me.TxtNo.Text = "0"
        Me.TxtNo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BtnPrint
        '
        Me.BtnPrint.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BtnPrint.Image = CType(resources.GetObject("BtnPrint.Image"), System.Drawing.Image)
        Me.BtnPrint.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(28, 30)
        Me.BtnPrint.Text = "طباعة الوثيقة"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label2.Location = New System.Drawing.Point(655, 1)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 28)
        Me.Label2.TabIndex = 176
        Me.Label2.Text = "ر.م"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ToolStripLabel3
        '
        Me.ToolStripLabel3.Name = "ToolStripLabel3"
        Me.ToolStripLabel3.Size = New System.Drawing.Size(52, 30)
        Me.ToolStripLabel3.Text = "نص البحث"
        '
        'BtnFind
        '
        Me.BtnFind.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BtnFind.Image = CType(resources.GetObject("BtnFind.Image"), System.Drawing.Image)
        Me.BtnFind.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BtnFind.Name = "BtnFind"
        Me.BtnFind.Size = New System.Drawing.Size(28, 30)
        Me.BtnFind.Text = "بحث عن وثيقة"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label4.Location = New System.Drawing.Point(169, 1)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(254, 28)
        Me.Label4.TabIndex = 182
        Me.Label4.Text = "اسم الكرت"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 33)
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(6, 33)
        '
        'DGVItems
        '
        Me.DGVItems.AllowUserToAddRows = False
        Me.DGVItems.AllowUserToDeleteRows = False
        Me.DGVItems.AllowUserToOrderColumns = True
        Me.DGVItems.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DGVItems.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DGVItems.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DGVItems.DefaultCellStyle = DataGridViewCellStyle2
        Me.DGVItems.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVItems.Location = New System.Drawing.Point(0, 0)
        Me.DGVItems.Name = "DGVItems"
        Me.DGVItems.ReadOnly = True
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DGVItems.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.DGVItems.RowsDefaultCellStyle = DataGridViewCellStyle4
        Me.DGVItems.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGVItems.Size = New System.Drawing.Size(730, 448)
        Me.DGVItems.TabIndex = 168
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label3.Location = New System.Drawing.Point(430, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(218, 28)
        Me.Label3.TabIndex = 178
        Me.Label3.Text = "نوع الخدمة"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.TxtNoItems})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(732, 29)
        Me.StatusStrip1.TabIndex = 1
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(61, 24)
        Me.ToolStripStatusLabel1.Text = "عدد البيانات"
        '
        'TxtNoItems
        '
        Me.TxtNoItems.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.TxtNoItems.Name = "TxtNoItems"
        Me.TxtNoItems.Size = New System.Drawing.Size(14, 24)
        Me.TxtNoItems.Text = "0"
        '
        'BtnExit
        '
        Me.BtnExit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BtnExit.Image = CType(resources.GetObject("BtnExit.Image"), System.Drawing.Image)
        Me.BtnExit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(28, 30)
        Me.BtnExit.Text = "خؤوج ESC"
        '
        'TxtCardName
        '
        Me.TxtCardName.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TxtCardName.Location = New System.Drawing.Point(169, 33)
        Me.TxtCardName.Name = "TxtCardName"
        Me.TxtCardName.Size = New System.Drawing.Size(254, 24)
        Me.TxtCardName.TabIndex = 180
        '
        'ToolStripSeparator6
        '
        Me.ToolStripSeparator6.Name = "ToolStripSeparator6"
        Me.ToolStripSeparator6.Size = New System.Drawing.Size(6, 33)
        '
        'BtnDel
        '
        Me.BtnDel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BtnDel.Image = CType(resources.GetObject("BtnDel.Image"), System.Drawing.Image)
        Me.BtnDel.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BtnDel.Name = "BtnDel"
        Me.BtnDel.Size = New System.Drawing.Size(28, 30)
        Me.BtnDel.Text = "حذف F4"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label5.Location = New System.Drawing.Point(4, 1)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(78, 28)
        Me.Label5.TabIndex = 183
        Me.Label5.Text = "التفعيل"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ChkIsActive
        '
        Me.ChkIsActive.AutoSize = True
        Me.ChkIsActive.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ChkIsActive.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ChkIsActive.Location = New System.Drawing.Point(4, 33)
        Me.ChkIsActive.Name = "ChkIsActive"
        Me.ChkIsActive.Size = New System.Drawing.Size(78, 22)
        Me.ChkIsActive.TabIndex = 181
        Me.ChkIsActive.UseVisualStyleBackColor = True
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 33)
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label1.Location = New System.Drawing.Point(89, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 28)
        Me.Label1.TabIndex = 184
        Me.Label1.Text = "السعر"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'BtnSave
        '
        Me.BtnSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BtnSave.Image = CType(resources.GetObject("BtnSave.Image"), System.Drawing.Image)
        Me.BtnSave.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BtnSave.Name = "BtnSave"
        Me.BtnSave.Size = New System.Drawing.Size(28, 30)
        Me.BtnSave.Text = "حفظ F3"
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.IsSplitterFixed = True
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.TableLayoutPanel1)
        Me.SplitContainer3.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer3.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.SplitContainer3.Size = New System.Drawing.Size(732, 595)
        Me.SplitContainer3.SplitterDistance = 102
        Me.SplitContainer3.TabIndex = 0
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.[Single]
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label10, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel2, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 35.29412!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 64.70588!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(732, 102)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.LightSkyBlue
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label10.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label10.Font = New System.Drawing.Font("Tahoma", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label10.ForeColor = System.Drawing.Color.White
        Me.Label10.Image = CType(resources.GetObject("Label10.Image"), System.Drawing.Image)
        Me.Label10.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label10.Location = New System.Drawing.Point(4, 1)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(724, 34)
        Me.Label10.TabIndex = 40
        Me.Label10.Text = "بيانات فئات الكروت"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.AutoSize = True
        Me.TableLayoutPanel2.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.[Single]
        Me.TableLayoutPanel2.ColumnCount = 5
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 24.31507!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 75.68493!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 260.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 79.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 83.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.Label4, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.TxtNo, 0, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Label2, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Label3, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.TxtCardName, 2, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Label5, 4, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.ChkIsActive, 4, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.Label1, 3, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.TxtPrice, 3, 1)
        Me.TableLayoutPanel2.Controls.Add(Me.CmbSrv, 1, 1)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(4, 39)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 2
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(724, 59)
        Me.TableLayoutPanel2.TabIndex = 0
        '
        'TxtPrice
        '
        Me.TxtPrice.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TxtPrice.Location = New System.Drawing.Point(89, 33)
        Me.TxtPrice.Name = "TxtPrice"
        Me.TxtPrice.Size = New System.Drawing.Size(73, 24)
        Me.TxtPrice.TabIndex = 185
        '
        'CmbSrv
        '
        Me.CmbSrv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CmbSrv.FormattingEnabled = True
        Me.CmbSrv.Location = New System.Drawing.Point(430, 33)
        Me.CmbSrv.Name = "CmbSrv"
        Me.CmbSrv.Size = New System.Drawing.Size(218, 24)
        Me.CmbSrv.TabIndex = 186
        '
        'SplitContainer4
        '
        Me.SplitContainer4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.IsSplitterFixed = True
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Name = "SplitContainer4"
        Me.SplitContainer4.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.ToolStrip1)
        Me.SplitContainer4.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.DGVItems)
        Me.SplitContainer4.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.SplitContainer4.Size = New System.Drawing.Size(732, 489)
        Me.SplitContainer4.SplitterDistance = 35
        Me.SplitContainer4.TabIndex = 0
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BtnNew, Me.BtnSave, Me.ToolStripSeparator3, Me.ToolStripLabel3, Me.TxtFind2, Me.ToolStripSeparator4, Me.BtnFind, Me.BtnPrint, Me.ToolStripSeparator5, Me.BtnDel, Me.ToolStripSeparator6, Me.BtnExit})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(730, 33)
        Me.ToolStrip1.TabIndex = 2
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'BtnNew
        '
        Me.BtnNew.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BtnNew.Image = CType(resources.GetObject("BtnNew.Image"), System.Drawing.Image)
        Me.BtnNew.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BtnNew.Name = "BtnNew"
        Me.BtnNew.Size = New System.Drawing.Size(28, 30)
        Me.BtnNew.Text = "جديد F2"
        '
        'TxtFind2
        '
        Me.TxtFind2.Name = "TxtFind2"
        Me.TxtFind2.Size = New System.Drawing.Size(174, 33)
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.IsSplitterFixed = True
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.SplitContainer3)
        Me.SplitContainer2.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.StatusStrip1)
        Me.SplitContainer2.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.SplitContainer2.Size = New System.Drawing.Size(732, 628)
        Me.SplitContainer2.SplitterDistance = 595
        Me.SplitContainer2.TabIndex = 0
        '
        'SplitContainer1
        '
        Me.SplitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label9)
        Me.SplitContainer1.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
        Me.SplitContainer1.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.SplitContainer1.Size = New System.Drawing.Size(734, 662)
        Me.SplitContainer1.SplitterDistance = 28
        Me.SplitContainer1.TabIndex = 2
        '
        'FrmTypeCards
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(734, 662)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.Name = "FrmTypeCards"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.RightToLeftLayout = True
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "أنواع فئات الكروت"
        CType(Me.DGVItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer4.ResumeLayout(False)
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TxtNo As System.Windows.Forms.Label
    Friend WithEvents BtnPrint As System.Windows.Forms.ToolStripButton
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ToolStripLabel3 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents BtnFind As System.Windows.Forms.ToolStripButton
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents DGVItems As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TxtNoItems As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents BtnExit As System.Windows.Forms.ToolStripButton
    Friend WithEvents TxtCardName As System.Windows.Forms.TextBox
    Friend WithEvents ToolStripSeparator6 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BtnDel As System.Windows.Forms.ToolStripButton
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ChkIsActive As System.Windows.Forms.CheckBox
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents BtnSave As System.Windows.Forms.ToolStripButton
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TxtPrice As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents BtnNew As System.Windows.Forms.ToolStripButton
    Friend WithEvents TxtFind2 As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents CmbSrv As System.Windows.Forms.ComboBox
End Class
