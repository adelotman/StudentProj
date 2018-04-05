<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmRepScheduls
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmRepScheduls))
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.DGVItems = New System.Windows.Forms.DataGridView()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.BtnRefrech = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BtnFirstBal = New System.Windows.Forms.ToolStripButton()
        Me.DisplayPay = New System.Windows.Forms.ToolStripButton()
        Me.BtnFindInvf = New System.Windows.Forms.ToolStripButton()
        Me.BtnCal = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel()
        Me.CmbTypePrice = New System.Windows.Forms.ToolStripComboBox()
        Me.BtnPrint = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.BtnExit = New System.Windows.Forms.ToolStripButton()
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.BtnCnv = New System.Windows.Forms.Button()
        Me.NumYearsCnv = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.CmbStocksCnv = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.NumRec = New System.Windows.Forms.NumericUpDown()
        Me.CmbGroups = New System.Windows.Forms.ComboBox()
        Me.BtnFind = New System.Windows.Forms.Button()
        Me.ChkFindMoney = New System.Windows.Forms.CheckBox()
        Me.CmbOpratorFind = New System.Windows.Forms.ComboBox()
        Me.ListRealItems = New System.Windows.Forms.ListBox()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.CmbYears = New System.Windows.Forms.ComboBox()
        Me.CmbCourses = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.TxtNoItems = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ProgBar = New System.Windows.Forms.ToolStripProgressBar()
        Me.ChkCurse = New System.Windows.Forms.CheckBox()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        CType(Me.DGVItems, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.NumYearsCnv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.NumRec, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.AutoSize = True
        Me.TableLayoutPanel2.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.[Single]
        Me.TableLayoutPanel2.ColumnCount = 1
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel3, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.DGVItems, 0, 1)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 2
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 6.055046!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 93.94495!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(819, 546)
        Me.TableLayoutPanel2.TabIndex = 2
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 1
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.71675!))
        Me.TableLayoutPanel3.Controls.Add(Me.Label7, 0, 0)
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(4, 4)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 1
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(811, 26)
        Me.TableLayoutPanel3.TabIndex = 16
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.HotTrack
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label7.ForeColor = System.Drawing.Color.White
        Me.Label7.Image = CType(resources.GetObject("Label7.Image"), System.Drawing.Image)
        Me.Label7.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label7.Location = New System.Drawing.Point(3, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(805, 26)
        Me.Label7.TabIndex = 45
        Me.Label7.Text = "الجدول الدراسي"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DGVItems
        '
        Me.DGVItems.AllowUserToAddRows = False
        Me.DGVItems.AllowUserToDeleteRows = False
        Me.DGVItems.AllowUserToOrderColumns = True
        Me.DGVItems.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DGVItems.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.DGVItems.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGVItems.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DGVItems.Location = New System.Drawing.Point(4, 37)
        Me.DGVItems.Name = "DGVItems"
        Me.DGVItems.ReadOnly = True
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.DGVItems.RowsDefaultCellStyle = DataGridViewCellStyle8
        Me.DGVItems.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGVItems.Size = New System.Drawing.Size(811, 505)
        Me.DGVItems.TabIndex = 17
        Me.DGVItems.VirtualMode = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
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
        Me.SplitContainer1.Size = New System.Drawing.Size(1174, 651)
        Me.SplitContainer1.SplitterDistance = 25
        Me.SplitContainer1.TabIndex = 1
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.White
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label9.Location = New System.Drawing.Point(0, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(1174, 25)
        Me.Label9.TabIndex = 34
        '
        'SplitContainer2
        '
        Me.SplitContainer2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
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
        Me.SplitContainer2.Size = New System.Drawing.Size(1174, 622)
        Me.SplitContainer2.SplitterDistance = 593
        Me.SplitContainer2.TabIndex = 0
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.ToolStrip1)
        Me.SplitContainer3.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer3.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.SplitContainer3.Size = New System.Drawing.Size(1172, 591)
        Me.SplitContainer3.SplitterDistance = 39
        Me.SplitContainer3.TabIndex = 0
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BtnRefrech, Me.ToolStripSeparator1, Me.BtnFirstBal, Me.DisplayPay, Me.BtnFindInvf, Me.BtnCal, Me.ToolStripSeparator4, Me.ToolStripLabel1, Me.CmbTypePrice, Me.BtnPrint, Me.ToolStripSeparator5, Me.BtnExit})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1172, 39)
        Me.ToolStrip1.Stretch = True
        Me.ToolStrip1.TabIndex = 3
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'BtnRefrech
        '
        Me.BtnRefrech.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BtnRefrech.Image = CType(resources.GetObject("BtnRefrech.Image"), System.Drawing.Image)
        Me.BtnRefrech.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.BtnRefrech.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BtnRefrech.Name = "BtnRefrech"
        Me.BtnRefrech.Size = New System.Drawing.Size(36, 36)
        Me.BtnRefrech.Text = "تحديث F2"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 39)
        '
        'BtnFirstBal
        '
        Me.BtnFirstBal.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BtnFirstBal.Image = CType(resources.GetObject("BtnFirstBal.Image"), System.Drawing.Image)
        Me.BtnFirstBal.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.BtnFirstBal.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BtnFirstBal.Name = "BtnFirstBal"
        Me.BtnFirstBal.Size = New System.Drawing.Size(36, 36)
        '
        'DisplayPay
        '
        Me.DisplayPay.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.DisplayPay.Image = CType(resources.GetObject("DisplayPay.Image"), System.Drawing.Image)
        Me.DisplayPay.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.DisplayPay.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.DisplayPay.Name = "DisplayPay"
        Me.DisplayPay.Size = New System.Drawing.Size(36, 36)
        '
        'BtnFindInvf
        '
        Me.BtnFindInvf.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BtnFindInvf.Image = CType(resources.GetObject("BtnFindInvf.Image"), System.Drawing.Image)
        Me.BtnFindInvf.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.BtnFindInvf.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BtnFindInvf.Name = "BtnFindInvf"
        Me.BtnFindInvf.Size = New System.Drawing.Size(36, 36)
        '
        'BtnCal
        '
        Me.BtnCal.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BtnCal.Image = CType(resources.GetObject("BtnCal.Image"), System.Drawing.Image)
        Me.BtnCal.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.BtnCal.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BtnCal.Name = "BtnCal"
        Me.BtnCal.Size = New System.Drawing.Size(36, 36)
        Me.BtnCal.Text = "الالة الحاسبة"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 39)
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.ForeColor = System.Drawing.Color.Red
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(76, 36)
        Me.ToolStripLabel1.Text = "خيارات الطباعة"
        '
        'CmbTypePrice
        '
        Me.CmbTypePrice.Items.AddRange(New Object() {"طباعة الجدول الدراسي", "تحويل الى اكسل"})
        Me.CmbTypePrice.Name = "CmbTypePrice"
        Me.CmbTypePrice.Size = New System.Drawing.Size(121, 39)
        '
        'BtnPrint
        '
        Me.BtnPrint.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BtnPrint.Image = CType(resources.GetObject("BtnPrint.Image"), System.Drawing.Image)
        Me.BtnPrint.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.BtnPrint.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BtnPrint.Name = "BtnPrint"
        Me.BtnPrint.Size = New System.Drawing.Size(36, 36)
        Me.BtnPrint.Text = "طباعة F3"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(6, 39)
        '
        'BtnExit
        '
        Me.BtnExit.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.BtnExit.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BtnExit.Image = CType(resources.GetObject("BtnExit.Image"), System.Drawing.Image)
        Me.BtnExit.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.BtnExit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(36, 36)
        Me.BtnExit.Text = "خروج"
        '
        'SplitContainer4
        '
        Me.SplitContainer4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.TableLayoutPanel1)
        Me.SplitContainer4.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.TableLayoutPanel2)
        Me.SplitContainer4.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.SplitContainer4.Size = New System.Drawing.Size(1172, 548)
        Me.SplitContainer4.SplitterDistance = 347
        Me.SplitContainer4.TabIndex = 0
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.AutoSize = True
        Me.TableLayoutPanel1.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.[Single]
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Label19, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5.206463!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 94.79353!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(345, 546)
        Me.TableLayoutPanel1.TabIndex = 1
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.Crimson
        Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label19.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label19.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label19.ForeColor = System.Drawing.Color.White
        Me.Label19.Image = CType(resources.GetObject("Label19.Image"), System.Drawing.Image)
        Me.Label19.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label19.Location = New System.Drawing.Point(4, 1)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(337, 28)
        Me.Label19.TabIndex = 44
        Me.Label19.Text = "خيارات البحث"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(4, 33)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(337, 509)
        Me.Panel1.TabIndex = 45
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BtnCnv)
        Me.GroupBox2.Controls.Add(Me.NumYearsCnv)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.CmbStocksCnv)
        Me.GroupBox2.Location = New System.Drawing.Point(9, 406)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(323, 100)
        Me.GroupBox2.TabIndex = 40
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "خيارات تحويل"
        '
        'BtnCnv
        '
        Me.BtnCnv.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnCnv.Image = CType(resources.GetObject("BtnCnv.Image"), System.Drawing.Image)
        Me.BtnCnv.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnCnv.Location = New System.Drawing.Point(6, 46)
        Me.BtnCnv.Name = "BtnCnv"
        Me.BtnCnv.Size = New System.Drawing.Size(97, 49)
        Me.BtnCnv.TabIndex = 67
        Me.BtnCnv.Text = "تحويل"
        Me.BtnCnv.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BtnCnv.UseVisualStyleBackColor = True
        '
        'NumYearsCnv
        '
        Me.NumYearsCnv.Location = New System.Drawing.Point(119, 60)
        Me.NumYearsCnv.Name = "NumYearsCnv"
        Me.NumYearsCnv.Size = New System.Drawing.Size(106, 20)
        Me.NumYearsCnv.TabIndex = 65
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(231, 60)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 13)
        Me.Label1.TabIndex = 66
        Me.Label1.Text = "السنة الدراسية :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Red
        Me.Label5.Location = New System.Drawing.Point(232, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(74, 13)
        Me.Label5.TabIndex = 38
        Me.Label5.Text = "اسم الفصل :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CmbStocksCnv
        '
        Me.CmbStocksCnv.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.CmbStocksCnv.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CmbStocksCnv.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbStocksCnv.FormattingEnabled = True
        Me.CmbStocksCnv.Location = New System.Drawing.Point(6, 19)
        Me.CmbStocksCnv.Name = "CmbStocksCnv"
        Me.CmbStocksCnv.Size = New System.Drawing.Size(219, 21)
        Me.CmbStocksCnv.TabIndex = 37
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ChkCurse)
        Me.GroupBox1.Controls.Add(Me.NumRec)
        Me.GroupBox1.Controls.Add(Me.CmbGroups)
        Me.GroupBox1.Controls.Add(Me.BtnFind)
        Me.GroupBox1.Controls.Add(Me.ChkFindMoney)
        Me.GroupBox1.Controls.Add(Me.CmbOpratorFind)
        Me.GroupBox1.Controls.Add(Me.ListRealItems)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.CmbYears)
        Me.GroupBox1.Controls.Add(Me.CmbCourses)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(323, 400)
        Me.GroupBox1.TabIndex = 39
        Me.GroupBox1.TabStop = False
        '
        'NumRec
        '
        Me.NumRec.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.NumRec.Location = New System.Drawing.Point(11, 295)
        Me.NumRec.Name = "NumRec"
        Me.NumRec.Size = New System.Drawing.Size(120, 20)
        Me.NumRec.TabIndex = 71
        '
        'CmbGroups
        '
        Me.CmbGroups.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.CmbGroups.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CmbGroups.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbGroups.FormattingEnabled = True
        Me.CmbGroups.Location = New System.Drawing.Point(11, 41)
        Me.CmbGroups.Name = "CmbGroups"
        Me.CmbGroups.Size = New System.Drawing.Size(202, 21)
        Me.CmbGroups.TabIndex = 70
        '
        'BtnFind
        '
        Me.BtnFind.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnFind.Image = CType(resources.GetObject("BtnFind.Image"), System.Drawing.Image)
        Me.BtnFind.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnFind.Location = New System.Drawing.Point(6, 346)
        Me.BtnFind.Name = "BtnFind"
        Me.BtnFind.Size = New System.Drawing.Size(311, 48)
        Me.BtnFind.TabIndex = 5
        Me.BtnFind.Text = "بحث"
        Me.BtnFind.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.BtnFind.UseVisualStyleBackColor = True
        '
        'ChkFindMoney
        '
        Me.ChkFindMoney.AutoSize = True
        Me.ChkFindMoney.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Bold)
        Me.ChkFindMoney.Location = New System.Drawing.Point(226, 298)
        Me.ChkFindMoney.Name = "ChkFindMoney"
        Me.ChkFindMoney.Size = New System.Drawing.Size(79, 16)
        Me.ChkFindMoney.TabIndex = 69
        Me.ChkFindMoney.Text = "تكرار المادة :"
        Me.ChkFindMoney.UseVisualStyleBackColor = True
        '
        'CmbOpratorFind
        '
        Me.CmbOpratorFind.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.CmbOpratorFind.FormattingEnabled = True
        Me.CmbOpratorFind.Items.AddRange(New Object() {"=", ">", "<", ">=", "<="})
        Me.CmbOpratorFind.Location = New System.Drawing.Point(149, 296)
        Me.CmbOpratorFind.Name = "CmbOpratorFind"
        Me.CmbOpratorFind.Size = New System.Drawing.Size(64, 21)
        Me.CmbOpratorFind.TabIndex = 67
        Me.CmbOpratorFind.Text = "="
        '
        'ListRealItems
        '
        Me.ListRealItems.FormattingEnabled = True
        Me.ListRealItems.Location = New System.Drawing.Point(11, 72)
        Me.ListRealItems.Name = "ListRealItems"
        Me.ListRealItems.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.ListRealItems.Size = New System.Drawing.Size(202, 173)
        Me.ListRealItems.TabIndex = 2
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Font = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Label14.Location = New System.Drawing.Point(226, 150)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(98, 13)
        Me.Label14.TabIndex = 66
        Me.Label14.Text = "الفصل الدراسي :"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.BackColor = System.Drawing.Color.Transparent
        Me.Label11.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Red
        Me.Label11.Location = New System.Drawing.Point(226, 44)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(63, 13)
        Me.Label11.TabIndex = 64
        Me.Label11.Text = "المجموعة :"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CmbYears
        '
        Me.CmbYears.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.CmbYears.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CmbYears.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbYears.FormattingEnabled = True
        Me.CmbYears.Items.AddRange(New Object() {"سنة أولى", "سنة ثانية", "سنة ثالثة"})
        Me.CmbYears.Location = New System.Drawing.Point(11, 13)
        Me.CmbYears.Name = "CmbYears"
        Me.CmbYears.Size = New System.Drawing.Size(202, 21)
        Me.CmbYears.TabIndex = 0
        '
        'CmbCourses
        '
        Me.CmbCourses.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.CmbCourses.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CmbCourses.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmbCourses.FormattingEnabled = True
        Me.CmbCourses.Location = New System.Drawing.Point(11, 255)
        Me.CmbCourses.Name = "CmbCourses"
        Me.CmbCourses.Size = New System.Drawing.Size(202, 21)
        Me.CmbCourses.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(226, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(94, 13)
        Me.Label4.TabIndex = 38
        Me.Label4.Text = "السنة الدراسية :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.TxtNoItems, Me.ToolStripStatusLabel2, Me.ProgBar})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1172, 23)
        Me.StatusStrip1.TabIndex = 6
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(68, 18)
        Me.ToolStripStatusLabel1.Text = "عدد العناصر :"
        '
        'TxtNoItems
        '
        Me.TxtNoItems.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.TxtNoItems.Name = "TxtNoItems"
        Me.TxtNoItems.Size = New System.Drawing.Size(14, 18)
        Me.TxtNoItems.Text = "0"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.ForeColor = System.Drawing.Color.Red
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(72, 18)
        Me.ToolStripStatusLabel2.Text = "حالة المعالجة :"
        '
        'ProgBar
        '
        Me.ProgBar.Name = "ProgBar"
        Me.ProgBar.Size = New System.Drawing.Size(100, 17)
        '
        'ChkCurse
        '
        Me.ChkCurse.AutoSize = True
        Me.ChkCurse.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Bold)
        Me.ChkCurse.Location = New System.Drawing.Point(231, 255)
        Me.ChkCurse.Name = "ChkCurse"
        Me.ChkCurse.Size = New System.Drawing.Size(78, 16)
        Me.ChkCurse.TabIndex = 72
        Me.ChkCurse.Text = "اسم المادة :"
        Me.ChkCurse.UseVisualStyleBackColor = True
        '
        'FrmRepScheduls
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1174, 651)
        Me.Controls.Add(Me.SplitContainer1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.KeyPreview = True
        Me.Name = "FrmRepScheduls"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.RightToLeftLayout = True
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "تقاير الجدول الدراسي"
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        CType(Me.DGVItems, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.PerformLayout()
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.ResumeLayout(False)
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel1.PerformLayout()
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        Me.SplitContainer4.Panel2.PerformLayout()
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer4.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.NumYearsCnv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.NumRec, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel3 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents DGVItems As System.Windows.Forms.DataGridView
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents BtnRefrech As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BtnFirstBal As System.Windows.Forms.ToolStripButton
    Friend WithEvents DisplayPay As System.Windows.Forms.ToolStripButton
    Friend WithEvents BtnFindInvf As System.Windows.Forms.ToolStripButton
    Friend WithEvents BtnCal As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents CmbTypePrice As System.Windows.Forms.ToolStripComboBox
    Friend WithEvents BtnPrint As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BtnExit As System.Windows.Forms.ToolStripButton
    Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents BtnCnv As System.Windows.Forms.Button
    Friend WithEvents NumYearsCnv As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents CmbStocksCnv As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents BtnFind As System.Windows.Forms.Button
    Friend WithEvents ChkFindMoney As System.Windows.Forms.CheckBox
    Friend WithEvents CmbOpratorFind As System.Windows.Forms.ComboBox
    Friend WithEvents ListRealItems As System.Windows.Forms.ListBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents CmbYears As System.Windows.Forms.ComboBox
    Friend WithEvents CmbCourses As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents TxtNoItems As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ProgBar As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents CmbGroups As System.Windows.Forms.ComboBox
    Friend WithEvents NumRec As System.Windows.Forms.NumericUpDown
    Friend WithEvents ChkCurse As System.Windows.Forms.CheckBox

End Class
