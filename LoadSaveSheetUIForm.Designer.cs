namespace WFormProjEstimateApp1
{
    partial class WFormProjEstimate
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WFormProjEstimate));
            menuStrip1 = new MenuStrip();
            SourceFilelStripMenu = new ToolStripMenuItem();
            OpenTaskItemsSource = new ToolStripMenuItem();
            OpenTaskItemsWithArchiectSource = new ToolStripMenuItem();
            TargetFileStripMenu = new ToolStripMenuItem();
            SaveQuotationReportTarget = new ToolStripMenuItem();
            版本資訊ToolStripMenuItem = new ToolStripMenuItem();
            授權聲明ToolStripMenuItem = new ToolStripMenuItem();
            版本變更記錄ToolStripMenuItem = new ToolStripMenuItem();
            作者ToolStripMenuItem = new ToolStripMenuItem();
            產生範本檔案ToolStripMenuItem = new ToolStripMenuItem();
            開啟無架構師工作清單AToolStripMenuItem = new ToolStripMenuItem();
            CreateSourceWorkItemB = new ToolStripMenuItem();
            statusLable = new StatusStrip();
            StatusBarLabel = new ToolStripStatusLabel();
            mainWindowPanel = new Panel();
            projDepartmentGroup = new GroupBox();
            techRepresentativeComboBox = new ComboBox();
            techDepartmentComboBox = new ComboBox();
            techRepresentativeExtensionLabel = new Label();
            techRepresentativePhoneExtension = new TextBox();
            techRepresentativeEmailLabel = new Label();
            techRepresentativeEmailAddress = new TextBox();
            techDepartmentRepresentativelabel = new Label();
            techDepartmentLabel = new Label();
            projSalesGroup = new GroupBox();
            salesRepresentativeComboBox = new ComboBox();
            salesDepartmentComboBox = new ComboBox();
            salesRepresentativeExtensionLabel = new Label();
            salesRepresentativePhoneExtension = new TextBox();
            salesRepresentativeEmailLabel = new Label();
            salesRepresentativeEmailAddress = new TextBox();
            salesRepresentativeLabel = new Label();
            salesDepartmentLabel = new Label();
            projIdentityGroup = new GroupBox();
            projDeliveryGroup = new GroupBox();
            deleteDeliverableButton = new Button();
            modifyDeliverableButton = new Button();
            addDeliverableButton = new Button();
            deliverableListBox = new ListBox();
            deliverableSelectionComboBox = new ComboBox();
            customerName = new Label();
            customerNameTextBox = new TextBox();
            projectNameTextBox = new TextBox();
            projectName = new Label();
            menuStrip1.SuspendLayout();
            statusLable.SuspendLayout();
            mainWindowPanel.SuspendLayout();
            projDepartmentGroup.SuspendLayout();
            projSalesGroup.SuspendLayout();
            projIdentityGroup.SuspendLayout();
            projDeliveryGroup.SuspendLayout();
            SuspendLayout();
            // 
            // menuStrip1
            // 
            menuStrip1.ImageScalingSize = new Size(20, 20);
            menuStrip1.Items.AddRange(new ToolStripItem[] { SourceFilelStripMenu, TargetFileStripMenu, 版本資訊ToolStripMenuItem, 產生範本檔案ToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(693, 24);
            menuStrip1.TabIndex = 0;
            menuStrip1.Text = "menuStrip1";
            // 
            // SourceFilelStripMenu
            // 
            SourceFilelStripMenu.DropDownItems.AddRange(new ToolStripItem[] { OpenTaskItemsSource, OpenTaskItemsWithArchiectSource });
            SourceFilelStripMenu.Name = "SourceFilelStripMenu";
            SourceFilelStripMenu.Size = new Size(91, 20);
            SourceFilelStripMenu.Text = "讀取來源檔案";
            // 
            // OpenTaskItemsSource
            // 
            OpenTaskItemsSource.Name = "OpenTaskItemsSource";
            OpenTaskItemsSource.Size = new Size(228, 22);
            OpenTaskItemsSource.Text = "開啟(無)架構師工作清單...(&F)";
            OpenTaskItemsSource.Click += OpenWorkItemsSource_Click;
            // 
            // OpenTaskItemsWithArchiectSource
            // 
            OpenTaskItemsWithArchiectSource.Name = "OpenTaskItemsWithArchiectSource";
            OpenTaskItemsWithArchiectSource.Size = new Size(228, 22);
            OpenTaskItemsWithArchiectSource.Text = "開啟(含)架構師工作清單...(&G)";
            OpenTaskItemsWithArchiectSource.Click += OpenWorkItemsWithArchiectSource_Click;
            // 
            // TargetFileStripMenu
            // 
            TargetFileStripMenu.DropDownItems.AddRange(new ToolStripItem[] { SaveQuotationReportTarget });
            TargetFileStripMenu.Name = "TargetFileStripMenu";
            TargetFileStripMenu.Size = new Size(103, 20);
            TargetFileStripMenu.Text = "產生目標報價單";
            // 
            // SaveQuotationReportTarget
            // 
            SaveQuotationReportTarget.Name = "SaveQuotationReportTarget";
            SaveQuotationReportTarget.Size = new Size(194, 22);
            SaveQuotationReportTarget.Text = "產生目標報價單試算表";
            SaveQuotationReportTarget.Click += SaveQuotationReportTarget_Click;
            // 
            // 版本資訊ToolStripMenuItem
            // 
            版本資訊ToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 授權聲明ToolStripMenuItem, 版本變更記錄ToolStripMenuItem, 作者ToolStripMenuItem });
            版本資訊ToolStripMenuItem.Name = "版本資訊ToolStripMenuItem";
            版本資訊ToolStripMenuItem.Size = new Size(67, 20);
            版本資訊ToolStripMenuItem.Text = "版本資訊";
            // 
            // 授權聲明ToolStripMenuItem
            // 
            授權聲明ToolStripMenuItem.Name = "授權聲明ToolStripMenuItem";
            授權聲明ToolStripMenuItem.Size = new Size(146, 22);
            授權聲明ToolStripMenuItem.Text = "授權聲明";
            授權聲明ToolStripMenuItem.Click += LicensingStatement_Click;
            // 
            // 版本變更記錄ToolStripMenuItem
            // 
            版本變更記錄ToolStripMenuItem.Name = "版本變更記錄ToolStripMenuItem";
            版本變更記錄ToolStripMenuItem.Size = new Size(146, 22);
            版本變更記錄ToolStripMenuItem.Text = "版本變更記錄";
            版本變更記錄ToolStripMenuItem.Click += VersionDescription_Click;
            // 
            // 作者ToolStripMenuItem
            // 
            作者ToolStripMenuItem.Name = "作者ToolStripMenuItem";
            作者ToolStripMenuItem.Size = new Size(146, 22);
            作者ToolStripMenuItem.Text = "作者";
            作者ToolStripMenuItem.Click += AuthorInformation_Click;
            // 
            // 產生範本檔案ToolStripMenuItem
            // 
            產生範本檔案ToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { 開啟無架構師工作清單AToolStripMenuItem, CreateSourceWorkItemB });
            產生範本檔案ToolStripMenuItem.Name = "產生範本檔案ToolStripMenuItem";
            產生範本檔案ToolStripMenuItem.Size = new Size(91, 20);
            產生範本檔案ToolStripMenuItem.Text = "產生範本檔案";
            // 
            // 開啟無架構師工作清單AToolStripMenuItem
            // 
            開啟無架構師工作清單AToolStripMenuItem.Name = "開啟無架構師工作清單AToolStripMenuItem";
            開啟無架構師工作清單AToolStripMenuItem.Size = new Size(227, 22);
            開啟無架構師工作清單AToolStripMenuItem.Text = "建立(無)架構師工作清單...(&A)";
            開啟無架構師工作清單AToolStripMenuItem.Click += CreateSourceWorkItemA_Click;
            // 
            // CreateSourceWorkItemB
            // 
            CreateSourceWorkItemB.Name = "CreateSourceWorkItemB";
            CreateSourceWorkItemB.Size = new Size(227, 22);
            CreateSourceWorkItemB.Text = "建立(含)架構師工作清單...(&B)";
            CreateSourceWorkItemB.Click += CreateSourceWorkItemB_Click;
            // 
            // statusLable
            // 
            statusLable.ImageScalingSize = new Size(20, 20);
            statusLable.Items.AddRange(new ToolStripItem[] { StatusBarLabel });
            statusLable.Location = new Point(0, 418);
            statusLable.Name = "statusLable";
            statusLable.Size = new Size(693, 22);
            statusLable.TabIndex = 1;
            statusLable.Text = "狀態";
            // 
            // StatusBarLabel
            // 
            StatusBarLabel.Name = "StatusBarLabel";
            StatusBarLabel.Size = new Size(31, 17);
            StatusBarLabel.Text = "狀態";
            // 
            // mainWindowPanel
            // 
            mainWindowPanel.BorderStyle = BorderStyle.FixedSingle;
            mainWindowPanel.Controls.Add(projDepartmentGroup);
            mainWindowPanel.Controls.Add(projSalesGroup);
            mainWindowPanel.Controls.Add(projIdentityGroup);
            mainWindowPanel.ForeColor = SystemColors.ActiveCaptionText;
            mainWindowPanel.Location = new Point(12, 35);
            mainWindowPanel.Name = "mainWindowPanel";
            mainWindowPanel.Size = new Size(658, 372);
            mainWindowPanel.TabIndex = 2;
            // 
            // projDepartmentGroup
            // 
            projDepartmentGroup.Controls.Add(techRepresentativeComboBox);
            projDepartmentGroup.Controls.Add(techDepartmentComboBox);
            projDepartmentGroup.Controls.Add(techRepresentativeExtensionLabel);
            projDepartmentGroup.Controls.Add(techRepresentativePhoneExtension);
            projDepartmentGroup.Controls.Add(techRepresentativeEmailLabel);
            projDepartmentGroup.Controls.Add(techRepresentativeEmailAddress);
            projDepartmentGroup.Controls.Add(techDepartmentRepresentativelabel);
            projDepartmentGroup.Controls.Add(techDepartmentLabel);
            projDepartmentGroup.Location = new Point(357, 193);
            projDepartmentGroup.Name = "projDepartmentGroup";
            projDepartmentGroup.Size = new Size(274, 164);
            projDepartmentGroup.TabIndex = 8;
            projDepartmentGroup.TabStop = false;
            projDepartmentGroup.Text = "專案團隊";
            // 
            // techRepresentativeComboBox
            // 
            techRepresentativeComboBox.FormattingEnabled = true;
            techRepresentativeComboBox.Location = new Point(84, 64);
            techRepresentativeComboBox.Name = "techRepresentativeComboBox";
            techRepresentativeComboBox.Size = new Size(168, 23);
            techRepresentativeComboBox.TabIndex = 10;
            // 
            // techDepartmentComboBox
            // 
            techDepartmentComboBox.FormattingEnabled = true;
            techDepartmentComboBox.Location = new Point(84, 35);
            techDepartmentComboBox.Name = "techDepartmentComboBox";
            techDepartmentComboBox.Size = new Size(168, 23);
            techDepartmentComboBox.TabIndex = 10;
            // 
            // techRepresentativeExtensionLabel
            // 
            techRepresentativeExtensionLabel.AutoSize = true;
            techRepresentativeExtensionLabel.Location = new Point(21, 125);
            techRepresentativeExtensionLabel.Name = "techRepresentativeExtensionLabel";
            techRepresentativeExtensionLabel.Size = new Size(55, 15);
            techRepresentativeExtensionLabel.TabIndex = 17;
            techRepresentativeExtensionLabel.Text = "電話分機";
            // 
            // techRepresentativePhoneExtension
            // 
            techRepresentativePhoneExtension.Location = new Point(86, 122);
            techRepresentativePhoneExtension.Name = "techRepresentativePhoneExtension";
            techRepresentativePhoneExtension.Size = new Size(168, 23);
            techRepresentativePhoneExtension.TabIndex = 16;
            // 
            // techRepresentativeEmailLabel
            // 
            techRepresentativeEmailLabel.AutoSize = true;
            techRepresentativeEmailLabel.Location = new Point(20, 96);
            techRepresentativeEmailLabel.Name = "techRepresentativeEmailLabel";
            techRepresentativeEmailLabel.Size = new Size(55, 15);
            techRepresentativeEmailLabel.TabIndex = 15;
            techRepresentativeEmailLabel.Text = "電子信箱";
            // 
            // techRepresentativeEmailAddress
            // 
            techRepresentativeEmailAddress.Location = new Point(85, 93);
            techRepresentativeEmailAddress.Name = "techRepresentativeEmailAddress";
            techRepresentativeEmailAddress.Size = new Size(168, 23);
            techRepresentativeEmailAddress.TabIndex = 14;
            // 
            // techDepartmentRepresentativelabel
            // 
            techDepartmentRepresentativelabel.AutoSize = true;
            techDepartmentRepresentativelabel.Location = new Point(21, 67);
            techDepartmentRepresentativelabel.Name = "techDepartmentRepresentativelabel";
            techDepartmentRepresentativelabel.Size = new Size(55, 15);
            techDepartmentRepresentativelabel.TabIndex = 13;
            techDepartmentRepresentativelabel.Text = "部門代表";
            // 
            // techDepartmentLabel
            // 
            techDepartmentLabel.AutoSize = true;
            techDepartmentLabel.Location = new Point(20, 38);
            techDepartmentLabel.Name = "techDepartmentLabel";
            techDepartmentLabel.Size = new Size(55, 15);
            techDepartmentLabel.TabIndex = 11;
            techDepartmentLabel.Text = "技術部門";
            // 
            // projSalesGroup
            // 
            projSalesGroup.Controls.Add(salesRepresentativeComboBox);
            projSalesGroup.Controls.Add(salesDepartmentComboBox);
            projSalesGroup.Controls.Add(salesRepresentativeExtensionLabel);
            projSalesGroup.Controls.Add(salesRepresentativePhoneExtension);
            projSalesGroup.Controls.Add(salesRepresentativeEmailLabel);
            projSalesGroup.Controls.Add(salesRepresentativeEmailAddress);
            projSalesGroup.Controls.Add(salesRepresentativeLabel);
            projSalesGroup.Controls.Add(salesDepartmentLabel);
            projSalesGroup.Location = new Point(357, 19);
            projSalesGroup.Name = "projSalesGroup";
            projSalesGroup.Size = new Size(274, 159);
            projSalesGroup.TabIndex = 7;
            projSalesGroup.TabStop = false;
            projSalesGroup.Text = "業務團隊";
            // 
            // salesRepresentativeComboBox
            // 
            salesRepresentativeComboBox.FormattingEnabled = true;
            salesRepresentativeComboBox.Location = new Point(83, 58);
            salesRepresentativeComboBox.Name = "salesRepresentativeComboBox";
            salesRepresentativeComboBox.Size = new Size(168, 23);
            salesRepresentativeComboBox.TabIndex = 10;
            // 
            // salesDepartmentComboBox
            // 
            salesDepartmentComboBox.FormattingEnabled = true;
            salesDepartmentComboBox.Location = new Point(83, 29);
            salesDepartmentComboBox.Name = "salesDepartmentComboBox";
            salesDepartmentComboBox.Size = new Size(168, 23);
            salesDepartmentComboBox.TabIndex = 10;
            // 
            // salesRepresentativeExtensionLabel
            // 
            salesRepresentativeExtensionLabel.AutoSize = true;
            salesRepresentativeExtensionLabel.Location = new Point(19, 119);
            salesRepresentativeExtensionLabel.Name = "salesRepresentativeExtensionLabel";
            salesRepresentativeExtensionLabel.Size = new Size(55, 15);
            salesRepresentativeExtensionLabel.TabIndex = 9;
            salesRepresentativeExtensionLabel.Text = "電話分機";
            // 
            // salesRepresentativePhoneExtension
            // 
            salesRepresentativePhoneExtension.Location = new Point(84, 116);
            salesRepresentativePhoneExtension.Name = "salesRepresentativePhoneExtension";
            salesRepresentativePhoneExtension.Size = new Size(168, 23);
            salesRepresentativePhoneExtension.TabIndex = 8;
            // 
            // salesRepresentativeEmailLabel
            // 
            salesRepresentativeEmailLabel.AutoSize = true;
            salesRepresentativeEmailLabel.Location = new Point(18, 90);
            salesRepresentativeEmailLabel.Name = "salesRepresentativeEmailLabel";
            salesRepresentativeEmailLabel.Size = new Size(55, 15);
            salesRepresentativeEmailLabel.TabIndex = 7;
            salesRepresentativeEmailLabel.Text = "電子信箱";
            // 
            // salesRepresentativeEmailAddress
            // 
            salesRepresentativeEmailAddress.Location = new Point(83, 87);
            salesRepresentativeEmailAddress.Name = "salesRepresentativeEmailAddress";
            salesRepresentativeEmailAddress.Size = new Size(168, 23);
            salesRepresentativeEmailAddress.TabIndex = 6;
            // 
            // salesRepresentativeLabel
            // 
            salesRepresentativeLabel.AutoSize = true;
            salesRepresentativeLabel.Location = new Point(19, 61);
            salesRepresentativeLabel.Name = "salesRepresentativeLabel";
            salesRepresentativeLabel.Size = new Size(55, 15);
            salesRepresentativeLabel.TabIndex = 5;
            salesRepresentativeLabel.Text = "業務代表";
            // 
            // salesDepartmentLabel
            // 
            salesDepartmentLabel.AutoSize = true;
            salesDepartmentLabel.Location = new Point(18, 32);
            salesDepartmentLabel.Name = "salesDepartmentLabel";
            salesDepartmentLabel.Size = new Size(55, 15);
            salesDepartmentLabel.TabIndex = 3;
            salesDepartmentLabel.Text = "業務部門";
            // 
            // projIdentityGroup
            // 
            projIdentityGroup.Controls.Add(projDeliveryGroup);
            projIdentityGroup.Controls.Add(customerName);
            projIdentityGroup.Controls.Add(customerNameTextBox);
            projIdentityGroup.Controls.Add(projectNameTextBox);
            projIdentityGroup.Controls.Add(projectName);
            projIdentityGroup.Location = new Point(16, 19);
            projIdentityGroup.Name = "projIdentityGroup";
            projIdentityGroup.Size = new Size(322, 338);
            projIdentityGroup.TabIndex = 6;
            projIdentityGroup.TabStop = false;
            projIdentityGroup.Text = "專案與交付項目";
            // 
            // projDeliveryGroup
            // 
            projDeliveryGroup.Controls.Add(deleteDeliverableButton);
            projDeliveryGroup.Controls.Add(modifyDeliverableButton);
            projDeliveryGroup.Controls.Add(addDeliverableButton);
            projDeliveryGroup.Controls.Add(deliverableListBox);
            projDeliveryGroup.Controls.Add(deliverableSelectionComboBox);
            projDeliveryGroup.Location = new Point(19, 94);
            projDeliveryGroup.Name = "projDeliveryGroup";
            projDeliveryGroup.Size = new Size(284, 238);
            projDeliveryGroup.TabIndex = 6;
            projDeliveryGroup.TabStop = false;
            projDeliveryGroup.Text = "專案交付項目";
            // 
            // deleteDeliverableButton
            // 
            deleteDeliverableButton.Location = new Point(186, 209);
            deleteDeliverableButton.Name = "deleteDeliverableButton";
            deleteDeliverableButton.Size = new Size(75, 23);
            deleteDeliverableButton.TabIndex = 2;
            deleteDeliverableButton.Text = "刪除項目";
            deleteDeliverableButton.UseVisualStyleBackColor = true;
            deleteDeliverableButton.Click += DeleteDeliverableButton_Click;
            // 
            // modifyDeliverableButton
            // 
            modifyDeliverableButton.Location = new Point(101, 209);
            modifyDeliverableButton.Name = "modifyDeliverableButton";
            modifyDeliverableButton.Size = new Size(75, 23);
            modifyDeliverableButton.TabIndex = 2;
            modifyDeliverableButton.Text = "修改項目";
            modifyDeliverableButton.UseVisualStyleBackColor = true;
            modifyDeliverableButton.Click += ModifyDeliverableButton_Click;
            // 
            // addDeliverableButton
            // 
            addDeliverableButton.Location = new Point(16, 209);
            addDeliverableButton.Name = "addDeliverableButton";
            addDeliverableButton.Size = new Size(75, 23);
            addDeliverableButton.TabIndex = 2;
            addDeliverableButton.Text = "新增項目";
            addDeliverableButton.UseVisualStyleBackColor = true;
            addDeliverableButton.Click += AddDeliverableButton_Click;
            // 
            // deliverableListBox
            // 
            deliverableListBox.FormattingEnabled = true;
            deliverableListBox.Location = new Point(16, 30);
            deliverableListBox.Name = "deliverableListBox";
            deliverableListBox.Size = new Size(245, 124);
            deliverableListBox.TabIndex = 1;
            // 
            // deliverableSelectionComboBox
            // 
            deliverableSelectionComboBox.FormattingEnabled = true;
            deliverableSelectionComboBox.Location = new Point(16, 173);
            deliverableSelectionComboBox.Name = "deliverableSelectionComboBox";
            deliverableSelectionComboBox.Size = new Size(245, 23);
            deliverableSelectionComboBox.TabIndex = 0;
            // 
            // customerName
            // 
            customerName.AutoSize = true;
            customerName.Location = new Point(19, 32);
            customerName.Name = "customerName";
            customerName.Size = new Size(55, 15);
            customerName.TabIndex = 3;
            customerName.Text = "客戶名稱";
            // 
            // customerNameTextBox
            // 
            customerNameTextBox.Location = new Point(84, 29);
            customerNameTextBox.Name = "customerNameTextBox";
            customerNameTextBox.Size = new Size(219, 23);
            customerNameTextBox.TabIndex = 2;
            // 
            // projectNameTextBox
            // 
            projectNameTextBox.Location = new Point(85, 58);
            projectNameTextBox.Name = "projectNameTextBox";
            projectNameTextBox.Size = new Size(218, 23);
            projectNameTextBox.TabIndex = 4;
            // 
            // projectName
            // 
            projectName.AutoSize = true;
            projectName.Location = new Point(20, 61);
            projectName.Name = "projectName";
            projectName.Size = new Size(55, 15);
            projectName.TabIndex = 5;
            projectName.Text = "專案名稱";
            // 
            // WFormProjEstimate
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(693, 440);
            Controls.Add(mainWindowPanel);
            Controls.Add(statusLable);
            Controls.Add(menuStrip1);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = menuStrip1;
            MaximizeBox = false;
            Name = "WFormProjEstimate";
            Text = "工作項目成本估算報價單 (整合架構師欄位) - 0.1c.3 版";
            Load += WFormProjEstimate_Load;
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            statusLable.ResumeLayout(false);
            statusLable.PerformLayout();
            mainWindowPanel.ResumeLayout(false);
            projDepartmentGroup.ResumeLayout(false);
            projDepartmentGroup.PerformLayout();
            projSalesGroup.ResumeLayout(false);
            projSalesGroup.PerformLayout();
            projIdentityGroup.ResumeLayout(false);
            projIdentityGroup.PerformLayout();
            projDeliveryGroup.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private MenuStrip menuStrip1;
        private ToolStripMenuItem SourceFilelStripMenu;
        private ToolStripMenuItem OpenTaskItemsSource;
        private StatusStrip statusLable;
        private ToolStripStatusLabel StatusBarLabel;
        private ToolStripMenuItem TargetFileStripMenu;
        private ToolStripMenuItem SaveQuotationReportTarget;
        private Panel mainWindowPanel;
        private GroupBox projDepartmentGroup;
        private GroupBox projSalesGroup;
        private Label salesDepartmentLabel;
        private GroupBox projIdentityGroup;
        private Label salesRepresentativeLabel;
        private Label salesRepresentativeExtensionLabel;
        protected internal TextBox salesRepresentativePhoneExtension;
        private Label salesRepresentativeEmailLabel;
        protected internal TextBox salesRepresentativeEmailAddress;
        private Label techRepresentativeExtensionLabel;
        protected internal TextBox techRepresentativePhoneExtension;
        private Label techRepresentativeEmailLabel;
        protected internal TextBox techRepresentativeEmailAddress;
        private Label techDepartmentRepresentativelabel;
        private Label techDepartmentLabel;
        private Label customerName;
        private TextBox customerNameTextBox;
        protected internal TextBox projectNameTextBox;
        private Label projectName;
        private GroupBox projDeliveryGroup;
        private ComboBox deliverableSelectionComboBox;
        private Button deleteDeliverableButton;
        private Button modifyDeliverableButton;
        private Button addDeliverableButton;
        private ListBox deliverableListBox;
        private ComboBox techRepresentativeComboBox;
        private ComboBox techDepartmentComboBox;
        private ComboBox salesRepresentativeComboBox;
        private ComboBox salesDepartmentComboBox;
        private ToolStripMenuItem 版本資訊ToolStripMenuItem;
        private ToolStripMenuItem 授權聲明ToolStripMenuItem;
        private ToolStripMenuItem 版本變更記錄ToolStripMenuItem;
        private ToolStripMenuItem 作者ToolStripMenuItem;
        private ToolStripMenuItem OpenTaskItemsWithArchiectSource;
        private ToolStripMenuItem 產生範本檔案ToolStripMenuItem;
        private ToolStripMenuItem 開啟無架構師工作清單AToolStripMenuItem;
        private ToolStripMenuItem CreateSourceWorkItemB;
    }
}
