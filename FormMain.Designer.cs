namespace WLTDBWUI
{
    partial class FormMain
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.listViewLogFiles = new System.Windows.Forms.ListView();
            this.columnHeaderLogFiles = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.groupBoxLogFiles = new System.Windows.Forms.GroupBox();
            this.buttonOpenBackupDirectory = new System.Windows.Forms.Button();
            this.buttonOpenLogDirectory = new System.Windows.Forms.Button();
            this.buttonLoadLogFilesFromLogDirectory = new System.Windows.Forms.Button();
            this.buttonClearLogFiles = new System.Windows.Forms.Button();
            this.buttonLogFilesToDb = new System.Windows.Forms.Button();
            this.tabControlMain = new System.Windows.Forms.TabControl();
            this.tabLog = new System.Windows.Forms.TabPage();
            this.tabCsv = new System.Windows.Forms.TabPage();
            this.buttonOpenCsvDirectory = new System.Windows.Forms.Button();
            this.checkBoxExecCsv = new System.Windows.Forms.CheckBox();
            this.groupBoxWlIdAndAlias = new System.Windows.Forms.GroupBox();
            this.buttonMoveDownInWlIdsAndAliases = new System.Windows.Forms.Button();
            this.buttonMoveUpInWlIdsAndAliases = new System.Windows.Forms.Button();
            this.buttonRemoveFromWlIdsAndAliases = new System.Windows.Forms.Button();
            this.labelWlIdsAndAliases = new System.Windows.Forms.Label();
            this.labelSelectedAlias = new System.Windows.Forms.Label();
            this.labelSelectedWlId = new System.Windows.Forms.Label();
            this.buttonUseForCsv = new System.Windows.Forms.Button();
            this.listViewWlIdsAndAliases = new System.Windows.Forms.ListView();
            this.columnHeaderWlIds = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeaderAliases = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.listViewNewWlIds = new System.Windows.Forms.ListView();
            this.columnHeaderWlId = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.buttonUpdateAlias = new System.Windows.Forms.Button();
            this.labelWlId = new System.Windows.Forms.Label();
            this.textAlias = new System.Windows.Forms.TextBox();
            this.buttonLoadNewWlIdsFromDb = new System.Windows.Forms.Button();
            this.groupBoxTargetDateRange = new System.Windows.Forms.GroupBox();
            this.dateTimePickerEndDate = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerStartDate = new System.Windows.Forms.DateTimePicker();
            this.labelStartDate = new System.Windows.Forms.Label();
            this.labelEndDate = new System.Windows.Forms.Label();
            this.buttonCreateCsvFromDb = new System.Windows.Forms.Button();
            this.textBoxConsole = new System.Windows.Forms.TextBox();
            this.buttonClearConsole = new System.Windows.Forms.Button();
            this.groupBoxLogFiles.SuspendLayout();
            this.tabControlMain.SuspendLayout();
            this.tabLog.SuspendLayout();
            this.tabCsv.SuspendLayout();
            this.groupBoxWlIdAndAlias.SuspendLayout();
            this.groupBoxTargetDateRange.SuspendLayout();
            this.SuspendLayout();
            // 
            // listViewLogFiles
            // 
            this.listViewLogFiles.AllowDrop = true;
            this.listViewLogFiles.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewLogFiles.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderLogFiles});
            this.listViewLogFiles.FullRowSelect = true;
            this.listViewLogFiles.HideSelection = false;
            this.listViewLogFiles.Location = new System.Drawing.Point(8, 18);
            this.listViewLogFiles.Name = "listViewLogFiles";
            this.listViewLogFiles.Size = new System.Drawing.Size(725, 387);
            this.listViewLogFiles.TabIndex = 0;
            this.listViewLogFiles.UseCompatibleStateImageBehavior = false;
            this.listViewLogFiles.View = System.Windows.Forms.View.Details;
            this.listViewLogFiles.DragDrop += new System.Windows.Forms.DragEventHandler(this.listViewLogFiles_DragDrop);
            this.listViewLogFiles.DragEnter += new System.Windows.Forms.DragEventHandler(this.listViewLogFiles_DragEnter);
            // 
            // columnHeaderLogFiles
            // 
            this.columnHeaderLogFiles.Text = ".logファイル名";
            this.columnHeaderLogFiles.Width = 545;
            // 
            // groupBoxLogFiles
            // 
            this.groupBoxLogFiles.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxLogFiles.Controls.Add(this.buttonOpenBackupDirectory);
            this.groupBoxLogFiles.Controls.Add(this.buttonOpenLogDirectory);
            this.groupBoxLogFiles.Controls.Add(this.buttonLoadLogFilesFromLogDirectory);
            this.groupBoxLogFiles.Controls.Add(this.buttonClearLogFiles);
            this.groupBoxLogFiles.Controls.Add(this.listViewLogFiles);
            this.groupBoxLogFiles.Location = new System.Drawing.Point(6, 6);
            this.groupBoxLogFiles.Name = "groupBoxLogFiles";
            this.groupBoxLogFiles.Size = new System.Drawing.Size(740, 440);
            this.groupBoxLogFiles.TabIndex = 0;
            this.groupBoxLogFiles.TabStop = false;
            this.groupBoxLogFiles.Text = ".logファイル一覧";
            // 
            // buttonOpenBackupDirectory
            // 
            this.buttonOpenBackupDirectory.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonOpenBackupDirectory.Location = new System.Drawing.Point(397, 411);
            this.buttonOpenBackupDirectory.Name = "buttonOpenBackupDirectory";
            this.buttonOpenBackupDirectory.Size = new System.Drawing.Size(136, 23);
            this.buttonOpenBackupDirectory.TabIndex = 4;
            this.buttonOpenBackupDirectory.Text = "バックアップフォルダを開く";
            this.buttonOpenBackupDirectory.UseVisualStyleBackColor = true;
            this.buttonOpenBackupDirectory.Click += new System.EventHandler(this.buttonOpenBackupDirectory_Click);
            // 
            // buttonOpenLogDirectory
            // 
            this.buttonOpenLogDirectory.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonOpenLogDirectory.Location = new System.Drawing.Point(279, 411);
            this.buttonOpenLogDirectory.Name = "buttonOpenLogDirectory";
            this.buttonOpenLogDirectory.Size = new System.Drawing.Size(112, 23);
            this.buttonOpenLogDirectory.TabIndex = 3;
            this.buttonOpenLogDirectory.Text = "ログフォルダを開く";
            this.buttonOpenLogDirectory.UseVisualStyleBackColor = true;
            this.buttonOpenLogDirectory.Click += new System.EventHandler(this.buttonOpenLogDirectory_Click);
            // 
            // buttonLoadLogFilesFromLogDirectory
            // 
            this.buttonLoadLogFilesFromLogDirectory.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonLoadLogFilesFromLogDirectory.Location = new System.Drawing.Point(98, 411);
            this.buttonLoadLogFilesFromLogDirectory.Name = "buttonLoadLogFilesFromLogDirectory";
            this.buttonLoadLogFilesFromLogDirectory.Size = new System.Drawing.Size(175, 23);
            this.buttonLoadLogFilesFromLogDirectory.TabIndex = 2;
            this.buttonLoadLogFilesFromLogDirectory.Text = "ログフォルダの内容を反映する";
            this.buttonLoadLogFilesFromLogDirectory.UseVisualStyleBackColor = true;
            this.buttonLoadLogFilesFromLogDirectory.Click += new System.EventHandler(this.buttonLoadLogFilesFromLogDirectory_Click);
            // 
            // buttonClearLogFiles
            // 
            this.buttonClearLogFiles.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonClearLogFiles.Location = new System.Drawing.Point(8, 411);
            this.buttonClearLogFiles.Name = "buttonClearLogFiles";
            this.buttonClearLogFiles.Size = new System.Drawing.Size(84, 23);
            this.buttonClearLogFiles.TabIndex = 1;
            this.buttonClearLogFiles.Text = "一覧をクリア";
            this.buttonClearLogFiles.UseVisualStyleBackColor = true;
            this.buttonClearLogFiles.Click += new System.EventHandler(this.buttonClearLogFiles_Click);
            // 
            // buttonLogFilesToDb
            // 
            this.buttonLogFilesToDb.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonLogFilesToDb.Location = new System.Drawing.Point(526, 452);
            this.buttonLogFilesToDb.Name = "buttonLogFilesToDb";
            this.buttonLogFilesToDb.Size = new System.Drawing.Size(220, 23);
            this.buttonLogFilesToDb.TabIndex = 1;
            this.buttonLogFilesToDb.Text = ".logファイルを内部データベースに取り込む";
            this.buttonLogFilesToDb.UseVisualStyleBackColor = true;
            this.buttonLogFilesToDb.Click += new System.EventHandler(this.buttonLogFilesToDb_Click);
            // 
            // tabControlMain
            // 
            this.tabControlMain.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControlMain.Controls.Add(this.tabLog);
            this.tabControlMain.Controls.Add(this.tabCsv);
            this.tabControlMain.Location = new System.Drawing.Point(12, 12);
            this.tabControlMain.Name = "tabControlMain";
            this.tabControlMain.SelectedIndex = 0;
            this.tabControlMain.Size = new System.Drawing.Size(760, 507);
            this.tabControlMain.TabIndex = 0;
            // 
            // tabLog
            // 
            this.tabLog.Controls.Add(this.groupBoxLogFiles);
            this.tabLog.Controls.Add(this.buttonLogFilesToDb);
            this.tabLog.Location = new System.Drawing.Point(4, 22);
            this.tabLog.Name = "tabLog";
            this.tabLog.Padding = new System.Windows.Forms.Padding(3);
            this.tabLog.Size = new System.Drawing.Size(752, 481);
            this.tabLog.TabIndex = 0;
            this.tabLog.Text = ".logファイルの取込";
            this.tabLog.UseVisualStyleBackColor = true;
            // 
            // tabCsv
            // 
            this.tabCsv.Controls.Add(this.buttonOpenCsvDirectory);
            this.tabCsv.Controls.Add(this.checkBoxExecCsv);
            this.tabCsv.Controls.Add(this.groupBoxWlIdAndAlias);
            this.tabCsv.Controls.Add(this.groupBoxTargetDateRange);
            this.tabCsv.Controls.Add(this.buttonCreateCsvFromDb);
            this.tabCsv.Location = new System.Drawing.Point(4, 22);
            this.tabCsv.Name = "tabCsv";
            this.tabCsv.Padding = new System.Windows.Forms.Padding(3);
            this.tabCsv.Size = new System.Drawing.Size(752, 481);
            this.tabCsv.TabIndex = 1;
            this.tabCsv.Text = "CSVの作成";
            this.tabCsv.UseVisualStyleBackColor = true;
            // 
            // buttonOpenCsvDirectory
            // 
            this.buttonOpenCsvDirectory.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonOpenCsvDirectory.Location = new System.Drawing.Point(6, 452);
            this.buttonOpenCsvDirectory.Name = "buttonOpenCsvDirectory";
            this.buttonOpenCsvDirectory.Size = new System.Drawing.Size(167, 23);
            this.buttonOpenCsvDirectory.TabIndex = 2;
            this.buttonOpenCsvDirectory.Text = "CSVファイルのフォルダを開く";
            this.buttonOpenCsvDirectory.UseVisualStyleBackColor = true;
            this.buttonOpenCsvDirectory.Click += new System.EventHandler(this.buttonOpenCSVDirectory_Click);
            // 
            // checkBoxExecCsv
            // 
            this.checkBoxExecCsv.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBoxExecCsv.AutoSize = true;
            this.checkBoxExecCsv.Location = new System.Drawing.Point(348, 456);
            this.checkBoxExecCsv.Name = "checkBoxExecCsv";
            this.checkBoxExecCsv.Size = new System.Drawing.Size(172, 16);
            this.checkBoxExecCsv.TabIndex = 3;
            this.checkBoxExecCsv.Text = "CSV作成後関連付け起動する";
            this.checkBoxExecCsv.UseVisualStyleBackColor = true;
            // 
            // groupBoxWlIdAndAlias
            // 
            this.groupBoxWlIdAndAlias.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxWlIdAndAlias.Controls.Add(this.buttonMoveDownInWlIdsAndAliases);
            this.groupBoxWlIdAndAlias.Controls.Add(this.buttonMoveUpInWlIdsAndAliases);
            this.groupBoxWlIdAndAlias.Controls.Add(this.buttonRemoveFromWlIdsAndAliases);
            this.groupBoxWlIdAndAlias.Controls.Add(this.labelWlIdsAndAliases);
            this.groupBoxWlIdAndAlias.Controls.Add(this.labelSelectedAlias);
            this.groupBoxWlIdAndAlias.Controls.Add(this.labelSelectedWlId);
            this.groupBoxWlIdAndAlias.Controls.Add(this.buttonUseForCsv);
            this.groupBoxWlIdAndAlias.Controls.Add(this.listViewWlIdsAndAliases);
            this.groupBoxWlIdAndAlias.Controls.Add(this.listViewNewWlIds);
            this.groupBoxWlIdAndAlias.Controls.Add(this.buttonUpdateAlias);
            this.groupBoxWlIdAndAlias.Controls.Add(this.labelWlId);
            this.groupBoxWlIdAndAlias.Controls.Add(this.textAlias);
            this.groupBoxWlIdAndAlias.Controls.Add(this.buttonLoadNewWlIdsFromDb);
            this.groupBoxWlIdAndAlias.Location = new System.Drawing.Point(6, 6);
            this.groupBoxWlIdAndAlias.Name = "groupBoxWlIdAndAlias";
            this.groupBoxWlIdAndAlias.Size = new System.Drawing.Size(740, 388);
            this.groupBoxWlIdAndAlias.TabIndex = 0;
            this.groupBoxWlIdAndAlias.TabStop = false;
            this.groupBoxWlIdAndAlias.Text = "WLIDとエイリアス名";
            // 
            // buttonMoveDownInWlIdsAndAliases
            // 
            this.buttonMoveDownInWlIdsAndAliases.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonMoveDownInWlIdsAndAliases.Enabled = false;
            this.buttonMoveDownInWlIdsAndAliases.Location = new System.Drawing.Point(643, 359);
            this.buttonMoveDownInWlIdsAndAliases.Name = "buttonMoveDownInWlIdsAndAliases";
            this.buttonMoveDownInWlIdsAndAliases.Size = new System.Drawing.Size(75, 23);
            this.buttonMoveDownInWlIdsAndAliases.TabIndex = 12;
            this.buttonMoveDownInWlIdsAndAliases.Text = "下へ";
            this.buttonMoveDownInWlIdsAndAliases.UseVisualStyleBackColor = true;
            this.buttonMoveDownInWlIdsAndAliases.Click += new System.EventHandler(this.buttonMoveDownInWlIdsAndAliases_Click);
            // 
            // buttonMoveUpInWlIdsAndAliases
            // 
            this.buttonMoveUpInWlIdsAndAliases.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonMoveUpInWlIdsAndAliases.Enabled = false;
            this.buttonMoveUpInWlIdsAndAliases.Location = new System.Drawing.Point(562, 359);
            this.buttonMoveUpInWlIdsAndAliases.Name = "buttonMoveUpInWlIdsAndAliases";
            this.buttonMoveUpInWlIdsAndAliases.Size = new System.Drawing.Size(75, 23);
            this.buttonMoveUpInWlIdsAndAliases.TabIndex = 11;
            this.buttonMoveUpInWlIdsAndAliases.Text = "上へ";
            this.buttonMoveUpInWlIdsAndAliases.UseVisualStyleBackColor = true;
            this.buttonMoveUpInWlIdsAndAliases.Click += new System.EventHandler(this.buttonMoveUpInWlIdsAndAliases_Click);
            // 
            // buttonRemoveFromWlIdsAndAliases
            // 
            this.buttonRemoveFromWlIdsAndAliases.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonRemoveFromWlIdsAndAliases.Enabled = false;
            this.buttonRemoveFromWlIdsAndAliases.Location = new System.Drawing.Point(403, 359);
            this.buttonRemoveFromWlIdsAndAliases.Name = "buttonRemoveFromWlIdsAndAliases";
            this.buttonRemoveFromWlIdsAndAliases.Size = new System.Drawing.Size(153, 23);
            this.buttonRemoveFromWlIdsAndAliases.TabIndex = 10;
            this.buttonRemoveFromWlIdsAndAliases.Text = "選択したWLIDを削除する";
            this.buttonRemoveFromWlIdsAndAliases.UseVisualStyleBackColor = true;
            this.buttonRemoveFromWlIdsAndAliases.Click += new System.EventHandler(this.buttonRemoveFromWlIdsAndAliases_Click);
            // 
            // labelWlIdsAndAliases
            // 
            this.labelWlIdsAndAliases.AutoSize = true;
            this.labelWlIdsAndAliases.Location = new System.Drawing.Point(4, 176);
            this.labelWlIdsAndAliases.Name = "labelWlIdsAndAliases";
            this.labelWlIdsAndAliases.Size = new System.Drawing.Size(171, 12);
            this.labelWlIdsAndAliases.TabIndex = 3;
            this.labelWlIdsAndAliases.Text = "CSVに使用するWLIDとエイリアス名:";
            // 
            // labelSelectedAlias
            // 
            this.labelSelectedAlias.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelSelectedAlias.AutoSize = true;
            this.labelSelectedAlias.Location = new System.Drawing.Point(4, 364);
            this.labelSelectedAlias.Name = "labelSelectedAlias";
            this.labelSelectedAlias.Size = new System.Drawing.Size(62, 12);
            this.labelSelectedAlias.TabIndex = 7;
            this.labelSelectedAlias.Text = "エイリアス名:";
            // 
            // labelSelectedWlId
            // 
            this.labelSelectedWlId.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelSelectedWlId.AutoSize = true;
            this.labelSelectedWlId.Location = new System.Drawing.Point(70, 343);
            this.labelSelectedWlId.Name = "labelSelectedWlId";
            this.labelSelectedWlId.Size = new System.Drawing.Size(0, 12);
            this.labelSelectedWlId.TabIndex = 6;
            // 
            // buttonUseForCsv
            // 
            this.buttonUseForCsv.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonUseForCsv.Enabled = false;
            this.buttonUseForCsv.Location = new System.Drawing.Point(603, 150);
            this.buttonUseForCsv.Name = "buttonUseForCsv";
            this.buttonUseForCsv.Size = new System.Drawing.Size(126, 23);
            this.buttonUseForCsv.TabIndex = 2;
            this.buttonUseForCsv.Text = "CSVに使用する";
            this.buttonUseForCsv.UseVisualStyleBackColor = true;
            this.buttonUseForCsv.Click += new System.EventHandler(this.buttonUseForCsv_Click);
            // 
            // listViewWlIdsAndAliases
            // 
            this.listViewWlIdsAndAliases.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewWlIdsAndAliases.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderWlIds,
            this.columnHeaderAliases});
            this.listViewWlIdsAndAliases.FullRowSelect = true;
            this.listViewWlIdsAndAliases.HideSelection = false;
            this.listViewWlIdsAndAliases.Location = new System.Drawing.Point(6, 191);
            this.listViewWlIdsAndAliases.MultiSelect = false;
            this.listViewWlIdsAndAliases.Name = "listViewWlIdsAndAliases";
            this.listViewWlIdsAndAliases.Size = new System.Drawing.Size(728, 149);
            this.listViewWlIdsAndAliases.TabIndex = 4;
            this.listViewWlIdsAndAliases.UseCompatibleStateImageBehavior = false;
            this.listViewWlIdsAndAliases.View = System.Windows.Forms.View.Details;
            this.listViewWlIdsAndAliases.SelectedIndexChanged += new System.EventHandler(this.listViewWlIdsAndAliases_SelectedIndexChanged);
            // 
            // columnHeaderWlIds
            // 
            this.columnHeaderWlIds.Text = "WLID";
            this.columnHeaderWlIds.Width = 379;
            // 
            // columnHeaderAliases
            // 
            this.columnHeaderAliases.Text = "エイリアス名";
            this.columnHeaderAliases.Width = 288;
            // 
            // listViewNewWlIds
            // 
            this.listViewNewWlIds.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listViewNewWlIds.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeaderWlId});
            this.listViewNewWlIds.FullRowSelect = true;
            this.listViewNewWlIds.HideSelection = false;
            this.listViewNewWlIds.Location = new System.Drawing.Point(6, 47);
            this.listViewNewWlIds.Name = "listViewNewWlIds";
            this.listViewNewWlIds.Size = new System.Drawing.Size(728, 97);
            this.listViewNewWlIds.TabIndex = 1;
            this.listViewNewWlIds.UseCompatibleStateImageBehavior = false;
            this.listViewNewWlIds.View = System.Windows.Forms.View.Details;
            this.listViewNewWlIds.SelectedIndexChanged += new System.EventHandler(this.listViewNewWlIds_SelectedIndexChanged);
            // 
            // columnHeaderWlId
            // 
            this.columnHeaderWlId.Text = "WLID";
            this.columnHeaderWlId.Width = 531;
            // 
            // buttonUpdateAlias
            // 
            this.buttonUpdateAlias.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonUpdateAlias.Enabled = false;
            this.buttonUpdateAlias.Location = new System.Drawing.Point(270, 359);
            this.buttonUpdateAlias.Name = "buttonUpdateAlias";
            this.buttonUpdateAlias.Size = new System.Drawing.Size(127, 23);
            this.buttonUpdateAlias.TabIndex = 9;
            this.buttonUpdateAlias.Text = "エイリアス名の更新";
            this.buttonUpdateAlias.UseVisualStyleBackColor = true;
            this.buttonUpdateAlias.Click += new System.EventHandler(this.buttonUpdateAlias_Click);
            // 
            // labelWlId
            // 
            this.labelWlId.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelWlId.AutoSize = true;
            this.labelWlId.Location = new System.Drawing.Point(6, 343);
            this.labelWlId.Name = "labelWlId";
            this.labelWlId.Size = new System.Drawing.Size(33, 12);
            this.labelWlId.TabIndex = 5;
            this.labelWlId.Text = "WLID:";
            // 
            // textAlias
            // 
            this.textAlias.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.textAlias.Enabled = false;
            this.textAlias.Location = new System.Drawing.Point(72, 361);
            this.textAlias.MaxLength = 20;
            this.textAlias.Name = "textAlias";
            this.textAlias.Size = new System.Drawing.Size(192, 19);
            this.textAlias.TabIndex = 8;
            // 
            // buttonLoadNewWlIdsFromDb
            // 
            this.buttonLoadNewWlIdsFromDb.Location = new System.Drawing.Point(6, 18);
            this.buttonLoadNewWlIdsFromDb.Name = "buttonLoadNewWlIdsFromDb";
            this.buttonLoadNewWlIdsFromDb.Size = new System.Drawing.Size(239, 23);
            this.buttonLoadNewWlIdsFromDb.TabIndex = 0;
            this.buttonLoadNewWlIdsFromDb.Text = "新たなWLIDを内部データベースから読み込む";
            this.buttonLoadNewWlIdsFromDb.UseVisualStyleBackColor = true;
            this.buttonLoadNewWlIdsFromDb.Click += new System.EventHandler(this.buttonLoadNewWlIdsFromDb_Click);
            // 
            // groupBoxTargetDateRange
            // 
            this.groupBoxTargetDateRange.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxTargetDateRange.Controls.Add(this.dateTimePickerEndDate);
            this.groupBoxTargetDateRange.Controls.Add(this.dateTimePickerStartDate);
            this.groupBoxTargetDateRange.Controls.Add(this.labelStartDate);
            this.groupBoxTargetDateRange.Controls.Add(this.labelEndDate);
            this.groupBoxTargetDateRange.Location = new System.Drawing.Point(6, 400);
            this.groupBoxTargetDateRange.Name = "groupBoxTargetDateRange";
            this.groupBoxTargetDateRange.Size = new System.Drawing.Size(740, 46);
            this.groupBoxTargetDateRange.TabIndex = 1;
            this.groupBoxTargetDateRange.TabStop = false;
            this.groupBoxTargetDateRange.Text = "対象期間";
            // 
            // dateTimePickerEndDate
            // 
            this.dateTimePickerEndDate.Location = new System.Drawing.Point(240, 15);
            this.dateTimePickerEndDate.Name = "dateTimePickerEndDate";
            this.dateTimePickerEndDate.Size = new System.Drawing.Size(130, 19);
            this.dateTimePickerEndDate.TabIndex = 3;
            this.dateTimePickerEndDate.ValueChanged += new System.EventHandler(this.dateTimePickerEndDate_ValueChanged);
            // 
            // dateTimePickerStartDate
            // 
            this.dateTimePickerStartDate.Location = new System.Drawing.Point(55, 15);
            this.dateTimePickerStartDate.Name = "dateTimePickerStartDate";
            this.dateTimePickerStartDate.Size = new System.Drawing.Size(130, 19);
            this.dateTimePickerStartDate.TabIndex = 1;
            this.dateTimePickerStartDate.ValueChanged += new System.EventHandler(this.dateTimePickerStartDate_ValueChanged);
            // 
            // labelStartDate
            // 
            this.labelStartDate.AutoSize = true;
            this.labelStartDate.Location = new System.Drawing.Point(6, 18);
            this.labelStartDate.Name = "labelStartDate";
            this.labelStartDate.Size = new System.Drawing.Size(43, 12);
            this.labelStartDate.TabIndex = 0;
            this.labelStartDate.Text = "開始日:";
            // 
            // labelEndDate
            // 
            this.labelEndDate.AutoSize = true;
            this.labelEndDate.Location = new System.Drawing.Point(191, 18);
            this.labelEndDate.Name = "labelEndDate";
            this.labelEndDate.Size = new System.Drawing.Size(43, 12);
            this.labelEndDate.TabIndex = 2;
            this.labelEndDate.Text = "終了日:";
            // 
            // buttonCreateCsvFromDb
            // 
            this.buttonCreateCsvFromDb.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCreateCsvFromDb.Location = new System.Drawing.Point(526, 452);
            this.buttonCreateCsvFromDb.Name = "buttonCreateCsvFromDb";
            this.buttonCreateCsvFromDb.Size = new System.Drawing.Size(220, 23);
            this.buttonCreateCsvFromDb.TabIndex = 4;
            this.buttonCreateCsvFromDb.Text = "内部データベースからCSVを作成";
            this.buttonCreateCsvFromDb.UseVisualStyleBackColor = true;
            this.buttonCreateCsvFromDb.Click += new System.EventHandler(this.buttonCreateCsvFromDb_Click);
            // 
            // textBoxConsole
            // 
            this.textBoxConsole.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxConsole.Location = new System.Drawing.Point(12, 525);
            this.textBoxConsole.MaxLength = 0;
            this.textBoxConsole.Multiline = true;
            this.textBoxConsole.Name = "textBoxConsole";
            this.textBoxConsole.ReadOnly = true;
            this.textBoxConsole.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxConsole.Size = new System.Drawing.Size(760, 129);
            this.textBoxConsole.TabIndex = 1;
            this.textBoxConsole.WordWrap = false;
            // 
            // buttonClearConsole
            // 
            this.buttonClearConsole.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonClearConsole.Location = new System.Drawing.Point(12, 660);
            this.buttonClearConsole.Name = "buttonClearConsole";
            this.buttonClearConsole.Size = new System.Drawing.Size(119, 23);
            this.buttonClearConsole.TabIndex = 2;
            this.buttonClearConsole.Text = "コンソールをクリア";
            this.buttonClearConsole.UseVisualStyleBackColor = true;
            this.buttonClearConsole.Click += new System.EventHandler(this.buttonClearConsole_Click);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 691);
            this.Controls.Add(this.buttonClearConsole);
            this.Controls.Add(this.textBoxConsole);
            this.Controls.Add(this.tabControlMain);
            this.MaximizeBox = false;
            this.MinimumSize = new System.Drawing.Size(800, 730);
            this.Name = "FormMain";
            this.Text = "WATCH LOGGERの .log 整理ツール";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormMain_FormClosed);
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.groupBoxLogFiles.ResumeLayout(false);
            this.tabControlMain.ResumeLayout(false);
            this.tabLog.ResumeLayout(false);
            this.tabCsv.ResumeLayout(false);
            this.tabCsv.PerformLayout();
            this.groupBoxWlIdAndAlias.ResumeLayout(false);
            this.groupBoxWlIdAndAlias.PerformLayout();
            this.groupBoxTargetDateRange.ResumeLayout(false);
            this.groupBoxTargetDateRange.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListView listViewLogFiles;
        private System.Windows.Forms.GroupBox groupBoxLogFiles;
        private System.Windows.Forms.Button buttonLoadLogFilesFromLogDirectory;
        private System.Windows.Forms.Button buttonClearLogFiles;
        private System.Windows.Forms.Button buttonLogFilesToDb;
        private System.Windows.Forms.TabControl tabControlMain;
        private System.Windows.Forms.TabPage tabLog;
        private System.Windows.Forms.TabPage tabCsv;
        private System.Windows.Forms.GroupBox groupBoxTargetDateRange;
        private System.Windows.Forms.Label labelEndDate;
        private System.Windows.Forms.Label labelStartDate;
        private System.Windows.Forms.Button buttonCreateCsvFromDb;
        private System.Windows.Forms.DateTimePicker dateTimePickerEndDate;
        private System.Windows.Forms.DateTimePicker dateTimePickerStartDate;
        private System.Windows.Forms.ColumnHeader columnHeaderLogFiles;
        private System.Windows.Forms.Button buttonOpenLogDirectory;
        private System.Windows.Forms.Button buttonOpenBackupDirectory;
        private System.Windows.Forms.GroupBox groupBoxWlIdAndAlias;
        private System.Windows.Forms.Button buttonUpdateAlias;
        private System.Windows.Forms.Label labelWlId;
        private System.Windows.Forms.TextBox textAlias;
        private System.Windows.Forms.Button buttonLoadNewWlIdsFromDb;
        private System.Windows.Forms.Button buttonUseForCsv;
        private System.Windows.Forms.ListView listViewWlIdsAndAliases;
        private System.Windows.Forms.ListView listViewNewWlIds;
        private System.Windows.Forms.Label labelWlIdsAndAliases;
        private System.Windows.Forms.Label labelSelectedAlias;
        private System.Windows.Forms.Label labelSelectedWlId;
        private System.Windows.Forms.Button buttonRemoveFromWlIdsAndAliases;
        private System.Windows.Forms.ColumnHeader columnHeaderWlId;
        private System.Windows.Forms.Button buttonMoveDownInWlIdsAndAliases;
        private System.Windows.Forms.Button buttonMoveUpInWlIdsAndAliases;
        private System.Windows.Forms.ColumnHeader columnHeaderWlIds;
        private System.Windows.Forms.ColumnHeader columnHeaderAliases;
        private System.Windows.Forms.Button buttonOpenCsvDirectory;
        private System.Windows.Forms.CheckBox checkBoxExecCsv;
        private System.Windows.Forms.TextBox textBoxConsole;
        private System.Windows.Forms.Button buttonClearConsole;
    }
}

