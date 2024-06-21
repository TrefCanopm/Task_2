namespace Task_2
{
    partial class MaynForm
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
            SearchFile = new FolderBrowserDialog();
            LoadSearch = new Button();
            SearchSave = new Button();
            LoadFile = new TextBox();
            Start = new Button();
            LoadLabel = new Label();
            SaveFile = new TextBox();
            labelSave = new Label();
            SuspendLayout();
            // 
            // LoadSearch
            // 
            LoadSearch.Location = new Point(308, 30);
            LoadSearch.Name = "LoadSearch";
            LoadSearch.Size = new Size(94, 29);
            LoadSearch.TabIndex = 0;
            LoadSearch.Text = "Поиск";
            LoadSearch.UseVisualStyleBackColor = true;
            LoadSearch.Click += SearchLoad_Click;
            // 
            // SearchSave
            // 
            SearchSave.Location = new Point(308, 85);
            SearchSave.Name = "SearchSave";
            SearchSave.Size = new Size(94, 29);
            SearchSave.TabIndex = 1;
            SearchSave.Text = "Поиск";
            SearchSave.UseVisualStyleBackColor = true;
            SearchSave.Click += SearchSave_Click;
            // 
            // LoadFile
            // 
            LoadFile.Location = new Point(12, 32);
            LoadFile.Name = "LoadFile";
            LoadFile.Size = new Size(290, 27);
            LoadFile.TabIndex = 2;
            // 
            // Start
            // 
            Start.Location = new Point(12, 120);
            Start.Name = "Start";
            Start.Size = new Size(94, 29);
            Start.TabIndex = 3;
            Start.Text = "Запуск";
            Start.UseVisualStyleBackColor = true;
            Start.Click += Start_Click;
            // 
            // LoadLabel
            // 
            LoadLabel.AutoSize = true;
            LoadLabel.Location = new Point(12, 9);
            LoadLabel.Name = "LoadLabel";
            LoadLabel.Size = new Size(149, 20);
            LoadLabel.TabIndex = 4;
            LoadLabel.Text = "Укажите файл Отчет";
            // 
            // SaveFile
            // 
            SaveFile.Location = new Point(12, 87);
            SaveFile.Name = "SaveFile";
            SaveFile.Size = new Size(290, 27);
            SaveFile.TabIndex = 5;
            // 
            // labelSave
            // 
            labelSave.AutoSize = true;
            labelSave.Location = new Point(12, 64);
            labelSave.Name = "labelSave";
            labelSave.Size = new Size(226, 20);
            labelSave.TabIndex = 6;
            labelSave.Text = "Укажите папку для сохранения";
            // 
            // MaynForm
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(436, 168);
            Controls.Add(labelSave);
            Controls.Add(SaveFile);
            Controls.Add(LoadLabel);
            Controls.Add(Start);
            Controls.Add(LoadFile);
            Controls.Add(SearchSave);
            Controls.Add(LoadSearch);
            Name = "MaynForm";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private FolderBrowserDialog SearchFile;
        private Button LoadSearch;
        private Button SearchSave;
        private TextBox LoadFile;
        private Button Start;
        private Label LoadLabel;
        private TextBox SaveFile;
        private Label labelSave;
    }
}
