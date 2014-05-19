namespace PowerPointLabs.Views
{
    partial class RecorderUI
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.statusLabel = new System.Windows.Forms.Label();
            this.recButton = new System.Windows.Forms.Button();
            this.stopButton = new System.Windows.Forms.Button();
            this.playButton = new System.Windows.Forms.Button();
            this.timerLabel = new System.Windows.Forms.Label();
            this.soundTrackBar = new System.Windows.Forms.TrackBar();
            this.recDisplay = new System.Windows.Forms.ListView();
            this.recNumber = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.recName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.recLength = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.recDate = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            ((System.ComponentModel.ISupportInitialize)(this.soundTrackBar)).BeginInit();
            this.SuspendLayout();
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.statusLabel.ForeColor = System.Drawing.Color.Red;
            this.statusLabel.Location = new System.Drawing.Point(12, 80);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(89, 12);
            this.statusLabel.TabIndex = 0;
            this.statusLabel.Text = "Recording...";
            // 
            // recButton
            // 
            this.recButton.Location = new System.Drawing.Point(12, 22);
            this.recButton.Name = "recButton";
            this.recButton.Size = new System.Drawing.Size(75, 41);
            this.recButton.TabIndex = 1;
            this.recButton.Text = "Record";
            this.recButton.UseVisualStyleBackColor = true;
            this.recButton.Click += new System.EventHandler(this.RecButtonClick);
            // 
            // stopButton
            // 
            this.stopButton.Location = new System.Drawing.Point(103, 22);
            this.stopButton.Name = "stopButton";
            this.stopButton.Size = new System.Drawing.Size(75, 41);
            this.stopButton.TabIndex = 2;
            this.stopButton.Text = "Stop";
            this.stopButton.UseVisualStyleBackColor = true;
            this.stopButton.Click += new System.EventHandler(this.StopButtonClick);
            // 
            // playButton
            // 
            this.playButton.Location = new System.Drawing.Point(195, 22);
            this.playButton.Name = "playButton";
            this.playButton.Size = new System.Drawing.Size(75, 41);
            this.playButton.TabIndex = 3;
            this.playButton.Text = "Play";
            this.playButton.UseVisualStyleBackColor = true;
            this.playButton.Click += new System.EventHandler(this.PlayButtonClick);
            // 
            // timerLabel
            // 
            this.timerLabel.AutoSize = true;
            this.timerLabel.Font = new System.Drawing.Font("Times New Roman", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.timerLabel.Location = new System.Drawing.Point(151, 69);
            this.timerLabel.Name = "timerLabel";
            this.timerLabel.Size = new System.Drawing.Size(116, 31);
            this.timerLabel.TabIndex = 4;
            this.timerLabel.Text = "00:00:00";
            // 
            // soundTrackBar
            // 
            this.soundTrackBar.Location = new System.Drawing.Point(9, 99);
            this.soundTrackBar.Maximum = 100;
            this.soundTrackBar.Name = "soundTrackBar";
            this.soundTrackBar.Size = new System.Drawing.Size(258, 45);
            this.soundTrackBar.TabIndex = 5;
            // 
            // recDisplay
            // 
            this.recDisplay.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.recNumber,
            this.recName,
            this.recLength,
            this.recDate});
            this.recDisplay.FullRowSelect = true;
            this.recDisplay.HideSelection = false;
            this.recDisplay.Location = new System.Drawing.Point(14, 138);
            this.recDisplay.MultiSelect = false;
            this.recDisplay.Name = "recDisplay";
            this.recDisplay.Size = new System.Drawing.Size(253, 191);
            this.recDisplay.TabIndex = 6;
            this.recDisplay.UseCompatibleStateImageBehavior = false;
            this.recDisplay.View = System.Windows.Forms.View.Details;
            // 
            // recNumber
            // 
            this.recNumber.Text = "No.";
            this.recNumber.Width = 38;
            // 
            // recName
            // 
            this.recName.Text = "Name";
            this.recName.Width = 47;
            // 
            // recLength
            // 
            this.recLength.Text = "Length";
            this.recLength.Width = 68;
            // 
            // recDate
            // 
            this.recDate.Text = "Record Date";
            this.recDate.Width = 81;
            // 
            // RecorderUI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(280, 341);
            this.Controls.Add(this.recDisplay);
            this.Controls.Add(this.soundTrackBar);
            this.Controls.Add(this.timerLabel);
            this.Controls.Add(this.playButton);
            this.Controls.Add(this.stopButton);
            this.Controls.Add(this.recButton);
            this.Controls.Add(this.statusLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "RecorderUI";
            this.ShowInTaskbar = false;
            this.Text = "MCIRecorder";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.RecorderUI_FormClosing);
            this.Load += new System.EventHandler(this.RecorderUI_Load);
            ((System.ComponentModel.ISupportInitialize)(this.soundTrackBar)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.Button recButton;
        private System.Windows.Forms.Button stopButton;
        private System.Windows.Forms.Button playButton;
        private System.Windows.Forms.Label timerLabel;
        private System.Windows.Forms.TrackBar soundTrackBar;
        private System.Windows.Forms.ListView recDisplay;
        private System.Windows.Forms.ColumnHeader recNumber;
        private System.Windows.Forms.ColumnHeader recName;
        private System.Windows.Forms.ColumnHeader recLength;
        private System.Windows.Forms.ColumnHeader recDate;
    }
}