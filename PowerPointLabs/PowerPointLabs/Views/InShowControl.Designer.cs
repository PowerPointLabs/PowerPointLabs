namespace PowerPointLabs.Views
{
    partial class InShowControl
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
            this.recButton = new System.Windows.Forms.Button();
            this.undoButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // recButton
            // 
            this.recButton.Font = new System.Drawing.Font("Arial Narrow", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.recButton.Location = new System.Drawing.Point(12, 12);
            this.recButton.Name = "recButton";
            this.recButton.Size = new System.Drawing.Size(162, 67);
            this.recButton.TabIndex = 0;
            this.recButton.Text = "Start Recording";
            this.recButton.UseVisualStyleBackColor = true;
            this.recButton.Click += new System.EventHandler(this.RecButtonClick);
            // 
            // undoButton
            // 
            this.undoButton.Font = new System.Drawing.Font("Arial Narrow", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.undoButton.Location = new System.Drawing.Point(180, 12);
            this.undoButton.Name = "undoButton";
            this.undoButton.Size = new System.Drawing.Size(70, 67);
            this.undoButton.TabIndex = 1;
            this.undoButton.Text = "Undo";
            this.undoButton.UseVisualStyleBackColor = true;
            this.undoButton.Click += new System.EventHandler(this.UndoButtonClick);
            // 
            // InShowControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(262, 93);
            this.Controls.Add(this.undoButton);
            this.Controls.Add(this.recButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "InShowControl";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "InShowControl";
            this.TopMost = true;
            this.MouseClick += new System.Windows.Forms.MouseEventHandler(this.InShowControlMouseClick);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button recButton;
        private System.Windows.Forms.Button undoButton;
    }
}