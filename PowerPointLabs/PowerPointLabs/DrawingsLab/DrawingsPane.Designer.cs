namespace PowerPointLabs
{
    partial class DrawingsPane
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.LineButton = new System.Windows.Forms.Button();
            this.RectButton = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.HideButton = new System.Windows.Forms.Button();
            this.CloneButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // LineButton
            // 
            this.LineButton.Image = global::PowerPointLabs.Properties.Resources.About;
            this.LineButton.Location = new System.Drawing.Point(22, 23);
            this.LineButton.Name = "LineButton";
            this.LineButton.Size = new System.Drawing.Size(44, 45);
            this.LineButton.TabIndex = 0;
            this.LineButton.UseVisualStyleBackColor = true;
            this.LineButton.Click += new System.EventHandler(this.LineButton_Click);
            // 
            // RectButton
            // 
            this.RectButton.Image = global::PowerPointLabs.Properties.Resources.About;
            this.RectButton.Location = new System.Drawing.Point(82, 23);
            this.RectButton.Name = "RectButton";
            this.RectButton.Size = new System.Drawing.Size(44, 45);
            this.RectButton.TabIndex = 1;
            this.RectButton.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Image = global::PowerPointLabs.Properties.Resources.About;
            this.button2.Location = new System.Drawing.Point(144, 23);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(44, 45);
            this.button2.TabIndex = 2;
            this.button2.UseVisualStyleBackColor = true;
            // 
            // HideButton
            // 
            this.HideButton.Image = global::PowerPointLabs.Properties.Resources.About;
            this.HideButton.Location = new System.Drawing.Point(144, 91);
            this.HideButton.Name = "HideButton";
            this.HideButton.Size = new System.Drawing.Size(44, 45);
            this.HideButton.TabIndex = 3;
            this.HideButton.UseVisualStyleBackColor = true;
            this.HideButton.Click += new System.EventHandler(this.HideButton_Click);
            // 
            // CloneButton
            // 
            this.CloneButton.Image = global::PowerPointLabs.Properties.Resources.About;
            this.CloneButton.Location = new System.Drawing.Point(208, 91);
            this.CloneButton.Name = "CloneButton";
            this.CloneButton.Size = new System.Drawing.Size(44, 45);
            this.CloneButton.TabIndex = 4;
            this.CloneButton.UseVisualStyleBackColor = true;
            this.CloneButton.Click += new System.EventHandler(this.CloneButton_Click);
            // 
            // DrawingsPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.CloneButton);
            this.Controls.Add(this.HideButton);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.RectButton);
            this.Controls.Add(this.LineButton);
            this.Name = "DrawingsPane";
            this.Size = new System.Drawing.Size(304, 883);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button LineButton;
        private System.Windows.Forms.Button RectButton;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button HideButton;
        private System.Windows.Forms.Button CloneButton;

    }
}
