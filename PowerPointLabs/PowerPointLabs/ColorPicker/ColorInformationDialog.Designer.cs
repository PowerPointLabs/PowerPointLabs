namespace PowerPointLabs.ColorPicker
{
    partial class ColorInformationDialog
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea3 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Series series3 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint7 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(0D, 10D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint8 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(0D, 98.5D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint9 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(0D, 55.6D);
            this.selectedColorPanel = new System.Windows.Forms.Panel();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.hexTextBox = new System.Windows.Forms.TextBox();
            this.rgbTextBox = new System.Windows.Forms.TextBox();
            this.HSLTextBox = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.SuspendLayout();
            // 
            // selectedColorPanel
            // 
            this.selectedColorPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.selectedColorPanel.Location = new System.Drawing.Point(12, 12);
            this.selectedColorPanel.Name = "selectedColorPanel";
            this.selectedColorPanel.Size = new System.Drawing.Size(219, 126);
            this.selectedColorPanel.TabIndex = 0;
            // 
            // chart1
            // 
            chartArea3.AxisX.Enabled = System.Windows.Forms.DataVisualization.Charting.AxisEnabled.True;
            chartArea3.AxisX.MajorGrid.Enabled = false;
            chartArea3.AxisX2.Enabled = System.Windows.Forms.DataVisualization.Charting.AxisEnabled.False;
            chartArea3.AxisY.Enabled = System.Windows.Forms.DataVisualization.Charting.AxisEnabled.False;
            chartArea3.AxisY.Maximum = 100D;
            chartArea3.AxisY.Minimum = 0D;
            chartArea3.AxisY2.Enabled = System.Windows.Forms.DataVisualization.Charting.AxisEnabled.False;
            chartArea3.BorderColor = System.Drawing.Color.Transparent;
            chartArea3.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea3);
            this.chart1.Location = new System.Drawing.Point(246, 11);
            this.chart1.Name = "chart1";
            this.chart1.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.None;
            series3.ChartArea = "ChartArea1";
            series3.Name = "Series1";
            dataPoint7.AxisLabel = "R";
            dataPoint7.Color = System.Drawing.Color.Red;
            dataPoint7.CustomProperties = "DrawingStyle=Emboss, LabelStyle=Top";
            dataPoint7.Font = new System.Drawing.Font("Century Schoolbook", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataPoint7.IsValueShownAsLabel = true;
            dataPoint7.Label = "10%";
            dataPoint7.MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.None;
            dataPoint8.AxisLabel = "G";
            dataPoint8.Color = System.Drawing.Color.Green;
            dataPoint8.CustomProperties = "DrawingStyle=Emboss, LabelStyle=Top";
            dataPoint8.Font = new System.Drawing.Font("Century Schoolbook", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataPoint8.IsValueShownAsLabel = true;
            dataPoint8.Label = "98.5%";
            dataPoint8.LabelFormat = "";
            dataPoint9.AxisLabel = "B";
            dataPoint9.Color = System.Drawing.Color.Blue;
            dataPoint9.CustomProperties = "DrawingStyle=Emboss";
            dataPoint9.Font = new System.Drawing.Font("Century Schoolbook", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataPoint9.Label = "55.6%";
            series3.Points.Add(dataPoint7);
            series3.Points.Add(dataPoint8);
            series3.Points.Add(dataPoint9);
            this.chart1.Series.Add(series3);
            this.chart1.Size = new System.Drawing.Size(282, 222);
            this.chart1.TabIndex = 1;
            this.chart1.Text = "chart1";
            // 
            // hexTextBox
            // 
            this.hexTextBox.BackColor = System.Drawing.SystemColors.ControlLight;
            this.hexTextBox.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.hexTextBox.Location = new System.Drawing.Point(12, 144);
            this.hexTextBox.Name = "hexTextBox";
            this.hexTextBox.ReadOnly = true;
            this.hexTextBox.Size = new System.Drawing.Size(219, 27);
            this.hexTextBox.TabIndex = 3;
            // 
            // rgbTextBox
            // 
            this.rgbTextBox.BackColor = System.Drawing.SystemColors.ControlLight;
            this.rgbTextBox.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rgbTextBox.Location = new System.Drawing.Point(12, 176);
            this.rgbTextBox.Name = "rgbTextBox";
            this.rgbTextBox.ReadOnly = true;
            this.rgbTextBox.Size = new System.Drawing.Size(219, 27);
            this.rgbTextBox.TabIndex = 4;
            // 
            // HSLTextBox
            // 
            this.HSLTextBox.BackColor = System.Drawing.SystemColors.ControlLight;
            this.HSLTextBox.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.HSLTextBox.Location = new System.Drawing.Point(12, 208);
            this.HSLTextBox.Name = "HSLTextBox";
            this.HSLTextBox.ReadOnly = true;
            this.HSLTextBox.Size = new System.Drawing.Size(219, 27);
            this.HSLTextBox.TabIndex = 5;
            // 
            // ColorInformationDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(540, 241);
            this.Controls.Add(this.HSLTextBox);
            this.Controls.Add(this.rgbTextBox);
            this.Controls.Add(this.hexTextBox);
            this.Controls.Add(this.chart1);
            this.Controls.Add(this.selectedColorPanel);
            this.Name = "ColorInformationDialog";
            this.Text = "Color Information Dialog";
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel selectedColorPanel;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        private System.Windows.Forms.TextBox hexTextBox;
        private System.Windows.Forms.TextBox rgbTextBox;
        private System.Windows.Forms.TextBox HSLTextBox;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}