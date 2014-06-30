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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint1 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(0D, 10D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint2 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(0D, 98.5D);
            System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint3 = new System.Windows.Forms.DataVisualization.Charting.DataPoint(0D, 55.6D);
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
            chartArea1.AxisX.Enabled = System.Windows.Forms.DataVisualization.Charting.AxisEnabled.True;
            chartArea1.AxisX.MajorGrid.Enabled = false;
            chartArea1.AxisX2.Enabled = System.Windows.Forms.DataVisualization.Charting.AxisEnabled.False;
            chartArea1.AxisY.Enabled = System.Windows.Forms.DataVisualization.Charting.AxisEnabled.False;
            chartArea1.AxisY.Maximum = 100D;
            chartArea1.AxisY.Minimum = 0D;
            chartArea1.AxisY2.Enabled = System.Windows.Forms.DataVisualization.Charting.AxisEnabled.False;
            chartArea1.BorderColor = System.Drawing.Color.Transparent;
            chartArea1.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea1);
            this.chart1.Location = new System.Drawing.Point(246, 11);
            this.chart1.Name = "chart1";
            this.chart1.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.None;
            series1.ChartArea = "ChartArea1";
            series1.Name = "Series1";
            dataPoint1.AxisLabel = "R";
            dataPoint1.Color = System.Drawing.Color.Red;
            dataPoint1.CustomProperties = "DrawingStyle=Emboss, LabelStyle=Top";
            dataPoint1.Font = new System.Drawing.Font("Century Schoolbook", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataPoint1.IsValueShownAsLabel = true;
            dataPoint1.Label = "10%";
            dataPoint1.MarkerStyle = System.Windows.Forms.DataVisualization.Charting.MarkerStyle.None;
            dataPoint2.AxisLabel = "G";
            dataPoint2.Color = System.Drawing.Color.Green;
            dataPoint2.CustomProperties = "DrawingStyle=Emboss, LabelStyle=Top";
            dataPoint2.Font = new System.Drawing.Font("Century Schoolbook", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataPoint2.IsValueShownAsLabel = true;
            dataPoint2.Label = "98.5%";
            dataPoint2.LabelFormat = "";
            dataPoint3.AxisLabel = "B";
            dataPoint3.Color = System.Drawing.Color.Blue;
            dataPoint3.CustomProperties = "DrawingStyle=Emboss";
            dataPoint3.Font = new System.Drawing.Font("Century Schoolbook", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataPoint3.Label = "55.6%";
            series1.Points.Add(dataPoint1);
            series1.Points.Add(dataPoint2);
            series1.Points.Add(dataPoint3);
            this.chart1.Series.Add(series1);
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
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ColorInformationDialog";
            this.Text = "Color Information Dialog";
            this.TopMost = true;
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