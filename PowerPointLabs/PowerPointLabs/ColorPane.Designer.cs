namespace PowerPointLabs
{
    partial class ColorPane
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
            this.components = new System.ComponentModel.Container();
            this.panel1 = new System.Windows.Forms.Panel();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.AnalogousColorPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.flowLayoutPanel6 = new System.Windows.Forms.FlowLayoutPanel();
            this.AnalogousLighter = new System.Windows.Forms.Panel();
            this.AnalogousSelected = new System.Windows.Forms.Panel();
            this.AnalogousDarker = new System.Windows.Forms.Panel();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.flowLayoutPanel7 = new System.Windows.Forms.FlowLayoutPanel();
            this.ComplementaryLighter = new System.Windows.Forms.Panel();
            this.ComplementarySelected = new System.Windows.Forms.Panel();
            this.ComplementaryDarker = new System.Windows.Forms.Panel();
            this.flowLayoutPanel3 = new System.Windows.Forms.FlowLayoutPanel();
            this.label3 = new System.Windows.Forms.Label();
            this.flowLayoutPanel8 = new System.Windows.Forms.FlowLayoutPanel();
            this.TriadicLower = new System.Windows.Forms.Panel();
            this.TriadicSelected = new System.Windows.Forms.Panel();
            this.TriadicHigher = new System.Windows.Forms.Panel();
            this.flowLayoutPanel4 = new System.Windows.Forms.FlowLayoutPanel();
            this.label4 = new System.Windows.Forms.Label();
            this.flowLayoutPanel9 = new System.Windows.Forms.FlowLayoutPanel();
            this.Tetradic1 = new System.Windows.Forms.Panel();
            this.TetradicSelected = new System.Windows.Forms.Panel();
            this.Tetradic2 = new System.Windows.Forms.Panel();
            this.Tetradic3 = new System.Windows.Forms.Panel();
            this.brightnessBar = new System.Windows.Forms.TrackBar();
            this.brightnessPanel = new System.Windows.Forms.Panel();
            this.saturationPanel = new System.Windows.Forms.Panel();
            this.saturationBar = new System.Windows.Forms.TrackBar();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.label5 = new System.Windows.Forms.Label();
            this.flowLayoutPanel5 = new System.Windows.Forms.FlowLayoutPanel();
            this.MonoPanel1 = new System.Windows.Forms.Panel();
            this.MonoPanel2 = new System.Windows.Forms.Panel();
            this.MonoPanel3 = new System.Windows.Forms.Panel();
            this.MonoPanel4 = new System.Windows.Forms.Panel();
            this.MonoPanel5 = new System.Windows.Forms.Panel();
            this.MonoPanel6 = new System.Windows.Forms.Panel();
            this.MonoPanel7 = new System.Windows.Forms.Panel();
            this.AnalogousColorPanel.SuspendLayout();
            this.flowLayoutPanel6.SuspendLayout();
            this.flowLayoutPanel2.SuspendLayout();
            this.flowLayoutPanel7.SuspendLayout();
            this.flowLayoutPanel3.SuspendLayout();
            this.flowLayoutPanel8.SuspendLayout();
            this.flowLayoutPanel4.SuspendLayout();
            this.flowLayoutPanel9.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.brightnessBar)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.saturationBar)).BeginInit();
            this.flowLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel5.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.AllowDrop = true;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Location = new System.Drawing.Point(20, 16);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(75, 47);
            this.panel1.TabIndex = 0;
            this.panel1.DragDrop += new System.Windows.Forms.DragEventHandler(this.panel1_DragDrop);
            this.panel1.DragEnter += new System.Windows.Forms.DragEventHandler(this.panel1_DragEnter);
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // colorDialog1
            // 
            this.colorDialog1.FullOpen = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(101, 16);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Eyedropper";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(186, 16);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 2;
            this.button2.Text = "Edit Color";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // AnalogousColorPanel
            // 
            this.AnalogousColorPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.AnalogousColorPanel.Controls.Add(this.label1);
            this.AnalogousColorPanel.Controls.Add(this.flowLayoutPanel6);
            this.AnalogousColorPanel.Location = new System.Drawing.Point(20, 301);
            this.AnalogousColorPanel.Name = "AnalogousColorPanel";
            this.AnalogousColorPanel.Size = new System.Drawing.Size(263, 67);
            this.AnalogousColorPanel.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("MS Reference Sans Serif", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(181, 23);
            this.label1.TabIndex = 5;
            this.label1.Text = "Analogous Colors";
            // 
            // flowLayoutPanel6
            // 
            this.flowLayoutPanel6.Controls.Add(this.AnalogousLighter);
            this.flowLayoutPanel6.Controls.Add(this.AnalogousSelected);
            this.flowLayoutPanel6.Controls.Add(this.AnalogousDarker);
            this.flowLayoutPanel6.Location = new System.Drawing.Point(3, 26);
            this.flowLayoutPanel6.Name = "flowLayoutPanel6";
            this.flowLayoutPanel6.Size = new System.Drawing.Size(255, 26);
            this.flowLayoutPanel6.TabIndex = 12;
            // 
            // AnalogousLighter
            // 
            this.AnalogousLighter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.AnalogousLighter.Location = new System.Drawing.Point(3, 3);
            this.AnalogousLighter.Name = "AnalogousLighter";
            this.AnalogousLighter.Size = new System.Drawing.Size(20, 20);
            this.AnalogousLighter.TabIndex = 2;
            this.AnalogousLighter.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // AnalogousSelected
            // 
            this.AnalogousSelected.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.AnalogousSelected.Location = new System.Drawing.Point(29, 3);
            this.AnalogousSelected.Name = "AnalogousSelected";
            this.AnalogousSelected.Size = new System.Drawing.Size(20, 20);
            this.AnalogousSelected.TabIndex = 3;
            this.AnalogousSelected.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // AnalogousDarker
            // 
            this.AnalogousDarker.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.AnalogousDarker.Location = new System.Drawing.Point(55, 3);
            this.AnalogousDarker.Name = "AnalogousDarker";
            this.AnalogousDarker.Size = new System.Drawing.Size(20, 20);
            this.AnalogousDarker.TabIndex = 4;
            this.AnalogousDarker.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.flowLayoutPanel2.Controls.Add(this.label2);
            this.flowLayoutPanel2.Controls.Add(this.flowLayoutPanel7);
            this.flowLayoutPanel2.Location = new System.Drawing.Point(20, 374);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Size = new System.Drawing.Size(263, 67);
            this.flowLayoutPanel2.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("MS Reference Sans Serif", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(3, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(239, 23);
            this.label2.TabIndex = 6;
            this.label2.Text = "Complementary Colors";
            // 
            // flowLayoutPanel7
            // 
            this.flowLayoutPanel7.Controls.Add(this.ComplementaryLighter);
            this.flowLayoutPanel7.Controls.Add(this.ComplementarySelected);
            this.flowLayoutPanel7.Controls.Add(this.ComplementaryDarker);
            this.flowLayoutPanel7.Location = new System.Drawing.Point(3, 26);
            this.flowLayoutPanel7.Name = "flowLayoutPanel7";
            this.flowLayoutPanel7.Size = new System.Drawing.Size(255, 26);
            this.flowLayoutPanel7.TabIndex = 13;
            // 
            // ComplementaryLighter
            // 
            this.ComplementaryLighter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ComplementaryLighter.Location = new System.Drawing.Point(3, 3);
            this.ComplementaryLighter.Name = "ComplementaryLighter";
            this.ComplementaryLighter.Size = new System.Drawing.Size(20, 20);
            this.ComplementaryLighter.TabIndex = 2;
            this.ComplementaryLighter.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // ComplementarySelected
            // 
            this.ComplementarySelected.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ComplementarySelected.Location = new System.Drawing.Point(29, 3);
            this.ComplementarySelected.Name = "ComplementarySelected";
            this.ComplementarySelected.Size = new System.Drawing.Size(20, 20);
            this.ComplementarySelected.TabIndex = 3;
            this.ComplementarySelected.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // ComplementaryDarker
            // 
            this.ComplementaryDarker.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ComplementaryDarker.Location = new System.Drawing.Point(55, 3);
            this.ComplementaryDarker.Name = "ComplementaryDarker";
            this.ComplementaryDarker.Size = new System.Drawing.Size(20, 20);
            this.ComplementaryDarker.TabIndex = 4;
            this.ComplementaryDarker.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // flowLayoutPanel3
            // 
            this.flowLayoutPanel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.flowLayoutPanel3.Controls.Add(this.label3);
            this.flowLayoutPanel3.Controls.Add(this.flowLayoutPanel8);
            this.flowLayoutPanel3.Location = new System.Drawing.Point(20, 444);
            this.flowLayoutPanel3.Name = "flowLayoutPanel3";
            this.flowLayoutPanel3.Size = new System.Drawing.Size(263, 67);
            this.flowLayoutPanel3.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("MS Reference Sans Serif", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(3, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(147, 23);
            this.label3.TabIndex = 6;
            this.label3.Text = "Triadic Colors";
            // 
            // flowLayoutPanel8
            // 
            this.flowLayoutPanel8.Controls.Add(this.TriadicLower);
            this.flowLayoutPanel8.Controls.Add(this.TriadicSelected);
            this.flowLayoutPanel8.Controls.Add(this.TriadicHigher);
            this.flowLayoutPanel8.Location = new System.Drawing.Point(3, 26);
            this.flowLayoutPanel8.Name = "flowLayoutPanel8";
            this.flowLayoutPanel8.Size = new System.Drawing.Size(255, 26);
            this.flowLayoutPanel8.TabIndex = 14;
            // 
            // TriadicLower
            // 
            this.TriadicLower.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TriadicLower.Location = new System.Drawing.Point(3, 3);
            this.TriadicLower.Name = "TriadicLower";
            this.TriadicLower.Size = new System.Drawing.Size(20, 20);
            this.TriadicLower.TabIndex = 2;
            this.TriadicLower.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // TriadicSelected
            // 
            this.TriadicSelected.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TriadicSelected.Location = new System.Drawing.Point(29, 3);
            this.TriadicSelected.Name = "TriadicSelected";
            this.TriadicSelected.Size = new System.Drawing.Size(20, 20);
            this.TriadicSelected.TabIndex = 3;
            this.TriadicSelected.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // TriadicHigher
            // 
            this.TriadicHigher.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TriadicHigher.Location = new System.Drawing.Point(55, 3);
            this.TriadicHigher.Name = "TriadicHigher";
            this.TriadicHigher.Size = new System.Drawing.Size(20, 20);
            this.TriadicHigher.TabIndex = 4;
            this.TriadicHigher.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // flowLayoutPanel4
            // 
            this.flowLayoutPanel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.flowLayoutPanel4.Controls.Add(this.label4);
            this.flowLayoutPanel4.Controls.Add(this.flowLayoutPanel9);
            this.flowLayoutPanel4.Location = new System.Drawing.Point(20, 517);
            this.flowLayoutPanel4.Name = "flowLayoutPanel4";
            this.flowLayoutPanel4.Size = new System.Drawing.Size(263, 67);
            this.flowLayoutPanel4.TabIndex = 4;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("MS Reference Sans Serif", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(3, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(161, 23);
            this.label4.TabIndex = 6;
            this.label4.Text = "Tetradic Colors";
            // 
            // flowLayoutPanel9
            // 
            this.flowLayoutPanel9.Controls.Add(this.Tetradic1);
            this.flowLayoutPanel9.Controls.Add(this.TetradicSelected);
            this.flowLayoutPanel9.Controls.Add(this.Tetradic2);
            this.flowLayoutPanel9.Controls.Add(this.Tetradic3);
            this.flowLayoutPanel9.Location = new System.Drawing.Point(3, 26);
            this.flowLayoutPanel9.Name = "flowLayoutPanel9";
            this.flowLayoutPanel9.Size = new System.Drawing.Size(255, 26);
            this.flowLayoutPanel9.TabIndex = 15;
            // 
            // Tetradic1
            // 
            this.Tetradic1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Tetradic1.Location = new System.Drawing.Point(3, 3);
            this.Tetradic1.Name = "Tetradic1";
            this.Tetradic1.Size = new System.Drawing.Size(20, 20);
            this.Tetradic1.TabIndex = 3;
            this.Tetradic1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // TetradicSelected
            // 
            this.TetradicSelected.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TetradicSelected.Location = new System.Drawing.Point(29, 3);
            this.TetradicSelected.Name = "TetradicSelected";
            this.TetradicSelected.Size = new System.Drawing.Size(20, 20);
            this.TetradicSelected.TabIndex = 4;
            this.TetradicSelected.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // Tetradic2
            // 
            this.Tetradic2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Tetradic2.Location = new System.Drawing.Point(55, 3);
            this.Tetradic2.Name = "Tetradic2";
            this.Tetradic2.Size = new System.Drawing.Size(20, 20);
            this.Tetradic2.TabIndex = 5;
            this.Tetradic2.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // Tetradic3
            // 
            this.Tetradic3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Tetradic3.Location = new System.Drawing.Point(81, 3);
            this.Tetradic3.Name = "Tetradic3";
            this.Tetradic3.Size = new System.Drawing.Size(20, 20);
            this.Tetradic3.TabIndex = 6;
            this.Tetradic3.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // brightnessBar
            // 
            this.brightnessBar.Location = new System.Drawing.Point(20, 111);
            this.brightnessBar.Maximum = 240;
            this.brightnessBar.Name = "brightnessBar";
            this.brightnessBar.Size = new System.Drawing.Size(256, 45);
            this.brightnessBar.TabIndex = 6;
            this.brightnessBar.TickStyle = System.Windows.Forms.TickStyle.TopLeft;
            this.brightnessBar.ValueChanged += new System.EventHandler(this.brightnessBar_ValueChanged);
            // 
            // brightnessPanel
            // 
            this.brightnessPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.brightnessPanel.Location = new System.Drawing.Point(28, 94);
            this.brightnessPanel.Name = "brightnessPanel";
            this.brightnessPanel.Size = new System.Drawing.Size(240, 25);
            this.brightnessPanel.TabIndex = 7;
            // 
            // saturationPanel
            // 
            this.saturationPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.saturationPanel.Location = new System.Drawing.Point(29, 162);
            this.saturationPanel.Name = "saturationPanel";
            this.saturationPanel.Size = new System.Drawing.Size(240, 25);
            this.saturationPanel.TabIndex = 9;
            // 
            // saturationBar
            // 
            this.saturationBar.Location = new System.Drawing.Point(21, 178);
            this.saturationBar.Maximum = 240;
            this.saturationBar.Name = "saturationBar";
            this.saturationBar.Size = new System.Drawing.Size(256, 45);
            this.saturationBar.TabIndex = 8;
            this.saturationBar.TickStyle = System.Windows.Forms.TickStyle.TopLeft;
            this.saturationBar.ValueChanged += new System.EventHandler(this.saturationBar_ValueChanged);
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.flowLayoutPanel1.Controls.Add(this.label5);
            this.flowLayoutPanel1.Controls.Add(this.flowLayoutPanel5);
            this.flowLayoutPanel1.Location = new System.Drawing.Point(20, 229);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(263, 67);
            this.flowLayoutPanel1.TabIndex = 10;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("MS Reference Sans Serif", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(3, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(230, 23);
            this.label5.TabIndex = 6;
            this.label5.Text = "Monochromatic Colors";
            // 
            // flowLayoutPanel5
            // 
            this.flowLayoutPanel5.Controls.Add(this.MonoPanel1);
            this.flowLayoutPanel5.Controls.Add(this.MonoPanel2);
            this.flowLayoutPanel5.Controls.Add(this.MonoPanel3);
            this.flowLayoutPanel5.Controls.Add(this.MonoPanel4);
            this.flowLayoutPanel5.Controls.Add(this.MonoPanel5);
            this.flowLayoutPanel5.Controls.Add(this.MonoPanel6);
            this.flowLayoutPanel5.Controls.Add(this.MonoPanel7);
            this.flowLayoutPanel5.Location = new System.Drawing.Point(3, 26);
            this.flowLayoutPanel5.Name = "flowLayoutPanel5";
            this.flowLayoutPanel5.Size = new System.Drawing.Size(255, 26);
            this.flowLayoutPanel5.TabIndex = 11;
            // 
            // MonoPanel1
            // 
            this.MonoPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.MonoPanel1.Location = new System.Drawing.Point(3, 3);
            this.MonoPanel1.Name = "MonoPanel1";
            this.MonoPanel1.Size = new System.Drawing.Size(20, 20);
            this.MonoPanel1.TabIndex = 0;
            this.MonoPanel1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // MonoPanel2
            // 
            this.MonoPanel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.MonoPanel2.Location = new System.Drawing.Point(29, 3);
            this.MonoPanel2.Name = "MonoPanel2";
            this.MonoPanel2.Size = new System.Drawing.Size(20, 20);
            this.MonoPanel2.TabIndex = 1;
            this.MonoPanel2.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // MonoPanel3
            // 
            this.MonoPanel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.MonoPanel3.Location = new System.Drawing.Point(55, 3);
            this.MonoPanel3.Name = "MonoPanel3";
            this.MonoPanel3.Size = new System.Drawing.Size(20, 20);
            this.MonoPanel3.TabIndex = 1;
            this.MonoPanel3.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // MonoPanel4
            // 
            this.MonoPanel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.MonoPanel4.Location = new System.Drawing.Point(81, 3);
            this.MonoPanel4.Name = "MonoPanel4";
            this.MonoPanel4.Size = new System.Drawing.Size(20, 20);
            this.MonoPanel4.TabIndex = 2;
            this.MonoPanel4.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // MonoPanel5
            // 
            this.MonoPanel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.MonoPanel5.Location = new System.Drawing.Point(107, 3);
            this.MonoPanel5.Name = "MonoPanel5";
            this.MonoPanel5.Size = new System.Drawing.Size(20, 20);
            this.MonoPanel5.TabIndex = 3;
            this.MonoPanel5.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // MonoPanel6
            // 
            this.MonoPanel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.MonoPanel6.Location = new System.Drawing.Point(133, 3);
            this.MonoPanel6.Name = "MonoPanel6";
            this.MonoPanel6.Size = new System.Drawing.Size(20, 20);
            this.MonoPanel6.TabIndex = 4;
            this.MonoPanel6.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // MonoPanel7
            // 
            this.MonoPanel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.MonoPanel7.Location = new System.Drawing.Point(159, 3);
            this.MonoPanel7.Name = "MonoPanel7";
            this.MonoPanel7.Size = new System.Drawing.Size(20, 20);
            this.MonoPanel7.TabIndex = 5;
            this.MonoPanel7.MouseDown += new System.Windows.Forms.MouseEventHandler(this.MatchingPanel_MouseDown);
            // 
            // ColorPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.flowLayoutPanel1);
            this.Controls.Add(this.AnalogousColorPanel);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.flowLayoutPanel4);
            this.Controls.Add(this.flowLayoutPanel3);
            this.Controls.Add(this.flowLayoutPanel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.saturationPanel);
            this.Controls.Add(this.brightnessPanel);
            this.Controls.Add(this.brightnessBar);
            this.Controls.Add(this.saturationBar);
            this.Name = "ColorPane";
            this.Size = new System.Drawing.Size(304, 631);
            this.AnalogousColorPanel.ResumeLayout(false);
            this.AnalogousColorPanel.PerformLayout();
            this.flowLayoutPanel6.ResumeLayout(false);
            this.flowLayoutPanel2.ResumeLayout(false);
            this.flowLayoutPanel2.PerformLayout();
            this.flowLayoutPanel7.ResumeLayout(false);
            this.flowLayoutPanel3.ResumeLayout(false);
            this.flowLayoutPanel3.PerformLayout();
            this.flowLayoutPanel8.ResumeLayout(false);
            this.flowLayoutPanel4.ResumeLayout(false);
            this.flowLayoutPanel4.PerformLayout();
            this.flowLayoutPanel9.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.brightnessBar)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.saturationBar)).EndInit();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.flowLayoutPanel5.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.FlowLayoutPanel AnalogousColorPanel;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel3;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TrackBar brightnessBar;
        private System.Windows.Forms.Panel brightnessPanel;
        private System.Windows.Forms.Panel saturationPanel;
        private System.Windows.Forms.TrackBar saturationBar;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Panel MonoPanel1;
        private System.Windows.Forms.Panel MonoPanel2;
        private System.Windows.Forms.Panel MonoPanel3;
        private System.Windows.Forms.Panel MonoPanel4;
        private System.Windows.Forms.Panel MonoPanel5;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel5;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel6;
        private System.Windows.Forms.Panel AnalogousLighter;
        private System.Windows.Forms.Panel AnalogousSelected;
        private System.Windows.Forms.Panel AnalogousDarker;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel7;
        private System.Windows.Forms.Panel ComplementaryLighter;
        private System.Windows.Forms.Panel ComplementarySelected;
        private System.Windows.Forms.Panel ComplementaryDarker;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel8;
        private System.Windows.Forms.Panel TriadicLower;
        private System.Windows.Forms.Panel TriadicSelected;
        private System.Windows.Forms.Panel TriadicHigher;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel9;
        private System.Windows.Forms.Panel Tetradic1;
        private System.Windows.Forms.Panel TetradicSelected;
        private System.Windows.Forms.Panel Tetradic2;
        private System.Windows.Forms.Panel Tetradic3;
        private System.Windows.Forms.Panel MonoPanel6;
        private System.Windows.Forms.Panel MonoPanel7;
    }
}
