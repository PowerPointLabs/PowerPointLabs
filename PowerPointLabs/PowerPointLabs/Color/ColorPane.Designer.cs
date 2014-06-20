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
            this.themeLayoutPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.label6 = new System.Windows.Forms.Label();
            this.flowLayoutPanel11 = new System.Windows.Forms.FlowLayoutPanel();
            this.ThemePanel1 = new System.Windows.Forms.Panel();
            this.ThemePanel2 = new System.Windows.Forms.Panel();
            this.ThemePanel3 = new System.Windows.Forms.Panel();
            this.ThemePanel4 = new System.Windows.Forms.Panel();
            this.ThemePanel5 = new System.Windows.Forms.Panel();
            this.ThemePanel6 = new System.Windows.Forms.Panel();
            this.ThemePanel7 = new System.Windows.Forms.Panel();
            this.ThemePanel8 = new System.Windows.Forms.Panel();
            this.ThemePanel9 = new System.Windows.Forms.Panel();
            this.SaveThemeButton = new System.Windows.Forms.Button();
            this.LoadButton = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.FontEyeDropperButton = new System.Windows.Forms.Button();
            this.LineEyeDropperButton = new System.Windows.Forms.Button();
            this.FillEyeDropperButton = new System.Windows.Forms.Button();
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
            this.themeLayoutPanel.SuspendLayout();
            this.flowLayoutPanel11.SuspendLayout();
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
            this.panel1.DragDrop += new System.Windows.Forms.DragEventHandler(this.panel_DragDrop);
            this.panel1.DragEnter += new System.Windows.Forms.DragEventHandler(this.panel_DragEnter);
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
            this.AnalogousColorPanel.Location = new System.Drawing.Point(16, 288);
            this.AnalogousColorPanel.Name = "AnalogousColorPanel";
            this.AnalogousColorPanel.Size = new System.Drawing.Size(124, 76);
            this.AnalogousColorPanel.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(96, 38);
            this.label1.TabIndex = 5;
            this.label1.Text = "Analogous Colors";
            // 
            // flowLayoutPanel6
            // 
            this.flowLayoutPanel6.Controls.Add(this.AnalogousLighter);
            this.flowLayoutPanel6.Controls.Add(this.AnalogousSelected);
            this.flowLayoutPanel6.Controls.Add(this.AnalogousDarker);
            this.flowLayoutPanel6.Location = new System.Drawing.Point(3, 41);
            this.flowLayoutPanel6.Name = "flowLayoutPanel6";
            this.flowLayoutPanel6.Size = new System.Drawing.Size(85, 26);
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
            this.flowLayoutPanel2.Location = new System.Drawing.Point(146, 288);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Size = new System.Drawing.Size(143, 76);
            this.flowLayoutPanel2.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(3, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(134, 38);
            this.label2.TabIndex = 6;
            this.label2.Text = "Complementary Colors";
            // 
            // flowLayoutPanel7
            // 
            this.flowLayoutPanel7.Controls.Add(this.ComplementaryLighter);
            this.flowLayoutPanel7.Controls.Add(this.ComplementarySelected);
            this.flowLayoutPanel7.Controls.Add(this.ComplementaryDarker);
            this.flowLayoutPanel7.Location = new System.Drawing.Point(3, 41);
            this.flowLayoutPanel7.Name = "flowLayoutPanel7";
            this.flowLayoutPanel7.Size = new System.Drawing.Size(120, 26);
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
            this.flowLayoutPanel3.Location = new System.Drawing.Point(16, 370);
            this.flowLayoutPanel3.Name = "flowLayoutPanel3";
            this.flowLayoutPanel3.Size = new System.Drawing.Size(124, 64);
            this.flowLayoutPanel3.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(3, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(115, 19);
            this.label3.TabIndex = 6;
            this.label3.Text = "Triadic Colors";
            // 
            // flowLayoutPanel8
            // 
            this.flowLayoutPanel8.Controls.Add(this.TriadicLower);
            this.flowLayoutPanel8.Controls.Add(this.TriadicSelected);
            this.flowLayoutPanel8.Controls.Add(this.TriadicHigher);
            this.flowLayoutPanel8.Location = new System.Drawing.Point(3, 22);
            this.flowLayoutPanel8.Name = "flowLayoutPanel8";
            this.flowLayoutPanel8.Size = new System.Drawing.Size(85, 26);
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
            this.flowLayoutPanel4.Location = new System.Drawing.Point(146, 370);
            this.flowLayoutPanel4.Name = "flowLayoutPanel4";
            this.flowLayoutPanel4.Size = new System.Drawing.Size(143, 64);
            this.flowLayoutPanel4.TabIndex = 4;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(3, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(125, 19);
            this.label4.TabIndex = 6;
            this.label4.Text = "Tetradic Colors";
            // 
            // flowLayoutPanel9
            // 
            this.flowLayoutPanel9.Controls.Add(this.Tetradic1);
            this.flowLayoutPanel9.Controls.Add(this.TetradicSelected);
            this.flowLayoutPanel9.Controls.Add(this.Tetradic2);
            this.flowLayoutPanel9.Controls.Add(this.Tetradic3);
            this.flowLayoutPanel9.Location = new System.Drawing.Point(3, 22);
            this.flowLayoutPanel9.Name = "flowLayoutPanel9";
            this.flowLayoutPanel9.Size = new System.Drawing.Size(104, 26);
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
            this.brightnessBar.Location = new System.Drawing.Point(20, 98);
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
            this.brightnessPanel.Location = new System.Drawing.Point(28, 81);
            this.brightnessPanel.Name = "brightnessPanel";
            this.brightnessPanel.Size = new System.Drawing.Size(240, 25);
            this.brightnessPanel.TabIndex = 7;
            // 
            // saturationPanel
            // 
            this.saturationPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.saturationPanel.Location = new System.Drawing.Point(29, 149);
            this.saturationPanel.Name = "saturationPanel";
            this.saturationPanel.Size = new System.Drawing.Size(240, 25);
            this.saturationPanel.TabIndex = 9;
            // 
            // saturationBar
            // 
            this.saturationBar.Location = new System.Drawing.Point(21, 165);
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
            this.flowLayoutPanel1.Location = new System.Drawing.Point(16, 216);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(273, 58);
            this.flowLayoutPanel1.TabIndex = 10;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(3, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(183, 19);
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
            this.flowLayoutPanel5.Location = new System.Drawing.Point(3, 22);
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
            // themeLayoutPanel
            // 
            this.themeLayoutPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.themeLayoutPanel.Controls.Add(this.label6);
            this.themeLayoutPanel.Controls.Add(this.flowLayoutPanel11);
            this.themeLayoutPanel.Controls.Add(this.SaveThemeButton);
            this.themeLayoutPanel.Controls.Add(this.LoadButton);
            this.themeLayoutPanel.Location = new System.Drawing.Point(16, 440);
            this.themeLayoutPanel.Name = "themeLayoutPanel";
            this.themeLayoutPanel.Size = new System.Drawing.Size(273, 91);
            this.themeLayoutPanel.TabIndex = 11;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(3, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(116, 19);
            this.label6.TabIndex = 6;
            this.label6.Text = "Theme Colors";
            // 
            // flowLayoutPanel11
            // 
            this.flowLayoutPanel11.Controls.Add(this.ThemePanel1);
            this.flowLayoutPanel11.Controls.Add(this.ThemePanel2);
            this.flowLayoutPanel11.Controls.Add(this.ThemePanel3);
            this.flowLayoutPanel11.Controls.Add(this.ThemePanel4);
            this.flowLayoutPanel11.Controls.Add(this.ThemePanel5);
            this.flowLayoutPanel11.Controls.Add(this.ThemePanel6);
            this.flowLayoutPanel11.Controls.Add(this.ThemePanel7);
            this.flowLayoutPanel11.Controls.Add(this.ThemePanel8);
            this.flowLayoutPanel11.Controls.Add(this.ThemePanel9);
            this.flowLayoutPanel11.Location = new System.Drawing.Point(3, 22);
            this.flowLayoutPanel11.Name = "flowLayoutPanel11";
            this.flowLayoutPanel11.Size = new System.Drawing.Size(255, 26);
            this.flowLayoutPanel11.TabIndex = 15;
            // 
            // ThemePanel1
            // 
            this.ThemePanel1.AllowDrop = true;
            this.ThemePanel1.BackColor = System.Drawing.Color.White;
            this.ThemePanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ThemePanel1.Location = new System.Drawing.Point(3, 3);
            this.ThemePanel1.Name = "ThemePanel1";
            this.ThemePanel1.Size = new System.Drawing.Size(20, 20);
            this.ThemePanel1.TabIndex = 3;
            this.ThemePanel1.Click += new System.EventHandler(this.ThemePanel_Click);
            this.ThemePanel1.DragDrop += new System.Windows.Forms.DragEventHandler(this.panel_DragDrop);
            this.ThemePanel1.DragEnter += new System.Windows.Forms.DragEventHandler(this.panel_DragEnter);
            // 
            // ThemePanel2
            // 
            this.ThemePanel2.AllowDrop = true;
            this.ThemePanel2.BackColor = System.Drawing.Color.White;
            this.ThemePanel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ThemePanel2.Location = new System.Drawing.Point(29, 3);
            this.ThemePanel2.Name = "ThemePanel2";
            this.ThemePanel2.Size = new System.Drawing.Size(20, 20);
            this.ThemePanel2.TabIndex = 4;
            this.ThemePanel2.Click += new System.EventHandler(this.ThemePanel_Click);
            this.ThemePanel2.DragDrop += new System.Windows.Forms.DragEventHandler(this.panel_DragDrop);
            this.ThemePanel2.DragEnter += new System.Windows.Forms.DragEventHandler(this.panel_DragEnter);
            // 
            // ThemePanel3
            // 
            this.ThemePanel3.AllowDrop = true;
            this.ThemePanel3.BackColor = System.Drawing.Color.White;
            this.ThemePanel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ThemePanel3.Location = new System.Drawing.Point(55, 3);
            this.ThemePanel3.Name = "ThemePanel3";
            this.ThemePanel3.Size = new System.Drawing.Size(20, 20);
            this.ThemePanel3.TabIndex = 5;
            this.ThemePanel3.Click += new System.EventHandler(this.ThemePanel_Click);
            this.ThemePanel3.DragDrop += new System.Windows.Forms.DragEventHandler(this.panel_DragDrop);
            this.ThemePanel3.DragEnter += new System.Windows.Forms.DragEventHandler(this.panel_DragEnter);
            // 
            // ThemePanel4
            // 
            this.ThemePanel4.AllowDrop = true;
            this.ThemePanel4.BackColor = System.Drawing.Color.White;
            this.ThemePanel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ThemePanel4.Location = new System.Drawing.Point(81, 3);
            this.ThemePanel4.Name = "ThemePanel4";
            this.ThemePanel4.Size = new System.Drawing.Size(20, 20);
            this.ThemePanel4.TabIndex = 6;
            this.ThemePanel4.Click += new System.EventHandler(this.ThemePanel_Click);
            this.ThemePanel4.DragDrop += new System.Windows.Forms.DragEventHandler(this.panel_DragDrop);
            this.ThemePanel4.DragEnter += new System.Windows.Forms.DragEventHandler(this.panel_DragEnter);
            // 
            // ThemePanel5
            // 
            this.ThemePanel5.AllowDrop = true;
            this.ThemePanel5.BackColor = System.Drawing.Color.White;
            this.ThemePanel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ThemePanel5.Location = new System.Drawing.Point(107, 3);
            this.ThemePanel5.Name = "ThemePanel5";
            this.ThemePanel5.Size = new System.Drawing.Size(20, 20);
            this.ThemePanel5.TabIndex = 7;
            this.ThemePanel5.Click += new System.EventHandler(this.ThemePanel_Click);
            this.ThemePanel5.DragDrop += new System.Windows.Forms.DragEventHandler(this.panel_DragDrop);
            this.ThemePanel5.DragEnter += new System.Windows.Forms.DragEventHandler(this.panel_DragEnter);
            // 
            // ThemePanel6
            // 
            this.ThemePanel6.AllowDrop = true;
            this.ThemePanel6.BackColor = System.Drawing.Color.White;
            this.ThemePanel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ThemePanel6.Location = new System.Drawing.Point(133, 3);
            this.ThemePanel6.Name = "ThemePanel6";
            this.ThemePanel6.Size = new System.Drawing.Size(20, 20);
            this.ThemePanel6.TabIndex = 8;
            this.ThemePanel6.Click += new System.EventHandler(this.ThemePanel_Click);
            this.ThemePanel6.DragDrop += new System.Windows.Forms.DragEventHandler(this.panel_DragDrop);
            this.ThemePanel6.DragEnter += new System.Windows.Forms.DragEventHandler(this.panel_DragEnter);
            // 
            // ThemePanel7
            // 
            this.ThemePanel7.AllowDrop = true;
            this.ThemePanel7.BackColor = System.Drawing.Color.White;
            this.ThemePanel7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ThemePanel7.Location = new System.Drawing.Point(159, 3);
            this.ThemePanel7.Name = "ThemePanel7";
            this.ThemePanel7.Size = new System.Drawing.Size(20, 20);
            this.ThemePanel7.TabIndex = 9;
            this.ThemePanel7.Click += new System.EventHandler(this.ThemePanel_Click);
            this.ThemePanel7.DragDrop += new System.Windows.Forms.DragEventHandler(this.panel_DragDrop);
            this.ThemePanel7.DragEnter += new System.Windows.Forms.DragEventHandler(this.panel_DragEnter);
            // 
            // ThemePanel8
            // 
            this.ThemePanel8.AllowDrop = true;
            this.ThemePanel8.BackColor = System.Drawing.Color.White;
            this.ThemePanel8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ThemePanel8.Location = new System.Drawing.Point(185, 3);
            this.ThemePanel8.Name = "ThemePanel8";
            this.ThemePanel8.Size = new System.Drawing.Size(20, 20);
            this.ThemePanel8.TabIndex = 10;
            this.ThemePanel8.Click += new System.EventHandler(this.ThemePanel_Click);
            this.ThemePanel8.DragDrop += new System.Windows.Forms.DragEventHandler(this.panel_DragDrop);
            this.ThemePanel8.DragEnter += new System.Windows.Forms.DragEventHandler(this.panel_DragEnter);
            // 
            // ThemePanel9
            // 
            this.ThemePanel9.AllowDrop = true;
            this.ThemePanel9.BackColor = System.Drawing.Color.White;
            this.ThemePanel9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ThemePanel9.Location = new System.Drawing.Point(211, 3);
            this.ThemePanel9.Name = "ThemePanel9";
            this.ThemePanel9.Size = new System.Drawing.Size(20, 20);
            this.ThemePanel9.TabIndex = 11;
            this.ThemePanel9.Click += new System.EventHandler(this.ThemePanel_Click);
            this.ThemePanel9.DragDrop += new System.Windows.Forms.DragEventHandler(this.panel_DragDrop);
            this.ThemePanel9.DragEnter += new System.Windows.Forms.DragEventHandler(this.panel_DragEnter);
            // 
            // SaveThemeButton
            // 
            this.SaveThemeButton.Location = new System.Drawing.Point(3, 54);
            this.SaveThemeButton.Name = "SaveThemeButton";
            this.SaveThemeButton.Size = new System.Drawing.Size(75, 23);
            this.SaveThemeButton.TabIndex = 16;
            this.SaveThemeButton.Text = "Save";
            this.SaveThemeButton.UseVisualStyleBackColor = true;
            this.SaveThemeButton.Click += new System.EventHandler(this.SaveThemeButton_Click);
            // 
            // LoadButton
            // 
            this.LoadButton.Location = new System.Drawing.Point(84, 54);
            this.LoadButton.Name = "LoadButton";
            this.LoadButton.Size = new System.Drawing.Size(75, 23);
            this.LoadButton.TabIndex = 17;
            this.LoadButton.Text = "Load";
            this.LoadButton.UseVisualStyleBackColor = true;
            this.LoadButton.Click += new System.EventHandler(this.LoadButton_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "thm";
            this.saveFileDialog1.Filter = "PPTLabsThemes|*.thm";
            this.saveFileDialog1.Title = "Save Theme";
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "thm";
            this.openFileDialog1.Filter = "PPTLabsTheme|*.thm";
            this.openFileDialog1.Title = "Load Theme";
            // 
            // FontEyeDropperButton
            // 
            this.FontEyeDropperButton.Location = new System.Drawing.Point(101, 40);
            this.FontEyeDropperButton.Name = "FontEyeDropperButton";
            this.FontEyeDropperButton.Size = new System.Drawing.Size(34, 23);
            this.FontEyeDropperButton.TabIndex = 12;
            this.FontEyeDropperButton.Text = "Font";
            this.FontEyeDropperButton.UseVisualStyleBackColor = true;
            this.FontEyeDropperButton.Click += new System.EventHandler(this.FontEyeDropperButton_Click);
            // 
            // LineEyeDropperButton
            // 
            this.LineEyeDropperButton.Location = new System.Drawing.Point(186, 40);
            this.LineEyeDropperButton.Name = "LineEyeDropperButton";
            this.LineEyeDropperButton.Size = new System.Drawing.Size(35, 23);
            this.LineEyeDropperButton.TabIndex = 13;
            this.LineEyeDropperButton.Text = "Line";
            this.LineEyeDropperButton.UseVisualStyleBackColor = true;
            this.LineEyeDropperButton.Click += new System.EventHandler(this.LineEyeDropperButton_Click);
            // 
            // FillEyeDropperButton
            // 
            this.FillEyeDropperButton.Location = new System.Drawing.Point(226, 40);
            this.FillEyeDropperButton.Name = "FillEyeDropperButton";
            this.FillEyeDropperButton.Size = new System.Drawing.Size(35, 23);
            this.FillEyeDropperButton.TabIndex = 14;
            this.FillEyeDropperButton.Text = "Fill";
            this.FillEyeDropperButton.UseVisualStyleBackColor = true;
            this.FillEyeDropperButton.Click += new System.EventHandler(this.FillEyeDropperButton_Click);
            // 
            // ColorPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.FillEyeDropperButton);
            this.Controls.Add(this.LineEyeDropperButton);
            this.Controls.Add(this.FontEyeDropperButton);
            this.Controls.Add(this.themeLayoutPanel);
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
            this.Size = new System.Drawing.Size(304, 542);
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
            this.themeLayoutPanel.ResumeLayout(false);
            this.themeLayoutPanel.PerformLayout();
            this.flowLayoutPanel11.ResumeLayout(false);
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
        private System.Windows.Forms.FlowLayoutPanel themeLayoutPanel;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel11;
        private System.Windows.Forms.Panel ThemePanel1;
        private System.Windows.Forms.Panel ThemePanel2;
        private System.Windows.Forms.Panel ThemePanel3;
        private System.Windows.Forms.Panel ThemePanel4;
        private System.Windows.Forms.Panel ThemePanel5;
        private System.Windows.Forms.Panel ThemePanel6;
        private System.Windows.Forms.Panel ThemePanel7;
        private System.Windows.Forms.Panel ThemePanel8;
        private System.Windows.Forms.Panel ThemePanel9;
        private System.Windows.Forms.Button SaveThemeButton;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button LoadButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button FontEyeDropperButton;
        private System.Windows.Forms.Button LineEyeDropperButton;
        private System.Windows.Forms.Button FillEyeDropperButton;
    }
}
