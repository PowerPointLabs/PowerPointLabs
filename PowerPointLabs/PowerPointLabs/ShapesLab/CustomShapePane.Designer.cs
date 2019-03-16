using System.Windows.Controls;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ShapesLab.Views;

namespace PowerPointLabs.ShapesLab
{
    partial class CustomShapePane
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public void UpdateOnSelectionChange(Selection sel)
        {
            this.CustomShapePaneWPF1.UpdateAddShapeButtonEnabledStatus(sel);
        }

        public void AddCustomShape(string shapeName, string shapePath, bool immediateEditing)
        {
            CustomShapePaneWPF1.AddCustomShape(shapeName, shapePath, immediateEditing);
        }

        public void RemoveCustomShape(string shapeName)
        {
            CustomShapePaneWPF1.RemoveCustomShape(shapeName);
        }

        public void RenameCustomShape(string shapeOldName, string shapeNewName)
        {
            CustomShapePaneWPF1.RenameCustomShape(shapeOldName, shapeNewName);
        }

        public void AddCategory(string newCategoryName)
        {
            CustomShapePaneWPF1.AddCategory(newCategoryName);
        }

        public void RemoveCategory(int removedCategoryIndex)
        {
            CustomShapePaneWPF1.RemoveCategory(removedCategoryIndex);
        }

        public void RenameCategory(int renameCategoryIndex, string newCategoryName)
        {
            CustomShapePaneWPF1.RenameCategory(renameCategoryIndex, newCategoryName);
        }

        public void InitCustomShapePaneStorage()
        {
            CustomShapePaneWPF1.SetStorageSettings();
        }

        #region Test Interface

        public CustomShapePaneItem GetShape(string shapeName)
        {
            return CustomShapePaneWPF1.GetShape(shapeName);
        }

        public void ImportLibrary(string pathToLibrary)
        {
            CustomShapePaneWPF1.ImportLibrary(pathToLibrary);
        }

        public void ImportShape(string pathToShape)
        {
            CustomShapePaneWPF1.ImportShape(pathToShape);
        }

        public Presentation GetShapeGallery()
        {
            return CustomShapePaneWPF1.GetShapeGallery();
        }

        public Button GetAddShapeButton()
        {
            return CustomShapePaneWPF1.addShapeButton;
        }

        public void SaveSelectedShapes()
        {
            CustomShapePaneWPF1.SaveSelectedShapes();
        }

        public System.Windows.Point GetShapeForClicking(string shapeName)
        {
            return CustomShapePaneWPF1.GetShapeForClicking(shapeName);
        }

        #endregion

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
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.CustomShapePaneWPF1 = new PowerPointLabs.ShapesLab.Views.CustomShapePaneWPF();
            this.SuspendLayout();
            // 
            // elementHost1
            // 
            this.elementHost1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost1.Location = new System.Drawing.Point(0, 0);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Size = new System.Drawing.Size(300, 833);
            this.elementHost1.TabIndex = 0;
            this.elementHost1.Text = "elementHost1";
            this.elementHost1.Child = this.CustomShapePaneWPF1;
            // 
            // CustomShapePane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.elementHost1);
            this.Name = "CustomShapePane";
            this.Size = new System.Drawing.Size(300, 833);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Integration.ElementHost elementHost1;
        public CustomShapePaneWPF CustomShapePaneWPF1 { get; private set; }

        public string CurrentCategory
        {
            get
            {
                return CustomShapePaneWPF1.CurrentCategory;
            }
            set
            {
                CustomShapePaneWPF1.CurrentCategory = value;
            }
        }

    }
}
