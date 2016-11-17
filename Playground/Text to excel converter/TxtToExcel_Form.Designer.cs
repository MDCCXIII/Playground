namespace Playground.Text_to_excel_converter
{
    partial class TxtToExcel_Form
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
            if (disposing && (components != null)) {
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
            this.InstructionLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // InstructionLabel
            // 
            this.InstructionLabel.AutoSize = true;
            this.InstructionLabel.CausesValidation = false;
            this.InstructionLabel.Enabled = false;
            this.InstructionLabel.Location = new System.Drawing.Point(13, 13);
            this.InstructionLabel.Name = "InstructionLabel";
            this.InstructionLabel.Size = new System.Drawing.Size(190, 26);
            this.InstructionLabel.TabIndex = 0;
            this.InstructionLabel.Text = "Drag and Drop the .txt file \r\nyou would like to convert into this form.";
            // 
            // TxtToExcel_Form
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.ClientSize = new System.Drawing.Size(367, 142);
            this.Controls.Add(this.InstructionLabel);
            this.Name = "TxtToExcel_Form";
            this.Text = "Txt to Excel Converter";
            this.Load += new System.EventHandler(this.TxtToExcel_Form_Load);
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.TxtToExcel_Form_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.TxtToExcel_Form_DragEnter);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label InstructionLabel;
    }
}