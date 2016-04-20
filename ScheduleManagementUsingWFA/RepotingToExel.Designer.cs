namespace ScheduleManagementUsingWFA
{
    partial class RepotingToExel
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
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource1 = new Microsoft.Reporting.WinForms.ReportDataSource();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.ScheduleManagementDataSet = new ScheduleManagementUsingWFA.ScheduleManagementDataSet();
            this.XepLichBindingSource = new System.Windows.Forms.BindingSource(this.components);
            
            ((System.ComponentModel.ISupportInitialize)(this.ScheduleManagementDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.XepLichBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource1.Name = "DataSetXepLich";
            reportDataSource1.Value = this.XepLichBindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "ScheduleManagementUsingWFA.Report1.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(0, 0);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(284, 261);
            this.reportViewer1.TabIndex = 0;
            // 
            // ScheduleManagementDataSet
            // 
            this.ScheduleManagementDataSet.DataSetName = "ScheduleManagementDataSet";
            this.ScheduleManagementDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // XepLichBindingSource
            // 
            this.XepLichBindingSource.DataMember = "XepLich";
            this.XepLichBindingSource.DataSource = this.ScheduleManagementDataSet;
            // 
            // XepLichTableAdapter
            // 
            
            // 
            // RepotingToExel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.reportViewer1);
            this.Name = "RepotingToExel";
            this.Text = "RepotingToExel";
            this.Load += new System.EventHandler(this.RepotingToExel_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ScheduleManagementDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.XepLichBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource XepLichBindingSource;
        private ScheduleManagementDataSet ScheduleManagementDataSet;
        
    }
}