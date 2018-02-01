using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Windows.Forms.Design;

namespace Gui.Wizard
{
	/// <summary>
	/// Summary description for UserControl1.
	/// </summary>
	[Designer(typeof(InfoContainerDesigner))]
	public class InfoContainer: System.Windows.Forms.UserControl
    {
        private Label lblTitle;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		/// <summary>
		/// 
		/// </summary>
		public InfoContainer()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
		}

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.lblTitle = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.AutoEllipsis = true;
            this.lblTitle.AutoSize = true;
            this.lblTitle.BackColor = System.Drawing.Color.Transparent;
            this.lblTitle.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(98, 19);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(257, 19);
            this.lblTitle.TabIndex = 8;
            this.lblTitle.Text = "BigFix Excel Connector Wizard";
            this.lblTitle.Visible = false;
            // 
            // InfoContainer
            // 
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.lblTitle);
            this.Name = "InfoContainer";
            this.Size = new System.Drawing.Size(480, 388);
            this.Load += new System.EventHandler(this.InfoContainer_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		private void InfoContainer_Load(object sender, System.EventArgs e)
		{
			//Handle really irating resize that doesn't take account of Anchor
			// lblTitle.Left = picImage.Width+8;
			lblTitle.Width = (this.Width-4)-lblTitle.Left;
		}

		/// <summary>
		/// Get/Set the title for the info page
		/// </summary>
		[Category("Appearance")]
		public string PageTitle
		{
			get
			{
				return lblTitle.Text;
			}
			set
			{
				lblTitle.Text = value;
			}
		}

		
		/// <summary>
		/// Gets/Sets the Icon
		/// </summary>
		[Category("Appearance")]
		public Image Image
		{
			get
			{
				// return picImage.Image;
                return null;
			}
			set
			{
				// picImage.Image = value;
			}
		}


	}
}
