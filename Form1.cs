using System;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using Microsoft.Win32;
using AxSHDocVw;
using mshtml;
using System.Runtime.InteropServices;
using System.IO;
using System.Linq;
using System.Media;
using System.Resources;
using System.Net;
using System.Security.Policy;
using System.Xml;
using NUnit.Framework;
using TCore.WebControl;

namespace upc
{

	public enum ADAS
		{
		Book,
		DVD,
		Generic,
		Wine, 
		Unknown	
		};
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
        string sConnString = "";
	    private StatusBox.StatusRpt m_srpt;

        //string sConnString = "Provider=SQLOLEDB; Data Source=cacofonix;Initial Catalog=db0902;Trusted_Connection=Yes";

		private AxSHDocVw.AxWebBrowser m_axBrowser;
		private System.Windows.Forms.TextBox m_ebIsbn;
		private System.Windows.Forms.Button m_pbManualAdd;
		private System.Windows.Forms.TextBox m_ebTitle;
		private System.Windows.Forms.TextBox m_ebLocation;
		private System.Windows.Forms.StatusBar m_sb;
        private System.Windows.Forms.ComboBox m_cbxAddAs;
        private Button m_pbNewScan;
        private TextBox m_ebShortDesc;
        private TextBox m_ebLongDesc;
        private Label label7;
        private TextBox m_ebManualCode;

        SoundPlayer m_snd;
        private Button m_pbBooks;
        private Button m_pbDvd;
        private Button m_pbWine;
        private RichTextBox m_recStatus;
        private PictureBox pictureBox1;
        private Button m_pbGeneric;
        private TextBox m_ebWine;
        private Label m_lblWine;
        private Label m_lblWineHeader;
        private TextBox m_ebVintage;
        private Label m_lblNotes;
        private TextBox m_ebNotes;
        private Button m_pbDrinkWine;
        private CheckBox m_cbScanningTabOrder;
        		
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
        private Button m_pbUpdateDvd;
        private Button m_pbCleanDVDs;
        private TextBox m_ebLogFile;
        private ResourceManager m_rsm;

		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
            m_snd = new SoundPlayer();
            EnsureResourcesLoaded();
		    m_rsm = new ResourceManager("upc.Secrets", typeof (Form1).Assembly);
		    sConnString = m_rsm.GetString("strAzureConnection");
		    sAccessKey = m_rsm.GetString("strIsbnDBAccess");
            SetIsbnFocus();
			ShowHideWine();
			//
			// TODO: Add any constructor code after InitializeComponent call
			//

		    m_srpt = new StatusBox.StatusRpt(m_recStatus);
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.Windows.Forms.Label label1;
            System.Windows.Forms.Label label2;
            System.Windows.Forms.Label label3;
            System.Windows.Forms.Label label4;
            System.Windows.Forms.Label label8;
            System.Windows.Forms.Label label9;
            System.Windows.Forms.Label label10;
            System.Windows.Forms.Label label11;
            System.Windows.Forms.Label label5;
            System.Windows.Forms.Label label6;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            System.Windows.Forms.Label label12;
            this.m_lblWineHeader = new System.Windows.Forms.Label();
            this.m_lblWine = new System.Windows.Forms.Label();
            this.m_ebIsbn = new System.Windows.Forms.TextBox();
            this.m_pbManualAdd = new System.Windows.Forms.Button();
            this.m_ebTitle = new System.Windows.Forms.TextBox();
            this.m_ebLocation = new System.Windows.Forms.TextBox();
            this.m_sb = new System.Windows.Forms.StatusBar();
            this.m_cbxAddAs = new System.Windows.Forms.ComboBox();
            this.m_pbNewScan = new System.Windows.Forms.Button();
            this.m_ebShortDesc = new System.Windows.Forms.TextBox();
            this.m_ebLongDesc = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.m_ebManualCode = new System.Windows.Forms.TextBox();
            this.m_pbBooks = new System.Windows.Forms.Button();
            this.m_pbDvd = new System.Windows.Forms.Button();
            this.m_pbWine = new System.Windows.Forms.Button();
            this.m_recStatus = new System.Windows.Forms.RichTextBox();
            this.m_pbGeneric = new System.Windows.Forms.Button();
            this.m_ebWine = new System.Windows.Forms.TextBox();
            this.m_ebVintage = new System.Windows.Forms.TextBox();
            this.m_lblNotes = new System.Windows.Forms.Label();
            this.m_ebNotes = new System.Windows.Forms.TextBox();
            this.m_pbDrinkWine = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.m_cbScanningTabOrder = new System.Windows.Forms.CheckBox();
            this.m_pbUpdateDvd = new System.Windows.Forms.Button();
            this.m_pbCleanDVDs = new System.Windows.Forms.Button();
            this.m_axBrowser = new AxSHDocVw.AxWebBrowser();
            this.m_ebLogFile = new System.Windows.Forms.TextBox();
            label1 = new System.Windows.Forms.Label();
            label2 = new System.Windows.Forms.Label();
            label3 = new System.Windows.Forms.Label();
            label4 = new System.Windows.Forms.Label();
            label8 = new System.Windows.Forms.Label();
            label9 = new System.Windows.Forms.Label();
            label10 = new System.Windows.Forms.Label();
            label11 = new System.Windows.Forms.Label();
            label5 = new System.Windows.Forms.Label();
            label6 = new System.Windows.Forms.Label();
            label12 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.m_axBrowser)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            label1.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label1.Location = new System.Drawing.Point(12, 43);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(81, 19);
            label1.TabIndex = 10;
            label1.Text = "Scan Code";
            // 
            // label2
            // 
            label2.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label2.Location = new System.Drawing.Point(12, 71);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(58, 21);
            label2.TabIndex = 1;
            label2.Text = "Title";
            // 
            // label3
            // 
            label3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label3.Location = new System.Drawing.Point(230, 41);
            label3.Name = "label3";
            label3.Size = new System.Drawing.Size(71, 24);
            label3.TabIndex = 8;
            label3.Text = "Location";
            // 
            // label4
            // 
            label4.Location = new System.Drawing.Point(420, 46);
            label4.Name = "label4";
            label4.Size = new System.Drawing.Size(69, 19);
            label4.TabIndex = 3;
            label4.Text = "Add items as";
            // 
            // label8
            // 
            label8.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            label8.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label8.Location = new System.Drawing.Point(531, 437);
            label8.Name = "label8";
            label8.Size = new System.Drawing.Size(94, 16);
            label8.TabIndex = 18;
            label8.Text = "UPC Code";
            // 
            // label9
            // 
            label9.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            label9.Cursor = System.Windows.Forms.Cursors.Default;
            label9.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label9.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            label9.Location = new System.Drawing.Point(12, 8);
            label9.Name = "label9";
            label9.Size = new System.Drawing.Size(794, 22);
            label9.TabIndex = 19;
            label9.Text = "Inventory Add && Update";
            label9.Paint += new System.Windows.Forms.PaintEventHandler(this.EH_RenderHeadingLine);
            // 
            // label10
            // 
            label10.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            label10.Cursor = System.Windows.Forms.Cursors.Default;
            label10.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label10.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            label10.Location = new System.Drawing.Point(12, 409);
            label10.Name = "label10";
            label10.Size = new System.Drawing.Size(794, 22);
            label10.TabIndex = 20;
            label10.Text = "Custom UPC Label Creation";
            label10.Paint += new System.Windows.Forms.PaintEventHandler(this.EH_RenderHeadingLine);
            // 
            // label11
            // 
            label11.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            label11.Cursor = System.Windows.Forms.Cursors.Default;
            label11.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label11.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            label11.Location = new System.Drawing.Point(12, 499);
            label11.Name = "label11";
            label11.Size = new System.Drawing.Size(794, 22);
            label11.TabIndex = 21;
            label11.Text = "Status";
            label11.Paint += new System.Windows.Forms.PaintEventHandler(this.EH_RenderHeadingLine);
            // 
            // label5
            // 
            label5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            label5.AutoEllipsis = true;
            label5.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label5.Location = new System.Drawing.Point(13, 437);
            label5.Name = "label5";
            label5.Size = new System.Drawing.Size(77, 16);
            label5.TabIndex = 12;
            label5.Text = "Short";
            // 
            // label6
            // 
            label6.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            label6.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label6.Location = new System.Drawing.Point(13, 463);
            label6.Name = "label6";
            label6.Size = new System.Drawing.Size(77, 18);
            label6.TabIndex = 14;
            label6.Text = "Long";
            // 
            // m_lblWineHeader
            // 
            this.m_lblWineHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_lblWineHeader.Cursor = System.Windows.Forms.Cursors.Default;
            this.m_lblWineHeader.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_lblWineHeader.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.m_lblWineHeader.Location = new System.Drawing.Point(37, 107);
            this.m_lblWineHeader.Name = "m_lblWineHeader";
            this.m_lblWineHeader.Size = new System.Drawing.Size(769, 22);
            this.m_lblWineHeader.TabIndex = 25;
            this.m_lblWineHeader.Text = "Drinking Information";
            this.m_lblWineHeader.Paint += new System.Windows.Forms.PaintEventHandler(this.EH_RenderHeadingLine);
            // 
            // m_lblWine
            // 
            this.m_lblWine.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_lblWine.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_lblWine.Location = new System.Drawing.Point(37, 141);
            this.m_lblWine.Name = "m_lblWine";
            this.m_lblWine.Size = new System.Drawing.Size(77, 23);
            this.m_lblWine.TabIndex = 26;
            this.m_lblWine.Text = "Wine";
            // 
            // m_ebIsbn
            // 
            this.m_ebIsbn.AcceptsReturn = true;
            this.m_ebIsbn.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_ebIsbn.Location = new System.Drawing.Point(93, 40);
            this.m_ebIsbn.Name = "m_ebIsbn";
            this.m_ebIsbn.Size = new System.Drawing.Size(134, 26);
            this.m_ebIsbn.TabIndex = 0;
            this.m_ebIsbn.Text = "Scan Code";
            this.m_ebIsbn.Click += new System.EventHandler(this.Control_OnClick);
            this.m_ebIsbn.Enter += new System.EventHandler(this.Control_OnEnter);
            this.m_ebIsbn.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.m_ebIsbn_KeyUp);
            // 
            // m_pbManualAdd
            // 
            this.m_pbManualAdd.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.m_pbManualAdd.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_pbManualAdd.Location = new System.Drawing.Point(681, 66);
            this.m_pbManualAdd.Name = "m_pbManualAdd";
            this.m_pbManualAdd.Size = new System.Drawing.Size(125, 29);
            this.m_pbManualAdd.TabIndex = 7;
            this.m_pbManualAdd.Text = "Manual Add";
            this.m_pbManualAdd.Click += new System.EventHandler(this.ManualAdd_Click);
            // 
            // m_ebTitle
            // 
            this.m_ebTitle.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_ebTitle.Location = new System.Drawing.Point(93, 68);
            this.m_ebTitle.Name = "m_ebTitle";
            this.m_ebTitle.Size = new System.Drawing.Size(569, 26);
            this.m_ebTitle.TabIndex = 2;
            this.m_ebTitle.Text = "Title";
            this.m_ebTitle.Click += new System.EventHandler(this.Control_OnClick);
            this.m_ebTitle.Enter += new System.EventHandler(this.Control_OnEnter);
            // 
            // m_ebLocation
            // 
            this.m_ebLocation.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_ebLocation.Location = new System.Drawing.Point(308, 38);
            this.m_ebLocation.Name = "m_ebLocation";
            this.m_ebLocation.Size = new System.Drawing.Size(106, 27);
            this.m_ebLocation.TabIndex = 9;
            this.m_ebLocation.Text = "!? LOC ?!";
            this.m_ebLocation.Click += new System.EventHandler(this.Control_OnClick);
            this.m_ebLocation.Enter += new System.EventHandler(this.Control_OnEnter);
            // 
            // m_sb
            // 
            this.m_sb.Location = new System.Drawing.Point(0, 662);
            this.m_sb.Name = "m_sb";
            this.m_sb.Size = new System.Drawing.Size(1016, 22);
            this.m_sb.TabIndex = 6;
            // 
            // m_cbxAddAs
            // 
            this.m_cbxAddAs.Items.AddRange(new object[] {
            "Book",
            "DVD",
            "Wine",
            "Generic"});
            this.m_cbxAddAs.Location = new System.Drawing.Point(495, 43);
            this.m_cbxAddAs.Name = "m_cbxAddAs";
            this.m_cbxAddAs.Size = new System.Drawing.Size(80, 21);
            this.m_cbxAddAs.TabIndex = 4;
            this.m_cbxAddAs.Text = "Book";
            this.m_cbxAddAs.SelectedValueChanged += new System.EventHandler(this.ChangeAddAs);
            // 
            // m_pbNewScan
            // 
            this.m_pbNewScan.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.m_pbNewScan.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.m_pbNewScan.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_pbNewScan.Location = new System.Drawing.Point(663, 467);
            this.m_pbNewScan.Name = "m_pbNewScan";
            this.m_pbNewScan.Size = new System.Drawing.Size(143, 29);
            this.m_pbNewScan.TabIndex = 11;
            this.m_pbNewScan.Text = "New CustomScan";
            this.m_pbNewScan.Click += new System.EventHandler(this.m_pbNewScan_Click);
            // 
            // m_ebShortDesc
            // 
            this.m_ebShortDesc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_ebShortDesc.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_ebShortDesc.Location = new System.Drawing.Point(93, 434);
            this.m_ebShortDesc.Name = "m_ebShortDesc";
            this.m_ebShortDesc.Size = new System.Drawing.Size(432, 26);
            this.m_ebShortDesc.TabIndex = 13;
            this.m_ebShortDesc.Text = "-- Short Description (To be printed on UPC label)";
            this.m_ebShortDesc.Click += new System.EventHandler(this.Control_OnClick);
            this.m_ebShortDesc.Enter += new System.EventHandler(this.Control_OnEnter);
            // 
            // m_ebLongDesc
            // 
            this.m_ebLongDesc.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_ebLongDesc.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_ebLongDesc.Location = new System.Drawing.Point(93, 460);
            this.m_ebLongDesc.Name = "m_ebLongDesc";
            this.m_ebLongDesc.Size = new System.Drawing.Size(432, 26);
            this.m_ebLongDesc.TabIndex = 15;
            this.m_ebLongDesc.Text = "-- Long Description (To be printed on UPC label)";
            this.m_ebLongDesc.Click += new System.EventHandler(this.Control_OnClick);
            this.m_ebLongDesc.Enter += new System.EventHandler(this.Control_OnEnter);
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(277, 473);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(69, 13);
            this.label7.TabIndex = 17;
            this.label7.Text = "Manual Code";
            // 
            // m_ebManualCode
            // 
            this.m_ebManualCode.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_ebManualCode.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_ebManualCode.Location = new System.Drawing.Point(631, 434);
            this.m_ebManualCode.Name = "m_ebManualCode";
            this.m_ebManualCode.Size = new System.Drawing.Size(175, 26);
            this.m_ebManualCode.TabIndex = 16;
            this.m_ebManualCode.Text = "--";
            this.m_ebManualCode.Click += new System.EventHandler(this.Control_OnClick);
            this.m_ebManualCode.Enter += new System.EventHandler(this.Control_OnEnter);
            // 
            // m_pbBooks
            // 
            this.m_pbBooks.Location = new System.Drawing.Point(857, 17);
            this.m_pbBooks.Name = "m_pbBooks";
            this.m_pbBooks.Size = new System.Drawing.Size(131, 145);
            this.m_pbBooks.TabIndex = 16;
            this.m_pbBooks.Tag = "0";
            this.m_pbBooks.Text = "Books";
            this.m_pbBooks.UseVisualStyleBackColor = true;
            this.m_pbBooks.Click += new System.EventHandler(this.ChangeInput);
            this.m_pbBooks.Paint += new System.Windows.Forms.PaintEventHandler(this.PaintButton);
            // 
            // m_pbDvd
            // 
            this.m_pbDvd.Location = new System.Drawing.Point(857, 179);
            this.m_pbDvd.Name = "m_pbDvd";
            this.m_pbDvd.Size = new System.Drawing.Size(131, 145);
            this.m_pbDvd.TabIndex = 17;
            this.m_pbDvd.Tag = "1";
            this.m_pbDvd.Text = "DVD";
            this.m_pbDvd.UseVisualStyleBackColor = true;
            this.m_pbDvd.Click += new System.EventHandler(this.ChangeInput);
            this.m_pbDvd.Paint += new System.Windows.Forms.PaintEventHandler(this.PaintButton);
            // 
            // m_pbWine
            // 
            this.m_pbWine.Location = new System.Drawing.Point(857, 341);
            this.m_pbWine.Name = "m_pbWine";
            this.m_pbWine.Size = new System.Drawing.Size(131, 145);
            this.m_pbWine.TabIndex = 18;
            this.m_pbWine.Tag = "2";
            this.m_pbWine.Text = "Wine";
            this.m_pbWine.UseVisualStyleBackColor = true;
            this.m_pbWine.Click += new System.EventHandler(this.ChangeInput);
            this.m_pbWine.Paint += new System.Windows.Forms.PaintEventHandler(this.PaintButton);
            // 
            // m_recStatus
            // 
            this.m_recStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_recStatus.Location = new System.Drawing.Point(15, 524);
            this.m_recStatus.Name = "m_recStatus";
            this.m_recStatus.Size = new System.Drawing.Size(791, 132);
            this.m_recStatus.TabIndex = 0;
            this.m_recStatus.Text = "";
            // 
            // m_pbGeneric
            // 
            this.m_pbGeneric.Location = new System.Drawing.Point(857, 503);
            this.m_pbGeneric.Name = "m_pbGeneric";
            this.m_pbGeneric.Size = new System.Drawing.Size(131, 145);
            this.m_pbGeneric.TabIndex = 24;
            this.m_pbGeneric.Tag = "3";
            this.m_pbGeneric.Text = "Generic";
            this.m_pbGeneric.UseVisualStyleBackColor = true;
            this.m_pbGeneric.Click += new System.EventHandler(this.ChangeInput);
            this.m_pbGeneric.Paint += new System.Windows.Forms.PaintEventHandler(this.PaintButton);
            // 
            // m_ebWine
            // 
            this.m_ebWine.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_ebWine.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_ebWine.Location = new System.Drawing.Point(120, 138);
            this.m_ebWine.Name = "m_ebWine";
            this.m_ebWine.Size = new System.Drawing.Size(579, 26);
            this.m_ebWine.TabIndex = 27;
            // 
            // m_ebVintage
            // 
            this.m_ebVintage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_ebVintage.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_ebVintage.Location = new System.Drawing.Point(705, 138);
            this.m_ebVintage.Name = "m_ebVintage";
            this.m_ebVintage.Size = new System.Drawing.Size(81, 26);
            this.m_ebVintage.TabIndex = 28;
            // 
            // m_lblNotes
            // 
            this.m_lblNotes.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_lblNotes.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_lblNotes.Location = new System.Drawing.Point(37, 170);
            this.m_lblNotes.Name = "m_lblNotes";
            this.m_lblNotes.Size = new System.Drawing.Size(77, 23);
            this.m_lblNotes.TabIndex = 29;
            this.m_lblNotes.Text = "Notes";
            // 
            // m_ebNotes
            // 
            this.m_ebNotes.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_ebNotes.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_ebNotes.Location = new System.Drawing.Point(120, 167);
            this.m_ebNotes.Name = "m_ebNotes";
            this.m_ebNotes.Size = new System.Drawing.Size(666, 26);
            this.m_ebNotes.TabIndex = 30;
            this.m_ebNotes.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.NotesKeyUp);
            // 
            // m_pbDrinkWine
            // 
            this.m_pbDrinkWine.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.m_pbDrinkWine.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_pbDrinkWine.Location = new System.Drawing.Point(681, 199);
            this.m_pbDrinkWine.Name = "m_pbDrinkWine";
            this.m_pbDrinkWine.Size = new System.Drawing.Size(125, 29);
            this.m_pbDrinkWine.TabIndex = 31;
            this.m_pbDrinkWine.Text = "Drink Wine";
            this.m_pbDrinkWine.Click += new System.EventHandler(this.DrinkWine);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.SystemColors.ScrollBar;
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pictureBox1.Location = new System.Drawing.Point(837, 4);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(173, 658);
            this.pictureBox1.TabIndex = 22;
            this.pictureBox1.TabStop = false;
            // 
            // m_cbScanningTabOrder
            // 
            this.m_cbScanningTabOrder.AutoSize = true;
            this.m_cbScanningTabOrder.Location = new System.Drawing.Point(40, 210);
            this.m_cbScanningTabOrder.Name = "m_cbScanningTabOrder";
            this.m_cbScanningTabOrder.Size = new System.Drawing.Size(178, 17);
            this.m_cbScanningTabOrder.TabIndex = 32;
            this.m_cbScanningTabOrder.Text = "Optimize scanning for inventory";
            this.m_cbScanningTabOrder.UseVisualStyleBackColor = true;
            // 
            // m_pbUpdateDvd
            // 
            this.m_pbUpdateDvd.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.m_pbUpdateDvd.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_pbUpdateDvd.Location = new System.Drawing.Point(553, 9);
            this.m_pbUpdateDvd.Name = "m_pbUpdateDvd";
            this.m_pbUpdateDvd.Size = new System.Drawing.Size(125, 29);
            this.m_pbUpdateDvd.TabIndex = 33;
            this.m_pbUpdateDvd.Text = "Update DVDs";
            this.m_pbUpdateDvd.Click += new System.EventHandler(this.DoUpdateDvdDescriptions);
            // 
            // m_pbCleanDVDs
            // 
            this.m_pbCleanDVDs.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.m_pbCleanDVDs.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_pbCleanDVDs.Location = new System.Drawing.Point(681, 9);
            this.m_pbCleanDVDs.Name = "m_pbCleanDVDs";
            this.m_pbCleanDVDs.Size = new System.Drawing.Size(125, 29);
            this.m_pbCleanDVDs.TabIndex = 34;
            this.m_pbCleanDVDs.Text = "Clean DVDs";
            this.m_pbCleanDVDs.Click += new System.EventHandler(this.DoDvdsCleanup);
            // 
            // m_axBrowser
            // 
            this.m_axBrowser.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.m_axBrowser.Enabled = true;
            this.m_axBrowser.Location = new System.Drawing.Point(16, 101);
            this.m_axBrowser.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("m_axBrowser.OcxState")));
            this.m_axBrowser.Size = new System.Drawing.Size(790, 298);
            this.m_axBrowser.TabIndex = 5;
            this.m_axBrowser.NavigateComplete2 += new AxSHDocVw.DWebBrowserEvents2_NavigateComplete2EventHandler(this.TriggerNavigateComplete);
            this.m_axBrowser.DocumentComplete += new AxSHDocVw.DWebBrowserEvents2_DocumentCompleteEventHandler(this.TriggerDocumentDone);
            // 
            // label12
            // 
            label12.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            label12.Location = new System.Drawing.Point(586, 41);
            label12.Name = "label12";
            label12.Size = new System.Drawing.Size(71, 24);
            label12.TabIndex = 35;
            label12.Text = "Log File:";
            // 
            // m_ebLogFile
            // 
            this.m_ebLogFile.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.m_ebLogFile.Location = new System.Drawing.Point(663, 38);
            this.m_ebLogFile.Name = "m_ebLogFile";
            this.m_ebLogFile.Size = new System.Drawing.Size(143, 27);
            this.m_ebLogFile.TabIndex = 36;
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
            this.ClientSize = new System.Drawing.Size(1016, 684);
            this.Controls.Add(this.m_ebLogFile);
            this.Controls.Add(label12);
            this.Controls.Add(this.m_pbCleanDVDs);
            this.Controls.Add(this.m_pbUpdateDvd);
            this.Controls.Add(this.m_cbScanningTabOrder);
            this.Controls.Add(this.m_pbDrinkWine);
            this.Controls.Add(this.m_ebNotes);
            this.Controls.Add(this.m_lblNotes);
            this.Controls.Add(this.m_ebVintage);
            this.Controls.Add(this.m_ebWine);
            this.Controls.Add(this.m_lblWine);
            this.Controls.Add(this.m_lblWineHeader);
            this.Controls.Add(this.m_pbGeneric);
            this.Controls.Add(label11);
            this.Controls.Add(this.m_recStatus);
            this.Controls.Add(label10);
            this.Controls.Add(label8);
            this.Controls.Add(label9);
            this.Controls.Add(this.m_ebLongDesc);
            this.Controls.Add(this.m_pbWine);
            this.Controls.Add(this.m_ebShortDesc);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.m_pbDvd);
            this.Controls.Add(this.m_ebManualCode);
            this.Controls.Add(label6);
            this.Controls.Add(this.m_pbNewScan);
            this.Controls.Add(label4);
            this.Controls.Add(label5);
            this.Controls.Add(this.m_cbxAddAs);
            this.Controls.Add(this.m_sb);
            this.Controls.Add(label3);
            this.Controls.Add(this.m_ebLocation);
            this.Controls.Add(label2);
            this.Controls.Add(label1);
            this.Controls.Add(this.m_ebTitle);
            this.Controls.Add(this.m_pbManualAdd);
            this.Controls.Add(this.m_axBrowser);
            this.Controls.Add(this.m_ebIsbn);
            this.Controls.Add(this.m_pbBooks);
            this.Controls.Add(this.pictureBox1);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Inventory by ScanCode";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.m_axBrowser)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
		    string s = System.Configuration.ConfigurationManager.AppSettings["sleepTime"];

			Application.Run(new Form1());
		}
#if nosnd		
		public class Win32
		{
			[DllImport("winmm.dll", EntryPoint="PlaySound",CharSet=CharSet.Auto)]
			public static extern int PlaySound(string pszSound, int hmod, int falgs);
			public enum SND
			{
			SND_SYNC         = 0x0000  ,/* play synchronously (default) */
			SND_ASYNC        = 0x0001 , /* play asynchronously */
			SND_NODEFAULT    = 0x0002 , /* silence (!default) if sound not found */
			SND_MEMORY       = 0x0004 , /* pszSound points to a memory file */
			SND_LOOP         = 0x0008 , /* loop the sound until next sndPlaySound */
			SND_NOSTOP       = 0x0010 , /* don't stop any currently playing sound */
			SND_NOWAIT       = 0x00002000, /* don't wait if the driver is busy */
			SND_ALIAS        = 0x00010000 ,/* name is a registry alias */
			SND_ALIAS_ID     = 0x00110000, /* alias is a pre d ID */
			SND_FILENAME     = 0x00020000, /* name is file name */
			SND_RESOURCE     = 0x00040004, /* name is resource name or atom */
			SND_PURGE        = 0x0040,  /* purge non-static events for task */
			SND_APPLICATION  = 0x0080 /* look for application specific association */
			};
		};
#endif		
		ADAS Adas()
		{
			if (String.Compare(m_cbxAddAs.Text, "Book", true/*ignoreCase*/) == 0)
				return ADAS.Book;
			else if (String.Compare(m_cbxAddAs.Text, "DVD", true/*ignoreCase*/) == 0)
				return ADAS.DVD;
			else if (String.Compare(m_cbxAddAs.Text, "Generic", true/*ignoreCase*/) == 0)
				return ADAS.Generic;
			else if (String.Compare(m_cbxAddAs.Text, "Wine", true/*ignoreCase*/) == 0)
				return ADAS.Wine;
			else 
				return ADAS.Unknown;
		}
		
		string SMakeCreateSp(string sScan, string sTitle, string sLocation)
		{
			string sNow = DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss");
			sTitle = MakeStringSqlFriendly(sTitle);
			
			switch (Adas())
				{
				case ADAS.Book:
                    return String.Format("sp_createbook '{0}', '{1}', '{2}', '{3}', '{4}'", sScan, MakeStringSqlFriendly(sTitle), sNow, sNow, sLocation);
				case ADAS.DVD:
                    return String.Format("sp_createdvd '{0}', '{1}', '{2}', '{3}', 'D'", sScan, MakeStringSqlFriendly(sTitle), sNow, sNow);
				case ADAS.Generic:
                    return String.Format("sp_createscan '{0}', '{1}', '{2}', '{3}'", sScan, MakeStringSqlFriendly(sTitle), sNow, sNow);
				}
			throw(new Exception("Invalid ADAS in SMakeCreateSp"));
		}

		string SMakeTitleQuery(string sScan)
		{
			switch (Adas())
				{
				case ADAS.Book:
					return String.Format("SELECT Title From upc_Books WHERE ScanCode = '{0}'", sScan);
				case ADAS.DVD:
                    return String.Format("SELECT Title From upc_DVD WHERE ScanCode = '{0}'", sScan);
				case ADAS.Wine:
                    return String.Format("SELECT Wine From upc_Wines WHERE ScanCode = '{0}'", sScan);
				case ADAS.Generic:
                    return String.Format("SELECT DescriptionShort From upc_Codes WHERE ScanCode = '{0}'", sScan);
				}
			throw(new Exception("Invalid ADAS in SMakeTitleQuery"));
		}

		void AddToFile()
		{
			string sNow = DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt");
			SetupConnection();
			string sLocation = null;
			if (m_ebLocation.Text.Length > 0)
				sLocation = m_ebLocation.Text;
				
			OleDbCommand oCmd = new OleDbCommand(SMakeCreateSp(m_ebIsbn.Text, m_ebTitle.Text, sLocation), oConn);
			oCmd.ExecuteNonQuery();
			UpdateStatus("Added "+ m_ebTitle.Text + " successfully!");
		}
		
		OleDbConnection oConn = null;
		
		void SetupConnection()
		{
			if (oConn != null)
				return;
				
			try
				{
				oConn = new OleDbConnection(sConnString); // "Provider=SQLOLEDB;Trusted_Connection=yes;server=cacofonix;Database=UPC");
				oConn.Open();
				}
			catch (Exception e)
				{
				MessageBox.Show(e.ToString(), "UPC");
				throw(new Exception("Could not setup ODBC connection", e));
				}
		}

		string MakeStringSqlFriendly(string s)
		{
			int i = 0;
			int iLast = 0;
			string sRet = "";
			
			while ((i = s.IndexOf('\'', iLast)) != -1)
				{
				sRet = sRet + s.Substring(iLast, i - iLast + 1) + '\'';
				iLast = i + 1;
				}
			sRet = sRet + s.Substring(iLast, s.Length - iLast);
			
			return sRet;
		}

        enum AlertType
            {
            GoodInfo,
            BadInfo,
            Halt,
            Duplicate,
			Drink
            };

                
		void DoAlert(AlertType at, string s)
		{
		    // SystemSound ssnd;
		    
            // string sSound;
            switch (at)
                {
                case AlertType.GoodInfo:
                    SystemSounds.Exclamation.Play();
                    // sSound = "SystemExclamation";
                    break;
                case AlertType.Duplicate:
                    m_snd.Stream = Resource1.doh; // m_rm.GetStream("doh.wav");
                    m_snd.PlaySync();
                    break;
				case AlertType.Drink:
					m_snd.Stream = Resource1.hicup_392; // m_rm.GetStream("doh.wav");
					m_snd.PlaySync();
					break;
                default:
                    SystemSounds.Hand.Play();
                    // sSound = "SystemHand";
                    break;
                }

           
       //      Win32.PlaySound(sSound, 0, (int)(Win32.SND.SND_ASYNC | Win32.SND.SND_ALIAS | Win32.SND.SND_NOWAIT));
            UpdateStatus(s);
        }

	    private string sAccessKey = null;
        const string sRequestTemplate = "http://isbndb.com/api/books.xml?access_key={0}&index1=isbn&value1={1}";

		string SLookupBookUpc(string sUpc)
		{

            string sIsbn;
            string sTitle = "!!NO TITLE FOUND";

            sIsbn = sUpc; // SConvertUpcToIsbn(sUpc);

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(String.Format(sRequestTemplate, sAccessKey, sIsbn));
            if (req != null)
                {
                HttpWebResponse resp = (HttpWebResponse)req.GetResponse();

                if (resp != null)
                    {
                    Stream stm = resp.GetResponseStream();
                    if (stm != null)
                        {
                        System.Xml.XmlDocument dom = new System.Xml.XmlDocument();

                        try
                            {
                            dom.Load(stm);

                            XmlNode node = dom.SelectSingleNode("/ISBNdb/BookList/BookData/Title");
                            if (node == null)
                                {
                                // try again scraping from bn.com...this is notoriously fragile, so its our last resort.
                                sTitle = SScrapeISBN(sIsbn);
                                }
                            else
                                {
                                sTitle = node.InnerText;
                                }
                            }
                        catch (Exception exc)
                            {
                            sTitle = "!!NO TITLE FOUND: (" + exc.Message + ")";
                            }
                        }
                    }
                }

            return sTitle;
		}

        private object NavigateBrowser(string sUrl)
        {
            object Zero = null;
            object EmptyString = "";
            fNavDone = false;

            m_axBrowser.Visible = true;
            //m_axBrowser.
            m_axBrowser.Navigate(sUrl, ref Zero, ref EmptyString, ref EmptyString, ref EmptyString);
            long s = 0;

            while (s < 20 && !fNavDone) // && m_axBrowser.Busy)
            {
                Application.DoEvents();
                System.Threading.Thread.Sleep(1000);
                s++;
            }

            if (s > 20000 || (!fNavDone && m_axBrowser.Busy))
            {
                m_axBrowser.Stop();
                //	m_axBrowser.Visible = m_cbShowBrowser.Checked;
                UpdateStatus("Browser timed out!");
                return null;
            }
            return m_axBrowser.Document;
        }

	    private string SScrapeISBN(string sIsbn)
	    {
	        string sUrl = "http://search.barnesandnoble.com/bookSearch/isbnInquiry.asp?isbn=" + sIsbn;

	        object oDocBrowser = NavigateBrowser(sUrl);
	        if (oDocBrowser == null)
	            return "!!NO TITLE FOUND!!";

	        IHTMLDocument2 oDoc = (IHTMLDocument2) oDocBrowser;
	        IHTMLDocument3 oDoc2 = (IHTMLDocument3) oDocBrowser;
	        // scrape the book info
	        string sTitle = oDoc.title;
	        string sTitle2 = null;

            if (sTitle.IndexOf(sIsbn) > 0 || sTitle.IndexOf("|") > 0)
                {
//	        if (sTitle.IndexOf("-") > 0)
//	            {
	            sTitle2 = sTitle.Substring(sTitle.IndexOf("|") + 2);
	            if ((sTitle2.IndexOf("Book") >= 0 && sTitle2.IndexOf("Search") >= 0))
	                sTitle2 = null;
	            else
	                {
	                try
	                    {
	                    IHTMLElementCollection hec = (IHTMLElementCollection) oDoc.all.tags("section");

	                    IHTMLElement ihe1 = (IHTMLElement) hec.item("prodSummary", 0);
	                    IHTMLElementCollection hec2 = (IHTMLElementCollection) ihe1.all;
	                    IHTMLElementCollection hec3 = (IHTMLElementCollection) hec2.tags("H1");
	                    IHTMLElement ihe2 = (IHTMLElement) hec3.item(null, 0);
	                    sTitle2 = ihe2.innerText;

	                    // now look for some telltale endings of the title...
	                    string sNew;
	                    if (!UpcUtils.FSanitizeStringCore(sTitle2, " \r", false, false, out sNew))
	                        {
	                        sNew = UpcUtils.SanitizeStringCore(sTitle2, "by", true, false);
	                        }
	                    sTitle2 = sNew.Trim();
	                    sTitle = sTitle2;
//                            sTitle = ((IHTMLElement)((IHTMLElement)hec.item("product-info", 0)).children
	                    //("h2").item(0)).innerText;
	                    // sTitle = ((IHTMLElement)hec.item("title", 0)).innerText;
	                    }
	                catch
	                    {
	                    sTitle2 = null;
	                    sTitle = null;
	                    m_ebTitle.Text = "!!NO TITLE FOUND!!";
	                    SetIsbnFocus();
	                    }
	                }
	            }
	        if (sTitle2 == null)
	            {
	            DoAlert(AlertType.BadInfo, "Cannot add book!  ISBN not found!!");
	            SetIsbnFocus();
	            return "!!NO TITLE FOUND!!";
	            }
	        return sTitle;
	    }

        string SLookupUpcGenericUpcDatabase()
		{
            string sTitle = null;

			string sUrl = "http://www.upcdatabase.com/item.asp?upc="+m_ebIsbn.Text;

			object oDocBrowser = NavigateBrowser(sUrl);
            if (oDocBrowser == null)
                return "!!NO TITLE FOUND!!";

		    IHTMLDocument2 oDoc = (IHTMLDocument2)oDocBrowser;
			IHTMLDocument3 oDoc2 = (IHTMLDocument3)oDocBrowser;

#if SCRAPE_BOOK
			// ok, split off...
			if (adas == ADAS.Book)
				{

			else
#endif // SCRAPE_BOOK
			// scrape from the UPC
			// find the TR that starts with Description
			try
				{
				IHTMLElementCollection hec = (IHTMLElementCollection)oDoc.all.tags("tr");
				foreach (IHTMLElement he in hec)
					{
					if (he.innerText.Length > 11 && String.Compare(he.innerText.Substring(0, 11), "Description") == 0)
						{
						// got it!
						sTitle = he.innerText.Substring(11);
						sTitle = UpcUtils.SanitizeString(sTitle);
						break;
						}
					}
				if (sTitle == null)
					throw(new Exception("not found"));
				}
			catch
				{
				sTitle = "!!NO TITLE FOUND!!";
				}

			return sTitle;
		}

	    string SLookupUpcGeneric() // searchupc.com
		{
            string sTitle = null;

			string sUrl = "http://www.searchupc.com/upc/"+m_ebIsbn.Text;

			object oDocBrowser = NavigateBrowser(sUrl);
            if (oDocBrowser == null)
                return "!!NO TITLE FOUND!!";

		    IHTMLDocument2 oDoc = (IHTMLDocument2)oDocBrowser;
			IHTMLDocument3 oDoc2 = (IHTMLDocument3)oDocBrowser;

            // grab the teable with id= 'searchresultdata'
	        IHTMLElement ihe = (IHTMLElement) oDoc2.getElementById("searchresultdata");
	        IHTMLElementCollection ihecRows;

	        if (ihe == null || (ihecRows = (IHTMLElementCollection)((IHTMLElementCollection)ihe.all).tags("tr")).length != 2)
	            return "!!NO TITLE FOUND";


			// scrape from the UPC
			// find the TR that starts with Description
			try
				{
    	        IHTMLTableRow iheRow = (IHTMLTableRow)ihecRows.item(1);


				sTitle = ((IHTMLElement) iheRow.cells.item(0)).innerText;
				sTitle = UpcUtils.SanitizeString(sTitle);
				if (sTitle == null)
					throw(new Exception("not found"));
				}
			catch
				{
				sTitle = "!!NO TITLE FOUND!!";
				}

			return sTitle;
		}

	    void LogToFile(string s)
	    {
	        if (String.IsNullOrEmpty(m_ebLogFile.Text))
	            return;

	        StreamWriter sw = new StreamWriter(m_ebLogFile.Text, true /*fAppend*/, System.Text.Encoding.Default);

	        sw.WriteLine(s);
	        sw.Close();
	    }
	    bool FSyncUpc()
	    {
	        // before we go looking around on the internet, let's first check to see if our database has
	        // this in it already...
	        SetupConnection();
	        string sNow = DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt");
	        OleDbCommand oCmd = new OleDbCommand(SMakeTitleQuery(m_ebIsbn.Text), oConn);
	        string sTitle = (string) oCmd.ExecuteScalar();

	        if (sTitle != null)
	            {
#if bar
                Barcodes.Barcode.BarcodeWebControl bwc = new Barcodes.Barcode.BarcodeWebControl();

                bwc.BarText = m_ebIsbn.Text;
                bwc.BarType = Barcodes.Barcode.BarcodeTypes.Code93;

                Graphics gr = m_pbxScan.CreateGraphics();

                bwc.DrawBarcodeToSize(0.0, 0.0, m_pbxScan.Width, m_pbxScan.Height, Barcodes.Barcode.Dimensions.dmPixels, gr);
#endif
	            // found the scancode -- update the inventory
	            CheckForDupe(m_ebIsbn.Text);
	            m_ebTitle.Text = sTitle;
	            if (Adas() == ADAS.Book && m_ebLocation.Text.Length > 0 && String.Compare(m_ebLocation.Text, "!? LOC ?!") != 0)
	                {
	                oCmd = new OleDbCommand(
	                    String.Format("sp_updatebookscanlocation '{0}', '{1}', '{2}', '{3}'", m_ebIsbn.Text, MakeStringSqlFriendly(sTitle), sNow, m_ebLocation.Text), oConn);
	                }
	            else
	                {
	                oCmd = new OleDbCommand(String.Format("sp_updatescan '{0}', '{1}', '{2}'", m_ebIsbn.Text, MakeStringSqlFriendly(sTitle), sNow), oConn);
	                }
	            oCmd.ExecuteNonQuery();
	            DoAlert(AlertType.GoodInfo, String.Format("Found {0}, updated last scan to {1}", sTitle, sNow));
	            LogToFile(m_ebIsbn.Text);
	            if (Adas() == ADAS.Wine)
	                {
	                oCmd = new OleDbCommand(String.Format("select Notes from upc_Wines WHERE ScanCode = '{0}'", m_ebIsbn.Text), oConn);
	                object o = oCmd.ExecuteScalar();
	                if (o == null || System.DBNull.Value != o)
	                    m_ebNotes.Text = (string) o;

	                if (m_cbScanningTabOrder.Checked)
	                    SetFocus(m_ebIsbn);
	                else
	                    MoveToNotes();
	                return false;
	                }
	            return true;
	            }

	        ADAS adas = Adas();

	        if (adas == ADAS.Wine)
	            {
	            MoveToWine();
	            return true;
	            }

//			if (Adas() == ADAS.Book)
	        // drat -- didn't find the scancode -- look up as an ISBN or a UPC...

	        if (adas == ADAS.Book)
	            {
	            m_ebTitle.Text = SLookupBookUpc(m_ebIsbn.Text);
	            }
	        else
	            {
	            m_ebTitle.Text = SLookupUpcGeneric();
	            }
	        if (!m_ebTitle.Text.StartsWith("!!"))
	            sTitle = m_ebTitle.Text;

	        // fallthrough to add
	        if (adas == ADAS.Book && String.Compare(m_ebLocation.Text, "!? LOC ?!") == 0)
	            {
	            SetFocus(m_ebLocation);
	            DoAlert(AlertType.BadInfo, "Cannot add book!  Location not set!");
	            return false;
	            }

	        if (sTitle == null)
	            {
	            DoAlert(AlertType.BadInfo, "Scan code not found!  Enter Title manually!");
	            return false;
	            }
	        m_ebTitle.Text = sTitle;
	        DoAlert(AlertType.GoodInfo, "");
	        AddToFile();
//			}
//		else
	        //{
	        //Win32.PlaySound("SystemHand", 0, (int)(Win32.SND.SND_ASYNC | Win32.SND.SND_ALIAS | Win32.SND.SND_NOWAIT));
	        //UpdateStatus("Scan code not found!  Enter Title manually!");
	        //m_ebTitle.Text = "!!NO TITLE FOUND!!";
	        //}

	        return true;
	    }

	    void CheckForDupe(string sCode)
		{
			if (Adas() == ADAS.Wine)
				return;
		    // query the database for the last scan date
            string sCmd = String.Format("select ScanCode, LastScanDate FROM upc_Codes where ScanCode='{0}'", sCode);
            SetupConnection();

            OleDbCommand oCmd = new OleDbCommand(sCmd, oConn);
            OleDbDataReader ord = oCmd.ExecuteReader();
            
            if (ord.HasRows)
                {
                ord.Read();
                DateTime dttm = ord.GetDateTime(1);
                
                if (DateTime.Now.AddDays(-2).CompareTo(dttm) <= 0)
                    {
                    DoAlert(AlertType.Duplicate, String.Format("DUPLICATE:  Last ScanDate: {0}", dttm.ToString()));
                    }
                }
            ord.Close();
        }
		
		bool fNavDone;
		private void TriggerDocumentDone(object sender, AxSHDocVw.DWebBrowserEvents2_DocumentCompleteEvent e)
		{
			fNavDone = true;
		}

        private void TriggerNavigateComplete(object sender, AxSHDocVw.DWebBrowserEvents2_NavigateComplete2Event e)
        {
           // fNavDone = true;
        }

		private void m_ebIsbn_TextChanged(object sender, System.EventArgs e) {
			FSyncUpc();
		}

		private void m_ebIsbn_KeyUp(object sender, System.Windows.Forms.KeyPressEventArgs e) {
			if (e.KeyChar == 13)
				{
				// if FSyncUpc fails, then it will set focus appropriately
				if (FSyncUpc())
					{
                    SetIsbnFocus();
					}
				e.Handled = true;
				}
		}

		void UpdateStatus(string s)
		{
			m_sb.Text = s;
            m_srpt.AddMessage(s);
		}

        void SetFocus(TextBox c)
        {
            c.Focus();
            c.SelectAll();
        }

        void SetIsbnFocus()
        {
            SetFocus(m_ebIsbn);
        }

		private void ChangeAddAs(object sender, System.EventArgs e) 
		{
            SetFocus(m_ebIsbn);
        }

        string SConvertUpcToIsbn(string sUpc)
        {
            if (sUpc.Length != 13 || !sUpc.StartsWith("978"))
                return sUpc;

            string isbn = sUpc.Substring(3,9);

            int xsum = 0;
            
            for (int i = 0; i < 9; i++)
                {
                xsum += (10 - i) * isbn[i] - '0';
                }
                
            xsum %= 11;
            xsum = 11 - xsum;
            
            string x_val = String.Format("{0}", xsum);
            
            switch (xsum)
                {
                case 10: x_val = "X"; break;
                case 11: x_val = "0"; break;
                }
            isbn += x_val;

            return isbn;
        }

        string SCheckCalcIsbn10(string sIsbn10)
        {
            if (sIsbn10.Length != 9)
                return null;
                
            Int32 n = 0;
            Int32 i;
            
            for (i = 0; i < 9; i++)
                {
//                n += ((10 - i) * Int32.Parse(sIsbn10.Substring(i, 1)));
                n += ((i + 1) * Int32.Parse(sIsbn10.Substring(i, 1)));
            }
            n = n % 11;
//            n = 11 - n;
            if (n == 10)
                return "X";
            else
                return String.Format("{0}", n);
        
        }

        string SCheckCalcIsbn13(string sIsbn13)
        {
            if (sIsbn13.Length != 12)
                return null;

            Int32 n = 0;
            Int32 i;

            for (i = 0; i < 12; i++)
                {   
                int nDigit = Int32.Parse(sIsbn13.Substring(i, 1));
                if (i % 2 == 0)
                    { // get a weight of 1
                    n += nDigit;
                    }
                else
                    {
                    n += nDigit * 3;
                    }
                }
            n = n % 10;
            n = 10 - n;
            n = n % 10;
            return String.Format("{0}", n);
        }
        
        private void m_pbNewScan_Click(object sender, EventArgs e)
        {
            string sCode = "null";
            // create a new scan code and associate it...
            if (m_ebShortDesc.Text.Substring(0, 2) == "--")
                {
                DoAlert(AlertType.BadInfo, "MUST enter a short description!");
                SetFocus(m_ebShortDesc);
                return;
                }
            if (m_ebLongDesc.Text.Substring(0, 2) == "--")
                {
                DoAlert(AlertType.BadInfo, "MUST enter a long description!");
                SetFocus(m_ebLongDesc);
                return;
                }

            if (m_ebManualCode.Text.Substring(0, 2) != "--")
                {
                if (Adas() == ADAS.Book)
                    {
                    string sIsbn13;
                    
                    if (m_ebManualCode.Text.Length == 13)
                        {
                        // already got the isbn13; just check the check code
                        if (m_ebManualCode.Text.Substring(12, 1) != SCheckCalcIsbn13(m_ebManualCode.Text.Substring(0, 12)))
                            {
                            DoAlert(AlertType.BadInfo, String.Format("ISBN13 Check value incorrect:  {0} != {1}", m_ebManualCode.Text.Substring(12, 1), SCheckCalcIsbn13(m_ebManualCode.Text.Substring(0,12))));
                            SetFocus(m_ebManualCode);
                            return;
                            }
                        sIsbn13 = m_ebManualCode.Text;
                        }
                    else
                        {    
                        // Check the ISBN
                        if (m_ebManualCode.Text.Length != 10)
                            {
                            DoAlert(AlertType.BadInfo, "ISBN Number must be 10 digits!");
                            SetFocus(m_ebManualCode);
                            return;
                            }
                        
                        if (m_ebManualCode.Text.Substring(9, 1) != SCheckCalcIsbn10(m_ebManualCode.Text.Substring(0, 9)))
                            {
                            DoAlert(AlertType.BadInfo, String.Format("ISBN Check value incorrect:  {0} != {1}", m_ebManualCode.Text.Substring(9, 1), SCheckCalcIsbn10(m_ebManualCode.Text.Substring(0,9))));
                            SetFocus(m_ebManualCode);
                            return;
                            }
                        // cool, now make a new isbn-13
                        sIsbn13 = "978" + m_ebManualCode.Text.Substring(0, 9);
                        sIsbn13 += SCheckCalcIsbn13(sIsbn13);
                        }
                    m_ebIsbn.Text = sIsbn13;
                    sCode = "'" + sIsbn13 + "'";
    //                    return;
                        }
                }

            // let's figure out the code to use...
			string sTitle = MakeStringSqlFriendly(m_ebLongDesc.Text);
            string sShortDesc = MakeStringSqlFriendly(m_ebShortDesc.Text);

            string sCmd = String.Format("sp_createcustomscan '{0}', '{1}', '03%', {2}, ? OUTPUT", sShortDesc, sTitle, sCode);
			SetupConnection();
				
			OleDbCommand oCmd = new OleDbCommand(sCmd, oConn);
            OleDbParameter odp = new OleDbParameter();
            odp.Direction = ParameterDirection.Output;
            odp.Size = 14;
            oCmd.Parameters.Add(odp);

			oCmd.ExecuteNonQuery();
            m_ebIsbn.Text = odp.Value.ToString();
            m_ebTitle.Text = sTitle;
            m_ebLongDesc.Text = "--"+m_ebLongDesc.Text;
            m_ebShortDesc.Text = "--" + m_ebShortDesc.Text;
            SetFocus(m_ebIsbn);
        }

        private void Control_OnClick(object sender, EventArgs e)
        {
            if ((int)((Control)sender).Tag != 0)
                {
                ((TextBox)sender).SelectAll();
                ((Control)sender).Tag = 0;
                }
        }

        private void Control_OnEnter(object sender, EventArgs e)
        {
            ((Control)sender).Tag = 1;
        }

        private void ManualAdd_Click(object sender, EventArgs e)
        {
            DoAlert(AlertType.GoodInfo, "");
            AddToFile();
            SetIsbnFocus();
        }

    
        private void ChangeInput(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            m_cbxAddAs.SelectedIndex = Int32.Parse((string)btn.Tag);
            m_pbBooks.Invalidate();
            m_pbDvd.Invalidate();
            m_pbWine.Invalidate();
			ShowHideWine();
			SetFocus(m_ebIsbn);
        }


        Bitmap[] m_rgbmp;

        void EnsureResourcesLoaded()
		{
			if (m_rgbmp == null)
				{
				m_rgbmp = new Bitmap[4];

				m_rgbmp[0] = new Bitmap(this.GetType(), "Resources.Books2.png");
                m_rgbmp[1] = new Bitmap(this.GetType(), "Resources.dvds3.png");
                m_rgbmp[2] = new Bitmap(this.GetType(), "Resources.wine1.png");
				m_rgbmp[3] = new Bitmap(this.GetType(), "Resources.upc1.png");
                }
		}

		private void EH_RenderHeadingLine(object sender, System.Windows.Forms.PaintEventArgs e)
		{
			Label lbl = (Label)sender;
			string s = (string)lbl.Text;

			SizeF sf = e.Graphics.MeasureString(s, lbl.Font);
			int nWidth = (int)sf.Width;
			int nHeight = (int)sf.Height;

//			e.Graphics.DrawString(s, lbl.Font, new SolidBrush(Color.SlateBlue), 0, 0);// new System.Drawing.Point(0, (lbl.Width - nWidth) / 2));
			e.Graphics.DrawLine(new Pen(new SolidBrush(lbl.ForeColor), 0.5f), 2 + nWidth + 1, (nHeight / 2), lbl.Width, (nHeight / 2));

		}

		void MoveToNotes()
		{
			// get ready to input tasting notes
			m_ebWine.Text = m_ebTitle.Text;
            OleDbCommand oCmd = new OleDbCommand(String.Format("select Vintage from upc_Wines WHERE ScanCode = '{0}'", m_ebIsbn.Text), oConn);
			m_ebVintage.Text = (string)oCmd.ExecuteScalar();
			m_ebNotes.Focus();
		}

        void MoveToWine()
        {
            // nothing in inventory for this -- let them enter all the info...
            m_ebWine.Text = "-- No inventory for this wine.  Please enter a brief title";
            m_ebVintage.Text = "--Year";
            m_ebWine.Focus();
        }
        
		private void ShowHideWine()
		{
			bool fShowWine;

			if (Adas() == ADAS.Wine)
				fShowWine = true;
			else
				fShowWine = false;

		    m_cbScanningTabOrder.Visible = fShowWine;
			m_axBrowser.Visible = !fShowWine;
			m_lblWineHeader.Visible = fShowWine;
			m_lblWine.Visible = fShowWine;
			m_ebWine.Visible = m_ebWine.Enabled = fShowWine;
			m_ebVintage.Visible = m_ebVintage.Enabled = fShowWine;
			m_lblNotes.Visible = fShowWine;
			m_ebNotes.Visible = m_ebNotes.Enabled = fShowWine;
			m_pbDrinkWine.Visible = m_pbDrinkWine.Enabled = fShowWine;
		}


        private void PaintButton(object sender, PaintEventArgs e)
        {
            Button btn = (Button)sender;
			int iButton = Int32.Parse((string)btn.Tag);

            if (m_cbxAddAs.SelectedIndex == iButton)
				{
                ControlPaint.DrawButton(e.Graphics, e.ClipRectangle, ButtonState.Pushed);
				}
            else
				{
				ControlPaint.DrawButton(e.Graphics, e.ClipRectangle, ButtonState.Normal);
				}
		    Rectangle rc = btn.ClientRectangle;
		    
		    rc.Offset(new Point(5, 10));
		    rc.Width -= 20;
		    rc.Height -= 40;

			if (iButton < m_rgbmp.Length)
			    {
				e.Graphics.DrawImage(m_rgbmp[iButton], rc);
				}
			
			Font fnt = new Font("Tahoma", 12.0f, FontStyle.Bold);
			int dxText = (int)e.Graphics.MeasureString(btn.Text, fnt).Width;
            int dyText = (int)e.Graphics.MeasureString(btn.Text, fnt).Height;
			e.Graphics.DrawString(btn.Text, fnt, Brushes.Black, new Point((btn.ClientRectangle.Width - dxText) / 2, (btn.ClientRectangle.Height - rc.Bottom - dyText) / 2 + rc.Bottom));
        }

		private void NotesKeyUp(object sender, System.Windows.Forms.KeyPressEventArgs e) {
			if (e.KeyChar == 13)
				{
				DrinkWine();
				e.Handled = true;
				}
		}
        private void DrinkWine(object sender, EventArgs e)
		{
			DrinkWine();
		}

		private void DrinkWine()
        {
			string sNow = DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss");

            string sCmd = String.Format("sp_drinkwine '{0}', '{1}', '{2}', '{3}', '{4}'", m_ebIsbn.Text, MakeStringSqlFriendly(m_ebWine.Text), MakeStringSqlFriendly(m_ebVintage.Text), MakeStringSqlFriendly(m_ebNotes.Text), sNow);

			OleDbCommand oCmd = new OleDbCommand(sCmd, oConn);
			oCmd.ExecuteNonQuery();
			DoAlert(AlertType.Drink, "");
			UpdateStatus("Drank Wine:  " + m_ebTitle.Text + " [ " + m_ebNotes.Text + " ] Done!");
			m_ebWine.Text = "--" + m_ebWine.Text;

			SetFocus(m_ebIsbn);
        }

	    class DvdElement
	    {
	        string m_sScanCode;
	        string m_sTitle;
	        string m_sSummary;
            private string m_sNotes;
            private string m_sQueryUrl;
	        private string m_sCoverSrc;
	        private string m_sClassification;
	        private string m_sMediaType;
	        private List<string> m_plsClasses;

	        public DvdElement(OleDbDataReader odr)
	        {
	            m_sScanCode = odr.IsDBNull(0) ? "" : odr.GetString(0);
	            m_sTitle = odr.IsDBNull(1) ? "" : odr.GetString(1);
	            m_sSummary = odr.IsDBNull(2) ? "" : odr.GetString(2);
	            m_sNotes = odr.IsDBNull(3) ? "" : odr.GetString(3);
                m_sQueryUrl = odr.IsDBNull(4) ? "" : odr.GetString(4);
                m_sCoverSrc = odr.IsDBNull(5) ? "" : odr.GetString(5);
                m_sClassification = odr.IsDBNull(6) ? "" : odr.GetString(6);
                m_sMediaType = odr.IsDBNull(7) ? "" : odr.GetString(7);
            }

	        public DvdElement(string sScanCode)
	        {
	            m_sScanCode = sScanCode;
	            m_sTitle = m_sSummary = "";
	        }

	        public string Notes
	        {
	            get { return m_sNotes; }
	            set { m_sNotes = value; }
	        }

	        public List<string> ClassList
            {
                get { return m_plsClasses; }
                set { m_plsClasses = value; }
            }
            public string ScanCode
            {
                get { return m_sScanCode; }
                set { m_sScanCode = value; }
            }
            public string MediaType
            {
                get { return m_sMediaType; }
                set { m_sMediaType = value; }
            }
            public string QueryUrl
            {
                get { return m_sQueryUrl; }
                set { m_sQueryUrl = value; }
            }
            public string CoverSrc
            {
                get { return m_sCoverSrc; }
                set { m_sCoverSrc = value; }
            }
            public string Classification
            {
                get { return m_sClassification; }
                set { m_sClassification = value; }
            }
            public string Summary
            {
                get { return m_sSummary; }
                set { m_sSummary = value; }
            }
            public string Title
            {
                get { return m_sTitle; }
                set { m_sTitle = value; }
            }
        }

	    private WebControl m_wc;

	    void NavToScrapePage(string sHost)
	    {
	        if (!m_wc.FNavToPage(sHost))
	            throw (new Exception($"could not navigate to {sHost}"));
	    }

	    void EnsureScrapeSetup(bool fShow)
	    {
            if (m_wc == null)
                m_wc = new WebControl(m_srpt);

            if (fShow)
                m_wc.Show();
        }

	    bool FScrapeDvd(ref DvdElement dvd)
	    {
	        if (string.IsNullOrEmpty(dvd.ScanCode))
	            return false;

	        EnsureScrapeSetup(true);
            IHTMLDocument2 doc = m_wc.Document2;

            // if (!WebControl.FCheckForControl(doc, "searchBarBN"))
                NavToScrapePage("http://barnesandnoble.com");

            doc = m_wc.Document2;
            if (!WebControl.FCheckForControl(doc, "searchBarBN"))
	            throw (new Exception("could not find search control on scrape page"));

	        if (!WebControl.FSetInputControlText(doc, null, "searchBarBN", dvd.ScanCode, true))
	            throw (new Exception("could not set search control text"));

	        m_wc.ResetNav();
	        if (!m_wc.FClickControl(doc, "searchSubmit"))
	            throw (new Exception("could not click on search control"));

	        m_wc.ResetNav();
	        if (!m_wc.FWaitForNavFinish("productInfoOverview"))
	            throw (new Exception("could not wait for nav to finish"));

            // at this point, we should have the results page...
	        doc = m_wc.Document2;

            // find the "product info overfiew" div "(id is productInfoOverview)

	        string s = WebControl.SGetElementValue(doc, "productInfoOverview");
	        if (s != null)
	            {
	            dvd.Summary = UpcUtils.SanitizeString(s, false);

	            IHTMLElement ihe = WebControl.ElementGet(doc, "pdpMainImage");
	            if (ihe != null)
	                dvd.CoverSrc = ((IHTMLImgElement) ihe).src;

	            ihe = WebControl.ElementGet(doc, "prodPromo");
	            string sMediaType = null;
	            foreach (IHTMLElement iheInner in (IHTMLElementCollection)ihe.children)
	                {
	                if (String.Compare(iheInner.tagName, "h2", StringComparison.OrdinalIgnoreCase) == 0)
	                    {
	                    sMediaType = iheInner.innerText;
	                    break;
	                    }
	                }

	            dvd.MediaType = UpcUtils.SanitizeMediaType(sMediaType);

	            List<string> pls = WebControl.ListItemsGet(doc, "relatedSubjects");
	            if (pls != null)
	                pls = UpcUtils.SanitizeClassList(pls);

	            dvd.Classification = pls == null ? "" : string.Join(",", pls);
	            dvd.ClassList = pls;

	            dvd.QueryUrl = doc.url;

	            dvd.Title = UpcUtils.SanitizeString(dvd.Title, true);
	            return true;
	            }

	        return false;
	    }

	    private void DownloadImage(ref DvdElement dvd, string sLocalPath)
	    {
	        if (!Directory.Exists(sLocalPath))
	            {
	            Directory.CreateDirectory(sLocalPath);
	            }

	        WebRequest req = WebRequest.Create(dvd.CoverSrc);
	        Uri uri = new Uri(dvd.CoverSrc);

	        string sPathname = uri.GetComponents(UriComponents.Path, UriFormat.SafeUnescaped);
	        string sFilename = Path.GetFileName(sPathname);
	        string sOutPath = $"{sLocalPath}\\{sFilename}";

	        WebResponse response = req.GetResponse();
            
	        Stream stream = response.GetResponseStream();
	        Stream stmOut = new FileStream(sOutPath, FileMode.Create);

	        stream.CopyTo(stmOut);
	        stmOut.Close();
	        stmOut.Dispose();
	        stream.Close();
	        stream.Dispose();
	        response.Close();
	        response.Dispose();

	        dvd.CoverSrc = sOutPath;
	    }

        private void DoUpdateDvdDescriptions(object sender, EventArgs e)
        {
#if test
            DvdElement dvd2 = new DvdElement("0876964007030");
            //DvdElement dvd = new DvdElement("0841887026505");
            //DvdElement dvd = new DvdElement("0027616152527");
            if (FScrapeDvd(ref dvd2))
                {
                DownloadImage(ref dvd2, "dvd_support");
                }
            return;
#endif
            SetupConnection();
            // get the set of DVD's we need to scrape info for
            string sCmd = String.Format("SELECT ScanCode, Title, Summary, Notes, QueryUrl, CoverSrc, Classification, MediaType From upc_DVD WHERE datalength(IsNull(Summary,'')) < 10 AND IsNull(Notes, '')<> 'LC1' AND ScanCode Is Not Null");

            OleDbCommand oCmd = new OleDbCommand(sCmd, oConn);

            OleDbDataReader ord = oCmd.ExecuteReader();

            List<DvdElement> dvdElements = new List<DvdElement>();

            if (ord.HasRows)
                {
                while (ord.Read())
                    if (ord.IsDBNull(2) || ord.GetString(2).Length < 2)
                        dvdElements.Add(new DvdElement(ord));
                }

            // at this point we have a list of dvd's that need to be scraped
            ord.Close();

            HashSet<string> classes = new HashSet<string>();

            // now fill in the search box with the UPC we are interested in
            foreach (DvdElement dvd in dvdElements)
                {
                DvdElement dvdT = dvd;
                if (FScrapeDvd(ref dvdT))
                    {
                    DownloadImage(ref dvdT, "dvd_support");
                    if (dvdT.Notes == "")
                        dvdT.Notes = "LC1";

                    UpdateDVDInfo(dvdT);
                    if (dvdT.ClassList != null)
                        {
                        foreach (string s in dvdT.ClassList)
                            classes.Add(s);
                        }
                    // add dvd;
                    ;
                    }
                else
                    {
                    if (dvdT.Notes == "")
                        {
                        dvdT.Notes = "LC1";
                        string sCmdUpdate = $"UPDATE upc_DVD SET Notes='LC1' WHERE ScanCode='{dvdT.ScanCode}'";

                        OleDbCommand oCmdUpdate = new OleDbCommand(sCmdUpdate, oConn);
                        oCmdUpdate.ExecuteNonQuery();
                    }
                }
                }
            StreamWriter sw = new StreamWriter("c:\\temp\\classes.txt");
            foreach (string s in classes)
                sw.WriteLine(s);
            sw.Close();
            sw.Dispose();

        }

	    private void UpdateDVDInfo(DvdElement dvd)
	    {
	        string sCmdUpdate = $"UPDATE upc_DVD SET Summary='{MakeStringSqlFriendly(dvd.Summary)}', " +
	                            $"CoverSrc='{MakeStringSqlFriendly(dvd.CoverSrc)}', " +
	                            $"Classification='{MakeStringSqlFriendly(dvd.Classification)}', " +
	                            $"MediaType='{MakeStringSqlFriendly(dvd.MediaType)}', " +
	                            $"QueryUrl='{MakeStringSqlFriendly(dvd.QueryUrl)}', " +
	                            $"Notes='{MakeStringSqlFriendly(dvd.Notes)}', " +
	                            $"Title='{MakeStringSqlFriendly(dvd.Title)}' WHERE ScanCode='{dvd.ScanCode}'";

	        OleDbCommand oCmdUpdate = new OleDbCommand(sCmdUpdate, oConn);

	        oCmdUpdate.ExecuteNonQuery();
	    }

	    private void DoDvdsCleanup(object sender, EventArgs e)
        {
            SetupConnection();
            // get the set of DVD's we need to scrape info for
            string sCmd = String.Format("SELECT ScanCode, Title, Summary, Notes, QueryUrl, CoverSrc, Classification, MediaType From upc_DVD");

            OleDbCommand oCmd = new OleDbCommand(sCmd, oConn);

            OleDbDataReader ord = oCmd.ExecuteReader();

            List<DvdElement> dvdElements = new List<DvdElement>();

            if (ord.HasRows)
            {
                while (ord.Read())
                        dvdElements.Add(new DvdElement(ord));
            }

            // at this point we have a list of dvd's that need to be scraped
            ord.Close();

	        foreach (DvdElement dvd in dvdElements)
	            {
	            DvdElement dvdT = dvd;

	            // now, do some various cleanup
	            bool fNeedUpdate = false;

	            string s = UpcUtils.SanitizeString(dvdT.Title, true);

                // cleanup the title
	            if (String.CompareOrdinal(s, dvdT.Title) != 0)
	                {
	                fNeedUpdate = true;
	                dvdT.Title = s;
	                }

	            // download the image
	            if (dvdT.CoverSrc.Contains("http://"))
	                {
	                DownloadImage(ref dvdT, "dvd_support");
	                fNeedUpdate = true;
	                }

	            if (fNeedUpdate)
	                UpdateDVDInfo(dvdT);

            }
        }
    }
}
