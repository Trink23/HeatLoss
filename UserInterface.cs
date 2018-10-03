using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HeatLoss
{
    public partial class UserInterface : Form
    {
        public UserInterface()
        {
            InitializeComponent();

            //new ImportExcel().ImportTab(resumeView, "Resume");

        }



        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.button1 = new System.Windows.Forms.Button();
            this.resumeView = new System.Windows.Forms.DataGridView();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.class1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.resumeView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.class1BindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(601, 523);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(136, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Export Excel";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // resumeView
            // 
            this.resumeView.AutoGenerateColumns = false;
            this.resumeView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.resumeView.DataSource = this.class1BindingSource;
            this.resumeView.Location = new System.Drawing.Point(12, 33);
            this.resumeView.Name = "resumeView";
            this.resumeView.Size = new System.Drawing.Size(725, 480);
            this.resumeView.TabIndex = 1;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(11, 523);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(90, 23);
            this.button2.TabIndex = 2;
            this.button2.Text = "Add Room";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(107, 523);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(105, 23);
            this.button3.TabIndex = 3;
            this.button3.Text = "Remove Room";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(497, 523);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(98, 23);
            this.button4.TabIndex = 4;
            this.button4.Text = "Import Excel";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // class1BindingSource
            // 
            this.class1BindingSource.DataSource = typeof(HeatLoss.Heatloss);
            // 
            // UserInterface
            // 
            this.ClientSize = new System.Drawing.Size(749, 557);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.resumeView);
            this.Controls.Add(this.button1);
            this.Name = "UserInterface";
            ((System.ComponentModel.ISupportInitialize)(this.resumeView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.class1BindingSource)).EndInit();
            this.ResumeLayout(false);

        }
        private void button4_Click(object sender, EventArgs e)
        {
            new ImportExcelOverview().ImportExcel(resumeView, "Resume");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();

            AddRoom aRoom = new AddRoom();

            aRoom.Show();

            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in resumeView.SelectedRows)
            {
                if (!row.IsNewRow)
                    resumeView.Rows.Remove(row);
            }
        }
    }
}
