using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SHSchool.Data;


namespace UpdateRecordReport
{
    public partial class PrintBranch : BaseForm
    {
        public PrintBranch()
        {
            InitializeComponent();
        }

        private void PrintBranch_Load(object sender, EventArgs e)
        {
            List<SHDeptGroupRecord>  lstdeptgroup = new List<SHDeptGroupRecord>();
            lstdeptgroup = SHDeptGroup.SelectAll();
            lstCourseKind.Items.Clear();
            //var newItem = new ComboBoxItem();
            foreach (SHDeptGroupRecord branch in lstdeptgroup)
            {
                ListViewItem lvi = new ListViewItem(branch.Name);
                lvi.SubItems.Add(branch.ID);

                lstCourseKind.Items.Add(lvi);

            }
            btnPrint.Enabled = false;

        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            string CourseIDs = "";
            for (int i = 0; i < lstCourseKind.Items.Count; i++)
                if (lstCourseKind.Items[i].Checked == true)
                    CourseIDs = CourseIDs + lstCourseKind.Items[i].SubItems[1].Text + ",";
            Form1 form = new Form1(CourseIDs, txtTitle.Text);
            form.ShowDialog();
            this.Close();
        }
        private void lstCourseKind_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnPrint.Enabled = false;
            txtTitle.Text = "";
            for (int i = 0; i < lstCourseKind.Items.Count; i++)
                if (lstCourseKind.Items[i].Checked == true)
                {
                    btnPrint.Enabled = true;
                    txtTitle.Text = lstCourseKind.Items[i].SubItems[0].Text;
                    break;
                }


        }

        private void lstCourseKind_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            btnPrint.Enabled = false;
            txtTitle.Text = "";
            for (int i = 0; i < lstCourseKind.Items.Count; i++)
                if (lstCourseKind.Items[i].Checked == true)
                {
                    btnPrint.Enabled = true;
                    txtTitle.Text = lstCourseKind.Items[i].SubItems[0].Text;
                    break;
                }

        }

        private void txtTitle_TextChanged(object sender, EventArgs e)
        {
            btnPrint.Enabled = false;
            for (int i = 0; i < lstCourseKind.Items.Count; i++)
                if (lstCourseKind.Items[i].Checked == true && txtTitle.Text != "")
                {
                    btnPrint.Enabled = true;
                    break;
                }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
