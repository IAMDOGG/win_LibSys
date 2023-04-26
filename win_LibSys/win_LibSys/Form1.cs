using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace win_LibSys
{
    public partial class Form1 : Form
    {
        private OleDbConnection con;
        public Form1()
        {
            InitializeComponent();
            con = new OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = Z:\\QQ134\\win_LibSys\\LibSys.accdb");
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void btn_Add_Click(object sender, EventArgs e)
        {
            con.Open();
            OleDbCommand com = new OleDbCommand("Insert into book values ('" + txtno.Text + "' , '" + txttitle.Text + "' , '" +
                txtauthor.Text + "')", con);
            com.ExecuteNonQuery();

            MessageBox.Show("Succesfully SAVED!", "info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            con.Close();
            loadDatagrid();
        }

        private void btn_Delete_Click(object sender, EventArgs e)
        {
            con.Open();
            string num = txtno.Text;

            DialogResult dr = MessageBox.Show("Are you sure you want to delete this?", "Confirmation Deletion",
              MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (dr == DialogResult.Yes)
            {
                OleDbCommand com = new OleDbCommand("Delete from book where [Accession Number]= '" + num + "'", con);
                com.ExecuteNonQuery();

                MessageBox.Show("Sucessfully DELETED!", "info", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else
            {
                MessageBox.Show("CANCELLED!", "info", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            con.Close();
            loadDatagrid();
        }

        private void btn_Edit_Click(object sender, EventArgs e)
        {
            con.Open();
            string no;
            no = txtno.Text;

            OleDbCommand com = new OleDbCommand("Update book SET title = '" + txttitle.Text + "', author='" +
                txtauthor.Text + "' where accession_number = '" + no + "'", con);
            com.ExecuteNonQuery();

            MessageBox.Show("Sucessfully UPDATED!", "info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            con.Close();
            loadDatagrid();
        }

        private void grid1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            txtno.Text = grid1.Rows[e.RowIndex].Cells["Accession Number"].Value.ToString();
            txttitle.Text = grid1.Rows[e.RowIndex].Cells["Title"].Value.ToString();
            txtauthor.Text = grid1.Rows[e.RowIndex].Cells["Author"].Value.ToString();
        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            con.Open();
            OleDbCommand com = new OleDbCommand("Select * from book where title like '%" + txtsearch.Text + "%'", con); com.ExecuteNonQuery();
            OleDbDataAdapter adap = new OleDbDataAdapter(com);
            DataTable tab = new DataTable();
            adap.Fill(tab);
            grid1.DataSource = tab;
            con.Close();
        }
        private void loadDatagrid()
        {
            con.Open();
            OleDbCommand com = new OleDbCommand("SELECT * FROM book order by [Accession Number] Asc", con);
            com.ExecuteNonQuery();

            OleDbDataAdapter adap = new OleDbDataAdapter(com);
            DataTable tab = new DataTable();

            adap.Fill(tab);
            grid1.DataSource = tab;
            con.Close();

        }
    }
}
