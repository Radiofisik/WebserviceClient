using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WebserviceClient
{
    public partial class itemform : Form
    {
        public srv.order orderi;
        public itemform()
        {
            InitializeComponent();
           
        }
        public void initdatagrid()
        {
            //////////////////////////////////////////////////////////////
            dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.AutoSizeColumnsMode =
            DataGridViewAutoSizeColumnsMode.AllCells;

            dataGridView1.ColumnCount = 2;
            dataGridView1.Columns[0].Name = "Наименование продукта";
            dataGridView1.Columns[1].Name = "Количество";


            foreach (srv.orderItem orderitemi in orderi.items)
            {
                string[] row = new string[] { orderitemi.Product, orderitemi.Quantity.ToString() };
                dataGridView1.Rows.Add(row);
            }
            ////////////////////////////////////////////////////////////
        }
    }
}
