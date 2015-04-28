using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SQLite;

namespace WebserviceClient
{
    
   
    public partial class Form1 : Form
    {
        public srv.order[] ord;
        public Form1()
        {
            InitializeComponent();
            this.Text = "мегаобработчик заказов КГ НИЦ v1 beta";
           
            //dataGridView1.DataSource.
        }

        

        private void Form1_Load(object sender, EventArgs e)
        {
          srv.mynamePortTypeClient cln = new srv.mynamePortTypeClient();
            
           ord = cln.get_message("3");
       //     MessageBox.Show(ord[0].items[1].Product);
            //dataGridView1.DataSource = ord;
            dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.AutoSizeColumnsMode =DataGridViewAutoSizeColumnsMode.AllCells;

            dataGridView1.ColumnCount=4;
            dataGridView1.Columns[0].Name = "Номер заказа";
            dataGridView1.Columns[1].Name = "Название организации";
            dataGridView1.Columns[2].Name = "Контактное лицо";
            dataGridView1.Columns[3].Name = "Телефон";

         

            DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
            dataGridView1.Columns.Add(btn);
            btn.HeaderText = "сформировать заказ";
            btn.Text = "выгрузить docx";
            btn.Name = "btn";
            btn.UseColumnTextForButtonValue = true;
       

             DataGridViewButtonColumn btn2 = new DataGridViewButtonColumn();
            dataGridView1.Columns.Add(btn2);
            btn2.HeaderText = "сформировать договор поставки";
            btn2.Text = "выгрузить docx";
            btn2.Name = "btn2";
            btn2.UseColumnTextForButtonValue = true;

            DataGridViewButtonColumn btn3 = new DataGridViewButtonColumn();
            dataGridView1.Columns.Add(btn3);
            btn3.HeaderText = "сформировать договор установки";
            btn3.Text = "выгрузить docx";
            btn3.Name = "btn3";
            btn3.UseColumnTextForButtonValue = true;

             



           dataGridView1.CellClick += new DataGridViewCellEventHandler(dataGridView1_CellClick);
           dataGridView1.CellDoubleClick += new DataGridViewCellEventHandler(dataGridView1_CellDoubleClick);




           // btn.Click += new RoutedEventHandler(Onb2Click);

            foreach (srv.order orderi in ord)
            {
                string[] row = new string[] { orderi.id.ToString(), orderi.Organizationshort, orderi.FIOcont, orderi.Phonecont };
                dataGridView1.Rows.Add(row);
            }
           // dataGridView1.Columns.Add(ord[0].Organizationshort);

            //погасить кнопки
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {

                DataGridViewTextBoxCell btnn = new DataGridViewTextBoxCell();
               // MessageBox.Show(row.Cells[0].Value.ToString());
             
                if (msql.getFileName(Int32.Parse((String)row.Cells[0].Value), msql.Ftypes.Zakaz).Count>0) row.Cells[4]=btnn;
                if (msql.getFileName(Int32.Parse((String)row.Cells[0].Value), msql.Ftypes.Postavka).Count>0) row.Cells[5] = btnn;
                if (msql.getFileName(Int32.Parse((String)row.Cells[0].Value), msql.Ftypes.Ustanovka).Count>0) row.Cells[6] = btnn;
            }

            dataGridView1.Refresh();


        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 4)
            {
                //MessageBox.Show((e.RowIndex + 1) + "  Row  " + (e.ColumnIndex + 1) + "  Column button clicked ");
                GeneratedCode.GeneratedClass a = new GeneratedCode.GeneratedClass();
                a.orderi = ord[e.RowIndex];
                String namestr = a.orderi.id.ToString() + "_" + DateTime.Today.ToShortDateString() + a.orderi.Organizationshort + ".docx";
                a.CreatePackage("C:/" + namestr);
                msql.setFileName(a.orderi.id, namestr, msql.Ftypes.Zakaz);
                MessageBox.Show(namestr + "сфромирован");
            }
            else
            {
               
                //msql.query("INSERT INTO ordertable(id, orderid) Values(50, 20)");
            /*    msql.setFileName(1, "file.txt", msql.Ftypes.Ustanovka);
                String a=msql.getFileName(1, msql.Ftypes.Ustanovka).First<String>();
               
                MessageBox.Show(a);*/



                 
                /////////////////////////////////////////////////////////////////


            }
        }
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show("hi");
            itemform itf = new itemform();
            itf.orderi = ord[e.RowIndex];
            itf.initdatagrid();
            itf.Show();
        }

    }
}
