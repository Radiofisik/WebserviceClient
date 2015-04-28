using System;
using System.Data.SQLite;
using System.Data.Common;
using System.Windows.Forms;
using System.Collections.Generic;

namespace WebserviceClient
{
    public  class msql
    {
        static String sqConnectionString = "Data Source=filename.db; Version=3;";
        static bool err = false;
        public  enum Ftypes { Postavka=1, Ustanovka=2, Zakaz=3 };

        public static void setFileName(int orderid, String name, Ftypes type)
        {
            msql.query("INSERT INTO orderfiles(orderid, name, type) Values("+orderid+", \""+name+"\", "+(int)type+")");
        }

        public static List<String> getFileName(int orderid, Ftypes type)
        {
            String sql = "SELECT name from orderfiles where type =" + (int)type + " AND orderid=" + orderid;
           
            List<String> lst = new List<String>();
            SQLiteConnection myConn = new SQLiteConnection(sqConnectionString);
            SQLiteCommand sqCommand = new SQLiteCommand(sql);
            sqCommand.Connection = myConn;
            myConn.Open();
            try
            {
                SQLiteDataReader r = sqCommand.ExecuteReader();
                while (r.Read())
                {
                    lst.Add(r["name"].ToString());
                }
                r.Close();
                err = false;
                return lst;
                
            }
            catch
            {
                err = true;
            }
            finally
            {
                myConn.Close();
            }
            return lst;
        }
     




        public static void query(String sql)
        {

            MessageBox.Show(sql);
                SQLiteConnection myConn = new SQLiteConnection(sqConnectionString);
                SQLiteCommand sqCommand = new SQLiteCommand(sql);
                sqCommand.Connection = myConn;
                myConn.Open();
                try
                {
                    sqCommand.ExecuteNonQuery();
                    err = false;
                }
                catch
                {
                    err = true;
                    MessageBox.Show("sql error");
                }
                finally
                {
                    myConn.Close();
                } 
        }
    }
}