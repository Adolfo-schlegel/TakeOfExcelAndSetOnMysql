using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;

namespace Get_on_excel_Set_in_mysql
{
    class ConexionBD
    {
        public MySqlConnection conectar()
        {
            try
            {
                string query = "Database = stock_eet2; Data Source = localhost; Port = 3306; password = 10deagosto; user id = root;";

                MySqlConnection newconnection = new MySqlConnection(query);

                newconnection.Open();

                return newconnection;
            }
            catch(MySqlException ex)
            {
                Console.WriteLine("error " + ex.Message);
                return null;
            }
           
        }
    }
}
