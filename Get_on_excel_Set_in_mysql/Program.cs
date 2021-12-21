using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetLight;
using MySql.Data.MySqlClient;

namespace Get_on_excel_Set_in_mysql
{
    class Program 
    {      
        static void Main(string[] args)
        {
            try
            {
                string path = @"D:\Descargas\Inevetario laboratorio mayo 2021 .xlsx"; //set path from Excel
                string path2 = @"D:\Descargas\deposito_comp_muebledeherramientas.xlsx";

                int irow = 1;

                SLDocument sl = new SLDocument(path2);

                ConexionBD ob = new ConexionBD();

                MySqlConnection newconnection = ob.conectar();

                while (!string.IsNullOrEmpty(sl.GetCellValueAsString(irow, 1)))
                {
                    string tipo = sl.GetCellValueAsString(irow, 1);
                    string area = sl.GetCellValueAsString(irow, 2);
                    string nombre = sl.GetCellValueAsString(irow, 3);
                    string marca = sl.GetCellValueAsString(irow, 4);
                    string modelo = sl.GetCellValueAsString(irow, 5);
                    int cantidad = sl.GetCellValueAsInt32(irow, 6);
                    string estado = sl.GetCellValueAsString(irow, 7);

                    
                    string consulta = "INSERT INTO deposito_comp (nombre, area, marca, tipo, modelo, cantidad, estado) VALUES ('" + nombre + "','" + area + "','" + marca + "','" + tipo + "','" + modelo + "','" + cantidad + "','" + estado + "')";

                    MySqlCommand comando = new MySqlCommand(consulta, newconnection);

                    comando.ExecuteNonQuery();

                    irow++;
                }                              
            }
            catch(MySqlException ex)
            {
                Console.WriteLine("error " + ex.Message);
            }

            Console.ReadKey();
        }
    }
}
