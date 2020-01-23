using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Data;
using System.Security;
using Microsoft.SharePoint.Client;
using System.Data.SqlClient;

namespace ConnectSharePoint_V1
{
    class Program 
    {
            static void Main(string[] args)
            {
                            
                string webUrl = @"https://miremingtonedu.sharepoint.com/sites/Pruebas";
               
                string userName = @"luis.agudelo.3228@miremington.edu.co";
               
                string password = "Hpru1113";
                string NameTable = "Tabla_Prueba";
            string cadena = "Data Source=xx;Initial Catalog=Orchestrator; Integrated Security=True";
                var resp = InsertListInBD(webUrl, userName , password, NameTable,cadena);
                
            }
            private static string InsertListInBD(string Uri, string User, string Password,string NameTable,string StringConnectionBD) {
                try {
                    //Instancia de la Uri (ejm: https://miremingtonedu.sharepoint.com/sites/Pruebas) en este link se encuentra la lista
                    var context = new ClientContext(Uri);
                    //se obtiene el contexto de la paginaa
                    Web web = context.Web;
                    //se convierte la contraseña a un objeto SecureString
                    SecureString password = GetPassword(Password);
                    //se ingresan las credenciales del usuario
                    context.Credentials = new SharePointOnlineCredentials(User, password);
                    //se obtiene la tabla indicada
                    List tabla = web.Lists.GetByTitle(NameTable);
                    //-------se realiza un query sobre todos los elementos de la lista-------
                    CamlQuery query = CamlQuery.CreateAllItemsQuery();//query para consultar
                    ListItemCollection columnas = tabla.GetItems(query);
                    //-----------------------------------------------------------------------
                    //se carga la consulta y se ejecuta
                    context.Load(columnas);
                    context.ExecuteQuery();
                    //se recorren todos los items pendientes
                    var list = columnas.Where(x => x["Estado"].ToString().Equals("Pendiente"));
                    int Countslopes = list.Count();
                    if ( Countslopes != 0 )
                    {                    
                        foreach (ListItem item in list)
                        {
                            Console.WriteLine("------------");
                            Console.WriteLine(item["Title"].ToString() + " | " + item["Fecha_Creacion"].ToString()
                                + " | " + item["Estado"].ToString());
                            Console.WriteLine("------------");
                            InsertTicket(StringConnectionBD,item["Title"].ToString(), item["Fecha_Creacion"].ToString(), item["Estado"].ToString());
                            item["Estado"] = "En Proceso";
                            item.Update();
                            context.ExecuteQuery();
                           
                        }
                        return "1|Succesfull";
                    }
                     else
                    {
                        return "1|No hay elementos en la lista";
                    }
                }
                catch (Exception e) {
                    return "0|Error no controlado: " + e.ToString();
                }
                
            }

            private static SecureString GetPassword(string contraseña)
            {               
                SecureString securePassword = new SecureString();
                foreach(Char x in contraseña)
            { 
                    if ( !x.ToString().Equals("\n") )
                    {
                      securePassword.AppendChar(x);
                    }
                }
            return securePassword;
            
        }

            private static void InsertTicket(string StringConnection,string id, string date,string estado) {
                SqlConnection conexion = new SqlConnection(StringConnection);
                conexion.Open();
                string Query = "insert into Prueba(id,date,estado) values ('" + id + "','" + date + "' , '" + estado + "'"+")";            
                SqlCommand comand = new SqlCommand(Query, conexion);
                comand.ExecuteNonQuery();                
                conexion.Close();
            }
    }
    
}
