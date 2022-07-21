using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StokListesi
{
    class Connections
    {
        public static MySqlConnection MySqlBaglanti = new MySqlConnection("Server=192.167.1.107;Port=3306;Database=xrm;Uid=root;Pwd=emp1881;");
        public static SqlConnection SqlBaglanti = new SqlConnection("Server=192.167.1.11;Database=STORE2022;User Id=sait;Password=Sait.1453;");
    }
}
