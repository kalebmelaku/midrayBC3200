namespace Mindray_BC3200
{
    using MySql.Data.MySqlClient;
    using System;

    internal class DBConnection
    {
        public static MySqlConnection GetConnection()
        {
            MySqlConnection connection = new MySqlConnection("server=127.0.0.1; database=drhibistpedriati_drhibist;  uid=root;pwd=");
            connection.Open();
            return connection;
        }
    }
}

