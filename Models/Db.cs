using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using Microsoft.Data.SqlClient;

namespace AppConvert.Models
{
    public class Db
    {
        SqlConnection con = new SqlConnection("Data Source=LAPTOP-V25BTTUD\\SQLEXPRESS; Initial Catalog=contoh; Integrated Security=True");

        public DataTable Getrecord()
        {
            SqlCommand com = new SqlCommand("select * from [dbo].[customers]", con);
            SqlDataAdapter da = new SqlDataAdapter(com);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }

    }
}