using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient; 
using System.Windows.Forms;

namespace WindowsFormsApplication1
{

    class Class1
    {
        private SqlConnection sqlconexion; //para sql server
        private SqlCommand sqlcomando; //para sqlserver
        private SqlDataAdapter sqladapter;
        private OleDbConnection oledbconexion;  //para acces
        private OleDbCommand oledbcomando; //para acces
        private OleDbDataAdapter oledbadapter; //para acces
        private DataTable objtabla;

        private void iniciaSQL()
        {
            try
            {
                sqlconexion = new SqlConnection();
                sqlconexion.ConnectionString = "aqui va la cadena de conexion de sql Server";
                sqlconexion.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show(Convert.ToString(ex));
            }
        }


        private void finalSQL()
        {
            try
            {
                if (sqlconexion.State == ConnectionState.Open)
                {
                    sqlconexion.Close(); //cierra la base de datos
                }

                sqlconexion.Dispose(); //elimina el objeto
            }
            catch (Exception ex)
            {
                MessageBox.Show(Convert.ToString(ex));
            }
        }


        private void iniciaOledb()
            {
        try
        {
            oledbconexion = new OleDbConnection;
            oledbconexion.ConnectionString = "aqui va la cadena de conexion de acces";
            oledbconexion.Open();
        }
        catch (Exception ex) 
            {
            MessageBox.Show(Convert.ToString(ex));
            }
    
            }



        private void finalOledb()
        {
            try
            {
                if (oledbconexion.State == ConnectionState.Open)
                {
                    oledbconexion.Close();
                }
                oledbconexion.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(Convert.ToString(ex));
            }
         
       
        }

        public Boolean operasql(string str)
        { 
        iniciaSQL(();
           try
           {
               sqlcomando=new SqlCommand;
               sqlcomando.CommandText=str;
                  sqlcomando.Connection=sqlconexion;
               sqlcomando.ExecuteNonQuery()
           }

        }


}



   

}
      
    
        
    

