using AppBlock;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebApplication1
{
    public partial class file_upload : System.Web.UI.Page
    {
        string sqlcon = ConfigurationManager.ConnectionStrings["formcreation"].ToString();
        //connection 
        OleDbConnection con;
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void ExcelConnection(String path)
        {
            //ms excel connection string
            string connection = $" Microsoft.ACE.OLEDB.12.0; Data Source = {path}; Extended Properties = 'Excel 12.0 Xml;HDR=YES'";
            // creating an new oledb connection
            con = new OleDbConnection();
        }

        protected void UploadButton_Click(object sender, EventArgs e)
        {
            // checks whether the file is uploaded
            if (UploadFileAsExcel.HasFile)
            {
                string path;
                try
                {
                    // string s = Server.MapPath("../UploadedEmandateFiles/") + newfile;

                    // creating an path
                    path = Server.MapPath("~/Uploads/");
                    // stores values into folder
                    if (!Directory.Exists(path))
                    {
                        //if folder not exists it will create an directory
                        Directory.CreateDirectory(path);
                    }


                    string filename = Path.GetFileName(UploadFileAsExcel.FileName);
                    //string newfile = filename.Replace(filename, "UpdatedEmandatefile.xls");
                    UploadFileAsExcel.SaveAs(Server.MapPath("~/Uploads/") + filename);
                    //alert for upload sucess
                    label.Text = "File uploaded successfully";
                    label.ForeColor = System.Drawing.Color.Green;
                    label.Visible = true;
                    //Excel Connection
                    //  ExcelConnection(path);
                    // con.Open();
                    // reads excel data
                     path = Server.MapPath("~/Uploads/" + filename );
                    DataTable dt = ReadExcelRecords(path);
                    insertDataIntoDatabase(dt);
                    ClientScript.RegisterStartupScript(GetType(), "alert", "alert('Excel Record entered into database Successfully!!!');", true);
                    
                }
                catch (Exception ex)
                {
                    ClientScript.RegisterStartupScript(GetType(), "alert", "alert('Error uploading file: " + ex.Message.Replace("'", "\\'") + "');", true);
                }
            }
        }

        protected DataTable ReadExcelRecords(string path)
        {
           
         try
            {
                //excel connection
                
                string filename = Path.GetFileName(UploadFileAsExcel.FileName);
                string fileExtension = Path.GetExtension(filename.ToLower());
                string connection = string.Empty;

                if (fileExtension == ".xls")
                {
                     connection = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={path};Extended Properties='Excel 8.0;HDR=YES'";
                }else if (fileExtension == ".xlsx")
                {
                     connection = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={path};Extended Properties='Excel 12.0 Xml;HDR=YES '";
                }else
                {
                    ClientScript.RegisterStartupScript(GetType(), "alert", "alert('Invalid file format. Please upload an Excel file');", true);
                }

                //oledb connection
                con = new OleDbConnection(connection);
                con.Open();
                //creating an new data table
                DataTable dt = new DataTable();
                //creating new sheet and storing table
                DataTable sheets = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = sheets.Rows[0]["TABLE_NAME"].ToString();
                if (sheets == null || sheets.Rows.Count == 0)
                {
                    ClientScript.RegisterStartupScript(GetType(), "alert", "alert('No sheets found in the Excel file.');", true);
                    return null; // Return null if no sheets are found
                }
                
                OleDbCommand cmd = new OleDbCommand($"select * from [{sheetName}]", con);
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                con.Close();
                return dt;
            }
            catch (Exception ex)
            {
                ClientScript.RegisterStartupScript(GetType(), "alert", "alert('Error uploading file: " + ex.Message.Replace("'", "\\'") + "');", true);
                return null;
            }
        }
     

     protected void insertDataIntoDatabase(DataTable dt)
        {
            SqlConnection sc = new SqlConnection(sqlcon);
            sc.Open();
            foreach(DataRow row in dt.Rows)
            {
                SqlCommand sqlCMD = new SqlCommand("INSERT INTO form (firstname,lastname,date,address) VALUES(@firstname,@lastname,@date,@address)", sc);
            
                    //add parameters to the sql command
                    sqlCMD.Parameters.AddWithValue("@firstname", row["firstname"]);
                    sqlCMD.Parameters.AddWithValue("@lastname", row["lastname"]);
                    sqlCMD.Parameters.AddWithValue("@date", row["date"]);
                    sqlCMD.Parameters.AddWithValue("@address", row["address"]);

                    sqlCMD.ExecuteNonQuery();
             


            }

        }




    }
}