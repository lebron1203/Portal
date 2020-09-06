using ExcelDataReader;
using Microsoft.Ajax.Utilities;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Portal.Web.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult FlotCharts()
        {
            return View("FlotCharts");
        }

        public ActionResult MorrisCharts()
        {
            return View("MorrisCharts");
        }

        public ActionResult Tables()
        {
            return View("Tables");
        }

        public ActionResult Forms()
        {
            return View("Forms");
        }

        public ActionResult Panels()
        {
            return View("Panels");
        }

        public ActionResult Buttons()
        {
            return View("Buttons");
        }

        public ActionResult Notifications()
        {
            return View("Notifications");
        }

        public ActionResult Typography()
        {
            return View("Typography");
        }

        public ActionResult Icons()
        {
            return View("Icons");
        }

        public ActionResult Grid()
        {
            return View("Grid");
        }

        public ActionResult Blank()
        {
            return View("Blank");
        }

        public ActionResult Login()
        {
            return View("Login");
        }

        public ActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Upload(HttpPostedFileBase upload)
        {
            if (ModelState.IsValid)
            {

                if (upload != null && upload.ContentLength > 0)
                {
                    // ExcelDataReader works with the binary Excel file, so it needs a FileStream
                    // to get started. This is how we avoid dependencies on ACE or Interop:
                    Stream stream = upload.InputStream;

                    // We return the interface, so that
                    IExcelDataReader reader = null;


                    if (upload.FileName.EndsWith(".xls"))
                    {
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    else if (upload.FileName.EndsWith(".xlsx"))
                    {
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    }
                    else
                    {
                        ModelState.AddModelError("File", "Este formato de archivo no es soportado");
                        return View();
                    }

                    //reader.IsFirstRowAsColumnNames = true;

                    DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                  
                   DataTable dtable = result.Tables[0];

                    //linea de insercion 
                    using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["conexion"].ToString()))
                    {
                        conn.Open();

                        string query = "INSERT INTO tblExcelData (Itemkey_sku,ItemKey,OrigenExcel,Mes_Inicial,Mes_1,Mes_2,Mes_3,Mes_4,Mes_5,Mes_6,Mes_7,Mes_8,Mes_9,Mes_10,Mes_11,Mes_12,EmpCode,location_itemkey,formulaid_excel) VALUES (@Itemkey_sku,@ItemKey,@OrigenExcel,@Mes_Inicial,@Mes_1,@Mes_2,@Mes_3,@Mes_4,@Mes_5,@Mes_6,@Mes_7,@Mes_8,@Mes_9,@Mes_10,@Mes_11,@Mes_12,@EmpCode,@location_itemkey,@formulaid_excel)";


                        SqlCommand cmd = new SqlCommand(query, conn);

                        var empcode = 25421;
                        var location_itemkey = "PT-01";
                        var formulaid = "FP-2013";

                        foreach (DataRow row in dtable.Rows)
                        {

                            cmd.Parameters.Clear();
                            cmd.Parameters.AddWithValue("@Itemkey_sku", row.ItemArray[0]);
                            cmd.Parameters.AddWithValue("@ItemKey", row.ItemArray[0]);
                            cmd.Parameters.AddWithValue("@OrigenExcel", row.ItemArray[0]);
                            cmd.Parameters.AddWithValue("@Mes_Inicial", row.ItemArray[1]);
                            cmd.Parameters.AddWithValue("@Mes_1", row.ItemArray[2]);
                            cmd.Parameters.AddWithValue("@Mes_2", row.ItemArray[3]);
                            cmd.Parameters.AddWithValue("@Mes_3", row.ItemArray[4]);
                            cmd.Parameters.AddWithValue("@Mes_4", row.ItemArray[5]);
                            cmd.Parameters.AddWithValue("@Mes_5", row.ItemArray[6]);
                            cmd.Parameters.AddWithValue("@Mes_6", row.ItemArray[7]);
                            cmd.Parameters.AddWithValue("@Mes_7", row.ItemArray[8]);
                            cmd.Parameters.AddWithValue("@Mes_8", row.ItemArray[9]);
                            cmd.Parameters.AddWithValue("@Mes_9", row.ItemArray[10]);
                            cmd.Parameters.AddWithValue("@Mes_10", row.ItemArray[11]);
                            cmd.Parameters.AddWithValue("@Mes_11", row.ItemArray[12]);
                            cmd.Parameters.AddWithValue("@Mes_12", row.ItemArray[13]);                      
                            cmd.Parameters.AddWithValue("@EmpCode", empcode);
                            cmd.Parameters.AddWithValue("@location_itemkey", location_itemkey);
                            cmd.Parameters.AddWithValue("@formulaid_excel", formulaid);

                            cmd.ExecuteNonQuery();


                        }


                    }


                    return View(result.Tables[0]);



                    //continuacion de datatable//


                    //fin de insercion de data 

                }


                else
                {
                    ModelState.AddModelError("File", "Favor seleccione su archivo");
                }


            }
            return View();
        }



     

        [HttpPost]
        public ActionResult ExecuteProcedure()
        {


            string cnnString = System.Configuration.ConfigurationManager.ConnectionStrings["conexion"].ConnectionString;

            SqlConnection cnn = new SqlConnection(cnnString);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cnn;
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            cmd.CommandText = "borrar";
            //add any parameters the stored procedure might require
            cnn.Open();
            object o = cmd.ExecuteScalar();
            cnn.Close();


            return View("Upload");

        }



    }
}