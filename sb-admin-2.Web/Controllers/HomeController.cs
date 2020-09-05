using ExcelDataReader;
using Microsoft.Ajax.Utilities;
using System;
using System.Collections.Generic;
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
                    using (SqlConnection conn = new SqlConnection("Data Source=MYPC;Initial Catalog=Productos;Integrated Security=True"))
                    {
                        conn.Open();

                        string query = "INSERT INTO tblExcelData (Itemkey_sku) VALUES (@param1)";
                        SqlCommand cmd = new SqlCommand(query, conn);


                        foreach (DataRow row in dtable.Rows)
                        {

                            cmd.Parameters.Clear();
                            cmd.Parameters.AddWithValue("@param1", row.ItemArray[0]);
                            //cmd.Parameters.AddWithValue("@param2", row[1]);

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


       

    }
}