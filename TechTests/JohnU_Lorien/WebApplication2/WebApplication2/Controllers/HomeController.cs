using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using System.IO;
using System.Xml;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

using WebApplication2.Models;
using WebApplication2.DAL;

namespace WebApplication2.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpGet]
        public ActionResult ProcessTaxReturns()
        {
            return View();
        }

        /// <summary>
        /// Uplaod the excel file and store the valid data into local database file
        /// </summary>
        /// <param name="uFile"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ProcessTaxReturns(HttpPostedFileBase uFile)
        {
            if (ModelState.IsValid)
            {
                if (Request.Files["uFile"] != null)
                {
                    if (Request.Files["uFile"].ContentLength > 0)
                    {
                        var fileExt = Path.GetExtension(Request.Files["uFile"].FileName);
                        if (fileExt == ".csv" || fileExt == ".xlsx")
                        {
                            var fileLocation = Server.MapPath("~/Content/") + Request.Files["uFile"].FileName;
                            if (System.IO.File.Exists(fileLocation))
                            {
                                System.IO.File.Delete(fileLocation);
                            }

                            uFile.SaveAs(fileLocation);


                            var excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                            var excelConn = new OleDbConnection(excelConnectionString);
                            try
                            {
                                excelConn.Open();
                                var dt = new DataTable();
                                var ds = new DataSet();
                                if (excelConn.State == ConnectionState.Open)
                                {
                                    excelConn.Close();
                                }

                                excelConn.Open();
                                dt = excelConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                                for (var i = 0; dt.Rows.Count > i; i++)
                                {
                                    if (Convert.ToString(dt.Rows[i].ItemArray[2]).ToLower() == "sheet1$")
                                    {
                                        var oleDbCmd = new OleDbCommand("SELECT * FROM [Sheet1$]", excelConn);
                                        var oldDbAdapter = new OleDbDataAdapter(oleDbCmd);
                                        oldDbAdapter.Fill(ds);
                                        break;
                                    }
                                }
                                var successCount = 0;
                                var errorCount = 0;
                                for (var i = 0; i < ds.Tables[0].Rows.Count; i++)
                                {
                                    var conn = ConfigurationManager.ConnectionStrings["TransactionsConnStr"].ConnectionString;
                                    var con = new SqlConnection(conn);
                                    var accountNumber = !String.IsNullOrEmpty(ds.Tables[0].Rows[i][0].ToString()) ? ds.Tables[0].Rows[i][0].ToString() : String.Empty;
                                    var description = !String.IsNullOrEmpty(ds.Tables[0].Rows[i][1].ToString()) ? ds.Tables[0].Rows[i][1].ToString() : String.Empty;
                                    var currencyCode = !String.IsNullOrEmpty(ds.Tables[0].Rows[i][2].ToString()) ? IsValidCurrencyCode(ds.Tables[0].Rows[i][2].ToString()) ? ds.Tables[0].Rows[i][2].ToString() : String.Empty : String.Empty;
                                    var amount = !String.IsNullOrEmpty(ds.Tables[0].Rows[i][1].ToString()) ? ds.Tables[0].Rows[i][3].ToString() : String.Empty;

                                    if (!String.IsNullOrEmpty(accountNumber) && !String.IsNullOrEmpty(description) && !String.IsNullOrEmpty(currencyCode) && !String.IsNullOrEmpty(amount))
                                    {
                                        string query = "Insert into testt(AccountNumber,Description,CCode,Amount) Values('" + accountNumber + "','" + description + "','" + currencyCode + "', " + Convert.ToDecimal(amount) + ")";
                                        con.Open();
                                        SqlCommand cmd = new SqlCommand(query, con);
                                        cmd.ExecuteNonQuery();
                                        con.Close();

                                        successCount++;
                                    }
                                    else
                                    { 
                                        errorCount++; 
                                    }
                                }

                                var result = "Data added to the database";
                                if (successCount > 0)
                                    result += "; Total " + successCount + " records added successfully.";
                                if (errorCount > 0)
                                    result += "; Total " + errorCount + " records skipped due to invalid data.";

                                ViewBag.Result =  result;
                                excelConn.Close();
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                            finally
                            {
                                if (excelConn.State == ConnectionState.Open)
                                {
                                    excelConn.Close();
                                }
                            }
                        }
                        else
                        {
                            ViewBag.Result = "Please select a valid .xlsx or .csv file";
                        }
                    }
                    else
                    {
                        ViewBag.Result = "Please select a file with non-empty contents";
                    }
                }
                else
                {
                    ViewBag.Result = "Error in Request";
                }

            }
            else
            {
                ViewBag.Result = "Error while processing";
            }

            return View();
        }
    
        /// <summary>
        /// Method to retrieve existing data from the database
        /// </summary>
        /// <returns></returns>
        public ActionResult ViewDbData()
        {
            var dbContext = new TransactionsContext();
            var transactionList = new List<Transaction>();
            foreach (var item in dbContext.testts)
            {
                transactionList.Add(new Transaction { Account = item.AccountNumber, Description = item.Description, CurrencyCode = item.CCode, Amount = item.Amount, Id = item.Id });
            }
            return View(transactionList);
        }

        /// <summary>
        /// Method to delete a single item from the database
        /// </summary>
        /// <param name="id">id of the item marked for deletion</param>
        /// <returns></returns>
        public ActionResult Delete(int id)
        {
            var dbContext = new TransactionsContext();
            var itemToRemove = dbContext.testts.SingleOrDefault(x => x.Id == id);
            if (itemToRemove != null)
            {
                dbContext.testts.Remove(itemToRemove);
                dbContext.SaveChanges();
            }

            return RedirectToAction("ViewDbData");
        }


        /// <summary>
        /// Checks if the passed string is a valid currency code
        /// </summary>
        /// <param name="str">string to check</param>
        /// <returns>'true' if the string is valid ccode ortherwise 'false'</returns>
        private bool IsValidCurrencyCode(string str)
        {
            var result = false;
            var allCurrencyCodes = new List<string>();
            var xmlDoc = new XmlDocument();
            xmlDoc.Load("http://www.currency-iso.org/dam/downloads/lists/list_one.xml");
            var xmlNodes = xmlDoc.SelectNodes(".//Ccy");
            if (xmlNodes.Count > 0)
            {
                foreach (XmlNode xn in xmlNodes)
                {
                    if (xn != null)
                    {
                        allCurrencyCodes.Add(xn.InnerText);
                    }
                }
            }

            if (allCurrencyCodes.Count > 0)
            {
                var cc = from cc1 in allCurrencyCodes
                         .Where(c => c.ToLower().Trim() == str.ToLower().Trim())
                         select cc1;

                if (cc != null)
                    return true;
                else
                    return false;
            }

            return result;
        }
    }
}