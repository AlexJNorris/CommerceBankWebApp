using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using CommerceBankWebApp.Data;
using CommerceBankWebApp.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

using System.Data.OleDb;
using System.Data;
using Microsoft.Extensions.Logging;

namespace CommerceBankWebApp.Pages
{
    public class PopulateDatabaseModel : PageModel
    {
        public List<Transaction> Transactions { get; set; }

        public PopulateDatabaseModel(ILogger<PopulateDatabaseModel> logger, ApplicationDbContext context)
        {


            if (!context.Transactions.Any())
            {

                String sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + "transaction_data.xlsx" + ";" + "Extended Properties=Excel 8.0";

                OleDbConnection objConn = new OleDbConnection(sConnectionString);

                objConn.Open();

                OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM ['Cust A$']", objConn);

                OleDbDataReader reader = objCmdSelect.ExecuteReader();

                while (reader.Read())
                {
                    string accountType;

                    try
                    {
                        accountType = reader["Account Type"].ToString();
                    }
                    catch (System.IndexOutOfRangeException e)
                    {
                        accountType = "Checking";
                    }

                    try
                    {

                        int accountNum = Int32.Parse(reader["Acct #"].ToString());
                        DateTime processingDate = DateTime.Parse(reader["Processing Date"].ToString());

                        double balance = Double.Parse(reader["Balance"].ToString());

                        string creditFlag = reader["CR (Deposit) or DR (Withdrawal)"].ToString();

                        bool isCredit = false;
                        if (creditFlag == "CR") isCredit = true;
                        else if (creditFlag == "DR") isCredit = false;
                        else if (creditFlag == "") continue;

                        double amount = Double.Parse(reader["Amount"].ToString());

                        string description = reader["Description 1"].ToString();

                        Transaction transaction = new Transaction(accountType, accountNum, processingDate, balance, isCredit, amount, description);

                        context.Transactions.Add(transaction);

                    }
                    catch (Exception e)
                    {

                    }

                }

                context.SaveChanges();

            }

            Transactions = context.Transactions.ToList();

        }

        public void OnGet()
        {
        }
    }
}
