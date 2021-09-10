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
        // will store all Transactions in the database for testing purposes (regardless of accountnumber)
        public List<Transaction> Transactions { get; set; }

        private readonly ILogger<PopulateDatabaseModel> _logger;
        private readonly ApplicationDbContext _context;

        // The constructor enables logging and access to the database
        public PopulateDatabaseModel(ILogger<PopulateDatabaseModel> logger, ApplicationDbContext context)
        {
            _logger = logger;
            _context = context;

            // if there are no transactions in the database yet we will read the excel file and add the data
            if (!_context.Transactions.Any())
            {
                // add all transactions in the excel file to the database
                foreach (Transaction transaction in ReadExcelData())
                {
                    _context.Add(transaction);
                }

                // save the changes
                _context.SaveChanges();
            }

            // read all transactions in the database into the Transactions property so we can read the data in the razor page
            Transactions = context.Transactions.ToList();

        }

        // returns a list of all transactions in the excel file
        public List<Transaction> ReadExcelData() {
            List<Transaction> transactionList = new List<Transaction>();

            // connection settings to load transaction_data.xlsx in the root directory of the project
            String sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + "transaction_data.xlsx" + ";" + "Extended Properties=Excel 8.0";

            OleDbConnection objConn = new OleDbConnection(sConnectionString);
            objConn.Open();

            /* command selects the first sheet in the excel file. Currently we are only reading the first sheet
            this can be changed */

            OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM ['Cust A$']", objConn);

            // run the command
            OleDbDataReader reader = objCmdSelect.ExecuteReader();

            // read each row
            while (reader.Read())
            {
                string accountType;

                // try to read the account type entry in the sheet. The sheet for Cust A contains this entry, Cust B does not
                try
                {
                    accountType = reader["Account Type"].ToString();
                }
                catch (System.IndexOutOfRangeException e)
                {
                    // if we werent able to read the account type assume checking
                    accountType = "Checking";
                }

                // try to read the data for each field
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

                    transactionList.Add(transaction);

                }
                catch (Exception e)
                {
                    //TODO: Something if we couldnt read the current row as a transaction
                }

            }

            return transactionList;
        }

        public void OnGet()
        {
        }
    }
}
