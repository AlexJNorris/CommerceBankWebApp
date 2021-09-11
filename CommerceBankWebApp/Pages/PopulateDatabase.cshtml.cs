using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using CommerceBankWebApp.Data;
using CommerceBankWebApp.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

using Microsoft.Extensions.Logging;
using SpreadsheetLight;

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
            if (!context.Transactions.Any())
            {
                // add all transactions in the excel file to the database
                foreach (Transaction transaction in ReadExcelData("transaction_data.xlsx", "Cust A").Union(ReadExcelData("transaction_data.xlsx", "Cust B")))
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
        public List<Transaction> ReadExcelData(string fileName, string sheetName) {
            List<Transaction> transactionList = new List<Transaction>();

            SLDocument document = new SLDocument(fileName, sheetName);

            SLWorksheetStatistics stats = document.GetWorksheetStatistics();

            for (short i = 2;i <= stats.EndRowIndex;i++)
            {
                string accountType = "Checking";
                long? accountNumber = null;
                DateTime? processingDate = null;
                double? balance = null;
                bool? isCredit = null;
                double? amount = null;
                string description = null;

                for (short j = 1;j <= 7;j++)
                {
                    if (!document.HasCellValue(i, j)) continue;

                    switch (document.GetCellValueAsString(1, j))
                    {
                        case "Account Type":
                            accountType = document.GetCellValueAsString(i, j);
                            break;
                        case "Acct #":
                            accountNumber = document.GetCellValueAsInt64(i, j);
                            break;
                        case "Processing Date":
                            processingDate = document.GetCellValueAsDateTime(i, j);
                            break;
                        case "Balance":
                            balance = document.GetCellValueAsDouble(i, j);
                            break;
                        case "CR (Deposit) or DR (Withdrawal":
                        case "CR (Deposit) or DR (Withdrawal)":
                            string cellText = document.GetCellValueAsString(i, j);

                            if (cellText == "CR") isCredit = true;
                            else isCredit = false;
                            break;
                        case "Amount":
                            amount = document.GetCellValueAsDouble(i, j);
                            break;
                        case "Description 1":
                            description = document.GetCellValueAsString(i, j);
                            break;
                    }

                    if (accountNumber.HasValue && processingDate.HasValue && balance.HasValue && isCredit.HasValue && amount.HasValue && description != null)
                    {
                        Transaction transaction = new Transaction(accountType, accountNumber.Value, processingDate.Value,
                            balance.Value, isCredit.Value, amount.Value, description);

                        transactionList.Add(transaction);
                    }
                }
            }

            return transactionList;
        }

        public void OnGet()
        {
        }
    }
}
