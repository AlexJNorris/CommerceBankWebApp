using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CommerceBankWebApp.Models
{
    [Table("Transactions")]
    public class Transaction
    {
        [Column("ID")]
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Required]
        public int ID { get; set; }

        [Column("AccountType")]
        [Required]
        public string AccountType { get; set; }

        [Column("AccountNumber")]
        [Required]
        public int AccountNumber { get; set; }
        [Column("ProcesingDate")]
        [Required]
        public DateTime ProcessingDate { get; set; }
        [Column("Balance")]
        [Required]
        public double Balance { get; set; }
        [Column("IsCredit")]
        [Required]
        public bool IsCredit { get; set; }
        [Column("Amount")]
        [Required]
        public double Amount { get; set; }
        [Column("Description")]
        [Required]
        public string Description { get; set; }

        public Transaction()
        {
        }

        public Transaction(string accountType, int accountNumber, DateTime processingDate, double balance, bool isCredit, double amount, string description)
        {
            AccountType = accountType;
            AccountNumber = accountNumber;
            ProcessingDate = processingDate;
            Balance = balance;
            IsCredit = isCredit;
            Amount = amount;
            Description = description;
        }
    }
}
