using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Transaction
{
    public class Transaction
    {
        int index;
        string customerName;
        string transactionName;
        string transactionDate;
        decimal supplyPrice;
        int department;
        string transactionCode;

        public Transaction(int index)
        {
            this.index = index;
        }
        
        public Transaction(int index, string customerName, string transactionName, string transactionDate, decimal supplyPrice, int department, string transactionCode)
        {
            this.index = index;
            this.customerName = customerName;
            this.transactionName = transactionName;
            this.transactionDate = transactionDate;
            this.supplyPrice = supplyPrice;
            this.department = department;
            this.transactionCode = transactionCode;
        }

        public int Index { get { return index; } set { index = value; } }
        public string CustomerName { get { return customerName; } set { customerName = value; } }
        public string TransactionName { get { return transactionName; } set { transactionName = value; } }
        public string TransactionDate { get { return transactionDate; } set { transactionDate = value; } }
        public decimal SupplyPrice { get { return supplyPrice; } set { supplyPrice = value; } }  
        public int Department { get { return department; } set { department = value; } }
        public string TransactionCode { get { return transactionCode; } set { transactionCode = value; } }
    }
}
