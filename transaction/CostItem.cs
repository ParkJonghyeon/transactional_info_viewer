using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Transaction
{
    public class CostItem
    {
        int costItemIndex;
        int transactionIndex;
        string supplier;
        decimal sum;
        string note;
        string credit;
        string cash;
        string card;

        public CostItem(int cIndex)
        {
            this.costItemIndex = cIndex;
        }

        public CostItem(int tIndex, string supplier, decimal sum, string note)
        {
            this.transactionIndex = tIndex;
            this.supplier = supplier;
            this.sum = sum;
            this.note = note;
        }

        public CostItem(int cIndex, int tIndex, string supplier, decimal sum, string note)
        {
            this.costItemIndex = cIndex;
            this.transactionIndex = tIndex;
            this.supplier = supplier;
            this.sum = sum;
            this.note = note;
        }
        
        public int CostItemIndex { get { return costItemIndex; } set { costItemIndex = value; } }
        public int TransactionIndex { get { return transactionIndex; } set { transactionIndex = value; } }
        public string Supplier { get { return supplier; } set { supplier = value; } }
        public decimal Sum { get { return sum; } set { sum = value; } }
        // public string Credit { get { return credit; } set { credit = value; } }
        // public string Cash { get { return cash; } set { cash = value; } }
        // public string Card { get { return card; } set { card = value; } }
        public string Note { get { return note; } set { note = value;; } }
    }
}
