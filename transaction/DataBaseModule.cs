using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.IO;

namespace Transaction
{
    public class DataBaseModule
    {
        private string strCon;
        OracleConnection conn;
        OracleCommand cmd;
        OracleDataAdapter adapter;
        DataSet dataSet;
        DataTable transactionDataTable, costItemDataTable, departmentTransactionDataTable;

        int connectedUserDepartment = -1;

        #region DBTableGetter
        public DataTable TransactionDataTable { get { return transactionDataTable; } }
        public DataTable CostItemDataTable { get { return costItemDataTable; } }
        public DataTable DepartmentTransactionDataTable { get { return departmentTransactionDataTable; } }
        #endregion

        public DataBaseModule(int calledForm, int department)
        {
            FileInfo exefileinfo = new FileInfo(Application.ExecutablePath);

            // string path = exefileinfo.Directory.FullName.ToString();  //프로그램 실행되고 있는 path 가져오기
            // string fileName = @"\setting.ini";  //파일명
            // string filePath = path + fileName;   //ini 파일 경로 
            
            string filePath = Application.StartupPath + @"\\setting.ini";

            iniUtil ini = new iniUtil(filePath);

            /*
            string PROTOCOL = "TCP";
            string HOST = "";
            string PORT = "";
            string SERVICE_NAME = "XE";
            string USER_ID = "system";
            string PASSWORD = "oracle";
            
            ini.SetIniValue("Oracle", "USER_ID", "system");
            ini.SetIniValue("Oracle", "PASSWORD", "oracle");
            ini.SetIniValue("Oracle", "HOST", "");
            ini.SetIniValue("Oracle", "PROTOCOL", "TCP");
            ini.SetIniValue("Oracle", "PORT", "");
            ini.SetIniValue("Oracle", "SERVICE_NAME", "XE");
            */
            
            string USER_ID = ini.GetIniValue("Oracle", "USER_ID");
            string PASSWORD = ini.GetIniValue("Oracle", "PASSWORD");
            string HOST = ini.GetIniValue("Oracle", "HOST");
            string PROTOCOL = ini.GetIniValue("Oracle", "PROTOCOL");
            string PORT = ini.GetIniValue("Oracle", "PORT"); ;
            string SERVICE_NAME = ini.GetIniValue("Oracle", "SERVICE_NAME"); ;

            /// tnsnames.ora
            strCon = @"Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL="+PROTOCOL+")"
                                + "(HOST=" + HOST + ")"
                                + "(PORT=" + PORT + ")))"
                                + "(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=" + SERVICE_NAME + ")));"
                                + "User ID=" + USER_ID + ";"
                                + "Password=" + PASSWORD + ";";

            if (calledForm == -1)
            {
                DBConnect();
            }
            else if (calledForm == 0)
            {
                connectedUserDepartment = department;
                getAllTables();
                getCustomerNameList(department);
            }
            else if (calledForm == 1)
            {
                connectedUserDepartment = department;
                getAllTables();
            }
        }

        public bool DBConnect()
        {
            try
            {
                if (conn == null)
                    conn = new OracleConnection(strCon);
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                    cmd = new OracleCommand();
                    cmd.Connection = conn;

                    adapter = new OracleDataAdapter();
                    dataSet = new DataSet();
                }
            }
            catch (Exception e)
            {
                if(MessageBox.Show("DataBase 접속 실패.\n 프로그램을 종료합니다.", "Error", MessageBoxButtons.OK) == DialogResult.OK)
                {
                    Application.Exit();
                }
            }

            return conn.State == ConnectionState.Open ? true : false;
        }

        public void DBDisconnect()
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
                conn.Dispose();
                conn = null;
            }
        }

        public bool Insert(Object obj)
        {
            bool returnResult = false;

            if (obj is Transaction)
            {
                Transaction castObj = (Transaction)obj;
                string sql = "INSERT INTO TRANSACTION_TABLE VALUES (" + castObj.Index + ", '" + castObj.CustomerName + "', '" + castObj.TransactionName + "', '" + castObj.TransactionDate + "', " + castObj.SupplyPrice + ", " + castObj.Department + ", '" + castObj.TransactionCode + "')";
                returnResult = excuteQuery(sql);
            }
            else if (obj is CostItem)
            {
                CostItem castObj = (CostItem)obj;
                string sql = "INSERT INTO COST_ITEM_TABLE VALUES  (" + castObj.CostItemIndex + ", " + castObj.TransactionIndex + ", '" + castObj.Supplier + "', '" + castObj.Sum + "', NULL, NULL, NULL, '" + castObj.Note + "')";
                returnResult = excuteQuery(sql);
            }
            else if (obj is User)
            {
                User castObj = (User)obj;
                string sql = "INSERT INTO TRANSACTION_USER_TABLE VALUES ('" + castObj.Id + "', '" + castObj.Password + "', " + castObj.Authority + ", " + castObj.Department + ", 'logout')";
                returnResult = excuteQuery(sql);
            }

            getAllTables();

            return returnResult;
        }

        public bool Delete(Object obj)
        {
            bool returnResult = false;

            if (obj is Transaction)
            {
                Transaction castObj = (Transaction)obj;
                string sql = "DELETE FROM TRANSACTION_TABLE WHERE TRANSACTION_INDEX = " + castObj.Index;
                returnResult = excuteQuery(sql);
                // 해당 거래명의 인덱스를 가진 값은 모두 삭제
                sql = "DELETE FROM COST_ITEM_TABLE WHERE TRANSACTION_INDEX = " + castObj.Index;
                returnResult = excuteQuery(sql);
            }
            else if (obj is CostItem)
            {
                CostItem castObj = (CostItem)obj;
                string sql = "DELETE FROM COST_ITEM_TABLE WHERE COST_ITEM_INDEX = " + castObj.CostItemIndex;
                returnResult = excuteQuery(sql);
            }
            else if (obj is User)
            {
                User castObj = (User)obj;
                string sql = "DELETE FROM TRANSACTION_USER_TABLE WHERE USER_ID = '" + castObj.Id + "'";
                returnResult = excuteQuery(sql);
            }

            getAllTables();

            return returnResult;
        }

        public bool Update(Object obj)
        {
            bool returnResult = false;
            
            if (obj is Transaction)
            {
                Transaction castObj = (Transaction)obj;
                if(isRowExist(castObj.Index, 0))
                {
                    string sql = "UPDATE TRANSACTION_TABLE SET CUSTOMER_NAME = '" + castObj.CustomerName + "', TRANSACTION_NAME = '" + castObj.TransactionName + "', TRANSACTION_DATE = '" + castObj.TransactionDate + "', SUPPLY_PRICE = " + castObj.SupplyPrice + ", TRANSACTION_CODE = '" + castObj.TransactionCode + "' WHERE TRANSACTION_INDEX =" + castObj.Index;
                    returnResult = excuteQuery(sql);
                }
            }
            else if (obj is CostItem)
            {
                CostItem castObj = (CostItem)obj;
                if (isRowExist(castObj.CostItemIndex, 1))
                {
                    string sql = "UPDATE COST_ITEM_TABLE SET SUPPLIER = '" + castObj.Supplier + "', SUM = " + castObj.Sum + ", NOTE = '" + castObj.Note + "' WHERE COST_ITEM_INDEX =" + castObj.CostItemIndex;
                    returnResult = excuteQuery(sql);
                }
            }
            else if (obj is User)
            {
                User castObj = (User)obj;
                string sql = "UPDATE TRANSACTION_USER_TABLE SET USER_PASSWORD = '" + castObj.Password + "', USER_AUTHORITY = " + castObj.Authority + ", DEPARTMENT = " + castObj.Department + "WHERE USER_ID = '" + castObj.Id + "'";
                returnResult = excuteQuery(sql);
            }

            getAllTables();

            return returnResult;
        }
        
        public bool excuteQuery(String sql)
        {
            if (!DBConnect())
            {
                return false;
            }
            else
            {
                try
                {
                    cmd.CommandText = sql;
                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();
                    //cmd.Transaction.Commit();
                }
                catch (Exception e)
                {
                    //cmd.Transaction.Rollback();
                    Console.WriteLine("Sqlerror msg : " + e.Message);
                    return false;
                }
                finally
                {
                    cmd.Dispose();
                    DBDisconnect();
                }
                return true;
            }
        }

        public bool isRowExist(int index, int num)
        {
            if (!DBConnect())
            {
                MessageBox.Show("데이터 베이스 접근 실패.", "Error", MessageBoxButtons.OK);
                return false;
            }
            else
            {
                string query;

                try
                {
                    if (num == 0)
                        query = "SELECT TRANSACTION_INDEX FROM TRANSACTION_TABLE WHERE (TRANSACTION_INDEX = " + index + ")";
                    else
                        query = "SELECT COST_ITEM_INDEX FROM COST_ITEM_TABLE WHERE (TRANSACTION_INDEX = " + index + ")";
                    cmd.CommandText = query;
                    adapter.SelectCommand = cmd;

                    DataTable tmpTable = new DataTable();
                    adapter.Fill(tmpTable);

                    DBDisconnect();

                    if (tmpTable.Rows.Count > 0)
                        return true;
                    else
                        return false;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.StackTrace);
                }
            }

            return false;
        }

        #region Using LoginForm
        public string isExistUser(User user)
        {
            if (!DBConnect())
            {
                MessageBox.Show("데이터 베이스 접근 실패.", "Error", MessageBoxButtons.OK);
                return "error";
            }
            else
            {
                string query;

                try
                {
                    query = "SELECT USER_AUTHORITY, DEPARTMENT, CONNECTED FROM TRANSACTION_USER_TABLE WHERE (USER_ID = '" + user.Id + "' AND USER_PASSWORD = '" + user.Password + "')";
                    cmd.CommandText = query;
                    adapter.SelectCommand = cmd;

                    DataTable tmpTable = new DataTable();
                    adapter.Fill(tmpTable);

                    DBDisconnect();

                    string connectedAnotherPC = "not user";

                    if (tmpTable.Rows.Count > 0)
                    {
                        user.Authority = Int32.Parse(tmpTable.Rows[0].ItemArray[0].ToString());
                        user.Department = Int32.Parse(tmpTable.Rows[0].ItemArray[1].ToString());
                        connectedAnotherPC = tmpTable.Rows[0].ItemArray[2].ToString();
                        return connectedAnotherPC;
                    }
                    else
                    {
                        return connectedAnotherPC;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.StackTrace);
                    return "error";
                }
            }
        }

        public bool isRegisteredUser(string userId)
        {
            if (!DBConnect())
            {
                MessageBox.Show("데이터 베이스 접근 실패.", "Error", MessageBoxButtons.OK);
                return false;
            }
            else
            {
                string query;

                try
                {
                    query = "SELECT USER_ID FROM TRANSACTION_USER_TABLE WHERE (USER_ID = '" + userId + "')";
                    cmd.CommandText = query;
                    adapter.SelectCommand = cmd;

                    DataTable tmpTable = new DataTable();
                    adapter.Fill(tmpTable);

                    DBDisconnect();

                    if (tmpTable.Rows.Count > 0)
                    {
                        return !userId.Equals(tmpTable.Rows[0].ItemArray[0].ToString());
                    }
                    else
                    {
                        return true;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.StackTrace);
                    return false;
                }
            }
        }

        public bool Login(Object obj)
        {
            bool returnResult = false;

            if (obj is User)
            {
                User castObj = (User)obj;
                string sql = "UPDATE TRANSACTION_USER_TABLE SET CONNECTED = 'login' WHERE USER_ID = '" + castObj.Id + "'";
                returnResult = excuteQuery(sql);
            }

            getAllTables();

            return returnResult;
        }

        public bool Logout(Object obj)
        {
            bool returnResult = false;

            if (obj is User)
            {
                User castObj = (User)obj;
                string sql = "UPDATE TRANSACTION_USER_TABLE SET CONNECTED = 'logout' WHERE USER_ID = '" + castObj.Id + "'";
                returnResult = excuteQuery(sql);
            }

            getAllTables();

            return returnResult;
        }
        #endregion

        // 시작 이후, DB 갱신 이후에 갱신 된 DB로 메모리 상의 Table들을 업데이트
        public bool getAllTables()
        {
            if (!DBConnect())
            {
                return false;
            }
            else
            {
                string query;

                try
                {
                    // dataTable의 묶음인 dataSet에 테이블 두개를 채워넣음
                    // 존재하지 않는 테이블의 경우 catch에서 테이블 생성 쿼리문 실행
                    query = "SELECT * FROM TRANSACTION_TABLE";
                    cmd.CommandText = query;
                    adapter.SelectCommand = cmd;
                    adapter.Fill(dataSet, "TRANSACTION_TABLE");

                    query = "SELECT * FROM COST_ITEM_TABLE";
                    cmd.CommandText = query;
                    adapter.SelectCommand = cmd;
                    adapter.Fill(dataSet, "COST_ITEM_TABLE");
                    
                    query = "SELECT * FROM TRANSACTION_TABLE WHERE DEPARTMENT = " + connectedUserDepartment;
                    cmd.CommandText = query;
                    adapter.SelectCommand = cmd;
                    adapter.Fill(dataSet, "DEPARTMENT_TRANSACTION_TABLE");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.StackTrace);
                    return false;
                }

                // dataSet에서 각각 해당하는 dataTable을 꺼내어 저장
                transactionDataTable = dataSet.Tables["TRANSACTION_TABLE"];
                costItemDataTable = dataSet.Tables["COST_ITEM_TABLE"];
                departmentTransactionDataTable = dataSet.Tables["DEPARTMENT_TRANSACTION_TABLE"];

                DBDisconnect();
                return true;
            }
        }

        public DataView getOutputTable(string query, string sort)
        {
            DataTable outputDataTable = new DataView(transactionDataTable, query, sort, DataViewRowState.CurrentRows).ToTable();

            outputDataTable.Columns.Add("TRANSACTION_NUM");
            outputDataTable.Columns.Add("TOTAL_COST");
            outputDataTable.Columns.Add("PROFIT");
            outputDataTable.Columns.Add("TAX");
            outputDataTable.Columns.Add("FINAL_PROFIT");
            outputDataTable.Columns.Add("NOTE");

            decimal outputTotalCost, outputProfit, outputTax;
            int rows = outputDataTable.Rows.Count;
            int count = 0;

            while (count < rows)
            {
                outputDataTable.Rows[count].BeginEdit();
                // 출력 - 거래날짜 항목
                string tDate = outputDataTable.Rows[count]["TRANSACTION_DATE"].ToString();
                outputDataTable.Rows[count]["TRANSACTION_DATE"] = string2DateFormat(tDate);
                // 출력 - 수량
                outputDataTable.Rows[count]["TRANSACTION_NUM"] = 1;
                // 출력 - 원가
                int transactionIndex = Int32.Parse(outputDataTable.Rows[count]["TRANSACTION_INDEX"].ToString());
                outputTotalCost = totalCostOfTransaction(transactionIndex);
                outputDataTable.Rows[count]["TOTAL_COST"] = outputTotalCost.ToString("N0");
                // 출력 - 이익
                decimal supplyPrice = (decimal)outputDataTable.Rows[count]["SUPPLY_PRICE"];
                outputProfit = supplyPrice - outputTotalCost;
                outputDataTable.Rows[count]["PROFIT"] = outputProfit.ToString("N0");
                // 출력 - 부가세
                if (outputProfit < 0)
                {
                    outputProfit = 0;
                }
                outputTax = (Decimal)Math.Round(Decimal.ToDouble(outputProfit) * 0.1);
                outputDataTable.Rows[count]["TAX"] = outputTax.ToString("N0");
                // 출력 - 순이익
                outputDataTable.Rows[count]["FINAL_PROFIT"] = (outputProfit - outputTax).ToString("N0");
                // 출력 - 비고
                outputDataTable.Rows[count]["NOTE"] = noteOfTransaction(transactionIndex);
                outputDataTable.Rows[count].EndEdit();
                outputDataTable.AcceptChanges();
                count++;

            }

            return new DataView(outputDataTable);
        }

        public string string2DateFormat(String stringSource)
        {
            if (stringSource != null)
            {
                string year, month, day;
                year = stringSource.Substring(0, 4);
                month = stringSource.Substring(4, 2);
                day = stringSource.Substring(6, 2);

                return year + "-" + month + "-" + day;
            }
            else
            {
                return "0000-00-00";
            }
        }

        public int transactionItemCount(int transactionIndex)
        {
            return costItemDataTable.Select("TRANSACTION_INDEX = " + transactionIndex).Length;
        }

        public decimal totalCostOfTransaction(int transactionIndex)
        {
            object sumObject;
            sumObject = costItemDataTable.Compute("SUM(SUM)", "TRANSACTION_INDEX = " + transactionIndex);
            return Decimal.Parse(sumObject.ToString());
        }

        string[] customerNameArray;
        DataTable transactionNameTable;
        string[] transactionNameArray;
        string[] supplierArray;
        public string[] getCustomerNameList(int departMent)
        {
            DataTable customerNameTable = new DataTable();

            if (departMent != 999)
                customerNameTable = transactionDataTable.Select("DEPARTMENT = " + departMent).CopyToDataTable().DefaultView.ToTable(true, "CUSTOMER_NAME");
            else
                customerNameTable = transactionDataTable.DefaultView.ToTable(true, "CUSTOMER_NAME");

            customerNameArray = customerNameTable.AsEnumerable().Select(row => row.Field<string>("CUSTOMER_NAME")).ToArray();
            Array.Sort(customerNameArray);

            return customerNameArray;
        }

        public string noteOfTransaction(int transactionIndex)
        {
            string transactionNote = "";
            string addNote;

            DataRow[] cDataRow = costItemDataTable.Select("TRANSACTION_INDEX = " + transactionIndex, "COST_ITEM_INDEX");

            int count = 0;
            while (count < cDataRow.Length)
            {
                addNote = cDataRow[count].ItemArray[7].ToString();
                transactionNote += !addNote.Equals("") ? addNote+"\n":addNote;
                count++;
            }

            if(transactionNote.Length > 0)
            {
                transactionNote = transactionNote.Substring(0, transactionNote.LastIndexOf('\n'));
            }

            return transactionNote;
        }

        public string[] getTransactionNameList(string customerName, string yearAndMonth, int departMent)
        {
            try
            {
                if (customerNameArray.Contains(customerName))
                {
                    if(departMent != 999)
                        transactionNameTable = transactionDataTable
                            .Select("CUSTOMER_NAME = '" + customerName + "' AND DEPARTMENT = " + departMent)
                            .CopyToDataTable().DefaultView.ToTable(true, "TRANSACTION_NAME");
                    else
                        transactionNameTable = transactionDataTable
                            .Select("CUSTOMER_NAME = '" + customerName + "'")
                            .CopyToDataTable().DefaultView.ToTable(true, "TRANSACTION_NAME");

                    transactionNameArray = transactionNameTable.AsEnumerable().Select(row => row.Field<string>("TRANSACTION_NAME")).ToArray();
                    Array.Sort(transactionNameArray);

                    return transactionNameArray;
                }
                else
                {
                    return null;
                }
            }
            catch(Exception e)
            {
                return null;
            }
        }

        public string[] getAllSupplierList()
        {
            DataTable supplierNameTable = costItemDataTable.DefaultView.ToTable(true, "SUPPLIER");
            supplierArray = supplierNameTable.AsEnumerable().Select(row => row.Field<string>("SUPPLIER")).ToArray();
            Array.Sort(supplierArray);

            return supplierArray;
        }

        public void getTransactionIndex(string customerName, string transactionName, string currentDateMonth, int userDepartment,
            out int transactionIndex, out decimal supplyCost, out string transactionCode)
        {
            transactionIndex = -1;
            supplyCost = 0;
            transactionCode = "CODE01";

            DataRow[] tDataRow = transactionDataTable.Select("TRANSACTION_NAME = '" + transactionName + "' AND TRANSACTION_DATE LIKE '" + currentDateMonth + "%' AND CUSTOMER_NAME = '" + customerName + "' AND DEPARTMENT = "+userDepartment);

            if (tDataRow != null && tDataRow.Length > 0)
            {
                transactionIndex = Int32.Parse(tDataRow[0]["TRANSACTION_INDEX"].ToString());
                supplyCost = Decimal.Parse(tDataRow[0]["SUPPLY_PRICE"].ToString());
                transactionCode = tDataRow[0]["TRANSACTION_CODE"].ToString();
            }
        }

        public int maxValueOfTable(string tableName)
        {
            if (!DBConnect())
            {

            }
            else
            {
                string query;

                try
                {
                    if (tableName.Equals("TRANSACTION_TABLE"))
                    {
                        query = "SELECT MAX (TRANSACTION_INDEX) FROM " + tableName;
                    }
                    else
                    {
                        query = "SELECT MAX (COST_ITEM_INDEX) FROM " + tableName;
                    }
                    cmd.CommandText = query;
                    adapter.SelectCommand = cmd;

                    DataTable tmpTable = new DataTable();
                    adapter.Fill(tmpTable);

                    DBDisconnect();

                    if (tmpTable.Rows.Count > 0)
                    {
                        return Int32.Parse(tmpTable.Rows[0].ItemArray[0].ToString())+1;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.StackTrace);
                }
            }

            return -1;
        }
    }
}
