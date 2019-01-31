using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using System.Text.RegularExpressions;
using System.Configuration;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Drawing.Printing;

namespace Transaction
{
    public partial class MainForm : Form
    {
        private User connectedUser;
        private DataBaseModule dbModule;
        string inquiryDate;
        Authority connectedUserAuthority;
        bool newUserRegist = false;

        #region Enum
        public enum DateFormat { YearAndMonth = 1, FullDate = 2, NormalFormat = 3, OutputFormat = 4 };
        public enum Authority { NotUser = -1, Admin = 0, User = 1 };
        // 01 상품 02 제품 03 용역
        public enum TransactionCode { CODENONE = 0, CODE01 = 1, CODE02 = 2, CODE03 = 4, CODE0102 = 3, CODE0103 = 5, CODE0203 = 6, CODEALL = 7 };
        #endregion

        public MainForm()
        {
            InitializeComponent();
            connectedUser = new User();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoginForm lgForm = new LoginForm(connectedUser);
            DialogResult loginResult = lgForm.ShowDialog();
            connectedUserAuthority = (Authority)connectedUser.Authority;
            if (loginResult == DialogResult.OK)
            {
                dbModule = new DataBaseModule(connectedUser.Authority, connectedUser.Department);
                screenDrawAboutAuthority();
            }
            else
            {
                Application.Exit();
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(connectedUser.Authority != -1)
            {
                dbModule.Logout(connectedUser);
            }
        }

        /////////////////////////////////////////////////////////////////////////////

        #region 권한에 따른 화면에 보여질 폼 구성요소 설정
        public void screenDrawAboutAuthority()
        {
            if (Authority.NotUser.Equals(connectedUserAuthority))
            {
                newUserRegist = true;
                tabControl1.TabPages.RemoveAt(3);
                tabControl1.TabPages.RemoveAt(2);
                tabControl1.TabPages.RemoveAt(1);
                tabControl1.TabPages.RemoveAt(0);
            }
            else
            {
                tabControl1.TabPages.RemoveAt(4);

                toolStripStatusLabel1.Text = "로그인 계정 : " + connectedUser.Id + ",\t 부서 : " + connectedUser.DepartmentString() + ",\t 권한 : " + connectedUser.AuthorityString();
                toolStripStatusLabel2.Text = "로그인 계정 : " + connectedUser.Id + ",\t 부서 : " + connectedUser.DepartmentString() + ",\t 권한 : " + connectedUser.AuthorityString();
                toolStripStatusLabel3.Text = "로그인 계정 : " + connectedUser.Id + ",\t 부서 : " + connectedUser.DepartmentString() + ",\t 권한 : " + connectedUser.AuthorityString();
                toolStripStatusLabel4.Text = "로그인 계정 : " + connectedUser.Id + ",\t 부서 : " + connectedUser.DepartmentString() + ",\t 권한 : " + connectedUser.AuthorityString();

                if (Authority.User.Equals(connectedUserAuthority))
                {
                    tabControl1.TabPages.RemoveAt(1);
                    inquiryEditButton.Visible = false;
                    inquiryDeleteColUp.Visible = false;
                    inquiryDeleteColDown.Visible = false;

                    inquiryDataGridView2.AllowUserToAddRows = false;
                }
                else
                {
                    inputCustomerName.Items.AddRange(dbModule.getCustomerNameList(connectedUser.Department));
                }

                inquiryDate = date2StringFormat(this.inquiryDateTime.Value.ToString(), DateFormat.YearAndMonth);
                inquiryDataGridView1Table();
            }
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            TabPage current = (sender as TabControl).SelectedTab;

            if (current.Equals(tabPage1))
            {
                inquiryDate = date2StringFormat(this.inquiryDateTime.Value.ToString(), DateFormat.YearAndMonth);
                // inquiryDataGridView1.Rows.Clear();
                inquiryDataGridView2.Rows.Clear();
                inquiryDataGridView1Table();
            }
            else if (current.Equals(tabPage2))
            {
                inputCustomerName.Text = "";
                inputCustomerName.Items.Clear();
                inputCustomerName.Items.AddRange(dbModule.getCustomerNameList(connectedUser.Department));
                inputTransactionName.Text = "";
                inputTransactionCode.SelectedIndex = 0;
                inputSupplyPrice.Clear();
                inputDataGridView.Rows.Clear();

                if(connectedUser.Department == 999)
                {
                    inputDepart.Visible = true;
                    departLabel.Visible = true;
                    inputDepart.SelectedIndex = 0;
                }
            }
            else if (current.Equals(tabPage3))
            {
                selectedDepartment = connectedUser.DepartmentString();

                // 로그인 유저의 권한이 사용자(조회/입력) 이면 항목에 '전체' 항목을 추가
                if(connectedUser.Department == 999)
                {
                    outputDepartment.Items.Clear();
                    outputDepartment.Items.AddRange(new string[] { "전체", "관리영업", "기획/연구개발", "SI사업", "광주지역" });
                    outputDepartment.SelectedIndex = 0;
                }
                else
                {
                    outputDepartment.SelectedIndex = connectedUser.Department;
                }                
                // 달력의 최대 최소 범위 제한
                outputSearchStartDate.MaxDate = outputSearchEndDate.Value;
                outputSearchEndDate.MinDate = outputSearchStartDate.Value;
                
            }
            else if (current.Equals(tabPage4))
            {
                currentPassText.Clear();
                changePassText.Clear();
                removeUserPassText.Clear();

                if (Authority.Admin.Equals(connectedUserAuthority))
                    changeAdminRadioButton.Checked = true;
                else
                    changeUserRadioButton.Checked = true;

                switch (connectedUser.Department)
                {
                    case 0:
                        changeDepartment1.Checked = true;
                        break;
                    case 1:
                        changeDepartment2.Checked = true;
                        break;
                    case 2:
                        changeDepartment3.Checked = true;
                        break;
                    case 3:
                        changeDepartment4.Checked = true;
                        break;
                    case 999:
                        changeDepartment999.Checked = true;
                        changeDepartment999.Visible = true;
                        tableLayoutPanel19.Enabled = false;
                        break;
                }
            }
            else if (current.Equals(tabPage5))
            {
                newUserRadioButton.Checked = true;
                newUserDepartment1.Checked = true;
            }
        }
        #endregion

        /////////////////////////////////////////////////////////////////////////////

        #region tab page1 - 조회 화면

        private int currentClickedTransactionIndex = -1;
        private bool inquiryDataGridView1LoadComplete = false;
        private int transactionCode = 7;
        private string transactionCodeQuery = "";

        private void transactionCodeCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (sender == this.tCodeCheckBox1)
            {
                if (tCodeCheckBox1.CheckState == CheckState.Checked)
                {
                    tCodeCheckBox2.Checked = true;
                    tCodeCheckBox3.Checked = true;
                    tCodeCheckBox4.Checked = true;
                    transactionCode = 7;
                }
                else if (tCodeCheckBox1.CheckState == CheckState.Unchecked)
                {
                    tCodeCheckBox2.Checked = false;
                    tCodeCheckBox3.Checked = false;
                    tCodeCheckBox4.Checked = false;
                    transactionCode = 0;
                }
            }
            else if (sender == this.tCodeCheckBox2)
            {
                if (tCodeCheckBox2.CheckState == CheckState.Checked && transactionCode < 7)
                {
                    transactionCode += 1;
                }
                else if (tCodeCheckBox2.CheckState == CheckState.Unchecked && transactionCode > 0)
                {
                    transactionCode -= 1;
                }
            }
            else if (sender == this.tCodeCheckBox3)
            {
                if (tCodeCheckBox3.CheckState == CheckState.Checked && transactionCode < 7)
                {
                    transactionCode += 2;
                }
                else if (tCodeCheckBox3.CheckState == CheckState.Unchecked && transactionCode > 0)
                {
                    transactionCode -= 2;
                }
            }
            else if (sender == this.tCodeCheckBox4)
            {
                if (tCodeCheckBox4.CheckState == CheckState.Checked && transactionCode < 7)
                {
                    transactionCode += 4;
                }
                else if (tCodeCheckBox4.CheckState == CheckState.Unchecked && transactionCode > 0)
                {
                    transactionCode -= 4;
                }
            }

            if (transactionCode == 7)
            {
                tCodeCheckBox1.CheckState = CheckState.Checked;
            }
            else if (transactionCode == 0)
            {
                tCodeCheckBox1.CheckState = CheckState.Unchecked;
            }
            else
            {
                tCodeCheckBox1.CheckState = CheckState.Indeterminate;
            }

            inquiryDataGridView2.DataSource = null;
            inquiryDataGridView2.Rows.Clear();

            inquiryDataGridView1Table();
        }

        public void setTransactionCodeQuery()
        {
            if (transactionCode == 0)
            {
                transactionCodeQuery = " AND TRANSACTION_CODE = NULL";
                return;
            }
            else if (transactionCode == 7)
            {
                transactionCodeQuery = "";
                return;
            }


            switch ((TransactionCode)transactionCode)
            {
                case TransactionCode.CODE01:
                    transactionCodeQuery = " AND TRANSACTION_CODE = '상품'";
                    break;
                case TransactionCode.CODE02:
                    transactionCodeQuery = " AND TRANSACTION_CODE = '제품'";
                    break;
                case TransactionCode.CODE03:
                    transactionCodeQuery = " AND TRANSACTION_CODE = '용역'";
                    break;
                case TransactionCode.CODE0102:
                    transactionCodeQuery = " AND TRANSACTION_CODE IN ('상품', '제품')";
                    break;
                case TransactionCode.CODE0103:
                    transactionCodeQuery = " AND TRANSACTION_CODE IN ('상품', '용역')";
                    break;
                case TransactionCode.CODE0203:
                    transactionCodeQuery = " AND TRANSACTION_CODE IN ('제품', '용역')";
                    break;
            }
        }

        public void inquiryDataGridView1Table()
        {
            DataView dv = new DataView();
            setTransactionCodeQuery();

            if(connectedUser.Department != 999)
                dv = new DataView(dbModule.DepartmentTransactionDataTable, "TRANSACTION_DATE like '" + inquiryDate + "%'" + transactionCodeQuery, "TRANSACTION_DATE", DataViewRowState.CurrentRows);
            else
                dv = new DataView(dbModule.TransactionDataTable, "TRANSACTION_DATE like '" + inquiryDate + "%'" + transactionCodeQuery, "TRANSACTION_DATE", DataViewRowState.CurrentRows);

            // 데이터 소스 추가 전에 자동으로 열 생성해서 추가하는 것을 끄고
            inquiryDataGridView1.AutoGenerateColumns = false;
            inquiryDataGridView1.DataSource = dv;

            // 열을 지정해서 데이터 소스의 해당 열의 내용을 넣도록 설정. 쿼리로 처리한 데이터 테이블을 DataSource로 지정한 뒤에 열의 이름을 넣어야 함.
            inquiryInvisibleIndexCol.DataPropertyName = "TRANSACTION_INDEX";
            inquiryTransactionCodeCol.DataPropertyName = "TRANSACTION_CODE";
            inquiryTransactionDateCol.DataPropertyName = "TRANSACTION_DATE";
            inquiryCustomerNameCol.DataPropertyName = "CUSTOMER_NAME";
            inquiryTransactionNameCol.DataPropertyName = "TRANSACTION_NAME";
            inquirySupplyPriceCol.DataPropertyName = "SUPPLY_PRICE";

            inquiryDataGridView1LoadComplete = true;
            inquiryDataGridView1.ClearSelection();
        }

        public void inquiryDataGridView2Table(int transactionIndex)
        {
            this.inquiryDataGridView2.Rows.Clear();

            DataView dv = new DataView(dbModule.CostItemDataTable, "TRANSACTION_INDEX = " + transactionIndex, "COST_ITEM_INDEX", DataViewRowState.CurrentRows);

            inquiryDataGridView2.AutoGenerateColumns = false;
            
            // 데이터 소스를 수동으로 작업할 경우 -> inputGrid도 아래 방법으로 해야 작업이 가능할 것으로 예상됨.

            int count = 0;
            while (count < dv.Count)
            {
                DataRow dr = dv[count].Row;
                inquiryDataGridView2.Rows.Add(count + 1, dr.ItemArray[0], dr.ItemArray[1], dr.ItemArray[2], dr.ItemArray[3], dr.ItemArray[7]);
                count++;
            }
        }

        private void inquiryDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (e.ColumnIndex == inquiryDeleteColUp.Index && e.RowIndex >= 0)
            {
                if (MessageBox.Show("정말 삭제하시겠습니까?", "Confirm delete", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    // 지우기 전 행의 특정 값(인덱스)을 객체로 빼내어 DB를 연결하고 해당 값을 갖는 행을 DB에서 삭제처리
                    int delTransactionIndex = Int32.Parse(inquiryDataGridView1.Rows[e.RowIndex].Cells[inquiryInvisibleIndexCol.Index].Value.ToString());
                    Transaction delTrans = new Transaction(delTransactionIndex);

                    inquiryDataGridView1.Rows.RemoveAt(e.RowIndex);
                    dbModule.Delete(delTrans);

                    inquiryDate = date2StringFormat(this.inquiryDateTime.Value.ToString(), DateFormat.YearAndMonth);
                    inquiryDataGridView1Table();

                    inquiryDataGridView2.DataSource = null;
                    inquiryDataGridView2.Rows.Clear();
                }
            }
        }

        private void inquiryDataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                int transactionIndex = Int32.Parse(inquiryDataGridView1.Rows[e.RowIndex].Cells[inquiryInvisibleIndexCol.Index].Value.ToString());
                currentClickedTransactionIndex = transactionIndex;
                inquiryDataGridView2Table(transactionIndex);
            }
        }

        private void inquiryDataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex < inquiryDataGridView2.RowCount - 1)
            {
                if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0)
                {
                    if (MessageBox.Show("정말 삭제하시겠습니까?", "Confirm delete", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        string cIndexString = inquiryDataGridView2.Rows[e.RowIndex].Cells[inquiryInvisibleCIndexCol.Index].FormattedValue.ToString();
                        string tIndexString = inquiryDataGridView2.Rows[e.RowIndex].Cells[inquiryInvisibleTIndexCol.Index].FormattedValue.ToString();
                        int tIndex = -1;
                        bool tIndexParseOk = Int32.TryParse(tIndexString, out tIndex);

                        // DB에 등록 되어 있는 값을 지울 경우 DB까지 접근하여 삭제해야함.
                        if (!cIndexString.Equals("") || tIndexParseOk)
                        {
                            // 지우기 전 행의 특정 값(인덱스)을 객체로 빼내어 DB를 연결하고 해당 값을 갖는 행을 DB에서 삭제처리
                            int cIndex = Int32.Parse(cIndexString);
                            CostItem delCostItem = new CostItem(cIndex);
                            dbModule.Delete(delCostItem);
                        }

                        inquiryDataGridView2.Rows.RemoveAt(e.RowIndex);
                        
                        // 원가 항목이 모두 삭제 된 경우 거래내역은 자동적으로 삭제 됨
                        if (dbModule.transactionItemCount(tIndex) == 0)
                        {
                            Transaction delTransaction = new Transaction(tIndex);
                            dbModule.Delete(delTransaction);
                        }

                        int colIndex = inquiryDataGridView1.CurrentCell.ColumnIndex;  
                        int rowIndex = inquiryDataGridView1.CurrentCell.RowIndex;
                        inquiryDataGridView1Table();
                        inquiryDataGridView1.Rows[rowIndex].Cells[colIndex].Selected = true;
                        inquiryDataGridView1.CurrentCell = inquiryDataGridView1.Rows[rowIndex].Cells[colIndex];
                        inquiryDataGridView1.FirstDisplayedCell = inquiryDataGridView1.Rows[rowIndex].Cells[colIndex];
                    }
                }
            }
        }

        private void inquiryDateTime_ValueChanged(object sender, EventArgs e)
        {
            inquiryDate = date2StringFormat(this.inquiryDateTime.Value.ToString(), DateFormat.YearAndMonth);
            inquiryDataGridView1Table();
        }

        private void inquiryEditButton_Click(object sender, EventArgs e)
        {
            if (currentClickedTransactionIndex == -1)
            {
                MessageBox.Show("변경할 원가항목의 계약명을 먼저 선택해주세요.", "Error", MessageBoxButtons.OK);
                return;
            }

            bool updateOk = false;
            bool parseOk = true;

            string testCIndex, supplier, note;
            int[] sum = new int[inquiryDataGridView2.RowCount];
            int cIndex;

            // sum 항목의 입력값 유효성 검사를 먼저 수행 후 통과되면 update를 시작
            int count = 0;
            while (count < inquiryDataGridView2.RowCount - 1)
            {
                parseOk = Int32.TryParse(inquiryDataGridView2.Rows[count].Cells[inquirySumCol.Index].Value.ToString(), out sum[count]);

                if (!parseOk)
                {
                    MessageBox.Show("금액 항목에 유효한 값을 입력해주세요.", "Error", MessageBoxButtons.OK);
                    return;
                }

                count++;
            }

            count = 0;
            while (parseOk && count < inquiryDataGridView2.RowCount - 1)
            {
                // 값에 null 값이 예상될 경우 formattedValue를 쓰면 ""으로 치환
                // sum과 같이 확실하게 숫자가 올 경우는 value를 사용하여 값을 가져온다.
                testCIndex = inquiryDataGridView2.Rows[count].Cells[inquiryInvisibleCIndexCol.Index].FormattedValue.ToString();
                supplier = inquiryDataGridView2.Rows[count].Cells[inquirySupplierCol.Index].FormattedValue.ToString();
                note = inquiryDataGridView2.Rows[count].Cells[inquiryNoteCol.Index].FormattedValue.ToString();

                // cIndex가 있는 값은 업데이트를, 없는 값은 insert를 실행
                if (!testCIndex.Equals(""))
                {
                    cIndex = Int32.Parse(testCIndex);
                    updateOk = dbModule.Update(new CostItem(cIndex, currentClickedTransactionIndex, supplier, sum[count], note));
                }
                else
                {
                    cIndex = dbModule.maxValueOfTable("COST_ITEM_TABLE");
                    updateOk = dbModule.Insert(new CostItem(cIndex, currentClickedTransactionIndex, supplier, sum[count], note));
                }

                count++;
            }

            if (updateOk)
                MessageBox.Show("변경사항이 저장 되었습니다.", "Save", MessageBoxButtons.OK);
            else
                MessageBox.Show("변경사항이 저장 중 에러가 발생했습니다.", "Error", MessageBoxButtons.OK);

            inquiryDataGridView1Table();
            inquiryDataGridView2Table(currentClickedTransactionIndex);
        }

        #endregion

        /////////////////////////////////////////////////////////////////////////////

        #region tab page2 - 입력 화면. 관리자 권한에서 작업
        private decimal totalSum = 0;
        private int transactionIndex = -1;

        // 콤보 박스에서 선택한 거래내역의 정보를 저장
        int storedTransactionIndex = -1;
        decimal storedSupplyCost = 0;
        string storedTCode = "상품";

        private void inputCustomerName_TextChanged(object sender, EventArgs e)
        {
            transactionIndex = -1;
            inputTransactionName.Text = "";
            inputSupplyPrice.Text = "";
            AutoCompleteTextboxResult();

            string inputCName = inputCustomerName.Text.ToString();
            string searchDate = date2StringFormat(inputTransactionDate.Value.ToString(), DateFormat.YearAndMonth);
            int userDepartment = connectedUser.Department;

            string[] transactionNameList = dbModule.getTransactionNameList(inputCName, searchDate, userDepartment);

            if (transactionNameList != null)
            {
                inputTransactionName.Items.Clear();
                inputTransactionName.Items.AddRange(transactionNameList);
            }
        }

        private void inputTransactionName_Changed(object sender, EventArgs e)
        {
            transactionIndex = -1;
            inputSupplyPrice.Text = "";
            AutoCompleteTextboxResult();

            string customerName = inputCustomerName.Text.ToString();
            string transactionName = inputTransactionName.Text.ToString();
            string currentYearAndMonth = date2StringFormat(inputTransactionDate.Value.ToString(), DateFormat.YearAndMonth);
            string fullDate = date2StringFormat(inputTransactionDate.Value.ToString(), DateFormat.FullDate);
            int userDepartment = connectedUser.Department;

            dbModule.getTransactionIndex(customerName, transactionName, fullDate, userDepartment, out transactionIndex, out storedSupplyCost, out storedTCode);

            inputDataGridViewTable();
        }

        private void inputTransactionDate_ValueChanged(object sender, EventArgs e)
        {
            inputSupplyPrice.Text = "";

            string customerName = inputCustomerName.Text.ToString();
            string transactionName = inputTransactionName.Text.ToString();
            string currentYearAndMonth = date2StringFormat(inputTransactionDate.Value.ToString(), DateFormat.YearAndMonth);
            string fullDate = date2StringFormat(inputTransactionDate.Value.ToString(), DateFormat.FullDate);
            int userDepartment = connectedUser.Department;

            string[] transactionNameList = dbModule.getTransactionNameList(customerName, currentYearAndMonth, userDepartment);

            if (transactionNameList != null)
            {
                inputTransactionName.Items.Clear();
                inputTransactionName.Items.AddRange(transactionNameList);
            }
            dbModule.getTransactionIndex(customerName, transactionName, fullDate, userDepartment, out transactionIndex, out storedSupplyCost, out storedTCode);

            inputDataGridViewTable();
        }

        private void inputDataGridViewTable()
        {
            // 기존의 거래 내역을 가져올 경우
            if (transactionIndex != -1)
            {
                // 불러온 기존 거래기록의 거래코드 설정
                if (TransactionCode.CODE01.ToString().Equals(storedTCode))
                {
                    inputTransactionCode.SelectedIndex = 0;
                }
                else if (TransactionCode.CODE02.ToString().Equals(storedTCode))
                {
                    inputTransactionCode.SelectedIndex = 1;
                }
                else
                {
                    inputTransactionCode.SelectedIndex = 2;
                }

                // update 준비. datagrid 정보를 가져온다

                inputSupplyPrice.Text = storedSupplyCost.ToString();

                this.inputDataGridView.Rows.Clear();

                DataView dv = new DataView(dbModule.CostItemDataTable, "TRANSACTION_INDEX = " + transactionIndex, "COST_ITEM_INDEX", DataViewRowState.CurrentRows);

                inputDataGridView.AutoGenerateColumns = false;
                                
                int count = 0;
                while (count < dv.Count)
                {
                    DataRow dr = dv[count].Row;
                    inputDataGridView.Rows.Add(count + 1, dr.ItemArray[0], dr.ItemArray[1], dr.ItemArray[2], dr.ItemArray[3], dr.ItemArray[7]);
                    count++;
                }
                
            }
            // 새로운 거래 내역을 작성할 경우
            else
            {
                //insert 준비.
                inputSupplyPrice.Clear();
                inputTransactionCode.SelectedIndex = 0;
                this.inputDataGridView.Rows.Clear();
            }

            totalSum = totalInputSum();
            AutoCompleteTextboxResult();
        }

        private void inputDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex < inputDataGridView.RowCount - 1)
            {
                if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0)
                {
                    inputDataGridView.Rows.RemoveAt(e.RowIndex);
                    totalSum = totalInputSum();
                    AutoCompleteTextboxResult();
                }
            }
        }

        private void inputSupplyPrice_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void inputSupplyPrice_TextChanged(object sender, EventArgs e)
        {
            // 공급가 입력 값의 유효성 검사
            // 포커스 유지 중 일때는 입력값이 유효하면 다른 값들이 연동 되어 변하도록
            // 포커스가 사라질 때(leave) 입력값의 유효성을 검사
            AutoCompleteTextboxResult();
            inputSupplyPrice.Select(inputSupplyPrice.Text.Length, 0);
        }

        private void inputDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // 매입처 별 금액의 입력 값의 유효성 검사
            if (e.ColumnIndex == inputSumCol.Index && e.RowIndex >= 0)
            {
                decimal parseResult = 0;
                bool vaildInput = decimal.TryParse(inputDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), out parseResult);

                if (vaildInput)
                {
                    totalSum = totalInputSum();
                    inputDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = parseResult.ToString("N0");
                    AutoCompleteTextboxResult();
                }
                else
                {
                    MessageBox.Show("유효한 범위의 값을 입력해주세요.", "Error", MessageBoxButtons.OK);
                    inputDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 0;
                }
            }
        }

        private decimal totalInputSum()
        {
            bool isDecimal = false;
            decimal returnVal = 0, parseVal = 0;
            int count = 0;

            while (count < inputDataGridView.RowCount - 1)
            {
                isDecimal = Decimal.TryParse(inputDataGridView.Rows[count].Cells[inputSumCol.Index].FormattedValue.ToString(), out parseVal);
                if (isDecimal)
                {
                    returnVal += parseVal;
                }
                count++;
            }

            return returnVal;
        }

        private void AutoCompleteTextboxResult()
        {
            decimal parseResult;
            bool vaildInput = Decimal.TryParse(inputSupplyPrice.Text.ToString(), out parseResult);

            if (vaildInput)
            {
                inputSupplyPrice.Text = parseResult.ToString("N0");
                inputCost.Text = totalSum.ToString("N0");   // 원가 합계 (데이터 그리드에서 원가 합계 뽑아내는 함수 구현해서 사용)

                decimal profit = (parseResult - totalSum);
                inputProfit.Text = profit.ToString("N0"); // 공급가 - 원가합계 = 이익

                if (profit > 0)
                {
                    inputTax.Text = Math.Round((Decimal.ToDouble(profit) * 0.1)).ToString("N0");    // 이익*0.1
                }
                else
                {
                    inputTax.Text = "0";
                }
            }
            else
            {
                inputCost.Text = "0";
                inputProfit.Text = "0";
                inputTax.Text = "0";
            }
        }

        private void inputSaveButton_Click(object sender, EventArgs e)
        {
            // 기존에 있던 값을 저장시에는 업데이트
            // 새로운 값을 저장시에는 insert
            // 위 두가지 경우를 판별하도록 bool변수 하나 설정
            string customerName = inputCustomerName.Text.ToString();
            string transactionName = inputTransactionName.Text.ToString();
            string transactionDate = date2StringFormat(inputTransactionDate.Value.ToString(), DateFormat.FullDate);

            if (inputCustomerName.Text.ToString().Equals(""))
            {
                MessageBox.Show("거래처명이 공란입니다.", "Error", MessageBoxButtons.OK);
                return;
            }
            else if (inputTransactionName.Text.ToString().Equals(""))
            {
                MessageBox.Show("계약명이 공란입니다.", "Error", MessageBoxButtons.OK);
                return;
            }
            else if (inputSupplyPrice.Text.ToString().Equals("") || inputSupplyPrice.Text.ToString().Equals("0"))
            {
                MessageBox.Show("계약금액이 잘못 입력되었습니다.", "Error", MessageBoxButtons.OK);
                return;
            }
            else if (inputDataGridView.RowCount <= 1)
            {
                MessageBox.Show("원가 항목의 항목이 비어있습니다.", "Error", MessageBoxButtons.OK);
                return;
            }

            decimal supplyPrice = Decimal.Parse(inputSupplyPrice.Text.ToString());
            storedSupplyCost = supplyPrice;

            bool saveResult = false;

            if (transactionIndex == -1)
            {
                transactionIndex = dbModule.maxValueOfTable("TRANSACTION_TABLE");
                if(connectedUser.Department != 999)
                    saveResult = dbModule.Insert(new Transaction(transactionIndex, customerName, transactionName, transactionDate, supplyPrice, connectedUser.Department, inputTransactionCode.Text.ToString()));
                else
                    saveResult = dbModule.Insert(new Transaction(transactionIndex, customerName, transactionName, transactionDate, supplyPrice, inputDepart.SelectedIndex, inputTransactionCode.Text.ToString()));
                saveResult = saveDataGridView();
            }
            else
            {
                if (connectedUser.Department != 999)
                    saveResult = dbModule.Update(new Transaction(transactionIndex, customerName, transactionName, transactionDate, supplyPrice, connectedUser.Department, inputTransactionCode.Text.ToString()));
                else
                    saveResult = dbModule.Update(new Transaction(transactionIndex, customerName, transactionName, transactionDate, supplyPrice, inputDepart.SelectedIndex, inputTransactionCode.Text.ToString()));
                saveResult = saveDataGridView();
            }

            if (saveResult)
            {
                MessageBox.Show("입력 내용이 저장 되었습니다.", "Save Success", MessageBoxButtons.OK);
                inputCustomerName.Items.Clear();
                inputCustomerName.Items.AddRange(dbModule.getCustomerNameList(connectedUser.Department));

                inputDataGridViewTable();
            }
            else
            {
                MessageBox.Show("입력 내용이 저장에 실패했습니다.", "Save Fail", MessageBoxButtons.OK);
                inputCustomerName.Items.Clear();
                inputCustomerName.Items.AddRange(dbModule.getCustomerNameList(connectedUser.Department));
                inputDataGridViewTable();
            }

        }

        private bool saveDataGridView()
        {
            filledDataGridViewBeforeInsert();

            bool savaResult = false;
            int count = 0;

            while (count < inputDataGridView.RowCount - 1)
            {
                string invisibleTIndex = inputDataGridView.Rows[count].Cells[inputInvisibleTIndexCol.Index].FormattedValue.ToString();
                string supplier = inputDataGridView.Rows[count].Cells[inputSupplierCol.Index].FormattedValue.ToString();
                decimal sum = Decimal.Parse(inputDataGridView.Rows[count].Cells[inputSumCol.Index].Value.ToString());
                string note = inputDataGridView.Rows[count].Cells[inputNoteCol.Index].FormattedValue.ToString();

                if (invisibleTIndex.Equals(""))
                {
                    int costItemIndex = dbModule.maxValueOfTable("COST_ITEM_TABLE");
                    savaResult = dbModule.Insert(new CostItem(costItemIndex, transactionIndex, supplier, sum, note));
                }
                else
                {
                    int costItemIndex = Int32.Parse(inputDataGridView.Rows[count].Cells[inputInvisibleCIndexCol.Index].Value.ToString());
                    savaResult = dbModule.Update(new CostItem(costItemIndex, transactionIndex, supplier, sum, note));
                }

                count++;
            }

            return savaResult;
        }

        private void CleaarInputPage()
        {
            inputCustomerName.Items.Clear();
            inputCustomerName.Items.AddRange(dbModule.getCustomerNameList(connectedUser.Department));
            inputDataGridView.Rows.Clear();
            inputCustomerName.Text = "";
            inputTransactionName.Text = "";
            inputTransactionCode.SelectedIndex = 0;
            inputSupplyPrice.Text = "";
            AutoCompleteTextboxResult();
        }

        private void filledDataGridViewBeforeInsert()
        {
            int count = 0;

            while (count < inputDataGridView.RowCount - 1)
            {
                string sum = inputDataGridView.Rows[count].Cells[inputSumCol.Index].FormattedValue.ToString();

                if (sum.Equals(""))
                {
                    inputDataGridView.Rows[count].Cells[inputSumCol.Index].Value = "0";
                }
                count++;
            }
        }

        #endregion

        /////////////////////////////////////////////////////////////////////////////

        #region tab page3 - 출력 화면
        private string searchStartDate, searchEndDate;
        private string selectedDepartment;
        private bool outputDataGridViewLoadComplete = false;

        private void outputSearchDate_ValueChanged(object sender, EventArgs e)
        {
            outputSearchStartDate.MaxDate = outputSearchEndDate.Value;
            outputSearchEndDate.MinDate = outputSearchStartDate.Value;
            outputDataGridViewTable();
        }

        private void outputDepartment_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedDepartment = outputDepartment.SelectedItem.ToString();
            if (selectedDepartment.Equals(connectedUser.DepartmentString()) || selectedDepartment.Equals("전체"))
            {
                outputDataGridViewTable();
            }
            else if (connectedUser.Department == 999)
            {
                outputDataGridViewTable();
            }
            else
            {
                MessageBox.Show("해당 부서에 접근 권한이 없습니다.", "Access Failed", MessageBoxButtons.OK);
                outputDepartment.SelectedIndex = connectedUser.Department;
                outputDataGridView.Rows.Clear();
            }
            
        }

        public string setDepartmentQuery()
        {
            string departmentQuery = "";

            switch (selectedDepartment)
            {
                case "관리영업":
                    departmentQuery = " AND DEPARTMENT = 0";
                    break;
                case "기획/연구개발":
                    departmentQuery = " AND DEPARTMENT = 1";
                    break;
                case "SI사업":
                    departmentQuery = " AND DEPARTMENT = 2";
                    break;
                case "광주지역":
                    departmentQuery = " AND DEPARTMENT = 3";
                    break;
                default:
                    break;
            }

            return departmentQuery;
        }

        public void outputDataGridViewTable()
        {
            searchStartDate = date2StringFormat(outputSearchStartDate.Value.ToString(), DateFormat.FullDate);
            searchEndDate = date2StringFormat(outputSearchEndDate.Value.ToString(), DateFormat.FullDate);

            string dataViewQuery = "(TRANSACTION_DATE >= '" + searchStartDate + "' AND TRANSACTION_DATE <= '" + searchEndDate + "')" + setDepartmentQuery();

            DataView dv = dbModule.getOutputTable(dataViewQuery, "TRANSACTION_DATE");

            // 데이터 소스 추가 전에 자동으로 열 생성해서 추가하는 것을 끄고
            outputDataGridView.AutoGenerateColumns = false;
            outputDataGridView.DataSource = dv;

            // 열을 지정해서 데이터 소스의 해당 열의 내용을 넣도록 설정. 쿼리로 처리한 데이터 테이블을 DataSource로 지정한 뒤에 열의 이름을 넣어야 함.
            outputTransactionDateCol.DataPropertyName = "TRANSACTION_DATE";
            outputCustomerNameCol.DataPropertyName = "CUSTOMER_NAME";
            outputTransactionNameCol.DataPropertyName = "TRANSACTION_NAME";
            outputSupplyPriceCol.DataPropertyName = "SUPPLY_PRICE";
            outputNumCol.DataPropertyName = "TRANSACTION_NUM";
            outputCostCol.DataPropertyName = "TOTAL_COST";
            outputProfitCol.DataPropertyName = "PROFIT";
            outputTaxCol.DataPropertyName = "TAX";
            outputFinalProfitCol.DataPropertyName = "FINAL_PROFIT";
            outputNoteCol.DataPropertyName = "NOTE";

            outputDataGridViewLoadComplete = true;
        }

        private void outputButton_Click(object sender, EventArgs e)
        {
            if (sender == outputPrintButton)
            {
                
            }
            else if (sender == outputXlsButton)
            {
                ExportExcel();
            }
        }

        SaveFileDialog saveFileDialog = new SaveFileDialog();
        private void ExportExcel()
        {
            this.saveFileDialog.FileName = "결산서";
            this.saveFileDialog.DefaultExt = "xls";
            this.saveFileDialog.Filter = "Excel files (*.xls)|*.xls";

            DialogResult result = saveFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                int num = 0;
                int colCount = outputDataGridView.ColumnCount - 4;
                object missingType = Type.Missing;

                Excel.Application objApp;
                Excel._Workbook objBook;
                Excel.Workbooks objBooks;
                Excel.Sheets objSheets;
                Excel._Worksheet objSheet;
                Excel.Range range;

                string[] headers = new string[colCount];
                string[] columns = new string[colCount];

                for (int c = 0; c < colCount; c++)
                {
                        headers[c] = outputDataGridView.Rows[0].Cells[c].OwningColumn.HeaderText.ToString();
                        num = c + 65;
                        columns[c] = Convert.ToString((char)num);
                }

                try
                {
                    objApp = new Excel.Application();
                    objBooks = objApp.Workbooks;
                    objBook = objBooks.Add(Missing.Value);
                    objSheets = objBook.Worksheets;
                    objSheet = (Excel._Worksheet)objSheets.get_Item(1);
                    
                    // column의 헤더 값을 엑셀 제일 상단에 출력
                    for (int c = 0; c < colCount; c++)
                    {                        
                        range = objSheet.get_Range(columns[c] + "1", Missing.Value);
                        range.set_Value(Missing.Value, headers[c]);
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        range.Borders.ColorIndex = 1;
                        objSheet.Cells[1, c + 1].Interior.ColorIndex = 31;
                    }
                    
                    // 각 row의 값을 포맷과 형식에 맞게 채움
                    for (int i = 0; i < outputDataGridView.RowCount; i++)
                    {
                        for (int j = 0; j < colCount; j++)
                        {
                            range = objSheet.get_Range(columns[j] + Convert.ToString(i + 2), Missing.Value);
                            range.set_Value(Missing.Value, outputDataGridView.Rows[i].Cells[j].Value.ToString());
                            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            range.Borders.ColorIndex = 1;

                            if(j < 4)
                            {
                                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                if (j < 1)
                                {
                                    range.NumberFormat = "mm\"월\"dd\"일\"";
                                }
                            }
                                
                            else if(j < 9)
                            {
                                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                range.NumberFormat = "#,###";
                            }                                
                            else
                                range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                        }
                    }
                    
                    // 헤더 값 bold 처리와 셀 크기 설정
                    objSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                    for (int i = 0; i < colCount; i++)
                    {
                        //if (i < colCount - 1)
                            objSheet.Cells[1, i + 1].EntireColumn.AutoFit();
                        //else
                            //((Excel.Range)objSheet.Cells[1, i + 1]).ColumnWidth = 20;
                    }

                    objApp.Visible = false;
                    objApp.UserControl = false;
                    objApp.DisplayAlerts = false;

                    objBook.SaveAs(@saveFileDialog.FileName,
                              Excel.XlFileFormat.xlWorkbookNormal,
                              missingType, missingType, missingType, missingType,
                              Excel.XlSaveAsAccessMode.xlNoChange,
                              missingType, missingType, missingType, missingType, missingType);
                    objBook.Close(false, missingType, missingType);

                    Cursor.Current = Cursors.Default;

                    MessageBox.Show("Save Success!");
                }
                catch (Exception theException)
                {
                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);

                    MessageBox.Show(errorMessage, "Error");
                }
            }
        }
        #endregion

        /////////////////////////////////////////////////////////////////////////////

        #region tab page 4,5 - 신규 사용자 ID 중복체크 및 사용자 등록 관련 함수
        private string IdPattern = @"^[a-zA-Z0-9]*$";
        private string PwPattern = @"^[a-zA-Z0-9`~!@#$%^&*()-_=+,<.>?;:{}\\\/\|\[\]\''\""]*.{6}$";

        // 새 계정 등록
        private bool availableId = false, availablePw = false, duplicationCheck = false, equalPass = false;

        private void newUserText_TextChanged(object sender, EventArgs e)
        {
            // ID 입력
            if (sender == this.newUserIdText)
            {
                string userId = this.newUserIdText.Text.ToString();
                this.availableId = Regex.IsMatch(this.newUserIdText.Text.ToString(), IdPattern);
                if (availableId)
                {
                    this.AlertId.Visible = false;
                }
                else
                {
                    this.AlertId.Visible = true;
                }

                if (duplicationCheck)
                {
                    this.duplicationCheck = false;
                    this.duplicationCheckButton.Enabled = true;
                    this.duplicationCheckButton.ForeColor = Color.Black;
                }
            }
            // PW 입력
            else
            {
                string userPw = this.newUserPassText.Text.ToString();
                string userPwConfirm = this.newUserPassTextConfirm.Text.ToString();

                if (Regex.IsMatch(this.newUserPassText.Text.ToString(), PwPattern))
                {
                    this.AlertPass.Visible = false;
                    this.availablePw = true;
                }
                else
                {
                    this.AlertPass.Visible = true;
                    this.availablePw = false;
                }

                if (userPwConfirm.Equals(userPw))
                {
                    this.AlertConfirmPass.Visible = false;
                    this.equalPass = true;
                }
                else
                {
                    this.AlertConfirmPass.Visible = true;
                    this.equalPass = false;
                }
                // 등록용 PW 설정
                if (availablePw == equalPass)
                {
                    this.connectedUser.Password = userPw;
                }
            }
        }

        private void duplicationCheckButton_Click(object sender, EventArgs e)
        {
            if (availableId)
            {
                duplicationCheck = dbModule.isRegisteredUser(this.newUserIdText.Text.ToString());

                if (duplicationCheck)
                {
                    MessageBox.Show("사용 가능한 ID 입니다.", "Success", MessageBoxButtons.OK);
                    this.AlertId.Visible = false;
                    this.duplicationCheckButton.Enabled = false;
                    this.duplicationCheckButton.ForeColor = Color.Gray;
                    // 등록용 ID 설정
                    this.connectedUser.Id = this.newUserIdText.Text.ToString();
                }
                else
                {
                    MessageBox.Show("이미 사용 중인 ID 입니다.", "Error", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("ID 조건을 확인해주세요.", "Error", MessageBoxButtons.OK);
                this.AlertId.Visible = true;
            }
        }

        private void radioButtons_CheckChanged(object sender, EventArgs e)
        {
            if (newUserRegist)
            {
                // 등록용 사용자 권한 설정
                connectedUser.Authority = newAdminRadioButton.Checked ? 0 : 1;

                // 등록용 사용자 부서 설정
                if (newUserDepartment1.Checked)
                {
                    connectedUser.Department = 0;
                }
                else if (newUserDepartment2.Checked)
                {
                    connectedUser.Department = 1;
                }
                else if (newUserDepartment3.Checked)
                {
                    connectedUser.Department = 2;
                }
                else if (newUserDepartment4.Checked)
                {
                    connectedUser.Department = 3;
                }
            }
            else
            {
                // 사용자 권한
                changeAuthority = changeAdminRadioButton.Checked ? 0 : 1;

                // 부서
                if (changeDepartment1.Checked)
                {
                    changeDepartment = 0;
                }
                else if (changeDepartment2.Checked)
                {
                    changeDepartment = 1;
                }
                else if (changeDepartment3.Checked)
                {
                    changeDepartment = 2;
                }
                else if (changeDepartment4.Checked)
                {
                    changeDepartment = 3;
                }
                else if (changeDepartment999.Checked)
                {
                    changeDepartment = 999;
                }
            }
        }

        private void newUserAddButton_Click(object sender, EventArgs e)
        {
            if (newAdminRadioButton.Checked)
            {
                connectedUser.Authority = 0;
            }
            if (newUserDepartment1.Checked)
            {
                connectedUser.Department = 0;
            }
            else if (newUserDepartment2.Checked)
            {
                connectedUser.Department = 1;
            }
            else if (newUserDepartment3.Checked)
            {
                connectedUser.Department = 2;
            }
            else if (newUserDepartment4.Checked)
            {
                connectedUser.Department = 3;
            }


            if (!duplicationCheck)
            {
                MessageBox.Show("ID 중복 체크를 확인해주세요.", "Error", MessageBoxButtons.OK);
            }
            else if (!availableId)
            {
                MessageBox.Show("ID 조건을 확인해주세요.", "Error", MessageBoxButtons.OK);
            }
            else if (!availablePw)
            {
                MessageBox.Show("패스워드 조건을 확인해주세요.", "Error", MessageBoxButtons.OK);
            }
            else if (!equalPass)
            {
                MessageBox.Show("패스워드 확인이 패스워드와 일치하지 않습니다.", "Error", MessageBoxButtons.OK);
            }
            else
            {
                if (dbModule.Insert(connectedUser))
                {
                    MessageBox.Show("회원 등록이 정상적으로 처리되었습니다.\n로그인 화면으로 돌아갑니다.", "Success", MessageBoxButtons.OK);
                    Application.Restart();
                }
                else
                {
                    MessageBox.Show("DataBase 접근 오류.\n회원 등록에 실패하였습니다.", "Failed", MessageBoxButtons.OK);
                }
            }
        }

        // 기존 계정 정보 수정
        private int changeAuthority, changeDepartment;  // tab page 4 열릴 때 tabContorl1_Selecting에서 connectedUser의 정보로 초기화 됨

        private void userInfoUpdateButton_Click(object sender, EventArgs e)
        {
            bool isUserPassChanged = false;
            bool isUserDepartChanged = false;

            string usingPw = this.currentPassText.Text.ToString();
            string changePw = this.changePassText.Text.ToString();
            if (usingPw.Equals(""))
            {
                MessageBox.Show("본인 확인용 패스워드가 비어있습니다. \n본인 확인용 패스워드를 입력해 주세요.", "Update Failed", MessageBoxButtons.OK);
                return;
            }            
            else if (!usingPw.Equals(connectedUser.Password))
            {
                // 본인 확인용 패스워드 인증 실패
                MessageBox.Show("본인 확인용 패스워드가 잘못 되었습니다. \n본인 확인용 패스워드를 다시 입력해 주세요.", "Update Failed", MessageBoxButtons.OK);
                return;
            }
            else
            {   
                // 본인 확인용 패스워드 인증 성공. 변경 사항 있을 경우 체크하여 update
                if (!changePw.Equals(""))
                {
                    if (Regex.IsMatch(changePw, PwPattern) && !usingPw.Equals(changePw))
                    {
                        connectedUser.Password = changePw;
                        isUserPassChanged = true;
                    }
                    else if (Regex.IsMatch(changePw, PwPattern) && usingPw.Equals(changePw))
                    {
                        MessageBox.Show("변경 할 패스워드가\n현재 사용 중인 패스워드와 같습니다.", "Update Failed", MessageBoxButtons.OK);
                        return;
                    }
                    else
                    {
                        MessageBox.Show("변경 할 패스워드는\n영문/숫자/특수문자 조합의 6~20자\n사이의 값을 넣어주세요.", "Update Failed", MessageBoxButtons.OK);
                        return;
                    }
                }

                if (connectedUser.Authority != changeAuthority)
                {
                    connectedUser.Authority = changeAuthority;
                    isUserDepartChanged = true;
                }

                if (connectedUser.Department != changeDepartment)
                {
                    connectedUser.Department = changeDepartment;
                    isUserDepartChanged = true;
                }

                if (dbModule.Update(connectedUser) && (isUserDepartChanged || isUserPassChanged))
                {
                    if(isUserDepartChanged && isUserPassChanged)
                    {
                        MessageBox.Show("변경 된 패스워드와 부서 적용을 위해 \n프로그램을 재시작 합니다.", "Update Success", MessageBoxButtons.OK);
                        Application.Restart();
                    }
                    else if (isUserPassChanged)
                    {
                        MessageBox.Show("변경 된 패스워드 적용을 위해 \n프로그램을 재시작 합니다.", "Update Success", MessageBoxButtons.OK);
                        Application.Restart();
                    }
                    else if (isUserDepartChanged)
                    {
                        MessageBox.Show("변경 된 부서 적용을 위해 \n프로그램을 재시작 합니다.", "Update Success", MessageBoxButtons.OK);
                        Application.Restart();
                    }
                }
                else if ((isUserDepartChanged || isUserPassChanged))
                {
                    MessageBox.Show("DataBase 접근 오류.\n사용자 정보 변경 실패.", "Error", MessageBoxButtons.OK);
                }
                else
                {
                    MessageBox.Show("변경 된 내용이 없습니다.", "Update Success", MessageBoxButtons.OK);
                }
            }
        }

        // 기존 계정 정보 삭제
        private void removeUserButton_Click(object sender, EventArgs e)
        {
            if (connectedUser.Password.Equals(this.removeUserPassText.Text.ToString()))
            {
                if (MessageBox.Show("정말 삭제하시겠습니까?", "Confirm delete", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    if (dbModule.Delete(connectedUser))
                    {
                        MessageBox.Show("사용자 정보 삭제에 성공했습니다.\n접속을 종료합니다.", "Delete", MessageBoxButtons.OK);
                        Application.Restart();
                    }
                    else
                    {
                        MessageBox.Show("DataBase 접근 오류.\n사용자 정보 삭제에 실패했습니다.", "Error", MessageBoxButtons.OK);
                    }

                }
            }
        }
        #endregion

        /////////////////////////////////////////////////////////////////////////////

        #region 문자열 포맷 관련 함수들
        private decimal inquiryTotalCost, inquiryProfit;
        private char[] delimiterChars = { ' ', '년', '월', '일', '-', ':', ' ' };

        private void DataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                if (sender == inquiryDataGridView1 && inquiryDataGridView1LoadComplete)
                {
                    // 조회 - 거래날짜 항목
                    if (e.ColumnIndex == inquiryTransactionDateCol.Index)
                    {
                        e.Value = string2DateFormat(e.Value.ToString(), DateFormat.NormalFormat);
                    }
                    // 조회 - 원가
                    if (e.ColumnIndex == inquiryCostCol.Index)
                    {
                        int transactionIndex = Int32.Parse(inquiryDataGridView1.Rows[e.RowIndex].Cells[inquiryInvisibleIndexCol.Index].Value.ToString());
                        inquiryTotalCost = dbModule.totalCostOfTransaction(transactionIndex);
                        e.Value = inquiryTotalCost;
                    }
                    // 조회 - 이익
                    if (e.ColumnIndex == inquiryProfitCol.Index)
                    {
                        decimal supplyPrice = (decimal)inquiryDataGridView1.Rows[e.RowIndex].Cells[inquirySupplyPriceCol.Index].Value;
                        inquiryProfit = supplyPrice - inquiryTotalCost;
                        e.Value = inquiryProfit;
                    }
                    // 조회 - 부가세
                    if (e.ColumnIndex == inquiryTaxCol.Index)
                    {
                        if (inquiryProfit < 0)
                        {
                            inquiryProfit = 0;
                        }
                        e.Value = Math.Round(Decimal.ToDouble(inquiryProfit) * 0.1);
                    }
                }
                else if (sender == outputDataGridView && outputDataGridViewLoadComplete)
                {
                    // 출력 - 날짜
                    if (e.ColumnIndex == outputTransactionDateCol.Index)
                    {
                        e.Value = string2DateFormat(date2StringFormat(e.Value.ToString(), DateFormat.FullDate), DateFormat.OutputFormat);
                    }
                }
                else
                {
                    // 조회 - 원가항목 - 인덱스
                    if (e.ColumnIndex == inquiryIndexCol.Index)
                    {
                        e.Value = e.RowIndex + 1;
                    }
                    // 입력 - 원가항목 - 인덱스
                    if (e.ColumnIndex == inputIndexCol.Index)
                    {
                        e.Value = e.RowIndex + 1;
                    }
                }

            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.StackTrace);
                Console.WriteLine(exc.Message);
            }
        }

        public String date2StringFormat(String dateSource, DateFormat mode)
        {
            string[] s = dateSource.Split(delimiterChars);

            if (mode == DateFormat.YearAndMonth)
            {
                return s[0] + s[1];
            }
            else
            {
                return s[0] + s[1] + s[2];
            }
        }

        public string string2DateFormat(String stringSource, DateFormat dateFormat)
        {
            if (stringSource != null)
            {
                string year, month, day;

                if (dateFormat == DateFormat.NormalFormat)
                {
                    year = stringSource.Substring(0, 4);
                    month = stringSource.Substring(4, 2);
                    day = stringSource.Substring(6, 2);

                    return year + "-" + month + "-" + day;
                }
                else
                {
                    month = stringSource.Substring(4, 2);
                    day = stringSource.Substring(6, 2);

                    return month + "월" + day + "일";
                }

            }
            else
            {
                return "0000-00-00";
            }
        }

        // 매입처 항목 편집 자동완성 함수
        private void dataGridView_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();

            string[] supplierList = dbModule.getAllSupplierList();
            autoComplete.AddRange(supplierList);

            int currentCol = 0;
            string headerText = "";

            if (sender == inputDataGridView)
            {
                currentCol = this.inputDataGridView.CurrentCell.ColumnIndex;
                headerText = this.inputDataGridView.Columns[currentCol].Name;
            }
            else if (sender == inquiryDataGridView2)
            {
                currentCol = this.inquiryDataGridView2.CurrentCell.ColumnIndex;
                headerText = this.inquiryDataGridView2.Columns[currentCol].Name;
            }

            if (headerText.Equals("inquirySupplierCol") || headerText.Equals("inputSupplierCol"))
            {
                TextBox tb = e.Control as TextBox;

                if (tb != null)
                {
                    tb.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    tb.AutoCompleteCustomSource = autoComplete;
                    tb.AutoCompleteSource = AutoCompleteSource.CustomSource;
                }
            }
            else
            {
                TextBox tb = e.Control as TextBox;

                if (tb != null)
                {
                    tb.AutoCompleteMode = AutoCompleteMode.None;
                }
            }

        }
        #endregion

    }
}
