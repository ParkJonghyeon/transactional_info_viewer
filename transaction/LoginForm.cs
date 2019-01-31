using System;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;

namespace Transaction
{
    public partial class LoginForm : Form
    {
        DataBaseModule dbModule;
        User loginTryingUser;
        bool loginFail = false;

        /// <summary>
        /// 기본적인 폼 구성과 이벤트
        /// </summary>
        public LoginForm(User user)
        {
            InitializeComponent();
            loginTryingUser = user;
            dbModule = new DataBaseModule(-1, -1);
            this.Activate();
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            Login();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.Equals(Keys.Enter) && !loginFail)
            {
                Login();
                textBox1.Focus();
            }
            else if (e.KeyCode.Equals(Keys.Escape))
            {
                this.Close();
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.Equals(Keys.Enter) && !loginFail)
            {
                Login();
                textBox1.Focus();
            }
            else if (e.KeyCode.Equals(Keys.Escape))
            {
                this.Close();
            }
        }

        private void Login()
        {
            if (LoginCheck())
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private bool LoginCheck()
        {
            loginTryingUser.Id = this.textBox1.Text.ToString();
            loginTryingUser.Password = this.textBox2.Text.ToString();

            string loginResult = dbModule.isExistUser(loginTryingUser);
            // authority 설정후 반환
            if (loginResult.Equals("logout") && dbModule.Login(loginTryingUser))
            {                
                return true;
            }
            else if (loginResult.Equals("login"))
            {
                loginFail = true;
                if (MessageBox.Show("다른 PC에서 로그인 중 입니다.\n접속 할 수 없습니다.", "Access Reject", MessageBoxButtons.OK) == DialogResult.OK)
                {
                    loginFail = false;
                }
                return false;
            }
            else if (loginResult.Equals("not user"))
            {
                loginFail = true;
                if (MessageBox.Show("ID 또는 패스워드가 잘못 되었습니다.", "Unregistered User", MessageBoxButtons.OK) == DialogResult.OK)
                {
                    loginFail = false;
                }
                return false;
            }
            else
            {
                loginFail = true;
                if (MessageBox.Show("로그인 과정 중 에러가 발생했습니다.", "Error", MessageBoxButtons.OK) == DialogResult.OK)
                {
                    loginFail = false;
                }
                return false;
            }
        }
                
        private void newUserLinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            loginTryingUser.Authority = -1;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
