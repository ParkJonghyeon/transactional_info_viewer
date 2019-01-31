using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Transaction
{
    public class User
    {
        private string id;
        private string pass;
        private int authority;
        // authority -1 : 비회원/ 0 : 사용자/ 1 : 조회사용자
        private int department;
        // department 0 : 관리영업/ 1 : 기획/연구개발/ 2 : SI사업/ 3 : 광주지역/ 999 : 관리자

        public User()
        {
            id = null;
            pass = null;
            authority = -99;
            department = 0;
        }

        public User(string id, string pass, int authority, int department)
        {
            this.id = id;
            this.pass = pass;
            this.authority = authority;
            this.department = department;
        }

        public string AuthorityString()
        {
            switch (this.authority)
            {
                case -1:
                    return "비회원";
                case 0:
                    return "사용자(조회/입력)";
                case 1:
                    return "조회사용자";
                default:
                    return "UNKNOWN";
            }
        }

        public string DepartmentString()
        {
            switch (this.department)
            {
                case 0:
                    return "관리영업";
                case 1:
                    return "기획/연구개발";
                case 2:
                    return "SI사업";
                case 3:
                    return "광주지역";
                case 999:
                    return "관리자";
                default:
                    return "UNKNOWN";
            }
        }

        public string Id { get { return id; } set { id = value; } }
        public string Password { get { return pass; } set { pass = value; } }
        public int Authority { get { return authority; } set { authority = value; } }
        public int Department { get { return department; } set { department = value; } }
    }
}
