using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SpreadExcelReportDemo
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Init();
            spreadsheetControl1.DocumentLoaded += (s, ea) =>{
                Init();
            };
        }
        void Init()
        {
            IWorkbook workbook = spreadsheetControl1.Document;
            workbook.MailMergeDataSource = CreateData();
        }
        List<Employee> CreateData()
        {
            return new List<Employee>() {
                new Employee(){Name="张三",Code="00001",Age=24,Address="广东深圳",EmailAddr="zhangsan@tao.com",PhoneNum="138000000001"},
                new Employee(){Name="李四",Code="00002",Age=24,Address="广东深圳",EmailAddr="lisi@tao.com",PhoneNum="138000000002"},
                new Employee(){Name="王五",Code="00003",Age=24,Address="广东深圳",EmailAddr="wangwu@tao.com",PhoneNum="138000000003"},
                new Employee(){Name="土豪",Code="00004",Age=24,Address="广东深圳",EmailAddr="tuhao@tao.com",PhoneNum="138000000004"},
                new Employee(){Name="敬业福",Code="00005",Age=24,Address="广东深圳",EmailAddr="jingyefu@tao.com",PhoneNum="138000000005"},
                new Employee(){Name="牛逼",Code="00006",Age=24,Address="广东深圳",EmailAddr="niubi@tao.com",PhoneNum="138000000006"},
                new Employee(){Name="超神",Code="00007",Age=24,Address="广东深圳",EmailAddr="chaoshen@tao.com",PhoneNum="138000000007"}
            };
        }
        class Employee
        {
            [DisplayName("姓名")]
            public string Name { get; set; }
            [DisplayName("员工编号")]
            public string Code { get; set; }
            [DisplayName("年龄")]
            public int Age { get; set; }
            [DisplayName("地址")]
            public string Address { get; set; }
            [DisplayName("电话")]
            public string PhoneNum { get; set; }
            [DisplayName("邮件地址")]
            public string EmailAddr { get; set; }
        }
    }
}
