using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.Skins;
using DevExpress.LookAndFeel;
using DevExpress.UserSkins;
using DevExpress.XtraBars.Helpers;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Ribbon;
using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
using System.IO;
using DevExpress.Spreadsheet.Export;
using DevExpress.XtraEditors.Filtering.Templates;
using System.Collections;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using Aspose.Cells;
using System.Reflection;

namespace ExcelTool
{
    public partial class Form1 : RibbonForm
    {
        public Form1()
        {
            InitializeComponent();
            InitSkinGallery();
            InitialBindingDataTable();
        }

        /// <summary>
        /// 文件管理框的初始化
        /// </summary>
        void InitialBindingDataTable()
        {
            Filemanager.Columns.Add("IsChoose", typeof(bool)).SetOrdinal(0);
            Filemanager.Columns.Add("FileName");
            this.gridControl1.DataSource = Filemanager;
        }


        public Hashtable Hashtable = new Hashtable();
        public List<string> FilePaths = new List<string>();
        public DataTable Filemanager = new DataTable();
        void InitSkinGallery()
        {
            SkinHelper.InitSkinGallery(rgbiSkins, true);
        }

        private void spreadsheetCommandBarButtonItem190_ItemClick(object sender, ItemClickEventArgs e)
        {
            //DirectoryInfo di = GetDirectoryInfo(3);
            //string file = di.FullName + "CTCS_ExcelModel.xlsx";
            //ImportExcel(file);
        }

        private List<TxtPoint> OpenCSV(string filePath)
        {
            List<TxtPoint> points = new List<TxtPoint>();
            Encoding encoding = Encoding.ASCII;
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            StreamReader sr = new StreamReader(fs, encoding);
            //记录每行读取的一行记录
            string strLine = "";
            string[] aryLine = null;
            string[] tableHead = null;
            //标示行数
            int columnCount = 0;
            bool isFirst = true;
            int lineCount = 0;
            while((strLine=sr.ReadLine())!=null)
            {
                if(isFirst)
                {
                    tableHead = strLine.Split(',');
                    isFirst = false;
                    columnCount = tableHead.Length;
                }
                else
                {
                    lineCount++;
                    aryLine = strLine.Split(',');
                    TxtPoint point = new TxtPoint();
                    point.Id = lineCount;
                    point.Longtitude = Convert.ToDouble(aryLine[0]) * 3600000;
                    point.Latitude = Convert.ToDouble(aryLine[1]) * 3600000;
                }
            }
            return points;
        }


        public static string OpenExcel()
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "所有文件(*.xls)|*.xls"; //设置要选择的文件的类型
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                return fileDialog.FileName;//返回文件的完整路径               
            }
            else
            {
                return null;
            }

        }


        private void ImportExcel(string filePath)//导入按钮
        {
            if (!string.IsNullOrEmpty(filePath))
            {
                IWorkbook workbook = spreadsheetControl.Document;
                workbook.LoadDocument(filePath);
            }
        }

        private void spreadsheetCommandBarButtonItem191_ItemClick(object sender, ItemClickEventArgs e)
        {
            SaveExcel();
        }


        private void SaveExcel()//保存按钮
        {
            try
            {
                SaveFileDialog s = new SaveFileDialog();
                s.Filter = "Excel文件(*.xlsx)|*.xlsx";
                s.Title = "保存文件";
                if(s.ShowDialog()==DialogResult.OK)
                {
                    spreadsheetControl.SaveDocument(s.FileName, DocumentFormat.Xlsx);
                }
            }
            catch (Exception ExError)
            {
                XtraMessageBox.Show("该文件正被别的地方占用");
                ExError.ToString();
            }
        }


        public static DirectoryInfo GetDirectoryInfo(int index)
        {
            string directoryindex = "";
            for (int i = 0; i < index; i++)
            {
                directoryindex += @"..\";
            }
            DirectoryInfo di = new DirectoryInfo(string.Format("{0}{1}", System.Windows.Forms.Application.StartupPath, directoryindex));
            return di;
        }

        /// <summary>
        /// 插入N行
        /// </summary>
        /// <param name="workbook">Excel工作区</param>
        /// <param name="sheetIndex">sheet索引</param>
        /// <param name="startRowIndex">开始插入行的索引(向下插入表的)</param>
        /// <param name="rowcount">插入的行数</param>
        /// <param name="formatRowIndex">格式行的索引值</param>
        private void InsertRows(DevExpress.Spreadsheet.Worksheet sheet,int startRowIndex,int rowcount,int formatRowIndex)
        {
            DevExpress.Spreadsheet.RowCollection rows = sheet.Rows;
            //example:
            //sheet.Rows.Insert(6, 5);//在行索引为6的行下方插入5行,即增加7、8、9、10、11
            sheet.Rows.Insert(startRowIndex, rowcount);
            for(int j=startRowIndex; j<startRowIndex+rowcount;j++)
            {
                sheet.Rows[j].CopyFrom(sheet.Rows[formatRowIndex]);//复制行格式
            }
        }

        private void btnImportData_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                OpenFileDialog o = new OpenFileDialog();
                o.Filter = "线路文件(*.txt)|*.txt|Excel文件(*.csv)|*.csv";
                o.Title = "打开文件";
                o.Multiselect = true;
                o.InitialDirectory = Application.StartupPath;
                o.RestoreDirectory = false;
                if (o.ShowDialog() == DialogResult.OK)
                {
                    string[] filepaths = o.FileNames;
                    foreach (string s in filepaths)
                    {
                        if (!FilePaths.Contains(s))//如果没有添加过此文件
                        {
                            //添加ComboBox
                            ((DevExpress.XtraEditors.Repository.RepositoryItemComboBox)DataSources.Edit).Items.Add(s.Substring(s.LastIndexOf("\\") + 1));
                            List<string[]> latlngs = new List<string[]>();
                            ReadFormatTxt(s, ref latlngs);//读格式的文件返回经纬度
                            string trackName = s.Substring(s.LastIndexOf("\\") + 1);
                            string filename = trackName;
                            trackName = trackName.Split(new char[] {','})[0];
                            List<TxtPoint> points = stringsToPoints(latlngs,trackName,filename);
                            //添加哈希表--文件名,和TxtPoints
                            Hashtable.Add(s.Trim(), points);//保存路径和TxtPoints
                            AddFileRow(s, true);//在文件管理栏加文件
                            //添加打开的文件
                            FilePaths.Add(s.Trim());//添加路径
                        }

                    }

                    //List<string[]> latlngs = new List<string[]>();
                    //ReadFormatTxt(filepath, ref latlngs);//读格式的文件返回经纬度
                    //List<TxtPoint> points = stringsToPoints(latlngs);
                    //Hashtable.Add(filepath, points);
                    if (spreadsheetControl.ActiveWorksheet != null)
                    {
                        IWorkbook workbook = spreadsheetControl.Document;
                        DevExpress.Spreadsheet.Worksheet sheet = spreadsheetControl.Document.Worksheets[0];
                        string filepath = filepaths.First();//第一个文件为默认值
                        List<TxtPoint> last = (List<TxtPoint>)Hashtable[filepath];
                        filepath = filepath.Substring(filepath.LastIndexOf("\\") + 1);
                        workbook.MailMergeDataSource = last;
                        DataSources.EditValue = filepath;

                        siStatus.Caption = string.Format("共有{0}行数据", last.Count);
                        siInfo.Caption = string.Format("当前绑定数据为{0}", filepath);
                        
                         //InsertRows(sheet, 6, points.Count - 4, 3);
                        //ImportData(sheet, 2, points);
                    }
                }
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        

        /// <summary>
        /// 添加文件
        /// </summary>
        /// <param name="filename">完整文件路径名称</param>
        /// <param name="isChoose">是否选择</param>
        private void AddFileRow(string filename,bool isChoose)
        {
            DataRow dr = Filemanager.NewRow();
            dr["IsChoose"] = isChoose;
            dr["FileName"] = filename.Substring(filename.LastIndexOf("\\")+1);
            Filemanager.Rows.Add(dr);
        }


        private void ReadFormatTxt(string filePath, ref List<string[]> list)
        {
            try
            {
                list.Clear();//首先清空list
                using (StreamReader sr = new StreamReader(filePath))
                {
                    string[] lines = File.ReadAllLines(filePath, Encoding.Default);
                    foreach (var line in lines)
                    {
                        string temp = line.Trim();
                        if (temp != "")
                        {
                            string[] arr = temp.Split(new char[] { '\t', ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
                            if (arr.Length > 0)
                            {
                                list.Add(arr);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        private void CheckSame()
        {
            try
            {
                List<string> selectFiles = FilePaths;
                string ss = string.Format("[{0}]", DateTime.Now.ToString("g")) + "\r\n";
                int count = 0;
                if (selectFiles != null && selectFiles.Count != 0)
                {
                    List<TxtPoint> txtPoints = new List<TxtPoint>();
                    foreach (string s in selectFiles)
                    {
                        if (Hashtable.ContainsKey(s))
                        {
                            txtPoints.AddRange((List<TxtPoint>)Hashtable[s]);
                        }
                    }
                    var group = txtPoints.GroupBy(a => a.KiloPos).Where(x => x.Count() > 1).ToList();
                    foreach (IGrouping<double, TxtPoint> e in group)
                    {
                        count++;
                        List<TxtPoint> newtxts = e.ToList();
                        foreach (TxtPoint tp in newtxts)
                        {
                            ss += string.Format("{0}{6}点编号{1}{6}经度{2}{6}纬度{3}{6}公里标{4}{6}航向角{5}{6}增量航向角{7}{6}", tp.OverlayName, tp.Id, tp.Longtitude, tp.Latitude, tp.KiloPos, tp.Bear,  "\t",tp.DeltaBear) + "\r\n";
                        }
                        ss += "\r\n";
                    }
                    if (!string.IsNullOrEmpty(ss))
                    {
                        DirectoryInfo d = GetDirectoryInfo(3);
                        string strFileName = d.FullName + "Log\\" + DateTime.Now.ToString("yyyy-MM-dd") + "SameKiloPos.txt";
                        FileStream fs;
                        StreamWriter sw;
                        if (File.Exists(strFileName))
                        {
                            fs = new FileStream(strFileName, FileMode.Append, FileAccess.Write);
                        }
                        else
                        {
                            fs = new FileStream(strFileName, FileMode.Create, FileAccess.ReadWrite);
                        }
                        sw = new StreamWriter(fs);
                        sw.WriteLine(ss);
                        sw.Close();
                        fs.Close();
                    }
                }
                string cc = string.Format("共有{0}组相同的公里标点", count);
                XtraMessageBox.Show(cc);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        private List<TxtPoint> stringsToPoints(List<string[]> ss,string trackName="UnKnown",string StationName="UnKnown",string Overlayname="UnKnown")
        {
            List<TxtPoint> points = new List<TxtPoint>();
            //
            foreach(string[] s in ss)
            {
                TxtPoint p = string2point(s,trackName,StationName,Overlayname);
                //只设置了经纬高,里程和备注
                p.Id = ss.IndexOf(s)+1;
                //赋值index
                points.Add(p);
            }
            //赋值Bear
            for(int i=0;i<ss.Count-1;i++)
            {
                double LatA = Convert.ToDouble(ss[i][1]);
                double LatB = Convert.ToDouble(ss[i + 1][1]);
                double LngA = Convert.ToDouble(ss[i][0]);
                double LngB = Convert.ToDouble(ss[i + 1][0]);
                points[i].Bear = GetBear(LatA, LatB, LngA, LngB);
            }
            //赋值DeltaBear
            for(int i=0;i<points.Count-1;i++)
            {
                points[i].DeltaBear = points[i + 1].Bear - points[i].Bear;
            }
            return points;
        }

        //通用
        private TxtPoint string2point(string[] s,string trackName,string StationName,string overlayname)
        {
            TxtPoint t = new TxtPoint();
            t.Longtitude = Convert.ToInt32(Convert.ToDouble(s[0]) * 3600000);
            t.Latitude = Convert.ToInt32(Convert.ToDouble(s[1]) * 3600000);
            t.Height = Convert.ToInt32(Convert.ToDouble(s[2]) * 100);
            t.KiloPos = Convert.ToDouble(s[3]);
            t.TrackName = trackName;
            t.StationName = StationName;
            t.OverlayName = overlayname;
            if (s.Length >= 5)
            {
                t.Tag = s[4].ToString();
            }
            return t;
        }

        private void btnTransferModel_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (spreadsheetControl.Document.Worksheets[2] != null)
            { TransferModel(spreadsheetControl.Document.Worksheets[2]); }
            if (spreadsheetControl.Document.Worksheets[3] != null)
            { TransferModel(spreadsheetControl.Document.Worksheets[3]); }

        }

        private void TransferModel(DevExpress.Spreadsheet.Worksheet sheet)
        {
            //取消合并
            sheet.UnMergeCells(sheet.Range["D2:J2"]);
            sheet.UnMergeCells(sheet.Range["K2:N2"]);
            //合并
            sheet.MergeCells(sheet.Range["D2:D3"]);
            sheet.MergeCells(sheet.Range["E2:E3"]);
            sheet.MergeCells(sheet.Range["F2:F3"]);
            sheet.MergeCells(sheet.Range["G2:G3"]);
            sheet.MergeCells(sheet.Range["H2:H3"]);
            sheet.MergeCells(sheet.Range["I2:I3"]);
            sheet.MergeCells(sheet.Range["J2:J3"]);
            sheet.MergeCells(sheet.Range["K2:K3"]);
            sheet.MergeCells(sheet.Range["L2:L3"]);
            sheet.MergeCells(sheet.Range["M2:M3"]);
            sheet.MergeCells(sheet.Range["N2:N3"]);
            //删除
            sheet.Rows[2].Delete();

            //更改样式
            sheet.Cells["D2"].Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Medium);
            sheet.Cells["E2"].Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Medium);
            sheet.Cells["F2"].Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Medium);
            sheet.Cells["H2"].Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Medium);
            sheet.Cells["I2"].Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Medium);
            sheet.Cells["J2"].Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Medium);
            sheet.Cells["K2"].Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Medium);
            sheet.Cells["L2"].Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Medium);
            sheet.Cells["M2"].Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Medium);
            sheet.Cells["N2"].Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Medium);
        }

        private void spreadsheetControl_ActiveSheetChanged(object sender, ActiveSheetChangedEventArgs e)
        {
            //XtraMessageBox.Show(e.NewActiveSheetName);
        }

        //Import Data Form txtPoints
        //轨道地理信息表
        private void ImportData(DevExpress.Spreadsheet.Worksheet sheet,int tableHeader,List<TxtPoint> txtPoints)
        {
            int j = 0;
            for(int i=tableHeader;i<tableHeader+txtPoints.Count;i++)
            {
                TxtPoint t = txtPoints[j];
                sheet.Cells[i, 0].SetValue(j+1);//数据编号
                sheet.Cells[i, 1].SetValue(t.StationName);//车站名称
                //sheet.Cells[i, 2].SetValue();//车站编号
                sheet.Cells[i, 3].SetValue(t.TrackName);//轨道名称
                //sheet.Cells[i,4].SetValue();//轨道编号
                sheet.Cells[i, 5].SetValue(t.Longtitude);//经度(毫秒)
                sheet.Cells[i, 6].SetValue(t.Latitude);//纬度(单位毫秒)
                sheet.Cells[i, 7].SetValue(t.Height);//高程(厘米)
                sheet.Cells[i, 8].SetValue(t.KiloPos);//里程
                sheet.Cells[i, 9].SetValue(t.Bear);//航向角
                sheet.Cells[i, 10].SetValue(t.DeltaBear);//增量航向角
                j++;
            }
        }

        private double GetBear(double LatA,double LatB,double LngA,double LngB)
        {
            double bear = 0.0;
            //
            double E = LngB - LngA;
            double N = LatB - LatA;

            bear = Math.Atan2(E, N);//得到弧度-3.14~3.14
            //if(bear<0)
            //{
            //    bear += 6.28;//弧度范围0~6.28
            //}

            //改成角度
            //bear=headingA*180/Math.PI;
            //if(bear<0)
            //{
            //    bear += 360;//
            //}
            ////方位角转换转弧度
            //bear *= 0.0174533;
            //if(bear>3.14)
            //{
            //    bear -= 6.28;
            //}
            return bear;
        }

        private void spreadsheetCommandBarButtonItem203_ItemClick(object sender, ItemClickEventArgs e)
        {

        }

        private void repositoryItemComboBox1_EditValueChanged(object sender, EventArgs e)
        {
        }

        private void repositoryItemComboBox1_Click(object sender, EventArgs e)
        {
        }

        private void repositoryItemComboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            IWorkbook workbook = spreadsheetControl.Document;

            string s = FilePaths.Find(a => a.Contains(DataSources.EditValue.ToString().Trim()));
            List<TxtPoint> points = (List<TxtPoint>)Hashtable[s];
            workbook.MailMergeDataSource = points;

            siStatus.Caption = string.Format("共有{0}行数据", points.Count);
            siInfo.Caption = string.Format("当前绑定数据为{0}", DataSources.EditValue);

        }

        //只能是检查文件中的公里标，还没办法检查表中改后的公里标
        private void btnCheckKilo_ItemClick(object sender, ItemClickEventArgs e)
        {
            CheckSame();
        }

        private void barButtonItem1_ItemClick(object sender, ItemClickEventArgs e)
        {
            IWorkbook workbook = spreadsheetControl.Document;
            IList<IWorkbook> workbooks=workbook.GenerateMailMergeDocuments();
            
        }

        private void btnMergeSheet_ItemClick(object sender, ItemClickEventArgs e)
        {
            OpenFileDialog o = new OpenFileDialog();
            o.Multiselect = true;
            o.Filter = "表格文件(*.xlsx)|*.xlsx";
            o.Title = "打开需要合并的Excel文件";
            List<Workbook> books = new List<Workbook>();
            if (o.ShowDialog()==DialogResult.OK)
            {
                foreach(string s in o.FileNames)
                {
                    Workbook sourcebook = new Workbook(s);
                    books.Add(sourcebook);
                }
                SaveFileDialog sFD = new SaveFileDialog();
                sFD.Filter = "表格文件(*.xlsx)|*.xlsx";
                sFD.InitialDirectory = GetDirectoryInfo(3).FullName;
                if (sFD.ShowDialog() == DialogResult.OK)
                {
                    Workbook targetbook = books.First();
                    books.Remove(targetbook);
                    foreach (Workbook w in books)
                    {
                        targetbook.Combine(w);
                        targetbook.Save(sFD.FileName);
                    }
                }
            }
        }

        private void btnExtractInfo_ItemClick(object sender, ItemClickEventArgs e)
        {
            List<TxtPoint> points = new List<TxtPoint>();
            List<string> selectFiles = GetSelectFiles();//文件选择框中的文件
            foreach(string s in selectFiles)
            {
                string file = FilePaths.Find(a => a.Contains(s));//完整路径
                points.AddRange((List<TxtPoint>)Hashtable[file]);
            }
            var deviceInfo = points.FindAll(t => t.Tag != "UnKnown".Trim());//得到设备的信息

            SaveFileDialog sFD = new SaveFileDialog();
            sFD.Title = "合并至";
            sFD.Filter = "文件(*.txt)|*.txt|CSV文件(*.csv)|*.csv";
            sFD.InitialDirectory = GetDirectoryInfo(3).FullName;
            if (sFD.ShowDialog() == DialogResult.OK)
            {
                //if(sFD.FileName.Contains("txt"))
                if(Path.GetExtension(sFD.FileName)==".txt")
                {
                    //保存为txt文件
                    string s = SaveTxt(deviceInfo);
                    using (StreamWriter sw = new StreamWriter(sFD.FileName))
                    {
                        sw.Write(s);
                        string ss = string.Format("文件保存至{0}", sFD.FileName);
                        XtraMessageBox.Show(ss);
                    }
                }
                else if(sFD.FileName.Contains("csv"))
                {
                    //保存为csv文件
                    //加上标题
                    if(SaveDataToCSVFile(deviceInfo,sFD.FileName))
                    {
                        string ss= string.Format("文件保存至{0}", sFD.FileName);
                        XtraMessageBox.Show(ss);
                    }

                }
            }
        }
        /// <summary>
        /// 获取类的属性集合（以便生成CSV文件的所有Column标题）
        /// </summary>
        /// <returns></returns>
        private PropertyInfo[] GetPropertyInfoArray()
        {
            PropertyInfo[] props = null;
            try
            {
                Type type = typeof(TxtPoint);
                object obj = Activator.CreateInstance(type);
                props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
            return props;
        }
        /// <summary>
        /// Save the List data to CSV file
        /// </summary>
        /// <param name="studentList">data source</param>
        /// <param name="filePath">file path</param>
        /// <returns>success flag</returns>
        private bool SaveDataToCSVFile(List<TxtPoint> txtpoints, string filePath)
        {
            bool successFlag = true;

            StringBuilder strColumn = new StringBuilder();
            StringBuilder strValue = new StringBuilder();
            StreamWriter sw = null;
            PropertyInfo[] props = GetPropertyInfoArray();

            try
            {
                sw = new StreamWriter(filePath);
                for (int i = 0; i < props.Length; i++)
                {
                    strColumn.Append(props[i].Name);
                    strColumn.Append(",");
                }
                strColumn.Remove(strColumn.Length - 1, 1);
                sw.WriteLine(strColumn);    //write the column name

                //这是按照属性顺序写的
                for (int i = 0; i < txtpoints.Count; i++)
                {
                    strValue.Remove(0, strValue.Length); //clear the temp row value
                    strValue.Append(txtpoints[i].Id);
                    strValue.Append(",");
                    strValue.Append(txtpoints[i].OverlayName);
                    strValue.Append(",");
                    strValue.Append(txtpoints[i].StationName);
                    strValue.Append(",");
                    strValue.Append(txtpoints[i].TypeName);
                    strValue.Append(",");
                    strValue.Append(txtpoints[i].KiloPos);
                    strValue.Append(",");
                    strValue.Append(txtpoints[i].TrackName);
                    strValue.Append(",");
                    strValue.Append(txtpoints[i].Longtitude / 3600000);
                    strValue.Append(",");
                    strValue.Append(txtpoints[i].Latitude/3600000);
                    strValue.Append(",");
                    strValue.Append(txtpoints[i].Height / 100);
                    strValue.Append(",");
                    strValue.Append(txtpoints[i].Bear);
                    strValue.Append(",");
                    strValue.Append(txtpoints[i].DeltaBear);
                    sw.WriteLine(strValue); //write the row value
                }
            }
            catch (Exception ex)
            {
                successFlag = false;
            }
            finally
            {
                if (sw != null)
                {
                    sw.Dispose();
                }
            }
            return successFlag;
        }

        public string SaveTxt(List<TxtPoint> points)
        {
            string s = "";
            string tab = "\t";
            foreach(TxtPoint p in points)
            {
                s += p.Longtitude / 3600000 + tab;//返回到毫秒
                s += p.Latitude / 3600000+tab;//
                s += p.Height / 100 + tab;
                s += p.KiloPos + tab;
                s += p.Tag + tab;//加上信号机的类型
                s += p.Bear + tab;//加上航向角
                s += p.DeltaBear + tab;//加上增量航向角
                s += "\r\n";//换行
            }
            return s;
        }


        private void DataSources_ItemClick(object sender, ItemClickEventArgs e)
        {

        }

        private void repositoryItemCheckEdit1_QueryCheckStateByValue(object sender, DevExpress.XtraEditors.Controls.QueryCheckStateByValueEventArgs e)
        {
            SetCheckStateValue(e);
        }

        private void SetCheckStateValue(DevExpress.XtraEditors.Controls.QueryCheckStateByValueEventArgs e)
        {
            string val = "";
            if (e.Value != null)
            {
                val = e.Value.ToString();
            }
            else
            {
                val = "True";//默认为选中 
            }
            switch (val)
            {
                case "True":
                    e.CheckState = CheckState.Checked;
                    break;
                case "False":
                    e.CheckState = CheckState.Unchecked;
                    break;
                case "Yes":
                    goto case "True";
                case "No":
                    goto case "False";
                case "1":
                    goto case "True";
                case "0":
                    goto case "False";
                default:
                    e.CheckState = CheckState.Checked;
                    break;
            }
            e.Handled = true;
        }

        private void barButtonItem2_ItemClick(object sender, ItemClickEventArgs e)
        {
            DoDelete();

        }
        /// <summary>
        /// 关闭且删除当前文件
        /// </summary>
        public void DoDelete()
        {
            List<string> selectedFiles = GetSelectFiles();
            if (selectedFiles != null && selectedFiles.Count != 0)
            {
                foreach (string s in selectedFiles)
                {
                    var obj = FilePaths.Find(a => a.Contains(s));
                    FilePaths.Remove(obj);//从保存文件路径的全局变量中删除文件
                    Hashtable.Remove(obj);//从哈希表删除关闭的文件
                    ((DevExpress.XtraEditors.Repository.RepositoryItemComboBox)DataSources.Edit).Items.Remove(s);
                }
            }
            this.gridView1.DeleteSelectedRows();
            gridView1.RefreshData();
            this.gridView1.OptionsBehavior.Editable = true;
        }

        /// <summary>
        /// 获取文件选择列的文件得到的是文件后缀
        /// </summary>
        /// <returns></returns>
        private List<string> GetSelectFiles()
        {
            List<string> selectedFiles = new List<string>();
            for (int i = 0; i < gridView1.DataRowCount; i++)
            {
                var value = gridView1.GetDataRow(i)["IsChoose"].ToString().Trim();
                if (value == "True")//选择
                {
                    gridView1.SelectRow(i);
                    selectedFiles.Add(gridView1.GetDataRow(i)["FileName"].ToString().Trim());
                }
                else if (value == "False")
                {
                    gridView1.UnselectRow(i);
                    continue;
                }
            }
            return selectedFiles;
        }

        private void repositoryItemCheckEdit1_CheckedChanged(object sender, EventArgs e)
        {
            if (!gridView1.IsNewItemRow(gridView1.FocusedRowHandle))
            {
                gridView1.CloseEditor();
                gridView1.UpdateCurrentRow();
            }
            CheckState check = (sender as DevExpress.XtraEditors.CheckEdit).CheckState;
            int index = this.gridView1.FocusedRowHandle;
            if(check==CheckState.Checked)
            {
            }
            else
            {

            }
        }

        private void gridView1_Click(object sender, EventArgs e)
        {
            try
            {
                Point pt = gridControl1.PointToClient(Control.MousePosition);
                DevExpress.XtraGrid.Views.Grid.ViewInfo.GridHitInfo info = gridView1.CalcHitInfo(pt);
                if (info.InColumn && info.Column != null)
                {
                    string s = info.Column.FieldName.ToString();
                    switch (s)
                    {
                        case "IsChoose":
                            SetAllCheck("IsChoose", GetIsAllCheck("IsChoose"));
                            break;
                        default:
                            break;
                    }
                }
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
         }
        /// <summary>
        /// 设置GridView8某列全选
        /// </summary>
        /// <param name="columnFieldName">列的FieldName</param>
        /// <param name="checkState">选中状态</param>
        private void SetAllCheck(string columFieldName, bool checkState)
        {
            for (int i = 0; i < gridView1.DataRowCount; i++)
            {
                gridView1.SetRowCellValue(i, gridView1.Columns[columFieldName], checkState);
            }
            gridControl1.Refresh();
            gridView1.RefreshData();
        }

        /// <summary>
        /// GridView8判断某列是否全选
        /// </summary>
        private bool GetIsAllCheck(string columnFieldName)
        {
            DataTable dt = (DataTable)this.gridControl1.DataSource;
            //如果不是全选，则返回全选，是全选返回未选
            List<bool> states = new List<bool>();
            for (int j = 0; j < dt.Rows.Count; j++)
            {
                var state = dt.Rows[j][columnFieldName];
                states.Add(Convert.ToBoolean(state));
            }
            if (states.TrueForAll(a => a))//判断是否全为false(全未选)-true;不是全选(即全为false)的情况下就返回false
            {
                //全为true返回false
                return false;
            }
            else
            {
                //否则返回true
                return true;
            }
        }


        // //访问行
        // Workbook workbook=new Workbook();
        // //Access a Collection of rows
        // RowCollection rows = workbook.Worksheets[0].Rows;
        // //Access the first row by its index in the collection of rows;
        // Row firstRow_byIndex = rows[0];
        // //Access ths first row by its unique name
        // Row firstRow_byName = rows["1"];


        //// 访问列
        // Workbook workbook=new Workbook();
        // //Access a collection of columns
        // ColumnCollection columns = workbook.Worksheets[0].Columns.
        // //ColumnCollection columns=workbook.Worksheets[0].Columns;
        // Column firstColumn_byIndex = columns[0];
        // //Access the first column by its unique name;
        // Column firstColumn_byName = columns["A"];

    }
}