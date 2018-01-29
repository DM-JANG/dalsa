using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace DBtxt
{
    public partial class Form1 : Form
    {
       // BackgroundWorker thdProecess = null; //로딩을 위한 함수
        static string ServerName;
        static string DBName;
        static string ID = "";
        static string PW = "";
        public String selectqry;
        public bool All = true;
        static string connString;
        public bool stop = true;
    
        public bool backupko = false; // 백업 유무에 대해서 알려준다.

        public static state mState = state.Disconnect;

        public enum state
        {
            Normal = 0, Disconnect = 1
        }

        public Form1()
        {
            InitializeComponent();

        }

        private void DBConnection_Click(object sender, EventArgs e) //DB 연결 메서드
        {

            ServerName = SNameTxt.Text.Trim(); // 텍스트의 값을 받아온다.
            DBName = DNameTxt.Text.Trim();
            ID = UIDTxt.Text.Trim();
            PW = UPWTxt.Text.Trim();

            if (ServerName == "" || DBName == "")
            {
                label5.Text = "Enter Value ServerName or DBName";
            }
            else
            {
                if (ID == "" && PW == "")
                {
                    connString = String.Format("Data Source={0};Initial Catalog={1};Integrated Security=True",
                                               ServerName, DBName);
                }
                else
                {
                    connString = String.Format("Data Source={0};Initial Catalog={1};User Id={2};Password={3}",
                                               ServerName, DBName, ID, PW);
                }

                try
                {
                    using (SqlConnection connection = new SqlConnection(connString))
                    {
                        connection.Open();
                        label5.Text = "Connection Suceeced!!";
                        mState = state.Normal;
                    }

                }
                catch (Exception ex)
                {
                    mState = state.Disconnect;
                    label5.Text = "Connection false..";
                }
            }

        }

        private void Close_Click(object sender, EventArgs e) // 폼 종료
        {
            Form1.ActiveForm.Close();
        }

        private void Open_Click_1(object sender, EventArgs e) // 경로 설정하는 메서드
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK) // OK버튼 눌럿을 시에만 동작
            {
                FOpenread.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (mState == state.Normal)
            {
                selectqry = "SELECT name FROM sysobjects WHERE type = 'U' AND category = 0 "; //해당 DB의 모든 테이블 리스트를 가져온다.
                System.Data.DataTable table = new System.Data.DataTable();

                using (SqlConnection connection = new SqlConnection(connString))
                {
                    connection.Open();
                    SqlDataAdapter adapter = new SqlDataAdapter(selectqry, connection);
                    adapter.Fill(table);

                }
                TBDG.DataSource = table;
                TBDG.AllowUserToAddRows = false;
                TBDG.Columns[0].ReadOnly = true;
                TBDG.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;


            }
            FOpenread.Clear(); //  탭이 바뀔시 경로도 초기화
            xlspath.Clear();
            All = true; // 탭이 바뀔시 전부선택도 초기화
            DBbck.Checked = false;
        }

        private void AllCheck_Click(object sender, EventArgs e)
        {
            if (All)
            {
                TBDG.SelectAll();
                All = false;
            }
            else
            {
                TBDG.ClearSelection();
                All = true;
            }
        }
      
        private void Exprot_Click(object sender, EventArgs e)
        {
    
            try
            {
                if (FOpenread.Text != "")
                {
                    if (mState == state.Normal)
                    {

                        bool Caption = true;
                        int count = TBDG.SelectedRows.Count;
                        int count2 = 0;
                        String make = "";
                        using (System.Data.DataTable table = new System.Data.DataTable())
                        {
                            using (SqlConnection connection = new SqlConnection(connString))
                            {
                                connection.Open();
                                String CVSname = "";

                                for (int i = 0; i < count; i++)
                                {
                                    CVSname = "";
                                    Caption = true;
                                    table.Clear();
                                    table.Columns.Clear(); // 이상하게 컬럼만 클리어가 안되서 추가.
                                    table.Dispose();


                                    selectqry = String.Format(" * FROM {0} ", TBDG.SelectedRows[count - i - 1].Cells[0].Value.ToString());
                                    switch (TBDG.SelectedRows[count - i - 1].Cells[0].Value.ToString())
                                    {
                                        case "AT1001":
                                            selectqry = "SELECT TOP(100)" + selectqry + " order by LOG_IN DESC";
                                            break;
                                        case "CT1001":
                                            selectqry = "SELECT TOP(100)" + selectqry + " order by DATTME DESC";
                                            break;
                                        case "CT1010":
                                            selectqry = "SELECT TOP(500)" + selectqry + " order by DATTME DESC";
                                            break;
                                        case "CT1011":
                                            selectqry = "SELECT TOP(500)" + selectqry + " order by DATTME DESC";
                                            break;
                                        case "CT1020":
                                            selectqry = "SELECT TOP(500)" + selectqry + " order by DATTME DESC";
                                            break;
                                        case "CT1030":
                                            selectqry = "SELECT TOP(500)" + selectqry + " order by DATTME DESC";
                                            break;
                                        case "CT1040":
                                            selectqry = "SELECT TOP(500)" + selectqry + " order by DATTME DESC";
                                            break;
                                        case "CT1050":
                                            selectqry = "SELECT TOP(500)" + selectqry + " order by DATTME DESC";
                                            break;
                                        case "CT2010":
                                            selectqry = "SELECT TOP(100)" + selectqry + " order by DATTME DESC";
                                            break;
                                        case "CT1021":
                                            selectqry = "SELECT TOP(500)" + selectqry + " order by DATTME DESC";
                                            break;
                                        default:
                                            selectqry = "SELECT" + selectqry;
                                            break;
                                    }
                                    string _Filestr = FOpenread.Text + "\\" + TBDG.SelectedRows[count - i - 1].Cells[0].Value.ToString() + ".txt"; // 파일 존재유무 확인을 위한 경로설정
                                    System.IO.FileInfo fi = new System.IO.FileInfo(_Filestr);

                                    using (SqlDataAdapter adapter = new SqlDataAdapter(selectqry, connection))
                                    {
                                        if (fi.Exists)//파일 존재할시
                                        {
                                            MessageBox.Show(TBDG.SelectedRows[count - i - 1].Cells[0].Value.ToString() + ".txt 파일이 이미 존재합니다.");
                                        }
                                        else// 파일 존재안할시
                                        {
                                            using (StreamWriter writer = new StreamWriter(FOpenread.Text + "\\" + TBDG.SelectedRows[count - i - 1].Cells[0].Value.ToString() + ".txt", true, Encoding.UTF8))
                                            {
                                                make += TBDG.SelectedRows[count - i - 1].Cells[0].Value.ToString() + ".txt 생성완료 \n";

                                                if (Caption) //제목 출력과 값출력시 두번 들어가지 않도록 하기 위해
                                                {
                                                    adapter.Fill(table);
                                                }
                                                if (table.Rows.Count == 0) //데이터가 없을경우
                                                {
                                                    for (int j = 0; j < table.Columns.Count; j++)
                                                    {
                                                        if (Caption) // 제목 출력
                                                        {
                                                            if (j == table.Columns.Count - 1)
                                                            {
                                                                CVSname += table.Columns[j].Caption.ToString() + "\n";
                                                                Caption = false;

                                                            }
                                                            else
                                                            {

                                                                CVSname += table.Columns[j].Caption.ToString() + "|";

                                                            }
                                                        }
                                                    }
                                                }
                                                for (int z = 0; z < table.Rows.Count; z++)
                                                {

                                                    for (int j = 0; j < table.Columns.Count; j++)
                                                    {
                                                        if (Caption)
                                                        {
                                                            if (j == table.Columns.Count - 1)
                                                            {
                                                                CVSname += table.Columns[j].Caption.ToString() + "\n";
                                                                Caption = false;

                                                                z--;
                                                                //writer.WriteLine(CVSname);
                                                                //CVSname = "";
                                                            }
                                                            else
                                                            {

                                                                CVSname += table.Columns[j].Caption.ToString() + "|";

                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (j == table.Columns.Count - 1)
                                                            {
                                                                if (table.Rows[z].ItemArray[j].ToString() == "System.Byte[]") //조건문??
                                                                {
                                                                    string ox = "0x";
                                                                    byte[] bt = ObjectByteArrayConverter(table.Rows[z].ItemArray[j]);
                                                                    for (int a = 27; a < bt.Length; a++)
                                                                        ox += bt[a].ToString("X");
                                                                    CVSname += ox + "\n";

                                                                }
                                                                else
                                                                {
                                                                    CVSname += table.Rows[z].ItemArray[j].ToString() + "\n";


                                                                    //writer.WriteLine(CVSname);
                                                                    //CVSname = "";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if (table.Rows[z].ItemArray[j].ToString() == "System.Byte[]")
                                                                {
                                                                    string ox = "0x";
                                                                    byte[] bt = ObjectByteArrayConverter(table.Rows[z].ItemArray[j]);
                                                                    for (int a = 27; a < bt.Length; a++)
                                                                        ox += bt[a].ToString("X");
                                                                    CVSname += ox + "|";


                                                                }
                                                                else
                                                                {
                                                                    CVSname += table.Rows[z].ItemArray[j].ToString() + "|";
                                                                    ;
                                                                }

                                                            }
                                                        }
                                                    }


                                                }

                                                writer.WriteLine(CVSname);

                                            }
                                        }
                                    }

                                }
                                if (make != "")// 파일 생성된거 확인
                                {
                                    MessageBox.Show(make);
                                }
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("경로를 설정해 주세요.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static byte[] ObjectByteArrayConverter(object p_obj) // byte[] -> object

        {

            MemoryStream ms = new MemoryStream();

            BinaryFormatter formatter = new BinaryFormatter();
            try

            {

                formatter.Serialize(ms, p_obj);

                return ms.ToArray();

            }

            catch

            {

                return null;

            }

            finally

            {

                ms.Close();

            }

        }




        public static object ByteArrayToObject(byte[] p_buffer)//Object -> byte[]
        {

            BinaryFormatter formatter = new BinaryFormatter();

            MemoryStream ms = new MemoryStream(p_buffer);


            try

            {

                return formatter.Deserialize(ms);

            }

            catch (System.Exception ex)

            {

                MessageBox.Show(ex.Message);

                return null;

            }

        }


        private void Form1_Load(object sender, EventArgs e)
        {
            //thdProecess = new BackgroundWorker();
            //thdProecess.DoWork += new DoWorkEventHandler(Thread_Process);
            SNameTxt.Text = SystemInformation.ComputerName;
            
        }

        //private void Thread_Process(object sender, DoWorkEventArgs e)

        //{
        //    Process form = new Process(); // 작업 진행창 표시
        //    form.Show();

        //    for (int i = 0; i < 33;)
        //    {
        //        num = i * 3;
        //        form.SetProcess(num); // 작업현황 표시
        //        Thread.Sleep(100); // 작업에 필요한 지연
        //    }
        //    form.Close(); // 작업 진행창 닫기
        //}



        private void CSVOpen_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK) // OK버튼 눌럿을 시에만 동작
            {
                CSVPath.Text = openFileDialog1.FileName;
                
            }
        }
        public void Tranforms(String[] FileNames,BackgroundWorker woker, DoWorkEventArgs e)
        {
            if (FileNames[0] != "")
            {
                Excel.Application App;
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            int z = 1;
            App = new Excel.Application();
            workbook = App.Workbooks.Add();
            worksheet = workbook.ActiveSheet;
                App.Visible = false;
            

          
                String path = FileNames[0];
                String pathOnly = Path.GetDirectoryName(path);
                String fileName;
                long txtrow = 0;
                int row = 0;
                double num = 0;
                for (int j = 0; j < FileNames.Length; j++)
                {
                    path = FileNames[j];
    
                    using (StreamReader sr = new StreamReader(path, Encoding.UTF8))
                    {
                        while (!sr.EndOfStream)
                        {
                            txtrow++;
                            sr.ReadLine();
                        }
                    }
                }
                    for (int j = 0; j < FileNames.Length; j++)
                {
                    path = FileNames[j];
                    fileName = Path.GetFileName(path);
                    worksheet = workbook.ActiveSheet;
                    int RowLast =0;
                   // long txtrow = 0;
             
                    worksheet.Name = fileName.Remove(fileName.Length - 4, 4);
                    //using (StreamReader sr = new StreamReader(path, Encoding.UTF8))
                    //{
                    //    while (!sr.EndOfStream)
                    //    {
                    //        txtrow++;
                    //        sr.ReadLine();
                    //    }
                    // }
                     using (StreamReader sr = new StreamReader(path, Encoding.UTF8))
                        {
                            while (!sr.EndOfStream)
                        {
                      
                            string s = sr.ReadLine();
                            string[] temp = s.Split('|');

                            for (int i = 0; i < temp.Length; i++)
                            {
                                worksheet.Cells[z, i + 1] =  temp[i];
                            }
                            row++;
                            double sum = (100 / ((double)(txtrow * FileNames.Length))) * ((double)(j + 1) * (double)(row));
                            num = (num > sum) ? num : sum; // 추후 로딩을 위한 값  
                            if (num != 0)
                                woker.ReportProgress((int)(num));

                            if (z==1)
                            RowLast = temp.Length;
                            z++;
                        }
                        Excel.Range heder = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[z, RowLast]];
                        heder.Columns.AutoFit();
                        Excel.Range filterRng = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[z, RowLast]];
                       if(RowLast != 1)
                        filterRng.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
                

                    }

                    z = 1;
                    if (j < FileNames.Length - 1)
                    {
                        workbook.Worksheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);

                    }
                }

                System.IO.FileInfo fi = new System.IO.FileInfo(path + "\\DB_EXProt.xls");
                if (fi.Exists)
                {
                    MessageBox.Show("이미 DB_ExPort.xls 파일이 존재합니다.");
                }
                else
                {
                    workbook.SaveAs(pathOnly + "\\DB_ExPort.xls", Excel.XlFileFormat.xlWorkbookNormal);
                }
                MessageBox.Show("DB_EXPort.xls 생성 완료!!", "생성 완료",
                                  MessageBoxButtons.OK,
                                 MessageBoxIcon.None,
                                 MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                workbook.Save();
                workbook.Close(0);

                App.Quit();
                ExcelKill();

            }
            else
            {
                MessageBox.Show("경로를 설정해 주세요.");
            }
        }
        private void Tranform_Click(object sender, EventArgs e)
        {
            Process pr = new Process();
            pr.Show();
            pr.TransRun(openFileDialog1.FileNames,pr);
          //  Tranforms();
            


        }

        private void xlsOpen_Click(object sender, EventArgs e)
        {

            if (openFileDialog2.ShowDialog() == DialogResult.OK) // OK버튼 눌럿을 시에만 동작
            {
                xlspath.Text = openFileDialog2.FileName;

            }
        }
        public void ExcelKill()
        {
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process p in process)
            {
                if(!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
        }
        private void xlsImport_Click(object sender, EventArgs e)
        {
            //try
            //{
                if (mState == state.Normal)
                {

                    if (xlspath.Text != "")
                    {
                        if (backupko)
                        {
                            String Qry = "";
                            System.Data.DataTable table = new System.Data.DataTable();
                            using (SqlConnection connection = new SqlConnection(connString))
                            {

                                connection.Open();
                                Excel.Application App = null;
                                Excel.Workbook workbook = null;
                                Excel.Worksheet worksheet = null;
                         
                                    App = new Excel.Application();
                                    workbook = App.Workbooks.Open(xlspath.Text);
              

                            for (int i = 1; i <= workbook.Sheets.Count; i++)
                            {

                                worksheet = workbook.Worksheets.Item[i] as Excel.Worksheet;
                                if (worksheet.Name.Substring(0, 2) == "BT")
                                { 
                                    //디비 삭제
                                    selectqry = "DELETE FROM " + worksheet.Name;

                                SqlCommand command = new SqlCommand(selectqry, connection);
                                command.ExecuteNonQuery();
                                command.Dispose();


                                Excel.Range rng = worksheet.UsedRange;
                                object[,] data = rng.Value;
                                    if(data == null)
                                        continue;
                                    else
                                    for (int r = 2; r <= data.GetLength(0); r++) //행의 갯수
                                    {
                                            int h = 0;

                                        Qry = "INSERT INTO " + worksheet.Name + " VALUES(";
                                        for (int c = 1; c <= data.GetLength(1); c++)
                                        {
                                            if (c == data.GetLength(1))
                                            {
                                                if (data[r, c] == null)
                                                {
                                                    Qry += "'')";
                                                }
                                                else
                                                {
                                                    if (worksheet.Name == "BT1001" && c == 6)
                                                    {
                                                            if (data[r, c].ToString().Substring(11, 2) == "오후")
                                                            {
                                                                if (data[r, c].ToString().Substring(15, 1) == ":")
                                                                {
                                                                    if (Convert.ToInt32(data[r, c].ToString().Substring(14, 1)) != 12)
                                                                        h = Convert.ToInt32(data[r, c].ToString().Substring(14, 1)) + 12;
                                                                    else
                                                                        h = Convert.ToInt32(data[r, c].ToString().Substring(14, 2));
                                                                    data[r, c].ToString().Replace(data[r, c].ToString().Substring(14, 1), h.ToString());
                                                                    Qry += "'" + data[r, c].ToString().Replace(data[r, c].ToString().Substring(14, 1), h.ToString()).Remove(11, 2) + "',";
                                                                }
                                                                else
                                                                {
                                                                    if(Convert.ToInt32(data[r, c].ToString().Substring(14, 2))!=12)
                                                                    h = Convert.ToInt32(data[r, c].ToString().Substring(14, 2)) + 12;
                                                                    else
                                                                        h = Convert.ToInt32(data[r, c].ToString().Substring(14, 2));

                                                                    data[r, c].ToString().Replace(data[r, c].ToString().Substring(14, 2), h.ToString());

                                                                    Qry += "'" + data[r, c].ToString().Replace(data[r, c].ToString().Substring(14, 2), h.ToString()).Remove(11, 2) + "',";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                Qry += "'" + data[r, c].ToString().Remove(11, 2) + "',";
                                                            }
                                                        }
                                                    else
                                                    {
                                                        Qry += "'" + data[r, c].ToString() + "')";
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (data[r, c] == null)
                                                {
                                                    Qry += "'',";
                                                }
                                                else
                                                {
                                                    if (worksheet.Name == "BT1001" && c == 6)
                                                        {
                                                            if (data[r, c].ToString().Substring(11, 2) == "오후")
                                                            {
                                                                if (data[r, c].ToString().Substring(15, 1) == ":")
                                                                {
                                                                    if (Convert.ToInt32(data[r, c].ToString().Substring(14, 1)) != 12)
                                                                        h = Convert.ToInt32(data[r, c].ToString().Substring(14, 1)) + 12;
                                                                    else
                                                                        h = Convert.ToInt32(data[r, c].ToString().Substring(14, 2));
                                                                    data[r, c].ToString().Remove(14, 1).Insert(14, h.ToString());
                                                              //     data[r, c].ToString().Replace(data[r, c].ToString().Substring(14, 1), h.ToString());

                                                                    Qry += "'" + data[r, c].ToString().Remove(14, 1).Insert(14, h.ToString()).Remove(11, 2) + "',";
                                                                }
                                                                else
                                                                {
                                                                    if (Convert.ToInt32(data[r, c].ToString().Substring(14, 2)) != 12)
                                                                        h = Convert.ToInt32(data[r, c].ToString().Substring(14, 2)) + 12;
                                                                    else
                                                                        h = Convert.ToInt32(data[r, c].ToString().Substring(14, 2));


                                                                    data[r, c].ToString().Remove(14, 2).Insert(14, h.ToString());
                                                                    //data[r, c].ToString().Replace(data[r, c].ToString().Substring(14, 2), h.ToString());

                                                                    Qry += "'" + data[r, c].ToString().Remove(14, 2).Insert(14, h.ToString()).Remove(11, 2) + "',";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                Qry += "'" + data[r, c].ToString().Remove(11, 2) + "',";
                                                            }
                                                        }
                                                     else
                                                    {
                                                        Qry += "'" + data[r, c].ToString() + "',";
                                                    }
                                                }
                                            }


                                        }
                                        //디비 삽입
                                        command = new SqlCommand(Qry, connection);
                                        command.ExecuteNonQuery();

                                    }
                                    
                                command.Dispose();
                            }
                                    }
                            MessageBox.Show("DB Import 완료", "Import 실시",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.None,
                              MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                         //   workbook.Save();
                            workbook.Close(0);
                            
                           // System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                            //GC.Collect();
                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                            //GC.Collect();
                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(App);
                            //GC.Collect();
                            App.Quit();
                            ExcelKill();

                        }
                           backupko = false;
                        }
                        else
                        {
                            MessageBox.Show("백업을 먼저 실시해 주십시오..", "백업 실시",
                              MessageBoxButtons.OK,
                              MessageBoxIcon.None,
                              MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                        }
                    }
                    else
                    {
                        MessageBox.Show("경로를 설정해 주세요.", "경로설정",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.None,
                               MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                }
           // }
            //catch(Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //   // mState = state.Disconnect;
            //}
        }

        private void DB_Back(object sender, EventArgs e)
        {
            try
            {
                if (mState == state.Normal)
                {
                    if (xlspath.Text != "")
                    {
                        using (SqlConnection connection = new SqlConnection(connString))
                        {

                            String path = openFileDialog2.FileName;
                            String pathOnly = Path.GetDirectoryName(path);
                            connection.Open();
                            selectqry = "BACKUP DATABASE " + DNameTxt.Text + " TO DISK ='" + pathOnly + "\\" + DNameTxt.Text + ".bak';"; //백업
                            SqlCommand command = new SqlCommand(selectqry, connection);
                            command.CommandTimeout = 300;
                            command.ExecuteNonQuery();
                            command.Dispose();
                            connection.Close();
                            MessageBox.Show(DNameTxt+" 백업 완료\n 백업 파일 경로 : " + pathOnly);
                            backupko = true;
                        }

                    }
                }
                else
                {
                    MessageBox.Show("경로를 설정해 주세요.","경로설정",
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.None,
                                 MessageBoxDefaultButton.Button1,MessageBoxOptions.DefaultDesktopOnly);
                }
                    
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                // mState = state.Disconnect;
            }
            
        }

        private void DBbck_CheckedChanged(object sender, EventArgs e)
        {
            if(DBbck.Checked == true)
            {
                backupko = true;
            }
            if(DBbck.Checked == false)
            {
                backupko = false;
            }
        }
    }    
}
