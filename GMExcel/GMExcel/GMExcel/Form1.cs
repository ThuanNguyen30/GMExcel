using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Timer = System.Timers.Timer;




namespace GMExcel
{
    public partial class Form1 : Form
    {
        //KHAI BAO CAC HANG
        public int maxdongda = 0;
        public int maxcotda = 0;
        public int maxdongbt = 0;
        public int maxcotbt = 0;
        public string diem = "ĐIỂM";
        public int dongdiem, cotdiem, dongbdbt, cotbdbt, dongbdda, cotbdda;
        public object[,] valueArraybt;
        public object[,] valueArray;
        public int dongout = 1;
        public int cotout = 1;
        public bool flag = true;

        private BackgroundWorker worker;

        //public string pathcham = txtthumuccham.Text;

        //bien luu vi tri loi
        public List<int> dongloi;
        public List<int> cotloi;


        //file excel bai thi
        public List<char> charList;
        public Range excelRangeBT;



        public Form1()
        {
            InitializeComponent();
        }


        private void Browse_Click(object sender, EventArgs e)
        {

            int size = -1;
            DialogResult result = dapanfile.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = dapanfile.FileName;
                try
                {
                    //   string text = File.ReadAllText(file);
                    //  size = text.Length;
                    txtdapan.Text = file;
                }
                catch (IOException)
                {
                }
            }
            Console.WriteLine(size); // <-- Shows file size in debugging mode.
            Console.WriteLine(result); // <-- For debugging use.
        }



        private void btnbrowsefo_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] files = Directory.GetFiles(fbd.SelectedPath);
                    txtthumuccham.Text = fbd.SelectedPath;

                    //System.Windows.Forms.MessageBox.Show("Files found: " + files.Length.ToString(), "Message");
                }
            }
        }


        public void open_dap_an_file()
        {
            try
            {
                //0. khai bao cac bien can thiet
                //int dongdiem, cotdiem, dongbd, cotbd, dongbdda, cotbdda;
                string giatribatdau = txttieude.Text;


                //1.  Mo file dap an xac dinh vi tri bat dau du lieu
                //create the Application object we can use in the member functions.
                Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
                _excelApp.Visible = true;

                string fileName = txtdapan.Text;

                //open the workbook
                Workbook workbook = _excelApp.Workbooks.Open(fileName,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                //select the first sheet        
                Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

                //find the used range in worksheet
                Range excelRange = worksheet.UsedRange;

                //get an object array of all of the cells in the worksheet (their values)
                valueArray = (object[,])excelRange.get_Value(
                            XlRangeValueDataType.xlRangeValueDefault);
                maxdongda = worksheet.UsedRange.Rows.Count;
                maxcotda = worksheet.UsedRange.Columns.Count;
                dongdiem = cotdiem = 1; dongbdda = cotbdda = 0;
                //tim kiem vi tri bat dau du lieu
                for (int row = 1; row <= worksheet.UsedRange.Rows.Count; ++row)
                {
                    for (int col = 1; col <= worksheet.UsedRange.Columns.Count; ++col)
                    {
                        if (valueArray[row, col] != null)
                        {
                            if (valueArray[row, col].ToString() == txttieude.Text)
                            {
                                dongbdda = row + 1; cotbdda = col; break;
                            }
                        }
                    }
                    if (dongbdda != 0) { break; }
                }

                //clean up stuffs
                workbook.Close(false, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(workbook);

                _excelApp.Quit();
                Marshal.FinalReleaseComObject(_excelApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi tại hàm open_dap_an_file: " + ex.ToString());
            }
          
        }

        public void open_file_bai_thi(string pathfilebaithi) {
            try
            {
                //create the Application object we can use in the member functions.
                Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
                _excelApp.Visible = true;

                string fileName = pathfilebaithi;

                //open the workbook
                Workbook workbook = _excelApp.Workbooks.Open(fileName,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                //select the first sheet        
                Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

                //find the used range in worksheet
                Range excelRangeBT = worksheet.UsedRange;



                //get an object array of all of the cells in the worksheet (their values)
                valueArraybt = (object[,])excelRangeBT.get_Value(
                        XlRangeValueDataType.xlRangeValueDefault);
                maxdongbt = worksheet.UsedRange.Rows.Count;
                maxcotbt = worksheet.UsedRange.Columns.Count;
                dongbdbt = cotbdbt = 0;

                //tim kiem vi tri bat dau du lieu
                for (int row = 1; row <= maxdongbt; ++row)
                {
                    for (int col = 1; col <= maxcotbt; ++col)
                    {
                        if (valueArraybt[row, col] != null)
                        {
                            if (valueArraybt[row, col].ToString().ToUpper() == txttieude.Text.ToUpper())
                            {
                                dongbdbt = row + 1; cotbdbt = col; break;
                            }
                        }
                    }
                    if (dongbdbt != 0) { break; }
                }

                //TODO
                charList = new List<char>();
                for (int i = 1; i <= maxcotbt; ++i)
                {
                    if (excelRangeBT.Cells[dongbdbt, i].text == "")
                    {
                        charList.Add(' ');
                    }
                    else
                    {
                        charList.Add(excelRangeBT.Cells[dongbdbt, i].Formula[0]);
                    }

                }


                //clean up stuffs
                workbook.Close(false, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(workbook);

                _excelApp.Quit();
                Marshal.FinalReleaseComObject(_excelApp);
               // movefile(Path.GetFileName(pathfilebaithi));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi tại hàm open_file_bai_thi: " + ex.ToString());
            }

        }

        private void btnthoat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btbbrowsekq_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] files = Directory.GetFiles(fbd.SelectedPath);
                    txtthumucketqua.Text = fbd.SelectedPath;

                    //System.Windows.Forms.MessageBox.Show("Files found: " + files.Length.ToString(), "Message");
                }
            }
        }

        public int getmarknew(string tf)
        {
            int d = 0;
            bool ok = true;
            int dongchay = maxdongda - dongbdda;
            int cotchay = maxcotda - cotbdda;
            int diemss = Int32.Parse(txtdiemsapxep.Text);
            dongloi = new List<int>();
            cotloi = new List<int>();
            if (txtdiemsapxep.Text!="0")
            {
                ok = true;
                for (int col = 0; col < cotchay; ++col)
                {
                   
                    for (int row = 0; row <= dongchay; ++row)
                    {
                        if (valueArraybt[dongbdbt + row, cotbdbt + col] != null && valueArray[dongbdda + row, cotbdda + col] != null)
                        {
                            if (valueArraybt[dongbdbt + row, cotbdbt + col].ToString() != valueArray[dongbdda + row, cotbdda + col].ToString())
                            {
                                ok = false;
                                luuloi(dongbdbt + row, cotbdbt + col);
                            }
                        }
                    }
                   // if (ok && valueArray[dongdiem, cotbdda + col] != null && charList[cotbdbt + col-1] == '=')
                    //{
                        //d = d + Int32.Parse(valueArray[dongdiem, cotbdda + col].ToString());
                   // }else
                   // {
                        //Nguoc lai, neu khong co diem, thi boi mau o bi loi
                        //if (valueArray[dongdiem, cotbdda + col] != null)
                       // {

                       // }
                   }
                if (ok==true) { d = d + diemss; }

            }
       

            //CO DIEM SAP XEP HAY KHONG CUNG TINH CAI NAY
                //1.lấy ra list điểm cột  ở đáp án

                //Duyệt từng dòng điểm
                // Lấy ra list điểm của từng cột (list các cột chứa điểm)
                List<int> listdiem = new List<int>();
                for (int i = 2; i <= maxcotda; i++)
                {
                    if (valueArray[dongdiem, i] != null)
                    {
                        listdiem.Add(i);
                    }
                }
                //Duyệt từng cột điểm ở đáp án
                int columnda, rowda, cotbt = 0;
                int tdiem = 0;
                for (int i = 0; i < listdiem.Count; i++)
                {
                    //Xác định các giá trị
                    //1. cột đáp án:	columnda = listdiem[i];
                    columnda = listdiem[i];
                    //2. dòng đáp án bắt đầu: 
                    rowda = dongbdda;
                    //3. số lượng ô cần check: dongchay
                    //4. cột bài thi
                    cotbt = cotbdbt + (columnda - 2);
                    //5. dòng bài thi bắt đầu: dongbd

                    //lấy ô đầu tiên ở cột bài thi, nếu là công thức thì gọi hàm compare
                    if (charList[cotbt-1] == '=')
                    {
                        tdiem = compareList(columnda, rowda, dongchay, cotbt, dongbdbt);
                        if (tdiem == 0)
                        {
                            luuloi(dongbdbt, cotbt);
                        }
                    }
                    else
                    {
                        // ngược lại cho 0 điểm cột đó.
                        tdiem = 0;
                        luuloi(dongbdbt, cotbt);
                    }
                    d = d + tdiem;
                }
            

            return d;
        }

        private void luuloi(int dong, int cot)
        {
            dongloi.Add(dong);
            cotloi.Add(cot);
            //throw new NotImplementedException();
        }

        private void boimau(string tf)
        {

            //1.  Mo file dap an xac dinh vi tri bat dau du lieu
            //create the Application object we can use in the member functions.
            Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
            _excelApp.Visible = true;

            string fileName = tf;

            //open the workbook
            Workbook workbook = _excelApp.Workbooks.Open(fileName,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            //select the first sheet        
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

            //find the used range in worksheet
            Range excelRange = worksheet.UsedRange;
            for (int i = 0; i < dongloi.Count; i++)
            {
               // worksheet.Cells[1, 1].Font.Color = Color.Red;
                worksheet.Cells[dongloi[i], cotloi[i]].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }


            //excelRange.Cells[1,1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            //clean up stuffs
            workbook.Close(true, Type.Missing, Type.Missing);
            Marshal.ReleaseComObject(workbook);

            _excelApp.Quit();
            Marshal.FinalReleaseComObject(_excelApp);
          //  throw new NotImplementedException();
        }

        private int compareList(int columnda, int rowda, int dongchay, int cotbt, int dongbd)
        {
            
            System.Collections.Generic.List<string> lsda = new System.Collections.Generic.List<string>();
            System.Collections.Generic.List<string> lsbt = new System.Collections.Generic.List<string>();
            int i = 0;
            while ((i<=dongchay) && (valueArray[rowda + i, columnda] != null))
            {
                lsda.Add(valueArray[rowda + i, columnda].ToString());
                lsbt.Add(valueArraybt[dongbd + i, cotbt].ToString());
                i++;
            }
                
            lsda.Sort();
            lsbt.Sort();
            
            for ( i = 0; i < lsda.Count; i++)
            {
                if (lsda[i] != lsbt[i])
                {
                    return 0;
                }
            }

            return Int32.Parse(valueArray[dongdiem, columnda].ToString());
        }

        public void ghidiem(string sbd, int diem,string tenfile,string fullpathfile)
        {
            try
            {
                string path = txtthumucketqua.Text + "\\" + txttenlop.Text;
                DateTime localDate = DateTime.Now;

                // This text is added only once to the file.
                if (!File.Exists(path))
                {
                    // Create a file to write to.
                    using (StreamWriter sw = File.CreateText(path))
                    {

                        sw.WriteLine("ID,DIEM,THOI GIAN");
                        sw.WriteLine(sbd + "," + diem + "," + localDate.ToString());

                    }
                }
                else
                {
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine(sbd + "," + diem + "," + localDate.ToString());
                    }

                }
                outlog(sbd, diem, tenfile);
                movefile(Path.GetFileName(fullpathfile));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi tại hàm ghidiem: " + ex.ToString());
            }
        }

        public void Save_file_ket_qua()
        {
            MessageBox.Show("Đã chấm xong, bạn có thể xem file kết quả tại: " + txtthumucketqua.Text + "\\ketqua.xls");
        }

        public void outlog(string ts, int diemts, string tf)
        {
            try
            {
                //1. kiểm tra thư mục Logs có trong thư mục KETQUA không nếu không có thì tạo
                string path = txtthumucketqua.Text + "\\Logs\\";
                bool folderExists = Directory.Exists(path);
                if (!folderExists)
                    Directory.CreateDirectory(path);
                //2. ghi diem cua thi sinh vao file log
                string tenfile = path + tf+".log";
                string noidung = ts + "‣" + "DIEM: " + diemts.ToString();
                System.IO.File.WriteAllText(tenfile, noidung);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi tại hàm outlog: " + ex.ToString());
            }



        }

        public void movefile(string tenfile)
        {
            try
            {
                string path = txtthumuccham.Text + "\\" + tenfile;
                string path2 = txtthumucketqua.Text + "\\" + tenfile;
                if (File.Exists(path2))
                {
                    File.Delete(path2);
                }


                File.Move(path, path2);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi tại hàm movefile: " + ex.ToString());
            }

        }
        public string getnametowritelog(string st)
        {
            bool bl = st.Contains("]");
            if (bl == true)
            {
                string s = st.Substring(st.IndexOf("]")+1);
                return s;
            }
            else return st;
        }
        public void xulychinh()
        {
            int i = 0;
            string[] allfile = Directory.GetFiles(txtthumuccham.Text, "*.xls*");
            string path = txtthumuccham.Text;
            try
            {
                int diemtong = 0;
                //1.Mo file dap an xac dinh vi tri bat dau du lieu
                open_dap_an_file();

                //2. duyet trong thu muc cham, mo tung file bai thi xac dinh vi tri bat dau du lieu
                 path = txtthumuccham.Text;
                allfile = Directory.GetFiles(path, "*.xls*");
                //doi ten file bo dau "_"

                for (i = 0; i <= allfile.Length - 1; i++)
                {
                    string directorypath = Path.GetDirectoryName(allfile[i]);
                    string filename = Path.GetFileName(allfile[i]);
                    filename =filename.Replace("_"," ");

                    System.IO.File.Move(allfile[i], Path.Combine(directorypath, filename));
                }

                allfile = Directory.GetFiles(path, "*.xls*");

                for ( i = 0; i <= allfile.Length - 1; i++)
                {
                    open_file_bai_thi(allfile[i]);
                    diemtong = getmarknew(allfile[i]);
                   // if (diemtong != 10)
                    //{
                        boimau(allfile[i]);
                    //}
                   
                    ghidiem(getnametowritelog(Path.GetFileNameWithoutExtension(allfile[i])), diemtong, Path.GetFileNameWithoutExtension(allfile[i]), allfile[i]);
                }
            /*  if (allfile.Length != 0)
              {
                  Save_file_ket_qua();
              }*/

            //3. tinh diem mapping 1-1
        }
            catch (Exception ex)
            {
                allfile = Directory.GetFiles(path, "*.xls*");
                movefileerror(allfile[i]);
                //MessageBox.Show("Đã xảy ra lỗi tại hàm xulychinh: " + ex.ToString());
            }

}

        private void movefileerror(string p)
        {
            string path = txtthumucketqua.Text + "\\LOI\\";
            bool folderExists = Directory.Exists(path);
            if (!folderExists)
                Directory.CreateDirectory(path);
            //2. move file
          
             path = path + Path.GetFileName(p);
          //  string path2 = txtthumucketqua.Text + "\\" + tenfile;
            if (File.Exists(path))
            {
                File.Delete(path);
            }


            File.Move(p, path);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                //0. khai bao cac bien can thiet
                //int dongdiem, cotdiem, dongbd, cotbd, dongbdda, cotbdda;
                string giatribatdau = txttieude.Text;


                //1.  Mo file dap an xac dinh vi tri bat dau du lieu
                //create the Application object we can use in the member functions.
                Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
                _excelApp.Visible = true;

                string fileName = txtdapan.Text;

                //open the workbook
                Workbook workbook = _excelApp.Workbooks.Open(fileName,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                //select the first sheet        
                Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

                //find the used range in worksheet
                Range excelRange = worksheet.UsedRange;

                worksheet.Cells[1,1].Font.Color = Color.Red;
                worksheet.Cells[1,1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                //excelRange.Cells[1,1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                //clean up stuffs
                workbook.Close(true, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(workbook);

                _excelApp.Quit();
                Marshal.FinalReleaseComObject(_excelApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi tại hàm open_dap_an_file: " + ex.ToString());
            }
        }



        private void btnthc_Click(object sender, EventArgs e)
        {
            try
            {
                Mainabc();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi tại hàm btnthc_click: " + ex.ToString());
            }
            
        }


       public  void t_Elapsed(object sender, System.Timers.ElapsedEventArgs e)

        {  
            string path = txtthumuccham.Text;
            string[] allfile = Directory.GetFiles(path, "*.xls*");
            if (allfile.Length == 0)
            {

            }else
            {
                xulychinh();
            }

        }


        public void Mainabc()
        {
            worker = new BackgroundWorker();
            worker.DoWork += worker_DoWork;



            Timer timer = new Timer(1000); // 1 sec = 1000, 60 sec = 60000

            // t.AutoReset = true;

            timer.Elapsed += new System.Timers.ElapsedEventHandler(timer_Elapsed);

            timer.Start();



        }

       public  void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (!worker.IsBusy)
                worker.RunWorkerAsync();
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            //whatever You want the background thread to do...
            string path = txtthumuccham.Text;
            string[] allfile = Directory.GetFiles(path, "*.xls*");
            if (allfile.Length == 0)
            {

            }
            else
            {
                //string namefile =
                xulychinh();
            }
        }

    }
}
