using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Exc = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;
using System.Timers;

namespace Sorting
{
    public partial class Form1 : Form
    {

        private delegate void UpdateDataGridViewCallback(int[] array);
        private delegate void UpdateLabelCallback(long time);

        public Form1()
        {
            InitializeComponent();
           
        }

        public double Interval { get; set; }
        private static UpdateDataGridViewCallback UpdateDataGridView5 { get; set; }

        

        
        public void Excel()
        {
            try
            {
                string str;
                int rCnt;
                int cCnt;

                OpenFileDialog opf = new OpenFileDialog();
                opf.Filter = "Файл Excel|*.XLSX;*.XLS";
                opf.ShowDialog();
                System.Data.DataTable tb = new System.Data.DataTable();
                string filename = opf.FileName;

                Exc.Application ExcelApp = new Exc.Application();
                Exc._Workbook ExcelWorkBook;
                Exc.Worksheet ExcelWorkSheet;
                Exc.Range ExcelRange;

                ExcelWorkBook = ExcelApp.Workbooks.Open(filename, 0, true, 5, "", "", true, Exc.XlPlatform.xlWindows, "\t", false,
                    false, 0, true, 1, 0);
                ExcelWorkSheet = (Exc.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

                ExcelRange = ExcelWorkSheet.UsedRange;
                for (rCnt = 1; rCnt <= ExcelRange.Rows.Count; rCnt++)
                {
                    dataGridView1.Rows.Add(1);
                    for (cCnt = 1; cCnt <= 2; cCnt++)
                    {
                        str = (string)(ExcelRange.Cells[rCnt, cCnt] as Exc.Range).Text;
                        dataGridView1.Rows[rCnt - 1].Cells[cCnt - 1].Value = str;
                    }
                }
                ExcelWorkBook.Close(true, null, null);
                ExcelApp.Quit();

                releaseObject(ExcelWorkSheet);
                releaseObject(ExcelWorkBook);
                releaseObject(ExcelApp);
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Невозможно очистить " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }

        }

        static void Swap(ref int e1, ref int e2)
        {
            try
            {
                var temp = e1;
                e1 = e2;
                e2 = temp;
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }


        public void Bubblesort()
        {
            try { 
            long ellapledTicks = DateTime.Now.Ticks;

            string[] x;

            int rows = dataGridView1.Rows.Count;

            x = new string[dataGridView1.RowCount - 1];

            for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
            {
                x[i] = dataGridView1[0, i].Value.ToString();
            }

            int[] array = x.Select(i => int.Parse(i)).ToArray();

            var len = array.Length;
            for (var i = 1; i < len; i++)
            {
                for (var j = 0; j < len - i; j++)
                {
                    if (array[j] > array[j + 1])
                    {

                        Swap(ref array[j], ref array[j + 1]);
                        System.Threading.Thread.Sleep(1000);
                        dataGridView1.Invoke(new UpdateDataGridViewCallback(UpdateDataGridView2), array);
                    }
                }
            }



            ellapledTicks = DateTime.Now.Ticks - ellapledTicks;
            label8.Invoke(new UpdateLabelCallback(UpdateLabel8), ellapledTicks);
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }

        private void UpdateDataGridView2(int[] array1)
        {
            try
            {
                dataGridView2.Rows.Add(dataGridView1.RowCount);
                for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
                {
                    //
                    dataGridView2[0, i].Value = array1[i];
                }
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }

        }

        static void Swap2(int[] array, int i, int j)
        {
            try
            {
                int temp = array[i];
                array[i] = array[j];
                array[j] = temp;
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }

        public void Vstavka()
        {
            try
            {
                long ellapledTicks = DateTime.Now.Ticks;
                string[] x;

                int rows = dataGridView1.Rows.Count;

                x = new string[dataGridView1.RowCount - 1];

                for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
                {
                    x[i] = dataGridView1[0, i].Value.ToString();

                }

                int[] array = x.Select(i => int.Parse(i)).ToArray();

                int key;
                int j;
                for (int i = 1; i < array.Length; i++)
                {
                    key = array[i];
                    j = i;
                    while (j > 0 && array[j - 1] > key)
                    {
                        Swap2(array, j, j - 1);
                        j -= 1;
                        System.Threading.Thread.Sleep(1000);
                        dataGridView1.Invoke(new UpdateDataGridViewCallback(UpdateDataGridView3), array);
                    }
                    array[j] = key;
                    System.Threading.Thread.Sleep(1000);
                    dataGridView1.Invoke(new UpdateDataGridViewCallback(UpdateDataGridView3), array);
                }
                ellapledTicks = DateTime.Now.Ticks - ellapledTicks;
                label9.Invoke(new UpdateLabelCallback(UpdateLabel9), ellapledTicks);
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }

        private void UpdateDataGridView3(int[] array1)
        {
            try
            {
                for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
                {
                    dataGridView3.Rows.Add();
                    dataGridView3[0, i].Value = array1[i];
                }
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }

        }


        public void Shake()
        {
            try
            {
                long ellapledTicks = DateTime.Now.Ticks;
                string[] x;

                int rows = dataGridView1.Rows.Count;

                x = new string[dataGridView1.RowCount - 1];

                for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
                {
                    x[i] = dataGridView1[0, i].Value.ToString();

                }

                int[] array = x.Select(i => int.Parse(i)).ToArray();

                for (var i = 0; i < array.Length / 2; i++)
                {
                    var swapFlag = false;
                    //проход слева направо
                    for (var j = i; j < array.Length - i - 1; j++)
                    {
                        if (array[j] > array[j + 1])
                        {
                            Swap(ref array[j], ref array[j + 1]);
                            swapFlag = true;
                        }
                        System.Threading.Thread.Sleep(1000);
                        dataGridView1.Invoke(new UpdateDataGridViewCallback(UpdateDataGridView4), array);
                    }

                    //проход справа налево
                    for (var j = array.Length - 2 - i; j > i; j--)
                    {
                        if (array[j - 1] > array[j])
                        {
                            Swap(ref array[j - 1], ref array[j]);
                            swapFlag = true;
                            System.Threading.Thread.Sleep(1000);
                            dataGridView1.Invoke(new UpdateDataGridViewCallback(UpdateDataGridView4), array);
                        }
                    }

                    //если обменов не было выходим
                    if (!swapFlag)
                    {
                        break;
                    }
                }

                ellapledTicks = DateTime.Now.Ticks - ellapledTicks;
                label10.Invoke(new UpdateLabelCallback(UpdateLabel10), ellapledTicks);
                //dataGridView1.Invoke(new UpdateDataGridViewCallback(UpdateDataGridView4), array);
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }

        private void UpdateDataGridView4(int[] array1)
        {
            try
            {
                dataGridView4.Rows.Add(dataGridView1.RowCount);

                for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
                {

                    dataGridView4[0, i].Value = array1[i];
                }
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }

        }

        static int Partition(int[] array, int minIndex, int maxIndex)
        {
            try
            {
                var pivot = minIndex - 1;
                for (var i = minIndex; i < maxIndex; i++)
                {
                    if (array[i] < array[maxIndex])
                    {
                        pivot++;
                        Swap(ref array[pivot], ref array[i]);
                    }
                }

                pivot++;
                Swap(ref array[pivot], ref array[maxIndex]);
                return pivot;
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
                return 0;
            }
        }

        public int[] QuickSort(int[] array, int minIndex, int maxIndex)
        {
            try
            {
                if (minIndex >= maxIndex)
                {

                    dataGridView1.Invoke(new UpdateDataGridViewCallback(UpdateDataGridView51), array);
                    return array;
                }
                System.Threading.Thread.Sleep(1000);
                var pivotIndex = Partition(array, minIndex, maxIndex);
                dataGridView1.Invoke(new UpdateDataGridViewCallback(UpdateDataGridView51), array);
                QuickSort(array, minIndex, pivotIndex - 1);
                System.Threading.Thread.Sleep(1000);
                dataGridView1.Invoke(new UpdateDataGridViewCallback(UpdateDataGridView51), array);
                QuickSort(array, pivotIndex + 1, maxIndex);
                System.Threading.Thread.Sleep(1000);
                dataGridView1.Invoke(new UpdateDataGridViewCallback(UpdateDataGridView51), array);


                return array;
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
                return array;
            }
        }

        public void Fast()
        {
            try
            {
                long ellapledTicks = DateTime.Now.Ticks;
                string[] x;

                int rows = dataGridView1.Rows.Count;

                x = new string[dataGridView1.RowCount - 1];

                for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
                {
                    x[i] = dataGridView1[0, i].Value.ToString();

                }

                int[] array = x.Select(i => int.Parse(i)).ToArray();

                QuickSort(array, 0, array.Length - 1);

                int N = dataGridView1.RowCount;

                ellapledTicks = DateTime.Now.Ticks - ellapledTicks;
                label11.Invoke(new UpdateLabelCallback(UpdateLabel11), ellapledTicks);
                //dataGridView1.Invoke(new UpdateDataGridViewCallback(UpdateDataGridView51), array);
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }

        private void UpdateDataGridView51(int[] array1)
        {
            try
            {
                dataGridView5.Rows.Add(dataGridView1.RowCount);
                for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
                {
                    //dataGridView5.Rows.Add();
                    dataGridView5[0, i].Value = array1[i];
                }
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }

        }

        //метод для проверки упорядоченности массива
        static bool IsSorted(int[] a)
        {
            try
            {
                for (int i = 0; i < a.Length - 1; i++)
                {
                    if (a[i] > a[i + 1])
                        return false;
                }

                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
                return false;
            }
        }

        //перемешивание элементов массива
        static int[] RandomPermutation(int[] a)
        {
            try
            {
                Random random = new Random();
                var n = a.Length;
                while (n > 1)
                {
                    n--;
                    var i = random.Next(n + 1);
                    var temp = a[i];
                    a[i] = a[n];
                    a[n] = temp;
                }

                return a;
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
                return a;
            }
        }

        //случайная сортировка
        static int[] BogoSorting(int[] a)
        {
            try
            {
                while (!IsSorted(a))
                {
                    a = RandomPermutation(a);

                }

                return a;
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
                return a;
            }
        }

        public void BOGOSORT()
        {
            try
            {
                long ellapledTicks = DateTime.Now.Ticks;

                string[] x;

                int rows = dataGridView1.Rows.Count;

                x = new string[dataGridView1.RowCount - 1];

                for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
                {
                    x[i] = dataGridView1[0, i].Value.ToString();

                }

                int[] array = x.Select(i => int.Parse(i)).ToArray();

                int[] a = array;

                while (!IsSorted(a))
                {
                    a = RandomPermutation(a);
                    System.Threading.Thread.Sleep(1000);
                    dataGridView1.Invoke(new UpdateDataGridViewCallback(UpdateDataGridView6), array);
                }

                ellapledTicks = DateTime.Now.Ticks - ellapledTicks;
                label12.Invoke(new UpdateLabelCallback(UpdateLabel12), ellapledTicks);
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }

        }

        private void UpdateLabel12(long time)
        {
            label12.Text = Convert.ToString(time);

        }
        private void UpdateLabel11(long time)
        {
            label11.Text = Convert.ToString(time);

        }
        private void UpdateLabel10(long time)
        {
            label10.Text = Convert.ToString(time);

        }

        private void UpdateLabel9(long time)
        {
            label9.Text = Convert.ToString(time);

        }

        private void UpdateLabel8(long time)
        {
            label8.Text = Convert.ToString(time);

        }

        private void UpdateDataGridView6(int[] array1)
        {
            try
            {
                dataGridView6.Rows.Add(dataGridView1.RowCount);
                for (int i = 0; i < dataGridView1.RowCount - 1; ++i)
                {

                    dataGridView6[0, i].Value = array1[i];
                }
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }

        }
        private void menuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (checkBox1.Checked == true)
                {

                    System.Threading.Thread readThread = new System.Threading.Thread(new System.Threading.ThreadStart(Bubblesort));
                    readThread.Start();

                }
                if (checkBox2.Checked == true)
                {

                    System.Threading.Thread readThread = new System.Threading.Thread(new System.Threading.ThreadStart(Vstavka));
                    readThread.Start();

                }
                if (checkBox3.Checked == true)
                {

                    System.Threading.Thread readThread = new System.Threading.Thread(new System.Threading.ThreadStart(Shake));
                    readThread.Start();

                }
                if (checkBox4.Checked == true)
                {

                    System.Threading.Thread readThread = new System.Threading.Thread(new System.Threading.ThreadStart(Fast));
                    readThread.Start();

                }
                if (checkBox5.Checked == true)
                {

                    System.Threading.Thread readThread = new System.Threading.Thread(new System.Threading.ThreadStart(BOGOSORT));
                    readThread.Start();


                }
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void чистимToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.Rows.Clear();
                dataGridView2.Rows.Clear();
                dataGridView3.Rows.Clear();
                dataGridView4.Rows.Clear();
                dataGridView5.Rows.Clear();
                dataGridView6.Rows.Clear();
                label8.Text = null;
                label9.Text = null;
                label10.Text = null;
                label11.Text = null;
                label12.Text = null;
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }

        }

        private void выбратьExcelДокументToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Excel();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }

        public void TimeTest()
        {
            try {
                long bub = 9999999999999999;
                long vst = 9999999999999999;
                long sh = 9999999999999999;
                long fa = 9999999999999999;
                long bog = 9999999999999999;


                if (label8.Text != "")
                {
                    bub = Convert.ToInt64(label8.Text);
                }

                if (label9.Text != "")
                {
                    vst = Convert.ToInt64(label9.Text);
                }
                if (label10.Text != "")
                {
                    sh = Convert.ToInt64(label10.Text);
                }
                if (label11.Text != "")
                {
                    fa = Convert.ToInt64(label11.Text);
                }
                if (label12.Text != "")
                {
                    bog = Convert.ToInt64(label12.Text);
                }




                if (bub < vst && bub < sh && bub < fa && bub < bog)
                {
                    label15.Text = "Пузырьки";
                }
                if (vst < bub && vst < sh && vst < fa && vst < bog)
                {
                    label15.Text = "Вставочки";
                }
                if (sh < vst && sh < bub && sh < fa && sh < bog)
                {
                    label15.Text = "Шейк";
                }
                if (fa < vst && fa < sh && fa < bub && fa < bog)
                {
                    label15.Text = "Быстрая";
                }
                if (bog < vst && bog < sh && bog < fa && bog < bub)
                {
                    label15.Text = "BOGOSORT";
                }
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }

        void ReverseDGVRows(DataGridView dgv)
        {
            try
            {
                List<DataGridViewRow> rows = new List<DataGridViewRow>();
                rows.AddRange(dgv.Rows.Cast<DataGridViewRow>());
                rows.Reverse();
                dgv.Rows.Clear();
                dgv.Rows.AddRange(rows.ToArray());
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }
        private void видПоУбываниюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                for (int i = dataGridView2.Rows.Count - 1; i > -1; i--)
                {
                    DataGridViewRow row = dataGridView2.Rows[i];
                    if (!row.IsNewRow && row.Cells[0].Value == null)
                    {
                        dataGridView2.Rows.RemoveAt(i);
                    }
                }
                for (int i = dataGridView3.Rows.Count - 1; i > -1; i--)
                {
                    DataGridViewRow row = dataGridView3.Rows[i];
                    if (!row.IsNewRow && row.Cells[0].Value == null)
                    {
                        dataGridView3.Rows.RemoveAt(i);
                    }
                }
                for (int i = dataGridView4.Rows.Count - 1; i > -1; i--)
                {
                    DataGridViewRow row = dataGridView4.Rows[i];
                    if (!row.IsNewRow && row.Cells[0].Value == null)
                    {
                        dataGridView4.Rows.RemoveAt(i);
                    }
                }
                for (int i = dataGridView5.Rows.Count - 1; i > -1; i--)
                {
                    DataGridViewRow row = dataGridView5.Rows[i];
                    if (!row.IsNewRow && row.Cells[0].Value == null)
                    {
                        dataGridView5.Rows.RemoveAt(i);
                    }
                }
                for (int i = dataGridView6.Rows.Count - 1; i > -1; i--)
                {
                    DataGridViewRow row = dataGridView6.Rows[i];
                    if (!row.IsNewRow && row.Cells[0].Value == null)
                    {
                        dataGridView6.Rows.RemoveAt(i);
                    }
                }

                ReverseDGVRows(dataGridView2);
                ReverseDGVRows(dataGridView3);
                ReverseDGVRows(dataGridView4);
                ReverseDGVRows(dataGridView5);
                ReverseDGVRows(dataGridView6);
            }
            catch (Exception err)
            {
                MessageBox.Show($"Ошибочка вышла: \n {err.Message}");
            }
        }

        private void самаяБыстраяСортировкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TimeTest();
        }
    }
}
