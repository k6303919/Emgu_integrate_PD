using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text.RegularExpressions;
using Emgu;
using Emgu.CV;
using Emgu.CV.Structure;
using Emgu.CV.Util;
using Emgu.Util;
using Emgu.CV.CvEnum;
using Emgu.CV.UI;
using Emgu.CV.Shape;
using Emgu.CV.XImgproc;
using Excel = Microsoft.Office.Interop.Excel;

namespace Emgu_integrate
{
    public partial class Form1 : Form
    {
       

        Image<Bgr, byte> imgInput = null;
        Image<Gray, byte> imgGray;
        //Image<Gray, byte> imghiseq;
        //Image<Gray, byte> imgBinarize;
        Image<Gray, byte> imgOutput = null;
        Image<Gray, byte> imgOutput2 = null;
        Image<Gray, byte> imgOutput3 = null;
        Image<Gray, byte> imgOutput4 = null;
        Image<Gray, byte> imgOutput5 = null;
        Image<Gray, byte> imgOutput6 = null;
        int num1 = 0;
        int count = 1;

        Point point1;
        Point point2;
        Point point3;
        Point point4;
        Point point5;
        Point point6;

        Rectangle Rect = new Rectangle();
        //Rectangle Rect2 = new Rectangle();
        Rectangle RealImageRect = new Rectangle();
        //Rectangle RealImageRect2 = new Rectangle();

        private Brush selectionBrush = new SolidBrush(Color.FromArgb(128, 64, 64, 64));
        private Brush selectionBrush2 = new SolidBrush(Color.FromArgb(128, 64, 64, 64));

        List<Image<Bgr, byte>> imageList1 = new List<Image<Bgr, byte>>();
        List<string> path = new List<string>();
        double VS = 0; //全腦組織面積
        //double HY1 = 0, HY2 = 0, HY3 = 0;
        double VTTR;
        double VTTL;
        double SUR_1;
        double SUR_2;
        double SUR_3;
        double SUR_4;
        double SUR1;
        double SUR2;
        double SUR3;
        double test5;
        double test6;
        double test11;
        double test12;
        double VTTO;
        double VTTT;
        double VPTT;
        double VPTT1;
        double VPTT2;
        double VPT1;
        double VPT2;
        double VPT0;
        double test1;
        double test2; 
        double test3;
        double test4;
        double test7; 
        double test8;
        double test9;
        double test10;
        double test66;
        double ASI1;
        double VPP;
        List<double> VT = new List<double>(); //腦組織
        List<double> VP = new List<double>(); //腦實質
        double VolCN ; //尾狀核體積
        double VolPM ; //殼核體積
        double VolOB ; //枕骨體積
        double VolSN ; //黑質體積
        double VPPCN ; //萎縮比尾狀核
        double VPPPM ; //萎縮比殼核
        double VPPOB ; //萎縮比枕骨
        double VPPSN ; //萎縮比黑質 
        List<double> VT1 = new List<double>();
        List<double> VT2 = new List<double>();
        List<double> VT3 = new List<double>();
        List<double> VTe = new List<double>();
        
        List<double> VPe = new List<double>();
        List<double> VP1 = new List<double>();
        List<double> VP2 = new List<double>();
        List<double> VP3 = new List<double>();
        List<double> vt = new List<double>();
        List<double> vp = new List<double>();
        List<double> VVT = new List<double>();
        List<double> VVT2 = new List<double>();
        List<double> VVT3 = new List<double>();
        List<double> CD_N = new List<double>();
        List<double> CD_N_PU_TR = new List<double>();
        List<double> CD_N_PU_TL = new List<double>();
        List<double> STN = new List<double>();
        List<double> VTS = new List<double>();
        List<double> PU_T = new List<double>();
        List<double> OC_B = new List<double>();
        List<double> OB_B = new List<double>();
        List<double> VT_T = new List<double>();
        List<double> CS_F = new List<double>();
        List<double> GM_M = new List<double>();
        List<double> WM_M = new List<double>();
        List<double> CS_F2 = new List<double>();
        List<double> Bg_a = new List<double>();
        List<double> DADS = new List<double>();
        List<double> DADSs = new List<double>();
        public Form1()
        {
            InitializeComponent();
            button8.Click += new System.EventHandler(button2_Click);
            button8.Click += new System.EventHandler(button3_Click);
            //button18.Click += new System.EventHandler(button2_Click);
            //button18.Click += new System.EventHandler(button3_Click);
            button7.Click += new System.EventHandler(numericUpDown4_ValueChanged);
            button7.Click += new System.EventHandler(button2_Click);
            button7.Click += new System.EventHandler(button3_Click);
            button19.Click += new System.EventHandler(numericUpDown4_ValueChanged);
            button19.Click += new System.EventHandler(button2_Click);
            button19.Click += new System.EventHandler(button3_Click);
            button19.Click += new System.EventHandler(button7_Click);
            //button21.Click += new System.EventHandler(numericUpDown4_ValueChanged);
            //button21.Click += new System.EventHandler(button2_Click);
            //button21.Click += new System.EventHandler(button3_Click);
            //button21.Click += new System.EventHandler(button7_Click);

            button7.Click += new System.EventHandler(numericUpDown4_ValueChanged);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                imgInput = new Image<Bgr, byte>(ofd.FileName);
                imageBox1.Image = imgInput;
            }
        }
        private void imageBox1_MouseDown(object sender, MouseEventArgs e)
        {
            point1 = e.Location;
            Invalidate();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Image<Gray, byte> image_Copy = imgInput.Convert<Gray, byte>();
            Image<Gray, byte> mask = new Image<Gray, byte>(imgInput.Width + 2, imgInput.Height + 2);
            Image<Gray, byte> imgasd;


            //Point point1 = new Point(256, 320);

            Rectangle dummy = new System.Drawing.Rectangle(0, 0, 0, 0);
            CvInvoke.FloodFill(image_Copy, mask, point1, new MCvScalar(1), out dummy,
                new MCvScalar(12),
                new MCvScalar(7),
                Emgu.CV.CvEnum.Connectivity.EightConnected,
                Emgu.CV.CvEnum.FloodFillType.Default);
            imgasd = (imgInput.Convert<Gray, byte>()) - image_Copy;
           
            //imageBox3.Image = imgssd;
            //MRI 分析
            //for (int i = 0; i < imgasd.Height; i++)
            {
                //for (int j = 0; j < imgasd.Width; j++)
                {
                    //if (imgasd.Data[i, j, 0] > 107) //GM 107 WM 101
                    {
                        //imgasd.Data[i, j, 0] = 0;
                    }
                }
            }


            imageBox2.Image = imgasd;

            imgOutput = imgasd;
        }
        private void button3_Click(object sender, EventArgs e)
        {

            if (imgOutput == null)
            {
                return;
            }

            Image<Gray, byte> imgBgr = imgOutput;
            Image<Gray, byte> median = new Image<Gray, byte>(imgOutput.Width, imgOutput.Height, new Gray(0));
            Image<Gray, byte> blur = new Image<Gray, byte>(imgOutput.Width, imgOutput.Height, new Gray(0));
            median = imgBgr.SmoothMedian((int)numericUpDown1.Value);
            //blur = median.SmoothBlur(3, 3);
            blur = median.Copy();

            imgOutput  = blur.Copy();
            imgOutput2 = blur.Copy();
            imgOutput3 = blur.Copy();
           

            imageBox2.Image = imgOutput;
            imageBox3.Image = imgOutput2;
            //imageBox5.Image = imgOutput3;
            



        }

        private void button4_Click(object sender, EventArgs e)
        {
            Image<Gray, byte> imgray = imgOutput.ThresholdBinary(new Gray(25), new Gray(255));
            Image<Gray, byte> imgray2 = imgOutput.ThresholdBinary(new Gray((int)numericUpDown7.Value), new Gray(255));
            double AT = 0;
            double AV = 0;
            int area = 0;
            for (int i = 0; i < imgray.Rows; i++)
            {
                for (int j = 0; j < imgray.Cols; j++)
                {
                    if (imgray.Data[i, j, 0] == 255)
                        area++;

                }
            }
            int area1 = 0;
            for (int i = 0; i < imgray2.Rows; i++)
            {
                for (int j = 0; j < imgray2.Cols; j++)
                {
                    if (imgray2.Data[i, j, 0] == 255)
                        area1++;

                }
            }
            AT = System.Convert.ToDouble(area) * 0.2304;
            AV = Convert.ToDouble(area1) * 0.2304;
            
            listBox15.Items.Add(AT.ToString("f4"));
            listBox16.Items.Add(AV.ToString("f4"));
            VT.Add(AT);
            VP.Add(AV);

            /*AT = System.Convert.ToDouble(area) * 0.2304; //腦組織
              AV = Convert.ToDouble(area1) * 0.2304; //腦實質比
              listBox1.Items.Add(AT.ToString("f4"));
              listBox2.Items.Add(AV.ToString("f4"));
              VT.Add(AT);
              VP.Add(AV);*/
        }

        private void button6_Click(object sender, EventArgs e)
        {

            FolderBrowserDialog folderBrowserDlg = new FolderBrowserDialog();
            DialogResult dlgResult = folderBrowserDlg.ShowDialog();
            //List<Image<Bgr, byte>> imageList1 = new List<Image<Bgr, byte>>();
            path.Clear();
            //Image<Bgr, byte> EmguImage = new Image<Bgr, byte>();
            if (dlgResult.Equals(DialogResult.OK))
            {
                foreach (string file in System.IO.Directory.GetFiles(folderBrowserDlg.SelectedPath, "*.jpg")) //.png, bmp, etc.
                {

                    path.Add(file);




                }
                //path.Sort(new ReverseStringComparer());
                //path = path.OrderBy<string, string > (f => f).ToList(); ;
                path.Sort((a, b) => new StringNum(a).CompareTo(new StringNum(b)));
                numericUpDown4.Maximum = path.Count();
                try
                {
                    imgInput = new Image<Bgr, byte>(path[(int)numericUpDown4.Value - 1]);

                    imageBox1.Image = imgInput;
                }
                catch (Exception ex)
                {

                    MessageBox.Show("資料夾內無符合格式影像");
                    //Application.Restart();

                }
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            numericUpDown4.Maximum = path.Count();
            if (numericUpDown4.Value < path.Count())
            {
                numericUpDown4.Value = numericUpDown4.Value + 1;
            }
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {

            numericUpDown4.Maximum = path.Count();
            try
            {
                imgInput = new Image<Bgr, byte>(path[(int)numericUpDown4.Value - 1]);
                imageBox1.Image = imgInput;
            }
            catch
            {
                //MessageBox.Show("資料夾內無符合格式影像");

            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                return;
            }
            Image<Gray, byte> image_copy2 = imgOutput2.Copy();
            Image<Gray, byte> mask = new Image<Gray, byte>(imgInput.Width + 2, imgInput.Height + 2);
            Image<Gray, byte> imgasd;
            Rectangle dummy = new System.Drawing.Rectangle(1, 1, 1, 1);
            CvInvoke.FloodFill(image_copy2, mask, point2, new MCvScalar(254), out dummy,
                new MCvScalar((double)numericUpDown2.Value),
                new MCvScalar((double)numericUpDown3.Value),
                Emgu.CV.CvEnum.Connectivity.FourConnected,
                Emgu.CV.CvEnum.FloodFillType.FixedRange);
            imgasd = image_copy2;
            imageBox3.Image = imgasd;
            imgOutput2 = imgasd.Copy();
        }

        private void imageBox2_MouseDown(object sender, MouseEventArgs e)
        {
            point2 = e.Location;
            Invalidate();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            
            imgOutput2 = imgOutput.Copy();
            imageBox3.Image = imgOutput2;
        }

        private void imageBox3_MouseUp(object sender, MouseEventArgs e)
        {
            if (checkBox1.Checked)
            {
                if (imgOutput2 == null)
                    return;
                //Define ROI. Valida altura e largura para evitar index range exception.
                if (RealImageRect.Width > 0 && RealImageRect.Height > 0)
                {
                    imgOutput2.ROI = RealImageRect;
                    imageBox4.Image = imgOutput2;

                }
            }
        }

        private void imageBox3_MouseMove(object sender, MouseEventArgs e)
        {
            int X0, Y0;
            Utilities.ConvertCoordinates(imageBox3, out X0, out Y0, e.X, e.Y);
            label1.Text = "Last Position: X:" + X0 + "  Y:" + Y0;
            if (checkBox1.Checked)
            {
                //Coordinates at input picture box
                if (e.Button != MouseButtons.Left)
                    return;
                Point tempEndPoint = e.Location;
                Rect.Location = new Point(
                    Math.Min(point2.X, tempEndPoint.X),
                    Math.Min(point2.Y, tempEndPoint.Y));
                Rect.Size = new Size(
                    Math.Abs(point2.X - tempEndPoint.X),
                    Math.Abs(point2.Y - tempEndPoint.Y));



                //Coordinates at real image - Create ROI
                Utilities.ConvertCoordinates(imageBox3, out X0, out Y0,
                point2.X, point2.Y);
                int X1, Y1;
                Utilities.ConvertCoordinates(imageBox3, out X1, out Y1, tempEndPoint.X, tempEndPoint.Y);
                RealImageRect.Location = new Point(
                    Math.Min(X0, X1),
                    Math.Min(Y0, Y1));
                RealImageRect.Size = new Size(
                    Math.Abs(X0 - X1),
                    Math.Abs(Y0 - Y1));
                ((PictureBox)sender).Invalidate();
            }
        }

        private void imageBox3_Paint(object sender, PaintEventArgs e)
        {
            if (checkBox1.Checked)
            {
                // Draw the rectangle...
                if (imageBox3.Image != null)
                {
                    if (Rect != null && Rect.Width > 0 && Rect.Height > 0)
                    {
                        //Seleciona a ROI
                        e.Graphics.SetClip(Rect, System.Drawing.Drawing2D.CombineMode.Exclude);
                        e.Graphics.FillRectangle(selectionBrush, new Rectangle
                (0, 0, ((PictureBox)sender).Width, ((PictureBox)sender).Height));
                        //e.Graphics.FillRectangle(selectionBrush, Rect);
                    }
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {


            Image<Gray, byte> imgGray = imgOutput2.Copy();
            Image<Gray, byte> imgbi = new Image<Gray, byte>(imgGray.Width, imgGray.Height);

            imgGray._GammaCorrect(3d); // 1d 2d 3d 對比度調整
            imageBox3.Image = imgGray;
            imgOutput2 = imgGray.Copy();
        }

        private void button10_Click(object sender, EventArgs e)
        {

            Image<Gray, byte> imgGray = imgOutput2.Copy();
            Image<Gray, byte> binary = new Image<Gray, byte>(imgGray.Width, imgGray.Height);
            //Image<Gray, byte> imghinq ;
            CvInvoke.Threshold(imgGray, binary, 103, 255, Emgu.CV.CvEnum.ThresholdType.Binary);
            imgOutput2 = binary.Copy();
            for (int i = 0; i < imgOutput2.Height; i++)
            {
                for (int j = 0; j < imgOutput2.Width; j++)
                {
                    if (imgOutput2.Data[i, j, 0] < 235) //根據前景閥值設定 來做調整 0~255
                    {
                        imgOutput2.Data[i, j, 0] = 0;
                    }

                }

            }
            //imghinq = new Image<Gray, byte>(imgGray.Width, imgGray.Height, new Gray(0));
            //CvInvoke.EqualizeHist(imgOutput2, imghinq);
            imageBox3.Image = imgOutput2;
            //imgOutput2 = imghinq.Copy();
        }

        private void button11_Click(object sender, EventArgs e)
        {

            if (checkedListBox2.CheckedItems.Count < 1)
            {
                MessageBox.Show("選擇腦室");
                return;
            }
            if (imgOutput2 == null)
            {
                return;
            }

            Image<Gray, byte> imgray = imgOutput2.ThresholdBinary(new Gray(253), new Gray(255));
            VectorOfVectorOfPoint con = new VectorOfVectorOfPoint();

            double area1 = new double();

            int area = 0;
            for (int i = 0; i < imgOutput2.Rows; i++)
            {
                for (int j = 0; j < imgOutput2.Cols; j++)
                {
                    if (imgray.Data[i, j, 0] == 255)
                        area++;

                }
            }



            area1 = area * 0.2304;
            textBox1.Text = (area1).ToString("f4");


            if (checkedListBox2.GetItemChecked(0))
            {
                CD_N_PU_TR.Add(area1);
                
                listBox12.Items.Add(area1 .ToString("f4"));

            }
            if (checkedListBox2.GetItemChecked(1))
            {
                CD_N_PU_TL.Add(area1);
                
                listBox4.Items.Add(area1 .ToString("f4"));

            }
            if (checkedListBox2.GetItemChecked(2))
            {
                OC_B.Add(area1);
                listBox20.Items.Add(area1 .ToString("f4"));
            }
                
            if (checkedListBox2.GetItemChecked(3))
            {
                STN.Add(area1);
                listBox1.Items.Add(area1.ToString("f4"));
            }
                



            checkBox1.Checked = false;
            for (int i = 0; i < checkedListBox2.Items.Count; i++)

            {

                checkedListBox2.SetItemChecked(i, false);

            }


        }



        public void ApplyRangeFitler(int min, int max)
        {
            try
            {
                //Image<Gray, byte> imggray = imgOutput3.Copy();
                Image<Gray, byte> image_Copy = imgInput.Convert<Gray, byte>();
                Image<Gray, byte> mask = new Image<Gray, byte>(imgInput.Width + 2, imgInput.Height + 2).SmoothMedian(5);
                imgOutput3 = imgInput.Convert<Gray, byte>().InRange(new Gray(min), new Gray(max));
                Image<Gray, byte> imgasd;
                imgasd = (imgInput.Convert<Gray, byte>()) - imgOutput3;
                imgOutput3 = imgasd.Convert<Gray, byte>();
                //imageBox5.Image = imgOutput3;
                //imageBox5.Invalidate();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex.Message);
            }
        }
        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                if (imgOutput3 == null)
                {
                    return;
                }
                Image<Bgr, byte> temp = imgInput.Clone();
                temp.SetValue(new Bgr(0, 0, 255), imgOutput3);
                //imageBox5.Image = temp;

            }
            catch (Exception)
            {

            }
        }

        private void button12_Click_1(object sender, EventArgs e)
        {
            formParameters fp = new formParameters(this);
            fp.Show();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            Image<Gray, byte> imgGray = imgOutput3.Copy();
            Image<Gray, byte> imghiseq;
            //for (int i = 0; i < imgGray.Height; i++)
            {
                //for (int j = 0; j < imgGray.Width; j++)
                {
                    //if (imgGray.Data[i, j, 0] < 115)
                    {
                        // imgGray.Data[i, j, 0] = 0;
                    }
                }
            }
            /*for (int i = 0; i < imgGray.Height; i++)
             {
                 for (int j = 0; j < imgGray.Width; j++)
                 {
                     if (imgGray.Data[i, j, 0] > 120)
                     {
                         imgGray.Data[i, j, 0] = 0;
                     }
                 }
             }*/
            imghiseq = new Image<Gray, byte>(imgGray.Width, imgGray.Height, new Gray(0));
            CvInvoke.EqualizeHist(imgGray, imghiseq);
            //imghiseq = imgGray;
            // for (int i = 0; i < imghiseq.Height; i++)
            {
                // for (int j = 0; j < imghiseq.Width; j++)
                {
                    // if (imghiseq.Data[i, j, 0] < 160)
                    {
                        //  imghiseq.Data[i, j, 0] = 0;
                    }
                }
            }
            //imageBox5.Image = imghiseq;
            imgOutput3 = imghiseq.Copy();

        }



        public class StringNum : IComparable<StringNum>
        {
            private List<string> _strings;
            private List<int> _numbers;
            public StringNum(string value)
            {
                _strings = new List<string>();
                _numbers = new List<int>();
                int pos = 0;
                bool number = false;
                while (pos < value.Length)
                {
                    int len = 0;
                    while (pos + len < value.Length && Char.IsDigit(value[pos + len]) == number)
                    {
                        len++;
                    }
                    if (number)
                    {
                        _numbers.Add(int.Parse(value.Substring(pos, len)));
                    }
                    else
                    {
                        _strings.Add(value.Substring(pos, len));
                    }
                    pos += len;
                    number = !number;
                }
            }
            public int CompareTo(StringNum other)
            {
                int index = 0;
                while (index < _strings.Count && index < other._strings.Count)
                {
                    int result = _strings[index].CompareTo(other._strings[index]);
                    if (result != 0) return result;
                    if (index < _numbers.Count && index < other._numbers.Count)
                    {
                        result = _numbers[index].CompareTo(other._numbers[index]);
                        if (result != 0) return result;
                    }
                    else
                    {
                        return index == _numbers.Count && index == other._numbers.Count ? 0 : index == _numbers.Count ? -1 : 1;
                    }
                    index++;
                }
                return index == _strings.Count && index == other._strings.Count ? 0 : index == _strings.Count ? -1 : 1;
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            //if (checkedListBox1.CheckedItems.Count < 1)
            {
                // MessageBox.Show("選擇區域");
                // return;
            }
            //if (imgOutput3 == null)
            {
                // return;
            }
            //checkBox2.Checked = false;
            //if (imgOutput3 == null)
            {
                //return;
            }

            Image<Gray, byte> imgray = imgOutput3.ThresholdBinary(new Gray(253), new Gray(255));
            //Image<Gray, byte> imgray2 = imgOutput3.ThresholdBinary(new Gray((int)numericUpDown9.Value), new Gray(255));
            VectorOfVectorOfPoint con = new VectorOfVectorOfPoint();
            double areal = 0;




            int area = 0;
            for (int i = 0; i < imgray.Rows; i++)
            {
                for (int j = 0; j < imgray.Cols; j++)
                {
                    if (imgray.Data[i, j, 0] == 255)
                        area++;

                }
            }
            //int area1 = 0;
            //for (int i = 0; i < imgray2.Rows; i++)
            {
                //for (int j = 0; j < imgray2.Cols; j++)
                {
                    //if (imgray2.Data[i, j, 0] == 255)
                    // area1++;

                }
            }


            areal = area * 0.2304; //MRI 0.1936 CT 0.2304
            //AVe = Convert.ToDouble(area1) * 0.2304;       //MRI 0.1936 CT 0.2304

            //textBox12.Text = areal.ToString("f4");
            //listBox6.Items.Add(AVe.ToString("f4"));
            /*if (checkedListBox1.GetItemChecked(0))
            {
                GM_M.Add(areal);


                //listBox17.Items.Add(areal.ToString("f4"));

            }
            if (checkedListBox1.GetItemChecked(1))
            {
                WM_M.Add(areal);


                //listBox17.Items.Add(areal.ToString("f4"));

            }
            if (checkedListBox1.GetItemChecked(2))
            {
                CD_N_PU_TR.Add(areal);


                //listBox17.Items.Add(areal.ToString("f4"));

            }



            //VPe.Add(AVe);*/





        }

        private void button17_Click(object sender, EventArgs e)
        {
            /*
             1. Edge detection(sobel)
             2. Dilation (10,1)
             3.FindContours
             4.Gemetrical Constrints
             */
            //sobel
            if (imgOutput2 == null)
            {
                return;
            }
            Image<Gray, byte> grayimg = new Image<Gray, byte>(imgOutput2.Bitmap).Resize((1.0), Emgu.CV.CvEnum.Inter.Linear).Dilate(1).PyrUp().PyrDown();
            grayimg = grayimg.ThresholdBinary(new Gray((int)(numericUpDown5.Value)), new Gray(255)); //閥值90適用帕金森氏症基底核分割//依據圈選範圍調整單一侵蝕與膨脹 .Dilate(2)   //90適用帕金森氏症 
            grayimg.SmoothGaussian(13);
            //Mat her = CvInvoke.GetStructuringElement(Emgu.CV.CvEnum.ElementShape.Rectangle, new Size(10, 1), new Point(-1, -1));
            //imgGray = imgGray.MorphologyEx(Emgu.CV.CvEnum.MorphOp.Gradient, her, new Point(-1, -1), 1, Emgu.CV.CvEnum.BorderType.Reflect, new MCvScalar(255));
            VectorOfVectorOfPoint cn = new VectorOfVectorOfPoint();
            VectorOfVectorOfPoint cY = new VectorOfVectorOfPoint();
            //HY1 = 0;
            //HY2 = 0;
            //num = 0;

            Image<Gray, byte> hier = new Image<Gray, byte>(grayimg.Width, grayimg.Height);
            CvInvoke.FindContours(grayimg, cn, hier, Emgu.CV.CvEnum.RetrType.External, Emgu.CV.CvEnum.ChainApproxMethod.ChainApproxSimple);
            CvInvoke.FindContours(grayimg, cY, hier, Emgu.CV.CvEnum.RetrType.List, Emgu.CV.CvEnum.ChainApproxMethod.ChainApproxSimple);

            CvInvoke.DrawContours(imgOutput2, cn, -1, new MCvScalar(0, 0, 190), 1);
            CvInvoke.DrawContours(imgOutput2, cY, -1, new MCvScalar(200, 0, 0), 1);

            Point[][] cons = cn.ToArrayOfArray();
            PointF[][] con2 = Array.ConvertAll<Point[], PointF[]>(cons, new Converter<Point[], PointF[]>(PointToPointF));

            for (int i = 0; i < cn.Size; i++)
            {
                PointF[] hull = CvInvoke.ConvexHull(con2[i], true);
                for (int j = 0; j < hull.Length; j++)
                {
                    Point p1 = new Point((int)(hull[j].X), (int)(hull[j].Y));
                    Point p2;
                    if (j == hull.Length - 1)
                        p2 = new Point((int)(hull[0].X), (int)(hull[0].Y));
                    else
                        p2 = new Point((int)(hull[j + 1].X), (int)(hull[j + 1].Y));
                    //CvInvoke.Circle(imgoutput3, p1, 3, new MCvScalar(0, 255, 255, 255), 0);
                    //CvInvoke.Line(imgoutput3, p1, p2, new MCvScalar(150), 2);


                    double perimeter = CvInvoke.ArcLength(cn[i], true);
                    VectorOfPoint cn2 = new VectorOfPoint();
                    CvInvoke.ApproxPolyDP(cn[i], cn2, 0.02 * perimeter, true);

                    for (int k = 0; k < cY.Size; k++)
                    {
                        double perimeter2 = CvInvoke.ArcLength(cY[k], true);
                        VectorOfPoint cb2 = new VectorOfPoint();
                        CvInvoke.ApproxPolyDP(cY[k], cb2, 0.02 * perimeter2, true);



                        //CvInvoke.FindContours(imgGray, cn, hier, Emgu.CV.CvEnum.RetrType.Ccomp, Emgu.CV.CvEnum.ChainApproxMethod.ChainApproxSimple);

                        //CvInvoke.DrawContours(imgOutput2, cY, k, new MCvScalar(200, 0, 0), 1);
                        CvInvoke.Line(imgOutput2, p1, p2, new MCvScalar(150), 2);
                        //HY1++;
                        //HY2++;
                        //num++;

                    }
                    //imageBox1.Image = imgOutput2;

                    imageBox3.Image = imgOutput2;


                }

            }

        }
        private void button32_Click(object sender, EventArgs e)
        {

        }
        public static PointF[] PointToPointF(Point[] pf)
        {
            PointF[] aaa = new PointF[pf.Length];
            int num = 0;
            foreach (var point in pf)
            {
                aaa[num].X = (int)point.X;
                aaa[num++].Y = (int)point.Y;
            }
            return aaa;
        }

        private void imageBox5_MouseDown(object sender, MouseEventArgs e)
        {
            point3 = e.Location;
            Invalidate();
        }

        /* private void imageBox5_MouseUp(object sender, MouseEventArgs e)

             if (checkBox2.Checked)
             {
                 if (imgOutput3 == null)
                     return;
                 //Define ROI. Valida altura e largura para evitar index range exception.
                 if (RealImageRect.Width > 0 && RealImageRect.Height > 0)
                 {
                     imgOutput3.ROI = RealImageRect;
                     imageBox6.Image = imgOutput3;

                 }
             }
         }*/

        /* private void imageBox5_MouseMove(object sender, MouseEventArgs e)
         {
             int X0, Y0;
             Utilities.ConvertCoordinates(imageBox5, out X0, out Y0, e.X, e.Y);
             label3.Text = "Last Position: X:" + X0 + "  Y:" + Y0;
             if (checkBox2.Checked)
             {
                 //Coordinates at input picture box
                 if (e.Button != MouseButtons.Left)
                     return;
                 Point tempEndPoint = e.Location;
                 Rect.Location = new Point(
                     Math.Min(point3.X, tempEndPoint.X),
                     Math.Min(point3.Y, tempEndPoint.Y));
                 Rect.Size = new Size(
                     Math.Abs(point3.X - tempEndPoint.X),
                     Math.Abs(point3.Y - tempEndPoint.Y));



                 //Coordinates at real image - Create ROI
                 Utilities.ConvertCoordinates(imageBox5, out X0, out Y0,
                 point3.X, point3.Y);
                 int X1, Y1;
                 Utilities.ConvertCoordinates(imageBox5, out X1, out Y1, tempEndPoint.X, tempEndPoint.Y);
                 RealImageRect.Location = new Point(
                     Math.Min(X0, X1),
                     Math.Min(Y0, Y1));
                 RealImageRect.Size = new Size(
                     Math.Abs(X0 - X1),
                     Math.Abs(Y0 - Y1));
                 ((PictureBox)sender).Invalidate();
             }
         }*/

        /*private void imageBox5_Paint(object sender, PaintEventArgs e)
        {
            if (checkBox2.Checked)
            {
                // Draw the rectangle...
                if (imageBox5.Image != null)
                {
                    if (Rect != null && Rect.Width > 0 && Rect.Height > 0)
                    {
                        //Seleciona a ROI
                        e.Graphics.SetClip(Rect, System.Drawing.Drawing2D.CombineMode.Exclude);
                        e.Graphics.FillRectangle(selectionBrush, new Rectangle
                (0, 0, ((PictureBox)sender).Width, ((PictureBox)sender).Height));
                        //e.Graphics.FillRectangle(selectionBrush, Rect);
                    }
                }
            }
        }*/

        private void button18_Click(object sender, EventArgs e)
        {
            imgOutput3 = imgOutput.Copy();
            //imageBox5.Image = imgOutput3;
        }

        private void button15_Click(object sender, EventArgs e)
        {

            //Image<Gray, byte> imgray = imgOutput3.ThresholdBinary(new Gray((int)numericUpDown6.Value), new Gray(255)); //閥值90適用帕金森氏症基底核分割
            //DepthType imgra = default(DepthType);
            //CvInvoke.Sobel(imgray, imgoutput, imgra, 1, 0);
            //Image<Bgr, byte> imgbgr = imgoutput.Convert<Bgr, byte>();
            VectorOfVectorOfPoint con = new VectorOfVectorOfPoint();
            VectorOfVectorOfPoint con2s = new VectorOfVectorOfPoint();
            //double area = new double();
            Image<Gray, byte> hier = new Image<Gray, byte>(imgOutput3.Width, imgOutput3.Height);
           // CvInvoke.FindContours(imgray, con, hier, RetrType.External, ChainApproxMethod.ChainApproxSimple);
           //CvInvoke.FindContours(imgray, con2s, hier, RetrType.Ccomp, ChainApproxMethod.ChainApproxSimple);
            CvInvoke.DrawContours(imgOutput3, con, -1, new MCvScalar(200, 0, 0), 1);
            CvInvoke.DrawContours(imgOutput3, con2s, -1, new MCvScalar(0, 0, 200), 1);
            Point[][] cons = con.ToArrayOfArray();
            PointF[][] con2 = Array.ConvertAll<Point[], PointF[]>(cons, new Converter<Point[], PointF[]>(PointToPointF));
            for (int i = 0; i < con.Size; i++)
            {
                PointF[] hull = CvInvoke.ConvexHull(con2[i], true);
                for (int j = 0; j < hull.Length; j++)
                {
                    Point p1 = new Point((int)(hull[j].X + 0.5), (int)(hull[j].Y + 0.5));
                    Point p2;
                    if (j == hull.Length - 1)
                        p2 = new Point((int)(hull[0].X + 0.5), (int)(hull[0].Y + 0.5));
                    else
                        p2 = new Point((int)(hull[j + 1].X + 0.5), (int)(hull[j + 1].Y + 0.5));
                    //CvInvoke.Circle(imgbgr, p1, 3, new MCvScalar(0, 255, 255, 255), 0);
                    //CvInvoke.Line(imgOutput3, p1, p2, new MCvScalar(150), 2);

                }
            }

            imgOutput3.Convert<Gray, byte>();
            //imageBox5.Image = imgOutput3;


        }

        //白質分割
        private void button21_Click(object sender, EventArgs e)
        {

        }

        //灰質分割
        private void button22_Click(object sender, EventArgs e)
        {
            Image<Gray, byte> image_copy = imgInput.Convert<Gray, byte>();
            Image<Gray, byte> mask = new Image<Gray, byte>(imgInput.Width + 2, imgInput.Height + 2);
            Image<Gray, byte> imgasd;
            Image<Gray, byte> imgssd;

            //Point point1 =new Point(256,320);

            Rectangle dummy = new System.Drawing.Rectangle(0, 0, 0, 0);
            CvInvoke.FloodFill(image_copy, mask, point1, new MCvScalar(1), out dummy,
                new MCvScalar(12),
                new MCvScalar(7),
                Emgu.CV.CvEnum.Connectivity.EightConnected,
                Emgu.CV.CvEnum.FloodFillType.Default);
            imgasd = (imgInput.Convert<Gray, byte>()) - image_copy;
            //imageBox7.Image = imgasd;
            //imageBox4.Image = imgasd;
            //imageBox5.Image = imgasd;
            //imageBox2.Image = image_copy;
            imgOutput5 = imgasd;
            imgssd = imgasd.ThresholdBinary(new Gray(50), new Gray(255));
            for (int i = 0; i < imgasd.Height; i++)
            {
                for (int j = 0; j < imgasd.Width; j++)
                {
                    if (imgasd.Data[i, j, 0] > 100)
                    {
                        imgasd.Data[i, j, 0] = 0;
                    }
                }
            }
        }

        //腦脊髓液分割
        private void button24_Click(object sender, EventArgs e)
        {

        }

        private void button25_Click(object sender, EventArgs e)
        {
            Image<Gray, byte> imgray = imgOutput6.ThresholdBinary(new Gray(25), new Gray(255));
            //Image<Gray, byte> imgray2 = imgOutput6.ThresholdBinary(new Gray((int)numericUpDown19.Value), new Gray(255));
            double AT = 0;
            double AV = 0;
            int area = 0;
            for (int i = 0; i < imgray.Rows; i++)
            {
                for (int j = 0; j < imgray.Cols; j++)
                {
                    if (imgray.Data[i, j, 0] == 255)
                        area++;

                }
            }
            int area1 = 0;
            // for (int i = 0; i < imgray2.Rows; i++)
            {
                // for (int j = 0; j < imgray2.Cols; j++)
                {
                    //if (imgray2.Data[i, j, 0] == 255)
                    area1++;

                }
            }

            AT = System.Convert.ToDouble(area) * 0.2304;
            AV = Convert.ToDouble(area1) * 0.2304;
            //textBox3.Text = AT.ToString("f4");
            //textBox14.Text = AV.ToString("f4");
            //listBox5.Items.Add(AT.ToString("f4"));

            VT1.Add(AT);
            VP1.Add(AV);


        }

        private void button26_Click(object sender, EventArgs e)
        {
            Image<Gray, byte> imgray = imgOutput5.ThresholdBinary(new Gray(25), new Gray(255));
            //Image<Gray, byte> imgray2 = imgOutput5.ThresholdBinary(new Gray((int)numericUpDown18.Value), new Gray(255));
            double AT = 0;
            double AV = 0;
            int area = 0;
            for (int i = 0; i < imgray.Rows; i++)
            {
                for (int j = 0; j < imgray.Cols; j++)
                {
                    if (imgray.Data[i, j, 0] == 255)
                        area++;

                }
            }
            int area1 = 0;
            //for (int i = 0; i < imgray2.Rows; i++)
            {
                // for (int j = 0; j < imgray2.Cols; j++)
                {
                    // if (imgray2.Data[i, j, 0] == 255)
                    area1++;

                }
            }

            AT = System.Convert.ToDouble(area) * 0.2304;
            AV = Convert.ToDouble(area1) * 0.2304;
            //textBox4.Text = AT.ToString("f4");
            //textBox4.Text = AV.ToString("f4");
            //listBox6.Items.Add(AT.ToString("f4"));
            VT2.Add(AT);
            //VP2.Add(AV);


        }

        /*private void button28_Click(object sender, EventArgs e)
        {

            VS = volume(VT, 5);
            double VSS = volume(VP, 5);
            VPP = VSS / VS * 100;
            double VPPS = 90 - VPP;
            textBox10.Text = VS.ToString("f4"); //全腦體積
            textBox6.Text = VSS.ToString("f4"); //腦組織體積
            textBox9.Text = VPP.ToString("f4") + "%"; //腦組織比例
            textBox7.Text = VPPS.ToString("f4") + "%";//腦萎縮比


        }*/


        private static double volume(List<double> LD, double H)

        {

            double SV = 0;
            double V = 0;
            if (LD.Count == 0)
            {
                return 0;
            }
            if (LD.Count == 1)
            {
                LD.Add(0);

            }
            if (LD.Count == 2)
            {
                LD.Add(0);
            }

            for (int i = 1; i <= LD.Count - 1; i++)
            {

                double A = LD[i] + LD[i - 1];
                double DA = Math.Sqrt(LD[i] * LD[i - 1]);
                V = (A + DA) * H / 3;
                /* double k = (H * LD[LD.Count - 1]) / (LD[LD.Count - 2] - LD[LD.Count - 1]);
                  vk = (LD[LD.Count - 1] * k) / 3;
                  */

                SV = SV + V;

            }

            return SV;
        }

        private void button29_Click(object sender, EventArgs e)
        {
            CD_N_PU_TR.Clear();
            CD_N_PU_TL.Clear();
            OC_B.Clear();
            STN.Clear();
            VT.Clear();
            VP.Clear();
            listBox1.Items.Clear();
            listBox4.Items.Clear();
            listBox12.Items.Clear();
            listBox15.Items.Clear();
            listBox16.Items.Clear();
            listBox20.Items.Clear();
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            //textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            //textBox8.Clear();
            textBox9.Clear();
            textBox10.Clear();
            //textBox11.Clear();
            //textBox12.Clear();
            textBox13.Clear();
            //textBox14.Clear();
            //textBox15.Clear();
            textBox16.Clear();
            //textBox17.Clear();
            //textBox18.Clear();
            //textBox19.Clear();
            textBox21.Clear();
            //textBox22.Clear();
            textBox25.Clear();
            textBox26.Clear();
            //textBox28.Clear();
            textBox35.Clear();
            textBox37.Clear();
            VTTR = 0;
            VTTL = 0;
            SUR_1 = 0;
            SUR_2 = 0;
            SUR_3 = 0;
            SUR_4 = 0;
            SUR1 = 0;
            SUR2 = 0;
            SUR3 = 0;
            test5 = 0;
            test6 = 0;
            test11 = 0;
            test12 = 0;
            VTTO = 0;
            VTTT = 0;
            VPTT = 0;
            VPTT1 = 0;
            VPTT2 = 0;
            VPT1 = 0;
            VPT2 = 0;
            VPT0 = 0;
            test1 = 0;
            test2 = 0;
            test3 = 0;
            test4 = 0;
            test7 = 0;
            test8 = 0;
            test9 = 0;
            test10 = 0;
            ASI1 = 0;
            VPP = 0;
            test66 = 0;
            VolCN = 0;
            VolPM = 0;
            VolOB = 0;
            VolSN = 0;
            VPPCN = 0;
            VPPPM = 0;
            VPPOB = 0;
            VPPSN = 0;


        }

        private void button24_Click_1(object sender, EventArgs e)
        {
            Image<Gray, byte> imgGray = imgOutput2.Copy();
            Image<Gray, byte> imgSucal;
            imgSucal = new Image<Gray, byte>(imgGray.Width, imgGray.Height, new Gray(0));
            CvInvoke.Threshold(imgGray, imgSucal, 253, 255, Emgu.CV.CvEnum.ThresholdType.Binary);
            imageBox3.Image = imgSucal;
            imgOutput2 = imgSucal.Copy();
        }

        private void button27_Click_1(object sender, EventArgs e)
        {
            Image<Gray, byte> imgGray = imgInput.Convert<Gray, byte>();
            imgOutput = imgGray.Convert<Gray, byte>();
            imageBox2.Image = imgOutput;
        }

        private void button30_Click(object sender, EventArgs e)
        {
            if (checkedListBox2.CheckedItems.Count < 1)
            {
                MessageBox.Show("選擇腦室");
                return;
            }
            if (imgOutput2 == null)
            {
                return;
            }

            Image<Gray, byte> imgray = imgOutput2.ThresholdBinary(new Gray(253), new Gray(255));
            double area1 = new double();
            /*VectorOfVectorOfPoint con = new VectorOfVectorOfPoint();
            double area = new double();
            double area1 = new double();
            Image<Gray, byte> hier = new Image<Gray, byte>(imgoutput2.Width, imgoutput2.Height);
            CvInvoke.FindContours(imgray, con, hier, RetrType.External, ChainApproxMethod.ChainApproxSimple);
            //CvInvoke.DrawContours(imgoutput2, con, -1, new MCvScalar(2));
            //imageBox3.Image = imgoutput2;
            int count = con.Size;
            for (int i = 0; i < count; i++)
            {
                using (VectorOfPoint contour = con[i])
                using (VectorOfPoint approxContour = new VectorOfPoint())
                {
                    area = CvInvoke.ContourArea(contour, false);
                }
            }*/
            int area = 0;
            for (int i = 0; i < imgOutput2.Rows; i++)
            {
                for (int j = 0; j < imgOutput2.Cols; j++)
                {
                    if (imgray.Data[i, j, 0] == 255)
                        area++;

                }
            }

            area1 = area * 0.1936;
            textBox1.Text = area1.ToString("f4");


            if (checkedListBox2.GetItemChecked(0))
            {
                CD_N_PU_TR.Add(area1);
                
                listBox12.Items.Add(area1.ToString("f4"));

            }
            if (checkedListBox2.GetItemChecked(1))
            {
                CD_N_PU_TL.Add(area1);
                
                listBox4.Items.Add(area1.ToString("f4"));

            }
            if (checkedListBox2.GetItemChecked(2))
            {
                OC_B.Add(area1);
                listBox20.Items.Add(area1.ToString("f4"));
            }
            
            if (checkedListBox2.GetItemChecked(3))
            {
                STN.Add(area1);
                listBox1.Items.Add(area1.ToString("f4"));
            }
            

            checkBox1.Checked = false;
            for (int i = 0; i < checkedListBox2.Items.Count; i++)

            {

                checkedListBox2.SetItemChecked(i, false);

            }


            //if (checkedListBox1.(0,true))
        }

        private void button31_Click(object sender, EventArgs e)
        {
            Image<Gray, byte> imgray = imgOutput.ThresholdBinary(new Gray(25), new Gray(255));
            Image<Gray, byte> imgray2 = imgOutput.ThresholdBinary(new Gray((int)numericUpDown7.Value), new Gray(255));
            double AT = 0;
            double AV = 0;
            int area = 0;
            for (int i = 0; i < imgray.Rows; i++)
            {
                for (int j = 0; j < imgray.Cols; j++)
                {
                    if (imgray.Data[i, j, 0] == 255)
                        area++;

                }
            }
            int area1 = 0;
            for (int i = 0; i < imgray2.Rows; i++)
            {
                for (int j = 0; j < imgray2.Cols; j++)
                {
                    if (imgray2.Data[i, j, 0] == 255)
                        area1++;

                }
            }
            AT = System.Convert.ToDouble(area) * 0.1936; //MRI 0.1936 CT 0.2304
            AV = Convert.ToDouble(area1) * 0.1936;       //MRI 0.1936 CT 0.2304
            listBox15.Items.Add(AT.ToString("f4"));
            listBox16.Items.Add(AV.ToString("f4"));
            VT.Add(AT);
            VP.Add(AV);
            //VT.Add(AT);
            //VP.Add(AV);
        }


        //GLCM分期測試1
        private void button33_Click(object sender, EventArgs e)
        {

        }

        //GLCM分期測試2
        private void button34_Click(object sender, EventArgs e)
        {
            double[] CM;
            double[] ET;
            double[] CO;
            double[] CON;
            double[] CS;
            double[] EV;
            double[] EN;
            double[] CU;
            double CC, EE, HH, HQ;
            Utilities.getP(imgOutput3.ToBitmap(), (imgOutput3.Height) / 2, (imgOutput3.Height) / 2, (imgOutput3.Height) / 2);
            Utilities.getEntropy(imgOutput3.Convert<Bgr, byte>().ToBitmap(), (imgOutput3.Height) / 2, ((imgOutput3.Height) / 2), ((imgOutput3.Height) / 2), out CM, out ET, out CO, out CON, out CS, out EV, out EN, out CU);
            CC = CS[0];
            EE = EN[0];
            HH = ET[0];
            HQ = EV[0];
            //textBox180.Text = JJ.ToString();
            if ((CC * (-0.2973) + EE * (-0.6508) + HH * (0.6946) + HQ * 0.0747 + 0) > -0.6220)
            {
                //textBox18.Text = "2";
            }
            else
            {
                //textBox18.Text = "3";
            }
        }


        private void button13_Click_1(object sender, EventArgs e)
        {
            CD_N_PU_TR.Clear();
            CD_N_PU_TL.Clear();
            OC_B.Clear();
            STN.Clear();
            VT.Clear();
            VP.Clear();
            listBox1.Items.Clear();
            listBox4.Items.Clear();
            listBox12.Items.Clear();
            listBox15.Items.Clear();
            listBox16.Items.Clear();
            listBox20.Items.Clear();
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            //textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            //textBox8.Clear();
            textBox9.Clear();
            textBox10.Clear();
            //textBox11.Clear();
            //textBox12.Clear();
            textBox13.Clear();
            //textBox14.Clear();
            //textBox15.Clear();
            textBox16.Clear();
            //textBox17.Clear();
            //textBox18.Clear();
            //textBox19.Clear();
            textBox21.Clear();
            //textBox22.Clear();
            textBox25.Clear();
            textBox26.Clear();
            //textBox28.Clear();
            textBox35.Clear();
            textBox37.Clear();
            VTTR = 0;
            VTTL = 0;
            SUR_1 = 0;
            SUR_2 = 0;
            SUR_3 = 0;
            SUR_4 = 0;
            SUR1 = 0;
            SUR2 = 0;
            SUR3 = 0;
            test5 = 0;
            test6 = 0;
            test11 = 0;
            test12 = 0;
            VTTO = 0;
            VTTT = 0;
            VPTT = 0;
            VPTT1 = 0;
            VPTT2 = 0;
            VPT1 = 0;
            VPT2 = 0;
            VPT0 = 0;
            test1 = 0;
            test2 = 0;
            test3 = 0;
            test4 = 0;
            test7 = 0;
            test8 = 0;
            test9 = 0;
            test10 = 0;
            ASI1 = 0;
            VPP = 0;
            test66 = 0;
            VolCN = 0;
            VolPM = 0;
            VolOB = 0;
            VolSN = 0;
            VPPCN = 0;
            VPPPM = 0;
            VPPOB = 0;
            VPPSN = 0;
        }

        

        private void button25_Click_1(object sender, EventArgs e)
        {

        }

        private void button25_Click_2(object sender, EventArgs e)
        {

        }

        private void button26_Click_1(object sender, EventArgs e)
        {

        }

        private void button23_Click_2(object sender, EventArgs e)
        {
            Image<Gray, byte> imgray = imgOutput.ThresholdBinary(new Gray(25), new Gray(255));
            Image<Gray, byte> imgray2 = imgOutput.ThresholdBinary(new Gray((int)numericUpDown7.Value), new Gray(255));

            double AT = 0;
            double AV = 0;
            int area = 0;
            for (int i = 0; i < imgray.Rows; i++)
            {
                for (int j = 0; j < imgray.Cols; j++)
                {
                    if (imgray.Data[i, j, 0] == 255)
                        area++;

                }
            }
            int area1 = 0;
            for (int i = 0; i < imgray2.Rows; i++)
            {
                for (int j = 0; j < imgray2.Cols; j++)
                {
                    if (imgray2.Data[i, j, 0] == 255)
                        area1++;

                }
            }
            
           
            AT = System.Convert.ToDouble(area) * 0.1521; //MRI 0.1936 CT 0.2304
            AV = Convert.ToDouble(area1) * 0.1521;       //MRI 0.1936 CT 0.2304
            listBox15.Items.Add(AT.ToString("f4"));
            listBox16.Items.Add(AV.ToString("f4"));
            VT.Add(AT);
            VP.Add(AV);
            //VT.Add(AT);
            //VP.Add(AV);
        }

        private void button26_Click_2(object sender, EventArgs e)
        {
            if (checkedListBox2.CheckedItems.Count < 1)
            {
                MessageBox.Show("選擇腦室");
                return;
            }
            if (imgOutput2 == null)
            {
                return;
            }

            Image<Gray, byte> imgray = imgOutput2.ThresholdBinary(new Gray(253), new Gray(255));
            double area1 = new double();
            /*VectorOfVectorOfPoint con = new VectorOfVectorOfPoint();
            double area = new double();
            double area1 = new double();
            Image<Gray, byte> hier = new Image<Gray, byte>(imgoutput2.Width, imgoutput2.Height);
            CvInvoke.FindContours(imgray, con, hier, RetrType.External, ChainApproxMethod.ChainApproxSimple);
            //CvInvoke.DrawContours(imgoutput2, con, -1, new MCvScalar(2));
            //imageBox3.Image = imgoutput2;
            int count = con.Size;
            for (int i = 0; i < count; i++)
            {
                using (VectorOfPoint contour = con[i])
                using (VectorOfPoint approxContour = new VectorOfPoint())
                {
                    area = CvInvoke.ContourArea(contour, false);
                }
            }*/
            int area = 0;
            for (int i = 0; i < imgOutput2.Rows; i++)
            {
                for (int j = 0; j < imgOutput2.Cols; j++)
                {
                    if (imgray.Data[i, j, 0] == 255)
                        area++;

                }
            }

            area1 = area * 0.1521;
            textBox1.Text = area1.ToString("f4");


            if (checkedListBox2.GetItemChecked(0))
            {
                CD_N_PU_TR.Add(area1);
                
                listBox12.Items.Add(area1.ToString("f4"));

            }
            if (checkedListBox2.GetItemChecked(1))
            {
                CD_N_PU_TL.Add(area1);
                
                listBox4.Items.Add(area1.ToString("f4"));

            }
            if (checkedListBox2.GetItemChecked(2))
            {
                OC_B.Add(area1);
                listBox20.Items.Add(area1.ToString("f4"));
            }
            
            if (checkedListBox2.GetItemChecked(3))
            {
                STN.Add(area1);
                listBox1.Items.Add(area1.ToString("f4"));
            }
            
        }

        private void button25_Click_3(object sender, EventArgs e)
        {
            //Image<Gray, byte> imgGray;
            Image<Gray, byte> imgBinarize;
            Image<Gray, byte> imgasd;
            Image<Gray, byte> image_Copy = imgInput.Convert<Gray, byte>();

            imgBinarize = new Image<Gray, byte>(imgInput.Width+2, imgInput.Height+2, new Gray(0));
            CvInvoke.Threshold(image_Copy, imgBinarize, 250, 255, Emgu.CV.CvEnum.ThresholdType.Binary);
            /*Image<Gray, byte>*/
            imgasd = imgInput.Convert<Gray,byte>() - image_Copy;
            imageBox2.Image = imgOutput;
            imgOutput = imgasd;
        }

        private void button27_Click(object sender, EventArgs e)
        {
          
            Image<Gray, byte> grayimg = imgOutput3.Copy();
            Image<Gray, byte> imgBinarize;
            imgBinarize = new Image<Gray, byte>(grayimg.Width, grayimg.Height,new Gray(0));
            

           // CvInvoke.AdaptiveThreshold(grayimg, imgBinarize, 255, AdaptiveThresholdType.GaussianC, ThresholdType.BinaryInv, (int)numericUpDown15.Value, 3); //閥值90適用帕金森氏症基底核分割


            VectorOfVectorOfPoint cn = new VectorOfVectorOfPoint();
            //HY1 = 0;
            //HY2 = 0;
            //num1 = 0;

            CvInvoke.FindContours(grayimg, cn, null, RetrType.External, ChainApproxMethod.ChainApproxNone);
            CvInvoke.FindContours(grayimg, cn, null, RetrType.Ccomp, ChainApproxMethod.ChainApproxNone);


            int count = cn.Size;
            for (int i = 0; i < count; i++)
            {
                using (VectorOfPoint contour = cn[i])
                using (VectorOfPoint approxContour = new VectorOfPoint())
                {
                    // 原始輪廓線
                    CvInvoke.DrawContours(imgOutput3, cn, i, new MCvScalar(170), 1);
                    CvInvoke.DrawContours(imgOutput3, cn, i, new MCvScalar(200), 1);

                    // 近似後輪廓線
                    CvInvoke.ApproxPolyDP(contour, approxContour, CvInvoke.ArcLength(contour, true) * 0.03, true);
                    Point[][] cons = cn.ToArrayOfArray();
                    PointF[][] con2 = Array.ConvertAll<Point[], PointF[]>(cons, new Converter<Point[], PointF[]>(PointToPointF));


                    {
                        PointF[] hull = CvInvoke.ConvexHull(con2[i], true);
                        for (int j = 0; j < hull.Length; j++)
                        {
                            Point p1 = new Point((int)(hull[j].X + 0.5), (int)(hull[j].Y + 0.5));
                            Point p2;
                            if (j == hull.Length - 1)
                                p2 = new Point((int)(hull[0].X + 0.5), (int)(hull[0].Y + 0.5));
                            else
                                p2 = new Point((int)(hull[j + 1].X + 0.5), (int)(hull[j + 1].Y + 0.5));




                        }
                    }
                    //imgOutput2._Not();
                    //imageBox5.Image = imgOutput3;


                    //imageBox8.Image = imgasd;
                }
            }
        }

        private void button27_Click_2(object sender, EventArgs e)
        {
        }

        private void numericUpDown15_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button28_Click(object sender, EventArgs e)
        {
            VS = volume(VT, 5);
            double VSS = volume(VP, 5);
            VPP = VSS / VS * 100;
            double VPPS = 90 - VPP;
            textBox10.Text = VS.ToString("f4");
            textBox6.Text = VSS.ToString("f4");
            textBox9.Text = VPP.ToString("f4") + "%";
            textBox7.Text = VPPS.ToString("f4") + "%";
        }

        private void button23_Click_1(object sender, EventArgs e)
        {

            double[] CM;
            double[] ET;
            double[] CO;
            double[] CON;
            double[] CS;
            double[] EV;
            double[] EN;
            double[] CU;
            //double CC, EE, HH, HQ;
            Utilities.getP(imgOutput2.ToBitmap(), (imgOutput2.Height) / 2, (imgOutput2.Height) / 2, (imgOutput2.Height) / 2);
            Utilities.getEntropy(imgOutput2.Convert<Bgr, byte>().ToBitmap(), (imgOutput2.Height) / 2, ((imgOutput2.Height) / 2), ((imgOutput2.Height) / 2), out CM, out ET, out CO, out CON, out CS, out EV, out EN, out CU);
            test1 = CS[0];
            test2 = EN[0];
            test9 = ET[0];
            test10 = EV[0];
            test7 = CM[0];

            //textBox180.Text = JJ.ToString();
            if ((test1 * (0.0008) + test2 * (-0.0016) + test9 * (-0.5851) + test10 * -0.5685 + test7 * (-0.0055) + 0) > -0.5783)
            {
                textBox2.Text = "1";
            }
            else
            {
                textBox2.Text = "0";
            }
            if ((test1 * (0.0007) + test2 * (-0.0003) + test9 * (-0.5779) + test10 * -0.5775 + test7 * (0.0034) + 0) > -0.5766)
            {
                textBox2.Text = "2";
            }
            else
            {
                textBox2.Text = "3";
            }
        }

        private void button22_Click_2(object sender, EventArgs e)
        {
            char[] Mychar = { '%' };

            test1 = (double.Parse(textBox21.Text.TrimEnd(Mychar))) ; //R區域吸收率
            test2 = (double.Parse(textBox5.Text.TrimEnd(Mychar))) ;  //L區域吸收率
            test9 = (double.Parse(textBox35.Text.TrimEnd(Mychar))) ; //R面積比
            test10 = (double.Parse(textBox16.Text.TrimEnd(Mychar))) ; //L面積比
            //test7 = (double.Parse(textBox38.Text.TrimEnd(Mychar))) ; //R+L面積比
            ASI1 = (double.Parse(textBox13.Text.TrimEnd(Mychar))) ; //不對稱性指標
            //VTTR  = (double.Parse(listBox12.Text.TrimEnd(Mychar))) / 100; //R區域總和
            //VTTL  = (double.Parse(listBox4.Text.TrimEnd(Mychar))) / 100;  //L區域總和
            //test4 = (double.Parse(listBox20.Text.TrimEnd(Mychar))) / 100; //OC區域總和
            test3 = (double.Parse(textBox26.Text.TrimEnd(Mychar))) ; //R+L區域總吸收率
            test8 = (double.Parse(textBox37.Text.TrimEnd(Mychar))); //OC面積比*/

            //double sum = (test1) + (test2) ;

            if (test1 > 0.45 || test1 == 0.58 && test2 > 0.42 && test2 == 0.66)
            {
                textBox2.Text = "1.0";
            }
            else if (test1 > 0.3 || test1 == 0.55 && test2 > 0.3 || test2 == 0.5)
            {
                textBox2.Text = "2.0";
            }
            else if (test1 > 0.2 || test1 > 0.1 || test1 == 0.298 && test2 > 0.2 || test2 < 0.1 || test2 == 0.315)

            {
                textBox2.Text = "3.0";
            }



            {

      {
                                                                        
                    
                    
                    //double n1 = (test9 * 0.1716) + (test10 * 0.1608) + (test7 * 0.3324) + (ASI1 * 12.6425);



                    //textBox2.Text = n.ToString();

                    //textBox2.Text = n1.ToString();

                    /* if ((test1 * (-0.1162) + test2 * (0.0535) + test9 * (0.7170) + test10 * (-0.4560) > -0.5115))
                     {

                         {

                             textBox2.Text = "1";
                         }
                         if ((test1 * (0.3010) + test2 * (-0.2096) + test9 * (-0.4690) > (-0.0016) ))
                         {

                             textBox2.Text = "2or3";
                         }

                     else
                     {
                             textBox2.Text = "0";
                         }
                     }*/


                    /*else
                    if ((test1 * (-0.0862) + test2 * (0.1383) + test9 * (0.6280) + test10 * (-0.6843) + test8 * 0.3269 + test7 * 0.0747) > -0.0629)
                    {
                        //textBox18.Text = "2";
                    }
                    else
                    {
                        //textBox18.Text = "3";
                    }*/


                    /*if ( < 0.449)
                     {
                         textBox2.Text = "1"; //判斷正確
                     }
                    else
                     {
                         textBox2.Text = "0"; //判斷錯誤
                     }*/





                    //if (n > -0.36010)
                    //textBox2.Text = "1";

                    // if (n > -0.18400)
                    //textBox2.Text = "1";

                }
            }
        }
  




       private void button20_Click(object sender, EventArgs e)
       {
           { 
               for (int z = 0; z < CD_N_PU_TR.Count;z++)
               {
                   SUR_1 = Convert.ToDouble(CD_N_PU_TR[z]); //SUR_1 += 等於讓這變數持續 + -x/
               }
               {

               }
               for (int z = 0; z < CD_N_PU_TL.Count; z++)
               {
                   SUR_2 = Convert.ToDouble(CD_N_PU_TL[z]); //SUR_2+= 等於讓這變數持續+-x/ 
               }
               {

               }
               for (int z = 0; z < OC_B.Count; z++)
               {
                   SUR_3 = Convert.ToDouble(OC_B[z]);      //SUR_3+= 等於讓這變數持續+-x/ 
               }


               test1 = ((SUR_1 - SUR_3) / SUR_3); //R區域攝取比值(對稱)
               test2 = ((SUR_2 - SUR_3) / SUR_3); //L區域攝取比值(同稱)  
               VTTR  = SUR_1;                      //R區域面積(對稱同稱區域)
               VTTL  = SUR_2;                      //L區域面積(同側區域)
               test4 = SUR_3;                      //OC區域面積(總合對稱同稱區域)
               test3 = VTTR + VTTL;                //R+L區域總面積(對稱+同稱)
               ASI1 = (2*(test1 - test2)  / (test1 + test2));

               for (int z = 0; z < VTS.Count;z++)
               {
                   VS= Convert.ToDouble(VTS[z]); //全腦面積   VTS+= 等於讓這變數持續+-x/ 
               }


                //test5 = test1 / VS; //SUR1面積比
                //test6 = test2 / VS; //SUR2面積比
               
               test11 = ((test3 - test4) / test4);
               //test7  = test3 / VS * 100;  //SUR紋狀體面積比
               test9  = VTTR  / VS * 100;   //R區域面積比
               test10 = VTTL  / VS * 100;  //L區域面積比
               test8  = test4 / VS * 100;  //SUR0面積比

              // HY1 = test1;
              // HY2 = test2;

               //VPTT1 = 90 - VPT1;
               //VPTT2 = 90 - VPT2;
               //VPTT  = 90 - VPT0;

               textBox21.Text = test1.ToString("f4"); //R區域面積
               textBox5.Text  = test2.ToString("f4"); //L區域面積
               textBox26.Text = test3.ToString("f4"); //R+L區域總面積
               textBox25.Text = test4.ToString("f4"); //OC區域面積
               textBox3.Text = test11.ToString("f4"); //Striatum攝取比值


               textBox35.Text = (test9*100).ToString("f4")+("%"); //R區域面積比
               textBox16.Text = (test10*100).ToString("f4")+("%"); //L區域面積比
               textBox37.Text = (test8*100).ToString("f4")+("%"); //OC區域面積比
               //textBox38.Text = (test7*100).ToString("f4")+("%"); //R+L區域紋狀體比

               textBox13.Text = (ASI1*100).ToString("f4") + ("%"); //不稱性指標比

               /*textBox4.Text  = VPT1.ToString("f4")+("%"); //R區域體積比
               textBox8.Text  = VPT2.ToString("f4")+("%"); //L區域體積比
               textBox18.Text = VPT0.ToString("f4")+("%"); //OC區域體積比
               textBox13.Text = ASI1.ToString("f4")+("%"); //不稱性指標比*/


            /*textBox14.Text = VPTT1.ToString("f4") + ("%");
              textBox15.Text = VPTT2.ToString("f4") + ("%");
              textBox27.Text = VPTT.ToString("f4") + ("%"); */

            

        }
    }

        /*private void button36_Click(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                return;
            }
            Image<Gray, byte> image_copy2 = imgOutput3.Copy();
            Image<Gray, byte> mask = new Image<Gray, byte>(imgInput.Width + 2, imgInput.Height + 2);
            Image<Gray, byte> imgasd;
            Rectangle dummy = new System.Drawing.Rectangle(1, 1, 1, 1);
            CvInvoke.FloodFill(image_copy2, mask, point3, new MCvScalar(254), out dummy,
                //new MCvScalar((double)numericUpDown11.Value),
                //new MCvScalar((double)numericUpDown12.Value),
                Emgu.CV.CvEnum.Connectivity.FourConnected,
                Emgu.CV.CvEnum.FloodFillType.FixedRange);
            imgasd =  image_copy2;
            //imageBox5.Image = imgasd;
            imgOutput3 = imgasd.Copy();
        }*/
        class WatershedSegmenter
        {
            private Image<Gray, Int32> _Markers;

            public void SetMakers(Image<Gray, byte> markers)
            {
                this._Markers = markers.Convert<Gray, Int32>();

            }

            public Image<Gray, Int32> Process(Image<Gray, byte> image)
            {
                CvInvoke.Watershed(image, this._Markers);
                return this._Markers;
            }
            public Image<Gray, Byte> GetWatersheds()
            {
                Image<Gray, Byte> watersheds = this._Markers.Convert<Gray, Byte>();
                watersheds._ThresholdBinary(new Gray(1), new Gray(255));
                return watersheds;
            }


        }

        private void button37_Click(object sender, EventArgs e)
        {
            //Read input image
            Image<Gray, byte> image = imgOutput3.Copy();

            //Get the binary image
            //Image<Gray, byte> binary =image.Convert<Gray, byte>().ThresholdBinary(new Gray((int)numericUpDown13.Value), new Gray(255));   //閥值90適用帕金森氏症基底核分割
            var closeElement = CvInvoke.GetStructuringElement(ElementShape.Rectangle,new Size(5,5),new Point(-1,-1));
            //binary = binary.MorphologyEx(Emgu.CV.CvEnum.MorphOp.Close,closeElement, 1, BorderType.D, new MCvScalar());
            //CvInvoke.MorphologyEx(binary, binary, MorphOp.Open, closeElement, new Point(-1, -1), 1, BorderType.Constant, new MCvScalar());

            //Eliminate noise and smaller objects
            //Image<Gray, byte> foreground = binary.Erode(1);

            //Identify image pixels without objects
            //Image<Gray, byte> background = binary.Dilate(1);
            //background._ThresholdBinary(new Gray(1), new Gray(128));
            //Create markers image
            //Image<Gray, byte> markers = background - foreground;

            //Create watershed segmentation object
            WatershedSegmenter watershedSegmenter = new WatershedSegmenter();
            //Set markers and process
            //watershedSegmenter.SetMakers(markers);
          
            //Image<Gray, Int32> boundaryImage = watershedSegmenter.Process(markers);

            //imageBox5.Image = markers;
            //imgOutput3 = markers.Copy();
        }

        private void button38_Click(object sender, EventArgs e)
        {

        }

        private void button39_Click(object sender, EventArgs e)
        {

        }

        private void button38_Click_1(object sender, EventArgs e)
        {

            //Image<Gray, byte> imgray = imgOutput2.ThresholdBinary(new Gray((int)numericUpDown8.Value), new Gray(255));
            //Image<Bgr, byte> img = new Image<Bgr, byte>(imgray.Bitmap);
            using (VectorOfVectorOfPoint contours = new VectorOfVectorOfPoint())
            {
               // CvInvoke.FindContours(imgray, contours, null, RetrType.Ccomp, ChainApproxMethod.ChainApproxSimple);

                int count = contours.Size;
                for (int i =1; i<count; i++)
                {
                    using (VectorOfPoint contour = contours[i])
                    using (VectorOfPoint approxContour = new VectorOfPoint())
                    {
                        CvInvoke.DrawContours(imgOutput2, contours, i, new MCvScalar(170, 0, 0), 1);

                        CvInvoke.ApproxPolyDP(contour, approxContour, CvInvoke.ArcLength(contour, true) * 0.05, true);
                        //Point[] pts = approxContour.ToArray();
                        PointF[] temp = Array.ConvertAll(contour.ToArray(), new Converter<Point, PointF>(Point2PointF));
                        PointF[] pts = CvInvoke.ConvexHull(temp, true);
                        for (int j = 0; j < pts.Length; j++)
                        {
                            Point p1 = new Point((int)pts[j].X, (int)pts[j].Y);
                            Point p2;

                            if (j == pts.Length - 1)
                                p2 = new Point((int)pts[0].X, (int)pts[0].Y);
                            else
                                p2 = new Point((int)pts[j + 1].X, (int)pts[j + 1].Y);

                            CvInvoke.Line(imgOutput2, p1, p2, new MCvScalar(255, 0, 255, 255), 3);
                        }
                        imageBox3.Image = imgOutput2;
                        //imgOutput2 = img.Convert<Gray,byte>().Copy();
                    }
                }
            }
        }
        private static PointF Point2PointF(Point p)
        {
            PointF PF = new PointF
            {
                X = p.X,
                Y = p.Y
            };
            return PF;
        }

        private void button39_Click_1(object sender, EventArgs e)
        {
            
            //Read input image
            Image<Bgr, byte> image = new Image<Bgr, byte>(imgInput.Bitmap);
           // Image<Gray, byte> img = new Image<Gray, byte>(imgInput.Width, imgInput.Height, new Gray(0));
            //Image<Gray, byte> imgasd;
            Point topLeftPoint = new Point(image.Width / 3, image.Height / 3);
            int wid = Math.Min(image.Width, image.Height) / 3;
            Rectangle rect = new Rectangle(topLeftPoint, new Size(wid, wid));
            //define bounding rectangle
            //the pixels outside this rectangle
            //will be labeled as background
            //Rectangle rect = new Rectangle(1, 1, 1, 1);
            //GrabCut segmentation
            Image<Gray, byte> mask = image.GrabCut(rect, 5);
            mask = mask.And(new Gray(3));
            Image<Gray, byte> result = imgOutput2.Copy(mask);
            
            //Generate output image
           // imgOutput2 = imgasd;
            
            imageBox3.Image = result;
        }

        private void button40_Click(object sender, EventArgs e)
        {
            Image<Gray, byte> img = imgInput.Convert<Gray,byte>();
            Image<Gray, byte> mask = new Image<Gray, byte>(img.Width + 2, img.Height + 2);
            //Image<Gray, byte> imgasd;
            Point center = new Point(img.Width / 2, img.Height / 2);
            Rectangle dummy = new System.Drawing.Rectangle(1, 1, 1, 1);
            CvInvoke.FloodFill(img, mask, center, new MCvScalar(254), out dummy,
                new MCvScalar(12),
                new MCvScalar(7),
                Emgu.CV.CvEnum.Connectivity.EightConnected,
                Emgu.CV.CvEnum.FloodFillType.MaskOnly);
            Image<Gray, byte> reducedMask = mask.Copy(new Rectangle(1, 1, img.Width, img.Height));
            //imgasd = imgInput.Convert<Gray, byte>() - img;
            Image<Gray, byte> morphImg = new Image<Gray, byte>(img.Size);
            var closeElement = CvInvoke.GetStructuringElement(ElementShape.Cross, new Size(3, 3), new Point(1, 1));
            CvInvoke.MorphologyEx(reducedMask, morphImg, MorphOp.Close, closeElement, new Point(-1, -1), 1, BorderType.Default, new MCvScalar());
            //SEGMENTING FOREGROUND OBJECT USING MASK
            Image<Gray, byte> imageROI = imgInput.Convert<Gray, byte>().Copy(morphImg);
            //imgasd = imgInput.Convert<Gray, byte>() - img;
            imgOutput3 = imageROI.Copy();
            //imageBox5.Image = imgOutput3.SmoothMedian(5);
            

        }

        private void button41_Click(object sender, EventArgs e)
        {
            Image<Gray, byte> img = imgInput.Convert<Gray, byte>();
            Size sz = img.Size;
           // Mat img2 = new Mat(sz);
            Image<Gray, byte> mask = new Image<Gray, byte>(img.Width + 2, img.Height + 2);
            
            Image<Gray, Byte> temp1 = img.Copy(mask);
            //imageBox5.Image = imgOutput3;
            imgOutput3 = mask.Copy();

               
            }

        private void button42_Click(object sender, EventArgs e)
        {
            Image<Gray, byte> imgGray = imgOutput2.Copy();
            Image<Gray, byte> imghiseq;
            for (int i = 0; i < imgGray.Height; i++)
            {
                for (int j = 0; j < imgGray.Width; j++)
                {
                    if (imgGray.Data[i, j, 0] < 25) //補充 85 臨界值 適用阿茲海默
                    {
                        imgGray.Data[i, j, 0] = 0;
                    }
                    
                }
            }
            imghiseq = new Image<Gray, byte>(imgGray.Width, imgGray.Height, new Gray(0));
            CvInvoke.EqualizeHist(imgGray, imghiseq);
            for (int i = 0; i < imghiseq.Height; i++)
            {
                for (int j = 0; j < imghiseq.Width; j++)
                {
                    if (imghiseq.Data[i, j, 0] < 110)//補充 85 臨界值 適用阿茲海默
                    {
                        imghiseq.Data[i, j, 0] = 0;
                    }
                    
                }
            }
            imageBox3.Image = imghiseq;
            
            
            imgOutput2 = imghiseq.Copy();

        }

        private void button35_Click(object sender, EventArgs e)
        {

        }

        private void button21_Click_1(object sender, EventArgs e)
        {
            numericUpDown4.Maximum = path.Count();
        }

        private void imageBox7_MouseDown(object sender, MouseEventArgs e)
        {
            point4 = e.Location;
            Invalidate();
        }

        private void imageBox8_MouseDown(object sender, MouseEventArgs e)
        {
            point5 = e.Location;
            Invalidate();
        }
        private void button10_1_Click(object sender,EventArgs e)
        {

        }




       /* private void ApplyFilter(bool preview)
        {
            if (imgOutput == null)
            {
                return;
            }
            Bitmap previewBitmap = new Bitmap(imgOutput.Bitmap);
            if (preview == true)
            


        }*/


        private void trackBar2_ValueChanged(object sender, EventArgs e)
        {

        }

        

        /*private void button10_Click_1(object sender, EventArgs e)
{


   Bitmap newBitmap = new Bitmap(imgOutput2.Bitmap);
   for (int x=0; x< imgOutput2.Width; x++)
   {
       for (int y=0; y< imgOutput2.Height; y++)
       {
           Color pixel = newBitmap.GetPixel(x, y);
           int red   = pixel.R;
           int green = pixel.G;
           int blue = pixel.B;
           newBitmap.SetPixel(x, y, Color.FromArgb(255 - red, 255 - green, 255 - blue));



       }


   }
   Emgu.CV.UI.ImageBox imageBox = new Emgu.CV.UI.ImageBox();

   imageBox3.Image = new Image<Bgr,byte>(newBitmap);

}*/

        private void button12_Click(object sender, EventArgs e)
        {
            if (imgOutput2 == null)
            {
                return;
            }

            Image<Gray, byte> grayimg = new Image<Gray, byte>(imgOutput2.Bitmap).Resize((1.0), Emgu.CV.CvEnum.Inter.Linear).Dilate(1).PyrUp().PyrDown();
            grayimg.SmoothGaussian(9);
            grayimg = grayimg.ThresholdBinary(new Gray((int)(numericUpDown8.Value)), new Gray(255)); //閥值90適用帕金森氏症基底核分割
            var element = CvInvoke.GetStructuringElement(ElementShape.Cross, new Size(3, 3), new Point(-1, -1));
            CvInvoke.MorphologyEx(grayimg, grayimg, MorphOp.Erode, element, new Point(-1, -1), 1, BorderType.Reflect, new MCvScalar());
            CvInvoke.MorphologyEx(grayimg, grayimg, MorphOp.Dilate, element, new Point(-1, -1), 1, BorderType.Reflect, new MCvScalar());

            VectorOfVectorOfPoint cn = new VectorOfVectorOfPoint();
            Image<Gray, byte> hier = new Image<Gray, byte>(imgOutput2.Width, imgOutput2.Height);

            //HY1 = 0;
            //HY2 = 0;
            //num1 = 0;

            CvInvoke.FindContours(grayimg, cn, hier, RetrType.Ccomp, ChainApproxMethod.ChainApproxSimple);
            //CvInvoke.FindContours(grayimg, cn, null, RetrType.Ccomp, ChainApproxMethod.ChainApproxNone);

           

            int count = cn.Size;
                for (int i = 0; i < count; i++)
                {
                    using (VectorOfPoint contour = cn[i])
                    using (VectorOfPoint approxContour = new VectorOfPoint())
                    {
                        // 原始輪廓線
                        CvInvoke.DrawContours(imgOutput2, cn, i, new MCvScalar(170), 1);
                        

                        // 近似後輪廓線
                        CvInvoke.ApproxPolyDP(contour, approxContour, CvInvoke.ArcLength(contour, true) * 0.05, true);
                        Point[][] cons = cn.ToArrayOfArray();
                        PointF[][] con2 = Array.ConvertAll<Point[], PointF[]>(cons, new Converter<Point[], PointF[]>(PointToPointF));

                      
                    {

                        PointF[] hull = CvInvoke.ConvexHull(con2[i], true);
                        for (int j = 0; j < hull.Length; j++)
                        {
                            Point p1 = new Point((int)(hull[j].X + 0.5), (int)(hull[j].Y + 0.5));
                            Point p2;
                            if (j == hull.Length - 1)
                                p2 = new Point((int)(hull[0].X + 0.5), (int)(hull[0].Y + 0.5));
                            else
                                p2 = new Point((int)(hull[j + 1].X + 0.5), (int)(hull[j + 1].Y + 0.5));
                            //CvInvoke.Line(imgOutput2, p1, p2, new MCvScalar(150), 2);
                            //CvInvoke.Line(imgOutput2, p1, p2, new MCvScalar(200), 1);
                            //HY1++;
                            //HY2++;
                            //num1++;


                        }
                    }
                    //imgOutput2._Not();
                    imageBox3.Image = imgOutput2;

                }
            }
        }

        private void button48_Click(object sender, EventArgs e)
        {
            if (imgOutput2 == null)
            {
                return;
            }

            Image<Gray, byte> grayimg = new Image<Gray, byte>(imgOutput2.Bitmap);

            //grayimg = grayimg.ThresholdBinary(new Gray((int)(numericUpDown10.Value)), new Gray(255)); //閥值90適用帕金森氏症基底核分割


            VectorOfVectorOfPoint cn = new VectorOfVectorOfPoint();
            //HY1 = 0;
            //HY2 = 0;
            //num1 = 0;

            CvInvoke.FindContours(grayimg, cn, null, RetrType.External, ChainApproxMethod.ChainApproxNone);
            CvInvoke.FindContours(grayimg, cn, null, RetrType.Ccomp, ChainApproxMethod.ChainApproxNone);


            int count = cn.Size;
            for (int i = 0; i < count; i++)
            {
                using (VectorOfPoint contour = cn[i])
                using (VectorOfPoint approxContour = new VectorOfPoint())
                {
                    // 原始輪廓線
                    CvInvoke.DrawContours(imgOutput2, cn, i, new MCvScalar(170), 1);
                    CvInvoke.DrawContours(imgOutput2, cn, i, new MCvScalar(200), 1);

                    // 近似後輪廓線
                    CvInvoke.ApproxPolyDP(contour, approxContour, CvInvoke.ArcLength(contour, true) * 0.03, true);
                    Point[][] cons = cn.ToArrayOfArray();
                    PointF[][] con2 = Array.ConvertAll<Point[], PointF[]>(cons, new Converter<Point[], PointF[]>(PointToPointF));


                    {
                        PointF[] hull = CvInvoke.ConvexHull(con2[i], true);
                        for (int j = 0; j < hull.Length; j++)
                        {
                            Point p1 = new Point((int)(hull[j].X + 0.5), (int)(hull[j].Y + 0.5));
                            Point p2;
                            if (j == hull.Length - 1)
                                p2 = new Point((int)(hull[0].X + 0.5), (int)(hull[0].Y + 0.5));
                            else
                                p2 = new Point((int)(hull[j + 1].X + 0.5), (int)(hull[j + 1].Y + 0.5));

                        


                        }
                    }
                    //imgOutput2._Not();
                    imageBox3.Image = imgOutput2;


                    //imageBox8.Image = imgasd;



                }
            }
        }

        private void button22_Click_1(object sender, EventArgs e)
        {
            Image<Gray, byte> image_Copy = imgInput.Convert<Gray, byte>();
            Image<Gray, byte> mask = new Image<Gray, byte>(imgInput.Width + 2, imgInput.Height + 2);
            Image<Gray, byte> imgasd;


            Rectangle dummy = new System.Drawing.Rectangle(0, 0, 0, 0);
            CvInvoke.FloodFill(image_Copy, mask, point1, new MCvScalar(1), out dummy,
                new MCvScalar(12),
                new MCvScalar(7),
                Emgu.CV.CvEnum.Connectivity.EightConnected,
                Emgu.CV.CvEnum.FloodFillType.Default);
            imgasd = (imgInput.Convert<Gray, byte>()) - image_Copy;

            //MRI 分析
            for (int i = 0; i < imgasd.Height; i++)
            {
                for (int j = 0; j < imgasd.Width; j++)
                {
                    if (imgasd.Data[i, j, 0] > 101) //GM 107 WM 101
                    {
                        imgasd.Data[i, j, 0] = 0;
                    }
                }
            }


            //imageBox7.Image = imgasd;

            imgOutput4 = imgasd;
        }

        private void button41_Click_1(object sender, EventArgs e)
        {
            Image<Gray, byte> imgray = imgOutput4.ThresholdBinary(new Gray(25), new Gray(255));
            Image<Gray, byte> imgray2 =imgOutput4.ThresholdBinary(new Gray((int)numericUpDown7.Value), new Gray(255));
            double AT = 0;
            double AV = 0;
            int area = 0;
            for (int i = 0; i < imgray.Rows; i++)
            {
                for (int j = 0; j < imgray.Cols; j++)
                {
                    if (imgray.Data[i, j, 0] == 255)
                        area++;

                }
            }
            int area1 = 0;
            for (int i = 0; i < imgray2.Rows; i++)
            {
                for (int j = 0; j < imgray2.Cols; j++)
                {
                    if (imgray2.Data[i, j, 0] == 255)
                        area1++;

                }
            }
            AT = System.Convert.ToDouble(area) * 0.2304; //腦組織
            AV = Convert.ToDouble(area1) * 0.2304; //腦實質比
            //listBox1.Items.Add(AT.ToString("f4"));
            //listBox2.Items.Add(AV.ToString("f4"));
            VT.Add(AT);
            VP.Add(AV);
        }

        private void button47_Click(object sender, EventArgs e)
        {
            Image<Gray, byte> imgray = imgOutput5.ThresholdBinary(new Gray(25), new Gray(255));
            Image<Gray, byte> imgray2 = imgOutput5.ThresholdBinary(new Gray((int)numericUpDown7.Value), new Gray(255));
            double AT = 0;
            double AV = 0;
            int area = 0;
            for (int i = 0; i < imgray.Rows; i++)
            {
                for (int j = 0; j < imgray.Cols; j++)
                {
                    if (imgray.Data[i, j, 0] == 255)
                        area++;

                }
            }
            int area1 = 0;
            for (int i = 0; i < imgray2.Rows; i++)
            {
                for (int j = 0; j < imgray2.Cols; j++)
                {
                    if (imgray2.Data[i, j, 0] == 255)
                        area1++;

                }
            }
            AT = System.Convert.ToDouble(area) * 0.2304; //腦組織
            AV = Convert.ToDouble(area1) * 0.2304; //腦實質比
            //listBox1.Items.Add(AT.ToString("f4"));
            //listBox2.Items.Add(AV.ToString("f4"));
            VT.Add(AT);
            VP.Add(AV);
        }

        private void button23_Click(object sender, EventArgs e)
        {

        }

        private void button9_Click_1(object sender, EventArgs e)
        {
             double V = volume(CD_N_PU_TR, 5);
             VolCN = V / VS*100;
             textBox4.Text = V.ToString("f4");
             textBox8.Text =  (VolCN * 100).ToString("f4") + "%";
             
             double V1 = volume(CD_N_PU_TL, 5);
             VolPM = V1 / VS*100;
             textBox11.Text = V1.ToString("f4");
             textBox14.Text = (VolPM * 100).ToString("f4") + "%";
             
             double V2 = volume(OC_B, 5);
             VolOB = V2 / VS*100;
             textBox15.Text = V2.ToString("f4");
             textBox17.Text = (VolOB * 100).ToString("f4") + "%";

            /*for (int z = 0; z < STN.Count; z++)
            {
                SUR_4 = Convert.ToDouble(STN[z]); //SUR_2+= 等於讓這變數持續+-x/ 
            }
             test66 = SUR_4;
             test5 = SUR_4 / VS *100;
             textBox28.Text = test66.ToString("f4");
             textBox22.Text = (test5*100).ToString("f4") + "%";
             double V3 = volume(STN, 5);
             textBox18.Text = V3.ToString("f4");
             VolSN = V3 / VS * 100;
             textBox19.Text = (VolSN * 100).ToString("f4") + "%";*/


        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            
            string Data = textBox20.Text.ToString();

            string filestr = "K:\\研究進度7月1日\\Emgu整合介面\\PD介面\\7月1日論文撰寫修改_目標7月底完成\\7月1日論文修改更新版\\PD論文_科技部_結案報告\\15科技部訓練組測試組與30論文訓練與測試組\\TEST" + Data;
            Excel.Application Excel_Data1 = new Excel.Application(); //設定Excel 應用程序
            Excel.Workbook Excel_Wb = Excel_Data1.Workbooks.Add();
            Excel.Worksheet Excel_Ws1 = new Excel.Worksheet();
            Excel.Worksheet Excel_Ws2 = new Excel.Worksheet();
            Excel.Worksheet Excel_Ws3 = new Excel.Worksheet();
            Excel_Ws1 = Excel_Wb.Worksheets[1];
            Excel_Ws1.Name = "TRODAT個腦資訊";
            Excel_Ws2 = Excel_Wb.Worksheets.Add();
            Excel_Ws2.Name = "CT個腦資訊";
            Excel_Ws3 = Excel_Wb.Worksheets.Add();
            Excel_Ws3.Name = "MRI個腦資訊";






            Excel_Data1.Cells[1, 1] = "編號";
            Excel_Data1.Cells[1, 2] = "全腦體積";
            Excel_Data1.Cells[1, 3] = "腦組織體積";
            Excel_Data1.Cells[1, 4] = "腦組織比例";
            Excel_Data1.Cells[1, 5] = "萎縮比";
            Excel_Data1.Cells[1, 6] = "Caudate Nucleus  Putamen(R)比值";
            Excel_Data1.Cells[1, 7] = "%"; //R面積比
            Excel_Data1.Cells[1, 8] = "Caudate Nucleus  Putamen(L)比值";
            Excel_Data1.Cells[1, 9] = "%"; //L面積比
            Excel_Data1.Cells[1, 10] = "Occipital Bone";
            Excel_Data1.Cells[1, 11] = "%";//OC面積比 
            //Excel_Data1.Cells[1, 12] = "Substantia nigra(SN)";
            //Excel_Data1.Cells[1, 13] = "%";//SN面積比
            //Excel_Data1.Cells[1, 14] = "Caudate Nucleus  Putamen(R)體積";
            //Excel_Data1.Cells[1, 15] = "%";//R體積比
            //Excel_Data1.Cells[1, 16] = "Caudate Nucleus  Putamen(L)體積";
            //Excel_Data1.Cells[1, 17] = "%";//L體積比
            //Excel_Data1.Cells[1, 18] = "Occipital Bone體積";
            //Excel_Data1.Cells[1, 19] = "%";//OC體積比
            //Excel_Data1.Cells[1, 20] = "Substantia nigra(SN)體積";
            //Excel_Data1.Cells[1, 21] = "%";//SN體積比
            Excel_Data1.Cells[1, 12] = "紋狀體總面積";
            Excel_Data1.Cells[1, 13] = "紋狀體攝取值";
            Excel_Data1.Cells[1, 14] = "不對稱性指標";
            Excel_Data1.Cells[1, 15] = "系統HY";

            
            

            Excel_Wb.SaveAs(filestr);

            Excel_Ws1 = null;
            Excel_Wb.Close();
            Excel_Wb = null;
            Excel_Data1.Quit();
            Excel_Data1 = null;
            num1 = 0;
            label33.Text = num1.ToString();

        }

        private void button24_Click_2(object sender, EventArgs e)
        {
            num1++;
            count++;
            string Data = textBox20.Text.ToString();

            string filestr = "K:\\研究進度7月1日\\Emgu整合介面\\PD介面\\7月1日論文撰寫修改_目標7月底完成\\7月1日論文修改更新版\\PD論文_科技部_結案報告\\15科技部訓練組測試組與30論文訓練與測試組\\TEST" + Data;
            Excel.Application Excel_Data1 = new Excel.Application(); //設定Excel 應用程序
            Excel.Workbook Excel_Wb = Excel_Data1.Workbooks.Open(filestr);
            Excel.Worksheet Excel_Ws1 = new Excel.Worksheet();
            Excel.Worksheet Excel_Ws2 = new Excel.Worksheet();
            Excel.Worksheet Excel_Ws3 = new Excel.Worksheet();
            Excel_Ws1 = Excel_Wb.Worksheets[1];
            Excel_Ws2 = Excel_Wb.Worksheets[2];
            Excel_Ws3 = Excel_Wb.Worksheets[3];


            //Excel_Ws1.Activate();
            Excel_Data1.Cells[count, 1] = num1;
            Excel_Data1.Cells[count, 2] = textBox10.Text;
            Excel_Data1.Cells[count, 3] = textBox6.Text;
            Excel_Data1.Cells[count, 4] = textBox9.Text;
            Excel_Data1.Cells[count, 5] = textBox7.Text;
            Excel_Data1.Cells[count, 6] = textBox21.Text;
            Excel_Data1.Cells[count, 7] = textBox35.Text;
            Excel_Data1.Cells[count, 8] = textBox5.Text;
            Excel_Data1.Cells[count, 9] = textBox16.Text;
            Excel_Data1.Cells[count, 10] = textBox25.Text;
            Excel_Data1.Cells[count, 11] = textBox37.Text;
            //Excel_Data1.Cells[count, 12] = textBox28.Text;
            //Excel_Data1.Cells[count, 13] = textBox22.Text;
            //Excel_Data1.Cells[count, 14] = textBox4.Text;
            //Excel_Data1.Cells[count, 15] = textBox8.Text;
            //Excel_Data1.Cells[count, 16] = textBox11.Text;
            //Excel_Data1.Cells[count, 17] = textBox14.Text;
            //Excel_Data1.Cells[count, 18] = textBox15.Text;
            //Excel_Data1.Cells[count, 19] = textBox17.Text;
            //Excel_Data1.Cells[count, 20] = textBox18.Text;
            //Excel_Data1.Cells[count, 21] = textBox19.Text;
            Excel_Data1.Cells[count, 12] = textBox26.Text;
            Excel_Data1.Cells[count, 13] = textBox3.Text;
            Excel_Data1.Cells[count, 14] = textBox13.Text;
            Excel_Data1.Cells[count, 15] = textBox2.Text;

            //Excel_Ws2.Activate();
            Excel_Data1.Cells[count, 1] = num1;
            Excel_Data1.Cells[count, 2] = textBox10.Text;
            Excel_Data1.Cells[count, 3] = textBox6.Text;
            Excel_Data1.Cells[count, 4] = textBox9.Text;
            Excel_Data1.Cells[count, 5] = textBox7.Text;
            Excel_Data1.Cells[count, 6] = textBox21.Text;
            Excel_Data1.Cells[count, 7] = textBox35.Text;
            Excel_Data1.Cells[count, 8] = textBox5.Text;
            Excel_Data1.Cells[count, 9] = textBox16.Text;
            Excel_Data1.Cells[count, 10] = textBox25.Text;
            Excel_Data1.Cells[count, 11] = textBox37.Text;
            //Excel_Data1.Cells[count, 12] = textBox28.Text;
            //Excel_Data1.Cells[count, 13] = textBox22.Text;
            //Excel_Data1.Cells[count, 14] = textBox4.Text;
            //Excel_Data1.Cells[count, 15] = textBox8.Text;
            //Excel_Data1.Cells[count, 16] = textBox11.Text;
            //Excel_Data1.Cells[count, 17] = textBox14.Text;
            //Excel_Data1.Cells[count, 18] = textBox15.Text;
            //Excel_Data1.Cells[count, 19] = textBox17.Text;
            //Excel_Data1.Cells[count, 20] = textBox18.Text;
            //Excel_Data1.Cells[count, 21] = textBox19.Text;
            Excel_Data1.Cells[count, 12] = textBox26.Text;
            Excel_Data1.Cells[count, 13] = textBox3.Text;
            Excel_Data1.Cells[count, 14] = textBox13.Text;
            Excel_Data1.Cells[count, 15] = textBox2.Text;

            //Excel_Ws3.Activate();
            Excel_Data1.Cells[count, 1] = num1;
            Excel_Data1.Cells[count, 2] = textBox10.Text;
            Excel_Data1.Cells[count, 3] = textBox6.Text;
            Excel_Data1.Cells[count, 4] = textBox9.Text;
            Excel_Data1.Cells[count, 5] = textBox7.Text;
            Excel_Data1.Cells[count, 6] = textBox21.Text;
            Excel_Data1.Cells[count, 7] = textBox35.Text;
            Excel_Data1.Cells[count, 8] = textBox5.Text;
            Excel_Data1.Cells[count, 9] = textBox16.Text;
            Excel_Data1.Cells[count, 10] = textBox25.Text;
            Excel_Data1.Cells[count, 11] = textBox37.Text;
            //Excel_Data1.Cells[count, 12] = textBox28.Text;
            //Excel_Data1.Cells[count, 13] = textBox22.Text;
            //Excel_Data1.Cells[count, 14] = textBox4.Text;
            //Excel_Data1.Cells[count, 15] = textBox8.Text;
            //Excel_Data1.Cells[count, 16] = textBox11.Text;
            //Excel_Data1.Cells[count, 17] = textBox14.Text;
            //Excel_Data1.Cells[count, 18] = textBox15.Text;
            //Excel_Data1.Cells[count, 19] = textBox17.Text;
            //Excel_Data1.Cells[count, 20] = textBox18.Text;
            //Excel_Data1.Cells[count, 21] = textBox19.Text;
            Excel_Data1.Cells[count, 12] = textBox26.Text;
            Excel_Data1.Cells[count, 13] = textBox3.Text;
            Excel_Data1.Cells[count, 14] = textBox13.Text;
            Excel_Data1.Cells[count, 15] = textBox2.Text;



            Excel_Wb.Save();
            Excel_Ws1 = null;
            Excel_Wb.Close();
            Excel_Wb = null;
            Excel_Data1.Quit();
            Excel_Data1 = null;
            label33.Text = num1.ToString();
        }

        private void button25_Click_4(object sender, EventArgs e)
        {
            num1--;
            count--;
            string Data = textBox20.Text.ToString();

            string filestr = "K:\\研究進度7月1日\\Emgu整合介面\\PD介面\\7月1日論文撰寫修改_目標7月底完成\\7月1日論文修改更新版\\15科技部訓練組測試組\\TEST" + Data;
            Excel.Application Excel_Data1 = new Excel.Application(); //設定Excel 應用程序
            Excel.Workbook Excel_Wb = Excel_Data1.Workbooks.Open(filestr);
            Excel.Worksheet Excel_Ws1 = new Excel.Worksheet();
            Excel.Worksheet Excel_Ws2 = new Excel.Worksheet();
            Excel.Worksheet Excel_Ws3 = new Excel.Worksheet();
            Excel_Ws1 = Excel_Wb.Worksheets[1];
            Excel_Ws2 = Excel_Wb.Worksheets[2];
            Excel_Ws3 = Excel_Wb.Worksheets[3];

            //Excel_Ws1.Activate();
            Excel_Data1.Cells[count, 1] = num1;
            Excel_Data1.Cells[count, 2] = textBox10.Text;
            Excel_Data1.Cells[count, 3] = textBox6.Text;
            Excel_Data1.Cells[count, 4] = textBox9.Text;
            Excel_Data1.Cells[count, 5] = textBox7.Text;
            Excel_Data1.Cells[count, 6] = textBox21.Text;
            Excel_Data1.Cells[count, 7] = textBox35.Text;
            Excel_Data1.Cells[count, 8] = textBox5.Text;
            Excel_Data1.Cells[count, 9] = textBox16.Text;
            Excel_Data1.Cells[count, 10] = textBox25.Text;
            Excel_Data1.Cells[count, 11] = textBox37.Text;
            //Excel_Data1.Cells[count, 12] = textBox28.Text;
            //Excel_Data1.Cells[count, 13] = textBox22.Text;
            //Excel_Data1.Cells[count, 14] = textBox4.Text;
            //Excel_Data1.Cells[count, 15] = textBox8.Text;
            //Excel_Data1.Cells[count, 16] = textBox11.Text;
            //Excel_Data1.Cells[count, 17] = textBox14.Text;
            //Excel_Data1.Cells[count, 18] = textBox15.Text;
            //Excel_Data1.Cells[count, 19] = textBox17.Text;
            //Excel_Data1.Cells[count, 20] = textBox18.Text;
            //Excel_Data1.Cells[count, 21] = textBox19.Text;
            Excel_Data1.Cells[count, 12] = textBox26.Text;
            Excel_Data1.Cells[count, 13] = textBox3.Text;
            Excel_Data1.Cells[count, 14] = textBox13.Text;
            Excel_Data1.Cells[count, 15] = textBox2.Text;

            //Excel_Ws2.Activate();
            Excel_Data1.Cells[count, 1] = num1;
            Excel_Data1.Cells[count, 2] = textBox10.Text;
            Excel_Data1.Cells[count, 3] = textBox6.Text;
            Excel_Data1.Cells[count, 4] = textBox9.Text;
            Excel_Data1.Cells[count, 5] = textBox7.Text;
            Excel_Data1.Cells[count, 6] = textBox21.Text;
            Excel_Data1.Cells[count, 7] = textBox35.Text;
            Excel_Data1.Cells[count, 8] = textBox5.Text;
            Excel_Data1.Cells[count, 9] = textBox16.Text;
            Excel_Data1.Cells[count, 10] = textBox25.Text;
            Excel_Data1.Cells[count, 11] = textBox37.Text;
            //Excel_Data1.Cells[count, 12] = textBox28.Text;
            //Excel_Data1.Cells[count, 13] = textBox22.Text;
            //Excel_Data1.Cells[count, 14] = textBox4.Text;
            //Excel_Data1.Cells[count, 15] = textBox8.Text;
            //Excel_Data1.Cells[count, 16] = textBox11.Text;
            //Excel_Data1.Cells[count, 17] = textBox14.Text;
            //Excel_Data1.Cells[count, 18] = textBox15.Text;
            //Excel_Data1.Cells[count, 19] = textBox17.Text;
            //Excel_Data1.Cells[count, 20] = textBox18.Text;
            //Excel_Data1.Cells[count, 21] = textBox19.Text;
            Excel_Data1.Cells[count, 12] = textBox26.Text;
            Excel_Data1.Cells[count, 13] = textBox3.Text;
            Excel_Data1.Cells[count, 14] = textBox13.Text;
            Excel_Data1.Cells[count, 15] = textBox2.Text;

            //Excel_Ws3.Activate();
            Excel_Data1.Cells[count, 1] = num1;
            Excel_Data1.Cells[count, 2] = textBox10.Text;
            Excel_Data1.Cells[count, 3] = textBox6.Text;
            Excel_Data1.Cells[count, 4] = textBox9.Text;
            Excel_Data1.Cells[count, 5] = textBox7.Text;
            Excel_Data1.Cells[count, 6] = textBox21.Text;
            Excel_Data1.Cells[count, 7] = textBox35.Text;
            Excel_Data1.Cells[count, 8] = textBox5.Text;
            Excel_Data1.Cells[count, 9] = textBox16.Text;
            Excel_Data1.Cells[count, 10] = textBox25.Text;
            Excel_Data1.Cells[count, 11] = textBox37.Text;
            //Excel_Data1.Cells[count, 12] = textBox28.Text;
            //Excel_Data1.Cells[count, 13] = textBox22.Text;
            //Excel_Data1.Cells[count, 14] = textBox4.Text;
            //Excel_Data1.Cells[count, 15] = textBox8.Text;
            //Excel_Data1.Cells[count, 16] = textBox11.Text;
            //Excel_Data1.Cells[count, 17] = textBox14.Text;
            //Excel_Data1.Cells[count, 18] = textBox15.Text;
            //Excel_Data1.Cells[count, 19] = textBox17.Text;
            //Excel_Data1.Cells[count, 20] = textBox18.Text;
            //Excel_Data1.Cells[count, 21] = textBox19.Text;
            Excel_Data1.Cells[count, 12] = textBox26.Text;
            Excel_Data1.Cells[count, 13] = textBox3.Text;
            Excel_Data1.Cells[count, 14] = textBox13.Text;
            Excel_Data1.Cells[count, 15] = textBox2.Text;



            Excel_Wb.Save();
            Excel_Ws1 = null;
            Excel_Wb.Close();
            Excel_Wb = null;
            Excel_Data1.Quit();
            Excel_Data1 = null;
            label33.Text = num1.ToString();
            label33.Invalidate();

        }

        private void button25_Click_5(object sender, EventArgs e)
        {
            for (int z = 0; z < STN.Count; z++)
            {
                SUR_4 = Convert.ToDouble(STN[z]); //SUR_2+= 等於讓這變數持續+-x/ 
            }
            test66 = SUR_4;
            test5 = SUR_4 / VS * 100;
            textBox28.Text = test66.ToString("f4");
            textBox22.Text = (test5 * 100).ToString("f4") + "%";
            double V3 = volume(STN, 5);
            textBox18.Text = V3.ToString("f4");
            VolSN = V3 / VS * 100;
            textBox19.Text = (VolSN * 100).ToString("f4") + "%";
        }
    }
    }
    

        






