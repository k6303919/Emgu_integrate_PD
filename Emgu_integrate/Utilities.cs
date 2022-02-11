using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using Emgu.CV.UI;
using System.Windows.Forms;
using System.Drawing;
namespace Emgu_integrate
{
    class Utilities
    {
        public static void ConvertCoordinates(ImageBox pic,
            out int X0, out int Y0, int x, int y)
        {
            int pic_hgt = pic.ClientSize.Height;
            int pic_wid = pic.ClientSize.Width;
            int img_hgt = pic.Height;
            int img_wid = pic.Width;

            X0 = x;
            Y0 = y;
            switch (pic.SizeMode)
            {
                case PictureBoxSizeMode.AutoSize:
                case PictureBoxSizeMode.Normal:
                    // These are okay. Leave them alone.
                    break;
                case PictureBoxSizeMode.CenterImage:
                    X0 = x - (pic_wid - img_wid) / 2;
                    Y0 = y - (pic_hgt - img_hgt) / 2;
                    break;
                case PictureBoxSizeMode.StretchImage:
                    X0 = (int)(img_wid * x / (float)pic_wid);
                    Y0 = (int)(img_hgt * y / (float)pic_hgt);
                    break;
                case PictureBoxSizeMode.Zoom:
                    float pic_aspect = pic_wid / (float)pic_hgt;
                    float img_aspect = img_wid / (float)img_wid;
                    if (pic_aspect > img_aspect)
                    {
                        // The PictureBox is wider/shorter than the image.
                        Y0 = (int)(img_hgt * y / (float)pic_hgt);

                        // The image fills the height of the PictureBox.
                        // Get its width.
                        float scaled_width = img_wid * pic_hgt / img_hgt;
                        float dx = (pic_wid - scaled_width) / 2;
                        X0 = (int)((x - dx) * img_hgt / (float)pic_hgt);
                    }
                    else
                    {
                        // The PictureBox is taller/thinner than the image.
                        X0 = (int)(img_wid * x / (float)pic_wid);

                        // The image fills the height of the PictureBox.
                        // Get its height.
                        float scaled_height = img_hgt * pic_wid / img_wid;
                        float dy = (pic_hgt - scaled_height) / 2;
                        Y0 = (int)((y - dy) * img_wid / pic_wid);
                    }
                    break;
            }
        }

        public static void getEntropy(Bitmap bmp, int radius, int cenX, int cenY, out double[] JJ, out double[] HH, out double[] GG, out double[] QQ, out double[] CC, out double[] HQ, out double[] EE, out double[] JH)
        {
            /// <param name="bmp">传入一张图片</param>
            /// <param name="radius">需要计算的区域的半径</param>
            /// <param name="cenX">圆心横坐标</param>
            /// <param name="cenY">圆心纵坐标</param>
            /// <param name="JJ">最大機率値</param>
            /// <param name="HH">熵</param>
            /// <param name="GG">对比度</param>
            /// <param name="QQ">同质性</param>
            /// <param name="CC">相关性</param>
            /// <param name="HQ">變異數</param>
            /// <param name="EE">能量
            /// <param name="JH">Cluster

            double[,,] p = getP(bmp, radius, cenX, cenY);//得到共生矩阵

            double J = 0;//最大機率値，J值越大，纹理越粗，反之越细
            double H = 0;//熵，H值越大，纹理较多
            double G = 0;//对比度，G值越大，纹理越清晰
            double Q = 0;//同质性，Q值越大，图像纹理不同区域间变化小，局部非常均匀
            double E = 0;//能量，J值越大，纹理越粗，反之越细
            JJ = new double[3];
            HH = new double[3];
            GG = new double[3];
            QQ = new double[3];
            CC = new double[3];//相关性
            HQ = new double[3];//變異數
            EE = new double[3];//能量
            JH = new double[3];//Cluster
            double[] ui = new double[3];
            double[] uj = new double[3];
            double[] si = new double[3];
            double[] sj = new double[3];
            for (int i = 0; i < 3; i++)
            {
                ui[i] = 0;
                uj[i] = 0;
                si[i] = 0;
                sj[i] = 0;
                CC[i] = 0;
                EE[i] = 0;
            }
            for (int n = 0; n < 3; n++)
            {
                for (int ii = 0; ii < 16; ii++)
                    for (int jj = 0; jj < 16; jj++)
                    {
                        if (p[ii, jj, n] != 0)
                        {
                            J += (ii) * (jj) * p[ii, jj, n];
                            H -= p[ii, jj, n] * Math.Log(p[ii, jj, n]);
                            G += (ii - jj) * (ii - jj) * p[ii, jj, n];
                            Q += p[ii, jj, n] / (1 + (ii - jj) * (ii - jj));
                            E += p[ii, jj, n] * p[ii, jj, n];
                        }
                        ui[n] = ii * p[ii, jj, n] + ui[n];
                        uj[n] = jj * p[ii, jj, n] + uj[n];
                    }
                JJ[n] = J;
                HH[n] = H;
                GG[n] = G;
                QQ[n] = Q;
                EE[n] = E;

            }
            for (int n = 0; n < 3; n++)
            {
                for (int ii = 0; ii < 16; ii++)
                    for (int jj = 0; jj < 16; jj++)
                    {
                        si[n] = (ii - ui[n]) * (ii - ui[n]) * p[ii, jj, n] + si[n];
                        sj[n] = (jj - uj[n]) * (ii - uj[n]) * p[ii, jj, n] + sj[n];
                        CC[n] = ii * jj * p[ii, jj, n] + CC[n];


                    }
                CC[n] = (CC[n] - ui[n] * uj[n]) / (si[n] * sj[n]);
            }
        }
        public static double[,,] getP(Bitmap bmp, int radius, int cenX, int cenY)//计算共生矩阵，灰度级弄成16
        {
            int[,] g = new int[2 * radius, 2 * radius];
            //转换灰度级
            for (int w = cenX - radius; w < cenX + radius; w++)
                for (int h = cenY - radius; h < cenY + radius; h++)
                {
                    int Gray = (30 * bmp.GetPixel(w, h).R + 59 * bmp.GetPixel(w, h).G + 11 * bmp.GetPixel(w, h).B) / 100;//计算灰度值
                    //int Gray = bmp.GetPixel(w, h).R;
                    g[w - cenX + radius, h - cenY + radius] = Gray / 16;
                }
            double[,,] p = new double[16, 16, 3];//三个方向的，向右一步，向下一步，向右下一步
            for (int n = 0; n < 3; n++)//都转换成概率
                for (int i = 0; i < 16; i++)
                    for (int j = 0; j < 16; j++)
                    {
                        p[i, j, n] = 0;
                    }
            for (int m = 0; m < 16; m++)
                for (int n = 0; n < 16; n++)
                    for (int w = 0; w < 2 * radius; w++)
                    {
                        for (int h = 0; h < 2 * radius; h++)
                        {
                            if (g[w, h] == m)
                            {
                                if (h < 2 * radius - 1)
                                    if (g[w, h + 1] == n)
                                        p[m, n, 0] = p[m, n, 0] + 1;
                                if (h < 2 * radius - 1 && w < 2 * radius - 1)
                                    if (g[w + 1, h + 1] == n)
                                        p[m, n, 1] = p[m, n, 1] + 1;
                                if (w < 2 * radius - 1)
                                    if (g[w + 1, h] == n)
                                        p[m, n, 2] = p[m, n, 2] + 1;

                            }
                        }

                    }

            for (int n = 0; n < 3; n++)//都转换成概率
            {
                double sum = 0;
                for (int i = 0; i < 16; i++)
                    for (int j = 0; j < 16; j++)
                    {
                        if (p[i, j, n] != 0)
                            sum += p[i, j, n];
                    }
                for (int i = 0; i < 16; i++)
                    for (int j = 0; j < 16; j++)
                    {
                        if (p[i, j, n] != 0)
                            p[i, j, n] = p[i, j, n] / sum;
                    }
            }
            return p;
        }
    }
}

