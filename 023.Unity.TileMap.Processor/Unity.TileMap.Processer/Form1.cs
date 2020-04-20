using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Unity.TileMap.Processer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label2_DragDrop(object sender, DragEventArgs e)
        {
            label2.Text = "";
            String[] files = e.Data.GetData(DataFormats.FileDrop, false) as String[];
            foreach(string f in files)
            {
                label2.Text = label2.Text + f + "|";
            }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            label2_DragDrop(sender, e);
        }

        private void Form1_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void label2_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] files = label2.Text.Split('|');
            foreach(string f in files)
            {
                if (File.Exists(f))
                {
                    Image i = Bitmap.FromFile(f);
                    int w = Convert.ToInt32(xBox.Text);
                    int h = Convert.ToInt32(yBox.Text);
                    int wc = i.Width / w;int hc = i.Height / h;
                    Image o = new Bitmap(i.Width + wc * 2,i.Height + hc * 2);
                    Graphics g = Graphics.FromImage(o);
                    for(int j = 0;j < wc; j++)
                    {
                        for (int k = 0; k < hc; k++)
                        {
                            g.DrawImage(i, 
                                new Rectangle((int)(j * (w + 2)+1), (int)(k * (h + 2)+1), w, h),
                                new Rectangle((int)(j * w), (int)(k * h), w, h),
                                GraphicsUnit.Pixel);
                            g.DrawImage(i,
                                new Rectangle((int)(j * (w + 2)), (int)(k * (h + 2)+1), 1, h),
                                new Rectangle((int)(j * w), (int)(k * h), 1, h),
                                GraphicsUnit.Pixel);
                            g.DrawImage(i,
                                new Rectangle((int)(j * (w + 2) + 1+w), (int)(k * (h + 2)+1), 1, h),
                                new Rectangle((int)(j * w + w - 1), (int)(k * h), 1, h),
                                GraphicsUnit.Pixel);
                            g.DrawImage(i,
                                new Rectangle((int)(j * (w + 2)+1), (int)(k * (h + 2)), w, 1),
                                new Rectangle((int)(j * w), (int)(k * h), w, 1),
                                GraphicsUnit.Pixel);
                            g.DrawImage(i,
                                new Rectangle((int)(j * (w + 2)+1), (int)(k * (h + 2) + 1+h), w, 1),
                                new Rectangle((int)(j * w ), (int)(k * h + h - 1), w, 1),
                                GraphicsUnit.Pixel);

                            g.DrawImage(i,
                                new Rectangle((int)(j * (w + 2)), (int)(k * (h + 2)), 1, 1),
                                new Rectangle((int)(j * w), (int)(k * h), 1, 1),
                                GraphicsUnit.Pixel);
                            g.DrawImage(i,
                                new Rectangle((int)(j * (w + 2) + 1 + w), (int)(k * (h + 2)), 1, 1),
                                new Rectangle((int)(j * w + w - 1), (int)(k * h), 1, 1),
                                GraphicsUnit.Pixel);
                            g.DrawImage(i,
                                new Rectangle((int)(j * (w + 2)), (int)(k * (h + 2) + 1 + h), 1, 1),
                                new Rectangle((int)(j * w), (int)(k * h + h - 1), 1, 1),
                                GraphicsUnit.Pixel);
                            g.DrawImage(i,
                                new Rectangle((int)(j * (w + 2) + 1 + w), (int)(k * (h + 2) + 1 + h), 1, 1),
                                new Rectangle((int)(j * w + w - 1), (int)(k * h + h - 1), 1, 1),
                                GraphicsUnit.Pixel);

                        }
                    }
                    o.Save(f.Replace(".",".convert."));
                }
            }
            MessageBox.Show("Success!","TileMaps",MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
