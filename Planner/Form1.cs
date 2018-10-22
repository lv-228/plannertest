using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelLibrary;//Для работы с xls
using ExcelLibrary.SpreadSheet;//Для работы с xls
using Excel = BotAgent.DataExporter.Excel;

namespace Planner
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(Nomenclatures.materials.Count == 0 || Nomenclatures.ovens.Count == 0 || Nomenclatures.ovensSpecifications.Count == 0)
            {
                MessageBox.Show("Пожалуйста, сначала загрузите файлы с данными о материалах и машинах", 
                    "Недостаточно данных", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            Nomenclatures.openPartiesFile(listView1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
                Nomenclatures.openMachineToolsFile(listView3);
        }

        private void button4_Click(object sender, EventArgs e)
        {
                Nomenclatures.openMaterialsFile(listView2, listView1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (Nomenclatures.ovens.Count == 0 || Nomenclatures.materials.Count == 0)
            {
                MessageBox.Show("Пожалуйста, сначала выберите файл с именами, идентификаторами машин, а так-же файл с материалами",
                    "Недостаточно данных", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            Nomenclatures.openSpecificationsFile();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (ShopPlanner.shop.Count == 0)
            {
                MessageBox.Show("Пожалуйста, для просмотра статистики загрузите партию",
                    "Недостаточно данных", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            Form2 formAboutOvens = new Form2();
            formAboutOvens.renderTabIndexMachineName();
            formAboutOvens.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (Nomenclatures.parties.Count == 0)
            {
                MessageBox.Show("Пожалуйста, для просмотра статистики загрузите партию",
                    "Недостаточно данных", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            Form3 formStatistic = new Form3();
            formStatistic.renderStatistic();
            formStatistic.Show();
        }
    }
}