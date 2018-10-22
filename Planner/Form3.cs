using BotAgent.DataExporter;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace Planner
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        public void renderStatistic()
        {
            this.Controls.Add(ShopPlanner.renderFormWithStatisticAboutBatch(Nomenclatures.setStatisticAboutBatch()));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Nomenclatures.saveStatInXlsxFile();
        }
    }
}
