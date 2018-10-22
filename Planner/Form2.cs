using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Planner
{
    public partial class Form2 : Form
    {
        //Если форма была срендарена меняется на true
        public static bool status = false;
        public Form2()
        {
            InitializeComponent();
        }

        public void renderTabIndexMachineName()
        {
            this.Controls.Add(ShopPlanner.renderFormWithOvensSpecifications());
        }
    }
}
