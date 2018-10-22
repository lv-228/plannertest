using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Planner
{
    /*Класс управлением и созданием цехов, по идее можно разбить на 
     * два класса Shop и Planner что-бы можно было создавать несколько цехов, 
     * управляемых допустим одним планировщиком, но недостаточно времени,
     * 
     * !!Выяснилось что в c# нельзя событийно создать глобальный объект, либо я не понял как, 
     * пришлось делать класс статическим, думаю правильный подход разделить данный класс на Shop и Planner, этот переименовать на Shop
     * а Planner как раз будет статическим и хранить объекты данного класса, собственно как данный класс хранит объекты класса machineTools
     * !!
     */
    public static class ShopPlanner
    {
        //Список хранящий в себе цех (все машины учавствующие в обработке материалов)
        public static List<machineTools> shop = new List<machineTools>();

        public static string name;//Название цеха

        static ShopPlanner()
        {
            
        }

        //Создать цэх
        public static bool constructShop()
        {
            for(int i = 1; i < Nomenclatures.ovens.Count();i++)
            {
                shop.Add(new machineTools(Nomenclatures.ovens[i][0]));
            }
            if(shop.Count == 0)
            {
                return false;
            }
            return true;
        }


        //Рендер формы с спецификациями машин
        public static TabControl renderFormWithOvensSpecifications()
        {
            //Создание tabControl объекта на форме
            TabControl tabControl1 = new TabControl();
            tabControl1.Location = new Point(5, 5);
            tabControl1.Size = new Size(386, 200);

            for(int i = 0; i < shop.Count; i++)
            {
                //Создание страницы у tabControl
                TabPage page1 = new TabPage();
                page1.Text = machineTools.getOvenNameById(ShopPlanner.shop[i].id);

                //Создание listView внутри таба
                ListView lv = new ListView();
                lv.Size = new Size(352, 150);
                lv.Location = new Point(5, 5);
                lv.View = View.Details;
                lv.Columns.Add("id", -2);
                lv.Columns.Add("Имя", -2);
                lv.Columns.Add("Материал", -2);
                lv.Columns.Add("Время обработки", -2);
                ListViewItem lvi = new ListViewItem(new string[] { shop[i].id, shop[i].name , Nomenclatures.viewMaterialById(shop[i].materialsTimes[0][1]), shop[i].materialsTimes[0][2] + " мин" });
                lv.Items.Add(lvi);
                for (int j = 1; j < shop[i].materialsTimes.Count; j++)
                {
                    ListViewItem lviMaterials = new ListViewItem(new string[] { "", "", Nomenclatures.viewMaterialById(shop[i].materialsTimes[j][1]), shop[i].materialsTimes[j][2] + " мин"});
                    lv.Items.Add(lviMaterials);
                }
                tabControl1.Controls.Add(page1);
                page1.Controls.Add(lv);
            }
            return tabControl1;
        }

        //Рендер формы с спецификациями машин
        public static ListView renderFormWithStatisticAboutBatch(List<string> array)
        {
            ListView lv = new ListView();
            lv.Size = new Size(250, 200);
            lv.Location = new Point(5, 5);
            lv.View = View.Details;
            lv.Columns.Add("Материал", -2);
            lv.Columns.Add("Количество", -2);
            for (int i = 0; i < array.Count; i += 2)
            {
                ListViewItem lviMaterials = new ListViewItem(new string[] {array[i], array[i+1]});
                lv.Items.Add(lviMaterials);
            }
            return lv;
        }

        //
    }
}