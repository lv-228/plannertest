using Planner;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
//Класс описывают одну машину из которых собирается цех в классе ShopPlanner
public class machineTools
{
    public bool MachineState;//Состояние печи, по умолчанию свободна.

    public string discription;//Описание машины

    public string name;//Описание машины

    public string id;//Идентификатор машины

    public List<List<string>> materialsTimes = new List<List<string>>();/*Подходящее сырье для данной машины,
    если сырье не указанно в данном массиве то машина не может с ним работать ключь [1] = id материала, 
    [2] = время обработки*/

	public machineTools(string ovenId)
	{
        this.id = ovenId;
        this.name = machineTools.getOvenNameById(ovenId);
        this.setMaterialsOvenById(ovenId);
        this.MachineState = false;
        this.discription = createDiscription(ovenId);
	}

    //Представление id и имени печей в ListView
    public static void renderOvensInListView(ListView lv)
    {
        for (int i = 1; i < Nomenclatures.ovens.Count; i++)
        {
            ListViewItem lvi = new ListViewItem(Nomenclatures.ovens[i][0]);
            lvi.SubItems.Add(Nomenclatures.ovens[i][1]);
            lv.Items.Add(lvi);
        }
    }

    //Возвращает имя печи по ее id
    public static string getOvenNameById(string ovenId)
    {
        for (int i = 0; i < Nomenclatures.ovens.Count; i++)
        {
            if (Nomenclatures.ovens[i][0] == ovenId)
                return Nomenclatures.ovens[i][1];
        }
        throw new DataIntegrityException($"Печь с id = {ovenId} не описана!");
    }
    
    //Возвращает материалы обрабатываемые печью
    public List<List<string>> getMaterialsOvenById(string ovenId)
    {
        List<List<string>> raw = new List<List<string>>();
        for (int i = 0; i < Nomenclatures.ovens.Count; i++)
        {
            if (Nomenclatures.ovensSpecifications[i][0] == ovenId)
            {
                raw.Add(Nomenclatures.ovensSpecifications[i]);
            }
        }
        if (this.materialsTimes.Count == 0)
            throw new DataIntegrityException($"Печь с id = {ovenId} и именем = {getOvenNameById(ovenId)} не обрабатывает ни одного материала из указанных в файле!");
        return raw;
    }

    //Записывает материалы обрабатываемые печью в объект класса
    public bool setMaterialsOvenById(string ovenId)
    {
        for (int i = 0; i < Nomenclatures.ovensSpecifications.Count; i++)
        {
            if (Nomenclatures.ovensSpecifications[i][0] == ovenId)
            {
                this.materialsTimes.Add(Nomenclatures.ovensSpecifications[i]);
            }
        }
        if (this.materialsTimes.Count == 0)
            throw new DataIntegrityException($"Печь с id = {ovenId} и именем = {getOvenNameById(ovenId)} не обрабатывает ни одного материала из указанных в файле!");
        return true;
    }

    //Создает описание машины из полученых данных для удобства пользователя
    protected string createDiscription(string ovenId)
    {
        string str = "Печь обрабатывает: ";
        for (int i = 0; i < this.materialsTimes.Count; i++)
        {
            str += $" {Nomenclatures.viewMaterialById(materialsTimes[i][1])} за {materialsTimes[i][2]} минут";
        }
        if (str != "Печь обрабатывает: ")
            return str;
        throw new NomenclaturesException("Не удалось составить описания для печи!");
    }
}