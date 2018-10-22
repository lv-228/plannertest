using BotAgent.DataExporter;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelLibrary.CompoundDocumentFormat;
using ExcelLibrary.SpreadSheet;

namespace Planner
{
    //Класс описывает получение данных из xlsx файлов
    abstract class Nomenclatures
    {
        //Сюда сохраняются материалы из загруженного файла
        public static List<List<string>> materials = new List<List<string>>();

        //Сюда сохраняются партии из загруженного файла
        public static List<List<string>> parties = new List<List<string>>();

        //Сюда сохраняются печи из загруженного файла
        public static List<List<string>> ovens = new List<List<string>>();

        //Сюда сохраняются данные о печах из загруженного файла
        public static List<List<string>> ovensSpecifications = new List<List<string>>();

        //Сюда сохраняются данные о печах из загруженного файла
        public static List<string> stat = new List<string>();

        //Сюда сохраняются данные о печах из загруженного файла
        public static List<string> columnsInMaterials = new List<string>(new string[] { "id", "nomenclature" });

        //Сюда сохраняются данные о печах из загруженного файла
        public static List<string> columnsInOvensIdName = new List<string>(new string[] { "id", "name" });

        //Сюда сохраняются данные о печах из загруженного файла
        public static List<string> columnsInParties = new List<string>(new string[] { "id", "nomenclature id" });

        //Сюда сохраняются данные о печах из загруженного файла
        public static List<string> columnsInOvensSpecifications = new List<string>(new string[] { "machine tool id", "nomenclature id", "operation time" });

        //Метод отвечающий за открытие xlsx файлов
        public static Excel openXslxFile(string path)
        {
            Excel xlsxfile = new Excel();
            xlsxfile.FileOpen(path);
            return xlsxfile;
        }

        //Открыть файл с описание и id материалов
        public static void openMaterialsFile(ListView lv, ListView lvp)
        {
            try
            {
                Nomenclatures.setMaterialsByFile(lv, lvp);
            }
            //toDo узнать можно ли как-то обертку в try - catch вынести в функцию или сократить для избежания использования повторного кода
            catch (NomenclaturesException ex)
            {
                MessageBox.Show("Ошибка структуры файла:\n" + ex.Message, "Ошибка структуры файла",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (System.IO.IOException ex)
            {
                MessageBox.Show("Ошибка доступа к файлу:\n" + ex.Message, "Ошибка доступа",
                MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (DataIntegrityException ex)
            {
                MessageBox.Show("Ошибка целостности данных:\n" + ex.Message, "Нарушение целостности данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        //Открыть файл с именами машин
        public static void openMachineToolsFile(ListView lv)
        {
            try
            {
                Nomenclatures.setOvensByFile(lv);
            }
            catch (NomenclaturesException ex)
            {
                MessageBox.Show("Ошибка структуры файла:\n" + ex.Message, "Ошибка структуры файла",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (System.IO.IOException ex)
            {
                MessageBox.Show("Ошибка доступа к файлу:\n" + ex.Message, "Ошибка доступа",
                MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (DataIntegrityException ex)
            {
                MessageBox.Show("Ошибка целостности данных:\n" + ex.Message, "Нарушение целостности данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        //Открывает файл с партиями
        public static void openPartiesFile(ListView lv)
        {
            try
            {
                Nomenclatures.setPartiesByFile(lv);
            }
            catch (NomenclaturesException ex)
            {
                MessageBox.Show("Ошибка структуры файла:\n" + ex.Message, "Ошибка структуры файла",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (DataIntegrityException ex)
            {
                MessageBox.Show("Ошибка целостности данных:\n" + ex.Message, "Нарушение целостности данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (System.IO.IOException ex)
            {
                MessageBox.Show("Ошибка доступа к файлу:\n" + ex.Message, "Ошибка доступа",
                MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        //Открывает файл с характеристиками машин
        public static void openSpecificationsFile()
        {
            try
            {
                setOvensSpecifications();
            }
            catch (NomenclaturesException ex)
            {
                MessageBox.Show("Ошибка структуры файла:\n" + ex.Message, "Ошибка структуры файла",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (DataIntegrityException ex)
            {
                MessageBox.Show("Ошибка целостности данных:\n" + ex.Message, "Нарушение целостности данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            catch (System.IO.IOException ex)
            {
                MessageBox.Show("Ошибка доступа к файлу:\n" + ex.Message, "Ошибка доступа",
                MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        //Метод сохраняющий спецификации машин в переменную ovensSpecifications
        public static void setOvensSpecifications()
        {
            if (ShopPlanner.shop.Count == 0)
            {
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Filter = "Excel files(*.xlsx)|*.xlsx";
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    Nomenclatures.validateSpecificationFile(openFile.FileName);
                    Excel xlsxfile = Nomenclatures.openXslxFile(openFile.FileName);
                    Nomenclatures.ovensSpecifications.AddRange(xlsxfile.Rows);
                    ShopPlanner.constructShop();
                    MessageBox.Show("Характеристики загруженны, цэх успешно создан");
                }
            }
            else
            {
                DialogResult result = MessageBox.Show("Спецификации машин уже загружены, хотите обновить?"
                    , "Подтвердите действие"
                    , MessageBoxButtons.OKCancel
                    , MessageBoxIcon.Warning
                );
                if (result == DialogResult.OK)
                {
                    OpenFileDialog openFile = new OpenFileDialog();
                    openFile.Filter = "Excel files(*.xlsx)|*.xlsx";
                    if (openFile.ShowDialog() == DialogResult.OK)
                    {
                        Nomenclatures.validateSpecificationFile(openFile.FileName);
                        Nomenclatures.ovensSpecifications.Clear();
                        Excel xlsxfile = Nomenclatures.openXslxFile(openFile.FileName);
                        Nomenclatures.ovensSpecifications.AddRange(xlsxfile.Rows);
                        ShopPlanner.shop.Clear();
                        ShopPlanner.constructShop();
                        MessageBox.Show("Характеристики загруженны, цэх успешно создан");
                    }
                }
            }
        }

        //Метод сохраняющий материалы в переменную materials
        public static void setMaterialsByFile(ListView lv,ListView lvp)
        {
            if (materials.Count == 0)
            {
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Filter = "Excel files(*.xlsx)|*.xlsx";
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    Nomenclatures.validateMaterialsFile(openFile.FileName);
                    Excel xlsxfile = Nomenclatures.openXslxFile(openFile.FileName);
                    materials.AddRange(xlsxfile.Rows);
                    Nomenclatures.renderMaterialsInListView(lv);
                }
            }
            else
            {
                DialogResult result = MessageBox.Show("Вы уже загружали ранее файл с описанием материалов, хотите ли вы перезаписать данные?(При перезаписи данные о партии будут сброшены, для сохранения целостности данных)"
                    , "Подтвердите действие"
                    , MessageBoxButtons.OKCancel
                    , MessageBoxIcon.Warning
                    );
                if (result == DialogResult.OK)
                {
                    OpenFileDialog openFile = new OpenFileDialog();
                    openFile.Filter = "Excel files(*.xlsx)|*.xlsx";
                    if (openFile.ShowDialog() == DialogResult.OK)
                    {
                        Nomenclatures.validateMaterialsFile(openFile.FileName);
                        Nomenclatures.parties.Clear();
                        Nomenclatures.materials.Clear();
                        lvp.Items.Clear();
                        lv.Items.Clear();
                        Excel xlsxfile = Nomenclatures.openXslxFile(openFile.FileName);
                        materials.AddRange(xlsxfile.Rows);
                        Nomenclatures.renderMaterialsInListView(lv);
                    }
                }
            }
        }

        //Метод сохраняющий партии в переменную parties
        public static void setPartiesByFile(ListView lv)
        {
            if(parties.Count == 0)
            {
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Filter = "Excel files(*.xlsx)|*.xlsx";
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    Nomenclatures.validatePartiesFile(openFile.FileName);
                    Excel xlsxfile = Nomenclatures.openXslxFile(openFile.FileName);
                    parties.AddRange(xlsxfile.Rows);
                    Nomenclatures.renderPartiesInListView(lv);
                }
            }
            else
            {
                DialogResult result = MessageBox.Show("Вы уже загружали ранее файл с партиями, хотите ли вы перезаписать данные?"
                    , "Подтвердите действие"
                    , MessageBoxButtons.OKCancel
                    , MessageBoxIcon.Warning
                    );
                if (result == DialogResult.OK)
                {
                    OpenFileDialog openFile = new OpenFileDialog();
                    openFile.Filter = "Excel files(*.xlsx)|*.xlsx";
                    if (openFile.ShowDialog() == DialogResult.OK)
                    {
                        Nomenclatures.validatePartiesFile(openFile.FileName);
                        Nomenclatures.parties.Clear();
                        lv.Items.Clear();
                        Excel xlsxfile = Nomenclatures.openXslxFile(openFile.FileName);
                        parties.AddRange(xlsxfile.Rows);
                        Nomenclatures.renderPartiesInListView(lv);
                    }
                }
            }
        }

        //Метод сохраняющий печи в переменную parties
        public static void setOvensByFile(ListView lv)
        {
            if (ovens.Count == 0)
            {
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Filter = "Excel files(*.xlsx)|*.xlsx";
                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    Nomenclatures.validateMachineToolsFile(openFile.FileName);
                    Excel xlsxfile = Nomenclatures.openXslxFile(openFile.FileName);
                    ovens.AddRange(xlsxfile.Rows);
                    machineTools.renderOvensInListView(lv);
                }
            }
            else
            {
                DialogResult result = MessageBox.Show("Вы уже загружали ранее файл с идентификаторма машин, хотите ли вы перезаписать данные?(Сведения о машинах будут сброшены, для сохранения целостности данных)"
                    , "Подтвердите действие"
                    , MessageBoxButtons.OKCancel
                    , MessageBoxIcon.Warning
                    );
                if (result == DialogResult.OK)
                {
                    OpenFileDialog openFile = new OpenFileDialog();
                    openFile.Filter = "Excel files(*.xlsx)|*.xlsx";
                    if (openFile.ShowDialog() == DialogResult.OK)
                    {
                        Nomenclatures.validateMachineToolsFile(openFile.FileName);
                        Nomenclatures.ovens.Clear();
                        Nomenclatures.ovensSpecifications.Clear();
                        lv.Items.Clear();
                        Excel xlsxfile = Nomenclatures.openXslxFile(openFile.FileName);
                        ovens.AddRange(xlsxfile.Rows);
                        machineTools.renderOvensInListView(lv);
                    }
                }
            }
        }

        //Проверка на соответствие название столбцов ожидаемому
        private static bool validateColumnsInFile(List<string> array, List<List<string>> outFile)
        {
            int count = 0;
            for(int j = 0; j < array.Count; j++)
                for(int i = 0; i < outFile[0].Count; i++)
                {
                    if (array[j] == outFile[0][i])
                        count++;
                }
            if (count == array.Count)
                return true;
            throw new NomenclaturesException(generateExceptionColumnsString(array));
        }

        //Генерирует ошибку при валидации столбцов
        protected static string generateExceptionColumnsString(List<string> array)
        {
            string str = "Ошибка в структуре файла, ожидаются столбцы:";
            for (int i = 0; i < array.Count; i++)
                str += $" [{array[i]}]";
            return str;
        }

        //Валидирует значение в файле по заданному регулярному выражению
        protected static bool validate(Regex reg, string row)
        {
            MatchCollection matche = reg.Matches(row);
            if (matche.Count == 0)
            {
                return false;
            }
            return true;
        }

        //Проверяет приходит ли в партии материал который не описан в файле предназначенном для описи материалов (тафтология)
        private static bool checkMaterialInParties(string materialsInPartiesId)
        {
            for (int i = 1; i < materials.Count; i++)
            {
                if (materialsInPartiesId == materials[i][0])
                    return true;
            }
            return false;
        }

        //Проверка, обрабатывает ли хотя-бы одна машина указанная в файле материал или нет
        public static bool canMaterialProcessedOrNot(string materialsInPartiesId)
        {
            for (int i = 1; i < ovensSpecifications.Count; i++)
            {
                if (materialsInPartiesId == ovensSpecifications[i][1])
                    return true;
            }
            return false;
        }

        //Валидирует файл с записями о названиях и id машин
        public static bool validateMachineToolsFile(string filePath)
        {
            Excel xlsxfile = new Excel();
            xlsxfile.FileOpen(filePath);
            Nomenclatures.validateColumnsInFile(columnsInOvensIdName, xlsxfile.Rows);
            Regex regex = new Regex(@"Печь [0-9]");
            for (int i = 1; i < xlsxfile.Rows.Count; i++)
            {
                if (!Nomenclatures.validate(regex, xlsxfile.Rows[i][1]))
                    throw new NomenclaturesException($"Не верное имя печи, ожидается:\n 'Печь [номер] (например: Печь 7)'\n получено:\n Имя: {xlsxfile.Rows[i][1]}\n Строка: {i}");
            }
            return true;
        }

        //Валидирует файл содержащий название материала и его id
        public static bool validateMaterialsFile(string filePath)
        {
            Excel xlsxfile = new Excel();
            xlsxfile.FileOpen(filePath);
            Nomenclatures.validateColumnsInFile(columnsInMaterials, xlsxfile.Rows);
            Regex regex = new Regex(@"[А-Яа-я]+\s?[А-Яа-я]*$");
            for (int i = 1; i < xlsxfile.Rows.Count; i++)
            {
                if(!Nomenclatures.validate(regex, xlsxfile.Rows[i][1]))
                    throw new NomenclaturesException($"Не верное имя сырья, ожидается:\n 'Одно или два слова на кирилице разделенных пробелом или переносом строки'\n получено:\n Имя: {xlsxfile.Rows[i][1]}\n Строка: {i}");
            }
            return true;
        }

        //Валидирует файл содержащий характеристики печей
        public static bool validateSpecificationFile(string filePath)
        {
            Excel xlsxfile = new Excel();
            xlsxfile.FileOpen(filePath);
            Nomenclatures.validateColumnsInFile(columnsInOvensSpecifications, xlsxfile.Rows);
            Regex regex = new Regex(@"[0-9]+$");
            for(int j = 0; j < xlsxfile.Rows[0].Count; j++)
                for (int i = 1; i < xlsxfile.Rows.Count; i++)
                {
                    if (!Nomenclatures.validate(regex, xlsxfile.Rows[i][j]))
                        throw new NomenclaturesException($"Ожидается числовое значение в столбце {xlsxfile.Rows[0][j]}, \n получено:\n id: {xlsxfile.Rows[i][j]}\n Строка: {i}");
                }
            return true;
        }

        /*Валидирует файл содержащий партии (id и nomenclatuer id)
         если в файле партий есть id которого небыло в файле с материалами выбрасывается исключение*/
        public static bool validatePartiesFile(string filePath)
        {
            Excel xlsxfile = new Excel();
            xlsxfile.FileOpen(filePath);
            Nomenclatures.validateColumnsInFile(columnsInParties, xlsxfile.Rows);
            Regex regex = new Regex(@"[0-9]+$");
            for (int i = 1; i < xlsxfile.Rows.Count; i++)
            {
                if (!Nomenclatures.validate(regex, xlsxfile.Rows[i][1]))
                    throw new NomenclaturesException($"Не верное id материала, ожидается:\n 'Числовое значение в столбце nomenclature id'\n получено:\n Имя: {xlsxfile.Rows[i][1]}\n Строка: {i}");
                if (!Nomenclatures.checkMaterialInParties(xlsxfile.Rows[i][1]))
                    throw new DataIntegrityException($"Не найден данный id в списке материалов:\n получено:\n id: {xlsxfile.Rows[i][1]}\n Строка: {i}");
                if (!Nomenclatures.canMaterialProcessedOrNot(xlsxfile.Rows[i][1]))
                    throw new DataIntegrityException($"Ни одна из загруженных машин не обрабатывает материал {viewMaterialById(xlsxfile.Rows[i][1])} :\n получено:\n id: {xlsxfile.Rows[i][1]}\n Название: {viewMaterialById(xlsxfile.Rows[i][1])} \n Строка: {i}");
            }
            return true;
        }

        //Выводит имя материала по его id
        public static string viewMaterialById(string materialId)
        {
            for(int i = 0; i < materials.Count; i++)
            {
                if (materials[i][0] == materialId)
                    return materials[i][1];
            }
            throw new DataIntegrityException($"Материал с id = {materialId} не описан!");
        }

        //Вывод материалов в ListView
        public static void renderMaterialsInListView(ListView lv)
        {
            for (int i = 1; i < Nomenclatures.materials.Count; i++)
            {
                ListViewItem lvi = new ListViewItem(Nomenclatures.materials[i][0]);
                lvi.SubItems.Add(Nomenclatures.materials[i][1]);
                lv.Items.Add(lvi);
            }
        }

        //Вывод партии в ListView
        public static void renderPartiesInListView(ListView lv)
        {
            for (int i = 1; i < Nomenclatures.parties.Count; i++)
            {
                ListViewItem lvi = new ListViewItem(Nomenclatures.parties[i][0]);
                lvi.SubItems.Add(Nomenclatures.viewMaterialById(Nomenclatures.parties[i][1]));
                lv.Items.Add(lvi);
            }
        }

        //Создание статистики по партии, запись количества материала и его наименование
        public static List<string> setStatisticAboutBatch()
        {
            List<string> array = new List<string>();
            int count = 0;
            for (int i = 1; i < materials.Count; i++)
            {
                array.Add(materials[i][1]);
                for(int j = 1; j < parties.Count; j++)
                {
                    if (parties[j][1] == materials[i][0])
                        count++;
                }
                array.Add(count.ToString());
                count = 0;
            }
            stat.AddRange(array);
            return array;
        }

        //Сохранить xlsx файл
        public static void saveStatInXlsxFile()
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "Excel files(*.xlsx)|*.xlsx";
            if (saveFile.ShowDialog() == DialogResult.Cancel)
                return;
            Workbook wb = new Workbook();
            Worksheet wh = new Worksheet("statFile");
            wh.Cells[0, 0] = new Cell("Материал партии");
            wh.Cells[0, 1] = new Cell("Количество");
            for (int i = 0, j = 1; i < Nomenclatures.stat.Count; i += 2, j++)
            {
                wh.Cells[j, 0] = new Cell(Nomenclatures.stat[i]);
                wh.Cells[j, 1] = new Cell(Nomenclatures.stat[i + 1]);
            }
            wb.Worksheets.Add(wh);
            wb.Save(saveFile.FileName);
            MessageBox.Show("Файл сохранен");
        }
    }
}