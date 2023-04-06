using MaterialSkin2DotNet;
using MaterialSkin2DotNet.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Lab_data.View
{
    public partial class Form2 : MaterialForm
    {
        readonly MaterialSkinManager skinManager;
        string s = "Сведения ";
        public Form2(Model.ApplicationContext db, int i)
        {
            InitializeComponent();
            skinManager = MaterialSkinManager.Instance;
            skinManager.AddFormToManage(this);
            skinManager.Theme = MaterialSkinManager.Themes.DARK;
            skinManager.ColorScheme = new ColorScheme(Primary.LightBlue800, Primary.LightBlue900, Primary.LightBlue500, Accent.LightBlue200, TextShade.WHITE);
            materialDataTable1.AutoGenerateColumns = false;
            switch (i)
            {
                case 0: var query = db.employees.GroupBy(x => x.gender).Select(g => new { g.Key, coun = g.Count() }).ToList(); materialDataTable1.DataSource = query; s += "о поле сотрудников"; break;
                case 1: var query1 = db.employees.GroupBy(x => x.Family).Select(g => new { g.Key, coun = g.Count() }).ToList(); materialDataTable1.DataSource = query1; s += "о семейном положении сотрудников"; break;
                case 2:
                    var query2 = db.employees.Join(db.Posts, x => x.Post_id, y => y.Id,
                    (x, y) => new
                    {
                        Key = y.Name,
                        Id = x.ID
                    }).GroupBy(x => x.Key).Select(g => new { g.Key, coun = g.Count() }).ToList();

                    materialDataTable1.DataSource = query2; s += "о должностях сотрудников"; break;

                case 3: var query3 = db.employees.GroupBy(x => x.degree).Select(g => new { g.Key, coun = g.Count() }).ToList(); materialDataTable1.DataSource = query3; s += "о ученых степенях сотрудников"; break;
                case 4: var query4 = db.employees.GroupBy(x => x.title).Select(g => new { g.Key, coun = g.Count() }).ToList(); materialDataTable1.DataSource = query4; s += "о ученых званиях сотрудников"; break;
            }
            materialDataTable1.Columns[0].DataPropertyName = "Key";
            materialDataTable1.Columns[1].DataPropertyName = "coun";
            this.Text = s;

        }

        private void materialButton1_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Сохранить данные о детях сотрудников?", "Сохранить", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    Excel.Application app = new Excel.Application
                    {
                        //Отобразить Excel
                        Visible = false,
                        //Количество листов в рабочей книге
                        SheetsInNewWorkbook = 1
                    };
                    //Добавить рабочую книгу
                    Excel.Workbook workBook = app.Workbooks.Add(Type.Missing);
                    //Отключить отображение окон с сообщениями
                    app.DisplayAlerts = false;
                    //Получаем первый лист документа (счет начинается с 1)
                    Excel.Worksheet sheet = (Excel.Worksheet)app.Worksheets.get_Item(1);
                    //Название листа (вкладки снизу)
                    sheet.Name = "Сведения";
                    Excel.Range r = sheet.Cells[1, 1];
                    Excel.Range r1 = sheet.Cells[1, 2];
                    Excel.Range range = sheet.get_Range(r, r1);
                    range.Merge(Type.Missing);
                    sheet.Cells[1, 1] = s;
                    sheet.Cells[2, 1] = "Название";
                    sheet.Cells[2, 2] = "Количество сотрудников";

                    int i = 2;
                    foreach (DataGridViewRow y in materialDataTable1.Rows)
                    {
                        i++;
                        sheet.Cells[i, 1] = y.Cells[0].Value;
                        sheet.Cells[i, 2] = y.Cells[1].Value;
                    }
                    Excel.Range r5 = sheet.Cells[1, 1];
                    Excel.Range r6 = sheet.Cells[i, 2];

                    Excel.Range range3 = sheet.get_Range(r5, r6);
                    range3.Borders.Color = ColorTranslator.ToOle(Color.Black);
                    range3.EntireColumn.AutoFit();
                    var save = new System.Windows.Forms.SaveFileDialog();
                    save.Filter = "Excel files(*.xlsx)|*.xlsx";    //формат выходных файлов
                                                                   //Сохраняем файл
                    if (save.ShowDialog() == DialogResult.Cancel)
                        return;
                    string filename = save.FileName;
                    app.Application.ActiveWorkbook.SaveAs(filename, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    app.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(app);

                    app = null;
                    workBook = null;
                    sheet = null;
                    GC.Collect(); // убрать за собой
                }
                catch (Exception ex) //возникает при ошибках
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message.ToString()); //выводим полученную ошибку
                    System.Windows.Forms.MessageBox.Show(ex.StackTrace.ToString());
                }
            }
        }
    }
}
