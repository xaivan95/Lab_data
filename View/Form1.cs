using MaterialSkin2DotNet;
using MaterialSkin2DotNet.Controls;
using Lab_data.Model;
using Excel = Microsoft.Office.Interop.Excel;
using Lab_data.View;
using System.Linq;

namespace Lab_data
{
    public partial class Form1 : MaterialForm
    {
        readonly MaterialSkinManager skinManager;
        Model.ApplicationContext db = new Model.ApplicationContext();
        public Form1()
        {
            InitializeComponent();
            skinManager = MaterialSkinManager.Instance;
            skinManager.AddFormToManage(this);
            skinManager.Theme = MaterialSkinManager.Themes.DARK;
            skinManager.ColorScheme = new ColorScheme(Primary.LightBlue800, Primary.LightBlue900, Primary.LightBlue500, Accent.LightBlue200, TextShade.WHITE);
            GetEmployer();
            GetPost("");
            GetChildren();
        }

        public void GetEmployer()
        {
            materialDataTable1.AutoGenerateColumns = false;
            materialDataTable1.DataSource = db.employees.ToList();
            materialDataTable1.Columns[0].DataPropertyName = "ID";
            materialDataTable1.Columns[1].DataPropertyName = "FIO";
            materialDataTable1.Columns[2].DataPropertyName = "gender";
            materialDataTable1.Columns[3].DataPropertyName = "Age";
            materialDataTable1.Columns[4].DataPropertyName = "Family";
            materialDataTable1.Columns[5].DataPropertyName = "childre";
            materialDataTable1.Columns[6].DataPropertyName = "Post_id";
            materialDataTable1.Columns[7].DataPropertyName = "degree";
            materialDataTable1.Columns[8].DataPropertyName = "title";

            (materialDataTable1.Columns[6] as DataGridViewComboBoxColumn).DisplayMember = "Name";
            (materialDataTable1.Columns[6] as DataGridViewComboBoxColumn).ValueMember = "Id";
            (materialDataTable1.Columns[6] as DataGridViewComboBoxColumn).DataSource = db.Posts.ToList();
            GetCount();
        }

        public void GetCount()
        {
            var query = db.employees.GroupJoin(db.childrens, x => x.ID, y => y.employee_id,
                (x, coun) => new employee
                {
                    ID = x.ID,
                    childre = coun.Count()

                }).ToList();

            foreach (var u in materialDataTable1.Rows)
                (u as DataGridViewRow).Cells[5].Value = query.First(x => x.ID == (int)(u as DataGridViewRow).Cells[0].Value).childre;
        }

        public void GetPost(string s)
        {
            materialDataTable2.AutoGenerateColumns = false;
            materialDataTable2.DataSource = db.Posts.Where(x => x.Name.ToLower().Contains(s)).ToList();
            materialDataTable2.Columns[0].DataPropertyName = "Id";
            materialDataTable2.Columns[1].DataPropertyName = "Name";
        }

        public void GetChildren()
        {
            materialDataTable3.AutoGenerateColumns = false;
            materialDataTable3.DataSource = db.childrens.ToList();
            materialDataTable3.Columns[0].DataPropertyName = "ID";
            materialDataTable3.Columns[1].DataPropertyName = "employee_id";
            materialDataTable3.Columns[2].DataPropertyName = "FIO";
            materialDataTable3.Columns[3].DataPropertyName = "birthdate";

            (materialDataTable3.Columns[1] as DataGridViewComboBoxColumn).DisplayMember = "FIO";
            (materialDataTable3.Columns[1] as DataGridViewComboBoxColumn).ValueMember = "Id";
            (materialDataTable3.Columns[1] as DataGridViewComboBoxColumn).DataSource = db.employees.ToList();
        }

        private void materialButton1_Click(object sender, EventArgs e)
        {
            materialDataTable1.EndEdit();
            db.SaveChanges();
        }

        private void materialButton3_Click(object sender, EventArgs e)
        {
            if (db.Posts.Count() > 0)
            {
                db.employees.Add(new employee(db.Posts.First().Id));
                db.SaveChanges();
                GetEmployer();
                GetChildren();
            }
            else
                MessageBox.Show("Добавте должность");
        }

        private void materialButton2_Click(object sender, EventArgs e)
        {
            if (materialDataTable1.SelectedRows.Count > 0)
                if ((int)materialDataTable1.SelectedRows[0].Cells[0].Value > 0)
                    db.employees.Remove(db.employees.First(x => x.ID == (int)materialDataTable1.SelectedRows[0].Cells[0].Value));
            db.SaveChanges();
            GetEmployer();
            GetChildren();
        }

        private void materialTextBox21_TextChanged(object sender, EventArgs e)
        {
            string s = materialTextBox21.Text.ToLower();
            //поиск 
            bool flag = false; //состояние поиска
            materialDataTable1.CurrentCell = null; //снимаем выделения строк с таблицы
            if (s.Equals("")) //если ничего не введено
            {
                foreach (DataGridViewRow row in materialDataTable1.Rows)
                {
                    row.Visible = true;//делаем все строчки видимыми
                }
            }
            else //если что-то ввели
            {
                foreach (DataGridViewRow row in materialDataTable1.Rows)
                {
                    flag = false;//состояние поиска - не найдено
                    if (row.Cells[1].Value != null)
                        if (row.Cells[1].Value.ToString().ToLower().Contains(s)) flag = true;//поиск
                    if (row.Cells[2].Value != null)
                        if (row.Cells[2].Value.ToString().ToLower().Contains(s)) flag = true;//поиск

                    if (row.Cells[3].Value != null)
                        if (row.Cells[3].Value.ToString().ToLower().Contains(s)) flag = true;//поиск
                    if (row.Cells[4].Value != null)
                        if (row.Cells[4].Value.ToString().ToLower().Contains(s)) flag = true;//поиск

                    if (row.Cells[5].Value != null)
                        if (row.Cells[5].Value.ToString().ToLower().Contains(s)) flag = true;//поиск

                    var a = db.Posts.First(x => x.Id == (int)(row.Cells[6] as DataGridViewComboBoxCell).Value).Name;
                    if (a.ToLower().Contains(s)) flag = true;//поиск
                    if (row.Cells[7].Value != null)
                        if (row.Cells[7].Value.ToString().ToLower().Contains(s)) flag = true;//поиск
                    if (row.Cells[8].Value != null)
                        if (row.Cells[8].Value.ToString().ToLower().Contains(s)) flag = true;//поиск
                    if (flag) row.Visible = true;//если нашли совпадение строчка видна
                    else row.Visible = false;//иначе скрываем
                }
            }
        }

        private void materialTextBox22_TextChanged(object sender, EventArgs e)
        {
            GetPost(materialTextBox22.Text);
        }

        private void materialButton7_Click(object sender, EventArgs e)
        {
            materialDataTable2.EndEdit();
            db.SaveChanges();
        }

        private void materialButton5_Click(object sender, EventArgs e)
        {
            db.Posts.Add(new Post());
            db.SaveChanges();
            GetPost("");
            GetEmployer();
        }

        private void materialButton6_Click(object sender, EventArgs e)
        {
            if (materialDataTable2.SelectedRows.Count > 0)
                if ((int)materialDataTable2.SelectedRows[0].Cells[0].Value > 0)
                {
                    db.Posts.Remove(db.Posts.First(x => x.Id == (int)materialDataTable2.SelectedRows[0].Cells[0].Value));
                    if (db.Posts.Count() == 0)
                        db.employees.RemoveRange(db.employees.Where(x => x.Post_id == (int)materialDataTable2.SelectedRows[0].Cells[0].Value));
                    else
                        foreach (var em in db.employees.Where(x => x.Post_id == (int)materialDataTable2.SelectedRows[0].Cells[0].Value))
                            em.Post_id = db.Posts.Last().Id;
                    db.SaveChanges();
                }

            GetPost("");
        }


        private void materialButton9_Click(object sender, EventArgs e)
        {
            if (materialDataTable3.SelectedRows.Count > 0)
                if ((int)materialDataTable3.SelectedRows[0].Cells[0].Value > 0)
                {
                    db.childrens.Remove(db.childrens.First(x => x.ID == (int)materialDataTable3.SelectedRows[0].Cells[0].Value));
                }
            db.SaveChanges();
            GetEmployer();
            GetChildren();
        }

        private void materialButton10_Click(object sender, EventArgs e)
        {
            materialDataTable3.EndEdit();
            db.SaveChanges();
            GetEmployer();
        }

        private void materialButton8_Click(object sender, EventArgs e)
        {
            if (db.employees.Count() > 0)
            {
                db.childrens.Add(new children(db.employees.First().ID));
                db.SaveChanges();
                GetChildren();
                GetEmployer();
            }
            else
                MessageBox.Show("Добавьте сотрудника");
        }

        private void materialTextBox23_TextChanged(object sender, EventArgs e)
        {
            string s = materialTextBox23.Text.ToLower();
            //поиск 
            bool flag = false; //состояние поиска
            materialDataTable3.CurrentCell = null; //снимаем выделения строк с таблицы
            if (s.Equals("")) //если ничего не введено
            {
                foreach (DataGridViewRow row in materialDataTable3.Rows)
                {
                    row.Visible = true;//делаем все строчки видимыми
                }
            }
            else //если что-то ввели
            {
                foreach (DataGridViewRow row in materialDataTable3.Rows)
                {
                    flag = false;//состояние поиска - не найдено
                    if (row.Cells[2].Value != null)
                        if (row.Cells[2].Value.ToString().ToLower().Contains(s)) flag = true;//поиск
                    if (row.Cells[3].Value != null)
                        if (row.Cells[3].Value.ToString().ToLower().Contains(s)) flag = true;//поиск

                    var a = db.employees.First(x => x.ID == (int)(row.Cells[1] as DataGridViewComboBoxCell).Value).FIO;
                    if (a.ToLower().Contains(s)) flag = true;//поиск

                    if (flag) row.Visible = true;//если нашли совпадение строчка видна
                    else row.Visible = false;//иначе скрываем
                }
            }
        }

        private void materialDataTable1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 5)
            {
                materialTabControl2.SelectedIndex = 2;
                materialTextBox23.Text = materialDataTable1.Rows[e.RowIndex].Cells[1].Value.ToString();
            }
        }

        private void materialButton4_Click(object sender, EventArgs e)
        {
            Form2 frm = new Form2(db, materialComboBox1.SelectedIndex);
            frm.ShowDialog();
        }

        private void materialButton11_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Сохранить данные о сотрудниках?", "Сохранить", MessageBoxButtons.YesNo);
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
                    sheet.Name = "Сотрудники";

                    sheet.Cells[1, 1] = "№ п/п";
                    sheet.Cells[1, 2] = "ФИО";
                    sheet.Cells[1, 3] = "Пол";
                    sheet.Cells[1, 4] = "Возраст";
                    sheet.Cells[1, 5] = "Семейное положение";
                    sheet.Cells[1, 6] = "Должность";
                    sheet.Cells[1, 7] = "Ученая степень";
                    sheet.Cells[1, 8] = "Ученое звание";

                    int i = 1;
                    var n = 0;
                    foreach (var y in db.employees.ToList())
                    {
                        n++;
                        i++;
                        sheet.Cells[i, 1] = n;
                        sheet.Cells[i, 2] = y.FIO;
                        sheet.Cells[i, 3] = y.gender;
                        sheet.Cells[i, 4] = y.Age;
                        sheet.Cells[i, 5] = y.Family;
                        sheet.Cells[i, 6] = db.Posts.First(x => x.Id == y.Post_id).Name;
                        sheet.Cells[i, 7] = y.degree;
                        sheet.Cells[i, 8] = y.title;
                    }
                    Excel.Range r5 = sheet.Cells[1, 1];
                    Excel.Range r6 = sheet.Cells[i, 8];
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

        private void materialButton12_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Сохранить данные о должностях?", "Сохранить", MessageBoxButtons.YesNo);
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
                    sheet.Name = "Должности";

                    sheet.Cells[1, 1] = "№ п/п";
                    sheet.Cells[1, 2] = "Название";


                    int i = 1;
                    var n = 0;
                    foreach (var y in db.Posts.ToList())
                    {
                        n++;
                        i++;
                        sheet.Cells[i, 1] = n;
                        sheet.Cells[i, 2] = y.Name;

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

        private void materialButton13_Click(object sender, EventArgs e)
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
                    sheet.Name = "Дети сотрудников";

                    sheet.Cells[1, 1] = "№ п/п";
                    sheet.Cells[1, 2] = "ФИО сотрудника";
                    sheet.Cells[1, 3] = "ФИО ребенка";
                    sheet.Cells[1, 4] = "Дата рождения";


                    int i = 1;
                    var n = 0;
                    foreach (var y in db.childrens.OrderBy(x => x.employee_id).ToList())
                    {
                        n++;
                        i++;
                        sheet.Cells[i, 1] = n;
                        sheet.Cells[i, 2] = db.employees.First(x => x.ID == y.employee_id).FIO;
                        sheet.Cells[i, 3] = y.FIO;
                        sheet.Cells[i, 4] = y.birthdate;

                    }
                    Excel.Range r5 = sheet.Cells[1, 1];
                    Excel.Range r6 = sheet.Cells[i, 4];
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