using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;
using Word = Microsoft.Office.Interop.Word;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Upload;
using Google.Apis.Util.Store;
using System.Threading;
using System.IO;
using Newtonsoft.Json;

namespace ProstaffProg
{
    public partial class Form1 : Form
    {
        List<UserData> users = new List<UserData>();
        UserData currentUser = null;
        List<DateTime> historyDTP4 = new List<DateTime>();
        List<DateTime> historyDTP5 = new List<DateTime>();
        List<DateTime> historyDTP6 = new List<DateTime>();
        List<DateTime> historyDTP7 = new List<DateTime>();
        List<DateTime> historyDTP8 = new List<DateTime>();
        private bool calendarOpen = false;
        private DateTime lastPickerValue;

        public Form1()
        {
            InitializeComponent();

            textBox10.TextChanged += AnyInputChanged;
            textBox17.TextChanged += AnyInputChanged;
            textBox19.TextChanged += AnyInputChanged;
            textBox24.TextChanged += AnyInputChanged;
            textBox30.TextChanged += AnyInputChanged;

            comboBox3.SelectedIndexChanged += comboBox3_SelectedIndexChanged;
            comboBox2.SelectedIndexChanged += comboBox2_SelectedIndexChanged;

            dateTimePicker4.ValueChanged += dateTimePicker4_ValueChanged;
            dateTimePicker5.ValueChanged += dateTimePicker5_ValueChanged;
            dateTimePicker6.ValueChanged += dateTimePicker6_ValueChanged;
            dateTimePicker7.ValueChanged += dateTimePicker7_ValueChanged;
            dateTimePicker8.ValueChanged += dateTimePicker8_ValueChanged;

        }
        List<string> photoPaths = new List<string>();
        int currentPhotoIndex = 0;
        private bool isUserChangingDate4 = false;
        private bool isUserChangingDate5 = false;
        private bool isUserChangingDate6 = false;
        private bool isUserChangingDate7 = false;
        private bool isUserChangingDate8 = false;
        private bool suppressDateEvents = false;
        private void button4_Click(object sender, EventArgs e)
        {
            string factoryName = textBox11.Text.Trim();

            if (!string.IsNullOrEmpty(factoryName))
            {
                // Проверяем, есть ли уже столбец с таким названием
                if (!dataGridView2.Columns.Contains(factoryName))
                {
                    // Добавляем новый столбец
                    dataGridView2.Columns.Add(factoryName, factoryName);

                    // Добавляем в ComboBox
                    comboBox5.Items.Add(factoryName);
                }
                else
                {
                    MessageBox.Show("Fabryka o takiej nazwie już istnieje.");
                }

                textBox11.Clear();
            }
            else
            {
                MessageBox.Show("Wprowadź nazwę fabryki.");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (comboBox5.SelectedItem != null)
            {
                string factoryName = comboBox5.SelectedItem.ToString();

                // Удаляем столбец
                if (dataGridView2.Columns.Contains(factoryName))
                {
                    dataGridView2.Columns.Remove(factoryName);
                }

                // Удаляем из ComboBox
                comboBox5.Items.Remove(factoryName);
            }
            else
            {
                MessageBox.Show("Wybierz zakład do usunięcia.");
            }
        }

        private List<Worker> workers = new List<Worker>();
        public class Worker
        {
            public string Imie { get; set; }
            public string Factory { get; set; }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            string name = textBox55.Text.Trim();
            string selectedFactory = comboBox5.SelectedItem?.ToString();

            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(selectedFactory))
            {
                MessageBox.Show("Wpisz imię i wybierz fabrykę.");
                return;
            }

            workers.Add(new Worker { Imie = name, Factory = selectedFactory });
            textBox55.Clear();

            // Показываем, если этот завод выбран
            if (selectedFactory == lastSelectedFactory)
                RefreshWorkers();
        }
        private void RefreshWorkers()
        {
            dataGridView1.Rows.Clear();

            if (string.IsNullOrEmpty(lastSelectedFactory))
                return;

            foreach (var w in workers.Where(w => w.Factory == lastSelectedFactory))
            {
                dataGridView1.Rows.Add(w.Imie);
            }
        }


        private string lastSelectedFactory = null;
        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                string name = dataGridView1.CurrentRow.Cells[0].Value?.ToString();

                var toRemove = workers.FirstOrDefault(w => w.Imie == name && w.Factory == lastSelectedFactory);
                if (toRemove != null)
                {
                    workers.Remove(toRemove);
                    RefreshWorkers();
                }
            }
        }


        private void RefreshWorkers(string factoryFilter)
        {
            dataGridView1.Rows.Clear();

            if (string.IsNullOrEmpty(lastSelectedFactory))
                return;

            foreach (var w in workers.Where(w => w.Factory == lastSelectedFactory))
            {
                dataGridView1.Rows.Add(w.Imie);
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 0)
            {
                string factoryName = dataGridView2.Columns[e.ColumnIndex].HeaderText;
                lastSelectedFactory = factoryName;
                RefreshWorkers(factoryName);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox7.Text))
            {
                MessageBox.Show("Imię i nazwisko są wymagane.");
                return;
            }
            if (currentUser != null)
            {
                SaveUserControls(currentUser);
                SaveUserToFile(currentUser);

                if (currentUser != null)
                {
                    string key = dateTimePicker9.Value.ToString("yyyy-MM");
                    currentUser.MonthlyValues[key] = new MonthData
                    {
                        TextBox17 = textBox17.Text,
                        TextBox19 = textBox19.Text
                    };

                }

                if (currentUser != null)
                {
                    SaveUserControls(currentUser);
                    SaveUserToFile(currentUser);
                    MessageBox.Show($"Dane użytkownika {currentUser.UserName} zachowane.");
                }
                MessageBox.Show($"Dane użytkownika {currentUser.UserName} zachowane.");
            }
            else
            {
                MessageBox.Show("Użytkownik nie został wybrany.");
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex >= 0)
            {
                lastSelectedFactory = dataGridView2.Columns[e.ColumnIndex].HeaderText;
                RefreshWorkers();
            }
        }


        private void AnyInputChanged(object sender, EventArgs e)
        {
            RecalculateAll();
        }
        private void RecalculateAll()
        {
            string cb7 = comboBox7.SelectedItem?.ToString();

            // Вводимые значения
            double.TryParse(textBox10.Text, out double tb10);
            double.TryParse(textBox17.Text, out double tb17);
            double.TryParse(textBox19.Text, out double tb19);
            double.TryParse(textBox24.Text, out double tb24);
            double.TryParse(textBox30.Text, out double tb30);
            double.TryParse(textBox15.Text, out double tb15); // Ввод вручную

            // Всегда считаются
            double tb16 = 42.5 * tb19;
            double tb18 = tb16 * 0.23;
            double tb23 = tb16 + tb18;

            textBox16.Text = tb16.ToString("0.00");
            textBox18.Text = tb18.ToString("0.00");
            textBox23.Text = tb23.ToString("0.00");

            if (cb7 == "Tak")
            {
                // Расчёты при Tak
                double tb29 = tb17 * tb19;
                double tb37 = tb19 * tb17;
                double tb35 = tb19 * tb15 - tb37;

                textBox29.Text = tb29.ToString("0.00");
                textBox37.Text = tb37.ToString("0.00");
                textBox35.Text = tb35.ToString("0.00");

                // Обнуление остальных полей
                textBox36.Text = "0.00";
                textBox39.Text = "0.00";
                textBox12.Text = "0.00";
                textBox22.Text = "0.00";
                textBox14.Text = "0.00";

                return;
            }
            else if (cb7 == "Nie")
            {
                // Расчёты при Nie
                double tb29Nie = tb17 * tb19 - tb10 - tb24 - tb30;
                double tb36 = tb29Nie * 9.76 / 100;
                double tb39 = tb29Nie * 1.5 / 100;
                double tb12 = tb29Nie * 1.67 / 100;

                double tb22 = tb36 + tb39 + tb12;
                double tb14 = tb29Nie - tb22;

                double tb37Nie = tb29Nie - tb22;
                double tb35Nie = (tb19 * tb15) - tb22 - tb29Nie;

                textBox29.Text = tb29Nie.ToString("0.00");
                textBox36.Text = tb36.ToString("0.00");
                textBox39.Text = tb39.ToString("0.00");
                textBox12.Text = tb12.ToString("0.00");
                textBox22.Text = tb22.ToString("0.00");
                textBox14.Text = tb14.ToString("0.00");
                textBox37.Text = tb37Nie.ToString("0.00");
                textBox35.Text = tb35Nie.ToString("0.00");

                return;
            }
        }





        private double originalValue17 = 0;
        private bool combo6DiscountApplied = false;
        private bool userEditingTextBox17 = false;

        private void textBox17_TextChanged(object sender, EventArgs e)
        {
            if (double.TryParse(textBox17.Text, out double value))
            {
                // Сохраняем оригинальное значение только если редактирует пользователь
                if (userEditingTextBox17)
                {
                    originalValue17 = value;
                    combo6DiscountApplied = false; // сбрасываем скидку
                }

                RecalculateAll();
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox6.SelectedItem?.ToString() == "Tak" && !combo6DiscountApplied)
            {
                double discounted = originalValue17 * (1 - 0.0245);
                textBox17.Text = discounted.ToString("0.00");
                combo6DiscountApplied = true;
            }
            else if (comboBox6.SelectedItem?.ToString() == "Nie" && combo6DiscountApplied)
            {
                textBox17.Text = originalValue17.ToString("0.00");
                combo6DiscountApplied = false;
            }

            RecalculateAll();
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
           RecalculateAll();
        }

        private void textBox17_Enter(object sender, EventArgs e)
        {
            userEditingTextBox17 = true;
        }

        private void textBox17_Leave(object sender, EventArgs e)
        {
            userEditingTextBox17 = false;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool show = comboBox3.SelectedItem?.ToString() == "Tak";

            dateTimePicker4.Visible = show;
            dateTimePicker5.Visible = show;
            label8.Visible = show;
            label15.Visible = show;
            label16.Visible = show;

            listBox1.Visible = show;
            listBox2.Visible = show;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool show = comboBox2.SelectedItem?.ToString() == "Tak";

            dateTimePicker6.Visible = show;
            dateTimePicker7.Visible = show;
            label9.Visible = show;
            label17.Visible = show;
            label18.Visible = show;

            listBox3.Visible = show;
            listBox4.Visible = show;
        }
        private bool manualDateSelection = false;

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadFactories(); // если есть
            LoadWorkers();   // ← ВАЖНО
            LoadAllUsers();

            dateTimePicker4.DoubleClick += (s, p) =>
            {
                string dateStr = dateTimePicker4.Value.ToString("dd.MM.yyyy");
                if (!listBox1.Items.Contains(dateStr))
                    listBox1.Items.Add(dateStr);
            };
            dateTimePicker5.DoubleClick += (s, p) =>
            {
                string dateStr = dateTimePicker5.Value.ToString("dd.MM.yyyy");
                if (!listBox2.Items.Contains(dateStr))
                    listBox2.Items.Add(dateStr);
            };

            dateTimePicker6.DoubleClick += (s, p) =>
            {
                string dateStr = dateTimePicker6.Value.ToString("dd.MM.yyyy");
                if (!listBox3.Items.Contains(dateStr))
                    listBox3.Items.Add(dateStr);
            };

            dateTimePicker7.DoubleClick += (s, p) =>
            {
                string dateStr = dateTimePicker7.Value.ToString("dd.MM.yyyy");
                if (!listBox4.Items.Contains(dateStr))
                    listBox4.Items.Add(dateStr);
            };

            dateTimePicker8.DoubleClick += (s, p) =>
            {
                string dateStr = dateTimePicker8.Value.ToString("dd.MM.yyyy");
                if (!listBox5.Items.Contains(dateStr))
                    listBox5.Items.Add(dateStr);
            };
            timerCheckDate3.Interval = 60000; // проверять каждую минуту
            timerCheckDate3.Tick += timerCheckDate3_Tick;

            dateTimePicker1.CustomFormat = "MMMM yyyy";
            dateTimePicker1.ShowUpDown = true;

            dateTimePicker4.DropDown += (s, p) => { calendarOpen = true; lastPickerValue = dateTimePicker4.Value; };
            dateTimePicker4.CloseUp += (s, p) => CheckDateChanged(dateTimePicker4, listBox1);

            dateTimePicker5.DropDown += (s, p) => { calendarOpen = true; lastPickerValue = dateTimePicker5.Value; };
            dateTimePicker5.CloseUp += (s, p) => CheckDateChanged(dateTimePicker5, listBox2);

            dateTimePicker6.DropDown += (s, p) => { calendarOpen = true; lastPickerValue = dateTimePicker6.Value; };
            dateTimePicker6.CloseUp += (s, p) => CheckDateChanged(dateTimePicker6, listBox3);

            dateTimePicker7.DropDown += (s, p) => { calendarOpen = true; lastPickerValue = dateTimePicker7.Value; };
            dateTimePicker7.CloseUp += (s, p) => CheckDateChanged(dateTimePicker7, listBox4);

            dateTimePicker8.DropDown += (s, p) => { calendarOpen = true; lastPickerValue = dateTimePicker8.Value; };
            dateTimePicker8.CloseUp += (s, p) => CheckDateChanged(dateTimePicker8, listBox5);
            comboBox2.SelectedIndexChanged += comboBox2_SelectedIndexChanged;
            comboBox3.SelectedIndexChanged += comboBox3_SelectedIndexChanged;

            dateTimePicker4.Enter += dateTimePicker4_Enter;
            dateTimePicker4.Leave += dateTimePicker4_Leave;
            dateTimePicker4.ValueChanged += dateTimePicker4_ValueChanged;

            dateTimePicker5.Enter += dateTimePicker5_Enter;
            dateTimePicker5.Leave += dateTimePicker5_Leave;
            dateTimePicker5.ValueChanged += dateTimePicker5_ValueChanged;

            dateTimePicker6.Enter += dateTimePicker6_Enter;
            dateTimePicker6.Leave += dateTimePicker6_Leave;
            dateTimePicker6.ValueChanged += dateTimePicker6_ValueChanged;

            dateTimePicker7.Enter += dateTimePicker7_Enter;
            dateTimePicker7.Leave += dateTimePicker7_Leave;
            dateTimePicker7.ValueChanged += dateTimePicker7_ValueChanged;

            dateTimePicker8.Enter += dateTimePicker8_Enter;
            dateTimePicker8.Leave += dateTimePicker8_Leave;
            dateTimePicker8.ValueChanged += dateTimePicker8_ValueChanged;
            dateTimePicker4.MouseDown += (s, p) => manualDateSelection = true;
            dateTimePicker5.MouseDown += (s, p) => manualDateSelection = true;
            dateTimePicker6.MouseDown += (s, p) => manualDateSelection = true;
            dateTimePicker7.MouseDown += (s, p) => manualDateSelection = true;
            dateTimePicker8.MouseDown += (s, p) => manualDateSelection = true;

            listBox5.Visible = true;
            listBox1.Visible = false;
            listBox2.Visible = false;
            listBox3.Visible = false;
            listBox4.Visible = false;
            dateTimePicker4.Visible = false;
            dateTimePicker5.Visible = false;
            dateTimePicker6.Visible = false;
            dateTimePicker7.Visible = false;
            label8.Visible = false;
            label9.Visible = false;

            label15.Visible = false;
            label16.Visible = false;
            label17.Visible = false;
            label18.Visible = false;

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;
            dataGridView1.SelectionChanged += new EventHandler(dataGridView1_SelectionChanged);
            textBox17.TextChanged += WarnIfDotEntered;
            textBox15.TextChanged += WarnIfDotEntered;
            textBox19.TextChanged += WarnIfDotEntered;
            
        }

        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {
            
        }

        private void dateTimePicker5_ValueChanged(object sender, EventArgs e)
        {
            
        }

        private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
        {
            
        }

        private void dateTimePicker7_ValueChanged(object sender, EventArgs e)
        {
            
        }

        private void dateTimePicker8_ValueChanged(object sender, EventArgs e)
        {
            
        }
        
        private string GetSelectedMonthKey()
        {
            return dateTimePicker1.Value.ToString("yyyy-MM"); // Пример: "2025-06"
        }

        public class UserData
        {

            public Dictionary<string, MonthData> MonthlyValues { get; set; } = new Dictionary<string, MonthData>();


            public List<DateTime> DateHistory4 { get; set; } = new List<DateTime>();
            public List<DateTime> DateHistory5 { get; set; } = new List<DateTime>();
            public List<DateTime> DateHistory6 { get; set; } = new List<DateTime>();
            public List<DateTime> DateHistory7 { get; set; } = new List<DateTime>();
            public List<DateTime> DateHistory8 { get; set; } = new List<DateTime>();
            public Dictionary<string, string> TextBoxValues = new Dictionary<string, string>();
            public Dictionary<string, string> ComboBoxValues = new Dictionary<string, string>();
            public Dictionary<string, DateTime> DateTimePickerValues = new Dictionary<string, DateTime>();
            public Dictionary<string, List<string>> ListBoxValues = new Dictionary<string, List<string>>();
            public string UserName; // имя, чтобы связать с dataGridView1   
        }
        private void SaveUserControls(UserData userData)
        {
            foreach (Control ctrl in GetAllControls(this))
            {
                if (ctrl is System.Windows.Forms.TextBox tb)
                    userData.TextBoxValues[tb.Name] = tb.Text;

                else if (ctrl is System.Windows.Forms.ComboBox cb)
                    userData.ComboBoxValues[cb.Name] = cb.SelectedItem != null ? cb.SelectedItem.ToString() : "";

                else if (ctrl is DateTimePicker dtp)
                    userData.DateTimePickerValues[dtp.Name] = dtp.Value;

                else if (ctrl is ListBox lb)
                {
                    List<string> items = new List<string>();
                    foreach (var item in lb.Items)
                        items.Add(item.ToString());

                    userData.ListBoxValues[lb.Name] = items;
                }
            }
        }


        private void LoadUserControls(UserData userData)
        {
            listBox1.Visible = false;
            listBox2.Visible = false;
            listBox3.Visible = false;
            listBox4.Visible = false;

            dateTimePicker4.Visible = false;
            dateTimePicker5.Visible = false;
            dateTimePicker6.Visible = false;
            dateTimePicker7.Visible = false;

            label15.Visible = false;
            label16.Visible = false;
            label17.Visible = false;
            label18.Visible = false;
            foreach (Control ctrl in GetAllControls(this))
            {
                if (ctrl is System.Windows.Forms.TextBox tb && userData.TextBoxValues.ContainsKey(tb.Name))
                    tb.Text = userData.TextBoxValues[tb.Name];

                else if (ctrl is System.Windows.Forms.ComboBox cb && userData.ComboBoxValues.ContainsKey(cb.Name))
                    cb.SelectedItem = userData.ComboBoxValues[cb.Name];

                else if (ctrl is DateTimePicker dtp && userData.DateTimePickerValues.ContainsKey(dtp.Name))
                    dtp.Value = userData.DateTimePickerValues[dtp.Name];

                else if (ctrl is ListBox lb && userData.ListBoxValues.ContainsKey(lb.Name))
                {
                    lb.Items.Clear();
                    foreach (string item in userData.ListBoxValues[lb.Name])
                        lb.Items.Add(item);
                }
                lastDateTimePicker9Key = dateTimePicker9.Value.ToString("yyyy-MM");
                if (currentUser.MonthlyValues.TryGetValue(lastDateTimePicker9Key, out var values))
                {
                    textBox17.Text = values.TextBox17;
                    textBox19.Text = values.TextBox19;
                }
                else
                {
                    textBox17.Text = "";
                    textBox19.Text = "";
                }
                // Принудительно вызываем события для обновления интерфейса
                comboBox3_SelectedIndexChanged(comboBox3, EventArgs.Empty);
                comboBox2_SelectedIndexChanged(comboBox2, EventArgs.Empty);
                comboBox4_SelectedIndexChanged(comboBox4, EventArgs.Empty);
                comboBox6_SelectedIndexChanged(comboBox6, EventArgs.Empty);
                comboBox7_SelectedIndexChanged(comboBox7, EventArgs.Empty);

            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0) return;

            string name = dataGridView1.SelectedRows[0].Cells[0].Value?.ToString();
            if (string.IsNullOrWhiteSpace(name)) return;

            // Сохраняем предыдущего пользователя
            if (currentUser != null)
            {
                SaveUserControls(currentUser);
                SaveUserToFile(currentUser);
            }

            // Проверка: есть ли уже файл для этого пользователя
            var loaded = LoadUserFromFile(name);
            if (loaded != null)
            {
                currentUser = loaded;
                if (!users.Any(u => u.UserName == name))
                    users.Add(currentUser);

                LoadUserControls(currentUser);
                LoadUserPhotos();
                LoadUserFiles();
            }
            else
            {
                // Новый пользователь — создаём пустой
                CreateNewUser(name);
            }
        }

        private IEnumerable<Control> GetAllControls(Control parent)
        {
            foreach (Control child in parent.Controls)
            {
                yield return child;

                foreach (Control grandChild in GetAllControls(child))
                    yield return grandChild;
            }
        }
        private string GetUserFilePath(string userName)
        {
            string folderPath = Path.Combine(Application.StartupPath, "UserData");
            Directory.CreateDirectory(folderPath); // Создаёт папку при необходимости
            return Path.Combine(folderPath, $"{userName}.json");
        }

        private void SaveUserToFile(UserData user)
        {
            string json = JsonConvert.SerializeObject(user, Formatting.Indented);
            File.WriteAllText(GetUserFilePath(user.UserName), json);
        }

        private UserData LoadUserFromFile(string userName)
        {
            string path = GetUserFilePath(userName);
            if (File.Exists(path))
            {
                string json = File.ReadAllText(path);
                return JsonConvert.DeserializeObject<UserData>(json);
            }
            return null;
        }
        

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox7.Text))
            {
                MessageBox.Show("Imię i nazwisko są wymagane.");
                return;
            }
            GenerateWordDocument("Aneks_do_umowy.docx", "Aneks_" + textBox1.Text + ".docx");
        }

        private void button7_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox7.Text))
            {
                MessageBox.Show("Imię i nazwisko są wymagane.");
                return;
            }
            GenerateWordDocument("umowa_zlecenia.docx", "Umowa_" + textBox1.Text + ".docx");
            GenerateUserInfoTxt();
        }


        private void GenerateWordDocument(string templateName, string outputName)
        {
            var wordApp = new Word.Application();
            string exeFolder = AppDomain.CurrentDomain.BaseDirectory;
            string templatePath = Path.Combine(exeFolder, templateName);

            string fullName = textBox1.Text + " " + textBox7.Text;
            string folderPath = Path.Combine(@"D:\Histora Umowy", fullName);
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            string savePath = Path.Combine(folderPath, outputName);

            var document = wordApp.Documents.Open(templatePath);

            try
            {
                // Форматы
                string nowDate = DateTime.Now.ToString("dd.MM.yyyy");
                string dateFrom = dateTimePicker1.Value.ToString("dd.MM.yyyy");
                string dateTo = dateTimePicker2.Value.ToString("dd.MM.yyyy");
                string rate = textBox17.Text;
                string fullNameText = textBox1.Text + " " + textBox7.Text;
                string passport = textBox4.Text;
                string birthDate = textBox9.Text;
                string region = comboBox4.Text;

                // Общие замены
                ReplaceWordText(document, "Synytsin Volodymyr", fullNameText);
                ReplaceWordText(document, "Prytula Larysa", fullNameText);
                ReplaceWordText(document, "Larysa Prytula", fullNameText);
                ReplaceWordText(document, "FU123547", passport);
                ReplaceWordText(document, "FH869263", passport);
                ReplaceWordText(document, "27.08.1983", birthDate);
                ReplaceWordText(document, "22.11.1965", birthDate);
                ReplaceWordText(document, "Ukraina, Obwód:", region);
                ReplaceWordText(document, "31,50", rate);

                // Уникальные замены по шаблону
                if (templateName.Contains("Aneks"))
                {
                    ReplaceWordText(document, "21.04.2025", nowDate);     // текущая дата
                    ReplaceWordText(document, "21.03.2025", dateFrom);    // дата начала
                    ReplaceWordText(document, "30.09.2025", dateTo);      // дата окончания
                }
                else if (templateName.Contains("umowa"))
                {

                    ReplaceWordText(document, "21.03.2025", dateFrom);
                    ReplaceWordText(document, "30.09.2025", dateTo);
                }

                document.SaveAs2(savePath);
                MessageBox.Show("Plik został utworzony: " + savePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd: " + ex.Message);
            }
            finally
            {
                document.Close(false);
                wordApp.Quit(false);
            }
        }


        // Метод замены текста
        private void ReplaceWordText(Word.Document doc, string findText, string replaceText)
        {
            Word.Find find = doc.Content.Find;
            find.ClearFormatting();
            find.Text = findText;
            find.Replacement.ClearFormatting();
            find.Replacement.Text = replaceText;

            object replaceAll = Word.WdReplace.wdReplaceAll;
            find.Execute(Replace: ref replaceAll);
        }

        private void buttonLoadPhotos_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Изображения (*.jpg;*.png)|*.jpg;*.png";
                ofd.Multiselect = true;

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string fullName = textBox1.Text + " " + textBox7.Text;
                    string targetFolder = Path.Combine(@"D:\Histora Umowy", fullName);

                    if (!Directory.Exists(targetFolder))
                        Directory.CreateDirectory(targetFolder);

                    int index = 1;
                    foreach (string file in ofd.FileNames)
                    {
                        string extension = Path.GetExtension(file);
                        string targetPath = Path.Combine(targetFolder, $"photo_{index}{extension}");
                        File.Copy(file, targetPath, true);
                        index++;
                    }

                    MessageBox.Show("Zdjęcia zostały zapisane.");
                }
            }
        }

        private void buttonLoadFiles_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Документы|*.pdf;*.docx;*.xlsx;*.txt";
                ofd.Multiselect = true;

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string fullName = textBox1.Text + " " + textBox7.Text;
                    string targetFolder = Path.Combine(@"D:\Histora Umowy", fullName);

                    if (!Directory.Exists(targetFolder))
                        Directory.CreateDirectory(targetFolder);

                    foreach (string file in ofd.FileNames)
                    {
                        string filename = Path.GetFileName(file);
                        string targetPath = Path.Combine(targetFolder, filename);
                        File.Copy(file, targetPath, true);
                    }

                    MessageBox.Show("Pliki zostały pomyślnie przesłane.");
                }
            }
        }

        private void GenerateUserInfoTxt()
        {
            string fullName = textBox1.Text + " " + textBox7.Text;
            string startDate = dateTimePicker1.Value.ToString("dd.MM.yyyy");
            string endDate = dateTimePicker2.Value.ToString("dd.MM.yyyy");

            string content = $"{fullName}:\n\nDate od: {startDate}\nDate do: {endDate}\n\nAneks: date {startDate} - {endDate}";

            string folderPath = Path.Combine(@"D:\Histora Umowy", fullName);
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            string filePath = Path.Combine(folderPath, "umowa_info.txt");
            File.WriteAllText(filePath, content);
        }
        private void LoadUserPhotos()
        {
            string fullName = textBox1.Text + " " + textBox7.Text;
            string folderPath = Path.Combine(@"D:\Histora Umowy", fullName);

            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            photoPaths = Directory.GetFiles(folderPath, "*.*")
                                  .Where(f => f.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) || f.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
                                  .ToList();

            currentPhotoIndex = 0;
            ShowCurrentPhoto();
        }
        private void ShowCurrentPhoto()
        {
            if (photoPaths.Count == 0)
            {
                pictureBoxPhoto.Image = null;
                return;
            }

            try
            {
                // Освобождаем старое изображение, если есть
                if (pictureBoxPhoto.Image != null)
                    pictureBoxPhoto.Image.Dispose();

                pictureBoxPhoto.Image = Image.FromFile(photoPaths[currentPhotoIndex]);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Błąd ładowania zdjęcia: " + ex.Message);
            }
        }

        private void buttonPrevPhoto_Click(object sender, EventArgs e)
        {
            if (photoPaths.Count == 0) return;

            currentPhotoIndex = (currentPhotoIndex - 1 + photoPaths.Count) % photoPaths.Count;
            ShowCurrentPhoto();
        }

        private void buttonNextPhoto_Click(object sender, EventArgs e)
        {
            if (photoPaths.Count == 0) return;

            currentPhotoIndex = (currentPhotoIndex + 1) % photoPaths.Count;
            ShowCurrentPhoto();
        }
        private void LoadUserFiles()
        {
            string fullName = textBox1.Text + " " + textBox7.Text;
            string folderPath = Path.Combine(@"D:\Histora Umowy", fullName);

            if (!Directory.Exists(folderPath)) return;

            var files = Directory.GetFiles(folderPath)
                                 .Where(f => !f.EndsWith(".jpg") && !f.EndsWith(".png")) // исключаем фото
                                 .ToList();

            listBoxFiles.Items.Clear();
            foreach (var file in files)
            {
                listBoxFiles.Items.Add(Path.GetFileName(file));
            }
        }

        private void listBox6_DoubleClick(object sender, EventArgs e)
        {
            if (listBoxFiles.SelectedItem == null) return;

            string fullName = textBox1.Text + " " + textBox7.Text;
            string folderPath = Path.Combine(@"D:\Histora Umowy", fullName);
            string fileName = listBoxFiles.SelectedItem.ToString();
            string fullPath = Path.Combine(folderPath, fileName);

            if (File.Exists(fullPath))
                System.Diagnostics.Process.Start(fullPath);
        }

        private void buttonGenerateInfoTxt_Click(object sender, EventArgs e)
        {
            string fullName = textBox1.Text + " " + textBox7.Text;
            string startDate = dateTimePicker1.Value.ToString("dd.MM.yyyy");
            string endDate = dateTimePicker2.Value.ToString("dd.MM.yyyy");

            string content = $"{fullName}:\n\nDate od: {startDate}\nDate do: {endDate}\n\nAneks: date {startDate} - {endDate}";

            string folderPath = Path.Combine(@"D:\Histora Umowy", fullName);
            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            string filePath = Path.Combine(folderPath, "umowa_info.txt");
            File.WriteAllText(filePath, content);

            MessageBox.Show("Informacja została zapisana w TXT.");
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selection = comboBox8.SelectedItem?.ToString();

            if (selection == "Pracuje")
            {
                this.BackColor = Color.DarkSeaGreen;
            }
            else if (selection == "Nie Pracuje")
            {
                this.BackColor = Color.IndianRed;
            }
            else
            {
                this.BackColor = SystemColors.Control; // Цвет по умолчанию
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveWorkers();
            SaveAllUsers();
            SaveFactories();
            UploadAppDataToDrive();
            UploadCurrentUserFolderToDrive();

            var factoryNames = dataGridView2.Columns
        .Cast<DataGridViewColumn>()
        .Select(col => col.HeaderText)
        .ToList();

            string savePath = Path.Combine(Application.StartupPath, "factories.json");
            File.WriteAllText(savePath, JsonConvert.SerializeObject(factoryNames, Formatting.Indented));
        }
        
        private DriveService GetDriveService()
        {
            string[] Scopes = { DriveService.Scope.DriveFile };
            string appName = "MyDriveUploader";

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = Path.Combine(Application.StartupPath, "token.json");

                var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

                return new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = appName
                });
            }
        }

        private void UploadFileToDrive(string localFilePath)
        {
            var service = GetDriveService(); // см. ниже

    var fileMeta = new Google.Apis.Drive.v3.Data.File
    {
        Name = Path.GetFileName(localFilePath)
    };

    using (var stream = new FileStream(localFilePath, FileMode.Open))
    {
        var request = service.Files.Create(fileMeta, stream, "application/octet-stream");
        request.Fields = "id";
        var result = request.Upload();

        if (result.Status != UploadStatus.Completed)
        {
            MessageBox.Show("Błąd ładowania: " + result.Exception?.Message);
        }
    }
        }
        private void SaveAllUsers()
        {
            string path = Path.Combine(Application.StartupPath, "users.json");
            string json = JsonConvert.SerializeObject(users, Formatting.Indented);
            File.WriteAllText(path, json);
        }
        private void LoadAllUsers()
        {
            string path = Path.Combine(Application.StartupPath, "users.json");
            if (!File.Exists(path)) return;

            string json = File.ReadAllText(path);
            users = JsonConvert.DeserializeObject<List<UserData>>(json);

            foreach (var user in users)
            {
                dataGridView1.Rows.Add(user.UserName);
            }
        }
        private void SaveFactories()
        {
            var factoryNames = dataGridView2.Columns
         .Cast<DataGridViewColumn>()
         .Select(col => col.HeaderText)
         .ToList();

            string path = Path.Combine(Application.StartupPath, "factories.json");
            File.WriteAllText(path, JsonConvert.SerializeObject(factoryNames, Formatting.Indented));
        }
        private void LoadFactories()
        {
            string path = Path.Combine(Application.StartupPath, "factories.json");
            if (!File.Exists(path)) return;

            var factoryNames = JsonConvert.DeserializeObject<List<string>>(File.ReadAllText(path));

            foreach (string name in factoryNames)
            {
                if (!dataGridView2.Columns.Contains(name))
                {
                    dataGridView2.Columns.Add(name, name);
                    comboBox5.Items.Add(name);
                }
            }
        }
        private void UploadCurrentUserFolderToDrive()
        {
            // Проверка, есть ли текущий пользователь
            if (currentUser == null || string.IsNullOrWhiteSpace(currentUser.UserName))
            {
                MessageBox.Show("Nie wybrano użytkownika do przesłania na Google Dysk.");
                return;
            }

            // Путь к папке пользователя
            string fullName = currentUser.UserName;
            string folderPath = Path.Combine(@"D:\Histora Umowy", fullName);

            // Проверка папки
            if (!Directory.Exists(folderPath))
            {
                
                return;
            }

            // Получаем все файлы в папке
            string[] files = Directory.GetFiles(folderPath);
            if (files.Length == 0)
            {
                MessageBox.Show("Nie ma plików do przesłania na Google Dysk.");
                return;
            }

            // Загружаем каждый файл
            foreach (string file in files)
            {
                try
                {
                    UploadFileToDrive(file); // метод, который ты уже реализовал
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Błąd podczas ładowania pliku {Path.GetFileName(file)}:\n{ex.Message}");
                }
            }
        }

        private void UploadAppDataToDrive()
        {
            string[] filesToUpload = { "users.json", "factories.json" };

            foreach (var file in filesToUpload)
            {
                string fullPath = Path.Combine(Application.StartupPath, file);
                if (File.Exists(fullPath))
                {
                    try
                    {
                        UploadFileToDrive(fullPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Błąd podczas ładowania pliku {file}:\n{ex.Message}");
                    }
                }
            }
        }

        private void dateTimePicker4_Leave(object sender, EventArgs e)
        {
            isUserChangingDate4 = false;   
        }

        private void dateTimePicker4_Enter(object sender, EventArgs e)
        {
            isUserChangingDate4 = true;
        }

        private void dateTimePicker5_Enter(object sender, EventArgs e)
        {
            isUserChangingDate5 = true;
        }

        private void dateTimePicker5_Leave(object sender, EventArgs e)
        {
            isUserChangingDate5 = false;
        }

        private void dateTimePicker6_Leave(object sender, EventArgs e)
        {
            isUserChangingDate6 = false;
        }

        private void dateTimePicker6_Enter(object sender, EventArgs e)
        {
            isUserChangingDate6 = true;
        }

        private void dateTimePicker7_Leave(object sender, EventArgs e)
        {
            isUserChangingDate7 = false;

        }

        private void dateTimePicker7_Enter(object sender, EventArgs e)
        {
            isUserChangingDate7 = true;

        }

        private void dateTimePicker8_Enter(object sender, EventArgs e)
        {
            isUserChangingDate8 = true;

        }

        private void dateTimePicker8_Leave(object sender, EventArgs e)
        {
            isUserChangingDate8 = false;

        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
                listBox1.Items.Remove(listBox1.SelectedItem);
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox2_DoubleClick(object sender, EventArgs e)
        {
            if (listBox2.SelectedItem != null)
                listBox2.Items.Remove(listBox2.SelectedItem);
        }
        private string lastDate4 = "";
        private string lastDate5 = "";
        private string lastDate6 = "";
        private string lastDate7 = "";
        private string lastDate8 = "";
        private void listBox3_DoubleClick(object sender, EventArgs e)
        {
            if (listBox3.SelectedItem != null)
                listBox3.Items.Remove(listBox3.SelectedItem);
        }

        private void listBox4_DoubleClick(object sender, EventArgs e)
        {
            if (listBox4.SelectedItem != null)
                listBox4.Items.Remove(listBox4.SelectedItem);
        }

        private void listBox5_DoubleClick(object sender, EventArgs e)
        {
            if (listBox5.SelectedItem != null)
                listBox5.Items.Remove(listBox5.SelectedItem);
        }

        private void CheckDateChanged(DateTimePicker picker, ListBox listBox)
        {
            calendarOpen = false;

            if (picker.Value != lastPickerValue)
            {
                string dateStr = picker.Value.ToString("dd.MM.yyyy");
                if (!listBox.Items.Contains(dateStr))
                    listBox.Items.Add(dateStr);
            }
        }

        private void textBox24_TextChanged(object sender, EventArgs e)
        {

        }

        private void label37_Click(object sender, EventArgs e)
        {

        }

        private void label36_Click(object sender, EventArgs e)
        {

        }

        private void textBox30_TextChanged(object sender, EventArgs e)
        {

        }

        private void label43_Click(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }
        private void SaveWorkers()
        {
            string path = Path.Combine(Application.StartupPath, "workers.json");
            File.WriteAllText(path, JsonConvert.SerializeObject(workers, Formatting.Indented));
        }

        private void LoadWorkers()
        {
            string path = Path.Combine(Application.StartupPath, "workers.json");
            if (File.Exists(path))
            {
                workers = JsonConvert.DeserializeObject<List<Worker>>(File.ReadAllText(path));
            }
        }


        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox5.SelectedItem != null)
            {
                string selectedFactory = comboBox5.SelectedItem.ToString();
                RefreshWorkers(selectedFactory);
            }
        }

        private void textBox37_TextChanged(object sender, EventArgs e)
        {

        }
        
        private void WarnIfDotEntered(object sender, EventArgs e)
        {
            var tb = sender as System.Windows.Forms.TextBox;
            if (tb != null && tb.Text.Contains("."))
            {
                MessageBox.Show("Nie używaj kropki. Proszę używać przecinka jako separatora dziesiętnego.", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tb.Text = tb.Text.Replace(".", ""); // можно также: tb.Clear();
            }
        }
        private bool isNewUser = false;
        private void CreateNewUser(string name)
        {
            currentUser = new UserData { UserName = name };
            users.Add(currentUser);
            ClearFormForNewUser(); // очищаем форму
            bool alreadyExists = dataGridView1.Rows
            .Cast<DataGridViewRow>()
            .Any(row => row.Cells[0].Value?.ToString() == name);

            if (!alreadyExists)
            {
                dataGridView1.Rows.Add(name);
            }
        }
        private void ClearFormForNewUser()
        {
            foreach (Control ctrl in GetAllControls(this))
            {
                if (ctrl is TextBox tb) tb.Text = "";
                else if (ctrl is ComboBox cb) cb.SelectedIndex = -1;
                else if (ctrl is ListBox lb) lb.Items.Clear();
                else if (ctrl is DateTimePicker dtp) dtp.Value = DateTime.Now;
            }

            // Скрыть те, что надо:
            listBox1.Visible = false;
            listBox2.Visible = false;
            listBox3.Visible = false;
            listBox4.Visible = false;
            label15.Visible = false;
            label16.Visible = false;
            label17.Visible = false;
            label18.Visible = false;
            dateTimePicker4.Visible = false;
            dateTimePicker5.Visible = false;
            dateTimePicker6.Visible = false;
            dateTimePicker7.Visible = false;
        }
        private string lastDateTimePicker9Key = "";

        private void dateTimePicker9_ValueChanged(object sender, EventArgs e)
        {
            if (currentUser == null) return;

            // Сохраняем значения для предыдущего месяца
            if (!string.IsNullOrEmpty(lastDateTimePicker9Key))
            {
                currentUser.MonthlyValues[lastDateTimePicker9Key] = new MonthData
                {
                    TextBox17 = textBox17.Text,
                    TextBox19 = textBox19.Text
                };

            }

            // Загружаем значения для нового месяца
            string newKey = dateTimePicker9.Value.ToString("yyyy-MM");
            lastDateTimePicker9Key = newKey;

            if (currentUser.MonthlyValues.TryGetValue(newKey, out var values))
            {
                textBox17.Text = values.TextBox17;
                textBox19.Text = values.TextBox19;

            }
            else
            {
                textBox17.Text = "";
                textBox19.Text = "";
            }
        }
        private void LoadMonthlyValues()
        {
            if (currentUser == null) return;

            string monthKey = GetSelectedMonthKey();

            if (currentUser.MonthlyValues.TryGetValue(monthKey, out var data))
            {
                textBox17.Text = data.TextBox17;
                textBox19.Text = data.TextBox19;
            }
            else
            {
                textBox17.Text = "";
                textBox19.Text = "";
            }
        }
        public class MonthData
        {
            public string TextBox17 { get; set; }
            public string TextBox19 { get; set; }
        }

        private void label26_Click(object sender, EventArgs e)
        {

        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {

        }

        private void textBox26_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox9.SelectedItem?.ToString() == "Tak")
            {
                timerCheckDate3.Start();
            }
            else
            {
                timerCheckDate3.Stop();
            }
        }

        private void timerCheckDate3_Tick(object sender, EventArgs e)
        {
            if (comboBox9.SelectedItem?.ToString() != "Tak") return;

            DateTime targetDate = dateTimePicker3.Value.Date;
            DateTime now = DateTime.Now.Date;

            if (now >= targetDate)
            {
                textBox17.Text = textBox26.Text;
                timerCheckDate3.Stop(); // Чтобы больше не срабатывал
            }
        }
    }
}
