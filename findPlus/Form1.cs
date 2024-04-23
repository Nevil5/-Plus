using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace findPlus
{
    public partial class Form1 : Form
    {
        private OleDbConnection connection;
        private const string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=F:\ДДМА\3 Курс\2а семестр\ОБДЗ Курсова робота\АІС для роботи бюро знахідок\findPlus\findPlus\Bureau of finds.accdb";

        public Form1()
        {
            connection = new OleDbConnection(connectionString);
            InitializeComponent();
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "bureau_of_findsDataSet.Контакти_клієнтів". При необходимости она может быть перемещена или удалена.
            this.контакти_клієнтівTableAdapter.Fill(this.bureau_of_findsDataSet.Контакти_клієнтів);
            LoadData(); // Завантаження даних з бази даних в dataGridView1
            FillComboBox3();
        }
        private void LoadData()
        {
            try
            {
                // Завантаження даних у dataGridView1 з таблиці "Втрачені речі"
                using (OleDbDataAdapter lostItemsAdapter = new OleDbDataAdapter("SELECT * FROM [Втрачені речі] WHERE [Дата знаходження втраченої речі] IS NOT NULL", connection))
                {
                    DataTable lostItemsTable = new DataTable();
                    lostItemsAdapter.Fill(lostItemsTable);
                    dataGridView1.DataSource = lostItemsTable;
                }

                // Завантаження даних у dataGridView2 з таблиці "Контакти клієнтів"
                using (OleDbDataAdapter contactsAdapter = new OleDbDataAdapter("SELECT * FROM [Контакти клієнтів]", connection))
                {
                    DataTable contactsTable = new DataTable();
                    contactsAdapter.Fill(contactsTable);
                    dataGridView2.DataSource = contactsTable;
                    dataGridView2.Columns["Код"].Visible = false;
                }
                // Завантаження даних у dataGridView3 з таблиці "Аукціон"
                using (OleDbDataAdapter contactsAdapter = new OleDbDataAdapter("SELECT * FROM [Аукціон]", connection))
                {
                    DataTable contactsTable = new DataTable();
                    contactsAdapter.Fill(contactsTable);
                    dataGridView3.DataSource = contactsTable;
                }
                // Встановлення автоматичного розширення стовпців для всіх DataGridView
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
        }

        // Очищення полів вводу
        private void ClearInputFields()
        {
            textBox2.Clear();
            textBox3.Clear();
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
        }

        private void FillComboBox3()
        {
            comboBox3.Items.Clear();
            // Заповнення комбо бокса комірками з діапазону
            string[] allCells = new string[] { "A01", "A02", "A03", "A04", "A05", "B01", "B02", "B03", "B04", "B05", "C01", "C02", "C03", "C04", "C05", "D01", "D02", "D03", "D04", "D05", "E01", "E02", "E03", "E04", "E05" };
            comboBox3.Items.AddRange(allCells);

            try
            {
                connection.Open();
                // Виконання запиту до бази даних, щоб отримати список зайнятих комірок
                string occupiedCellsQuery = "SELECT [Номер комірки зберігання] FROM [Втрачені речі]";
                using (OleDbCommand occupiedCellsCommand = new OleDbCommand(occupiedCellsQuery, connection))
                using (OleDbDataReader reader = occupiedCellsCommand.ExecuteReader())
                {
                    List<string> occupiedCells = new List<string>();
                    while (reader.Read())
                    {
                        occupiedCells.Add(reader["Номер комірки зберігання"].ToString());
                    }

                    // Вилучення зайнятих комірок зі списку
                    foreach (string cell in occupiedCells)
                    {
                        if (comboBox3.Items.Contains(cell))
                        {
                            comboBox3.Items.Remove(cell);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                // Занесення даних у таблицю Втрачені речі
                string insertLostItemsQuery = "INSERT INTO [Втрачені речі] ([Дата знаходження втраченої речі], [Локація знахідки], [Опис речі], [Особисті прикмети], [Тип речі], [Номер комірки зберігання]) VALUES (?, ?, ?, ?, ?, ?)";
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                using (OleDbCommand insertLostItemsCommand = new OleDbCommand(insertLostItemsQuery, connection))
                {
                    connection.Open();
                    // Отримання дати з dateTimePicker2, обрізаючи час
                    DateTime selectedDate = dateTimePicker2.Value.Date;

                    insertLostItemsCommand.Parameters.AddWithValue("@p1", selectedDate); // Дата знаходження втраченої речі
                    insertLostItemsCommand.Parameters.AddWithValue("@p2", comboBox1.SelectedItem.ToString()); // Локація знаходки
                    insertLostItemsCommand.Parameters.AddWithValue("@p3", textBox2.Text); // Опис речі
                    insertLostItemsCommand.Parameters.AddWithValue("@p4", textBox3.Text); // Особисті прикмети
                    insertLostItemsCommand.Parameters.AddWithValue("@p5", comboBox2.SelectedItem.ToString()); // Тип речі
                    insertLostItemsCommand.Parameters.AddWithValue("@p6", comboBox3.SelectedItem.ToString()); // Номер комірки зберігання

                    insertLostItemsCommand.ExecuteNonQuery();
                    MessageBox.Show("Дані успішно занесені до бази даних!");
                }

                // Після додавання запису оновити вміст таблиці
                LoadData();

                // Після додавання запису оновити вміст comboBox3
                FillComboBox3();

                // Очищення полів
                ClearInputFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
        }
        private void button2_Click_1(object sender, EventArgs e)
        {
            // Очистка полів
            ClearInputFields();
        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            // Перевірка, чи вибрано хоча б один рядок для видалення
            if (dataGridView1.SelectedRows.Count > 0)
            {
                // Підтвердження видалення
                DialogResult result = MessageBox.Show("Ви впевнені, що хочете видалити вибраний рядок?", "Підтвердження видалення", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        // Використовую окреме з'єднання для видалення
                        using (OleDbConnection tempConnection = new OleDbConnection(connectionString))
                        {
                            tempConnection.Open();

                            // Отримання ID вибраного рядка
                            int selectedRowIndex = dataGridView1.SelectedRows[0].Index;
                            int selectedItemId = Convert.ToInt32(dataGridView1.Rows[selectedRowIndex].Cells["Ідентифікаційний номер"].Value);

                            // Видалення рядка з бази даних
                            string deleteQuery = "DELETE FROM [Втрачені речі] WHERE [Ідентифікаційний номер] = ?";
                            using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, tempConnection))
                            {
                                deleteCommand.Parameters.AddWithValue("@p1", selectedItemId);
                                int rowsAffected = deleteCommand.ExecuteNonQuery();

                                if (rowsAffected > 0)
                                {
                                    MessageBox.Show("Рядок успішно видалено!");
                                    // Оновлення відображення таблиці
                                    LoadData(); // Оновлення вмісту таблиці
                                    FillComboBox3();
                                }
                                else
                                {
                                    MessageBox.Show("Не вдалося видалити рядок.");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Помилка: " + ex.Message);
                    }
                }
            }
            else
            {
                MessageBox.Show("Будь ласка, виберіть рядок для видалення.");
            }
        }
        private void buttonClear_Click_1(object sender, EventArgs e)
        {
            // Очищення полів
            textBox4.Clear();
            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            dateTimePicker1.Checked = false;

            // Очищення результатів пошуку та відображення повного вмісту таблиці
            LoadData();
            FillComboBox3();
        }

        private void SearchLostItems(string descriptionKeyword, string locationKeyword, string typeKeyword, DateTime? date)
        {
            try
            {
                using (OleDbCommand searchCommand = connection.CreateCommand())
                {
                    StringBuilder queryBuilder = new StringBuilder("SELECT * FROM [Втрачені речі] WHERE 1=1");

                    // Пошук за допомогою поля "Локація знахідки"
                    if (!string.IsNullOrEmpty(locationKeyword))
                    {
                        queryBuilder.Append(" AND ([Локація знахідки] = ?)");
                        searchCommand.Parameters.AddWithValue("@pLocation", locationKeyword);
                    }

                    // Пошук за допомогою поля "Тип речі"
                    if (!string.IsNullOrEmpty(typeKeyword))
                    {
                        queryBuilder.Append(" AND ([Тип речі] = ?)");
                        searchCommand.Parameters.AddWithValue("@pType", typeKeyword);
                    }

                    // Пошук за допомогою поля "Дата знаходження втраченої речі"
                    if (date.HasValue)
                    {
                        queryBuilder.Append(" AND ([Дата знаходження втраченої речі] = ?)");
                        searchCommand.Parameters.AddWithValue("@pDate", date.Value);
                    }

                    // Пошук за допомогою поля "Опис речі"
                    if (!string.IsNullOrEmpty(descriptionKeyword))
                    {
                        queryBuilder.Append(" AND (UCASE([Опис речі]) LIKE ?)");
                        searchCommand.Parameters.AddWithValue("@pDescription", "%" + descriptionKeyword.ToUpper() + "%");
                    }

                    // Пошук за допомогою поля "Особисті прикмети"
                    if (!string.IsNullOrEmpty(descriptionKeyword))
                    {
                        queryBuilder.Append(" OR (UCASE([Особисті прикмети]) LIKE ?)");
                        searchCommand.Parameters.AddWithValue("@pDescription", "%" + descriptionKeyword.ToUpper() + "%");
                    }

                    searchCommand.CommandText = queryBuilder.ToString();

                    OleDbDataAdapter searchAdapter = new OleDbDataAdapter(searchCommand);
                    DataTable searchTable = new DataTable();
                    searchAdapter.Fill(searchTable);

                    dataGridView1.DataSource = searchTable;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка під час пошуку: " + ex.Message);
            }
        }
        // Виконання пошуку
        private void buttonSearch_Click_1(object sender, EventArgs e)
        {
            string descriptionKeyword = textBox4.Text;
            string locationKeyword = comboBox4.SelectedItem?.ToString();
            string typeKeyword = comboBox5.SelectedItem?.ToString();
            DateTime? date = null;

            // Перевіряємо, чи доступний dateTimePicker1
            if (this.dateTimePicker1 != null && this.dateTimePicker1.Checked)
            {
                date = dateTimePicker1.Value.Date;
            }

            SearchLostItems(descriptionKeyword, locationKeyword, typeKeyword, date);
        }
        private void button4_Click_1(object sender, EventArgs e)
        {
            try
            {
                // Занесення даних у таблицю "Контакти клієнтів"
                string insertContactQuery = "INSERT INTO [Контакти клієнтів] ([ПІБ клієнта], [Номер телефону], [Номер документу], [Опис документу]) VALUES (?, ?, ?, ?)";
                using (OleDbConnection contactsConnection = new OleDbConnection(connectionString))
                using (OleDbCommand insertContactCommand = new OleDbCommand(insertContactQuery, contactsConnection))
                {
                    contactsConnection.Open();

                    insertContactCommand.Parameters.AddWithValue("@p1", textBox6.Text); // ПІБ клієнта
                    insertContactCommand.Parameters.AddWithValue("@p2", textBox7.Text); // Номер телефону
                    insertContactCommand.Parameters.AddWithValue("@p3", textBox8.Text); // Номер документу
                    insertContactCommand.Parameters.AddWithValue("@p4", textBox1.Text); // Опис документу

                    insertContactCommand.ExecuteNonQuery();
                    MessageBox.Show("Дані успішно занесені до бази даних!");
                }

                // Після додавання запису оновити вміст dataGridView2
                LoadData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
        }
        private void button5_Click_1(object sender, EventArgs e)
        {
            try
            {
                // Формування SQL-запиту для пошуку
                string searchQuery = "SELECT * FROM [Контакти клієнтів] WHERE ";
                List<string> conditions = new List<string>();

                if (!string.IsNullOrEmpty(textBox6.Text))
                    conditions.Add("[ПІБ клієнта] LIKE '%" + textBox6.Text + "%'");

                if (!string.IsNullOrEmpty(textBox7.Text))
                    conditions.Add("[Номер телефону] LIKE '%" + textBox7.Text + "%'");

                if (!string.IsNullOrEmpty(textBox8.Text))
                    conditions.Add("[Номер документу] LIKE '%" + textBox8.Text + "%'");

                if (!string.IsNullOrEmpty(textBox1.Text))
                    conditions.Add("[Опис документу] LIKE '%" + textBox1.Text + "%'");

                if (conditions.Count > 0)
                {
                    searchQuery += string.Join(" AND ", conditions);
                    OleDbDataAdapter searchAdapter = new OleDbDataAdapter(searchQuery, connection);
                    DataTable searchTable = new DataTable();
                    searchAdapter.Fill(searchTable);
                    dataGridView2.DataSource = searchTable;
                }
                else
                {
                    MessageBox.Show("Будь ласка, введіть критерії пошуку.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка під час пошуку: " + ex.Message);
            }
        }
        private void button3_Click_1(object sender, EventArgs e)
        {
            // Перевірка, чи вибрано хоча б один рядок для видалення
            if (dataGridView2.SelectedRows.Count > 0)
            {
                // Підтвердження видалення
                DialogResult result = MessageBox.Show("Ви впевнені, що хочете видалити вибраний рядок?", "Підтвердження видалення", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        // Використання окремого з'єднання для видалення
                        using (OleDbConnection tempConnection = new OleDbConnection(connectionString))
                        {
                            tempConnection.Open();

                            // Отримання ID вибраного рядка
                            int selectedRowIndex = dataGridView2.SelectedRows[0].Index;
                            int selectedItemId = Convert.ToInt32(dataGridView2.Rows[selectedRowIndex].Cells["Код"].Value);

                            // Видалення рядка з бази даних
                            string deleteQuery = "DELETE FROM [Контакти клієнтів] WHERE ([Код]) = ?";
                            using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, tempConnection))
                            {
                                deleteCommand.Parameters.AddWithValue("@p1", selectedItemId);
                                int rowsAffected = deleteCommand.ExecuteNonQuery();

                                if (rowsAffected > 0)
                                {
                                    MessageBox.Show("Рядок успішно видалено!");
                                    // Оновлення відображення dataGridView2
                                    LoadData();
                                }
                                else
                                {
                                    MessageBox.Show("Не вдалося видалити рядок.");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Помилка: " + ex.Message);
                    }
                }
            }
            else
            {
                MessageBox.Show("Будь ласка, виберіть рядок для видалення.");
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            // Очищення полів textBox6, textBox8, textBox1
            textBox6.Clear();
            textBox7.Clear();
            textBox8.Clear();
            textBox1.Clear();

            // Повернення dataGridView2 у режим відображення всього вмісту БД
            LoadData();
        }
        private void buttonExamination_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Open();

                // Отримуємо поточну дату
                DateTime currentDate = DateTime.Today;

                // Отримуємо дані з таблиці "Втрачені речі"
                OleDbDataAdapter lostItemsAdapter = new OleDbDataAdapter("SELECT * FROM [Втрачені речі]", connection);
                DataTable lostItemsTable = new DataTable();
                lostItemsAdapter.Fill(lostItemsTable);

                // Перевірка, чи є рядки для копіювання в "Аукціон"
                bool rowsInserted = false;

                // Ініціалізуємо рядок, щоб зберегти ідентифікаційний номер успішно занесеного рядка
                string insertedItemId = "";

                // Проходимося по кожному рядку таблиці "Втрачені речі"
                foreach (DataRow row in lostItemsTable.Rows)
                {
                    DateTime foundDate = Convert.ToDateTime(row["Дата знаходження втраченої речі"]);

                    // Перевіряємо, чи пройшов рік з моменту знаходження речі
                    if (foundDate.AddYears(1) <= currentDate)
                    {
                        // Копіюємо дані до таблиці "Аукціон"
                        string insertAuctionQuery = "INSERT INTO [Аукціон] ([Ідентифікаційний номер], [Дата початку сберігання], [Дата розміщення на аукціон], [Статус], [Кому продано ПІБ], [Номер телефону]) VALUES (?, ?, ?, ?, ?, ?)";
                        using (OleDbCommand insertAuctionCommand = new OleDbCommand(insertAuctionQuery, connection))
                        {
                            insertAuctionCommand.Parameters.AddWithValue("@p1", row["Ідентифікаційний номер"]);
                            insertAuctionCommand.Parameters.AddWithValue("@p2", row["Дата знаходження втраченої речі"]);
                            insertAuctionCommand.Parameters.AddWithValue("@p3", currentDate); // Сьогоднішня дата
                            insertAuctionCommand.Parameters.AddWithValue("@p4", "На продажі"); // Статус "На продажі"
                            insertAuctionCommand.Parameters.AddWithValue("@p5", DBNull.Value); // Поки що не продано
                            insertAuctionCommand.Parameters.AddWithValue("@p6", DBNull.Value); // Поки що не вказано номер телефону

                            insertAuctionCommand.ExecuteNonQuery();

                            // Зберігаємо ідентифікаційний номер успішно занесеного рядка
                            insertedItemId = row["Ідентифікаційний номер"].ToString();

                            rowsInserted = true;
                        }
                    }
                }

                if (rowsInserted)
                {
                    MessageBox.Show($"Рядок(и) успішно занесено у таблицю 'Аукціон'. Ідентифікаційний номер: {insertedItemId}", "Успіх");
                }
                else
                {
                    MessageBox.Show("Відсутні рядки для занесення у таблицю 'Аукціон'.", "Увага");
                }

                LoadData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        private void sell_Click(object sender, EventArgs e)
        {
            // Перевірка, чи обрано рядок у dataGridView3
            if (dataGridView3.SelectedRows.Count > 0)
            {
                // Отримуємо ідентифікаційний номер обраного рядка
                int identificationNumber = Convert.ToInt32(dataGridView3.SelectedRows[0].Cells["ідентифікаційний Номер"].Value);

                // Перевірка наявності даних у стовбцях "Кому продано ПІБ" та "Номер телефону"
                if (!string.IsNullOrEmpty(dataGridView3.SelectedRows[0].Cells["Кому продано ПІБ"].Value.ToString()) &&
                    !string.IsNullOrEmpty(dataGridView3.SelectedRows[0].Cells["Номер телефону"].Value.ToString()))
                {
                    MessageBox.Show("Помилка. Річ вже продано.", "Увага");
                    return;
                }

                // Відкриваємо Form2
                using (Form2 form2 = new Form2(identificationNumber, this))
                {
                    if (form2.ShowDialog() == DialogResult.OK)
                    {
                        // Отримуємо дані з Form2
                        string buyerName = form2.textBox1.Text;
                        string phoneNumber = form2.textBox2.Text;
                        // Заносимо дані у відповідні стовбці таблиці "Аукціон"
                        try
                        {
                            connection.Open();

                            string updateQuery = "UPDATE [Аукціон] SET [Кому продано ПІБ] = ?, [Номер телефону] = ?, [Статус] = ? WHERE [Ідентифікаційний номер] = ?";
                            using (OleDbCommand updateCommand = new OleDbCommand(updateQuery, connection))
                            {
                                updateCommand.Parameters.AddWithValue("@p1", buyerName);
                                updateCommand.Parameters.AddWithValue("@p2", phoneNumber);
                                updateCommand.Parameters.AddWithValue("@p3", "Продано");
                                updateCommand.Parameters.AddWithValue("@p4", identificationNumber);

                                int rowsAffected = updateCommand.ExecuteNonQuery();
                           
                                if (rowsAffected > 0)
                                {
                                    MessageBox.Show("Дані про продаж успішно збережено.", "Успіх");
                                }
                                else
                                {
                                    MessageBox.Show("Не вдалося змінити дані про продаж.", "Помилка");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Помилка: " + ex.Message);
                        }
                        finally
                        {
                            connection.Close();
                        }

                    }
                }
            }
            else
            {
                MessageBox.Show("Спочатку оберіть рядок для продажу.", "Увага");
            }
        }
        // Метод, який буде викликаний при спрацьовуванні події DataUpdated
        private void textBox5_TextChanged(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                string identificationNumber = textBox5.Text;
                if (!string.IsNullOrWhiteSpace(identificationNumber))
                {
                    try
                    {
                        connection.Open();

                        // Запит для вибірки рядків за ідентифікаційним номером
                        string selectQuery = "SELECT * FROM [Аукціон] WHERE [Ідентифікаційний номер] = ?";
                        using (OleDbCommand selectCommand = new OleDbCommand(selectQuery, connection))
                        {
                            selectCommand.Parameters.AddWithValue("@p1", identificationNumber);

                            OleDbDataAdapter adapter = new OleDbDataAdapter(selectCommand);
                            DataTable dataTable = new DataTable();
                            adapter.Fill(dataTable);

                            dataGridView3.DataSource = dataTable;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Помилка: " + ex.Message);
                    }
                    finally
                    {
                        connection.Close();
                    }
                }
            }
        }
        private void clearAK_Click_1(object sender, EventArgs e)
        {
            textBox5.Clear();
            LoadData();
        }
        // Видалення рядку з таблиці "Аукціон" за допомогою поля "Ідентифікаційний Номер"
        private void deleteRowButton_Click(object sender, EventArgs e)
        {
            if (dataGridView3.SelectedRows.Count > 0)
            {
                DialogResult result = MessageBox.Show("Ви впевнені, що хочете видалити обраний рядок?", "Підтвердження видалення", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    foreach (DataGridViewRow row in dataGridView3.SelectedRows)
                    {
                        int identificationNumber = Convert.ToInt32(row.Cells["ідентифікаційний Номер"].Value);

                        try
                        {
                            connection.Open();

                            string deleteQuery = "DELETE FROM [Аукціон] WHERE [Ідентифікаційний номер] = ?";
                            using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                            {
                                deleteCommand.Parameters.AddWithValue("@p1", identificationNumber);
                                deleteCommand.ExecuteNonQuery();
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Помилка видалення рядка: " + ex.Message, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            connection.Close();
                        }

                        dataGridView3.Rows.Remove(row);
                    }
                }
            }
            else
            {
                MessageBox.Show("Спочатку оберіть рядок для видалення.", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public void RefreshDataGridView()
        {
            // Очищаємо вміст dataGridView3 перед оновленням
            dataGridView3.DataSource = null;

            try
            {
                connection.Open();

                // SQL запит для вибору всіх даних з таблиці "Аукціон"
                string auctionQuery = "SELECT * FROM [Аукціон]";

                // Створюємо DataAdapter для виконання запиту та заповнення DataGridView
                OleDbDataAdapter auctionAdapter = new OleDbDataAdapter(auctionQuery, connection);

                // Створюємо об'єкт DataTable для зберігання результатів запиту
                DataTable auctionDataTable = new DataTable();

                // Заповнюємо DataTable результатами запиту
                auctionAdapter.Fill(auctionDataTable);

                // Встановлюємо DataTable як джерело даних для dataGridView3
                dataGridView3.DataSource = auctionDataTable;

                // Очищаємо вміст dataGridView1 перед оновленням
                dataGridView1.DataSource = null;

                // SQL запит для вибору всіх даних з таблиці "Втрачені речі"
                string lostItemsQuery = "SELECT * FROM [Втрачені речі]";

                // Заповнюємо DataTable результатами запиту для "Втрачені речі"
                DataTable lostItemsDataTable = new DataTable();
                auctionAdapter.SelectCommand.CommandText = lostItemsQuery;
                auctionAdapter.Fill(lostItemsDataTable);

                // Встановлюємо DataTable як джерело даних для dataGridView1
                dataGridView1.DataSource = lostItemsDataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка при оновленні даних: " + ex.Message);
            }
            finally
            {
                // Закриваємо з'єднання
                connection.Close();
            }
        }
    }
}