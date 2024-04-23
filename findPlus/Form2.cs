using System;
using System.Data.OleDb;
using System.Windows.Forms;

namespace findPlus
{
    public partial class Form2 : Form
    {
        private OleDbConnection connection;
        private const string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=F:\ДДМА\3 Курс\2а семестр\ОБДЗ Курсова робота\АІС для роботи бюро знахідок\findPlus\findPlus\Bureau of finds.accdb";
        private int identificationNumber;

        // Подія, що відбувається при закритті форми і передає дані про оновлення назад до Form1
        public event EventHandler<DataUpdatedEventArgs> DataUpdated;
        private Form1 parentForm;
        public Form2(int identificationNumber, Form1 parentForm)
        {
            InitializeComponent();
            this.identificationNumber = identificationNumber;
            this.parentForm = parentForm;
            connection = new OleDbConnection(connectionString);
        }

        private void agree_Click(object sender, EventArgs e)
        {
            // Отримуємо дані з текстових полів
            string buyerName = textBox1.Text;
            string phoneNumber = textBox2.Text;

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

                    // Видаляємо рядок з таблиці "Втрачені речі" за Ідентифікаційним номером
                    string deleteQuery = "DELETE FROM [Втрачені речі] WHERE [Ідентифікаційний номер] = ?";
                    using (OleDbCommand deleteCommand = new OleDbCommand(deleteQuery, connection))
                    {
                        deleteCommand.Parameters.AddWithValue("@p4", identificationNumber);

                        int rowsDeleted = deleteCommand.ExecuteNonQuery();
                    }

                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Дані про продаж успішно збережено.", "Успіх");

                        // Оновлення вмісту dataGridView
                        parentForm.RefreshDataGridView();
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

            // Після оновлення даних, викликаємо функцію DataUpdated і передаємо нові дані
            OnDataUpdated(new DataUpdatedEventArgs(identificationNumber, buyerName, phoneNumber));

            // Закриваємо форму
            this.Close();
        }
        // Метод для спрацьовування події DataUpdated
        protected virtual void OnDataUpdated(DataUpdatedEventArgs e)
        {
            DataUpdated?.Invoke(this, e);
        }
    }

    // Клас, що містить дані, які були оновлені в Form2 і потрібно передати назад до Form1
    public class DataUpdatedEventArgs : EventArgs
    {
        public int IdentificationNumber { get; }
        public string BuyerName { get; }
        public string PhoneNumber { get; }

        public DataUpdatedEventArgs(int identificationNumber, string buyerName, string phoneNumber)
        {
            IdentificationNumber = identificationNumber;
            BuyerName = buyerName;
            PhoneNumber = phoneNumber;
        }
    }
}