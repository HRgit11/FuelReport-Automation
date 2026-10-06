using mamaRabota.Classes;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Management;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
namespace mamaRabota
{
    /// <summary>
    /// Логика взаимодействия для LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : Window
    {
        public LoginWindow()
        {
            InitializeComponent();
            
        }

        private async void enterBtn_Click(object sender, RoutedEventArgs e)
        {
            string hwId = HardwareUtils.GetHwId();
            string clientKey = txtKey.Text;

            var data = new { key = clientKey, hwid = hwId };
            string jsonString = JsonConvert.SerializeObject(data);
            var content = new StringContent(jsonString, Encoding.UTF8, "application/json");

            var client = new HttpClient();

            //ХАРДКОР!!!!!
            try
            {
                var response = await client.PostAsync("http://localhost:5000/api/License/activate", content);

                if (response.IsSuccessStatusCode)
                {
                    MessageBox.Show("Успешно! Программа активирована.");

                    string basePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
                    string appFolder = Path.Combine(basePath, "mamaRabota");
                    string configFolder = Path.Combine(appFolder, "config");

                    Directory.CreateDirectory(configFolder);

                    string filePath = Path.Combine(configFolder, "license.key");
                    File.WriteAllText(filePath, clientKey);

                    var main = new MainWindow();
                    main.Show();
                    this.Close();
                }
                else
                {
                    MessageBox.Show($"Ошибка: {response.StatusCode}");
                }
            }
            catch
            {
                MessageBox.Show("Сервер не запущен!");
            }
            
        }



    }
}
