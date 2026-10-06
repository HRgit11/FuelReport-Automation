using mamaRabota.Classes;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Windows;

namespace mamaRabota
{
    public partial class App : Application
    {
        private async void Application_Startup(object sender, StartupEventArgs e)
        {
            bool isLicenseValid = false;

            // Тот же путь, что и при записи
            string basePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            string appFolder = Path.Combine(basePath, "mamaRabota");
            string configFolder = Path.Combine(appFolder, "config");
            string licensePath = Path.Combine(configFolder, "license.key");

            if (File.Exists(licensePath))
            {
                try
                {
                    string savedKey = File.ReadAllText(licensePath).Trim();
                    string hwId = HardwareUtils.GetHwId();

                    var data = new { key = savedKey, hwid = hwId };
                    string json = JsonConvert.SerializeObject(data);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var client = new HttpClient();
                    var response = await client.PostAsync("http://localhost:5000/api/License/check", content);

                    if (response.IsSuccessStatusCode)
                    {
                        isLicenseValid = true;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Ошибка проверки лицензии: {ex.Message}");
                }
            }

            if (isLicenseValid)
            {
                var main = new MainWindow();
                main.Show();
            }
            else
            {
                var login = new LoginWindow();
                login.Show();
            }
        }

    }
}