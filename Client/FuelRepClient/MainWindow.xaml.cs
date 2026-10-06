using mamaRabota.Classes;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;



namespace mamaRabota
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private double fSum;
        private double fLitrs;
        private string month;
        private string season;
        private List<CheckBox> allCheckBoxex;
        public MainWindow()
        {
            InitializeComponent();
            allCheckBoxex = new List<CheckBox>()
            {
                chekBox1,
                chekBox2, chekBox3, chekBox4, chekBox5,
                chekBox6, chekBox7,
                chekBox8, chekBox9,
                chekBox10,
                chekBox11,
                chekBox12,
                chekBox13,
                chekBox14,
                chekBox15,
                chekBox16,
                chekBox17,
                chekBox18,
                chekBox19,
                chekBox20,
                chekBox21,
                chekBox22,
                chekBox23,
                chekBox24,
                chekBox25,
                chekBox26,
                chekBox27,
                chekBox28,
                chekBox29,
                chekBox30,
                chekBox31,
            }; 

        }

        private void upLoadBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                fSum = Convert.ToDouble(fullSumtxt.Text);
                fLitrs = Convert.ToDouble(fullLitrstxt.Text);
                var nList=Convert.ToInt32(txtNPutList.Text);
                var nList2=Convert.ToInt32(txtTireN.Text);


                if (season != "зима" && season != "лето")
                {
                    MessageBox.Show("Выберите время года!");
                    return;
                }

                if (string.IsNullOrEmpty(cmbMonth.Text))
                {
                    MessageBox.Show("Выберите месяц!");
                    return;
                }

                var openFileDialog = new OpenFileDialog();
                openFileDialog.ShowDialog();
                if (openFileDialog != null)
                {
                    var selectedFile = openFileDialog.FileName;
                    var checkedBoxes = allCheckBoxex.Where(cb => cb.IsChecked == true).ToList();
                    GetListNameDate();
                    var data = new ExcelData
                    {
                        File = selectedFile,
                        FullDays = checkedBoxes.Count,
                        FullSum = fSum,
                        FullLitrs = fLitrs,
                        AllChekBoxes = checkedBoxes,
                        Month = month,
                        NuberList = nList,
                        NuberList2 = nList2,
                        DaysLitrs = GetLitrsInOneDay(),
                        Season = GetSeason(),
                        Target = Convert.ToDouble(txtTarget.Text)

                    };
                    var excel = new ExcelClass(data);
                    excel.OpenFile();
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Проверьте правильность заполнения полей");
            }
            
        }
        
        private void GetListNameDate()
        {
            try
            {
                var months = new Dictionary<string, string>
{
    {"Январь", "01"}, {"Февраль", "02"}, {"Март", "03"},
    {"Апрель", "04"}, {"Май", "05"}, {"Июнь", "06"},
    {"Июль", "07"}, {"Август", "08"}, {"Сентябрь", "09"},
    {"Октябрь", "10"}, {"Ноябрь", "11"}, {"Декабрь", "12"}
};

                if (months.TryGetValue(cmbMonth.Text, out string monthValue))
                {
                    month = monthValue;
                }
            }
            catch( Exception ex )
            {
                MessageBox.Show("Ошибка получения месяца");
            }
        }

        private Dictionary<string, string> GetLitrsInOneDay()
        {
            var daysLitrs = new Dictionary<string, string>();
            try
            {
                List<TextBox> listTxt = new List<TextBox>()
        {
            txtDay1, txtLitrs1, txtDay2, txtLitrs2,
            txtDay3, txtLitrs3, txtDay4, txtLitrs4,
            txtDay5, txtLitrs5, txtDay6, txtLitrs6,
            txtDay7, txtLitrs7, txtDay8,txtLitrs8
        };

                for (int i = 0; i < listTxt.Count; i += 2)
                {
                    string day = listTxt[i].Text?.Trim();
                    string litrs = listTxt[i + 1].Text?.Trim();

                    if (!string.IsNullOrEmpty(day) && !string.IsNullOrEmpty(litrs))
                    {
                        if (!daysLitrs.ContainsKey(day))
                        {
                            daysLitrs.Add(day, litrs);
                        }
                    }
                }
                return daysLitrs;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " Ошибка заполнения дней заправки");
                return null;
            }
        }

        public string GetSeason()
        {
            if (season == "зима")
            {
                return "12,60";
            }
            else if(season == "лето")
            {
                return "11,68";
            }
            else
            {
                MessageBox.Show("Выберите время года!");
                return "";
            }
        }

        private void rbWinter_Checked(object sender, RoutedEventArgs e)
        {
            season = "зима";
        }

        private void rbSummer_Checked(object sender, RoutedEventArgs e)
        {
            season = "лето";
        }
    }
}
