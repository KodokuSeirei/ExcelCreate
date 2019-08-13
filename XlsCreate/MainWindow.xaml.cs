using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections.ObjectModel;
using static ExcelCreate.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Globalization;

namespace ExcelCreate
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {           
            InitializeComponent();
            textSavePath.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            textVAT.Text = "20";
        }

        Data dt = new Data();
        public string FullPath { get { return textSavePath.Text + "/" + textName.Text; } set { } }
        private void Button_Click(object sender, RoutedEventArgs e)
        {

            if (datePicker.Text!="")
            {
                        if (textVAT.Text == "20" || textVAT.Text == "18" || textVAT.Text == "10" || textVAT.Text == "10/110" || textVAT.Text == "18/118" || textVAT.Text == "20/120")
                        {  
                           dt.GetVATExcel(datePicker.Text, textVAT.Text, FullPath);
                        }
                        else
                        {
                           MessageBox.Show("Неверный НДС. Выберите одно из значений: 20 || 18 || 10 || 10/110 ||18/118 || 20/120");
                        }                
            }
            else
            {
                MessageBox.Show("Введите дату формата дд.мм.гггг");
                return;
            }

        }
            
    }

}


    
