
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Word = Microsoft.Office.Interop.Word;

namespace Potolki
{
    public partial class Base_form : Form
    {
        public Base_form()
        {
            // Инициализация компонентов
            InitializeComponent();
        }
        // глобальные переменные для обленгчения передачи данных между методами

        public static double razmer; // итоговый размер потолка
        public static double rezult; // итоговая сумма потолка

        public static string chek_text1;
        public static string chek_text2;
        public static string chek_text3;

        public static bool triger_chek = false;

        // метод нажатия кнопки "Расчитать"
        private void button1_Click(object sender, EventArgs e)
        {
            //значение добавочной стоимости при фотопечати и многоуровневости если эти параметры небыли выбраны пользователем
            double price_percent_over_level = 1;
            double price_percent_photo_print = 1;


            if (checkBox1.Checked == true)
            {
                //включение умножения конечной стоимости на 30% если был выбран многоуровневый потолок
                price_percent_over_level = 1.3;
                chek_text1 = " многоуровневый";
            }

            if (checkBox2.Checked == true)
            {
                //включение умножения конечной стоимости на 26% если была выбрана фотопечать
                price_percent_photo_print = 1.26;
                chek_text2 = " с фотопечатью";
            }

            // результат если будет выбран глянцевый тип потолков
            if (radioButton1.Checked == true)
            {
                try
                {
                    // расчет квадратуры и итоговой суммы потолка с дальнейшим выводом в listBOX
                    razmer = Convert.ToDouble(calculation_meter(textBox1.Text, textBox2.Text));
                    rezult = (razmer * 213.15) * price_percent_over_level * price_percent_photo_print;
                    listBox1.Items.Clear();
                    listBox1.Items.Add($"Расчет показал что:");
                    listBox1.Items.Add($"------------------------------------------------------");
                    listBox1.Items.Add($"Указанная квадратура равняется: {razmer} м²");
                    listBox1.Items.Add($"------------------------------------------------------");
                    listBox1.Items.Add($"Стоимость потолка в {razmer} м² с учетом  ");
                    listBox1.Items.Add($"указанных параметров выйдет: {rezult} руб.");

                    chek_text3 = $"Глянцевый потолок {razmer} м² ";
                    triger_chek = true;
                }
                catch
                {
                    // сообщение об ошибке ввода высоты или ширины
                    MessageBox.Show(
"Не верно введены значения ширины или высоты!",
"Ошибка",
MessageBoxButtons.OK,
MessageBoxIcon.Error,
MessageBoxDefaultButton.Button1,
MessageBoxOptions.DefaultDesktopOnly);
                }
               
            }
            // результат если будет выбран матовый тип потолков
            else if (radioButton2.Checked == true)
            {
                try
                {
                    // расчет квадратуры и итоговой суммы потолка с дальнейшим выводом в listBOX
                    razmer = Convert.ToDouble(calculation_meter(textBox1.Text, textBox2.Text));
                    rezult = (razmer * 265.80) * price_percent_over_level * price_percent_photo_print;
                    listBox1.Items.Clear();
                    listBox1.Items.Add($"Расчет показал что:");
                    listBox1.Items.Add($"------------------------------------------------------");
                    listBox1.Items.Add($"Указанная квадратура равняется: {razmer} м²");
                    listBox1.Items.Add($"------------------------------------------------------");
                    listBox1.Items.Add($"Стоимость потолка в {razmer} м² с учетом  ");
                    listBox1.Items.Add($"указанных параметров выйдет: {rezult} руб.");

                    chek_text3 = $"Матовый потолок {razmer} м² ";
                    triger_chek = true;
                }
                catch
                {
                    // сообщение об ошибке ввода высоты или ширины
                    MessageBox.Show(
"Не верно введены значения ширины или высоты!",
"Ошибка",
MessageBoxButtons.OK,
MessageBoxIcon.Error,
MessageBoxDefaultButton.Button1,
MessageBoxOptions.DefaultDesktopOnly);
                }
               
            }

          
        }
        //Метод подсчета квадратных метрос с отсеиванием неверныо введенных в поле ввода параметров
        public static string calculation_meter(string width, string height)
        {
            string answer = "null";

            try
            {
                // Преобразование тескта в число с плавующей запятой
                double width_double = Convert.ToDouble(width);
                double height_double = Convert.ToDouble(height);

                if (width_double <= 0 || height_double <= 0)
                {
                    answer = "null";
                }
                else
                {
                    // подсчет квадратуры помещения
                    answer = "" + (width_double* height_double);
                }
            }
            catch
            {
                // присваивание методу сообщение об ошибка при неверно занесенных данных
                answer = "error";
            }

            // возврат ответа от метода
            return answer;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (triger_chek == true)
            {
                // Создаём объект документа
                Microsoft.Office.Interop.Word.Document doc = null;
                try
                {
                    // Создаём объект приложения
                    Word.Application app = new Word.Application();
                    // Путь до шаблона документа
                    string source = System.IO.Path.GetFullPath("..\\Шаблон квитанции\\Квитанция.docx");
                    // Открываем
                    doc = app.Documents.Add(source);
                    doc.Activate();

                    // Добавляем информацию
                    // wBookmarks содержит все закладки
                    Microsoft.Office.Interop.Word.Bookmarks wBookmarks = doc.Bookmarks;
                    Word.Range wRange;
                    int i = 0;
                    Random random = new Random();
                    int randoms = random.Next(100000000, 999999999);
                    DateTime dateTime = DateTime.Now;
                    string[] data = new string[4] { $"{randoms}", $"{dateTime}", $"{chek_text3} ( Дополнительные особоенности:{chek_text1}{chek_text2} )", $"{rezult}" };
                    foreach (Word.Bookmark mark in wBookmarks)
                    {

                        wRange = mark.Range;
                        wRange.Text = data[i];
                        i++;
                    }

                    // Закрываем документ
                    doc.Close();
                    doc = null;

                    MessageBox.Show(
"Квитанция успешно сформирована!",
"Информация",
MessageBoxButtons.OK,
MessageBoxIcon.Information,
MessageBoxDefaultButton.Button1,
MessageBoxOptions.DefaultDesktopOnly);
                }
                catch (Exception ex)
                {
                    doc = null;
                    Console.WriteLine("Во время выполнения произошол системный сбой пожалуста презапустите приложение!");
                    Console.ReadLine();
                }
            }
            else
            {
                MessageBox.Show(
"Для формирования квмтанции нужно произвести расчет!",
"Ошибка",
MessageBoxButtons.OK,
MessageBoxIcon.Error,
MessageBoxDefaultButton.Button1,
MessageBoxOptions.DefaultDesktopOnly);
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
"В данной версии програмного обеспечения этого сделать нельзя!",
"Ошибка",
MessageBoxButtons.OK,
MessageBoxIcon.Error,
MessageBoxDefaultButton.Button1,
MessageBoxOptions.DefaultDesktopOnly);
        }
    }
}
