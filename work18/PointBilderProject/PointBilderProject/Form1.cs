using System;
using LibrarsStraightLine;  // Импорт  библиотеки классов для работы с геометрическими объектами
using System.Windows.Forms;

namespace PointBilderProject
{
    // Основной класс формы приложения для работы с геометрическими операциями
    public partial class Parallel_OX_Check : Form
    {
        // Поле для хранения текущего выбранного инструмента (типа линии)
        public string ToolsUse;

        
        public Parallel_OX_Check()
        {
            InitializeComponent();
        }

        // Поле для хранения введенных пользователем координат
        public string Check1Tools;

        // Обработчик события загрузки формы (в текущей реализации не содержит кода)
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        // Обработчик клика по метке, отображающей текущий инструмент
        private void labeltools_Click(object sender, EventArgs e)
        {

        }

        // Обработчик кнопки проверки параллельности линии оси OX
        // Вызывает метод из библиотеки для определения параллельности
        private void ParallelOX_Chek_Click(object sender, EventArgs e)
        {
            string lineParallel = Check1Tools;  // Получение сохраненных координат
            // Вызов метода проверки параллельности и вывод результата
            textBox1.Text = (PointBilder.ConvertToPointParal(lineParallel));
        }

        // Обработчик выбора инструмента "Прямая линия" из меню
        private void straightLineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OutputInfLine.Select();  // Установка фокуса на поле ввода координат
            ToolsUse = "StraightLine";  // Сохранение типа выбранного инструмента
            labeltools.Text = "straight line";  // Обновление отображаемого названия инструмента
        }

        // Обработчик выбора инструмента "Луч" из меню
        private void beamLineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OutputInfLine.Select();  // Установка фокуса на поле ввода координат
            // Вывод подсказки о специфике работы с лучом
            textBox1.Text = "учтите что первая точка статична, а вторая задает направление луча";
            ToolsUse = "BeamLine";  // Сохранение типа выбранного инструмента
            labeltools.Text = "beam line";  // Обновление отображаемого названия инструмента
        }

        // Обработчик выбора инструмента "Отрезок" из меню
        private void lengthToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            OutputInfLine.Select();  // Установка фокуса на поле ввода координат
            ToolsUse = "LimitedLine";  // Сохранение типа выбранного инструмента
            labeltools.Text = "limited line";  // Обновление отображаемого названия инструмента
        }

        // Основной обработчик кнопки ввода - активирует интерфейс и обрабатывает координаты
        private void Input_Click(object sender, EventArgs e)
        {
            // Активация элементов интерфейса, которые были скрыты
            Intersection_of_axes.Visible = true;
            Length_Limited_Line.Visible = true;
            labeltools.Visible = true;
            textBox1.Visible = true;
            angle_OX_Chek.Visible = true;
            ParallelOX_Chek.Visible = true;

            string lineBild = OutputInfLine.Text;  // Получение введенных координат
            Check1Tools = OutputInfLine.Text;  // Сохранение координат для дальнейшего использования
            // Преобразование введенных координат в форматированную строку и вывод обратно
            OutputInfLine.Text = (PointBilder.ConvertToPointString(lineBild));
        }

        // Обработчик изменения текста в поле ввода координат
        private void OutputInfLine_TextChanged(object sender, EventArgs e)
        {

        }

        // Обработчик изменения текста в поле вывода результатов
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
             
        }

        // Обработчик кнопки проверки угла между линией и осью OX
        // Вызывает метод из библиотеки для вычисления угла
        private void angle_OX_Chek_Click(object sender, EventArgs e)
        {
            string lineParallel = Check1Tools;  // Получение сохраненных координат
            // Вызов метода вычисления угла и вывод результата
            textBox1.Text = (PointBilder.ConvertToPointCornerOX(lineParallel));
        }

        // Обработчик выбора меню "Инструменты" - скрывает справочную информацию
        private void toolsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            labelHelp.Visible = false;  // Скрытие метки со справочной информацией
        }

        // Обработчик клика по метке справочной информации
        private void labelHelp_Click(object sender, EventArgs e)
        {

        }

        // Обработчик кнопки вычисления длины отрезка
        // Доступен только для инструмента "Отрезок"
        private void Length_Limited_Line_Click(object sender, EventArgs e)
        {
            switch (ToolsUse)  // Проверка текущего выбранного инструмента
            {
                case "LimitedLine":  // Если выбран инструмент "Отрезок"
                    // Вычисление длины отрезка и вывод результата
                    textBox1.Text = (Limited_Line.Length_Limited_Line(Check1Tools));
                    break;
                default:  // Для других инструментов
                    // Вывод сообщения об ошибке доступности
                    textBox1.Text = "Данный метод доступен только классу инструментов: Отрезки (LimitedLine)";
                    break;
            }
        }

        // Обработчик кнопки проверки пересечения линии с осями координат
        // В зависимости от типа линии вызывает соответствующий метод
        private void Intersection_of_axes_Click(object sender, EventArgs e)
        {
            switch (ToolsUse)  // Проверка текущего выбранного инструмента
            {
                case "BeamLine":  // Для луча
                    // Проверка пересечения луча с осями координат
                    textBox1.Text = BeamLine.Intersection_Of_Axes_BeamLine(Check1Tools);
                    break;
                case "LimitedLine":  // Для отрезка
                    // Проверка пересечения отрезка с осями координат
                    textBox1.Text = Limited_Line.Intersection_Of_Axes_Limited_Line(Check1Tools);
                    break;
                case "StraightLine":  // Для прямой
                    // Проверка пересечения прямой с осями координат
                    textBox1.Text = (StraightLine.Intersection_Of_AxesStraightLine(Check1Tools));
                    break;
                default:  // Если инструмент не выбран
                    textBox1.Text = "Ошибка";
                    break;
            }
        }

        // Обработчик кнопки очистки поля ввода координат
        private void Clean_Click(object sender, EventArgs e)
        {
            OutputInfLine.Text = "";  // Очистка поля ввода
        }

        // Обработчик кнопки выхода из приложения
        private void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();  // Завершение работы приложения
        }
    }
}