using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Net;
using System.Reflection;
using System.Net.Sockets;
#region --- лицензия GNU
//                           LICENSE INFORMATION
// Urine Analyzer Driver, serial port & SQL communications. Version 1.001.01.
// Thos driver is for managing data from serial port & its transmission to MS-SQL Server 
// according  to the rules described in Russuan Specification "UriLit-150. Мочевой анализатор.
// Руководство по эксплуатации".
//
// Copyright (C)  2020 Vladimir A. Maltapar
// Email: maltapar@gmail.com
// Created: 18 August 2020
//
// This program is free software: you can redistribute it and/or modify  it under the terms of
// the GNU General Public License as published by the Free Software Foundation, 
// either version 3 of the License, or (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,  but WITHOUT ANY WARRANTY; 
// without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
// See the GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License  along with this program.
// If not, see <http://www.gnu.org/licenses/>.
//********************************************************************************************
#endregion --- лицензия GNU

namespace Uri150
{
    public partial class Form1 : Form
    {
        #region --- Общие параметры 
        // параметры в строчках ini-файла 
        private static int AnalyzerID;     // для UriLi-150 ID=13 - д.б. указан в настройках!
        private static string ComPortNo;   // назваеие COM-порта: Сom1, COM2, ...
        private static string connStr;     // строка коннекта к SQL
        private static int MaxCntAn = 22;  // максимальное кол-во анализов одного пациента, 
                                           // передаваемых в MS-SQL за 1 раз
                                           // (это ограничение структуры AnalyserResults :((  )
        private static int MaxHistNo = 9;  // максимальный номер истории (для ограничения ошибок ввода на анализаторе - там 13 цифр!) 
                                           // берётся из .ini-файла
        private static string sModes;      // строка списка режимов работы 
                                           // режим: WriteToSQL=Yes или No
                                           // режим: (NoComPort) - работать, даже ели на комп. нет COM-порта
                                           // режим: "Все анализы одного пациента в одну строку"
                                           // режим: дата анализа берётся в SQL по GetDate()
        private static string sDebugModes; // строка списка режимов отладки:
                                           // режим: (лог частей) 
                                           // режим: (лог приёма)
                                           // режим: (лог квиточка)
                                           // режим: (лог SQL) 
                                           // режим: (Log_Excel)
                                           // режим: (лог пациентов без номера истории)
        private static string PathLog;     // путь к лог-файлу с вычисленными каталогом и датой ...\GGGG-MM\Uri150_2020-08-06.txt 
        private static string PathLogParm; // путь к лог-файлу, заданный в параметрах ini-файла  
        private static string PathErrLog;  // путь к лог-файлу ошибок 
        private static string pathIniFile; // путь+имя ini-файла (как и имя запущенной программы)
        private static string[] str_ini;   // строки ini-файла
        
        // глобальные
        private static DateTime dt0 = DateTime.Now; // время старта / (пред)последнего приёма данных
        private static DateTime dtm = dt0;          // время последнего приёма данных
        private string UserName = System.Environment.UserName;
        private string ComputerName = System.Environment.MachineName;
        public SerialPort _serialPort;
        private string PathIni;         // Путь к ini-файлу (пусть будет там, откуда запуск? пока так!)
        private static string AppName;  // static т.к. исп. в background-процессе
        private static string AppVer="Версия 1.7 ";   // 2020-09-03, 2020-10-21
        private static string qkrq = ""; // ToDo запрос о работе ("кукареку") 
        private static string strV1, strV2, strV3, strV4, strV5; // на форме - напоминалки :)
        private static string dateDone = "2020-12-31 23:59";  // дата-время выполнения анализа по часам на анализаторе
        private static string dateDone999 = "31-12-2020 23:59:00.000";  // дата-время выполнения анализа для SQL ResultDate
        // начальные значения дла нового пациента
        private static int    kAn = 0; // количество анализов у одного пациента - для UriLi-150 ВСЕГДА 11 анализов!
        //private static string s0 = ""; // формирование строки для SQL ResultText
        private static string s1 = "Insert into ... инициализация в InitSQLstr";
        private static string s2 = "values (... ... инициализация в InitSQLstr";
        private string inputString = ""; // полученные и собранные вместе за сеанс передачи данные для парсинга
        private static Int64 nHistNo = -1;
        private string testNo = "";      // порядковый номер теста, присваиваемый анализатором, и он его каждый раз увеличивет на 1.
        private int kTest = 0;           // кол-во переданных тестов с начала работы драйвера - считает программа!
        private int kPart = 0;           // кол-во частей передачи данных для одного пациетна
        private int kLines = 0;          // кол-во строк в переданном анализе
        private string msg = "---";      // последнее китайстое предупреждение :))
        private char STX = '\x02';       // (символ "улыбка") 02h -  начало передачи
        private char ETX = '\x03';       // (символ "черви")  03h -  конец передачи
        private int F_Remind = 0;        // флаг - напоминалки. =0 по умолчанию - не отображать напоминалки
        private Random rnd = new Random(32123);
        #endregion --- Общие параметры 
        public Form1()      // всё, что нажито непосильным трудом, ... 
        {
            InitializeComponent();  // всё начиналось здесь... 
            ReadParmsIni();         // берём параметры из ini-файла 
            FormIni();              // установки на форме, которые не делает Visual Studio - IP-адреса
            PortIni();              // начальные установки сом-порта
        }
        private void FormIni()  // начальный вывод на форму - что прочитали из ini-файла
        {
            var host = Dns.GetHostEntry(Dns.GetHostName()).AddressList;
            var list_adr = from ip in host where ip.ToString().Length < 16 select ip.ToString();
            string IP = "";
            foreach (string st in list_adr) IP += st + " ";
            Lbl_IP.Text = "IP: " + IP;
            CmbTest.SelectedIndex = 0;  // первый (нулевой) элемент - текущий, видимый.

            notifyIcon1.Visible = false; // невидимая иконка в трее

            // добавляем Эвент или событие по 2му клику мышки, 

            //вызывая функцию  notifyIcon1_MouseDoubleClick
            this.notifyIcon1.MouseDoubleClick += new MouseEventHandler(notifyIcon1_MouseDoubleClick);

            // добавляем событие на изменение окна
            this.Resize += new System.EventHandler(this.Form1_Resize);
            this.ShowInTaskbar = false;
        }
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();
            notifyIcon1.Visible = false;
            WindowState = FormWindowState.Normal;
        }
        private void Form1_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                notifyIcon1.Visible = true; // иконка в трее видимая
                notifyIcon1.ShowBalloonTip(1000); // 1 секунду показать предупреждение о сворачивании в трей
            }
            else if (FormWindowState.Normal == this.WindowState)
            { notifyIcon1.Visible = false; }
        }
        #region --- методы для Com-порта: инициализация (PortIni) и чтения (Sp_DataReceived)
        private delegate void SetTextDeleg(string text);  // Делегат используется для записи в UI control из потока не-UI
        // Все опции для последовательного устройства могут быть отправлены через конструктор класса SerialPort
        // PortName = "COM1", Baud Rate = 19200, Parity = None, Data Bits = 8, Stop Bits = One, Handshake = None
        //public SerialPort _serialPort = new SerialPort("COM2", 9600, Parity.None, 8, StopBits.One);
        private void PortIni() // инициализация: проверка порта, ...
        {
            if (sModes.IndexOf("NoComPort") >= 0)  // (NoComPort)   
            {
                lblMes1.Text = $"Тест - без Com-порта.";
                return;
            }
            _serialPort = new SerialPort(ComPortNo, 9600, Parity.None, 8, StopBits.One);
            _serialPort.Handshake = Handshake.None;
            // _serialPort.ReadTimeout = 500; // в милисекундах
            _serialPort.WriteTimeout = 500;
            _serialPort.DataReceived += new SerialDataReceivedEventHandler(Sp_DataReceived);
            // Открытие последовательного порта
            // Перед попыткой записи убедимся, что порт открыт.
            try
            {
                if (!(_serialPort.IsOpen))
                {
                    _serialPort.Open();
                    lblMes1.Text = $"Открыт порт {_serialPort.PortName}.";
                }
                else
                    lblMes1.Text += " Уже открыт!";
                    //_serialPort.Write("что-то...\r\n");
            }
            catch (Exception ex)
            {
                string mes = "Ошибка открытия порта";
                WErrLog(mes);
                ExitApp($"{mes} {_serialPort.PortName}.\n{ex.Message}", 2);
            }
        }

        void Sp_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string dataCom = _serialPort.ReadExisting();
            // Привлечение делегата на потоке UI и отправка данных, 
            // которые были приняты привлеченным методом.
            // --- Метод "CombineParts" будет выполнен в потоке UI,
            // который позволит заполнить текстовое поле RTBout.
            this.BeginInvoke(new SetTextDeleg(CombineParts),
                             new object[] { dataCom });
        }
        #endregion --- методы для Com-порта: инициализация (PortIni) и чтения (Sp_DataReceived)

        private void CombineParts(string receivedData) // ======== склеивание частей полученных данных =========
        {
            /*int rb = receivedData.Count();
            //uint tst01 = uint.Parse(receivedData);
            byte[] stb = new byte[4096];     // готовим место для принятия данных
            // 1.
            char[] ar1 = new char[receivedData.Length];
            for (int i = 0; i < receivedData.Length; i++)    ar1[i] = receivedData[i];
            // 2.
            char[] ar2 = receivedData.ToCharArray();
            //stb = Convert.ToByteArray(receivedData);
            */
            Pic_Cat.Visible = true; 
            Lbl_State.ForeColor = Color.Red;
            Lbl_State.Text = "Приём данных...";
            string st = receivedData;
            st = st.Replace("?", " "); // заменить непоняно откуда взявшиеся смиволы вопроса на пробел (FIXME костыль №1 26.08.2020 :) 
            inputString += st;
            int stLength = st.Length;
            int inputStringLength = inputString.Length;
            kPart++;

            dtm = DateTime.Now;
            if (dt0.ToString("yyyy-MM-dd") != dtm.ToString("yyyy-MM-dd")) // изменилась дата
                SetPathLog(); // изменить имя каталога и лог-файла при смене даты

            dt0 = dtm;

            if (sDebugModes.IndexOf("лог частей") >= 0) // для режима отладки - лог частей получаемых данных     
                WLog($"--- {kPart} часть, {receivedData.Count()} байт, receivedData: {receivedData}");

            int nEoT = receivedData.IndexOf(ETX);
            if (nEoT >= 0)
            {
                kTest++;
                string s1 = $"=== {kTest}-й тест. Длина приёма {inputString.Count()} байт, все данные:\n";
                string s2 = $"{inputString}";
                if (sDebugModes.IndexOf("лог приёма") >= 0)   // для режима отладки - логирование
                    WLog(s1 + s2);

                if (sDebugModes.IndexOf("лог квиточка") >= 0) // для режима отладки - логирование
                    WTest("Log_Results", inputString);

                string ss = dtm.ToString("yyyy-MM-dd HH:mm:ss").Replace("-", ".");
                Add_RTB(RTBout, $"\n{ss} Begin test {kTest}.", Color.DarkMagenta);
                Add_RTB(RTBout, $"\n{inputString}", Color.Black);

                Parse150();  // обработчик для UriLit150

                Add_RTB(RTBout, $"\n{ss} End test {kTest}, lines {kLines}.", Color.DarkMagenta);
                // очистить inputString  для повторного приёма данных
                inputString = ""; kPart = 0; kLines = 0;
                Lbl_State.Text = "... ждём-с...";
                Lbl_State.ForeColor = Color.Green;
            }
            else
            {
                msg = $"{kPart}- часть, {stLength} байт. Время:{dtm}. - msg изм. в CombineParts()";
            }
        }

        private void Parse150() // ================= ПАРСИНГ ПОЛУЧЕННЫХ ДАННЫХ для UriLit-150 ===================
        {
            #region --- описание получаемых данных
            /* Пример передаваемых данных: 
             * (в первой строке - <02h>, затем - новая строка  <0D0Ah>, <...данные... >, самый последний байт с новой строки - <03h>)
            00<02h>
            01NO.000003
            02ID:0000000003351
            03 Sex:        
            04
            05 Age:
            06 Color:        
            07 
            08 Clarity:        
            09 2020-08-04  10:26:39
            10
            11
            12 *LEU +2    125 CELL/uL
            13
            14  KET -        0 mmol/L
            15  NIT -                
            16  URO            Normal
            17 *BIL +3     100 umol/L
            18 *PRO +3      >=3.0 g/L
            19  GLU -        0 mmol/L
            20  SG         1.005     
            21 *BLD +3    200 CELL/uL
            22
            23  pH         7.5       
            24  Vc  -        0 mmol/L
            25  
            26 <03h>
                Это уже третий алгоритм разбора полученных данных! Почему-то количество реально передаваемых строк
                не соответствует количеству строк, описанной в документации. Реально не совпадает количество полученных байт!
                И иногода (спонтпнно) оно меняется! Почему - не удалось выяснить. (Результаты не воспроизводились).
                Из практического опыта исвестно, что в переданных данных должны присутствовать 11 результатов измерения
                (кроме случаев ошибки - тогда в данных должна быть строка trouble-n, где n-номер ошибки,
                и кроме калибровочных тестов, которые мне не дали делать, т.к. калибровочных тест-полосои всего две).
                Поэтому алглритм следующий:
                Из каждой полученной строчки пробуем выделить один из возможных 11 анализов.
                Если есть нужный анализ из 11-ти - формируем строку SQL, дописывая выделенные даные в конец.
                Тогда (предположительно) даже для других типов тест-полосок, на которых есть дополнительные анализы:
                - "CR", "MA", "Ca", "ACR" - этот алгоритм должен работать! (Не проверял, т.к. нет тест-полосок 14G и других)
            */
            #endregion --- описание получаемых данных
            Pic_Cat.Visible = false;
            if (F_Remind == 1) RemindText(); // обновить напоминалки в основном окне
            Lbl_State.ForeColor = Color.Blue;
            Lbl_State.Text = "парсинг...";
            string[] Line = inputString.Split(new[] {'\r', '\n', STX, ETX}, StringSplitOptions.RemoveEmptyEntries);
            //string[] Line = inputString.Split(new[] { '\r', '\n' });  // без  игнорирования пустых строк.
            kLines = Line.Count();
            msg = $"Количество строк: {kLines}.";
            Add_RTB(RTBout, msg, Color.DarkBlue);
            WLog(msg);

            //for (int i = 0; i < kLine; i++) Add_RTB(RTBout, $"\n{i}-я строка: {Line[i]}", Color.DarkBlue);
            string st = ";"; // тестовая строка - для EXCEL
            string sc = ";"; // для EXCEL PivTab 
            Boolean Fl_PivTab = (sDebugModes.IndexOf("Log_PivTab") >= 0); // для вывода в Excel's Pvot Table
            for (int i = 0; i < Line.Length; i++)
            {
                Line[i] = Line[i].Replace("?", " ");    // заменить непоняно откуда взявшиеся смиволы вопроса на пробел
            }
            // получаем всё для формирования строки SQL
            int n = 0; // номер строкИ после разбиения данных на стрОки
            string s = "---test---", cHistNo = "";
            string[] PivTab = new string[27]; // строчки для результатов, которые пойдут в сводную таблицу

            // 0-я строка содержит 1 пробел (почему?) (признаки начала и конца 02h и 03h отбросили при разбиении по строкам :)
            // 1-я строка - порядковый номер пробы (исследования) - анализатор автоматически увеличивает его на 1 для следующей пробы
            //  0123456789 123456789 1
            //01NO.000003  - это порядковый номер теста, присваиваемый анализатором, увеличивающийся на 1.
            n = 1; s = Line[n]; 
            testNo = s.Substring(3);  // порядковый номер теста, присваиваемый анализатором, и он его каждый раз увеличивет на 1.
                                      //dateDone999 += ":00." + String.Format("{0:d3}", testNo); // testNo (27) в виде "027" - для тысячных долей секунд
            testNo = testNo.Replace(" ", "0"); // заменить пробелы на нули (FIXME костыль №3 26.08.2020 :) 

            // 2-я строка - номер истории пациента (HistoryNumber)
            //  0123456789 123456789 1
            //02ID:0000000003351
            n = 2; s = Line[n];
            if (s.Substring(0, 3) != "ID:")  // это "контроли" мочевика?
            {
                msg = "Это контроль? (Нет ID:) - Контроли игнорируются!";
                Add_RTB(RTBout, msg, Color.Red);
                WLog(msg);
                return;
            }
            cHistNo = s.Substring(3);
            cHistNo = cHistNo.Replace(" ", "0"); // заменить пробелы на нули (FIXME костыль №2 26.08.2020 :) 
            cHistNo = cHistNo.TrimStart('0');  // все нули слева удалить
            nHistNo = cHistNo.Length == 0 ? (0) : (Convert.ToInt64(cHistNo));
            if (nHistNo > MaxHistNo)
            {
                string err1 = $"Введённый номер истории {nHistNo} больше максимально допустимого {MaxHistNo}!";
                Add_RTB(RTBout, $"\n\n{err1}", Color.Red);
                nHistNo = 0;
                WLog(err1);
                //MessageBox.Show(err1, "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop); // нельзя останавливать процесс работы!
            }

            if (nHistNo == 0)
            {
                string err1 = $"Не введён номер истории! No: {testNo}, дата: {dateDone}.";
                if (sDebugModes.IndexOf("лог пациентов без номера истории") >= 0) // для режима отладки - логирование
                    WTest("Log_HistNo_0", err1);
                Add_RTB(RTBout, $"\n{err1}", Color.Red);
            }

            // далее 6 строк игнорируем:  (на анализаторе результат не печатается, лаборанты пишут ручками :))
            //03 Sex:        
            //04
            //05 Age:
            //06 Color:        
            //07 
            //08 Clarity:       

            //09 2020-08-04  10:26:39
            //  0123456789 123456789 1
            n = 7; s = Line[n];
            dateDone = s.Substring(1, 20);
            //                        ДД                        ММ                        ГГГГ                        
            dateDone999 = s.Substring(9, 2) + "-" + s.Substring(6, 2) + "-" + s.Substring(1, 4) + " "
                + s.Substring(13, 2) + ":" + s.Substring(16, 2) + ":" + s.Substring(19, 2);  // надо в формате ДД-ММ-ГГГГ ЧЧ:ММ:СС.МММ
            //                ЧЧ                         ММ                         CC                        
            //    + s.Substring(13, 2) + ":" + s.Substring(16, 2) + ":" + s.Substring(19, 2) + ".000"; // надо в формате ДД-ММ-ГГГГ ЧЧ:ММ:СС.МММ
            //dateDone999 += ":" + s1.Substring(0, 2) + "." + s1.Substring(2, 2) + "0";
            st += $";{dateDone}";
            sc += $";{dateDone}";  // для EXCEL PivTab

            st += $";{nHistNo};";
            sc += $";{nHistNo};";  // для EXCEL PivTab
            //s1 = "INSERT into AnalyzerResults(Analyzer_id,ResultDate,ResultText,HostName,CntParam,HistoryNumber";
            s1 = "INSERT into AnalyzerResults(Analyzer_id,ResultDate,HostName,CntParam,HistoryNumber";
            string si = inputString.Substring(1, inputString.Length - 2);
            si = si.Replace("\r", "");
            si = si.Replace("\n", "");
            //s2 = $"VALUES({AnalyzerID},'{dateDone999}','{si.Replace($"\r\n", "__")}',host_name(),10,{nHistNo.ToString()}";// CntParam=10 !
            // FIXME ?передаём в SQL без разделителей строк и признаков начала и конца передачи
            //s2 = $"VALUES({AnalyzerID},'{dateDone999}','{inputString.Replace($"\r\n"+STX+ETX, "")}',host_name(),10,{nHistNo.ToString()}";// CntParam=10 !
            //s2 = $"VALUES({AnalyzerID},'{dateDone999}','{si}',host_name(),10,{nHistNo.ToString()}";// CntParam=10 !
            s2 = $"VALUES({AnalyzerID},'{dateDone999}',host_name(),11,{nHistNo.ToString()}";    // ВНИМАНИЕ! CntParam=11 всегда - для тест-полосок G11. 

            FindAndConcatSqlSring();

            // записать строчки для результатов в лог для сводной таблицы
            if (Fl_PivTab) for (int i = 1; i <= 10; i++) WTest("Log_PivTab", PivTab[i]);

            // тестовая строка для выгрузки в EXCEL
            if (sDebugModes.IndexOf("Log_Excel") >= 0)
            {
                //string ss1 = "\n"+'\x0D'.ToString();
                //string ss2 = st.Replace(ss1, "__");
                WTest("Log_Excel", st.Replace($"\r\n", "__"));
            }

            // 2020-03-13 st = ";";    // очистили для следующего пациента для выгрузки в EXCEL

            // закрыть скобки
            s1 += ",ResultText)"; // далее идёт VALUES(...  
            s2 += $", 'Номер на анализаторе: {testNo}, дата-время: {dateDone}.'); ";
            if (sDebugModes.IndexOf("лог SQL") >= 0)
                WTest($"Log_SQL", s1 + s2);

            if (sModes.IndexOf("WriteToSQL=Yes") >= 0 & kAn>0 ) // строка SQL сформирована успешно
            {
                WLog("SQL " + s1 + s2);
                ToSQL(s1 + s2);
                msg = $"SQL: Номер истории {nHistNo}, дата {dateDone999}. Номер пробы: {testNo}.";
                Add_RTB(RTBout, $"\n{msg}\n", Color.DarkGreen);
            }

            // добавить в конец строки SQL полученный результат измерения (всего 11 измерений для тест-полосок G11)
            //void ConcatSqlSring(int nParm, string cnam, string nam, string val, string mgr)
            void FindAndConcatSqlSring()
            {
                // найти нужные анадизы и сформировать SQL-строку
                //throw new NotImplementedException();
                //rc = 1; // RetCode = 1 по умолчанию
                int k = 0; // номер параметра (результата) для формирования строки SQL - ParamName1 (n=1), ParamValue1, ParamMsr1 ...
                string sx = "+++"; // одна (текущая) строчка Line[i] из квиточка
                string val, nam, mgr, attention;
                for (int i = 8; i < Line.Length; i++)  // не с первой: строки с результатами начинаются с ~8..11-й. FIXME !
                {
                    sx = Line[i];
                    if (sx.Trim().Length<=1)    // =1 для последней строки == 03h
                    {
                        msg = $"Строка {i}: пустая.";
                        WLog(msg);
                        //Add_RTB(RTBout, $"\n{msg}", Color.Red);
                        continue; // пропускаем пустую строку
                    }

                    nam = "";
                    val = "";
                    mgr = "";
                    attention = sx.Substring(1, 1);  // признак: " " -норма, "*" - результат выходит за референсные значения 

                    nam = sx.Substring(2, 3);
                    if (nam == "LEU" | nam == "KET" | nam == "BIL" | nam == "GLU" | nam == "BLD" | nam == "Vc ")
                    {
                        val = sx.Substring(5, 11);
                        mgr = sx.Substring(16);
                    }

                    if (nam == "NIT" | nam == "URO" | nam == "SG " | nam == "pH ")
                    {
                        val = sx.Substring(5);
                        mgr = "";
                    }

                    if (nam == "PRO" | nam == "CR " | nam == "MA " | nam == "Ca " | nam == "ACR")
                    {
                        val = sx.Substring(5, 15);
                        mgr = sx.Substring(20);
                    }

                    if (val.Length == 0)   // нет результата, ничего не нашли
                         continue; // пропускаем неопознанную строку
 
                    // логирование белка
                    if (((nam == "PRO") & sModes.IndexOf("лог белка") >= 0) & (sx.Substring(1, 1) == "*"))
                        WTest("Log_PROTEIN", $"Номер истории {nHistNo}, No: {testNo}, дата: {dateDone}, {s}");
                    #region --- old calls
                    ////  0123456789 123456789 123 // 12-я строка: ЛЕЙКОЦИТЫ   
                    ////12 *LEU +2    125 CELL/uLg              
                    //ConcatSqlSring(12, "LEU", s.Substring(2, 3), s.Substring(5, 11), s.Substring(16)); if (rc < 1) return;
                    ////14  KET -        0 mmol/L
                    //ConcatSqlSring(14, "KET", s.Substring(2, 3), s.Substring(5, 11), s.Substring(16)); if (rc < 1) return;
                    ////  0123456789 123456789 123 // 15 строка - НИТРАТЫ 
                    ////15  NIT -                    
                    //ConcatSqlSring(15, "NIT", s.Substring(2, 3), s.Substring(5), ""); if (rc < 1) return;
                    ////  0123456789 123456789 123 // 16 строка - Уробилиноген
                    ////16  URO Normal
                    //ConcatSqlSring(16, "URO", s.Substring(2, 3), s.Substring(5), ""); if (rc < 1) return;
                    ////  0123456789 123456789 123 // 17 строка - БИЛИРУБИН
                    ////17 *BIL +3     100 umol/L
                    //ConcatSqlSring(17, "BIL", s.Substring(2, 3), s.Substring(5, 12), s.Substring(17)); if (rc < 1) return;
                    ////  0123456789 123456789 123 // 18 строка - ПРОТЕИН - БЕЛОК
                    ////18 *PRO +3      >=3.0 g/L
                    //ConcatSqlSring(18, "PRO", s.Substring(2, 3), s.Substring(5, 15), s.Substring(20)); if (rc < 1) return;
                    ////  0123456789 123456789 123 // 19 строка --ххххх--
                    ////19  GLU -        0 mmol/L
                    //ConcatSqlSring(19, "GLU", s.Substring(2, 3), s.Substring(5, 12), s.Substring(17)); if (rc < 1) return;
                    ////  0123456789 123456789 123 // 20 строка - УДЕЛЬНЫЙ ВЕС (ПЛОТНОСТЬ)
                    ////20  SG         1.005     
                    //ConcatSqlSring(20, "SG", s.Substring(2, 2), s.Substring(4), ""); if (rc < 1) return;
                    ////  0123456789 123456789 123 // 21 строка - BLOOD - КРОВЬ
                    ////21 *BLD +3    200 CELL/uL
                    //ConcatSqlSring(21, "BLD", s.Substring(2, 3), s.Substring(5, 11), s.Substring(16)); if (rc < 1) return;
                    ////23  pH         7.5     
                    //ConcatSqlSring(23, "pH", s.Substring(2, 2), s.Substring(4), ""); if (rc < 1) return;
                    ////24  Vc  -        0 mmol/L
                    //ConcatSqlSring(24, "Vc", s.Substring(2, 2), s.Substring(4, 12), s.Substring(17)); if (rc < 1) return;
                    #endregion --- old calls

                    nam = nam.Trim();
                    val = val.Trim().Replace("    ", " "); //т.к. val для URO не влезает в 16 символов (ограничение длины поля в SQL) 
                    mgr = mgr.Trim();
                    k++;   // k-тый параметр Param<k>
                    //s1 += $",ParamName{k},ParamValue{k},ParamMsr{k}";
                    //s2 += $",'{nam}','{val}','{mgr}'";
                    s1 += $",ParamName{k},ParamValue{k},ParamMsr{k},Attention{k}";  // 2020-10-20
                    s2 += $",'{nam}','{val}','{mgr}','{attention}'";

                    //st += $";{nam}";  //2020-08-20
                    //sc += $";{nam}";  // для EXCEL PivTab
                    //st += $"{nam};{val};{mgr};";  // 2020-10-20
                    st += $"{nam};{val};{mgr};{attention};";
                    WLog($"Строка {i} (длина {sx.Length}): {sx}. Выделен {k}-й: nam='{nam}', val='{val}', mgr='{mgr}', attention='{attention}'.");
                    PivTab[k] = sc + $"{nam};{val};{mgr};";
                }

                if (k == 0)   // нет результата, ничего не нашли
                {
                    kAn = 0;    // кол-во выделенных анализов - используется как признак, писАть ли в SQL.
                    s1 = ""; s2 = "";   // SQL строку не формируем 
                    msg = "Не нашли ни одного результата!";
                    Add_RTB(RTBout, "\n"+msg, Color.Red);
                    WLog(msg);
                    return;
                }
                kAn = k;
            }
        }
        // ---
        private void ReadParmsIni()   // читать настройки из ini-файла 
        {
            PathIni = Application.StartupPath;
            AppName = AppDomain.CurrentDomain.FriendlyName;
            AppName = AppName.Substring(0, AppName.IndexOf(".exe"));
            pathIniFile = PathIni + @"\" + $"{AppName}" + ".ini";
            if (!File.Exists(pathIniFile))
            {
                string errmsg = "Не найден файл " + pathIniFile;
                //WErrLog(errmsg);  /3/ А потому что нет ещё пути, где ErrLog! :)))
                MessageBox.Show(errmsg, " Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                Environment.Exit(1);
            }
            //string dtVers = new System.IO.FileInfo(PathIni).CreationTime.ToString();
            string dtVers = new System.IO.FileInfo(AppName+".exe").LastWriteTime.ToString();
            AppVer += $" от {dtVers}.";  // 2020-09-01
            Lbl_dtm_ver.Text = AppVer;
            ///*
            //var version = Assembly.GetExecutingAssembly().GetName().Version;
            //var builtDate = new DateTime(2019, 1, 1).AddDays(version.Build).AddSeconds(version.Revision * 2);
            //var versionString = String.Format("версия {0} изм. {1}", version.ToString(2), builtDate.ToShortDateString());
            // из AutoVersion 2:
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var fileVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(assemblyLocation).FileVersion;
            var versionString = version + ", " + fileVersion;
            Lbl_AutoVersion.Text = versionString;
            //*/
            lbl_dtm.Text = "Время старта: " + dt0.ToString();  //+ToDo: вывод на форму - время старта и путь ini-файла 
            lbl_ini.Text = "ini-файл: " + pathIniFile;
            string[] lines = File.ReadAllLines(pathIniFile, Encoding.GetEncoding(1251));

            //  2-я строка: AnalyzerID: 13
            string s = lines[1];    // 2-я строкa
            AnalyzerID = Convert.ToInt32(s.Substring(s.IndexOf(':') + 1, 2));
            //if (int.TryParse(lines[5], out AnalyzerID) ) AnalyzerID = 339;
            //AnalyzerID = 
            Lbl_AnalyserID.Text = $"AnalyzerID: {AnalyzerID}";

            ComPortNo = lines[2];   // 3-я строка 
            connStr = lines[3];     // 4-я строка
            string nameSQLsrev = connStr.Substring(0, connStr.IndexOf(';'));
            Lbl_SQL.Text = nameSQLsrev;     // Data Source=asu-911; или bsmp-server0
            MaxCntAn = Convert.ToInt32(lines[4].Trim());    // 5-я строка
            sModes = lines[5].Trim();                       // 6-я строка: Режимы работы 
            Lbl_sModes.Text = sModes;
            // (sModes.IndexOf("Все анализы одного пациента в одну строку") >= 0);
            // (sModes.IndexOf("WriteToSQL=Yes") >= 0);   
            // (sModes.IndexOf("Дата в SQL") >= 0);  // (GetDate)         
            sDebugModes = lines[6].Trim();           // 7-я строкa: Режимы отладки 
            Lbl_DebugModes.Text = sDebugModes;
            MaxHistNo = Convert.ToInt32(lines[7].Trim());          // 8-я строка:  макс. номер истории
            PathLogParm = lines[8].Trim();  // 9-я строка 
            Lbl_log.Text = "путь к логам: " + PathLogParm;
            SetPathLog();
            PathErrLog = lines[9];          // 10-я строка 
            string sRemind = lines[10];  // 11-я строка: флаг изменения режима отображения напоминалок: 1 - показывать, иначе - нет.
            int.TryParse(sRemind, out F_Remind); // F_Remind == 0 or 1
                               // 12-я строка: <резерв>
            qkrq  = lines[12]; // 13-я строка  ToDo ! для ответа на запрос о рaботе = доделать!
            strV1 = lines[13]; // 14-я строка
            strV2 = lines[14]; // 15-я строка
            strV3 = lines[15]; // 16-я строка
            strV4 = lines[16]; // 17-я строка
            strV5 = lines[17]; // 18-я  cтрока
            Lbl_v1.Text = "";
            Lbl_v2.Text = "";
            Lbl_v3.Text = "";
            Lbl_v4.Text = "";
            Lbl_v5.Text = "";
            if (F_Remind == 1)
            {
                Lbl_v1.Text = strV1;
                Lbl_v2.Text = strV2;
                Lbl_v3.Text = strV3;
                Lbl_v4.Text = strV4;
                Lbl_v5.Text = strV5;
            }

            Lbl_Comp_User.Text = $"Computer: {ComputerName}, User: {UserName}";
            WLog($"--- Запуск {AppName} {ComputerName} {UserName} {dtVers}");
            //WLog($"--- ini-файл: {fnPathIni}");
            //string parmIni = "";
            //for (int i = 0; i < 19; i++) parmIni += "\n" + lines[i];
            //WLog("--- параметры в ini-файле:" + parmIni);
        }
        private void ToSQL(string st) // запись в MS-SQL сформированной строки
        {
            //using (SqlConnection sqlConn = new SqlConnection(Properties.Settings.Default.connStr))
            using (SqlConnection sqlConn = new SqlConnection(connStr))
            {
                //SqlCommand sqlCmd = new SqlCommand("INSERT INTO [AnalyzerResults]([Analyzer_Id],[ResultText],[ResultDate]
                // ,[Hostname],[HistoryNumber])VALUES(@AnalyzerId,@ResultText,GETDATE(),@PCname,@HistoryNumber)", sqlConn);
                SqlCommand sqlCmd = new SqlCommand(st, sqlConn);
                try
                {
                    //Lbl_State.Text = "запись в SQL...";
                    sqlCmd.CommandType = System.Data.CommandType.Text;
                    sqlConn.Open();
                    sqlCmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    string mes = $"Ошибка при записи в SQL. Номер истории: {nHistNo}, номер теста: {testNo}."; //2019-11-13
                    WErrLog(mes + "\n" + ex.ToString());    // в файл ошибок...
                    WLog(mes + "\n" + ex.ToString());       // в лог файл тоже!
                    Add_RTB(RTBout, $"\n!!! {mes} !!!", Color.Red);
                    //ExitApp(mes, 3);
                }
            }
        }
        #region --- ( Easter eggs :))
        private void RemindText() // Текст напоминалок - отображается в основном окне
        {
            str_ini = File.ReadAllLines(pathIniFile, Encoding.GetEncoding(1251));
            // для случайного
            int k1 = 0, k2 = 0, k3 = 0, k4 = 0;
            for (int i = 0; i < str_ini.Length; i++)
            {
                if (str_ini[i].IndexOf("*** start test ***") != -1) k1 = i; // ограничитель начала напоминалок
                if (str_ini[i].IndexOf("*** end test ***") != -1) k2 = i; // ограничитель конца  напоминалок
            }
            if ((k1 < k2) & (k1 != 0) & (k2 != 0))
            {
                k1++; k2--;  // текст по 5 строк с k1 - начало, по k2 - конец  (весь текст в ini-файле :)
                //Random rnd = new Random();       // сразу в глобальных c инициализацией
                int irand = rnd.Next(0, k2 - k1);  //очередное  случайное число в диапазоне k1 - k2.
                k3 = irand / 5;
                k4 = k1 + k3 * 5; // по 5 строк на "напоминалку" - strV1-strV5 
                strV1 = str_ini[k4];
                strV2 = str_ini[k4 + 1];
                strV3 = str_ini[k4 + 2];
                strV4 = str_ini[k4 + 3];
                strV5 = str_ini[k4 + 4];
                Lbl_v1.Text = strV1;
                Lbl_v2.Text = strV2;
                Lbl_v3.Text = strV3;
                Lbl_v4.Text = strV4;
                Lbl_v5.Text = strV5;
                Lbl_v1.Show();
                Lbl_v2.Show();
                Lbl_v3.Show();
                Lbl_v4.Show();
                Lbl_v5.Show();
                //Stat1.Text = $"ir={irand}, k1={k1}, k2={k2}, k3={k3}, k4={k4}.";
            }
        }
        private void PictureBox2_Click(object sender, EventArgs e)
        {
            //pictureBox2.Visible = !pictureBox2.Visible;
        }

        private void PictureBox1_Click(object sender, EventArgs e)
        {
            //pictureBox2.Visible = !pictureBox2.Visible;
        }
        private void Pic_Cat_Click(object sender, EventArgs e)
        {

        }
        private void Lbl_dtm_ver_Click_1(object sender, EventArgs e)
        {
            Pic_Cat_Click(sender, e);
        }
        private void Pic_An_Click(object sender, EventArgs e)
        {
            Pic_Cat_Click(sender, e);
        }
        private void Lbl_dtm_ver_Click(object sender, EventArgs e)
        {
            Pic_Cat_Click(sender, e);
        }
        #endregion --- ( Easter eggs :))
        #region --- методы Wlog, WErrLog, WTest;  SetPathLog, Add_RTB, ExitApp...
        private static void SetPathLog()     // нужна, если программа работает много дней и меняется текущая дата //2019-08-06
        {
            dtm = DateTime.Now;
            PathLogParm = Path.GetFullPath(PathLogParm + @"\.");
            string PathLogGodMes = PathLogParm + @"\" + dtm.ToString("yyyy-MM");
            if (!Directory.Exists(PathLogGodMes))
                Directory.CreateDirectory(PathLogGodMes);
            //PathLog += @"\BeckLog" + $"{dtm.Year}-" +
            //    $"{dtm.Month.ToString().PadLeft(2, '0')}-" +
            //    $"{dtm.Day.ToString().PadLeft(2,'0')}.txt";
            PathLog = PathLogGodMes + @"\" + $"{AppName}_" + dtm.ToString("yyyy-MM-dd") + ".txt";
        }
        private static void WLog(string st) // записать в лог FLog
        {
            FileStream fn = new FileStream(PathLog, FileMode.Append);
            StreamWriter sw = new StreamWriter(fn, Encoding.GetEncoding(1251));
            dtm = DateTime.Now;
            string ss = dtm.ToString("yyyy-MM-dd HH:mm:ss").Replace("-", ".");
            sw.WriteLine($"{ss} {st}");
            sw.Close();
        }
        private static void WErrLog(string st) // записать ErrLog
        {
            string fnPathErrLog = PathErrLog + @"\Log_ERR.txt";
            fnPathErrLog = Path.GetFullPath(fnPathErrLog);
            FileStream fn = new FileStream(fnPathErrLog, FileMode.Append);
            StreamWriter sw = new StreamWriter(fn, Encoding.GetEncoding(1251));
            dtm = DateTime.Now;
            string ss = dtm.ToString("yyyy-MM-dd HH:mm:ss").Replace("-", ".");
            sw.WriteLine($"\n{ss} {st}");
            sw.Close();
        }
        private static void WTest(string FileNam, string st) // записать в FileNam.txt
        {
            FileStream fn = new FileStream(PathErrLog + "\\" + FileNam + ".txt", FileMode.Append);
            StreamWriter sw = new StreamWriter(fn, Encoding.GetEncoding(1251));
            dtm = DateTime.Now;
            string ss = dtm.ToString("yyyy-MM-dd HH:mm:ss").Replace("-", ".");
            sw.WriteLine($"\n{ss} {st}");
            sw.Close();
        }

        private static void Add_RTB(RichTextBox rtbOut, string addText)
        {
            Add_RTB(rtbOut, addText, Color.Black);
        }

        private static void Add_RTB(RichTextBox rtbOut, string addText, Color myColor)
        {
            Int32 p1, p2;
            p1 = rtbOut.TextLength;
            p2 = addText.Length;
            rtbOut.AppendText(addText);
            rtbOut.Select(p1, p2);
            rtbOut.SelectionColor = myColor;
            // 1 rtbOut.Select(0, 0);
            // 2 rtbOut.Select(p1 + p2, 0);
            // 2 rtbOut.AppendText("");
            rtbOut.SelectionStart = rtbOut.Text.Length;
            rtbOut.ScrollToCaret();
            // или: rtbOut.Select(p1, p2);
            //      SendKeys.Send("^{END}");  // это прокрутка в конец :)
        }

        private static void ExitApp(string mess, int ErrCode = 1001) // Завершение работы по ошибке.
        {
            string title = "Аварийное завершение работы.";
            WErrLog($"{title}\n{mess}");
            MessageBox.Show(mess, title
                , MessageBoxButtons.OK, MessageBoxIcon.Stop);
            Environment.Exit(ErrCode);
        }
        private static void ExitApp(string mess) // Нормальное завершение работы (по кнопке Х)
        {
            //    WLog("--- передумал выходить :) ");
            //    return; // просто передумал :)
            WLog("--- " + mess);
            Environment.Exit(0);
        }

        private void RTBout_TextChanged(object sender, EventArgs e)
        {
            // Nohing to do! :))
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //MessageBox.Show("mess - Что нажали?", "title - заголовок"
            //    , MessageBoxButtons.OK, MessageBoxIcon.Stop);
        }

        private void ПараметрыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Меню / Сервис / Параметры в.ini - файле:
            string s = "";
            string nameSQLsrv = connStr.Substring(0, connStr.IndexOf(';'));
            string sep = $"---------------------------------------------------------------------------\n";
            // 2020-04-07 из надстойки Automatic Version 2
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var fileVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(assemblyLocation).FileVersion;
            var productVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(assemblyLocation).ProductVersion;
            //s += $"         {sHeader1}\n";
            int n1 = connStr.IndexOf('=') + 1;
            int n2 = connStr.IndexOf(';') - n1;
            s += $"AppName: {AppName}\n";
            s += $"AnalyzerID: {AnalyzerID.ToString()}\n";
            s += $"ComPortNo: {ComPortNo}\n";
            s += $"Server: {connStr.Substring(n1, n2)}\n";
            s += sep;
            s += $"ComputerName: {ComputerName}, UserName: {UserName}\n";
            s += sep;
            s += $"путь к логам:\n";
            s += $"PathLog: {PathLogParm}\n";
            s += $"PathErrLog: {PathErrLog}\n";
            s += $"PathIni: {PathIni}\n";
            s += sep;
            s += $"Режимы работы: {sModes}\n";
            //s += $"SQL: {nameSQLsrv}, Analyzer_Id: {Analyzer_Id}\n";
            s += sep + "\n\n\n\n";
            DialogResult result = MessageBox.Show(s, "  Параметры в .ini-файле:"
                , MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ВыходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // ExitApp("Выход по меню <Выход>");
            string mess = "Завершение работы драйвера анализатора.";
            DialogResult result = MessageBox.Show("Вы действительно хотите завершить работу?"
                , mess, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                ExitApp("Выход из меню<Выход>");
            }
            else
            {
                WLog("--- хотел выйти из меню <Выход>, но передумал :) ");
            }
        }

        private void ОпрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s = $"Описание драйвера смотрите в файле \n'ДРАЙВЕР АНАЛИЗАТОРА МОЧИ UILIT-150.docx' ";
            DialogResult result = MessageBox.Show(s, "  О пограмме:"
                , MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)  // закрыть App по крестику X
        {
            string mess = "Завершение работы драйвера анализатора.";
            DialogResult result = MessageBox.Show("Вы действительно хотите завершить работу? \n\n   Результаты передаваться не будут!"
                , mess, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                WLog("--- хотел выйти, но передумал :) " + e.CloseReason.ToString()
                    + " " + result.ToString()); 
                e.Cancel = true;
            }
            else
            {
                WLog("--- " + mess);
            }

        }
        private void ReRead_ini_File()
        {
            ReadParmsIni();   // читать настройки из ini-файла 
            string[] lines = File.ReadAllLines(pathIniFile, Encoding.GetEncoding(1251));
            string st = "";
            for (int i = 0; i < lines.Length; i++)
                st += $"{i + 1} :  " + lines[i] + "\n";
            MessageBox.Show(st, " Параметры в .ini-файле:");
        }
        #endregion --- методы Wlog, WErrLog;  SetPathLog, Add_RTB, ExitApp...
        #region --- Тесты по кнопке Выполнить
        private void BtnRunTest_Click(object sender, EventArgs e)
        {
            string testSelect = CmbTest.SelectedItem.ToString();
            int nSelected = CmbTest.SelectedIndex;

            switch (testSelect)
            {
                case "перечитать ini-файл":
                    ReRead_ini_File();
                    break;
                case "001":
                    string cHistNo = "100 000062060";    // 13 digits
                    //cHistNo = "0000000000000";
                    cHistNo = cHistNo.Replace(" ", "0"); // заменить пробелы на нули (FIXME костыль №2 26.08.2020 :) 
                    Add_RTB(RTBout, $"\n cHistNo: {cHistNo}.\n", Color.BlueViolet);
                    cHistNo = cHistNo.TrimStart( '0' );  // все нули слева удалить
                    nHistNo = cHistNo.Length == 0 ? (0) : (Convert.ToInt64(cHistNo));
                    Add_RTB(RTBout, $"\n cHistNo: {cHistNo}, nHistNo: {nHistNo}.", Color.Red);
                    if (nHistNo > MaxHistNo)
                    {
                        Add_RTB(RTBout, $"\n\n Введённый номер истории {nHistNo} больше максимально допустимого {MaxHistNo}!", Color.Red);
                        nHistNo = 0;
                    }
                    break;
                case "0A1":
                    //4 UBG  Normal 3.4umol/L
                    // 0123456789012345678901
                    string s = " UBG  Normal 3.4umol/L";
                    string namUBG = s.Substring(1, 3);
                    //string rstUBG = s.Substring(4, s.Length - 4);
                    string mgrUBG = s.Substring(s.Length - 6, 6);// последние 6 имволов, 6=длина "umol/L"
                    string valUBG = s.Substring(4, s.Length - (4 + 6)).Trim(); // остаток посередине
                    //string[] ms = Regex.Split(s, "\n\\s*");
                    Add_RTB(RTBout, $"\n nam: {namUBG},\n val: {valUBG},\n mgr: {mgrUBG}.", Color.Red);
                    break;
                case "002":
                    //4 UBG  Normal 3.4umol/L
                    // 0123456789012345678901
                    int i1 = 9, i2 = 12, i3 = 321, i4 = 4321;
                    //string st99 = String.Format("{0:d3}",testNo);
                    string s1 = "", s2 = "", s3 = "", s4 = "", s0 = "";
                    //s1 = i1.ToString("##-##");
                    s1 = i1.ToString("D4");
                    s2 = i2.ToString("D4");
                    s3 = i3.ToString("D4");
                    s4 = i4.ToString("D4");
                    s0 = ":" + s1.Substring(0, 2) + "." + s1.Substring(2, 2) + "0";
                    s0 = ":" + s2.Substring(0, 2) + "." + s2.Substring(2, 2) + "0";
                    s0 = ":" + s3.Substring(0, 2) + "." + s3.Substring(2, 2) + "0";
                    s0 = ":" + s4.Substring(0, 2) + "." + s4.Substring(2, 2) + "0";
                    break;
                case "003":
                    kTest++; dtm = DateTime.Now;
                    Add_RTB(RTBout, $"00 Test00, номер {kTest} КРАСНЫЙ\n", Color.Red);
                    Add_RTB(RTBout, $"{dtm.ToString("yyyy-MM-dd HH:mm:ss")} ЗЕЛЁНЫЙ \n", Color.Green);
                    break;
                case "004":
                    kTest++;
                    dt0 = DateTime.Now;
                    string st1 = dt0.ToString("yyyy-MM-dd");
                    string st2 = dt0.ToString("HH:mm:ss") + ".000";
                    break;
                /*string st = "R,05-08-2019,1908051005,ЕЛИСЕЕВА Е А,44463,ОЖОГ,,Admin,,Общий белок,Колич.,46.4,г/л";
                string[] pt;
                //char[] sep = { ",", "," };
                pt = st.Split(new char[] { ',' }, 13);
                break;
                */
                case "RTBout 1":
                    kTest++; dtm = DateTime.Now;
                    Add_RTB(RTBout, $"1 номер {testNo}" + dtm.ToString("yyyy-MM-dd HH:mm:ss") + $"\n");
                    break;
                case "RTBout 2":
                    kTest++; dtm = DateTime.Now;
                    Add_RTB(RTBout, $"2 номер {testNo}" + dtm.ToString("yyyy-MM-dd HH:mm:ss") + $"\n");
                    break;
                case "01. Тест 01.":
                    kTest++; dtm = DateTime.Now;
                    Add_RTB(RTBout, $"номер {testNo} КРАСНЫЙ\n", Color.Red);
                    Add_RTB(RTBout, $"{dtm.ToString("yyyy-MM-dd HH:mm:ss")} ЗЕЛЁНЫЙ \n", Color.Green);
                    break;
                case "02. Тест 02.":
                    Add_RTB(RTBout, $"test 02{dtm.ToString("yyyy-MM-dd HH:mm:ss")}\n", Color.Pink);
                    break;
                case "03. test03":
                    string AppName = AppDomain.CurrentDomain.FriendlyName;
                    Add_RTB(RTBout, $"03. {AppName} --- {dtm.ToString("yyyy-MM-dd HH:mm:ss")}\n", Color.Navy);
                    break;
                default:
                    MessageBox.Show($"Нет теста для:  {CmbTest.SelectedItem}\n");
                    break;
            }
        }
        #endregion --- Тесты по кнопке Выполнить
    }
}
