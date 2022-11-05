using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
//using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop.Excel;
//using Excel;

namespace Itogovie_Test
{
    public partial class Glavnoe : Form
    {
        public Glavnoe()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (vibor_banka.Text == "")//Проверка выбран банк или нет
            {
                MessageBox.Show("Выберите банк или калькулятор");
            }
            else
            {
                if (vidi.Text == "" || sum_credit.Text == "" || srok_credit.Text == "" || procent_kredit.Text == "")
                {
                    MessageBox.Show("Заполните все поля");
                }
                else
                {
                    double sum = double.Parse(sum_credit.Text);        //Общая сумма
                    double srok = double.Parse(srok_credit.Text);      //Срок взятия
                    double procent = double.Parse(procent_kredit.Text);//Процент введенные
                    double ezemeh = sum / srok;//Ежемесячный платеж
                    int chet = 0;//Счетчик
                    double sum1 = 0;
                    string plata_txt;
                    double plata = 0;//Общая плата
                    double procent_sum = 0;//Проценты годовые
                    double itog_proc = 0;
                    double sum_itog = sum;
                    double anyit = 0;
                    double anyit_mes = 0;
                    string sum_txt;
                    DateTime today = DateTime.Now;
                    if (procent > 365)
                    {
                        MessageBox.Show("Процеты больше 365%");
                    }
                    else
                    {
                        if (vidi.Text == "Дифференцированый")//вычисляем дифференцированый кредит
                        {
                            tabel2.Rows.Clear();
                            while (srok > chet)
                            {
                                procent_sum = sum * procent / 100 * 31 / 365;
                                procent_sum = Math.Round(procent_sum, 2);
                                plata = procent_sum + ezemeh;
                                sum -= ezemeh;
                                itog_proc += procent_sum;
                                chet++;
                                DateTime answer = today.AddMonths(1);
                                tabel2.Rows.Add(answer, Math.Round(plata, 2), Math.Round(procent_sum, 2), Math.Round(ezemeh, 2), Math.Round(sum, 2));//Заполняем таблицу
                                today = answer;//Заполняем таблицу
                            }

                            plata_txt = Convert.ToString(Math.Round(itog_proc));
                            sum_itog += itog_proc;
                            sum_txt = Convert.ToString(Math.Round(sum_itog));
                            if (itog_proc < 0.4)
                            {
                                pereplat_credit.ForeColor = Color.Black;
                                pereplat_credit.Text = plata_txt;
                                vs9_suma_credit.Text = sum_txt;
                            }
                            else
                            {
                                pereplat_credit.ForeColor = Color.Red;
                                pereplat_credit.Text = plata_txt;
                                vs9_suma_credit.Text = sum_txt;
                            }

                        }
                        else if (vidi.Text == "Аннуитетный")//вычисляем аннуитетный кредит
                        {

                            tabel2.Rows.Clear();
                            procent_sum = Math.Round(procent / 12 / 100, 7);
                            anyit = Math.Round((procent_sum * (Math.Pow((1 + procent_sum), srok)) / (Math.Pow((1 + procent_sum), srok) - 1)), 7);
                            anyit_mes = anyit * sum;
                            sum1 = anyit_mes * srok - sum;
                            plata = anyit_mes * srok;
                            while (srok > chet)
                            {
                                ezemeh = Math.Round(sum * procent_sum, 2);
                                sum_itog = Math.Round(anyit_mes - ezemeh, 2);
                                if (sum < anyit_mes)
                                {
                                    sum_itog = Math.Round(sum, 0);
                                    sum = sum - Math.Round(sum_itog, 2);
                                }
                                else
                                {
                                    sum = sum - Math.Round(sum_itog, 0);
                                }

                                ++chet;

                                DateTime answer = today.AddMonths(1);//Это для изменения месяца
                                tabel2.Rows.Add(answer, Math.Round(anyit_mes, 2), Math.Round(ezemeh, 2), Math.Round(sum_itog, 2), Math.Round(sum, 2));//Заполняем таблицу
                                today = answer;
                            }
                            plata_txt = Convert.ToString(Math.Round(sum1));
                            sum_txt = Convert.ToString(Math.Round(plata));
                            if (sum1 < 0.4)
                            {
                                pereplat_credit.ForeColor = Color.Black;
                                pereplat_credit.Text = plata_txt;
                                vs9_suma_credit.Text = sum_txt;
                            }
                            else
                            {
                                pereplat_credit.ForeColor = Color.Red;
                                pereplat_credit.Text = plata_txt;
                                vs9_suma_credit.Text = sum_txt;
                            }
                            save.Enabled = true;
                        }
                    }
                }

            }
           
        }


        private void srok_credit_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar)))
            {
                if (e.KeyChar != (char)Keys.Back && e.KeyChar > 120)
                {
                    e.Handled = true;
                    MessageBox.Show("Вы вводите не то", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }
        
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)//Выставление процентов по вкладу
        {
            if (vidi.Text == "1")
            {
                proc_vklad.Text = "7";
            }
            else if (vidi.Text == "3")
            {
                proc_vklad.Text = "6";
            }
            else if (vidi.Text == "6")
            {
                proc_vklad.Text = "5";
            }
            else if (vidi.Text == "12")
            {
                proc_vklad.Text = "8";
            }
        }
        
        private void button3_Click_1(object sender, EventArgs e)
        {

            if (vibor_banka.Text == "")//Проверка выбран ли банк
            {
                MessageBox.Show("Выберите банк или калькулятор");
            }
            else
            {
                if (por9dok.Text == "" || sum_ipotek.Text == "" || srok_ipotek.Text == "" || procent_ipotek.Text == "")//Проверка заполнены ли поля
                {
                    MessageBox.Show("Заполните все поля");
                }
                else
                {
                    double vznos = 0;
                    if (vznos_ipotek.Text == "")
                    {
                        vznos_ipotek.Text = "0";
                    }
                    else
                    {
                        vznos = double.Parse(vznos_ipotek.Text);
                    }
                    DateTime today = DateTime.Now;
                    double sum = double.Parse(sum_ipotek.Text);        //Общая сумма
                    double srok = double.Parse(srok_ipotek.Text);      //Срок взятия
                    double procent = double.Parse(procent_ipotek.Text);//Процент введенные
                    double ezemeh = (sum - vznos) / srok;//Ежемесячный платеж
                    int chet = 0;//Счетчик
                    double sum1 = 0;
                    string plata_txt;
                    double plata = 0;//Общая плата
                    double procent_sum = 0;//Проценты годовые
                    double itog_proc = 0;
                    double sum_itog = sum;
                    double anyit = 0;
                    double anyit_mes = 0;
                    string sum_txt;
                    double sum2 = sum - vznos;
                    if (por9dok.Text == "Дифференцированый")//вычисляем дифференцированную ипотеку
                    {
                        table_ipotek.Rows.Clear();
                        while (srok > chet)
                        {
                            procent_sum = sum2 * procent / 100 * 31 / 365;
                            procent_sum = Math.Round(procent_sum, 2);
                            plata = procent_sum + ezemeh;
                            sum2 -= ezemeh;
                            itog_proc += procent_sum;
                            chet++;
                            DateTime answer = today.AddMonths(1);
                            table_ipotek.Rows.Add(answer, Math.Round(plata, 2), Math.Round(procent_sum, 2), Math.Round(ezemeh, 2), Math.Round(sum2, 2));//Заполняем таблицу
                            today = answer;
                        }
                        plata_txt = Convert.ToString(Math.Round(itog_proc));
                        sum_itog += itog_proc;
                        sum_txt = Convert.ToString(Math.Round(sum_itog));
                        if (itog_proc < 0.4)
                        {
                            pereplat_ipoteka.ForeColor = Color.Black;
                            pereplat_ipoteka.Text = plata_txt;
                            vs9_suma_ipoteka.Text = sum_txt;
                        }
                        else
                        {
                            pereplat_ipoteka.ForeColor = Color.Red;
                            pereplat_ipoteka.Text = plata_txt;
                            vs9_suma_ipoteka.Text = sum_txt;
                        }
                    }
                    else if (por9dok.Text == "Аннуитетный")//вычисляем аннуитетную ипотеку
                    {
                        table_ipotek.Rows.Clear();
                        procent_sum = Math.Round(procent / 12 / 100, 7);
                        anyit = Math.Round((procent_sum * (Math.Pow((1 + procent_sum), srok)) / (Math.Pow((1 + procent_sum), srok) - 1)), 7);
                        anyit_mes = anyit * (sum - vznos);
                        sum1 = anyit_mes * srok - sum2;
                        plata = anyit_mes * srok;
                        while (srok > chet)
                        {
                            ezemeh = Math.Round(sum2 * procent_sum, 2);
                            sum_itog = Math.Round(anyit_mes - ezemeh, 2);
                            if (sum2 < anyit_mes)
                            {
                                sum_itog = Math.Round(sum2, 2);
                                sum2 = sum2 - Math.Round(sum_itog, 2);
                            }
                            else
                            {
                                sum2 = sum2 - Math.Round(sum_itog, 2);
                            }
                            DateTime answer = today.AddMonths(1);//Это для изменения месяца
                            table_ipotek.Rows.Add(answer, Math.Round(anyit_mes, 2), Math.Round(ezemeh, 2), Math.Round(sum_itog, 2), Math.Round(sum2, 2));//Заполняем таблицу
                            today = answer;

                            ++chet;
                        }
                        plata += vznos;
                        plata_txt = Convert.ToString(Math.Round(sum1 ));
                        sum_txt = Convert.ToString(Math.Round(plata));
                        if (sum1 < 0.4)                                     //Определяем цвет текста
                        {
                            pereplat_ipoteka.ForeColor = Color.Black;
                            pereplat_ipoteka.Text = plata_txt;
                            vs9_suma_ipoteka.Text = sum_txt;
                        }
                        else
                        {
                            pereplat_ipoteka.ForeColor = Color.Red;
                            pereplat_ipoteka.Text = plata_txt;
                            vs9_suma_ipoteka.Text = sum_txt;
                        }
                    }
                    pdf_ipotek.Enabled = true;
                }

            }
        }
        
        private void button2_Click_1(object sender, EventArgs e)
        {
            if (vibor_banka.Text == "")
            {
                MessageBox.Show("Выберите банк или калькулятор");
            }
            else
            {
                if ( sum_vklad.Text == "" || srok_vklad.Text == "" || proc_vklad.Text == "" || vidi_vklad.Text == "")
                {
                    MessageBox.Show("Заполните все поля");
                }
                else
                {
                    double period = 0;
                    DateTime today = DateTime.Now;
                    table_vklad.Rows.Clear();
                    double srok = double.Parse(srok_vklad.Text);
                    double srok1 = double.Parse(srok_vklad.Text);
                    if (period_vklad.Text == "Ежемесячная")
                    {
                        period = 12;

                    }
                    else if (period_vklad.Text == "Ежеквартальная")
                    {
                        period = 4;
                        srok = srok / 3;

                    }
                    else if (period_vklad.Text == "Ежегодная")
                    {
                        period = 1;
                        srok = srok / 12;
                    }

                    double sum = double.Parse(sum_vklad.Text);
                    double proc = double.Parse(proc_vklad.Text);
                    double sum1 = sum;
                    double Procent = 0;
                    double procent1 = 0;
                    double procent2 = 0;
                    string proc_text;
                    string suma_text;
                    int n = 0;
                    if (vidi_vklad.Text == "Выплата процентов")
                    {
                        Procent = (sum * proc * (srok * 31) / 365) / 100;
                        proc_text = Convert.ToString(Math.Round(Procent, 2));
                        pribil_vklad.Text = proc_text;
                        pribil_vklad.ForeColor = Color.Blue;
                        while (n < srok1)
                        {
                            procent2 = sum * ((proc / 100) / 12);
                            procent1 += procent2;
                            sum += procent2;
                            n++;
                            DateTime answer = today.AddMonths(1);
                            table_vklad.Rows.Add(answer, Math.Round(procent2, 2), Math.Round(sum, 2));
                            today = answer;
                        }
                        sum1 += Procent;
                       suma_text = Convert.ToString(Math.Round(sum1, 2));
                        suma.Text = suma_text;
                    }
                    else if (vidi_vklad.Text == "Kапитализацией процентов")
                    {
                        if (period_vklad.Text == "" )
                        {
                            MessageBox.Show("Заполните все поля");
                        }
                        else
                        {
                            Procent = (sum * (Math.Pow((1 + ((proc / 100) / period)), srok))) - sum;

                            proc_text = Convert.ToString(Math.Round(Procent, 2));
                            pribil_vklad.Text = proc_text;
                            pribil_vklad.ForeColor = Color.Blue;
                            while (n < srok)
                            {
                                procent2 = sum * ((proc / 100) / period);
                                procent1 += procent2;
                                sum += procent2;
                                n++;
                                DateTime answer = today.AddMonths(1);
                                table_vklad.Rows.Add(answer, Math.Round(procent2, 2), Math.Round(sum, 2));
                                today = answer;
                            }
                            sum1 += Procent;
                            suma_text = Convert.ToString(Math.Round(sum1, 2));
                            suma.Text = suma_text;
                        }
                    }
                    pdf_vklad.Enabled = true;

                }
            }
        }
        private void Chek()
        {
            ctrahovka_kredit.Visible = true;
            karta_ktedit.Visible = true;
            pension_kredit.Visible = true;
            vozrast_kredit.Visible = true;
            sibir_ipotek.Visible = true;
            maser_ipotek.Visible = true;
            karta_ipotek.Visible = true;
            ctrahovka_ipotek.Visible = true;
            ctrahovka_kredit.Checked = false;
            karta_ktedit.Checked = false;
            pension_kredit.Checked = false;
            vozrast_kredit.Checked = false;
            sibir_ipotek.Checked = false;
            maser_ipotek.Checked = false;
            karta_ipotek.Checked = false;
            ctrahovka_ipotek.Checked = false;
            procent_kredit.ReadOnly = true;
            procent_ipotek.ReadOnly = true;
            proc_vklad.ReadOnly = true;
        }
        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)//Проверка какой банк выбран и выставление условий банка
        {
            if (vibor_banka.Text == "Банк ТПК")
            {
                Chek();
                karta_ktedit.Text = "Я получаю зарплату на карту банка ТПК";
                karta_ipotek.Text = "Я получаю зарплату на карту банка ТПК";
                kredit.BackgroundImage = Properties.Resources.TPK;
                ipotek.BackgroundImage = Properties.Resources.TPK;
                vklad.BackgroundImage = Properties.Resources.TPK;
                procent_kredit.Text = "17";
                procent_ipotek.Text = "15";
                proc_vklad.Text = "50";
            }
            else if (vibor_banka.Text == "БТВ")
            {
                Chek();
                kredit.BackgroundImage = Properties.Resources.VTB;
                ipotek.BackgroundImage = Properties.Resources.VTB;
                vklad.BackgroundImage = Properties.Resources.VTB;
                procent_kredit.Text = "24,2";
                procent_ipotek.Text = "20";
                proc_vklad.Text = "10";
                karta_ktedit.Text = "Я получаю зарплату на карту БТВ";
                karta_ipotek.Text = "Я получаю зарплату на карту БТВ";
            }
            else if (vibor_banka.Text == "Ёлкабанк")
            {
                Chek();
                kredit.BackgroundImage = Properties.Resources.SBER;
                ipotek.BackgroundImage = Properties.Resources.SBER;
                vklad.BackgroundImage = Properties.Resources.SBER;
                procent_kredit.Text = "30";
                procent_ipotek.Text = "27";
                proc_vklad.Text = "5";
                karta_ktedit.Text = "Я получаю зарплату на карту Ёлкабанк";
                karta_ipotek.Text = "Я получаю зарплату на карту Ёлкабанк";
            }
            else if (vibor_banka.Text == "Бета банк")
            {
                Chek();
                kredit.BackgroundImage = Properties.Resources.ALFA;
                ipotek.BackgroundImage = Properties.Resources.ALFA;
                vklad.BackgroundImage = Properties.Resources.ALFA;
                procent_kredit.Text = "25,9";
                procent_ipotek.Text = "17";
                proc_vklad.Text = "15";
                karta_ktedit.Text = "Я получаю зарплату на карту Бета банка";
                karta_ipotek.Text = "Я получаю зарплату на карту Бета банка";

            }
            else if (vibor_banka.Text == "Калькулятор")
            {
                procent_kredit.Text = "";
                procent_ipotek.Text = "";
                ctrahovka_kredit.Visible = false;
                karta_ktedit.Visible = false;
                pension_kredit.Visible = false;
                vozrast_kredit.Visible = false;
                sibir_ipotek.Visible = false;
                maser_ipotek.Visible = false;
                karta_ipotek.Visible = false;
                ctrahovka_ipotek.Visible = false;
                kredit.BackgroundImage = Properties.Resources.white;
                ipotek.BackgroundImage = Properties.Resources.white;
                vklad.BackgroundImage = Properties.Resources.white;
                procent_kredit.ReadOnly = false;
                procent_kredit.Text = "";
                proc_vklad.ReadOnly = false;
                procent_ipotek.Text = "";
                procent_ipotek.ReadOnly = false;
                proc_vklad.Text = "";
                button1.BackColor = Color.White;
                button1.ForeColor = Color.Black;
            }
        }

        private void sum_credit_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar)))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    MessageBox.Show("Вы вводите не то", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        private void srok_credit_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar)))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    MessageBox.Show("Вы вводите не то", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        private void procent_kredit_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar)))
            {
                
                if (Char.IsDigit(e.KeyChar) || e.KeyChar == ',' || e.KeyChar == (char)Keys.Back)
                {
                    e.Handled = false;
                }
                else
                {
                    e.Handled = true;
                    MessageBox.Show("Вы вводите не то", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)//Проверка как условия выбраны и выставление соответсвующих условий
        {
            string procent_txt;
            double procent = double.Parse(procent_kredit.Text);
            if (procent >= 10)
            {
                if (ctrahovka_kredit.Checked == true)
                {
                    procent_txt = Convert.ToString(procent - 10);
                    procent_kredit.Text = procent_txt;
                }
                else
                {
                    procent_txt = Convert.ToString(procent + 10);
                    procent_kredit.Text = procent_txt;
                }
            }
            else
            {
                procent_kredit.Text = "0";
                if (ctrahovka_kredit.Checked == false)
                {
                    procent_txt = Convert.ToString(procent + 10);
                    procent_kredit.Text = procent_txt;
                }
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)//Проверка как условия выбраны и выставление соответсвующих условий
        {
            string procent_txt;
            double procent = double.Parse(procent_kredit.Text);
            
                if (karta_ktedit.Checked == true)
                {
                    procent_txt = Convert.ToString(procent - 1.3);
                    procent_kredit.Text = procent_txt;
                    pension_kredit.Checked = false;
                }
                else
                {
                    procent_txt = Convert.ToString(procent + 1.3);
                    procent_kredit.Text = procent_txt;
                }
            

        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)//Проверка как условия выбраны и выставление соответсвующих условий
        {
            string procent_txt;
            double procent = double.Parse(procent_kredit.Text);
            if (procent >= 1)
            {
                if (pension_kredit.Checked == true)
                {
                    procent_txt = Convert.ToString(procent - 1);
                    procent_kredit.Text = procent_txt;
                    karta_ktedit.Checked = false;
                    vozrast_kredit.Checked = false;
                }
                else
                {
                    procent_txt = Convert.ToString(procent + 1);
                    procent_kredit.Text = procent_txt;
                }
            }
            else
            {
                procent_kredit.Text = "0";
                if (pension_kredit.Checked == true)
                {
                    procent_txt = Convert.ToString(procent - 1);
                    procent_kredit.Text = procent_txt;
                    karta_ktedit.Checked = false;
                    vozrast_kredit.Checked = false;
                }
                else
                {
                    procent_txt = Convert.ToString(procent + 1);
                    procent_kredit.Text = procent_txt;
                }
            }

        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)//Проверка как условия выбраны и выставление соответсвующих условий
        {
            string procent_txt;
            double procent = double.Parse(procent_kredit.Text);
            if (procent <= 98)
            {
                if (vozrast_kredit.Checked == true)
                {
                    procent_txt = Convert.ToString(procent + 2);
                    procent_kredit.Text = procent_txt;
                    pension_kredit.Checked = false;
                }
                else
                {
                    procent_txt = Convert.ToString(procent - 2);
                    procent_kredit.Text = procent_txt;
                }
            }
            else
            {
                procent_kredit.Text = "100";
                if (vozrast_kredit.Checked == true)
                {
                    procent_txt = Convert.ToString(procent + 2);
                    procent_kredit.Text = procent_txt;
                    pension_kredit.Checked = false;
                }
                else
                {
                    procent_txt = Convert.ToString(procent - 2);
                    procent_kredit.Text = procent_txt;
                }
            }
        }

        private void checkBox10_CheckedChanged(object sender, EventArgs e)//Проверка как условия выбраны и выставление соответсвующих условий
        {
            string procent_txt;
            double procent = double.Parse(procent_ipotek.Text);
            
                if (ctrahovka_ipotek.Checked == true)
                {
                    procent_txt = Convert.ToString(procent - 10);
                    procent_ipotek.Text = procent_txt;
                }
                else
                {
                    procent_txt = Convert.ToString(procent + 10);
                    procent_ipotek.Text = procent_txt;
                }
            
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)//Проверка как условия выбраны и выставление соответсвующих условий
        {
            string procent_txt;
            double procent = double.Parse(procent_ipotek.Text);

            if (karta_ipotek.Checked == true)
            {
                procent_txt = Convert.ToString(procent - 1.3);
                procent_ipotek.Text = procent_txt;
            }
            else
            {
                procent_txt = Convert.ToString(procent + 1.3);
                procent_ipotek.Text = procent_txt;
            }
        }



        private void checkBox8_CheckedChanged(object sender, EventArgs e)//Проверка как условия выбраны и выставление соответсвующих условий
        {
            
            string procent_txt;
            double procent = double.Parse(procent_ipotek.Text);

            if (maser_ipotek.Checked == true)
            {
                procent_txt = Convert.ToString(procent - 1);
                procent_ipotek.Text = procent_txt;
            }
            else
            {
                procent_txt = Convert.ToString(procent + 1);
                procent_ipotek.Text = procent_txt;
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)//Проверка как условия выбраны и выставление соответсвующих условий
        {
            string procent_txt;
            double procent = double.Parse(procent_ipotek.Text);

            if (sibir_ipotek.Checked == true)
            {
                procent_txt = Convert.ToString(procent - 1.5);
                procent_ipotek.Text = procent_txt;
            }
            else
            {
                procent_txt = Convert.ToString(procent + 1.5);
                procent_ipotek.Text = procent_txt;
            }
        }

        

        private void sum_ipotek_KeyPress(object sender, KeyPressEventArgs e)//Проверка на правильность заполненых полей
        {
            if (!(char.IsDigit(e.KeyChar)))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    MessageBox.Show("Вы вводите не то", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        private void srok_ipotek_KeyPress(object sender, KeyPressEventArgs e)//Проверка на правильность заполненых полей
        {
            if (!(char.IsDigit(e.KeyChar)))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    MessageBox.Show("Вы вводите не то", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        private void vznos_ipotek_KeyPress(object sender, KeyPressEventArgs e)//Проверка на правильность заполненых полей
        {
            if (!(char.IsDigit(e.KeyChar)))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    MessageBox.Show("Вы вводите не то", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        private void procent_ipotek_KeyPress(object sender, KeyPressEventArgs e)//Проверка на правильность заполненых полей
        {
            if (Char.IsDigit(e.KeyChar) || e.KeyChar == ',' || e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
                MessageBox.Show("Вы вводите не то", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void vidi_vklad_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(vidi_vklad.Text== "Выплата процентов")
            {
                period_vklad.SelectedIndex = -1;
               period_vklad.Enabled = false;
            }
            else if(vidi_vklad.Text == "Kапитализацией процентов")
            {
                period_vklad.Enabled = true;
            }
        }

        private void proc_vklad_KeyPress(object sender, KeyPressEventArgs e)//Проверка на правильность заполненых полей
        {
            if (Char.IsDigit(e.KeyChar) || e.KeyChar == ',' || e.KeyChar == (char)Keys.Back)
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
                MessageBox.Show("Вы вводите не то", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }

        private void sum_vklad_KeyPress(object sender, KeyPressEventArgs e)//Проверка на правильность заполненых полей
        {
            if (!(char.IsDigit(e.KeyChar)))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    MessageBox.Show("Вы вводите не то", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        private void srok_vklad_KeyPress(object sender, KeyPressEventArgs e)//Проверка на правильность заполненых полей
        {
            if (!(char.IsDigit(e.KeyChar)))
            {
                if (e.KeyChar != (char)Keys.Back)
                {
                    e.Handled = true;
                    MessageBox.Show("Вы вводите не то", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        //private void save_Click(object sender, EventArgs e)//Сохранение файла в PDF формат
        //{
        //    var app = new Excel.Application();//Открывает Excel
        //    SaveFileDialog sfd = new SaveFileDialog();
        //    sfd.Filter = "PDF |*.pdf";
        //    if (sfd.ShowDialog() == DialogResult.OK)
        //    {
        //        var book = app.Workbooks.Add();
        //        Excel.Worksheet sheet = book.ActiveSheet;
        //        int row = 0,col =0;
              
        //        for (col = 0; col < tabel2.Columns.Count; col++)
        //        {
        //            for (row = 0; row < tabel2.Rows.Count; row++)
        //            {
        //                sheet.Cells[row + 2, col + 2] = tabel2[col, row].Value.ToString();
        //            }
        //        }
        //        sheet.Cells[1,2] = "Дата платежа   ";
        //        sheet.Cells[1, 3] = "Платеж   ";
        //        sheet.Cells[1, 4] = "Процент   ";
        //        sheet.Cells[1, 5] = "Тело кредита   ";
        //        sheet.Cells[1, 6] = "Остаток   ";
        //        book.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, sfd.FileName);
        //        app.DisplayAlerts = false;
        //    }
        //    app.Quit();
        //}

        //private void pdf_vklad_Click(object sender, EventArgs e)//Сохранение файла в PDF формат
        //{
        //    var app = new Excel.Application();//Открывает Excel
        //    SaveFileDialog sfd = new SaveFileDialog();
        //    sfd.Filter = "PDF |*.pdf";
        //    if (sfd.ShowDialog() == DialogResult.OK)
        //    {
        //        var book = app.Workbooks.Add();
        //        Excel.Worksheet sheet = book.ActiveSheet;
        //        for (int col = 0; col < table_vklad.Columns.Count; col++)
        //        {
        //            for (int row = 0; row < table_vklad.Rows.Count; row++)
        //            {
        //                sheet.Cells[row + 2, col + 2] = table_vklad[col, row].Value.ToString();
        //            }
        //        }
        //        sheet.Cells[1, 2] = "Дата платежа   ";
        //        sheet.Cells[1, 3] = "Процент   ";
        //        sheet.Cells[1, 4] = "Остаток суммы   ";
        //        book.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, sfd.FileName);
        //        app.DisplayAlerts = false;
        //    }
        //    app.Quit();
        //}

        //private void pdf_ipotek_Click(object sender, EventArgs e)//Сохранение файла в PDF формат
        //{
        //    var app = new Excel.Application();//Открывает Excel
        //    SaveFileDialog sfd = new SaveFileDialog();
        //    sfd.Filter = "PDF |*.pdf";
        //    if (sfd.ShowDialog() == DialogResult.OK)
        //    {
        //        var book = app.Workbooks.Add();
        //        Excel.Worksheet sheet = book.ActiveSheet;
        //        for (int col = 0; col < table_ipotek.Columns.Count; col++)
        //        {
        //            for (int row = 0; row < table_ipotek.Rows.Count; row++)
        //            {
        //                sheet.Cells[row + 2, col + 2] = table_ipotek[col, row].Value.ToString();
        //            }
        //        }
        //        sheet.Cells[1, 2] = "Дата платежа   ";
        //        sheet.Cells[1, 3] = "Платеж   ";
        //        sheet.Cells[1, 4] = "Процент   ";
        //        sheet.Cells[1, 5] = "Тело кредита   ";
        //        sheet.Cells[1, 6] = "Остаток   ";
        //        book.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, sfd.FileName);
        //        app.DisplayAlerts = false;
        //    }
        //    app.Quit();
        //}

        private void procent_kredit_TextChanged(object sender, EventArgs e)
        {
            if (procent_kredit.Text !="")//Проверкка textbox
            {
                double pr = Convert.ToDouble(procent_kredit.Text);
                if (pr > 365)
                {
                    MessageBox.Show("Проценты больше 365");
                    procent_kredit.Text = "";
                }
            }
        }

        private void vznos_ipotek_TextChanged(object sender, EventArgs e)
        {
            
           
                if (vznos_ipotek.Text != "" )//Проверкка textbox
                {
                    double vznos = Convert.ToDouble(vznos_ipotek.Text);
                    double sum = Convert.ToDouble(sum_ipotek.Text);
                    double chek = sum / 100 * 90;
                if (vznos > chek)
                    {

                        MessageBox.Show("Первоначальный взнос не должен дыть больше суммы на 90%");
                        vznos_ipotek.Text = Convert.ToString(chek);
                    }
                }
            
        }

        private void sum_ipotek_TextChanged(object sender, EventArgs e)
        {
            if (sum_ipotek.Text != "")//Проверкка textbox
            {
                double vznos = 0;
                if (vznos_ipotek.Text != "")//Проверкка textbox
                {
                   vznos = Convert.ToDouble(vznos_ipotek.Text);
                }
                double sum = Convert.ToDouble(sum_ipotek.Text);
                double chek = sum / 100 * 90;
                if (vznos > chek)
                {

                    MessageBox.Show("Первоначальный взнос не должен дыть больше суммы на 90%");
                    vznos_ipotek.Text = Convert.ToString(chek);
                }
            }
            if (sum_ipotek.Text != "")
            {
                vznos_ipotek.Enabled = true;
            }
            else
            {
                vznos_ipotek.Text = "";
                vznos_ipotek.Enabled = false;
            }
        }

        private void srok_credit_TextChanged(object sender, EventArgs e)
        {
            if (srok_credit.Text != "")//Проверкка textbox
            {
                double pr = Convert.ToDouble(srok_credit.Text);
                if (pr > 120||pr<1)
                {
                    MessageBox.Show("Срок от 1 до 120 месяцев");
                    srok_credit.Text = "1";
                }
            }
        }

        private void procent_ipotek_TextChanged(object sender, EventArgs e)
        {
            if (procent_ipotek.Text != "")//Проверкка textbox
            {
                double pr = Convert.ToDouble(procent_ipotek.Text);
                if (pr > 365)
                {
                    MessageBox.Show("Проценты больше 365");
                    procent_ipotek.Text = "";
                }
            }
        }

        private void srok_ipotek_TextChanged(object sender, EventArgs e)
        {
            if (srok_ipotek.Text != "")//Проверкка textbox
            {
                double pr = Convert.ToDouble(srok_ipotek.Text);
                if (pr > 120 || pr < 1)
                {
                    MessageBox.Show("Срок от 1 до 120 месяцев");
                    srok_ipotek.Text = "1";
                }
            }
        }

        private void srok_vklad_TextChanged(object sender, EventArgs e)
        {
            if (srok_vklad.Text != "")//Проверкка textbox
            {
                double pr = Convert.ToDouble(srok_vklad.Text);
                if (pr > 120 || pr < 1)
                {
                    MessageBox.Show("Срок от 1 до 120 месяцев");
                    srok_vklad.Text = "1";
                }
            }
        }

        private void proc_vklad_TextChanged(object sender, EventArgs e)
        {
            if (proc_vklad.Text != "")//Проверкка textbox
            {
                double pr = Convert.ToDouble(proc_vklad.Text);
                if (pr > 365)
                {
                    MessageBox.Show("Проценты больше 365");
                    proc_vklad.Text = "";
                }
            }
        }
    }
}
