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

namespace DZ_22_02_Datagrid_view_to_csv
{
    public partial class ConvertCSV : Form
    {
        public ConvertCSV()
        {
            InitializeComponent();
        }

        private void ConvertCSV_Load(object sender, EventArgs e)
        {
            Random random = new Random();

            // Список для того, чтобы значения не повторялись
            List<char> lstSave = new List<char>();

            char[] array = { 'а', 'б', 'в', 'г', 'д', 'е', 'ё', 'ж', 'з', 'и', 'й', 'к', 'л', 'м', 'н', 'о', 'п', 'р', 'с', 'т', 'у', 'ф', 'х', 'ц', 'ч', 'ш', 'щ', 'ъ', 'ы', 'ь', 'э', 'ю', 'я'};
            dataGridView1.RowCount = 5;
            dataGridView1.ColumnCount = 5;

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    int tmp = random.Next(array.Length);
                    if (Chech(lstSave, array[tmp]))
                    {
                        dataGridView1.Rows[i].Cells[j].Value = array[tmp];
                    }
                    else
                    {
                        j--;
                    }
                }
            }

            }
            private bool Chech(List<char> lst, char ch)
            {
                for (int i = 0; i < lst.Count; i++)
                {
                    if (lst[i] == ch)
                    {
                        return false;
                    }
                }
                // Данного значения не было в DataGridView
                lst.Add(ch);
                return true;
            }
        

        private void export_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "Excel files (*.csv)|*.csv|Text files (*.txt)|*.txt";

            if(dlg.ShowDialog()==DialogResult.OK)
            {
                Stream myStream;
                myStream = dlg.OpenFile();
                StreamWriter sw = new StreamWriter(myStream, System.Text.Encoding.GetEncoding(-0));
                string columnTitle = "";
                try
                {
                    for (int i=0;i< dataGridView1.ColumnCount;i++)
                    {
                        if(i>0)
                        {
                            columnTitle += " ";
                        }
                        columnTitle += dataGridView1.Columns[i].HeaderText;
                    }
                    sw.WriteLine(columnTitle);

                    // Написать содержимое столбца

                    for(int j = 0; j<dataGridView1.Rows.Count;j++)
                    {
                        string columnValue = "";
                        for(int k=0;k<dataGridView1.Columns.Count; k++)
                        {
                            if(k>0)
                            {
                                columnValue += " ";
                            }
                            if (dataGridView1.Rows[j].Cells[k].Value == null)
                                columnValue += "";
                            else if (dataGridView1.Rows[j].Cells[k].Value.ToString().Contains(" "))
                            {
                                columnValue += "\"" + dataGridView1.Rows[j].Cells[k].Value.ToString().Trim() + "\"";
                            }
                            else
                            {
                                columnValue += dataGridView1.Rows[j].Cells[k].Value.ToString().Trim() + "\t";
                            }
                        }
                        sw.WriteLine(columnValue);

                    }
                    sw.Close();
                    myStream.Close();
                    MessageBox.Show("Данные экспортированы");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Форма экспорта не удалась");
                }
                finally
                {
                    sw.Close();
                    myStream.Close();
                }
            }
            else
            {
                MessageBox.Show("Отмените операцию экспорта");
            }


        }

        private void exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
