using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Free_PST_Backup
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        int count = 0;

        String currectDirectory = Directory.GetCurrentDirectory();

        private void Form1_Load(object sender, EventArgs e)
        {
            progressBar1.Maximum = 1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            showMessage("This process may take a while \n Please, don't close the application \n This message will close automatically in 4 seconds", 4000);  /* 1 segundo = 1000 */
            progressBar1.Value = 0;

            if (radioButton1.Checked)
            {
                String userprofile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

                checkedListBox.Items.Clear();
                List<string> file_list = new List<string>();
                SearchDirectory(userprofile, file_list);

                label3.Text = "Total PST: " + checkedListBox.Items.Count.ToString();


            }
            else if (radioButton2.Checked)
            {
                checkedListBox.Items.Clear();
                List<string> file_list = new List<string>();
                SearchDirectory(@"c:\\", file_list);

                label2.Text = "Total PST: " + checkedListBox.Items.Count.ToString();
            }


            if (checkedListBox.Items.Count <= 0)
            {
                MessageBox.Show("No PST files found");
            }
        }


        private void SearchDirectory(string path, List<string> file_list)
        {
            DirectoryInfo dir_info = new DirectoryInfo(path);
            try
            {
                foreach (DirectoryInfo subdir_info in dir_info.GetDirectories())
                {
                    SearchDirectory(subdir_info.FullName, file_list);
                }
            }
            catch
            {
            }
            try
            {

                foreach (FileInfo file_info in dir_info.GetFiles())
                {
                    if (file_info.Extension == ".pst" || file_info.Extension == ".ost")
                        checkedListBox.Items.Add(file_info);
                }

            }
            catch
            {
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            //comprobar si existe la carpeta antes de crearla

            if (Directory.Exists(DateTime.Now.ToString("[yy-MM-dd]" + " PST Backup")))
            {
                MessageBox.Show("The backup directory alredy exists \n please, rename or delete it before continue");
            }
            else
            {

                progressBar1.Value = 0;


                //Cerrar Outlook antes de copiar
                killOutlook();

                // Determine if there are any items checked.  
                if (checkedListBox.CheckedItems.Count != 0)
                {


                    showMessage("This process may take a while \n Please, don't close the application \n This message will close automatically in 4 seconds", 4000);  /* 1 segundo = 1000 */

                    Directory.CreateDirectory(DateTime.Now.ToString("[yy-MM-dd]" + " PST Backup"));
                    // If so, loop through all checked items and print results.  

                    string dest_path = DateTime.Now.ToString("[yy-MM-dd]" + " PST Backup");

                    for (int i = 0; i < checkedListBox.CheckedItems.Count; i++)
                    {

                        string path_item = Path.GetFullPath(checkedListBox.CheckedItems[i].ToString());
                        string path_only = Path.GetDirectoryName(checkedListBox.CheckedItems[i].ToString());
                        String file_name = Path.GetFileName(checkedListBox.CheckedItems[i].ToString());
                        string path_folderDate = dest_path + "/" + file_name;

                        //Meter condicion si existe archivo
                        if (File.Exists(path_folderDate))
                        {
                            count++;
                            FileInfo file = new FileInfo(path_item);
                            file.CopyTo(dest_path + "/" + "(" + count + ")" + file_name);


                        }
                        else
                        {
                            FileInfo file = new FileInfo(path_item);
                            file.CopyTo(dest_path + "/" + file_name);
                        }

                    }
                    //cuando se termina la copia se llena la barra indicando el proceso completado
                    progressBar1.Value = 1;


                    //abre la ubicacion
                    Process.Start("explorer.exe", dest_path);
                }


                else
                {
                    MessageBox.Show("You must select a PST");
                }
            }
        }

        private void killOutlook()
        {

            //Comprobar si el proceso Outlook.exe esta corriendo y si no lo esta, lo cierra.

            Process[] proc = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");

            if (proc.Length > 0)
            {
                proc[0].Kill();
            }
            else
            {
                //Nothing to do
            }
        }

 
        private void timeTick(object sender, EventArgs e)
        {
            (sender as Timer).Stop();  /* Detiene el Timer */
            SendKeys.Send("{ESC}"); /* Hace la simulación de la tecla Escape, también puedes usar {ENTER} */
        }


        private void showMessage(string msg, int duration)
        {
            using (Timer t = new Timer())
            {
                Timer time = new Timer();
                time.Interval = duration;
                time.Tick += timeTick;  /* Evento enlazado */

                time.Start();

                /* Muestras el texto en el MB */
                MessageBox.Show(msg);
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/antonio-castillo");
        }
    }
}
