using System.Windows.Forms;
using System.IO;

namespace Task_2
{
    public partial class MaynForm : Form
    {
        bool isStart = true;

        public MaynForm()
        {
            InitializeComponent();
        }

        private void SearchLoad_Click(object sender, EventArgs e)
        {
            if (SearchFile.ShowDialog() == DialogResult.Cancel)
                return;

            LoadFile.Text = SearchFile.SelectedPath;
        }

        private void SearchSave_Click(object sender, EventArgs e)
        {
            if (SearchFile.ShowDialog() == DialogResult.Cancel)
                return;

            SaveFile.Text = SearchFile.SelectedPath;
        }

        private void Start_Click(object sender, EventArgs e)
        {
            if(LoadFile.Text == "" || SaveFile.Text == "")
            {
                Message.MessageError("Заполните поля");
            }
            else
            {
                if (!Directory.Exists(LoadFile.Text))
                {
                    Message.MessageError("Не удалось найти папку для обработки");
                    isStart = false;
                }

                if(!Directory.Exists(SaveFile.Text))
                {
                    Message.MessageError("Не удалось найти папку для сохранения Описи");
                    isStart = false;
                }

                if(isStart)
                {
                    WorkFile.ChekFile(LoadFile.Text, SaveFile.Text);
                    Message.MessageNotification("Опись произведена");
                }
            }
        }
    }
}
