namespace Task_2
{
    public class Message
    {
        static public void MessageNotification(string message)
        {
            string caption = "Уведомление";
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result;

            result = MessageBox.Show(message, caption, buttons);
        }

        static public void MessageError(string message)
        {
            string caption = "Ошибка";
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result;

            result = MessageBox.Show(message, caption, buttons);
        }
    }
}
