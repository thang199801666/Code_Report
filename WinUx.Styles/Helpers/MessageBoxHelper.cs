namespace WinUx.Controls
{
    public static class MessageBoxHelper
    {
        public static string? GetIconUri(MessageBoxIcon icon)
        {
            return icon switch
            {
                MessageBoxIcon.Information => "pack://application:,,,/WinUx.Styles;component/Resources/Information.png",
                MessageBoxIcon.Warning => "pack://application:,,,/WinUx.Styles;component/Resources/Warning.png",
                MessageBoxIcon.Error => "pack://application:,,,/WinUx.Styles;component/Resources/Error.png",
                MessageBoxIcon.Question => "pack://application:,,,/WinUx.Styles;component/Resources/Question.png",
                MessageBoxIcon.Success => "pack://application:,,,/WinUx.Styles;component/Resources/Success.png",
                _ => "pack://application:,,,/WinUx.Styles;component/Resources/None.png"
            };
        }
    }
}
