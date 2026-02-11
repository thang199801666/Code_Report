namespace WinUx.Controls
{
    public enum MessageBoxButtons
    {
        OK,
        OKCancel,
        YesNo,
        YesNoCancel,
        YesNoCancelYesToAll
    }

    public readonly struct ButtonConfig
    {
        public bool ShowOk { get; init; }
        public bool ShowCancel { get; init; }
        public bool ShowYes { get; init; }
        public bool ShowNo { get; init; }
        public bool ShowYesToAll { get; init; }

        public ButtonConfig(bool showOk = false, bool showCancel = false, bool showYes = false, bool showNo = false, bool showYesToAll = false)
        {
            ShowOk = showOk;
            ShowCancel = showCancel;
            ShowYes = showYes;
            ShowNo = showNo;
            ShowYesToAll = showYesToAll;
        }
    }

    public static class MessageBoxButtonsExtensions
    {
        public static ButtonConfig GetConfig(this MessageBoxButtons buttons) =>
            buttons switch
            {
                MessageBoxButtons.OK => new ButtonConfig(showOk: true),
                MessageBoxButtons.OKCancel => new ButtonConfig(showOk: true, showCancel: true),
                MessageBoxButtons.YesNo => new ButtonConfig(showYes: true, showNo: true),
                MessageBoxButtons.YesNoCancel => new ButtonConfig(showYes: true, showNo: true, showCancel: true),
                MessageBoxButtons.YesNoCancelYesToAll => new ButtonConfig(showYes: true, showNo: true, showYesToAll: true, showCancel: true),
                _ => new ButtonConfig(showOk: true)
            };
    }
}