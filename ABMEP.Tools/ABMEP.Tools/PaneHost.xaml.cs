using System.Windows.Controls;

namespace ABMEP.Tools
{
    public partial class PaneHost : UserControl
    {
        // Singleton instance Revit will dock and we keep alive
        public static PaneHost Instance { get; } = new PaneHost();

        public PaneHost()
        {
            InitializeComponent(); // wires up PART_Content from the XAML
        }

        // Called later to inject the hotloaded ABMEP.Work UI
        public void SetContent(UserControl content)
        {
            PART_Content.Content = content;
        }
    }
}
