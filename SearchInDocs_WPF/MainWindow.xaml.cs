using SearchInDocs_WPF.Cmds;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SearchInDocs_WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ICommand dragWindowCommand = null;
        private ICommand minimizeWindowCommand = null;
        private ICommand closeWindowCommand = null;

        public MainWindow()
        {
            InitializeComponent();
        }

        public ICommand DragWindowCommand => 
            dragWindowCommand ?? (dragWindowCommand = new DragWindowCommand());

        public ICommand MinimizeWindowCommand =>
            minimizeWindowCommand ?? (minimizeWindowCommand = new MinimizeWindowCommand());

        public ICommand CloseWindowCommand =>
            closeWindowCommand ?? (closeWindowCommand = new CloseWindowCommand());
    }
}
