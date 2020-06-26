using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace SearchInDocs_WPF.Cmds
{
    class MinimizeWindowCommand : CommandBase
    {
        public override bool CanExecute(object parameter) =>
            (parameter as Window) != null;

        public override void Execute(object parameter) =>
            ((Window)parameter).WindowState = WindowState.Minimized;
    }
}
