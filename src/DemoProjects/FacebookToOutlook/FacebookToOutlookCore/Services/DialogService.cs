using System;
using System.Linq;
using System.Windows;
using FacebookToOutlookCore.Services.Interfaces;

namespace FacebookToOutlookCore.Services
{
    public class DialogService : IDialogService
    {
        public bool? ShowDialog<T>(object ownerViewModel, object viewModel) where T : Window
        {
            return ShowDialog<T>(ownerViewModel, viewModel, null);
        }

        public bool? ShowDialog<T>(object ownerViewModel, object viewModel, Action<Window> setupWindow) where T : Window
        {
            // Create dialog and set properties
            var dialog = Activator.CreateInstance<T>();

            if (setupWindow != null)
                setupWindow(dialog);

            if (ownerViewModel != null)
                dialog.Owner = FindOwnerWindow(ownerViewModel);
            dialog.DataContext = viewModel;


            // Show dialog
            return dialog.ShowDialog();
        }

        public void Show<T>(object ownerViewModel, object viewModel) where T : Window
        {
            // Create dialog and set properties
            var dialog = Activator.CreateInstance<T>();
            if (ownerViewModel != null)
                dialog.Owner = FindOwnerWindow(ownerViewModel);
            dialog.DataContext = viewModel;

            // Show dialog
            dialog.Show();
        }


        public MessageBoxResult ShowMessageBox(object ownerViewModel, string messageBoxText, string caption,
                                               MessageBoxButton button, MessageBoxImage icon)
        {
            return MessageBox.Show(messageBoxText, caption, button, icon);
        }

        private static Window FindOwnerWindow(object viewModel)
        {
            var view = Application.Current.Windows.Cast<Window>().SingleOrDefault(v => ReferenceEquals(v.DataContext, viewModel));

            if (view == null)
                throw new ArgumentException("Viewmodel is not referenced by any registered View.");

            return view;
        }

        public void CloseDialog(object viewModel, bool? result)
        {
            var view = FindOwnerWindow(viewModel);
            view.DialogResult = result;
            view.Close();
        }

        public void Close(object viewModel)
        {
            var view = FindOwnerWindow(viewModel);
            view.Close();
        }

        public bool? ShowDialog<T>(object ownerViewModel, object viewModel, bool callOnUiThread) where T : Window
        {
            if (!Application.Current.CheckAccess())
                return (bool?)Application.Current.Dispatcher.Invoke((Func<bool?>)(() => ShowDialog<T>(ownerViewModel, viewModel)));
            
            return ShowDialog<T>(ownerViewModel, viewModel);
        }
    }
}