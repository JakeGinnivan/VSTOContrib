using System;
using System.Windows;

namespace FacebookToOutlookCore.Services.Interfaces
{
    public interface IDialogService
    {
        bool? ShowDialog<T>(object ownerViewModel, object viewModel) where T : Window;
        bool? ShowDialog<T>(object ownerViewModel, object viewModel, Action<Window> setupWindow) where T : Window;
        bool? ShowDialog<T>(object ownerViewModel, object viewModel, bool callOnUiThread) where T : Window;
        void Show<T>(object ownerViewModel, object viewModel) where T : Window;
        void CloseDialog(object viewModel, bool? result);
        void Close(object viewModel);
        MessageBoxResult ShowMessageBox(object ownerViewModel, string messageBoxText, string caption,
                                        MessageBoxButton button, MessageBoxImage icon);
    }
}