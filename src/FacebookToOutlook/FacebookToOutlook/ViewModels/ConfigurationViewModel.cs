using System;
using System.ComponentModel;
using System.Windows.Input;
using AutoMapper;
using FacebookToOutlook.Core;
using FacebookToOutlook.Presentation.Commands;
using FacebookToOutlook.Presentation.ViewModels;
using FacebookToOutlook.Services;

namespace FacebookToOutlook.ViewModels
{
    public class ConfigurationViewModel : ErrorInfoViewModelBase, IEditableObject, IConfigurationSettings
    {
        private readonly IConfigurationSettings _configurationSettings;
        private readonly IDialogService _dialogService;
        private readonly EventConfigurationViewModel _eventConfigurationSettings;
        private readonly ContactConfigurationViewModel _contactConfigurationSettings;

        private DelegateCommand _saveCommand;
        private ConfigurationTab _selectedConfigurationTab;

        public ConfigurationViewModel(IConfigurationSettings configurationSettings, IDialogService dialogService,
            EventConfigurationViewModel eventConfigurationSettings,
            ContactConfigurationViewModel contactConfigurationSettings)
        {
            _configurationSettings = configurationSettings;
            _dialogService = dialogService;
            _eventConfigurationSettings = eventConfigurationSettings;
            _contactConfigurationSettings = contactConfigurationSettings;
            BeginEdit();
        }

        public ConfigurationTab SelectedConfigurationTab
        {
            get
            {
                return _selectedConfigurationTab;
            }
            set
            {
                _selectedConfigurationTab = value;
                RaisePropertyChanged("SelectedConfigurationTab");
            }
        }

        public ICommand SaveCommand
        {
            get
            {
                return _saveCommand ?? (_saveCommand = new DelegateCommand(EndEdit, ()=>IsValid));
            }
        }

        public ICommand CancelCommand
        {
            get
            {
                return _saveCommand ?? (_saveCommand = new DelegateCommand(CancelEdit));
            }
        }

        public IEventConfigurationSettings EventConfigurationSettings
        {
            get
            {
                return _eventConfigurationSettings;
            }
        }

        public IContactConfigurationSettings ContactConfigurationSettings
        {
            get
            {
                return _contactConfigurationSettings;
            }
        }

        public void Save()
        {
            throw new NotSupportedException("Call EndEdit instead");
        }
        
        public void BeginEdit()
        {
            Mapper.Map(_configurationSettings.EventConfigurationSettings, _eventConfigurationSettings);
            Mapper.Map(_configurationSettings.ContactConfigurationSettings, _contactConfigurationSettings);
        }

        public void EndEdit()
        {
            Mapper.Map(_eventConfigurationSettings, _configurationSettings.EventConfigurationSettings);
            Mapper.Map(_contactConfigurationSettings, _configurationSettings.ContactConfigurationSettings);
            _configurationSettings.Save();
            _dialogService.CloseDialog(this, true);
        }

        public void CancelEdit()
        {
            _dialogService.CloseDialog(this, false);
        }
    }
}
