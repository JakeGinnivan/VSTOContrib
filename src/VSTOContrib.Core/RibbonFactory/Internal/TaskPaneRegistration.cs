namespace VSTOContrib.Core.RibbonFactory.Internal
{
    class TaskPaneRegistration
    {
        public TaskPaneRegistration(TaskPaneRegistrationInfo registrationInfo, OneToManyCustomTaskPaneAdapter adapter)
        {
            RegistrationInfo = registrationInfo;
            Adapter = adapter;
        }

        public TaskPaneRegistrationInfo RegistrationInfo { get; private set; }

        public OneToManyCustomTaskPaneAdapter Adapter { get; private set; }
    }
}