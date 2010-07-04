using System;
using System.ComponentModel;
using System.Deployment.Application;
using System.Reflection;
using Office.Utility.Extensions;

namespace Office.Utility
{
    /// <summary>
    /// Generic clickonce deployment helper class.
    /// </summary>
    public class ClickOnceUpdater
    {
        private readonly BackgroundWorker _worker = new BackgroundWorker();
        private Action<UpdateResult> _callback;
        private readonly string _applicationName;

        /// <summary>
        /// Initializes a new instance of the <see cref="ClickOnceUpdater"/> class.
        /// </summary>
        public ClickOnceUpdater() : this(Assembly.GetCallingAssembly().GetName().Name)
        { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ClickOnceUpdater"/> class.
        /// </summary>
        /// <param name="applicationName">Name of the application.</param>
        public ClickOnceUpdater(string applicationName)
        {
            _applicationName = applicationName;
            _worker.DoWork += (sender, e) => e.Result = UpdateApplication();
            _worker.RunWorkerCompleted += (sender, e) =>
            {
                if (_callback == null) return;

                _callback(e.Result as UpdateResult);
                _callback = null;
            };
        }

        /// <summary>
        /// Checks for update asynchronously.
        /// </summary>
        /// <param name="complete">The completed callback.</param>
        public void CheckForUpdateAsync(Action<UpdateResult> complete)
        {
            if (_worker.IsBusy) return;

            _callback = complete;
            _worker.RunWorkerAsync();
        }

        /// <summary>
        /// Updates the ClickOnce application.
        /// </summary>
        /// <returns></returns>
        public UpdateResult UpdateApplication()
        {
            try
            {
                if (!ApplicationDeployment.IsNetworkDeployed)
                    return NotNetworkDeployedResult();

                var currentDeployment = ApplicationDeployment.CurrentDeployment;

                return UpdateApplication(currentDeployment);
            }
            catch (Exception ex)
            {
                return FailResult(string.Concat(ex.ToMessageStack(), 
                    Environment.NewLine, "Stack Trace: ", Environment.NewLine,
                    ex.ToFullStackTrace()));
            }
        }

        /// <summary>
        /// Updates the ClickOnce application.
        /// </summary>
        /// <param name="currentDeployment">The current deployment.</param>
        /// <returns></returns>
        protected virtual UpdateResult UpdateApplication(ApplicationDeployment currentDeployment)
        {
            var info = currentDeployment.CheckForDetailedUpdate();

            if (!info.UpdateAvailable)
                return NoUpdateNeededResult();

            var message = string.Empty;
            return UpdateCurrentDeployment(currentDeployment, ref message) ? 
                SuccessResult(info) : FailResult(message);
        }

        /// <summary>
        /// Updates the current deployment.
        /// </summary>
        /// <param name="deployment">The deployment.</param>
        /// <param name="message">The message.</param>
        /// <returns></returns>
        protected virtual bool UpdateCurrentDeployment(ApplicationDeployment deployment, ref string message)
        {
            return deployment.Update();
        }

        private UpdateResult SuccessResult(UpdateCheckInfo info)
        {
            return new UpdateResult
                       {
                           Success = true,
                           Updated = true,
                           Message = string.Format(Properties.Resources.Deployment_Success,
                                                   _applicationName,
                                                   info.AvailableVersion)
                       };
        }

        private static UpdateResult FailResult(string message)
        {
            return new UpdateResult
                           {
                               Success = false,
                               Updated = false,
                               Message = string.Format(Properties.Resources.Deployment_UpdateFailed, message)
                           };
        }

        private static UpdateResult NoUpdateNeededResult()
        {
            return new UpdateResult
                       {
                           Success = true,
                           Updated = false,
                           Message = Properties.Resources.Deployment_NoUpdateAvailable
                       };
        }

        private static UpdateResult NotNetworkDeployedResult()
        {
            return new UpdateResult
                       {
                           Success = true,
                           Updated = false,
                           Message = Properties.Resources.Deployment_NotNetworkDeployed
                       };
        }
    }
}
