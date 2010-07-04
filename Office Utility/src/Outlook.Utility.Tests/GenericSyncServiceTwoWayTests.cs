using System;
using Outlook.Utility.Services;
using Outlook.Utility.Services.Conficts;
using Rhino.Mocks;
using Xunit;

namespace Outlook.Utility.Tests
{
    public class GenericSyncServiceTwoWayTests
    {
        private readonly ISynchronisationProvider<SyncClass, int> _localProvider;
        private readonly ISynchronisationProvider<SyncClass, int> _remoteProvider;
        private readonly GenericSynchronisationService<SyncClass, int> _syncService;
        private readonly ISyncSettings _settings;

        public GenericSyncServiceTwoWayTests()
        {
            _localProvider = MockRepository.GenerateStub<ISynchronisationProvider<SyncClass, int>>();
            _remoteProvider = MockRepository.GenerateStub<ISynchronisationProvider<SyncClass, int>>();
            _settings = MockRepository.GenerateStub<ISyncSettings>();
            _settings.LastSync = DateTime.Today;
            _syncService = new GenericSynchronisationService<SyncClass, int>(c => c.Id, _localProvider, _remoteProvider, _settings, SyncDirection.TwoWay);
        }

        [Fact]
        public void SingleRemoteEntrySavesLocally()
        {
            //Arrange
            var syncClass = new SyncClass{ Id = 1};
            SetupProviders(null, new[] { syncClass }, null, null);

            //Act
            _syncService.PerformSynchronisation();

            //Assert
            _localProvider.AssertWasCalled(p => p.SaveEntries(new[] { syncClass }));
        }

        [Fact]
        public void SingleLocalEntrySavesRemotely()
        {
            //Arrange
            var syncClass = new SyncClass { Id = 1 };
            SetupProviders(new[] { syncClass }, null, null, null);

            //Act
            _syncService.PerformSynchronisation();

            //Assert
            _remoteProvider.AssertWasCalled(p => p.SaveEntries(new[] { syncClass }));
        }

        [Fact]
        public void ConflictedEntryWithRemoteWinsResolver()
        {
            //Arrange
            var syncClass = new SyncClass { Id = 1, Data = "Original Data" };
            var modifiedSyncClass = new SyncClass { Id = 1, Data = "Local modified Data" };
            SetupProviders(new[] { modifiedSyncClass }, new[] { syncClass }, null, null);

            //Act
            _syncService.PerformSynchronisation();

            //Assert
            _localProvider.AssertWasCalled(p => p.SaveEntries(new[] { syncClass }));
        }

        [Fact]
        public void ConflictedEntryWithLocalWinsResolver()
        {
            //Arrange
            var syncClass = new SyncClass { Id = 1, Data = "Original Data" };
            var modifiedSyncClass = new SyncClass { Id = 1, Data = "Local modified Data" };
            SetupProviders(new[] { modifiedSyncClass }, new[] { syncClass }, null, null);
            _syncService.ConflcitResolver = new LocalWinsResolver<SyncClass>();

            //Act
            _syncService.PerformSynchronisation();

            //Assert
            _remoteProvider.AssertWasCalled(p => p.SaveEntries(new[] { modifiedSyncClass }));
        }

        private void SetupProviders(SyncClass[] localModified, SyncClass[] remoteModified, int[] localDeletedIds, int[] remoteDeletedIds)
        {
            _remoteProvider.Stub(p => p.GetModifiedEntries(Arg<DateTime>.Is.Anything)).Return(remoteModified?? new SyncClass[0]);
            _localProvider.Stub(p => p.GetModifiedEntries(Arg<DateTime>.Is.Anything)).Return(localModified?? new SyncClass[0]);
            _remoteProvider.Stub(p => p.GetDeletedEntries(Arg<DateTime>.Is.Anything)).Return(remoteDeletedIds ?? new int[0]);
            _localProvider.Stub(p => p.GetDeletedEntries(Arg<DateTime>.Is.Anything)).Return(localDeletedIds ?? new int[0]);
        }
    }

    public class SyncClass
    {
        public int Id { get; set; }
        public string Data { get; set; }
    }
}
