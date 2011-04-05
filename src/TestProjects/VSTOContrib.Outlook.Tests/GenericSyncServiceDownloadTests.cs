using System;
using System.Collections.Generic;
using System.Linq;
using NSubstitute;
using VSTOContrib.Outlook.Services;
using VSTOContrib.Outlook.Services.Conficts;
using Xunit;

namespace VSTOContrib.Outlook.Tests
{
    public class GenericSyncServiceDownloadTests
    {
        private readonly ISynchronisationProvider<SyncClass, int> _localProvider;
        private readonly ISynchronisationProvider<SyncClass, int> _remoteProvider;
        private readonly GenericSynchronisationService<SyncClass, int> _syncService;
        private readonly ISyncSettings _settings;

        public GenericSyncServiceDownloadTests()
        {
            _localProvider = Substitute.For<ISynchronisationProvider<SyncClass, int>>();
            _remoteProvider = Substitute.For<ISynchronisationProvider<SyncClass, int>>();
            _settings = Substitute.For<ISyncSettings>();
            _settings.LastSync = DateTime.Today;
            _syncService = new GenericSynchronisationService<SyncClass, int>(c => c.Id, _localProvider, _remoteProvider, _settings, SyncDirection.Download);
        }

        [Fact]
        public void SingleRemoteEntrySavesLocally()
        {
            //Arrange
            var syncClass = new SyncClass { Id = 1 };
            SetupProviders(null, new[] { syncClass }, null, null);

            //Act
            var results = _syncService.PerformSynchronisation();

            //Assert
            _localProvider.Received().SaveEntries(Arg.Is<List<SyncClass>>(l=>l.First() == syncClass));
            Assert.Equal(0, results.NumberConflicts); //download doesn't track conflicts
            Assert.Equal(0, results.NumberUpdatedRemote);
            Assert.Equal(1, results.NumberUpdatedLocal);
            Assert.Equal(0, results.NumberDeletedRemote);
            Assert.Equal(0, results.NumberDeletedLocal);
        }

        [Fact]
        public void SingleLocalEntryDoesNotSaveRemotely()
        {
            //Arrange
            var syncClass = new SyncClass { Id = 1 };
            SetupProviders(new[] { syncClass }, null, null, null);

            //Act
            var results = _syncService.PerformSynchronisation();

            //Assert
            _remoteProvider.DidNotReceive().SaveEntries(Arg.Is<List<SyncClass>>(l => l.First() == syncClass));
            Assert.Equal(0, results.NumberConflicts); //download doesn't track conflicts
            Assert.Equal(0, results.NumberUpdatedRemote);
            Assert.Equal(0, results.NumberUpdatedLocal);
            Assert.Equal(0, results.NumberDeletedRemote);
            Assert.Equal(0, results.NumberDeletedLocal);
        }

        [Fact]
        public void ConflictedEntryWithRemoteWinsResolver()
        {
            //Arrange
            var syncClass = new SyncClass { Id = 1, Data = "Original Data" };
            var modifiedSyncClass = new SyncClass { Id = 1, Data = "Local modified Data" };
            SetupProviders(new[] { modifiedSyncClass }, new[] { syncClass }, null, null);
            _syncService.ConflcitResolver = new RemoteWinsResolver<SyncClass>();

            //Act
            var results = _syncService.PerformSynchronisation();

            //Assert
            _localProvider.Received().SaveEntries(Arg.Is<List<SyncClass>>(l => l.First() == syncClass));
            Assert.Equal(0, results.NumberConflicts); //download doesn't track conflicts
            Assert.Equal(0, results.NumberUpdatedRemote);
            Assert.Equal(1, results.NumberUpdatedLocal);
            Assert.Equal(0, results.NumberDeletedRemote);
            Assert.Equal(0, results.NumberDeletedLocal);
        }

        private void SetupProviders(SyncClass[] localModified, SyncClass[] remoteModified, int[] localDeletedIds, int[] remoteDeletedIds)
        {
            _remoteProvider.GetModifiedEntries(Arg.Any<DateTime?>()).Returns(remoteModified ?? new SyncClass[0]);
            _localProvider.GetModifiedEntries(Arg.Any<DateTime?>()).Returns(localModified ?? new SyncClass[0]);
            _remoteProvider.GetDeletedEntries(Arg.Any<DateTime?>()).Returns(remoteDeletedIds ?? new int[0]);
            _localProvider.GetDeletedEntries(Arg.Any<DateTime?>()).Returns(localDeletedIds ?? new int[0]);
        }
    }
}