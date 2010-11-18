using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Rhino.Mocks;
using Xunit;

namespace Office.Utility.Tests
{
    public class ComCleanupEnumeratorFixture
    {
        [Fact]
        public void EnumeratorCallsCleanuponEnumerator()
        {
            var cleanupCalled = false;
            using (new ComCleanupEnumerator<string>(Enumerable.Empty<string>().GetEnumerator(), null, o =>cleanupCalled = true))
            {}

            Assert.True(cleanupCalled);
        }

        [Fact]
        public void EnumeratorCallsCleanuponEvenWhenEnumeratorIsNotDisposable()
        {
            var cleanupCalled = false;
            //IEnumerator does not inherit IDisposable
            var enumerator = MockRepository.GenerateMock<IEnumerator>();

            using (new ComCleanupEnumerator<string>(enumerator, null, o => cleanupCalled = true))
            { }

            Assert.True(cleanupCalled);
        }

        [Fact]
        public void EnumeratorDisposesOfWrappedEnumerableIfDisposable()
        {
            var enumerator = MockRepository.GenerateMock<IEnumerator<string>>();

            using (new ComCleanupEnumerator<string>(enumerator, o => { }, null))
            { }

            enumerator.AssertWasCalled(e => e.Dispose());
        }
        
        [Fact]
        public void EnumeratorCallsCleanupOnSingleElement()
        {
            var cleanupCalled = false;
            var enumerable = new[] { "string" };

            using (var comCleanupEnumerator = new ComCleanupEnumerator<string>(enumerable.GetEnumerator(), o => cleanupCalled = true))
            {
                //Move to first item, then move next will return false because no more elements
                comCleanupEnumerator.MoveNext();
                comCleanupEnumerator.MoveNext();
            }

            Assert.True(cleanupCalled);
        }

        [Fact]
        public void EnumeratorCallsCleanupIfDisposeCalledBeforeMoveNextReturnsFalse()
        {
            var cleanupCalled = false;
            var enumerable = new[] { "string" };

            using (var comCleanupEnumerator = new ComCleanupEnumerator<string>(enumerable.GetEnumerator(), o => cleanupCalled = true))
            {
                //MoveNext will return true, leaving the current value not cleaned up yet.
                comCleanupEnumerator.MoveNext();
            }

            Assert.True(cleanupCalled);
        }

        [Fact]
        public void EnumeratorOnlyReturnsItemsOfCorrectType()
        {
            var items = new object[] { 1, "string" };

            using (var comCleanupEnumerator = new ComCleanupEnumerator<string>(items.GetEnumerator(), o => {}))
            {
                comCleanupEnumerator.MoveNext();
                Assert.Equal("string", comCleanupEnumerator.Current);
                Assert.Equal(false, comCleanupEnumerator.MoveNext());
            }
        }
    }
}
