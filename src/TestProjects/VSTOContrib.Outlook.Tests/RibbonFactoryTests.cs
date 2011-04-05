using Microsoft.Office.Core;
using Rhino.Mocks;
using Xunit;

namespace Office.Utility.Tests
{
    public class RibbonFactoryTests
    {
        private readonly RibbonFactory _factoryUnderTest;
        private readonly TestRibbonViewModelBase _testRibbonViewModelBase;

        public RibbonFactoryTests()
        {
            _factoryUnderTest = (RibbonFactory)RibbonFactory.Instance;
            _testRibbonViewModelBase = new TestRibbonViewModelBase();
            _factoryUnderTest.InitialiseFactory(new[] {_testRibbonViewModelBase});
        }

        [Fact]
        public void RibbonFactoryForwardsCallback()
        {
            //Arrange
            var button = MockRepository.GenerateStub<IRibbonControl>();
            button.Stub(b => b.Id).Return("buttonId");
            button.Stub(b => b.Tag).Return(typeof (TestRibbonViewModelBase).FullName + "buttonId");

            //Act
            _factoryUnderTest.GetCustomUI(RibbonType.OutlookAppointment.GetEnumDescription());
            _factoryUnderTest.OnAction(button);

            //Assert
            _testRibbonViewModelBase.Mock.AssertWasCalled(r => r.TestButtonClick(button));
        }
    }
}
