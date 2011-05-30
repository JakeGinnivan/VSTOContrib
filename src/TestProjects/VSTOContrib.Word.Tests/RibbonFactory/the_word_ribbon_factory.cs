using System.Runtime.InteropServices;
using Castle.DynamicProxy.Generators;
using Microsoft.Office.Interop.Word;
using NSubstitute;
using VSTOContrib.Word.RibbonFactory;
using Xunit;

namespace VSTOContrib.Word.Tests.RibbonFactory
{
    public class the_word_ribbon_factory
    {
        private readonly Application _application;

        public the_word_ribbon_factory()
        {
            AttributesToAvoidReplicating.Add<MarshalAsAttribute>();
            _application = Substitute.For<Application>();
        }

        [Fact]
        public void can_initialise()
        {
            // arrange
            WordRibbonFactory.SetApplication(_application);

            // act

            // assert

        }
    }
}
