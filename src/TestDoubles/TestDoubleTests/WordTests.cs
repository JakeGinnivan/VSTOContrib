using Word.TestDoubles;
using Xunit;

namespace TestDoubleTests
{
    public class WordTests
    {
        [Fact]
        public void WordFacade()
        {
            var word2013Facade = new Word2013Facade();

            var newDocument = word2013Facade.NewDocumentInNewWindow();

            Assert.Equal(1, word2013Facade.Application.Windows.Count);
            Assert.Equal(1, word2013Facade.Application.Documents.Count);
            var window = word2013Facade.Application.Windows[1];
            var document = word2013Facade.Application.Documents[1];
            Assert.Equal(1, document.Windows.Count);
            Assert.Equal(window, document.Windows[1]);
            Assert.Equal(window, newDocument.Item2);
            Assert.Equal(document, newDocument.Item1);
        }
    }
}
