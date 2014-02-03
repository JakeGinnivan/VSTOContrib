using System;

namespace Word.TestDoubles
{
    public class Word2013Facade
    {
        public Word2013Facade()
        {
            Application = new ApplicationTestDouble();
        }

        public ApplicationTestDouble Application { get; private set; }

        public Tuple<DocumentTestDouble, WindowTestDouble> NewDocumentInNewWindow()
        {
            var window = (WindowTestDouble) Application.NewWindow();
            var documentTestDouble = new DocumentTestDouble(Application, window);
            ((DocumentsTestDouble)Application.Documents).Add(documentTestDouble);

            return Tuple.Create(documentTestDouble, window);
        }
    }
}