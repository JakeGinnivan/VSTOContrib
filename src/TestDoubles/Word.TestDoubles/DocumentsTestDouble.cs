using System;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace Word.TestDoubles
{
#pragma warning disable 0067
    public class DocumentsTestDouble : Documents
    {
        readonly List<DocumentTestDouble> documents = new List<DocumentTestDouble>();
        readonly Application application;

        public DocumentsTestDouble(Application application)
        {
            this.application = application;
        }

        IEnumerator Documents.GetEnumerator()
        {
            return documents.GetEnumerator();
        }

        public void Close(ref object saveChanges, ref object originalFormat, ref object routeDocument)
        {
            throw new NotImplementedException();
        }

        public Document AddOld(ref object template, ref object newTemplate)
        {
            throw new NotImplementedException();
        }

        public Document OpenOld(ref object fileName, ref object confirmConversions, ref object readOnly,
            ref object addToRecentFiles,
            ref object passwordDocument, ref object passwordTemplate, ref object revert,
            ref object writePasswordDocument,
            ref object writePasswordTemplate, ref object format)
        {
            throw new NotImplementedException();
        }

        public void Save(ref object noPrompt, ref object originalFormat)
        {
            throw new NotImplementedException();
        }

        public Document Add(ref object template, ref object newTemplate, ref object documentType, ref object visible)
        {
            throw new NotImplementedException();
        }

        public Document Open2000(ref object fileName, ref object confirmConversions, ref object readOnly,
            ref object addToRecentFiles,
            ref object passwordDocument, ref object passwordTemplate, ref object revert,
            ref object writePasswordDocument,
            ref object writePasswordTemplate, ref object format, ref object encoding, ref object visible)
        {
            throw new NotImplementedException();
        }

        public void CheckOut(string fileName)
        {
            throw new NotImplementedException();
        }

        public bool CanCheckOut(string fileName)
        {
            throw new NotImplementedException();
        }

        public Document Open2002(ref object fileName, ref object confirmConversions, ref object readOnly,
            ref object addToRecentFiles,
            ref object passwordDocument, ref object passwordTemplate, ref object revert,
            ref object writePasswordDocument,
            ref object writePasswordTemplate, ref object format, ref object encoding, ref object visible,
            ref object openAndRepair, ref object documentDirection, ref object noEncodingDialog)
        {
            throw new NotImplementedException();
        }

        public Document Open(ref object fileName, ref object confirmConversions, ref object readOnly,
            ref object addToRecentFiles,
            ref object passwordDocument, ref object passwordTemplate, ref object revert,
            ref object writePasswordDocument,
            ref object writePasswordTemplate, ref object format, ref object encoding, ref object visible,
            ref object openAndRepair, ref object documentDirection, ref object noEncodingDialog, ref object xmlTransform)
        {
            throw new NotImplementedException();
        }

        public Document OpenNoRepairDialog(ref object fileName, ref object confirmConversions, ref object readOnly,
            ref object addToRecentFiles, ref object passwordDocument, ref object passwordTemplate, ref object revert,
            ref object writePasswordDocument, ref object writePasswordTemplate, ref object format, ref object encoding,
            ref object visible, ref object openAndRepair, ref object documentDirection, ref object noEncodingDialog,
            ref object xmlTransform)
        {
            if (encoding == null) throw new ArgumentNullException("encoding");
            throw new NotImplementedException();
        }

        public Document AddBlogDocument(string providerId, string postUrl, string blogName, string postId = "")
        {
            throw new NotImplementedException();
        }

        public int Count { get { return documents.Count; } }

        public Application Application { get { return application; } }
        public int Creator { get; private set; }
        public object Parent { get; private set; }

        public Document get_Item(ref object index)
        {
            return documents[((int) index) - 1];
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return documents.GetEnumerator();
        }

        public void Add(DocumentTestDouble document)
        {
            documents.Add(document);
            ((ApplicationTestDouble) Application).OnDocumentOpen(document);
        }
    }
}