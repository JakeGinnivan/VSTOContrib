using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace VSTOContrib.Word.Tests
{
    public class TestDocumentsCollection : List<Document>, Documents
    {
        public IEnumerator GetEnumerator()
        {
            return base.GetEnumerator();
        }

        public void Close(ref object SaveChanges, ref object OriginalFormat, ref object RouteDocument)
        {
            throw new System.NotImplementedException();
        }

        public Document AddOld(ref object Template, ref object NewTemplate)
        {
            throw new System.NotImplementedException();
        }

        public Document OpenOld(ref object FileName, ref object ConfirmConversions, ref object ReadOnly, ref object AddToRecentFiles,
            ref object PasswordDocument, ref object PasswordTemplate, ref object Revert, ref object WritePasswordDocument,
            ref object WritePasswordTemplate, ref object Format)
        {
            throw new System.NotImplementedException();
        }

        public void Save(ref object NoPrompt, ref object OriginalFormat)
        {
            throw new System.NotImplementedException();
        }

        public Document Add(ref object Template, ref object NewTemplate, ref object DocumentType, ref object Visible)
        {
            throw new System.NotImplementedException();
        }

        public Document Open2000(ref object FileName, ref object ConfirmConversions, ref object ReadOnly, ref object AddToRecentFiles,
            ref object PasswordDocument, ref object PasswordTemplate, ref object Revert, ref object WritePasswordDocument,
            ref object WritePasswordTemplate, ref object Format, ref object Encoding, ref object Visible)
        {
            throw new System.NotImplementedException();
        }

        public void CheckOut(string FileName)
        {
            throw new System.NotImplementedException();
        }

        public bool CanCheckOut(string FileName)
        {
            throw new System.NotImplementedException();
        }

        public Document Open2002(ref object FileName, ref object ConfirmConversions, ref object ReadOnly, ref object AddToRecentFiles,
            ref object PasswordDocument, ref object PasswordTemplate, ref object Revert, ref object WritePasswordDocument,
            ref object WritePasswordTemplate, ref object Format, ref object Encoding, ref object Visible,
            ref object OpenAndRepair, ref object DocumentDirection, ref object NoEncodingDialog)
        {
            throw new System.NotImplementedException();
        }

        public Document Open(ref object FileName, ref object ConfirmConversions, ref object ReadOnly, ref object AddToRecentFiles,
            ref object PasswordDocument, ref object PasswordTemplate, ref object Revert, ref object WritePasswordDocument,
            ref object WritePasswordTemplate, ref object Format, ref object Encoding, ref object Visible,
            ref object OpenAndRepair, ref object DocumentDirection, ref object NoEncodingDialog, ref object XMLTransform)
        {
            throw new System.NotImplementedException();
        }

        public Document OpenNoRepairDialog(ref object FileName, ref object ConfirmConversions, ref object ReadOnly,
            ref object AddToRecentFiles, ref object PasswordDocument, ref object PasswordTemplate, ref object Revert,
            ref object WritePasswordDocument, ref object WritePasswordTemplate, ref object Format, ref object Encoding,
            ref object Visible, ref object OpenAndRepair, ref object DocumentDirection, ref object NoEncodingDialog,
            ref object XMLTransform)
        {
            throw new System.NotImplementedException();
        }

        public Document AddBlogDocument(string ProviderID, string PostURL, string BlogName, string PostID = "")
        {
            throw new System.NotImplementedException();
        }

        public Application Application { get; private set; }
        public int Creator { get; private set; }
        public object Parent { get; private set; }
        public Document get_Item(ref object Index)
        {
            throw new System.NotImplementedException();
        }
    }
}