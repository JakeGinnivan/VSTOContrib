using System;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Excel.TestDoubles
{
    public class WorkbooksTestDouble : Workbooks
    {
        readonly List<Workbook> workbooks = new List<Workbook>();

        public WorkbooksTestDouble(ApplicationTestDouble applicationTestDouble)
        {
            Application = applicationTestDouble;
        }

        public Workbook Add(object template)
        {
            var workbookTestDouble = new WorkbookTestDouble(Application);
            workbooks.Add(workbookTestDouble);
            return workbookTestDouble;
        }

        public void Close()
        {
            throw new NotImplementedException();
        }

        IEnumerator Workbooks.GetEnumerator()
        {
            return workbooks.GetEnumerator();
        }

        public Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password,
            object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable,
            object notify, object converter, object addToMru)
        {
            throw new NotImplementedException();
        }

        public void __OpenText(string filename, object origin, object startRow, object dataType,
            XlTextQualifier textQualifier, object consecutiveDelimiter,
            object tab,
            object semicolon, object comma, object space, object other, object otherChar, object fieldInfo,
            object textVisualLayout)
        {
            throw new NotImplementedException();
        }

        public void _OpenText(string filename, object origin, object startRow, object dataType,
            XlTextQualifier textQualifier, object consecutiveDelimiter,
            object tab,
            object semicolon, object comma, object space, object other, object otherChar, object fieldInfo,
            object textVisualLayout, object decimalSeparator, object thousandsSeparator)
        {
            throw new NotImplementedException();
        }

        public Workbook Open(string filename, object updateLinks, object readOnly, object format, object password,
            object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable,
            object notify, object converter, object addToMru, object local, object corruptLoad)
        {
            throw new NotImplementedException();
        }

        public void OpenText(string filename, object origin, object startRow, object dataType,
            XlTextQualifier textQualifier, object consecutiveDelimiter,
            object tab,
            object semicolon, object comma, object space, object other, object otherChar, object fieldInfo,
            object textVisualLayout, object decimalSeparator, object thousandsSeparator, object trailingMinusNumbers,
            object local)
        {
            throw new NotImplementedException();
        }

        public Workbook OpenDatabase(string filename, object commandText, object commandType, object backgroundQuery,
            object importDataAs)
        {
            throw new NotImplementedException();
        }

        public void CheckOut(string filename)
        {
            throw new NotImplementedException();
        }

        public bool CanCheckOut(string filename)
        {
            throw new NotImplementedException();
        }

        public Workbook _OpenXML(string filename, object stylesheets)
        {
            throw new NotImplementedException();
        }

        public Workbook OpenXML(string filename, object stylesheets, object loadOption)
        {
            throw new NotImplementedException();
        }

        public Application Application { get; private set; }
        public XlCreator Creator { get; private set; }
        public object Parent { get; private set; }

        public int Count
        {
            get { return workbooks.Count; }
        }

        Workbook Workbooks.get_Item(object index)
        {
            throw new NotImplementedException();
        }

        public Workbook this[object index]
        {
            get { throw new NotImplementedException(); }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return workbooks.GetEnumerator();
        }
    }
}