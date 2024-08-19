using BitlyAPI;
using Newtonsoft.Json;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;

namespace DocLinkConverter.Utility
{
    public class Helper
    {
        private List<WField> hyperlinks;
        private const string APIkey = "AIzaSyAT5gOGUChAnPFJ29DrTGwGfLqjkSLV024";
        public Helper()
        {
            hyperlinks?.Clear();
            hyperlinks = new List<WField>();
        }
        public List<WField> FindAllHyperlinks(WordDocument document)
        {
            //Processes the body contents for each section in the Word document
            foreach (WSection section in document.Sections)
            {
                //Accesses the Body of section where all the contents in document are apart
                WTextBody sectionBody = section.Body;
                IterateTextBody(sectionBody);
                WHeadersFooters headersFooters = section.HeadersFooters;
                //consider that OddHeader & OddFooter are applied to this document
                //Iterates through the TextBody of OddHeader and OddFooter
                IterateTextBody(headersFooters.OddHeader);
                IterateTextBody(headersFooters.OddFooter);
            }
            return hyperlinks;
        }
        private void IterateTextBody(WTextBody textBody)
        {
            //Iterates through each of the child items of WTextBody
            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            {
                //IEntity is the basic unit in DocIO DOM. 
                //Accesses the body items (should be either paragraph or table) as IEntity
                IEntity bodyItemEntity = textBody.ChildEntities[i];
                //A Text body has 3 types of elements - Paragraph, Table and Block Content Control
                //Decides the element type by using EntityType
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = bodyItemEntity as WParagraph;
                        //Processes the paragraph contents
                        //Iterates through the paragraph's DOM
                        IterateParagraph(paragraph);
                        break;
                    case EntityType.Table:
                        //Table is a collection of rows and cells
                        //Iterates through table's DOM
                        IterateTable(bodyItemEntity as WTable);
                        break;
                    case EntityType.BlockContentControl:
                        //Iterates to the body items of Block Content Control
                        IterateTextBody((bodyItemEntity as BlockContentControl).TextBody);
                        break;
                }
            }
        }
        private void IterateParagraph(WParagraph paragraph)
        {
            for (int i = 0; i < paragraph.ChildEntities.Count; i++)
            {
                Entity entity = paragraph.ChildEntities[i];
                //A paragraph can have child elements such as text, image, hyperlink, symbols, etc.,
                //Decides the element type by using EntityType
                switch (entity.EntityType)
                {
                    case EntityType.Field:
                        WField field = entity as WField;
                        if (field.FieldType == FieldType.FieldHyperlink)
                            hyperlinks.Add(field);
                        break;
                }
            }
        }
        private void IterateTable(WTable table)
        {
            //Iterates the row collection in a table
            foreach (WTableRow row in table.Rows)
            {
                //Iterates the cell collection in a table row
                foreach (WTableCell cell in row.Cells)
                {
                    //Table cell is derived from (also a) TextBody
                    //Reusing the code meant for iterating TextBody
                    IterateTextBody(cell);
                }
            }
        }
        private string GetHyperlinkText(int hyperlinkIndex, WParagraph paragraph)
        {
            string text = string.Empty;
            //Add the hyperlink field in stack to get the textrange from nested fields.
            Stack<Entity> fieldStack = new Stack<Entity>();
            fieldStack.Push(paragraph.ChildEntities[hyperlinkIndex]);
            //Flag to get the text from textrange between field separator and end.
            bool isFieldCode = true;
            int i = (hyperlinkIndex + 1);
            while (i < paragraph.Items.Count)
            {
                Entity item = paragraph.ChildEntities[i];
                //If it is nested field, maintain in stack.
                if ((item is WField))
                {
                    fieldStack.Push(item);
                    //Set flag to skip getting text from textrange.
                    isFieldCode = true;
                }
                else if ((item is WFieldMark) && (((WFieldMark)(item)).Type == FieldMarkType.FieldSeparator))
                    //If separator is reached, set flag to read text from textrange.
                    isFieldCode = false;
                else if ((item is WFieldMark) && (((WFieldMark)(item)).Type == FieldMarkType.FieldEnd))
                {
                    //If field end is reached, check whether it is end of hyperlink field and skip the iteration.
                    if (fieldStack.Count == 1)
                    {
                        fieldStack.Clear();
                        return text;
                    }
                    else
                        fieldStack.Pop();
                }
                else if (!isFieldCode && (item is WTextRange))
                    text = (text + ((WTextRange)(item)).Text);
                i = (i + 1);
            }
            return text;
        }
        public async Task RemoveHyperlink(WField field)
        {
            Uri uriResult;
            bool result = Uri.TryCreate(Uri.EscapeUriString(field.FieldValue).Replace("%22", ""), UriKind.Absolute, out uriResult) && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
            if (result)
            {
                var shortUrl = await MyURLShorten(uriResult.OriginalString);
                WParagraph paragraph = field.OwnerParagraph;
                int itemIndex = paragraph.ChildEntities.IndexOf(field);
                if (!string.IsNullOrEmpty(GetHyperlinkText(itemIndex, paragraph).Replace(" ", "")))
                {
                    WTextRange textRange = new WTextRange(paragraph.Document);
                    //Gets the text from hyperlink field.
                    textRange.Text = $"{GetHyperlinkText(itemIndex, paragraph)}({shortUrl})";
                    //Removes the hyperlink field
                    paragraph.ChildEntities.RemoveAt(itemIndex);
                    //Inserts the hyperlink text
                    paragraph.ChildEntities.Insert(itemIndex, textRange);
                }
                else
                    paragraph.AppendText($"({shortUrl})");
            }
        }


        public async Task<string> MyURLShorten(string Longurl)
        {
            var newLink = Longurl;
            try
            {
                var bitly = new Bitly("40dbf4d4ce22aaf025141f22a169737a73454038");
                var linkResponse = await bitly.PostShorten(Longurl);
                newLink = linkResponse?.Link ?? Longurl;
            }
            catch (Exception ex)
            {

            }

            return newLink;
        }

    }

    public class Rootobject
    {
        public string longDynamicLink { get; set; }
        public Suffix suffix { get; set; }
    }

    public class Suffix
    {
        public string option { get; set; }
    }

    public class Response
    {
        public string shortLink { get; set; }
        public string previewLink { get; set; }
    }

}
