using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Configuration;
using System.Drawing;
using System.IO;

namespace WorkingDocAndPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            var sourceFilePath = ConfigurationManager.AppSettings["SourceFilePath"];

            var saveFilePath = ConfigurationManager.AppSettings["SaveFilePath"];

            var document = new Document(sourceFilePath);

            var bookmarksNavigator = new BookmarksNavigator(document);

            bookmarksNavigator.MoveToBookmark("client_name");
            bookmarksNavigator.ReplaceBookmarkContent("Ramil", true);

            bookmarksNavigator.MoveToBookmark("client_taxno");
            bookmarksNavigator.ReplaceBookmarkContent("VN-12300254178XY6", true);

            bookmarksNavigator.MoveToBookmark("amount");
            bookmarksNavigator.ReplaceBookmarkContent("871 AZN", true);

            bookmarksNavigator.MoveToBookmark("date");
            bookmarksNavigator.ReplaceBookmarkContent(DateTime.Now.ToString("dd.MM.yyyy"), true);

            var sealPath = ConfigurationManager.AppSettings["SealPath"];

            bookmarksNavigator.MoveToBookmark("seal", true, true);

            var section = document.AddSection();

            var image = Image.FromFile(sealPath);

            var paragraph = section.AddParagraph();

            paragraph.AppendPicture(image);

            bookmarksNavigator.InsertParagraph(paragraph);

            document.Sections.Remove(section);

            if (!Directory.Exists(saveFilePath))
                Directory.CreateDirectory(saveFilePath);

            var saveFileFullPath = $"{saveFilePath}\\{Guid.NewGuid()}.pdf";

            document.IsUpdateFields = true;

            document.SaveToFile(saveFileFullPath, FileFormat.PDF);
        }
    }
}
