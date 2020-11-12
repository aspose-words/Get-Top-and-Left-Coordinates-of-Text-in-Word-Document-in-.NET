using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Replacing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace TopLeftPositionOfString
{
    class Program
    {
        static void Main(string[] args)
        {
            ApplyLicense();

            Document doc = new Document("eSignature.Test.02.docx");
            //Find the text between <<>> and insert bookmark
            doc.Range.Replace(new Regex(@"\<<.*?\>>"), "", new FindReplaceOptions() { ReplacingCallback = new FindAndInsertBookmark() });

            LayoutCollector layoutCollector = new LayoutCollector(doc);
            LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);

            //Display the left top position of text between angle bracket.
            foreach (Bookmark bookmark in doc.Range.Bookmarks)
            {
                if (bookmark.Name.StartsWith("bookmark_"))
                {
                    layoutEnumerator.Current = layoutCollector.GetEntity(bookmark.BookmarkStart);
                    Console.WriteLine(" --> Left : " + layoutEnumerator.Rectangle.Left + " Top : " + layoutEnumerator.Rectangle.Top);
                }
            }
            doc.Save("20.10.docx");

            System.Diagnostics.Process.Start("20.10.docx");
        }

        public class FindAndInsertBookmark : IReplacingCallback
        {
            int i = 1;
            DocumentBuilder builder;
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                // This is a Run node that contains either the beginning or the complete match.
                Node currentNode = e.MatchNode;

                if (builder == null)
                    builder = new DocumentBuilder((Document)currentNode.Document);

                // The first (and may be the only) run can contain text before the match, 
                // in this case it is necessary to split the run.
                if (e.MatchOffset > 0)
                    currentNode = SplitRun((Run)currentNode, e.MatchOffset);

                ArrayList runs = new ArrayList();

                // Find all runs that contain parts of the match string.
                int remainingLength = e.Match.Value.Length;
                while (
                    (remainingLength > 0) &&
                    (currentNode != null) &&
                    (currentNode.GetText().Length <= remainingLength))
                {
                    runs.Add(currentNode);
                    remainingLength = remainingLength - currentNode.GetText().Length;

                    // Select the next Run node. 
                    // Have to loop because there could be other nodes such as BookmarkStart etc.
                    do
                    {
                        currentNode = currentNode.NextSibling;
                    }
                    while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
                }

                // Split the last run that contains the match if there is any text left.
                if ((currentNode != null) && (remainingLength > 0))
                {
                    SplitRun((Run)currentNode, remainingLength);
                    runs.Add(currentNode);
                }

                Run run = (Run)runs[0];
                builder.MoveTo(run);
                builder.StartBookmark("bookmark_" + i);
                builder.EndBookmark("bookmark_" + i);
                i++; ;

                // Signal to the replace engine to do nothing because we have already done all what we wanted.
                return ReplaceAction.Skip;
            }

            /// <summary>
            /// Splits text of the specified run into two runs.
            /// Inserts the new run just after the specified run.
            /// </summary>
            private static Run SplitRun(Run run, int position)
            {
                Run afterRun = (Run)run.Clone(true);
                afterRun.Text = run.Text.Substring(position);
                run.Text = run.Text.Substring(0, position);
                run.ParentNode.InsertAfter(afterRun, run);
                return afterRun;
            }
        }
        private static void ApplyLicense()
        {
            Aspose.Words.License lic = new Aspose.Words.License();
            lic.SetLicense(@"Aspose.Words.lic");
        }
    }
}
