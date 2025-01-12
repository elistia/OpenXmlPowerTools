﻿using Codeuctivity.OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

internal class FieldRetriever01
{
    private static void Main()
    {
        var n = DateTime.Now;
        var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
        tempDi.Create();

        var docWithFooter = new FileInfo("../../DocWithFooter1.docx");
        var scrubbedDocument = new FileInfo(Path.Combine(tempDi.FullName, "DocWithFooterScrubbed1.docx"));
        File.Copy(docWithFooter.FullName, scrubbedDocument.FullName);
        using (var wDoc = WordprocessingDocument.Open(scrubbedDocument.FullName, true))
        {
            ScrubFooter(wDoc);
        }

        docWithFooter = new FileInfo("../../DocWithFooter2.docx");
        scrubbedDocument = new FileInfo(Path.Combine(tempDi.FullName, "DocWithFooterScrubbed2.docx"));
        File.Copy(docWithFooter.FullName, scrubbedDocument.FullName);
        using (var wDoc = WordprocessingDocument.Open(scrubbedDocument.FullName, true))
        {
            ScrubFooter(wDoc);
        }
    }

    private static void ScrubFooter(WordprocessingDocument wDoc)
    {
        foreach (var footer in wDoc.MainDocumentPart.FooterParts)
        {
            FieldRetriever.AnnotateWithFieldInfo(footer);
            var root = footer.GetXDocument().Root;
            RemoveAllButSpecificFields(root);
            footer.PutXDocument();
        }
    }

    private static void RemoveAllButSpecificFields(XElement root)
    {
        var cachedAnnotationInformation = root.Annotation<Dictionary<int, List<XElement>>>();
        var runsToKeep = new List<XElement>();
        foreach (var item in cachedAnnotationInformation)
        {
            var runsForField = root
                .Descendants()
                .Where(d =>
                {
                    var stack = d.Annotation<Stack<FieldRetriever.FieldElementTypeInfo>>();
                    if (stack == null)
                    {
                        return false;
                    }

                    if (stack.Any(stackItem => stackItem.Id == item.Key))
                    {
                        return true;
                    }

                    return false;
                })
                .Select(d => d.AncestorsAndSelf(W.r).FirstOrDefault())
                .GroupAdjacent(o => o)
                .Select(g => g.First())
                .ToList();
            foreach (var r in runsForField)
            {
                runsToKeep.Add(r);
            }
        }
        foreach (var paragraph in root.Descendants(W.p).ToList())
        {
            if (paragraph.Elements(W.r).Any(r => runsToKeep.Contains(r)))
            {
                paragraph.Elements(W.r)
                    .Where(r => !runsToKeep.Contains(r) &&
                        !r.Elements(W.tab).Any())
                    .Remove();
                paragraph.Elements(W.r)
                    .Where(r => !runsToKeep.Contains(r))
                    .Elements()
                    .Where(rc => rc.Name != W.rPr &&
                        rc.Name != W.tab)
                    .Remove();
            }
            else
            {
                paragraph.Remove();
            }
        }
        root.Descendants(W.tbl).Remove();
    }
}