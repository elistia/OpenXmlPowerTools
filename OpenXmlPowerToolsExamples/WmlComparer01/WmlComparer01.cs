﻿using  Codeuctivity.OpenXmlPowerTools;
using Codeuctivity.OpenXmlPowerTools.WmlComparer;
using System;
using System.IO;

namespace WmlComparer01
{
    internal class WmlComparer01
    {
        private static void Main()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            var settings = new WmlComparerSettings();
            var result = WmlComparer.Compare(
                new WmlDocument("../../Source1.docx"),
                new WmlDocument("../../Source2.docx"),
                settings);
            result.SaveAs(Path.Combine(tempDi.FullName, "Compared.docx"));

            var revisions = WmlComparer.GetRevisions(result, settings);
            foreach (var rev in revisions)
            {
                Console.WriteLine("Author: " + rev.Author);
                Console.WriteLine("Revision type: " + rev.RevisionType);
                Console.WriteLine("Revision text: " + rev.Text);
                Console.WriteLine();
            }
        }
    }
}