using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace asu_docx_validator.Validation
{
    public static class TitleValidator
    {
        public static String TitleErrorMessage = "DOC_TITLE";

        public static void Validate(Dictionary<string, List<string>> errors, Document document, string filePath)
        {
            IEnumerable<Run> paragraphEnumerable = document.Descendants<Run>();
            List<Run> runs = paragraphEnumerable.Where(run => run.RunProperties?.FontSize?.Val == "32")
                .Where(run => run.InnerText != "")
                .ToList();
            if (runs.Count == 0 || runs.Count > 2)
            {
                Console.WriteLine(runs.Count);
                List<String> errorsList = errors.GetValueOrDefault(filePath, new List<string>());
                errorsList.Add(TitleErrorMessage);
                errors[filePath] = errorsList;
            }
        }
    }
}