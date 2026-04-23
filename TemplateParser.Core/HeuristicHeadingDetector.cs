using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace TemplateParser.Core
{
    /// <summary>
    /// Implements multi-signal heuristic heading detection for DOCX paragraphs.
    /// </summary>
    public class HeuristicHeadingDetector
    {
        private readonly int _baselineFontSize;
        public HeuristicHeadingDetector(int baselineFontSize)
        {
            _baselineFontSize = baselineFontSize;
        }

        /// <summary>
        /// Infers heading level and type for a paragraph using font size, bold, spacing, and numbering signals.
        /// Returns: "section", "subsection", "subsubsection", or null if not a heading.
        /// </summary>
        public string? InferHeadingLevel(Paragraph p)
        {
            if (p == null) return null;
            string text = p.InnerText?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text)) return null;

            // --- Font size (max in paragraph) ---
            int maxFontSize = 0;
            foreach (var run in p.Elements<Run>())
            {
                var sz = run.RunProperties?.FontSize?.Val;
                if (sz != null && int.TryParse(sz, out int size))
                {
                    if (size > maxFontSize) maxFontSize = size;
                }
            }

            // --- Bold weight ---
            int boldRuns = 0, totalRuns = 0;
            foreach (var run in p.Elements<Run>())
            {
                totalRuns++;
                if (run.RunProperties?.Bold != null) boldRuns++;
            }

            // --- Spacing ---
            var spacing = p.ParagraphProperties?.SpacingBetweenLines;
            int? before = null, after = null;
            if (spacing != null)
            {
                if (spacing.Before != null && int.TryParse(spacing.Before, out int b)) before = b;
                if (spacing.After != null && int.TryParse(spacing.After, out int a)) after = a;
            }

            // --- Numbering pattern ---
            string numberingPattern = "";
            // Fixed regex: matches patterns like '1.', 'A.', 'a.', '1.2.3', etc.
            var match = Regex.Match(text, @"^(\d+\.|[A-Z]\.\s|[a-z]\.\s|[a-zA-Z]\.)|^\d+(\.\d+)+");
            if (match.Success)
            {
                numberingPattern = match.Value;
            }

            // --- Signal scoring ---
            double score = 0;
            // Font size: +2 if >= 1.2x baseline
            if (maxFontSize >= _baselineFontSize * 1.2) score += 2;
            // Bold: +1 if all or most runs are bold
            if (totalRuns > 0 && boldRuns >= totalRuns * 0.8) score += 1;
            // Spacing before: +1 if large
            if (before.HasValue && before.Value > 200) score += 1;
            // Spacing after: -0.5 if large (body text)
            if (after.HasValue && after.Value > 100) score -= 0.5;
            // Numbering: +1 if heading-like
            if (!string.IsNullOrEmpty(numberingPattern)) score += 1;
            // Short text: +0.5 if <= 10 words
            int wordCount = text.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length;
            if (wordCount <= 10) score += 0.5;

            if (score >= 2.5)
            {
                // Infer level: numbering depth or font size rank
                int level = 1;
                if (!string.IsNullOrEmpty(numberingPattern))
                {
                    int dotCount = numberingPattern.Count(c => c == '.');
                    level = dotCount + 1;
                }
                else if (maxFontSize >= _baselineFontSize * 1.5)
                    level = 1;
                else if (maxFontSize >= _baselineFontSize * 1.3)
                    level = 2;
                else
                    level = 3;
                return level == 1 ? "section" : level == 2 ? "subsection" : "subsubsection";
            }
            return null;
        }
    }
}
