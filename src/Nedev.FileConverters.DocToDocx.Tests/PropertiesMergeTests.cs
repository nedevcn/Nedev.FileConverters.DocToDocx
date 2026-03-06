#nullable enable
using Xunit;
using Nedev.FileConverters.DocToDocx.Models;

namespace Nedev.FileConverters.DocToDocx.Tests
{
    public class PropertiesMergeTests
    {
        [Fact]
        public void ParagraphProperties_MergeWith_BaseOnlyFillsDefaults()
        {
            var baseProps = new ParagraphProperties
            {
                Alignment = ParagraphAlignment.Center,
                IndentLeft = 720,
                SpaceBefore = 240,
                KeepWithNext = true
            };
            var p = new ParagraphProperties(); // defaults

            p.MergeWith(baseProps);

            Assert.Equal(ParagraphAlignment.Center, p.Alignment);
            Assert.Equal(720, p.IndentLeft);
            Assert.Equal(240, p.SpaceBefore);
            Assert.True(p.KeepWithNext);
        }

        [Fact]
        public void ParagraphProperties_MergeWith_OverridesDefaultLeft()
        {
            var baseProps = new ParagraphProperties
            {
                Alignment = ParagraphAlignment.Right,
                IndentLeft = 720
            };
            var p = new ParagraphProperties
            {
                Alignment = ParagraphAlignment.Left,
                IndentLeft = 360
            };

            p.MergeWith(baseProps);

            // left is treated as "default" and should be overridden by a non-left base
            Assert.Equal(ParagraphAlignment.Right, p.Alignment);
            Assert.Equal(360, p.IndentLeft);
        }

        [Fact]
        public void RunProperties_MergeWith_BaseOnlyFillsDefaults()
        {
            var baseProps = new RunProperties
            {
                FontSize = 48,
                IsBold = true,
                UnderlineType = UnderlineType.Single,
                Color = 5
            };
            var r = new RunProperties();

            r.MergeWith(baseProps);

            Assert.Equal(48, r.FontSize);
            Assert.True(r.IsBold);
            Assert.Equal(UnderlineType.Single, r.UnderlineType);
            Assert.Equal(5, r.Color);
        }

        [Fact]
        public void RunProperties_MergeWith_OverridesDefaultSize()
        {
            var baseProps = new RunProperties
            {
                FontSize = 36,
                IsItalic = true
            };
            var r = new RunProperties
            {
                FontSize = 24,
                IsItalic = false
            };

            r.MergeWith(baseProps);

            // default size 24 should be replaced and italic inherits from base
            Assert.Equal(36, r.FontSize);
            Assert.True(r.IsItalic);
        }
    }
}
