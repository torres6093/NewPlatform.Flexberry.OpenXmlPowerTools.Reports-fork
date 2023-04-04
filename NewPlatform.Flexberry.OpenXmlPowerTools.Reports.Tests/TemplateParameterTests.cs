using NewPlatform.Flexberry.Reports;
using System;
using Xunit;

namespace NewPlatform.Flexberry.OpenXmlPowerTools.Reports.Tests
{
    public class TemplateParameterTests
    {
        [Fact]
        public void TemplateParameter_ParseNameAndFormat_FormatDateTime()
        {
            // Test
            var subject = new TemplateParameter("ConstructorParameterName:dd-MM-yyyy");
            var dateTime = new DateTime(2021, 1, 1);
            string result = subject.FormatObject(dateTime);

            Assert.Equal("ConstructorParameterName", subject.Name);
            Assert.Equal("01-01-2021", result);
        }
    }
}
