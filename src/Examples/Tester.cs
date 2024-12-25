using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Examples
{
    public static class Tester
    {
        const string ResultFolder = "Results";
        static Tester()
        {
            Directory.CreateDirectory(ResultFolder);
        }

        public static void RunAllTests()
        {
            Example00_NewDocument();
            Example01_SimpleWordTmplRendering();
        }

        public static void Example00_NewDocument()
        {
            var docxBytes = WordRender.CreateNew();
            File.WriteAllBytes(
    Path.Combine(ResultFolder, $"NewDocx-{DateTime.Now:HHmmss}.docx"),
    docxBytes);
        }

        public static void Example01_SimpleWordTmplRendering()
        {
            var docxBytes = WordRender.GenerateDocx(File.ReadAllBytes("AnnounceTemplate.docx"),
                new Dictionary<string, string>()
                {
                    ["Title"] = "澄清黑暗執行緒部落格併購傳聞",
                    ["SeqNo"] = "2021-FAKE-001",
                    ["PubDate"] = "2021-04-01",
                    ["Source"] = "亞太地區公關部",
                    ["Content"] = @"
　　坊間媒體盛傳「史塔克工業將以美金 18 億元併購黑暗執行緒部落格」一事，
本站在此澄清並無此事。\n\n
　　史塔克公司執行長日前確實曾派遣代表來訪，雙方就技術合作一事交換意見，
相談甚歡，惟本站暫無出售計劃，且傳聞金額亦不符合本站預估價值(謎之聲：180 元都嫌貴)，
純屬不實資訊。\n\n  
　　本站將秉持初衷，持續發揚野人獻曝、敝帚自珍精神，歡迎舊雨新知繼續支持。"
                });
            File.WriteAllBytes(
                Path.Combine(ResultFolder, $"TmplRender-{DateTime.Now:HHmmss}.docx"),
                docxBytes);
        }


    }
}
