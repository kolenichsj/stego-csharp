using NUnit.Framework;
using System.IO;

namespace tests
{
    [TestFixture]
    public class hipsHTMLTest
    {
        [SetUp]
        public void Setup()
        {

        }

        [Test]
        public void HideInHTML()
        {
            string[] resrcdir = Directory.GetDirectories(TestContext.CurrentContext.WorkDirectory, "resources", (new EnumerationOptions() { MatchCasing = MatchCasing.CaseSensitive }));
            string testResourcesDir = resrcdir[0];
            var overt_inPath = Path.Combine(testResourcesDir, "overt_1.html");
            var overt_outPath = Path.Combine(testResourcesDir, "overt_outPath.html");
            var covert_Path = Path.Combine(testResourcesDir, "covert_1.txt");

            hips.hipsHTML.hideInHTML(overt_inPath, overt_outPath, covert_Path);
            var file1 = new FileStream(overt_outPath, FileMode.Open, FileAccess.Read);
            var file2 = new FileStream(System.IO.Path.Combine(testResourcesDir, "tstCompare.html"), FileMode.Open, FileAccess.Read);

            Assert.That(Utils.FileCompare(file1, file2));
        }
    }
}