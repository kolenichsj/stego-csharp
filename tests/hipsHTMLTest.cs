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
        public void hideInHTML()
        {
            string[] resrcdir = Directory.GetDirectories(TestContext.CurrentContext.WorkDirectory, "resources", (new EnumerationOptions() { MatchCasing = MatchCasing.CaseSensitive }));
            string testResourcesDir = resrcdir[0];
            var overt_inPath = System.IO.Path.Combine(testResourcesDir, "overt_1.html");
            var overt_outPath = System.IO.Path.Combine(testResourcesDir, "overt_outPath.html");
            var covert_Path = System.IO.Path.Combine(testResourcesDir, "covert_1.txt");

            hips.hipsHTML.hideInHTML(overt_inPath, overt_outPath, covert_Path);
            FileStream file1 = new FileStream(overt_outPath, FileMode.Open, FileAccess.Read);
            FileStream file2 = new FileStream(System.IO.Path.Combine(testResourcesDir, "tstCompare.html"), FileMode.Open, FileAccess.Read);

            Assert.True(Utils.FileCompare(file1, file2));
        }
    }
}