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

            Assert.True(FileCompare(file1, file2));
        }

        /// <summary>
        /// Executes binary compares of two file streams
        /// </summary>
        /// <param name="fs1"></param>
        /// <param name="fs2"></param>
        /// <returns>True if the files are binary identical</returns>
        private static bool FileCompare(FileStream fs1, FileStream fs2)
        {
            // Check the file sizes. If they are not the same, the files are not the same.
            if (fs1.Length == fs2.Length)
            {
                // Compare each byte from each file until either a non-matching set
                // of bytes is found or the end of file1 is reached.
                do { } while ((fs1.ReadByte() == fs2.ReadByte()) && (fs1.Position != fs1.Length));

                return (fs1.Position == fs1.Length); //if end of file is reached, files were equal
            }

            return false; // Return false to indicate files are different
        }
    }
}