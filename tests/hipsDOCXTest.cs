using NUnit.Framework;
using System.IO;

namespace tests
{
    [SingleThreaded]
    [TestFixture]
    public class hipsDOCXTest
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        [TestCaseSource("hideTextInDOC")]
        public void getTextFromFile(int testId, string covertText, string hipsNamespace)
        {
            string[] resrcdir = Directory.GetDirectories(TestContext.CurrentContext.WorkDirectory, "resources", (new EnumerationOptions() { MatchCasing = MatchCasing.CaseSensitive }));
            string testResourcesDir = resrcdir[0];
            var testcase_path = Path.Combine(testResourcesDir, "testcase" + testId.ToString() + ".docx");
            string extract_text = hips.hipsDOCX.getText(testcase_path, hipsNamespace);

            Assert.That(covertText, Is.EqualTo(extract_text));
        }

        static object[] hideTextInDOC =
        {
            new object[] {0, "The best laid plans of mice and men", string.Empty},
            new object[] {1, "Cry havoc and let slip the dogs of war", "https://docs.nunit.org/articles/nunit/technical-notes/usage/"}
        };
    }
}