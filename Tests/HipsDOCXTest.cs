using NUnit.Framework;
using System.IO;

namespace Tests
{
    [SingleThreaded]
    [TestFixture]
    public class HipsDOCXTest
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        [TestCaseSource("hideTextInDOC")]
        public void GetTextFromFile(int testId, string covertText, string hipsNamespace)
        {
            var resrcdir = Directory.GetDirectories(TestContext.CurrentContext.WorkDirectory, "resources", (new EnumerationOptions() { MatchCasing = MatchCasing.CaseSensitive }));
            var testResourcesDir = resrcdir[0];
            var testcase_path = Path.Combine(testResourcesDir, "testcase" + testId.ToString() + ".docx");
            var extract_text = Hips.HipsDOCX.GetText(testcase_path, hipsNamespace);

            Assert.That(covertText, Is.EqualTo(extract_text));
        }

        static readonly object[] hideTextInDOC =
        [
            new object[] {0, "The best laid plans of mice and men", string.Empty},
            new object[] {1, "Cry havoc and let slip the dogs of war", "https://docs.nunit.org/articles/nunit/technical-notes/usage/"}
        ];
    }
}