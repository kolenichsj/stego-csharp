using System.IO;

public static class utils
{
        /// <summary>
        /// Executes binary compares of two file streams
        /// </summary>
        /// <param name="fs1"></param>
        /// <param name="fs2"></param>
        /// <returns>True if the files are binary identical</returns>
        public static bool FileCompare(Stream fs1, Stream fs2)
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