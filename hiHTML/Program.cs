﻿using System;
using System.IO;

namespace hiHTML
{
    class Program
    {
        static void Main(string[] args)
        {
            switch (args.Length)
            {
                case 2:
                    try
                    {
                        hipsHTML.getFromHTML(args[0], args[1]);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(ex.Message);
                    }
                    break;
                case 3:
                    try
                    {
                        hipsHTML.hideInHTML(args[0], args[1], args[2]);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(ex.Message);
                    }
                    break;
                default:
                    DisplayHelp();
                    break;
            }
        }

        private static void DisplayHelp()
        {
            using (var reader = new StreamReader(System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("hiHTML.helpOut.txt")))
            {
                while (!reader.EndOfStream)
                {
                    Console.WriteLine(reader.ReadLine());
                }
            }
        }
    }

    public static class hipsHTML
    {
        private const byte ONE = 32; // space
        private const byte ZERO = 9; // tab

        public static void hideInHTML(string overt_inPath, string overt_outPath, string covert_Path)
        {
            using (FileStream overt_in = File.OpenRead(overt_inPath))
            {
                using (FileStream covert_in = File.OpenRead(covert_Path))
                {
                    using (FileStream overt_out = File.OpenWrite(overt_outPath))
                    {
                        int nextchar = 0;
                        int nextcovbyte = 0;
                        int mask = 256;

                        if (covert_in.CanRead && overt_in.CanRead)
                        {
                            FileInfo covert_fi = new FileInfo(covert_Path);
                            string sizeString = "<!--" + covert_fi.Length.ToString("00000000") + "-->";
                            byte[] sizeinfo = System.Text.Encoding.UTF8.GetBytes(sizeString);
                            overt_out.Write(sizeinfo);
                        }

                        while (overt_in.Position < overt_in.Length || covert_in.Position < covert_in.Length)
                        {
                            nextchar = overt_in.ReadByte();

                            if ((nextchar == ZERO || nextchar == ONE || nextchar == -1) && nextcovbyte != -1)
                            {
                                if ((mask & 256) == 256)
                                {
                                    nextcovbyte = covert_in.ReadByte();
                                    mask = 1;
                                }

                                overt_out.WriteByte(((nextcovbyte & mask) == 0) ? ZERO : ONE);
                                overt_out.FlushAsync();
                                mask = mask << 1;
                            }
                            else
                            {
                                overt_out.WriteByte((byte)nextchar);
                                overt_out.FlushAsync();
                            }
                        }
                    }
                }
            }
        }

        public static void getFromHTML(string overt_inPath, string covert_outPath)
        {
            using (FileStream overt_in = File.OpenRead(overt_inPath))
            {
                using (FileStream covert_out = File.OpenWrite(covert_outPath))
                {
                    byte[] firstline = new byte[15];
                    overt_in.Read(firstline, 0, 15);
                    string marker = System.Text.Encoding.Default.GetString(firstline);
                    int totalbits = int.Parse(marker.Substring(4, 8)) * 8;

                    int mask = 1;
                    byte nextbyte = 0;
                    int bitcount = 0;
                    int nextchar = 0;

                    while (bitcount < totalbits && overt_in.CanRead)
                    {
                        nextchar = overt_in.ReadByte();
                        if (nextchar == ONE)
                        {
                            nextbyte = (byte)(nextbyte | mask);
                            mask = mask << 1;
                            bitcount++;
                        }
                        else if (nextchar == ZERO)
                        {
                            mask = mask << 1;
                            bitcount++;
                        }

                        if ((mask & 256) == 256)
                        {
                            covert_out.WriteByte(nextbyte);
                            mask = 1;
                            nextbyte = 0;
                        }
                    }
                }
            }
        }
    }
}