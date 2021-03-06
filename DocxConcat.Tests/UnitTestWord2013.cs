﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Alegor.DocxConcat.Tests
{
    [TestClass]
    public class UnitTestWord2013
    {
        private static readonly string PathWord2013File01 = Environment.CurrentDirectory + @"\..\..\TestFiles\Word 2013\01.docx";
        private static readonly string PathWord2013File02 = Environment.CurrentDirectory + @"\..\..\TestFiles\Word 2013\02.docx";
        private static readonly string PathWord2013File03 = Environment.CurrentDirectory + @"\..\..\TestFiles\Word 2013\03.docx";
        private static readonly string PathOutputWord2013 = Environment.CurrentDirectory + @"\..\..\TestFiles\Word 2013\out\Out.docx";

        [TestMethod]
        public void Test()
        {
            Console.WriteLine($"Currend directory: {Environment.CurrentDirectory}");

            Console.WriteLine($"PathWord2013File01: {PathWord2013File01}");
            Console.WriteLine($"PathWord2013File02: {PathWord2013File02}");
            Console.WriteLine($"PathWord2013File03: {PathWord2013File03}");
            Console.WriteLine($"PathOutputWord2013: {PathOutputWord2013}");
        }

        [TestMethod]
        public void TestConcat()
        {
            var properties = new Properties();
            properties.InputDocumentPathList.Add(PathWord2013File01);
            properties.InputDocumentPathList.Add(PathWord2013File02);
            properties.OutputDocumentPath = PathOutputWord2013;

            var program = new Program(properties);
            program.Run();
        }

        [TestMethod]
        public void TestConcatInsertAppend()
        {
            var properties = new Properties();
            properties.InputDocumentPathList.Add(PathWord2013File01);
            properties.InputDocumentPathList.Add(PathWord2013File02);
            properties.InputDocumentPathList.Add(PathWord2013File03);
            properties.BaseInputDocumentIndex = 1;
            properties.OutputDocumentPath = PathOutputWord2013;

            var program = new Program(properties);
            program.Run();
        }
    }
}
