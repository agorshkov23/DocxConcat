using System;
using System.IO;
using System.Runtime.Serialization.Json;

namespace Alegor.DocxConcat
{
    internal class Launcher
    {
        private const string GenerateProjectFile = "-GenerateProjectFile";

        public static int Main(string[] args)
        {
            if (args.Length == 1)
            {
                var serializerSettings = new DataContractJsonSerializerSettings()
                {
                    UseSimpleDictionaryFormat = true
                };

                var serializer = new DataContractJsonSerializer(typeof(Properties), serializerSettings);

                Properties properties = null;
                using (var fileStream = new FileStream(args[0], FileMode.Open))
                {
                    properties = serializer.ReadObject(fileStream) as Properties;
                }

                if (properties == null)
                {
                    Console.Error.WriteLine("Error reading DocxConcat project file!");
                }

                var program = new Program(properties);
                program.Run();

                return 0;
            }

            if (args.Length == 2 && GenerateProjectFile.Equals(args[0], StringComparison.CurrentCultureIgnoreCase))
            {
                var serializerSettings = new DataContractJsonSerializerSettings()
                {
                    UseSimpleDictionaryFormat = true
                };

                var properties = new Properties();
                var serializer = new DataContractJsonSerializer(typeof(Properties), serializerSettings);

                using (var fileStream = new FileStream(args[1], FileMode.OpenOrCreate))
                {
                    serializer.WriteObject(fileStream, properties);
                }

                return 0;
            }

            Console.WriteLine("DocxConcat");
            Console.WriteLine();
            Console.WriteLine("Using:");
            Console.WriteLine("    DocxConcat.exe ProjectFile");
            Console.WriteLine("    DocxConcat.exe -GenerateProjectFile");

            return 0;
        }
    }
}