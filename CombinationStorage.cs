using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace DEFOR_Combinations
{
    public static class CombinationStorage
    {
        private static readonly string Folder;
        private static readonly string FilePath;

        static CombinationStorage()
        {
            Folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "DEFOR_Combinations");

            FilePath = Path.Combine(Folder, "combinations.json");
        }

        public static List<CombinationModel> Load()
        {
            try
            {
                if (!Directory.Exists(Folder))
                {
                    Directory.CreateDirectory(Folder);
                }

                if (!File.Exists(FilePath))
                {
                    return new List<CombinationModel>();
                }

                string json = File.ReadAllText(FilePath);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return new List<CombinationModel>();
                }

                var list = JsonConvert.DeserializeObject<List<CombinationModel>>(json);
                return list ?? new List<CombinationModel>();
            }
            catch
            {
                // If anything goes wrong, return empty list
                return new List<CombinationModel>();
            }
        }

        public static void Save(List<CombinationModel> combinations)
        {
            if (combinations == null)
            {
                combinations = new List<CombinationModel>();
            }

            if (!Directory.Exists(Folder))
            {
                Directory.CreateDirectory(Folder);
            }

            string json = JsonConvert.SerializeObject(combinations, Formatting.Indented);
            File.WriteAllText(FilePath, json);
        }
    }
}
