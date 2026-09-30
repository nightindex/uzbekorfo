using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using UzbekOrfoAddIn.Models;

namespace UzbekOrfoAddIn.Services
{
    /// <summary>
    /// Loads and validates the versioned morphology JSON contract.
    /// </summary>
    public static class MorphologyRuleSetLoader
    {
        public static MorphologyRuleSet Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentNullException(nameof(path));
            if (!File.Exists(path)) throw new FileNotFoundException("Morphology rule file was not found.", path);

            string json = File.ReadAllText(path, Encoding.UTF8);
            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var data = serializer.Deserialize<Dictionary<string, object>>(json);
            if (data == null) throw new InvalidDataException("Morphology rule JSON has no root object.");

            var rules = new MorphologyRuleSet
            {
                Version = RequiredString(data, "version"),
                MaxSuffixDepth = RequiredInt(data, "maxSuffixDepth"),
                MinimumRootLength = RequiredInt(data, "minimumRootLength"),
                Vowels = StringList(data, "vowels", required: true),
                Suffixes = ObjectList(data, "suffixes").Select(ParseSuffix).ToList(),
                Families = ObjectList(data, "families").Select(ParseFamily).ToList(),
                AllomorphGroups = ObjectList(data, "allomorphGroups").Select(ParseAllomorphGroup).ToList(),
                RootMutations = ObjectList(data, "rootMutations").Select(ParseRootMutation).ToList()
            };

            List<string> errors = rules.Validate();
            if (errors.Count > 0)
                throw new InvalidDataException("Invalid morphology rules: " + string.Join("; ", errors));
            return rules;
        }

        private static SuffixEntry ParseSuffix(Dictionary<string, object> value)
        {
            if (!value.ContainsKey("productive"))
                throw new InvalidDataException("Every suffix must declare productive=true or false.");

            return new SuffixEntry
            {
                Id = RequiredString(value, "id"),
                Cyrillic = RequiredString(value, "cyrillic"),
                Latin = RequiredString(value, "latin"),
                Category = RequiredString(value, "category"),
                Family = RequiredString(value, "family"),
                Productive = RequiredBool(value, "productive"),
                StandaloneOnly = value.ContainsKey("standaloneOnly") && RequiredBool(value, "standaloneOnly"),
                RequiredRootFlags = OptionalString(value, "requiredRootFlags"),
                RootMutation = OptionalString(value, "rootMutation"),
                AttachesTo = OptionalString(value, "attachesTo"),
                Order = OptionalInt(value, "order"),
                AllomorphCondition = OptionalString(value, "allomorphCondition"),
                MutatesRoot = OptionalString(value, "mutatesRoot"),
                Person = OptionalString(value, "person"),
                Description = OptionalString(value, "description")
            };
        }

        private static MorphologyFamilyRule ParseFamily(Dictionary<string, object> value)
        {
            return new MorphologyFamilyRule
            {
                Id = RequiredString(value, "id"),
                RootLookupSuffix = OptionalString(value, "rootLookupSuffix") ?? string.Empty,
                AllowedRootPartOfSpeech = StringList(value, "allowedRootPartOfSpeech", required: true),
                AllowUnknownRootPartOfSpeech = RequiredBool(value, "allowUnknownRootPartOfSpeech"),
                CategoryStages = IntDictionary(value, "categoryStages"),
                OccurrenceLimits = ObjectList(value, "occurrenceLimits")
                    .Select(item => new CategoryOccurrenceLimit
                    {
                        Categories = StringList(item, "categories", required: true),
                        Max = RequiredInt(item, "max")
                    }).ToList(),
                Requirements = ObjectList(value, "requirements", required: false)
                    .Select(item => new CategoryRequirementRule
                    {
                        IfAny = StringList(item, "ifAny", required: true),
                        RequireAny = StringList(item, "requireAny", required: true)
                    }).ToList()
            };
        }

        private static AllomorphGroupRule ParseAllomorphGroup(Dictionary<string, object> value)
        {
            return new AllomorphGroupRule
            {
                Id = RequiredString(value, "id"),
                SuffixIds = StringList(value, "suffixIds", required: true),
                FallbackSuffixId = RequiredString(value, "fallbackSuffixId")
            };
        }

        private static RootMutationRule ParseRootMutation(Dictionary<string, object> value)
        {
            return new RootMutationRule
            {
                Id = RequiredString(value, "id"),
                FinalCharacterMap = StringDictionary(value, "finalCharacterMap")
            };
        }

        private static List<Dictionary<string, object>> ObjectList(
            Dictionary<string, object> source,
            string key,
            bool required = true)
        {
            if (!source.TryGetValue(key, out object raw) || raw == null)
            {
                if (required) throw new InvalidDataException($"Missing array '{key}'.");
                return new List<Dictionary<string, object>>();
            }

            var enumerable = raw as IEnumerable;
            if (enumerable == null) throw new InvalidDataException($"'{key}' must be an array.");
            var result = new List<Dictionary<string, object>>();
            foreach (object item in enumerable)
            {
                var dictionary = item as Dictionary<string, object>;
                if (dictionary == null) throw new InvalidDataException($"'{key}' contains a non-object value.");
                result.Add(dictionary);
            }
            return result;
        }

        private static List<string> StringList(Dictionary<string, object> source, string key, bool required)
        {
            if (!source.TryGetValue(key, out object raw) || raw == null)
            {
                if (required) throw new InvalidDataException($"Missing array '{key}'.");
                return new List<string>();
            }

            var enumerable = raw as IEnumerable;
            if (enumerable == null || raw is string)
                throw new InvalidDataException($"'{key}' must be an array.");
            return enumerable.Cast<object>().Select(v => v?.ToString()).ToList();
        }

        private static Dictionary<string, int> IntDictionary(Dictionary<string, object> source, string key)
        {
            var values = RequiredObject(source, key);
            var result = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var pair in values)
            {
                if (!int.TryParse(pair.Value?.ToString(), out int number))
                    throw new InvalidDataException($"'{key}.{pair.Key}' must be an integer.");
                result[pair.Key] = number;
            }
            return result;
        }

        private static Dictionary<string, string> StringDictionary(Dictionary<string, object> source, string key)
        {
            return RequiredObject(source, key).ToDictionary(
                pair => pair.Key,
                pair => pair.Value?.ToString(),
                StringComparer.Ordinal);
        }

        private static Dictionary<string, object> RequiredObject(Dictionary<string, object> source, string key)
        {
            if (!source.TryGetValue(key, out object value) ||
                !(value is Dictionary<string, object> result))
                throw new InvalidDataException($"Missing object '{key}'.");
            return result;
        }

        private static string RequiredString(Dictionary<string, object> source, string key)
        {
            string value = OptionalString(source, key);
            if (string.IsNullOrWhiteSpace(value))
                throw new InvalidDataException($"Missing string '{key}'.");
            return value;
        }

        private static string OptionalString(Dictionary<string, object> source, string key)
        {
            return source.TryGetValue(key, out object value) && value != null ? value.ToString() : null;
        }

        private static int RequiredInt(Dictionary<string, object> source, string key)
        {
            if (!source.TryGetValue(key, out object value) ||
                !int.TryParse(value?.ToString(), out int result))
                throw new InvalidDataException($"Missing integer '{key}'.");
            return result;
        }

        private static int OptionalInt(Dictionary<string, object> source, string key)
        {
            return source.TryGetValue(key, out object value) &&
                   int.TryParse(value?.ToString(), out int result)
                ? result
                : 0;
        }

        private static bool RequiredBool(Dictionary<string, object> source, string key)
        {
            if (!source.TryGetValue(key, out object value) || value == null ||
                !bool.TryParse(value.ToString(), out bool result))
                throw new InvalidDataException($"Missing boolean '{key}'.");
            return result;
        }
    }
}
