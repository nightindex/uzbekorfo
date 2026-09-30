using System;
using System.Collections.Generic;
using System.Linq;

namespace UzbekOrfoAddIn.Models
{
    /// <summary>
    /// Fully data-driven morphology rules loaded from uzbek_suffixes.json.
    /// </summary>
    public sealed class MorphologyRuleSet
    {
        public string Version { get; set; }
        public int MaxSuffixDepth { get; set; }
        public int MinimumRootLength { get; set; }
        public List<string> Vowels { get; set; } = new List<string>();
        public List<SuffixEntry> Suffixes { get; set; } = new List<SuffixEntry>();
        public List<MorphologyFamilyRule> Families { get; set; } = new List<MorphologyFamilyRule>();
        public List<AllomorphGroupRule> AllomorphGroups { get; set; } = new List<AllomorphGroupRule>();
        public List<RootMutationRule> RootMutations { get; set; } = new List<RootMutationRule>();

        /// <summary>
        /// Validates schema-level and cross-reference constraints. Invalid rule
        /// sets are rejected as a unit so spelling acceptance fails closed.
        /// </summary>
        public List<string> Validate()
        {
            var errors = new List<string>();
            if (!string.Equals(Version, "uzbekorfo-suffixes-v2", StringComparison.Ordinal))
                errors.Add("version must be 'uzbekorfo-suffixes-v2'");
            if (MaxSuffixDepth < 1 || MaxSuffixDepth > 12)
                errors.Add("maxSuffixDepth must be between 1 and 12");
            if (MinimumRootLength < 1 || MinimumRootLength > 5)
                errors.Add("minimumRootLength must be between 1 and 5");
            if (Vowels == null || Vowels.Count == 0 || Vowels.Any(v => v == null || v.Length != 1))
                errors.Add("vowels must contain single-character values");
            if (Suffixes == null || Suffixes.Count == 0)
                errors.Add("suffixes must contain at least one entry");
            if (Families == null || Families.Count == 0)
                errors.Add("families must contain at least one entry");

            var familyIds = UniqueIds(Families, f => f?.Id, "family", errors);
            var suffixIds = UniqueIds(Suffixes, s => s?.Id, "suffix", errors);
            var mutationIds = UniqueIds(RootMutations, r => r?.Id, "root mutation", errors);
            UniqueIds(AllomorphGroups, g => g?.Id, "allomorph group", errors);

            foreach (var family in Families ?? new List<MorphologyFamilyRule>())
            {
                if (family == null || string.IsNullOrWhiteSpace(family.Id)) continue;
                if (family.CategoryStages == null || family.CategoryStages.Count == 0)
                    errors.Add($"family '{family.Id}' has no categoryStages");
                else if (family.CategoryStages.Any(kv => string.IsNullOrWhiteSpace(kv.Key) || kv.Value < 1))
                    errors.Add($"family '{family.Id}' has an invalid category stage");

                foreach (var limit in family.OccurrenceLimits ?? new List<CategoryOccurrenceLimit>())
                {
                    if (limit == null || limit.Max < 1 || limit.Categories == null || limit.Categories.Count == 0)
                        errors.Add($"family '{family.Id}' has an invalid occurrence limit");
                    else if (limit.Categories.Any(c => !family.CategoryStages.ContainsKey(c)))
                        errors.Add($"family '{family.Id}' occurrence limit references an unknown category");
                }

                foreach (var requirement in family.Requirements ?? new List<CategoryRequirementRule>())
                {
                    if (requirement == null || requirement.IfAny == null || requirement.IfAny.Count == 0 ||
                        requirement.RequireAny == null || requirement.RequireAny.Count == 0)
                    {
                        errors.Add($"family '{family.Id}' has an invalid requirement");
                        continue;
                    }
                    if (requirement.IfAny.Concat(requirement.RequireAny)
                        .Any(c => !family.CategoryStages.ContainsKey(c)))
                        errors.Add($"family '{family.Id}' requirement references an unknown category");
                }
            }

            foreach (var suffix in Suffixes ?? new List<SuffixEntry>())
            {
                if (suffix == null || string.IsNullOrWhiteSpace(suffix.Id)) continue;
                if (string.IsNullOrWhiteSpace(suffix.Cyrillic) || string.IsNullOrWhiteSpace(suffix.Latin))
                    errors.Add($"suffix '{suffix.Id}' must have Cyrillic and Latin forms");
                if (string.IsNullOrWhiteSpace(suffix.Category))
                    errors.Add($"suffix '{suffix.Id}' has no category");
                if (string.IsNullOrWhiteSpace(suffix.Family) || !familyIds.Contains(suffix.Family))
                    errors.Add($"suffix '{suffix.Id}' references an unknown family");
                else
                {
                    var family = Families.First(f => f.Id == suffix.Family);
                    if (family.CategoryStages == null || !family.CategoryStages.ContainsKey(suffix.Category))
                        errors.Add($"suffix '{suffix.Id}' category is not configured for family '{suffix.Family}'");
                }
                if (!string.IsNullOrWhiteSpace(suffix.RootMutation) && !mutationIds.Contains(suffix.RootMutation))
                    errors.Add($"suffix '{suffix.Id}' references an unknown root mutation");
                if (!IsSupportedCondition(suffix.AllomorphCondition))
                    errors.Add($"suffix '{suffix.Id}' has an unsupported allomorph condition");
                if (suffix.StandaloneOnly && string.IsNullOrEmpty(suffix.RequiredRootFlags))
                    errors.Add($"suffix '{suffix.Id}' needs requiredRootFlags for standalone import");
                if (!string.IsNullOrEmpty(suffix.RequiredRootFlags) &&
                    suffix.RequiredRootFlags.Any(c => !char.IsLetter(c)))
                    errors.Add($"suffix '{suffix.Id}' has invalid requiredRootFlags");
            }

            foreach (var group in AllomorphGroups ?? new List<AllomorphGroupRule>())
            {
                if (group == null || string.IsNullOrWhiteSpace(group.Id)) continue;
                if (group.SuffixIds == null || group.SuffixIds.Count < 1)
                    errors.Add($"allomorph group '{group.Id}' must contain at least one suffix");
                else if (group.SuffixIds.Any(id => !suffixIds.Contains(id)))
                    errors.Add($"allomorph group '{group.Id}' references an unknown suffix");
                if (string.IsNullOrWhiteSpace(group.FallbackSuffixId) ||
                    group.SuffixIds == null || !group.SuffixIds.Contains(group.FallbackSuffixId))
                    errors.Add($"allomorph group '{group.Id}' has an invalid fallback suffix");
            }

            foreach (var mutation in RootMutations ?? new List<RootMutationRule>())
            {
                if (mutation == null || string.IsNullOrWhiteSpace(mutation.Id)) continue;
                if (mutation.FinalCharacterMap == null || mutation.FinalCharacterMap.Count == 0 ||
                    mutation.FinalCharacterMap.Any(kv => kv.Key == null || kv.Key.Length != 1 ||
                                                         kv.Value == null || kv.Value.Length != 1))
                    errors.Add($"root mutation '{mutation.Id}' must map single final characters");
            }

            return errors;
        }

        private static HashSet<string> UniqueIds<T>(
            IEnumerable<T> values,
            Func<T, string> getId,
            string label,
            List<string> errors)
        {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var value in values ?? Enumerable.Empty<T>())
            {
                string id = value == null ? null : getId(value);
                if (string.IsNullOrWhiteSpace(id))
                    errors.Add($"{label} has an empty id");
                else if (!ids.Add(id))
                    errors.Add($"duplicate {label} id '{id}'");
            }
            return ids;
        }

        private static bool IsSupportedCondition(string condition)
        {
            if (string.IsNullOrWhiteSpace(condition) ||
                condition.Equals("default", StringComparison.OrdinalIgnoreCase) ||
                condition.Equals("stemEndsInVowel", StringComparison.OrdinalIgnoreCase))
                return true;

            const string prefix = "stemEndsIn:";
            if (!condition.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return false;
            return condition.Substring(prefix.Length).Split(',')
                .All(value => value.Trim().Length == 1);
        }
    }

    public sealed class MorphologyFamilyRule
    {
        public string Id { get; set; }
        public string RootLookupSuffix { get; set; }
        public List<string> AllowedRootPartOfSpeech { get; set; } = new List<string>();
        public bool AllowUnknownRootPartOfSpeech { get; set; }
        public Dictionary<string, int> CategoryStages { get; set; } =
            new Dictionary<string, int>(StringComparer.Ordinal);
        public List<CategoryOccurrenceLimit> OccurrenceLimits { get; set; } =
            new List<CategoryOccurrenceLimit>();
        public List<CategoryRequirementRule> Requirements { get; set; } =
            new List<CategoryRequirementRule>();
    }

    public sealed class CategoryOccurrenceLimit
    {
        public List<string> Categories { get; set; } = new List<string>();
        public int Max { get; set; }
    }

    public sealed class CategoryRequirementRule
    {
        public List<string> IfAny { get; set; } = new List<string>();
        public List<string> RequireAny { get; set; } = new List<string>();
    }

    public sealed class AllomorphGroupRule
    {
        public string Id { get; set; }
        public List<string> SuffixIds { get; set; } = new List<string>();
        public string FallbackSuffixId { get; set; }
    }

    public sealed class RootMutationRule
    {
        public string Id { get; set; }
        public Dictionary<string, string> FinalCharacterMap { get; set; } =
            new Dictionary<string, string>(StringComparer.Ordinal);
    }
}
