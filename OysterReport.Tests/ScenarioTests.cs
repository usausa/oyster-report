// <copyright file="ScenarioTests.cs" company="PlaceholderCompany">
// Copyright (c) PlaceholderCompany. All rights reserved.
// </copyright>

namespace OysterReport.Tests;

/// <summary>
/// シナリオ・ユースケーステスト。
/// 複数の機能を組み合わせた実践的な利用シナリオを検証する。
/// Scenario and use-case level tests.
/// Tests realistic usage scenarios that combine multiple features.
/// </summary>
public sealed partial class ScenarioTests
{
    private static int CountSubstringOccurrences(string source, string value)
    {
        var count = 0;
        var index = 0;
        while ((index = source.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += value.Length;
        }

        return count;
    }
}
