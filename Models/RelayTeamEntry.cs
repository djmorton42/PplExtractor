using System.Collections.Generic;
using CommunityToolkit.Mvvm.ComponentModel;

namespace PplExtractor.Models;

/// <summary>
/// Relay team entry for PPL file: key, name, unused, affiliation
/// </summary>
public partial class RelayTeamEntry : ObservableObject
{
    [ObservableProperty]
    private string _key = string.Empty;

    [ObservableProperty]
    private string _name = string.Empty;

    [ObservableProperty]
    private string _unused = string.Empty;

    [ObservableProperty]
    private string _affiliation = string.Empty;

    public RelayTeamEntry() { }

    public RelayTeamEntry(string key, string name, string unused, string affiliation)
    {
        Key = key;
        Name = name;
        Unused = unused;
        Affiliation = affiliation;
    }

    public RelayTeamEntry CreateCopy() => new(Key, Name, Unused, Affiliation);

    /// <summary>
    /// Default relay teams: key, name, unused, affiliation
    /// </summary>
    public static IReadOnlyList<RelayTeamEntry> GetDefaults() => new[]
    {
        new RelayTeamEntry("LIGHTBLUE", "TEAM LIGHT BLUE", "", "SSO"),
        new RelayTeamEntry("RED", "TEAM RED", "", "SSO"),
        new RelayTeamEntry("BLACK", "TEAM BLACK", "", "SSO"),
        new RelayTeamEntry("WHITE", "TEAM WHITE", "", "SSO"),
        new RelayTeamEntry("YELLOW", "TEAM YELLOW", "", "SSO"),
        new RelayTeamEntry("PINK", "TEAM PINK", "", "SSO"),
        new RelayTeamEntry("GREY", "TEAM GREY", "", "SSO"),
        new RelayTeamEntry("ROYALBLUE", "TEAM ROYAL BLUE", "", "SSO"),
    };
}
