
using System;
using System.Collections.Generic;
using NUnit.Framework;

public class UpcUtils
{
    public static bool FSanitizeStringCore(string s, string sFilter, bool fBackwards, bool fTruncBackwards, out string sNew)
    {
        int i;

        if (fBackwards)
            {
            if ((i = s.LastIndexOf(sFilter)) >= 0)
                {
                if (fTruncBackwards)
                    sNew = s.Substring(i + sFilter.Length);
                else
                    sNew = s.Substring(0, i);

                return true;
                }
            }
        else
            {
            if ((i = s.IndexOf(sFilter)) >= 0)
                {
                if (fTruncBackwards)
                    sNew = s.Substring(i + sFilter.Length);
                else
                    sNew = s.Substring(0, i);
                return true;
                }
            }
        sNew = s;
        return false;
    }

    public static string SanitizeStringCore(string s, string sFilter, bool fBackwards, bool fTruncBackwards)
    {
        string sNew;
        FSanitizeStringCore(s, sFilter, fBackwards, fTruncBackwards, out sNew);
        return sNew;
    }

    public static string SanitizeString(string sTitle, bool fTitle = true)
    {
        // try to get rid of the fluff...
        if (fTitle)
            {
            sTitle = SanitizeStringCore(sTitle, " (DVD)", false, false);
            sTitle = SanitizeStringCore(sTitle, " [DVD", false, false);
            sTitle = SanitizeStringCore(sTitle, " [Blu-ray", false, false);
            sTitle = SanitizeStringCore(sTitle, " (Widescreen", false, false);
            sTitle = SanitizeStringCore(sTitle, "(Widescreen", false, false);
            sTitle = SanitizeStringCore(sTitle, " (Special Edition)", false, false);
            sTitle = SanitizeStringCore(sTitle, " (Blu-ray", false, false);
            sTitle = SanitizeStringCore(sTitle, " Blu-ray", false, false);
            sTitle = SanitizeStringCore(sTitle, " Bluray", false, false);
            sTitle = SanitizeStringCore(sTitle, " (Bluray", false, false);
            sTitle = SanitizeStringCore(sTitle, " DVD ", false, false);
            sTitle = SanitizeStringCore(sTitle, " DVD", false, false);
            sTitle = SanitizeStringCore(sTitle, " [Includes Digital", false, false);
            sTitle = SanitizeStringCore(sTitle, " [3 Discs", false, false);

        }

        sTitle = SanitizeStringCore(sTitle, "Overview\r\n\r\n\r\n\r\n", false, true);
        sTitle = SanitizeStringCore(sTitle, "\r\n\r\n\r\nAdvertising", true, false);
        return sTitle;
    }


    [Test]
    [TestCase("Test", true, "Test")]
    [TestCase("Test DVD ", true, "Test")]
    [TestCase("Test", true, "Test")]
    [TestCase("Test Starring Bob Smith", false, "Test Starring Bob Smith")]
    [TestCase("Overview\r\n\r\n\r\n\r\nRichard", false, "Richard")]
    [TestCase("Festival.\r\n\r\n\r\nAdvertising\r\n$(document)", false, "Festival.")]
    [TestCase("Overview\r\n\r\n\r\n\r\nRichard Festival.\r\n\r\n\r\nAdvertising\r\n$(document)", false, "Richard Festival.")]
    [TestCase("Summers (DVD)", true, "Summers")]
    [TestCase("Box (Special Edition)", true, "Box")]
    [TestCase("Lucy (Blu-ray +", true, "Lucy")]
    [TestCase("Lucy (Blu-ray/DVD/Digital HD)", true, "Lucy")]
    [TestCase("The Visit (Blu-ray + DIGITAL HD)", true, "The Visit")]
    [TestCase("I Am Legend DVD", true, "I Am Legend")]
    [TestCase("Boyhood (Blu-ray +", true, "Boyhood")]
    [TestCase("Stargate, Ark DVD", true, "Stargate, Ark")]
    [TestCase("Testing Blu-ray + DVD + Something else", true, "Testing")]
    [TestCase("Killer (Bluray/DVD Combo) [Blu-ray]", true, "Killer")]
    [TestCase("Session 9 [DVD] [English] [2001]", true, "Session 9")]
    [TestCase("Frontier(s) [DVD] [French] [2007]", true, "Frontier(s)")]
    [TestCase("Independence Day: Resurgence [Blu-ray/DVD]", true, "Independence Day: Resurgence")]
    [TestCase("Star Trek Beyond [Includes Digital Copy] [Blu-ray/DVD]", true, "Star Trek Beyond")]
    [TestCase("Martyrs [DVD] [2008] [Region 1] [US Import] [NTSC]", true, "Martyrs")]
    [TestCase("The Hobbit: The Desolation of Smaug [3 Discs] [Blu-ray/DVD]", true, "The Hobbit: The Desolation of Smaug")]
    [TestCase("Salt (Unrated) (Deluxe Extended Edition) (Blu-ray) (With INSTAWA", true, "Salt (Unrated) (Deluxe Extended Edition)")]
    [TestCase("", true, "")]
    [TestCase("", true, "")]
    [TestCase("", true, "")]
    [TestCase("", true, "")]
    [TestCase("", true, "")]
    public static void TestSanitizeString(string sIn, bool fTitle, string sExpected)
    {
        Assert.AreEqual(sExpected, SanitizeString(sIn, fTitle));
    }

    public static string SanitizeMediaType(string sMediaType)
    {
        if (sMediaType != null)
            {
            sMediaType = sMediaType.ToLower();

            if (sMediaType.IndexOf("blu-ray", StringComparison.Ordinal) >= 0)
                return "BLU-RAY";

            if (sMediaType.IndexOf("dvd", StringComparison.Ordinal) >= 0)
                return "DVD";

            if (sMediaType.IndexOf("laserdisc", StringComparison.Ordinal) >= 0)
                return "LD";
        }

        return "";
    }

    [Test]
    [TestCase("DVD", "DVD")]
    [TestCase("dvd", "DVD")]
    [TestCase("--DVD", "DVD")]
    [TestCase("DVD--", "DVD")]
    [TestCase("blu-ray", "BLU-RAY")]
    [TestCase("laserdisc", "LD")]
    [TestCase("DVD BLU-RAY", "BLU-RAY")]
    [TestCase("blu-ray DVD laserdisc", "BLU-RAY")]
    [TestCase("", "")]
    public static void TestSanitizeMediaType(string sInput, string sExpected)
    {
        Assert.AreEqual(sExpected, SanitizeMediaType(sInput));
    }

    static Dictionary<string, string> s_mpGeneric = new Dictionary<string, string>
    {
        {"comedy", "Comedy"},
        { "action", "Action"},
        { "adventure", "Adventure"},
        { "thriller", "Thriller"},
        { "fantasy", "Fantasy"},
        { "horror", "Horror" },
        { "family", "Family" },
        {"sci-fi", "Sci-Fi" },
        {"drama", "Drama" },
        {"mystery", "Mystery" },
        {"science fiction", "Sci-Fi" },
        {"historical", "Historical" }
    };

    static string[] s_rgsIgnore = new string[] {"Children", "Parody", "Period", "Prehistoric", "Superhero", "Sword", "UK", "War", "French", "Farce", "Costume", "Coming of Age", "Animation", "Alien", "Feature", "Future", "Gay/Lesbian", "Ireland", "Monster", "New Zealand", "Occult", "Military", "Supernatural", "Television", "Armed Forces", "Commandos", "Literary"};
    public static List<string> SanitizeClassList(List<string> pls)
    {
        HashSet<string> classes = new HashSet<string>();

        List<string> plsNew = new List<string>();

        foreach (string sCheck in pls)
            {
            string sCanon = sCheck.ToLower();
            bool fGotGeneric = false;

            foreach (string sGeneric in s_mpGeneric.Keys)
                {
                if (sCanon.Contains(sGeneric))
                    {
                    fGotGeneric = true;
                    classes.Add(s_mpGeneric[sGeneric]);
                    }
                }

            if (fGotGeneric)
                continue; // don't look for specifics if we got one or more generics

            bool fIgnore = false;

            foreach (string sIgnore in s_rgsIgnore)
                if (sCanon.Contains(sIgnore.ToLower()))
                    fIgnore = true;

            if (fIgnore)
                continue;

            // fallthrough
            classes.Add(sCheck);
            }

        foreach (string s in classes)
            plsNew.Add(s);

        return plsNew;
    }

    static List<string> PlsFromString(string s)
    {
        if (s == "")
            return new List<string>();

        return new List<string>(s.Split('|'));
    }

    [Test]
    [TestCase("Comedy", "Comedy")]
    [TestCase("Anarchic Comedy", "Comedy")]
    [TestCase("Comedy Action", "Comedy|Action")]
    [TestCase("Comedy Action|Unknown", "Comedy|Action|Unknown")]
    [TestCase("Sci-Fi|Thriller|Monster|Alien", "Sci-Fi|Thriller")]
    [TestCase("Action Thriller", "Action|Thriller")]
    [TestCase("Adventure Drama", "Adventure|Drama")]
    [TestCase("Aliens", "")]
    [TestCase("Anarchic Comedy", "Comedy")]
    [TestCase("Animation", "")]
    [TestCase("Animation - Features", "")]
    [TestCase("Animation - Kids", "")]
    [TestCase("Biographical Feature", "")]
    [TestCase("Children", "")]
    [TestCase("Children - Animation", "")]
    [TestCase("Comedy - General", "Comedy")]
    [TestCase("Comedy - Teen", "Comedy")]
    [TestCase("Comedy Adventure", "Comedy|Adventure")]
    [TestCase("Coming of Age", "")]
    [TestCase("Computer Animation", "")]
    [TestCase("Costume Adventure", "Adventure")]
    [TestCase("Drama - General", "Drama")]
    [TestCase("Family", "Family")]
    [TestCase("Fantasy Adventure", "Fantasy|Adventure")]
    [TestCase("Fantasy Comedy", "Fantasy|Comedy")]
    [TestCase("Farce", "")]
    [TestCase("French", "")]
    [TestCase("Future Dystopias", "")]
    [TestCase("Gay/Lesbian", "")]
    [TestCase("Historical Film", "Historical")]
    [TestCase("Ireland", "")]
    [TestCase("Monster", "")]
    [TestCase("New Zealand", "")]
    [TestCase("Occult Horror", "Horror")]
    [TestCase("Parody (spoof)", "")]
    [TestCase("Period Drama", "Drama")]
    [TestCase("Prehistoric Fantasy", "Fantasy")]
    [TestCase("Psychological Thriller", "Thriller")]
    [TestCase("Romantic Mystery", "Mystery")]
    [TestCase("Sci-Fi Action", "Sci-Fi|Action")]
    [TestCase("Sci-Fi, Fantasy & Horror - International", "Sci-Fi|Fantasy|Horror")]
    [TestCase("Superhero", "")]
    [TestCase("Sword-and-Sorcery", "")]
    [TestCase("UK", "")]
    [TestCase("War Drama", "Drama")]
    [TestCase("", "")]
    [TestCase("", "")]
    [TestCase("", "")]
    [TestCase("", "")]
    public static void TestSanitizeClassList(string sInput, string sExpected)
    {
        List<string> plsInput = PlsFromString(sInput);
        List<string> plsExpected = PlsFromString(sExpected);

        plsInput = SanitizeClassList(plsInput);
        plsInput.Sort();
        plsExpected.Sort();

        Assert.AreEqual(plsExpected, plsInput);
    }

}