using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// Summary description for Cls_Gen
/// </summary>
public class Cls_Gen
{
    public Cls_Gen()
    {
        //
        // TODO: Add constructor logic here
        //
    }
    public static string GetSettings(string str_SettingsName)
    {
        string str_FullSettingsFileString = System.IO.File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + ("/settings.log"));
        string str_StartString = "<" + str_SettingsName + ">";
        string str_EndString = "</" + str_SettingsName + @">";
        string str_SettingsValue = GetStringBetweenTwoStrings(str_FullSettingsFileString, str_StartString, str_EndString);
        return str_SettingsValue;
    }
    public static void SetSettings(string str_SettingsName, string str_SettingsValue)
    {
        string str_FullSettingsFileString = System.IO.File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + ("/settings.log"));
        string str_StartString = "<" + str_SettingsName + ">";
        string str_EndString = "</" + str_SettingsName + @">";
        string str_SettingsOrigValue = GetStringBetweenTwoStrings(str_FullSettingsFileString, str_StartString, str_EndString);
        string str_OriginalString = "<" + str_SettingsName + ">" + str_SettingsOrigValue + "</" + str_SettingsName + @">";
        string str_RevisedString = "<" + str_SettingsName + ">" + str_SettingsValue + "</" + str_SettingsName + @">";
        str_FullSettingsFileString = str_FullSettingsFileString.Replace(str_OriginalString, str_RevisedString);
        System.IO.File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + ("/settings.log"), str_FullSettingsFileString);
    }
    public static string GetStringBetweenTwoStrings(string str_FullString, string str_StartString, string str_EndString)
    {
        //Test with System.Diagnostics.Debug.WriteLine(CLS_Generic.Chp05_7GetStringBetweenTwoStrings("<html><p>Trial Test</p></html>", "<p>", "</p>"));
        return str_FullString.Split(new string[] { str_StartString }, StringSplitOptions.None)[1].Split(new string[] { str_EndString }, StringSplitOptions.None)[0].Trim();
    }
}