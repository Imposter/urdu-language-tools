using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Text.Json;

namespace UrduLanguageTools
{
    public static class DocumentSettingsExtensions
    {
        private static string GetSettingName<T>() => $"{nameof(UrduLanguageTools)}_{typeof(T).GetType().Name}";

        public static T GetSettings<T>(this Document document, T defaultValue)
            where T : new()
        {
            var settingName = GetSettingName<T>();
            var properties = (DocumentProperties)document.CustomDocumentProperties;
            try
            {
                var property = properties[settingName];
                string serializedValue = property.Value;
                var settings = JsonSerializer.Deserialize<T>(serializedValue);
                return settings;
            }
            catch
            {
                return defaultValue;
            }
        }

        public static void SetSettings<T>(this Document document, T settings)
            where T : new()
        {
            var settingName = GetSettingName<T>();
            var properties = (DocumentProperties)document.CustomDocumentProperties;
            var serializedValue = JsonSerializer.Serialize(settings);
            try
            {
                var property = properties[settingName];
                property.Value = serializedValue;
            }
            catch
            {
                properties.Add(settingName, false, MsoDocProperties.msoPropertyTypeString, serializedValue);
            }
        }

        public static TProperty GetSetting<T, TProperty>(this Document document, Func<T, TProperty> propertySelector)
            where T : new()
        {
            var settings = document.GetSettings(new T());
            return propertySelector(settings);
        }

        public static void SetSetting<T, TProperty>(this Document document, Action<T, TProperty> propertySetter, TProperty value)
            where T : new()
        {
            var settings = document.GetSettings(new T());
            propertySetter(settings, value);
            document.SetSettings(settings);
        }
    }
}
