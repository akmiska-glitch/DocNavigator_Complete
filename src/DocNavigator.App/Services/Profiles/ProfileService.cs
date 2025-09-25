using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using DocNavigator.App.Models;

namespace DocNavigator.App.Services.Profiles
{
    public class ProfileService
    {
        private readonly string _folder;
        public ProfileService(string folder) => _folder = folder;

        public IEnumerable<DbProfile> LoadAll()
        {
            if (!Directory.Exists(_folder))
                yield break;

            var files = Directory.EnumerateFiles(_folder, "*.json", SearchOption.TopDirectoryOnly);

            // Разрешаем хвостовые запятые и пропускаем комментарии
            var jsonOptions = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                AllowTrailingCommas = true,
                ReadCommentHandling = JsonCommentHandling.Skip
            };

            foreach (var file in files)
            {
                DbProfile? profile = null;
                try
                {
                    var json = File.ReadAllText(file);
                    profile = JsonSerializer.Deserialize<DbProfile>(json, jsonOptions);
                }
                catch (Exception ex)
                {
                    // Не валим всё приложение из-за одного файла; просто пропускаем и логируем
                    Console.WriteLine($"[profiles] Skip '{file}': {ex.Message}");
                }

                if (profile == null)
                    continue;

                // Если какие-то поля отсутствуют — используем значения по умолчанию из модели DbProfile
                profile.Name = string.IsNullOrWhiteSpace(profile.Name) ? Path.GetFileNameWithoutExtension(file) : profile.Name;
                profile.Kind = string.IsNullOrWhiteSpace(profile.Kind) ? "postgres" : profile.Kind;
                profile.ConnectionString ??= string.Empty;
                // DescBaseUrl/DescUrlTemplate/DescVersion уже имеют дефолты в DbProfile

                yield return profile;
            }
        }
    }
}
