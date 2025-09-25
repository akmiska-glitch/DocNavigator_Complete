using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Collections.Generic;
using DocNavigator.App.Models;

namespace DocNavigator.App.Services.Metadata
{
    public class DescParser
    {
        private readonly string _folder;
        public DescParser(string folder) => _folder = folder;

       public DescriptorMeta? Load(string descriptorFileName)
{
    // Старый путь: читаем локальный .desc, если он вообще есть.
    // Теперь этот метод используется как бэкап, а основная загрузка идёт по сети.
    var path = Path.Combine(_folder, descriptorFileName);
    if (!File.Exists(path))
        return null;

    var xml = File.ReadAllText(path);
    return ParseFromText(xml);
}

       public DescriptorMeta? ParseFromText(string xml)
{
    if (string.IsNullOrWhiteSpace(xml))
        return null;

    var doc = XDocument.Parse(xml);
    var root = doc.Root;
    if (root == null)
        return null;

    var meta = new DescriptorMeta();

    // content → главная таблица
    var contentEl = root.Descendants()
        .FirstOrDefault(e => e.Name.LocalName.Equals("content", StringComparison.OrdinalIgnoreCase));
    var contentTable = contentEl?.Attribute("table")?.Value;
    if (!string.IsNullOrWhiteSpace(contentTable))
    {
        meta.ContentTable = contentTable!;
        if (!meta.TableCaptions.ContainsKey(contentTable!))
            meta.TableCaptions[contentTable!] = "Content";
    }

    // fieldset-def / nested-fieldset → набор таблиц и их подписи
    foreach (var node in root.Descendants().Where(e =>
                 e.Name.LocalName.Equals("fieldset-def", StringComparison.OrdinalIgnoreCase) ||
                 e.Name.LocalName.Equals("nested-fieldset", StringComparison.OrdinalIgnoreCase)))
    {
        var t = node.Attribute("table")?.Value;
        if (!string.IsNullOrWhiteSpace(t))
        {
            meta.FieldsetTables.Add(t!);

            var caption = node.Attribute("caption")?.Value
                          ?? node.Attribute("documentation")?.Value;
            if (!string.IsNullOrWhiteSpace(caption))
                meta.TableCaptions[t!] = caption!;
        }
    }

    // columns / fields → RU подписи и типы
    // Ищем элементы field/column и собираем:
    //  - sys name: @name
    //  - RU name: @desc || @documentation || @caption
    //  - type:    @type
    //  - id:      @id (для ColumnCaptionsById)
    foreach (var f in root.Descendants()
                 .Where(e => e.Name.LocalName.Equals("field", StringComparison.OrdinalIgnoreCase) ||
                             e.Name.LocalName.Equals("column", StringComparison.OrdinalIgnoreCase)))
    {
        var sys  = f.Attribute("name")?.Value;
        var ru   = f.Attribute("desc")?.Value
                   ?? f.Attribute("documentation")?.Value
                   ?? f.Attribute("caption")?.Value;
        var type = f.Attribute("type")?.Value;

        if (!string.IsNullOrWhiteSpace(sys))
        {
            if (!meta.FieldsBySystemName.ContainsKey(sys!))
                meta.FieldsBySystemName[sys!] = new FieldMeta(sys!, ru, type);
        }

        // КЛЮЧЕВОЕ: id → русское имя (для маппинга колонок по id)
        var id = f.Attribute("id")?.Value;
        if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(ru))
        {
            meta.ColumnCaptionsById[id!] = ru!;
        }
    }

    // Уберём дубли и пустые
    meta.FieldsetTables = meta.FieldsetTables
        .Where(s => !string.IsNullOrWhiteSpace(s))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();

    return meta;

        }
    }
}
