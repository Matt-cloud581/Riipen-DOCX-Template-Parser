using System;
using System.Text.Json;
using System.Text.Json.Serialization;
using TemplateParser.Core;

namespace TemplateParser.Core
{
    public class NodeJsonConverter : JsonConverter<Node>
    {
        public override Node Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            throw new NotImplementedException("Deserialization not supported.");
        }

        public override void Write(Utf8JsonWriter writer, Node value, JsonSerializerOptions options)
        {
            writer.WriteStartObject();
            writer.WriteString("id", value.Id);
            writer.WriteString("templateId", value.TemplateId);
            if (value.ParentId.HasValue)
                writer.WriteString("parentId", value.ParentId.Value);
            else
                writer.WriteNull("parentId");
            writer.WriteString("type", value.Type);
            writer.WriteString("title", value.Title);
            writer.WriteNumber("orderIndex", value.OrderIndex);
            // Write metadataJson as raw JSON
            writer.WritePropertyName("metadataJson");
            using (JsonDocument doc = JsonDocument.Parse(value.MetadataJson ?? "{}"))
            {
                doc.RootElement.WriteTo(writer);
            }
            writer.WriteEndObject();
        }
    }
}
