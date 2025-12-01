using System.ComponentModel;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace {{SafeProjectName}}.Interop.TeamsSDK;

internal class EnumDescriptionConverter<T> : JsonConverter<T> where T : struct, Enum
{
    public override T Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        string value = reader.GetString();

        if (string.IsNullOrEmpty(value)) return default;

        var type = typeof(T);

        if (!Enum.TryParse(type, value, ignoreCase: false, out object result) &&
            !Enum.TryParse(type, value, ignoreCase: true, out result))
        {
            throw new JsonException(
                $"Unable to convert \"{value}\" to Enum \"{type}\".");
        }

        return (T)result;
    }

    public override void Write(Utf8JsonWriter writer, T value, JsonSerializerOptions options)
    {
        FieldInfo fi = value.GetType().GetField(value.ToString());

        var description = (DescriptionAttribute)fi.GetCustomAttribute(typeof(DescriptionAttribute), false);

        writer.WriteStringValue(description.Description);
    }
}