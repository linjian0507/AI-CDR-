using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace AIVectorCore
{
    public sealed class ApiProfile
    {
        [JsonProperty("name")]
        public string Name { get; set; } = "";

        [JsonProperty("apiType")]
        public string ApiType { get; set; } = "openai";

        [JsonProperty("baseUrl")]
        public string BaseUrl { get; set; } = "";

        [JsonProperty("model")]
        public string Model { get; set; } = "";

        [JsonProperty("apiKey")]
        public string ApiKey { get; set; } = "";

        [JsonProperty("vision")]
        public bool Vision { get; set; }

        [JsonProperty("models")]
        public List<string> Models { get; set; } = new List<string>();

        [JsonProperty("imageModel")]
        public string ImageModel { get; set; } = "";

        [JsonIgnore]
        public bool IsAnthropic
        {
            get { return string.Equals(ApiType, "anthropic", StringComparison.OrdinalIgnoreCase); }
        }

        public ApiProfile Clone()
        {
            return new ApiProfile
            {
                Name = Name,
                ApiType = ApiType,
                BaseUrl = BaseUrl,
                Model = Model,
                ApiKey = ApiKey,
                Vision = Vision,
                Models = Models == null ? new List<string>() : new List<string>(Models),
                ImageModel = ImageModel
            };
        }
    }

    public sealed class AppConfig
    {
        [JsonProperty("profiles")]
        public List<ApiProfile> Profiles { get; set; } = new List<ApiProfile>();

        [JsonProperty("activeIndex")]
        public int ActiveIndex { get; set; }

        [JsonProperty("proxy")]
        public string Proxy { get; set; } = "";

        [JsonProperty("cdrProgId")]
        public string CdrProgId { get; set; } = "CorelDRAW.Application.20";

        [JsonProperty("autoLayer")]
        public bool AutoLayer { get; set; } = true;

        [JsonProperty("reverseLayers")]
        public bool ReverseLayers { get; set; }

        [JsonProperty("svgW")]
        public int SvgWidth { get; set; } = 1024;

        [JsonProperty("svgH")]
        public int SvgHeight { get; set; } = 1024;

        [JsonProperty("layerCount")]
        public string LayerCount { get; set; } = "5";

        [JsonProperty("refMode")]
        public string ReferenceMode { get; set; } = "copy";

        [JsonProperty("lastPrompt")]
        public string LastPrompt { get; set; } = "";

        [JsonProperty("styleIndex")]
        public int StyleIndex { get; set; }

        [JsonProperty("variantCount")]
        public string VariantCount { get; set; } = "1";

        [JsonProperty("creativity")]
        public string Creativity { get; set; } = "0.6";

        [JsonProperty("palette")]
        public string Palette { get; set; } = "";

        [JsonProperty("noBg")]
        public bool NoBackground { get; set; }

        [JsonIgnore]
        public ApiProfile ActiveProfile
        {
            get
            {
                if (Profiles == null || Profiles.Count == 0) return null;
                var index = Math.Max(0, Math.Min(ActiveIndex, Profiles.Count - 1));
                return Profiles[index];
            }
        }

        public void Normalize()
        {
            if (Profiles == null) Profiles = new List<ApiProfile>();
            if (ActiveIndex < 0) ActiveIndex = 0;
            if (ActiveIndex >= Profiles.Count && Profiles.Count > 0) ActiveIndex = Profiles.Count - 1;
            if (SvgWidth <= 0) SvgWidth = 1024;
            if (SvgHeight <= 0) SvgHeight = 1024;
            if (string.IsNullOrWhiteSpace(LayerCount)) LayerCount = "5";
            if (string.IsNullOrWhiteSpace(VariantCount)) VariantCount = "1";
            if (string.IsNullOrWhiteSpace(Creativity)) Creativity = "0.6";
            if (Proxy == null) Proxy = "";
            if (CdrProgId == null) CdrProgId = "CorelDRAW.Application.20";
        }

        public static AppConfig Load(string path)
        {
            if (!File.Exists(path))
            {
                var empty = new AppConfig();
                empty.Normalize();
                return empty;
            }

            var json = File.ReadAllText(path);
            var config = JsonConvert.DeserializeObject<AppConfig>(json) ?? new AppConfig();
            config.Normalize();
            return config;
        }

        public void Save(string path)
        {
            Normalize();
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
            File.WriteAllText(path, JsonConvert.SerializeObject(this, Formatting.Indented));
        }
    }

    public sealed class GenerationOptions
    {
        public int Width { get; set; } = 1024;
        public int Height { get; set; } = 1024;
        public string LayerCount { get; set; } = "5";
        public string StyleDescription { get; set; } = "";
        public string Palette { get; set; } = "";
        public bool NoBackground { get; set; }
        public string ReferenceMode { get; set; } = "copy";
        public double Temperature { get; set; } = 0.6;
        public int MaxTokens { get; set; } = 16000;
    }

    public sealed class ImageInput
    {
        public string DataUrl { get; set; } = "";

        public bool HasValue
        {
            get { return !string.IsNullOrWhiteSpace(DataUrl); }
        }
    }

    public sealed class ChatMessage
    {
        public string Role { get; set; } = "user";
        public object Content { get; set; }

        public ChatMessage() { }

        public ChatMessage(string role, object content)
        {
            Role = role;
            Content = content;
        }
    }
}
