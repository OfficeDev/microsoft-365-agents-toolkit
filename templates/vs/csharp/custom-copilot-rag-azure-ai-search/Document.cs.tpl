using Azure.Search.Documents.Indexes;
using Azure.Search.Documents.Indexes.Models;
using System.Text.Json;

namespace {{SafeProjectName}}
{
    public class Document
    {
        [SimpleField(IsKey = true, IsFilterable = true, IsSortable = true)]
        public string DocId { get; set; }

        [SearchableField(IsFilterable = true, IsSortable = true)]
        public string DocTitle { get; set; }

        [SearchableField(AnalyzerName = LexicalAnalyzerName.Values.EnLucene)]
        public string Description { get; set; }

        [VectorSearchField(VectorSearchDimensions = 1536, VectorSearchProfileName = "my-vector-config")]
        public IReadOnlyList<float>? DescriptionVector { get; set; } = null;

        public override string ToString()
        {
            return JsonSerializer.Serialize(this);
        }
    }
}
