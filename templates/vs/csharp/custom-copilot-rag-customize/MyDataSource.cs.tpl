namespace {{SafeProjectName}}
{
    public class MyDataSource
    {
        private List<string> _data = new List<string>();
        public MyDataSource()
        {
            Init();
        }
        public string RenderData(string query)
        {
            if (query == null)
            {
                return string.Empty;
            }

            foreach (var data in _data)
            {
                if (data.Contains(query))
                {
                    //Console.WriteLine($"return rag data for data contains {query}");
                    return formatDocument(data);
                }
            }
            if (query.ToLower().Contains("perksplus"))
            {
                //Console.WriteLine("return rag data for query contains perksplus");
                return formatDocument(_data[0]);
            }
            else if (query.ToLower().Contains("company") || query.ToLower().Contains("history"))
            {
                //Console.WriteLine("return rag data for query contains company");
                return formatDocument(_data[1]);
            }
            else if (query.ToLower().Contains("northwind") || query.ToLower().Contains("plan"))
            {
                //Console.WriteLine("return rag data for query contains northwind");
                return formatDocument(_data[2]);
            }

            return string.Empty;
        }
        private void Init()
        {
            string[] Documents = Directory.GetFiles("data");

            foreach (string doc in Documents)
            {
                string readText = File.ReadAllText(doc);
                _data.Add(readText);
            }
            return;
        }

        private string formatDocument(string result)
        {
            return $"<context>{result}</context>";
        }
    }
}

