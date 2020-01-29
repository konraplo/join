using System.Collections.Generic;

namespace TestConsole
{
    public class SinequaProfile
    {
        public string Title { get; set; }
        public List<SinequaSearch> SearchItems { get; } = new List<SinequaSearch>();

    }
}